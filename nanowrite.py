import pywinauto
import pywinauto.clipboard

import tempfile
import shutil
import time
import datetime
import os
import os.path
import re

import winpaths

PATH = r"C:\Program Files\Nanoscribe\NanoWrite\NanoWrite.exe"

SETTINGS = {
    '1.7.1': {
        'positions': {
            # Tab options and fields therein
            'advanced_settings': (268, 40),
            'advanced_settings_textfield': (268, 100),
            'advanced_settings_submit': (582, 110),

            'camera': (190, 40),
            'graph': (140, 40),

            # Buttons on the left
            'exchange_holder': (50, 125),
            'load_structure': (50, 180),
            'approach_sample': (50, 230),
            'start_dlw': (50, 510),
            'abort': (50, 580),

            # Position of some text fields
            'progress_txt': (190, 610),
            'progress_estimate_txt': (615, 560),

            'piezo_x_txt': (720, 390),
            'piezo_y_txt': (720, 410),
            'piezo_z_txt': (720, 435),

            'stage_x_txt': (820, 390),
            'stage_y_txt': (820, 410),
            'stage_z_txt': (820, 435),

            # Some pixel definitions
            'finished_pixel': (641, 608),
            'inverted_z_axis_pixel': (960, 408)
        }
    }
}


class NanoWrite(object):
    class NotReady(Exception):
        pass

    def __init__(self, nanowrite_path=PATH):
        self._tmpfolder = None

        self._pwa_app = pywinauto.application.Application()
        self._pwa_app.connect_(path=nanowrite_path)

        self._main_dlg = self._pwa_app.window_(title_re='.*NanoWrite .+')

        self._version = self._main_dlg.WindowText().split(' ')[-1]
        assert self._version in SETTINGS, 'Program version not known'

        self._settings = SETTINGS[self._version]

        self._tmpfolder = tempfile.mkdtemp(suffix='nanowriteserver')
        self._job_running = False

        self._piezo_range = (300, 300, 300)

    def __del__(self):
        if self._tmpfolder is not None:
            shutil.rmtree(self._tmpfolder)

    def get_current_log(self):
        """
        Get the current log file content.
        """

        # This script assumes that the latest log file in %localappdata%\Nanoscribe\Messages
        # contains all the recent logs since the program started.

        msgs_dir_path = os.path.join(winpaths.get_local_appdata(), 'Nanoscribe\Messages')
        assert os.path.exists(msgs_dir_path), 'NanoWrite messages log path does not exist'

        # Get the most recent log file name
        # FIXME: This fails if there are other file names than '2013-07-08_16-17-00_Messages.log'
        log_file_name = os.listdir(msgs_dir_path)[-1]

        f = open(os.path.join(msgs_dir_path, log_file_name), 'r')

        results = list()
        current_timestamp = None
        for line in f.readlines():
            if len(line) <= 30:
                continue
            line = line.decode('latin-1')
            timestamp_txt = line[:28]
            msg_txt = line[29:]

            if len(timestamp_txt.strip()) == 0:
                results[-1][1] += msg_txt
            else:
                # FIXME: Don't ignore time zone offset here
                timestamp_struct = time.strptime(timestamp_txt[:19], '%Y-%m-%dT%H:%M:%S')
                timestamp = datetime.datetime.fromtimestamp(time.mktime(timestamp_struct))
                results.append([timestamp, msg_txt])
        return results

    def execute_mini_gwl(self, commands, execute=True, append_safeguard=True):
        if not self.has_finished():
            raise NanoWrite.NotReady()

        if append_safeguard:
            commands = 'MessageOut ***Seperator**\n' + commands + '\nwait 0.01'

        # Make sure that the correct dialog has the focus
        self._main_dlg.SetFocus()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['advanced_settings'])
        self._main_dlg.ClickInput(coords=self._settings['positions']['advanced_settings_textfield'])

        # Select all and delete existing text
        # We do this be going to the end of existing input by CTRL+END
        # Select all existing text upwards via SHIFT+CTRL+HOME
        # Delete the text via DEL
        self._main_dlg.TypeKeys('^{END}')
        self._main_dlg.TypeKeys('+^{HOME}')
        self._main_dlg.TypeKeys('{DEL}')

        import win32clipboard
        import win32con
        win32clipboard.OpenClipboard()
        win32clipboard.SetClipboardData(win32con.CF_TEXT, commands)
        win32clipboard.CloseClipboard()
        self._main_dlg.TypeKeys('^v')

        # And execute command if asked for
        if execute:
            self._main_dlg.ClickInput(coords=self._settings['positions']['advanced_settings_submit'])
            # Wait for the log to refresh
            time.sleep(1.0)

            self._job_running = True

            # Switch to camera view, since there might be something of interest there and it does not hurt us
            self.show_camera()

    def load_gwl_file(self, file_path):
        if not self.has_finished():
            raise NanoWrite.NotReady()

        # Make sure that the correct dialog has the focus
        self._main_dlg.SetFocus()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['load_structure'])

        open_dlg = self._pwa_app['Open file']
        open_dlg['Edit'].SetEditText(file_path)
        #open_dlg['Open'].Click()
        open_dlg.TypeKeys('{ENTER}')

        # Sleep some time, to allow the progress bar to update
        time.sleep(0.5)

    def show_camera(self):
        # Make sure that the correct dialog has the focus
        self._main_dlg.SetFocus()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['camera'])


    def start_dlw(self):
        if not self.has_finished():
            raise NanoWrite.NotReady()

        # Make sure that the correct dialog has the focus
        self._main_dlg.SetFocus()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['start_dlw'])

        # Show camera for progress
        self.show_camera()

        self._job_running = True

    def execute_complex_gwl_files(self, start_name, gwl_files, readback_files=None):
        assert start_name in gwl_files, 'Invalid start name given'

        for filename, content in gwl_files.items():
            # Make sure that an attacker might not access stuff outside of out temporary folder
            file_path = os.path.join(self._tmpfolder, os.path.basename(filename))
            with open(file_path, 'w') as f:
                f.write(content)

        self.load_gwl_file(os.path.join(self._tmpfolder, os.path.basename(start_name)))
        self.wait_until_finished()
        self.start_dlw()

        if readback_files is not None:
            time.sleep(5)
            self.wait_until_finished()
            results = dict()
            for filename in readback_files:
                file_path = os.path.join(self._tmpfolder, os.path.basename(filename))
                #print 'Read back:', file_path
                with open(file_path, 'rb') as f:
                    results[filename] = f.read()
            return results
        return {}

    def _get_value_from_selectable_field(self, dlg, pos):
        # Make sure that the correct dialog has the focus
        dlg.SetFocus()
        dlg.ClickInput(coords=pos)

        self._main_dlg.TypeKeys('^{END}')
        self._main_dlg.TypeKeys('+^{HOME}')
        self._main_dlg.TypeKeys('^c')

        return pywinauto.clipboard.GetData(format=13)

    def get_progress_time(self):
        val = self._get_value_from_selectable_field(self._main_dlg,
                                self._settings['positions']['progress_txt'])
        return val

    def get_progress_estimate(self):
        # Make sure that the correct dialog has the focus
        self._main_dlg.SetFocus()

        # Switch to graph view
        self._main_dlg.ClickInput(coords=self._settings['positions']['graph'])

        val = self._get_value_from_selectable_field(self._main_dlg,
                                self._settings['positions']['progress_estimate_txt'])
        return val

    def _get_pixel(self, coord):
        # Make sure that the correct dialog has the focus
        self._main_dlg.SetFocus()

        img = self._main_dlg.CaptureAsImage()
        return img.convert('RGB').getpixel(coord)

    def has_finished(self):
        """
        Check if the instrument has finished with all of its tasks.

        This is a rather tricky thing to do, since there is no common and reliable indicator.
        For now we resort in a sequence of these:

        If we now, that a job is running, the last line must contain a 'done' or 'aborted'.
        In case of no running job we need to fall back on the progress bar:
        - If the last pixel of the progress bar is not blue, it has not finished, except it has a
          'done.' at the end.
        """

        #print 'Checking finish state:'
        #print 'Job running:', self._job_running
        if self._job_running:
            last_msg = self.get_current_log()[-1][1].splitlines(True)[-1]
            #print 'Last message:', last_msg
            if re.match(r'.*((done)|(aborted))\.', last_msg):
                self._job_running = False
                return True
            else:
                return False

        pixel_val = self._get_pixel(self._settings['positions']['finished_pixel'])

        # The process has finished, when the pixel is more or less blue
        bar_full = pixel_val[2] > 240 and pixel_val[1] < 100 and pixel_val[0] < 100

        #print 'Bar is full:', bar_full
        if bar_full:
            return True
        else:
            last_msg = self.get_current_log()[-1][1]

            if re.match(r'.*done\.', last_msg):
                return True

        return False

    def wait_until_finished(self, poll_interval=0.5):
        while not self.has_finished():
            time.sleep(poll_interval)

    def abort(self):
        """
        Try to abort the currently running task.

        This is done be clicking on the 'Abort' button.
        """
        # Make sure that the correct dialog has the focus
        self._main_dlg.SetFocus()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['abort'])

        self.wait_until_finished()

    def get_camera_picture(self):
        """
        Get a camera picture as tiff file.

        This is implemented via the mini gwl command window.

        NOTE: This requires that the camera is actually enabled. Otherwise nanowrite just hangs...
        """
        img_path = os.path.join(self._tmpfolder, 'captured.tif')
        img_meta_path = img_path + '_meta.txt'
        self.execute_mini_gwl("CapturePhoto %s" % img_path)
        self.wait_until_finished()

        img_data = None
        with open(img_path, 'rb') as f:
            img_data = f.read()

        meta_data = None
        with open(img_meta_path, 'rb') as f:
            meta_data = f.read()

        return meta_data, img_data

    def get_piezo_position(self):
        valX = float(self._get_value_from_selectable_field(self._main_dlg,
                                self._settings['positions']['piezo_x_txt']))
        valY = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['piezo_y_txt']))
        valZ = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['piezo_z_txt']))

        if self.is_z_inverted():
            valX = self._piezo_range[0] - valX
            valZ = self._piezo_range[2] - valZ

        return valX, valY, valZ

    def is_z_inverted(self):
        return self._get_pixel(self._settings['positions']['inverted_z_axis_pixel'])[1] > 100

    def set_z_inverted(self, state):
        if self.is_z_inverted() ^ state:
            # Make sure that the correct dialog has the focus
            self._main_dlg.SetFocus()

            # Go to advanced settings tab and click into text field
            self._main_dlg.ClickInput(coords=self._settings['positions']['inverted_z_axis'])

        assert self.is_z_inverted() == state, "Invert z-state does not match"

    def get_stage_position(self):
        valX = float(self._get_value_from_selectable_field(self._main_dlg,
                                self._settings['positions']['stage_x_txt']))
        valY = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['stage_y_txt']))
        valZ = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['stage_z_txt']))
        return valX, valY, valZ

    def _get_screenshot(self):
        return self._main_dlg.CaptureAsImage()

def main():
    nanowrite = NanoWrite()

    nanowrite.execute_mini_gwl('test')
    nanowrite.wait_until_finished()
    #nanowrite.load_gwl_file(r'F:\testgwl\test.gwl')
    print nanowrite.get_progress_time()
    print nanowrite.get_progress_estimate()
    print nanowrite.has_finished()
    time.sleep(5)
    nanowrite.abort()
    print nanowrite.has_finished()
    print nanowrite.get_camera_picture()[0]
    #nanowrite._get_screenshot().save(r'F:\screenshot.png')
    print nanowrite.get_piezo_position(), nanowrite.get_stage_position()

    print nanowrite.is_z_inverted()
    nanowrite.set_z_inverted(True)

if __name__ == '__main__':
    print 'Starting...'
    main()