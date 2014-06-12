"""
Simple wrapper class around the NanoWrite software.

This class was intentionally implemented to be as simple as possible to allow easy distribution.
Currently, the only dependency not included in standard python is pywinauto.
"""

import tempfile
import shutil
import time
import datetime
import os
import os.path
import re

import pywinauto
import pywinauto.clipboard
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
    },
    '1.7.5': {
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

    class ExecutionError(Exception):
        pass

    def __init__(self, nanowrite_path=PATH, cache_piezo_position=True):
        """
        Constructor of the NanoWrite class.

        @param nanowrite_path: The path to the NanoWrite executabe.
            The path is used to find the running instance of NanoWrite.
        @type nanowrite_path: str
        """
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

        self._cache_piezo_position = cache_piezo_position
        self._cached_piezo_position = None

    def __del__(self):
        if self._tmpfolder is not None:
            shutil.rmtree(self._tmpfolder)

    def set_dialog_foreground(self, dlg=None):
        self._close_teamviewer_window()
        dlg = dlg if dlg is not None else self._main_dlg
        dlg.Restore()
        dlg.SetFocus()

    @staticmethod
    def _close_teamviewer_window():
        try:
            pwa_app = pywinauto.application.Application()
            pwa_app.connect_(path='teamviewer.exe')
            sponsored_session_window = pwa_app.window_(title_re='Sponsored session')

            if sponsored_session_window.Exists(timeout=0, retry_interval=0):
                sponsored_session_window['OK'].Click()
        except Exception:
            pass

    def get_piezo_range(self):
        return self._piezo_range

    def is_within_piezo_range(self, x, y, z):
        return ((0 <= x <= self._piezo_range[0]) and
                (0 <= y <= self._piezo_range[1]) and
                (0 <= z <= self._piezo_range[2]))

    @staticmethod
    def get_current_log():
        """
        Returns the current log since the start of the NanoWrite program.

        @note:
            This script assumes that the latest log file in %localappdata%\Nanoscribe\Messages
            contains all the recent logs since the program started.


        @return: List of two elements lists, containing a datetime object and the log message.
            The latest log message is the last element of the list.
        @rtype: list
        """
        msgs_dir_path = os.path.join(winpaths.get_local_appdata(), 'Nanoscribe\Messages')
        assert os.path.exists(msgs_dir_path), 'NanoWrite messages log path does not exist'

        # Get the most recent log file name
        # FIXME: This fails if there are other file names than '2013-07-08_16-17-00_Messages.log'
        log_file_name = os.listdir(msgs_dir_path)[-1]

        f = open(os.path.join(msgs_dir_path, log_file_name), 'r')

        results = list()
        for line in f.readlines():
            if len(line) <= 30:
                continue
            line = line.decode('latin-1')
            timestamp_txt = line[:28]
            msg_txt = line[29:]

            if len(timestamp_txt.strip()) == 0:
                assert len(results) > 0, 'No previous timestamp available in log file'
                results[-1][1] += msg_txt
            else:
                # FIXME: Don't ignore time zone offset here
                timestamp_struct = time.strptime(timestamp_txt[:19], '%Y-%m-%dT%H:%M:%S')
                timestamp = datetime.datetime.fromtimestamp(time.mktime(timestamp_struct))
                results.append([timestamp, msg_txt])
        return results

    def execute_mini_gwl(self, commands, execute=True, append_safeguard=True, invalidate_piezo=True):
        """
        Execute gwl commands by inserting them into the mini gwl window.

        @note: This command does not stall until the gwl command has finished.

        @param commands: String or list of strings to insert into the GWL window.
        @type commands: str, list

        @param execute: Execute the command or just insert it.
        @type execute: bool

        @raise NanoWrite.NotReady: Raised if the last command has not finished.
        """
        if not self.has_finished():
            raise NanoWrite.NotReady()

        # Append a safeguard, this way the "done." of the last command does not bother us.
        # Also insert a wait command to force a progress bar.
        commands = 'MessageOut ***Seperator***\n' + commands + '\nwait 0.01'

        # Make sure that the correct dialog has the focus
        self.set_dialog_foreground()

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

            self._job_running = True

            # Switch to camera view, since there might be something of interest there and it does not hurt us
            self.show_camera()

            # Wait for the log to refresh
            time.sleep(1.0)

            if invalidate_piezo:
                self.invalidate_piezo_position()

    def get_command_log(self):
        # Get log of the last command command
        # This assumes the use of the seperator.
        cmd_log = list()
        for timestamp, msg in reversed(self.get_current_log()):
            cmd_log.append((timestamp, msg))
            if '***Seperator***' in msg:
                break

        return reversed(cmd_log)

    def load_gwl_file(self, file_path):
        """
        Load a GWL file from a given path. The file is not automatically executed. Use @p start_dlw for this.

        @param file_path: Path to GWL file.
        @type file_path: str

        @rtype: None
        """
        if not self.has_finished():
            raise NanoWrite.NotReady()

        # Make sure that the correct dialog has the focus
        self.set_dialog_foreground()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['load_structure'])

        while True:
            try:
                time.sleep(0.5)
                open_dlg = self._pwa_app['Open file']
                time.sleep(0.5)
                open_dlg['Edit'].SetEditText(file_path)
                if open_dlg['Edit'].TextBlock() != file_path:
                    continue

                #open_dlg['Open'].Click()
                time.sleep(0.5)
                open_dlg.TypeKeys('{ENTER}')
                break
            except Exception:
                continue

        # Sleep some time, to allow the progress bar to update
        time.sleep(1.0)
        self._job_running = True
        self.wait_until_finished()

    def show_camera(self):
        """
        Switch to the camera view.
        """
        # Make sure that the correct dialog has the focus
        self.set_dialog_foreground()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['camera'])

    def start_dlw(self, invalidate_piezo=True):
        """
        Start writing the loaded DLW file.

        @raise NanoWrite.NotReady: Raised if the last command has not finished.
        """
        if not self.has_finished():
            raise NanoWrite.NotReady()

        # Make sure that the correct dialog has the focus
        self.set_dialog_foreground()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['start_dlw'])

        # Show camera for progress
        self.show_camera()

        self._job_running = True
        if invalidate_piezo:
            self.invalidate_piezo_position()

    def execute_complex_gwl_files(self, start_name, gwl_files, readback_files=None, invalidate_piezo=True):
        """
        Execute a set of possibly several GLW files and read back generated output files.

        @note: The NanoWriteRPC class overwrites this method and encodes the binary return values with BASE64 to
            allow marshaling in XML.

        @param start_name: Name of the executed GLW file.
        @type start_name: str

        @param gwl_files: Dictionary containing the GLW files. Where the key is the filename and the value is the
         content of the file.
        @type gwl_files: dict

        @param readback_files: List of generated files to read back. In most cases these will be pictures.
        @type readback_files: list, tuple

        @return: Dictionary containing the files to read back in @p readback_files.
        @rtype: dict
        """
        assert start_name in gwl_files, 'Invalid start name given'

        # Append a safeguard, this way the "done." of the last command does not bother us.
        # Also insert a wait command to force a progress bar.
        gwl_files[start_name] = 'MessageOut ***Seperator***\n' + gwl_files[start_name] + '\nwait 0.01'

        for filename, content in gwl_files.items():
            # Make sure that an attacker might not access stuff outside of out temporary folder
            file_path = os.path.join(self._tmpfolder, os.path.basename(filename))
            with open(file_path, 'w') as f:
                f.write(content)

        self.load_gwl_file(os.path.join(self._tmpfolder, os.path.basename(start_name)))
        self.wait_until_finished()
        self.start_dlw(invalidate_piezo=invalidate_piezo)

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

    def _get_value_from_selectable_field(self, dlg, pos, sleeps=0.2):
        """
        Get the content of a selectable text field.

        @note: This uses evil hacks which include sending keys and using the clipboard.

        @param dlg: Dialog handle.
        @param pos: Pixel position of the field.
        @return: The value of the text field.
        @rtype: str
        """

        # Make sure that the dialog has the focus
        self.set_dialog_foreground(dlg)
        #dlg.ClickInput(coords=pos)
        #dlg.TypeKeys('^{END}')
        #dlg.TypeKeys('+^{HOME}')
        dlg.DoubleClickInput(coords=pos)
        time.sleep(sleeps)
        dlg.TypeKeys('^c')
        time.sleep(sleeps)
        return pywinauto.clipboard.GetData(format=13)

    def get_progress_time(self):
        """
        Read the progress time field.

        The progress field is located left of the progress bar and counts seconds/minutes spent on the current job.

        @return: The progress time in seconds.
        @rtype: int
        """
        val = self._get_value_from_selectable_field(self._main_dlg,
                                                    self._settings['positions']['progress_txt']).split(':')
        val = [float(x) for x in val]
        seconds = val[0] * 60 * 60 + val[1] * 60 + val[2]
        return seconds

    def get_progress_estimate(self):
        """
        Read the progress estimate field.

        @return: The projected time to complete the job in seconds.
        @rtype: int
        """
        # Make sure that the correct dialog has the focus
        self.set_dialog_foreground()

        # Switch to graph view
        self._main_dlg.ClickInput(coords=self._settings['positions']['graph'])

        val = self._get_value_from_selectable_field(self._main_dlg,
                                self._settings['positions']['progress_estimate_txt']).split(':')

        val = [float(x) for x in val]
        seconds = val[0] * 60 * 60 + val[1] * 60 + val[2]
        return seconds

    def _get_pixel(self, coord):
        """
        Get the (R, G, B) value of the pixel at the given position.
        @param coord: Pixel position.
        @return: Tuple with the (R, G, B) value
        @rtype: tuple
        """
        self.set_dialog_foreground()

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

        print 'Checking finish state:'
        print 'Job running:', self._job_running
        if self._job_running:

            cmd_log = [txt for _, txt in self.get_command_log()]
            for submsg in cmd_log:
                if '!!!' in submsg:
                    self._job_running = False
                    raise NanoWrite.ExecutionError(submsg)

            last_msg = cmd_log[-1]
            print 'Last message:', last_msg
            #if re.match(r'.*((done)|(aborted))\.', last_msg):
            if 'done.' in last_msg or 'aborted.' in last_msg:
                self._job_running = False
                return True
            else:
                return False

        pixel_val = self._get_pixel(self._settings['positions']['finished_pixel'])

        # The process has finished, when the pixel is more or less blue
        bar_full = pixel_val[2] > 240 and pixel_val[1] < 100 and pixel_val[0] < 100

        print 'Bar is full:', bar_full
        if bar_full:
            return True
        else:
            last_msg = self.get_current_log()[-1][1]

            if re.match(r'.*done\.', last_msg):
                return True

        return False

    def wait_until_finished(self, poll_interval=0.5):
        """
        Stall execution until the current job has finished.

        @param poll_interval: Polling interval in seconds
        @type poll_interval: float
        """
        while not self.has_finished():
            time.sleep(poll_interval)

    def abort(self):
        """
        Try to abort the currently running task.

        This is done be clicking on the 'Abort' button.
        """
        # Make sure that the correct dialog has the focus
        self.set_dialog_foreground()

        # Go to advanced settings tab and click into text field
        self._main_dlg.ClickInput(coords=self._settings['positions']['abort'])

        self.wait_until_finished()
        self.invalidate_piezo_position()

    def get_camera_picture(self):
        """
        Get a camera picture as tiff file.

        This is implemented via the mini gwl command window.

        @note: This requires that the camera is actually enabled. Otherwise NanoWrite just hangs...

        @return: The binary tif file.
        @rtype: str
        """
        img_path = os.path.join(self._tmpfolder, 'captured.tif')
        img_meta_path = img_path + '_meta.txt'
        self.execute_mini_gwl("CapturePhoto %s" % img_path, invalidate_piezo=False)
        self.wait_until_finished()

        with open(img_path, 'rb') as f:
            img_data = f.read()

        with open(img_meta_path, 'rb') as f:
            meta_data = f.read()

        return meta_data, img_data

    def invalidate_piezo_position(self):
        """
        Invalidate the chached piezo position.

        Use, when you know that an outside instance manipulated the piezo position.
        """
        print('Invalidating piezo position')
        self._cached_piezo_position = None

    def get_piezo_position(self):
        """
        Returns the current piezo position corrected by the z-inversion feature.

        These coordinates correspond directly to the coordinates used in GLW commands.

        @return: Tuple of x, y, z coordinates
        @rtype: tuple
        """

        if self._cached_piezo_position is not None and self._cache_piezo_position:
            return self._cached_piezo_position

        val_x = float(self._get_value_from_selectable_field(self._main_dlg,
                                self._settings['positions']['piezo_x_txt']))
        val_y = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['piezo_y_txt']))
        val_z = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['piezo_z_txt']))

        if self.is_z_inverted():
            val_x = self._piezo_range[0] - val_x
            val_z = self._piezo_range[2] - val_z

        piezo_position = val_x, val_y, val_z
        self._cached_piezo_position = piezo_position
        return piezo_position

    def is_z_inverted(self):
        """
        Check if the z-Axis inversion is enabled.

        @return: Returns True if the z-axis is inverted.
        @rtype: bool
        """
        return self._get_pixel(self._settings['positions']['inverted_z_axis_pixel'])[1] > 100

    def set_z_inverted(self, state):
        """
        Set the z-axis inversion to the given state.

        If the state is already set, no action is executed.

        @param state: True if the z-axis is to be inverted.
        @type state: bool
        """
        if self.is_z_inverted() ^ state:
            # Make sure that the correct dialog has the focus
            self.set_dialog_foreground()

            # Go to advanced settings tab and click into text field
            self._main_dlg.ClickInput(coords=self._settings['positions']['inverted_z_axis_pixel'])

        assert self.is_z_inverted() == state, "Invert z-state does not match"

    def get_stage_position(self):
        """
        Get the current stage position.

        @return: Tuple of x, y, z coordinates
        @rtype: tuple
        """
        #FIXME: This might also need a z-inversion correction.

        val_x = float(self._get_value_from_selectable_field(self._main_dlg,
                                self._settings['positions']['stage_x_txt']))
        val_y = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['stage_y_txt']))
        val_z = float(self._get_value_from_selectable_field(self._main_dlg,
                        self._settings['positions']['stage_z_txt']))
        return val_x, val_y, val_z

    def _get_screenshot(self):
        """
        Get a screenshot of the main window.

        @return: A PIL image object.
        """
        self.set_dialog_foreground()
        return self._main_dlg.CaptureAsImage()

    def find_interface(self, at=50):
        gwl = 'findInterfaceAt %f' % at
        self.execute_mini_gwl(gwl)
        self.wait_until_finished()

    def move_piezo(self, x, y, z=None):
        if z is None:
            z = self.get_piezo_position()[2]

        new_pos = (x, y, z)
        new_pos_valid = self.is_within_piezo_range(*new_pos)
        gwl = '%f %f %f 0\nwrite' % new_pos
        self.execute_mini_gwl(gwl, invalidate_piezo=True)
        self.wait_until_finished()

        self._cached_piezo_position = new_pos if new_pos_valid else None
        # Give it some time to settle
        time.sleep(0.5)

    def move_piezo_relative(self, dx=0, dy=0, dz=0):
        piezo_position = self.get_piezo_position()
        self.move_piezo(piezo_position[0] + dx, piezo_position[1] + dy, piezo_position[2] + dz)

    def move_stage(self, x, y, z=None):
        current_stage_pos = self.get_stage_position()
        if z is None:
            z = current_stage_pos[2]

        delta_z = z-current_stage_pos[2]

        if self.is_z_inverted():
            delta_z *= -1

        self.move_stage_relative(x-current_stage_pos[0],
                                 y-current_stage_pos[1],
                                 delta_z)

    def move_stage_relative(self, dx=0, dy=0, dz=0):
        gwl = 'MoveStageX %f\nMoveStageY %f\nAddZDrivePosition %f\nwrite' % (dx, dy, dz)
        self.execute_mini_gwl(gwl, invalidate_piezo=False)
        self.wait_until_finished()

        # Give it some time to settle
        time.sleep(0.5)

    def move_piezo_to_same_location_by_stage(self, x, y):
        """
        Move the piezo to the given location and the stage in the opposite direction.

        In the end, the microscope should be at the same location.
        """
        piezo_position = self.get_piezo_position()
        piezo_correction = (x - piezo_position[0], y - piezo_position[1])
        self.move_piezo(x, y, piezo_position[2])
        self.move_stage_relative(-piezo_correction[0], -piezo_correction[1])


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