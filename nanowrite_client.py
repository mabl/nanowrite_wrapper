"""
This file implements a simple XML-RPC server for the nanowrite wrapper.

The NanoWrite class is extended to wrap the binary files in BASE64 to allow
easy marshalling into XML. No further changes are implemented.
"""

import time
import base64
import xmlrpclib

class NanoWriteRPCClient(object):
    """
    This class mimics the same behaviour as the NanoWrite class but connects over network to the XML-RPC server.

    You can easily substitute instances of and NanoWriteRPCClient without any loss in functionality.
    """

    def __init__(self, uri, *args, **nargs):
        self._proxy = xmlrpclib.ServerProxy(uri, *args, allow_none=True, **nargs)

    def __getattr__(self, item):
        if item not in self.__dict__:
            return self.__dict__['_proxy'].__getattr__(item)
        else:
            return self.__dict__[item]

    def get_camera_picture(self):
        meta, img = self._proxy.get_camera_picture()
        return meta, img.data

    def execute_complex_gwl_files(self, start_name, gwl_files, readback_files=None):
        results = self._proxy.execute_complex_gwl_files(start_name, gwl_files, readback_files)
        return {key: value.data for key, value in results.items()}

    def wait_until_finished(self, poll_interval=0.5):
        # Wait on the client side to avoid timeouts.
        while not self._proxy.has_finished():
            time.sleep(poll_interval)