"""
This file implements a simple XML-RPC server for the nanowrite wrapper.

The NanoWrite class is extended to wrap the binary files in BASE64 to allow
easy marshalling into XML. No further changes are implemented.
"""

import time
import base64

from DocXMLRPCServer import DocXMLRPCServer
from nanowrite import NanoWrite


class NanoWriteRPC(NanoWrite):
    def __init__(self, *args, **nargs):
        NanoWrite.__init__(self, *args, **nargs)

    def get_camera_picture(self):
        """
        Get a camera picture as BASE64 encoded tiff file.

        This is implemented via the mini gwl command window.

        @note: This requires that the camera is actually enabled. Otherwise NanoWrite just hangs...

        @return: The BASE64 encoded tif file
        @rtype: str
        """
        meta, pic = NanoWrite.get_camera_picture(self)
        return meta, base64.encodestring(pic)

    def execute_complex_gwl_files(self, start_name, gwl_files, readback_files=None):
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

        @return: Dictionary containing the files to read back in @p readback_files encoded into BASE64 strings.
        @rtype: dict
        """
        results = NanoWrite.execute_complex_gwl_files(self, start_name, gwl_files, readback_files)

        return {key: base64.encodestring(value) for key, value in results.items()}

if __name__ == '__main__':
    server = DocXMLRPCServer(('', 60000), logRequests=1, allow_none=True)
    server.register_introspection_functions()
    server.register_instance(NanoWriteRPC())

    print time.asctime(), 'Server starting'
    server.serve_forever()
    print time.asctime(), 'Server finishing'
