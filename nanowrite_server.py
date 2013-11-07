"""
This file implements a simple XML-RPC server for the nanowrite wrapper.

The NanoWrite class is extended to wrap the binary files in BASE64 to allow
easy marshalling into XML. No further changes are implemented.
"""

import time
import base64
import xmlrpclib
import signal

from DocXMLRPCServer import DocXMLRPCServer, DocXMLRPCRequestHandler
from nanowrite import NanoWrite


class VerifyingDocXMLRPCServer(DocXMLRPCServer):
    """
    This class implements a documented XML-RPC server which requires authentication.

    Additionally, this class allows graceful shutdown on signals
    """
    def __init__(self, users_auth, addr, timeout=0.5, *args, **kargs):
        # we use an inner class so that we can call out to the
        # authenticate method
        class VerifyingRequestHandler(DocXMLRPCRequestHandler):
            def parse_request(myself):
                # first, call the original implementation which returns
                # True if all OK so far
                if DocXMLRPCRequestHandler.parse_request(myself):
                    # next we authenticate
                    if self.authenticate(myself.headers):
                        return True
                    else:
                        # if authentication fails, tell the client
                        myself.send_error(401, 'Authentication failed')
                return False
        self._users_auth = users_auth
        self._finished = False
        self.timeout = timeout
        DocXMLRPCServer.__init__(self, addr, *args, requestHandler=VerifyingRequestHandler, **kargs)

    def authenticate(self, headers):
        # If no user was specified, don't require authentication
        if self._users_auth is None or len(self._users_auth) == 0:
            return True

        # We need an authentication
        if not 'Authorization' in headers:
            return False

        (basic, _, encoded) = headers.get('Authorization').partition(' ')

        assert basic == 'Basic', 'Only basic authentication supported'
        (username, _, password) = base64.b64decode(encoded).partition(':')

        # Check if username is valid
        if username in self._users_auth and password == self._users_auth[username]:
            return True

        # User was not authenticated
        return False

    def register_shutdown_signal(self, signum):
        signal.signal(signum, self._signal_handler)

    def _signal_handler(self, signum, frame):
        self.shutdown()

    def shutdown(self):
        self._finished = True

    def serve_forever(self):
        self._finished = False
        while not self._finished:
            server.handle_request()


class NanoWriteRPC(NanoWrite):
    def __init__(self, *args, **nargs):
        NanoWrite.__init__(self, *args, **nargs)
        pass

    def get_camera_picture(self):
        """
        Get a camera picture as BASE64 encoded tiff file.

        This is implemented via the mini gwl command window.

        @note: This requires that the camera is actually enabled. Otherwise NanoWrite just hangs...

        @return: The BASE64 encoded tif file
        @rtype: str
        """
        meta, pic = NanoWrite.get_camera_picture(self)
        return meta, xmlrpclib.Binary(pic)

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

        return {key: xmlrpclib.Binary(value) for key, value in results.items()}

if __name__ == '__main__':
    user_auth = {'user': 'password'}
    server = VerifyingDocXMLRPCServer(user_auth, ('', 60000), logRequests=1, allow_none=True)
    server.register_introspection_functions()
    server.register_instance(NanoWriteRPC())
    server.register_shutdown_signal(signal.SIGINT)

    print time.asctime(), 'Server starting'
    server.serve_forever()
    print time.asctime(), 'Server finishing'
