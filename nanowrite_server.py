import time
import base64

from DocXMLRPCServer import DocXMLRPCServer
from nanowrite import NanoWrite

class NanoWriteRPC(NanoWrite):
    def __init__(self, *args, **nargs):
        NanoWrite.__init__(self, *args, **nargs)

    def get_camera_picture(self):
        meta, pic = NanoWrite.get_camera_picture(self)
        return meta, base64.encodestring(pic)

    def execute_complex_gwl_files(self, start_name, gwl_files, readback_files=None):
        results = NanoWrite.execute_complex_gwl_files(self, start_name, gwl_files, readback_files)

        return {key: base64.encodestring(value) for key, value in results.items()}

if __name__ == '__main__':
    server = DocXMLRPCServer(('', 60000), logRequests=1, allow_none=True)
    server.register_introspection_functions()
    server.register_instance(NanoWriteRPC())

    print time.asctime(), 'Server starting'
    server.serve_forever()
    print time.asctime(), 'Server finishing'
