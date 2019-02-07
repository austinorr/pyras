"""
"""
import win32com.client

from . import ras500


class RASGeometry(object):

    def __init__(self):
        super(RASGeometry, self).__init__()
        try:
            self._geometry = win32com.client.DispatchEx(
                "{0}.HECRASGeometry".format(self._ras_version))
        except Exception:
            msg = "{0}.HECRASGeometry not found.".format(self._ras_version)
            raise ImportError(msg)


class RAS500(RASGeometry, ras500.Geometry):
    """HEC-RAS Geometry version RAS500."""

    def __init__(self, exe_version=None):
        if exe_version is None:
            version = '500'
        else:
            version = str(exe_version).replace('.', '')
        self._ras_version = 'RAS' + version
        self._ras = ras500
        super(RAS500, self).__init__()
