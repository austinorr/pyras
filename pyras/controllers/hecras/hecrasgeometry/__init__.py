"""
"""
import win32com.client

from . import ras410
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


class RAS410(RASGeometry, ras410.Geometry):
    """HEC-RAS Geometry version RAS41."""

    def __init__(self, version=None):
        if version is None:
            self._ras_version = 'RAS410'
        else:
            self._ras_version = 'RAS' + str(version).replace('.', '')[:-1]
        self._ras = ras410
        super(RAS410, self).__init__()


class RAS500(RASGeometry, ras500.Geometry):
    """HEC-RAS Geometry version RAS500."""

    def __init__(self, version=None):
        if version is None:
            self._ras_version = 'RAS500'
        else:
            self._ras_version = 'RAS' + str(version).replace('.', '')
        self._ras = ras500
        super(RAS500, self).__init__()