"""
"""
import win32com.client

from . import ras500, ras41


class RASGeometry(object):

    def __init__(self):
        super(RASGeometry, self).__init__()
        try:
            self._geometry = win32com.client.DispatchEx(
                "{0}.HECRASGeometry".format(self._ras_version))
        except Exception:
            msg = "{0}.HECRASGeometry not found.".format(self._ras_version)
            raise ImportError(msg)


class RAS41(RASGeometry, ras41.Geometry):
    """HEC-RAS Geometry version RAS41."""

    def __init__(self, progid='RAS41'):
        self._ras_version = progid
        self._ras = ras41
        super(RAS41, self).__init__()


class RAS500(RASGeometry, ras500.Geometry):
    """HEC-RAS Geometry version RAS500."""

    def __init__(self, progid='RAS500'):
        self._ras_version = progid
        self._ras = ras500
        super(RAS500, self).__init__()
