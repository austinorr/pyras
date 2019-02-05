"""
"""
import win32com.client

from . import ras41
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


class RAS41(RASGeometry, ras41.Geometry):
    """HEC-RAS Geometry version RAS41."""

    def __init__(self):
        self._ras_version = 'RAS41'
        self._ras = ras41
        super(RAS41, self).__init__()


class RAS500(RASGeometry, ras500.Geometry):
    """HEC-RAS Geometry version RAS500."""

    def __init__(self):
        self._ras_version = 'RAS500'
        self._ras = ras500
        super(RAS500, self).__init__()


class RAS503(RASGeometry, ras500.Geometry):
    """HEC-RAS Geometry version RAS500."""

    def __init__(self):
        self._ras_version = 'RAS503'
        self._ras = ras500
        super(RAS503, self).__init__()
