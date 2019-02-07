"""
"""
import win32com.client

from . import ras500
from .. import hecrasgeometry
from ..runtime import Runtime


class RASController(object):
    """ """

    def __init__(self, filename=None):
        super(RASController, self).__init__()
        try:
            self._rc = win32com.client.DispatchEx(
                "{0}.HECRASController".format(self._ras_version))
            # self._rc = win32com.client.DispatchWithEvents(
            #    "{0}.HECRASController".format(ras_version), ras.RASEvents)
            self._events = win32com.client.WithEvents(self._rc,
                                                      self._ras.RASEvents)
        except Exception:
            msg = "{0}.HECRASController not found.".format(self._ras_version)
            raise ImportError(msg)

        self._error = 'Not available in version "{0}" of controller'.format(
            self._ras_version)
        self._filename = filename
        self._runtime = Runtime(self)

        if filename:
            self.Project_Open(filename)

    def __enter__(self):
        """ """
        self.ShowRas()
        return self

    def __exit__(self, type, value, traceback):
        """ """
        self.close()

    def pause(self, time):
        """ """
        self._runtime.pause(time)

    def runtime(self):
        """ """
        return self._runtime

    def version(self):
        """ """
        return self._ras_version

    def close(self):
        """ """
        self._runtime.close()


class RAS500(RASController, ras500.Controller):
    """HEC-RAS Controller version RAS500.

    Parameters
    ----------
    filename : str
        path to a HEC-RAS project file to open (*.prj).
    """

    def __init__(self, filename=None, exe_version=None):
        if exe_version is None:
            version = '500'
        else:
            version = str(exe_version).replace('.', '')
        self._ras_version = 'RAS' + version
        self._ras = ras500
        self._geometry = hecrasgeometry.RAS500(exe_version=exe_version)
        super(RAS500, self).__init__(filename)
