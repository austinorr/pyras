"""
HEC-RAS Controller
==================
"""

import os

import win32api
from win32api import GetFileVersionInfo, LOWORD, HIWORD
import win32con
from six import PY3

from . import hecrascontroller


def get_controller(exe_version=None):
    if exe_version is None:
        exe_version = max(get_installed_ras_versions())

    version = str(exe_version).replace('.', '')
    controller_name = 'RAS' + _get_le_controller_version(version)

    rc = getattr(hecrascontroller, controller_name)

    return rc(exe_version=version)


def get_installed_ras_versions():
    """
    """
    ldic = _get_registered_typelibs()

    available_versions = []

    for dic in ldic:
        fname = dic['filename']
        if os.path.isfile(fname):
            available_versions.append(dic['version'])
    return available_versions


def kill_ras():
    """ """
    import subprocess

    ras_process_string = 'ras.exe'
    proc = subprocess.Popen('TASKLIST /FO "CSV"', stdout=subprocess.PIPE)
    if PY3:
        tasklist = proc.stdout.read().decode('utf-8').split('\n')
    else:
        tasklist = proc.stdout.read().split('\n')
    tasks = []
    pids = []
    for line in tasklist:
        l = line.lower()
        if ras_process_string in l:
            items = l.split(',')
            tasks.append(items)
            pids.append(int(eval(items[1])))

    for pid in pids:
        try:
            # FIXME:
            os.system('TASKKILL /PID {0} /F >nul'.format(pid))
        except Exception as e:
            print(e)


def _get_controller_versions():
    controller_versions = []
    for attr in dir(hecrascontroller):
        if 'RAS' in attr:
            _, ver = attr.split('RAS')
            controller_versions.append(ver)

    return controller_versions


def _get_le_controller_version(version):
    """
    version : str
        version string e.g., 410
    """
    avail_controller_versions = sorted(_get_controller_versions())[::-1]
    for avail_ver in avail_controller_versions:
        if avail_ver <= version:
            return avail_ver
    raise ValueError('version not found: ', version)


def _get_typelib_info(keyid, version):
    """
    adapted from pywin32

    # Copyright (c) 1996-2008, Greg Stein and Mark Hammond.
    """
    collected = []
    help_path = ""
    key = win32api.RegOpenKey(win32con.HKEY_CLASSES_ROOT,
                              "TypeLib\\%s\\%s" % (keyid, version))
    try:
        num = 0
        while 1:
            try:
                sub_key = win32api.RegEnumKey(key, num)
            except win32api.error:
                break
            h_sub_key = win32api.RegOpenKey(key, sub_key)
            try:
                value, typ = win32api.RegQueryValueEx(h_sub_key, None)
                if typ == win32con.REG_EXPAND_SZ:
                    value = win32api.ExpandEnvironmentStrings(value)
            except win32api.error:
                value = ""
            if sub_key == "HELPDIR":
                help_path = value
            elif sub_key == "Flags":
                flags = value
            else:
                try:
                    lcid = int(sub_key)
                    lcidkey = win32api.RegOpenKey(key, sub_key)
                    # Enumerate the platforms
                    lcidnum = 0
                    while 1:
                        try:
                            platform = win32api.RegEnumKey(lcidkey, lcidnum)
                        except win32api.error:
                            break
                        try:
                            hplatform = win32api.RegOpenKey(lcidkey, platform)
                            fname, typ = win32api.RegQueryValueEx(
                                hplatform, None)
                            if typ == win32con.REG_EXPAND_SZ:
                                fname = win32api.ExpandEnvironmentStrings(
                                    fname)
                        except win32api.error:
                            fname = ""
                        collected.append((lcid, platform, fname))
                        lcidnum = lcidnum + 1
                    win32api.RegCloseKey(lcidkey)
                except ValueError:
                    pass
            num = num + 1
    finally:
        win32api.RegCloseKey(key)

    return fname, lcid


def _get_ras_version_number(filename):
    try:
        info = GetFileVersionInfo(filename, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        return HIWORD(ms), LOWORD(ms), LOWORD(ls)  # major, minor, revision
    except:
        return 0, 0, 0


def _get_registered_typelibs(match='HEC River Analysis System'):
    """
    adapted from pywin32

    # Copyright (c) 1996-2008, Greg Stein and Mark Hammond.
    """
    # Explicit lookup in the registry.
    result = []
    key = win32api.RegOpenKey(win32con.HKEY_CLASSES_ROOT, "TypeLib")
    try:
        num = 0
        while 1:
            try:
                key_name = win32api.RegEnumKey(key, num)
            except win32api.error:
                break
            # Enumerate all version info
            sub_key = win32api.RegOpenKey(key, key_name)
            name = None
            try:
                sub_num = -1
                best_version = 0.0
                while 1:

                    try:
                        sub_num = sub_num + 1
                        version_str = win32api.RegEnumKey(sub_key, sub_num)

                    except win32api.error:
                        break

                    try:
                        version_flt = float(version_str)
                    except ValueError:
                        version_flt = 0  # ????
                    if version_flt > best_version:
                        best_version = version_flt
                        name = win32api.RegQueryValue(sub_key, version_str)

            finally:
                win32api.RegCloseKey(sub_key)
            if name is not None and match in name:
                fname, lcid = _get_typelib_info(key_name, version_str)

                # Split version
                major, minor, rev = _get_ras_version_number(fname)

                dct = {
                    'name': name,
                    'filename': fname,
                    'lcid': lcid,
                    'version': ".".join(map(str, [major, minor, rev])),
                    'major': int(major),
                    'minor': int(minor),
                    'revision': int(rev),
                }

                result.append(dct)
            num = num + 1
    finally:
        win32api.RegCloseKey(key)
    return result


class HECRASImportError(Exception):

    def __init__(self, message=''):
        msg = '"HEC River Analysis System" type library not found. ' \
              'Please install HEC-RAS'
        if message:
            msg = message

        # Call the base class constructor with the parameters it needs
        super(HECRASImportError, self).__init__(msg)


# %%
# kill_ras()
# __available_versions__ = get_available_versions()
