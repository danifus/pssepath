from __future__ import with_statement

import platform
import struct
import os
import sys
from functools import wraps

try:
    # Py2
    import _winreg as winreg
except ImportError:
    # Py3
    import winreg


def memoize(fn):
    """
    args and kwargs to fn must be hashable.
    """
    cache = {}

    @wraps(fn)
    def wrap(*args, **kwargs):
        if args:
            key_args = frozenset(args)
        else:
            key_args = None
        if kwargs:
            key_kwargs = frozenset(kwargs.items())
        else:
            key_kwargs = None

        key = (key_args, key_kwargs)
        if key not in cache:
            cache[key] = fn(*args, **kwargs)
        return cache[key]

    return wrap


def run_once(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not getattr(fn, "hasrun", False):
            setattr(fn, "hasrun", True)
            fn(*args, **kwargs)

    return wrapped


def get_python_ver():
    """Returns (python_version, nbits) eg. ("2.7", "32bit")"""
    py_ver = "%s.%s" % sys.version_info[:2]
    return py_ver, platform.architecture()[0]


# winreg helpers:
def get_reg_value(key, value_name):
    try:
        return winreg.QueryValueEx(key, value_name)[0]
    except WindowsError:
        return None


def enum_reg_keys(key):
    i = 0
    while True:
        try:
            yield winreg.EnumKey(key, i)
        except OSError:
            break
        i += 1


# pyc magic number (the code that hints what python can read the compiled
# python file) helpers:
PYC_MAGIC_NUMS = {
    20121: "1.5.x",
    50428: "1.6",
    50823: "2.0.x",
    60202: "2.1.x",
    60717: "2.2",
    62011: "2.3a0",
    62021: "2.3a0",
    62011: "2.3a0",
    62041: "2.4a0",
    62051: "2.4a3",
    62061: "2.4b1",
    62071: "2.5a0",
    62081: "2.5a0",
    62091: "2.5a0",
    62092: "2.5a0",
    62101: "2.5b3",
    62111: "2.5b3",
    62121: "2.5c1",
    62131: "2.5c2",
    62151: "2.6a0",
    62161: "2.6a1",
    62171: "2.7a0",
    62181: "2.7a0",
    62191: "2.7a0",
    62201: "2.7a0",
    62211: "2.7a0",
    3000: "3.0",
    3010: "3.0",
    3020: "3.0",
    3030: "3.0",
    3040: "3.0",
    3050: "3.0",
    3060: "3.0",
    3061: "3.0",
    3071: "3.0",
    3081: "3.0",
    3091: "3.0",
    3101: "3.0",
    3103: "3.0",
    3111: "3.0a4",
    3131: "3.0a5",
    3141: "3.1a0",
    3151: "3.1a0",
    3160: "3.2a0",
    3170: "3.2a1",
    3180: "3.2a2",
    3190: "3.3a0",
    3200: "3.3a0",
    3210: "3.3a0",
    3220: "3.3a1",
    3230: "3.3a4",
    3250: "3.4a1",
    3260: "3.4a1",
    3270: "3.4a1",
    3280: "3.4a1",
    3290: "3.4a4",
    3300: "3.4a4",
    3310: "3.4rc2",
    3320: "3.5a0",
    3330: "3.5b1",
    3340: "3.5b2",
    3350: "3.5b2",
    3360: "3.6a0",
    3361: "3.6a0",
    3370: "3.6a1",
    3371: "3.6a1",
    3372: "3.6a1",
    3373: "3.6b1",
    3375: "3.6b1",
    3376: "3.6b1",
    3377: "3.6b1",
    3378: "3.6b2",
    3379: "3.6rc1",
    3390: "3.7a1",
    3391: "3.7a2",
    3392: "3.7a4",
    3393: "3.7b1",
    3394: "3.7b5",
    3400: "3.8a1",
    3401: "3.8a1",
    3410: "3.8a1",
    3411: "3.8b2",
    3412: "3.8b2",
    3413: "3.8b4",
    3420: "3.9a0",
    3421: "3.9a0",
    3422: "3.9a0",
    3423: "3.9a2",
    3424: "3.9a2",
    3425: "3.9a2",
}


def read_magic_number(fname):
    pyc_file = open(fname, "rb")
    magic_bytes = pyc_file.read(2)
    pyc_file.close()
    magic = struct.unpack("<H", magic_bytes)[0]
    return PYC_MAGIC_NUMS[magic]


# Windows program files 32bit vs 64bit helpers:
def is_win64():
    return "PROGRAMFILES(X86)" in os.environ


def get_programfiles_32():
    if is_win64():
        return os.environ["PROGRAMFILES(X86)"]
    else:
        # "PROGRAMFILES" will return 'Program Files (x86)' on 32bit
        # python and 'Program Files' on 64bit.
        return os.environ["PROGRAMFILES"]


def get_programfiles_64():
    if is_win64():
        return os.environ["PROGRAMW6432"]
    else:
        return None
