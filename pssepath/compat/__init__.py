import sys
from contextlib import contextmanager

try:
    # Py2
    import _winreg as winreg
except ImportError:
    # Py3
    import winreg

py_major_version = sys.version_info[0]


if py_major_version == 3:
    from ._compat3 import compat_input, simple_print  # noqa: F401
else:
    from ._compat2 import compat_input, simple_print  # noqa: F401


if py_major_version == 3:
    open_hkey_ctxmg = winreg.OpenKey
else:
    @contextmanager
    def open_hkey_ctxmg(*args, **kwargs):
        key = winreg.OpenKey(*args, **kwargs)
        try:
            yield key
        finally:
            winreg.CloseKey(key)
