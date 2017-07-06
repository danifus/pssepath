import logging
import os
import sys
from functools import wraps
from textwrap import dedent
import _winreg


logger = logging.getLogger(__name__)


psse_version = None
req_python_exec = None
initialized = False


class PsseImportError(Exception):
    pass


def check_psspy_already_in_path():
    """Return True if psspy.pyc in the sys.path and os.environ['PATH'] dirs.

    Otherwise, print a warning message and return False so the paths get
    reconfigured.
    """
    syspath = find_file_on_path('psspy.pyc', sys.path)

    if syspath:
        # file in one of the files on the sys.path (python's path) list.
        envpaths = os.environ['PATH'].split(';')
        envpath = find_file_on_path('psspy.pyc', envpaths)
        if envpath:
            # lets check to see that PSSBIN is also on the windows path. If it
            # isn't, psspy will not function properly.
            if syspath == envpath:
                return True
            else:
                log_pathmismatch_warning(syspath, envpath)
        else:
            log_path_noenviron_warning()

    return False


def check_initialized(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if initialized:
            logger.info("psspath has already added PSSBIN to the system, continuing.")
        elif check_psspy_already_in_path():
            check_already_present_psse()
            logger.info("PSSBIN already in path, adding PSSBIN from pssepath skipped.")
            set_status(initialized=True)
        else:
            fn(*args, **kwargs)
    return wrapped


def run_once(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        if not getattr(fn, 'hasrun', False):
            setattr(fn, 'hasrun', True)
            fn(*args, **kwargs)
    return wrapped


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


@run_once
def log_path_noenviron_warning():
    logger.warning(
        dedent("""\
       pssepath: Warning - PSSBIN found on sys.path, but not os.environ['PATH'].
                           Running pssepath.add_pssepath() will reconfigure.

                 Running pssepath.add_pssepath() will attempt to reconfigure
                 your paths for you.  If you wish to find the root cause of
                 this message, check your Python scripts to see if they set up
                 sys.path or os.environ['PATH'] and remove that code.  If the
                 scripts do not attempt to configure these variables, you may
                 need to check your Windows PATH variables from windows, as they
                 may have been configured there.
                 """)
    )


@run_once
def log_pathmismatch_warning(syspath, envpath):
    logger.warning(
        dedent("""\
       pssepath: Warning - PSSBIN path mismatch.
                           Running pssepath.add_pssepath() will reconfigure.

                 Two different paths for PSSBIN were found in sys.path and
                 os.environ[PATH].

                 sys.path:           %s
                 os.environ['PATH']: %s

                 Running pssepath.add_pssepath() will attempt to reconfigure
                 your paths for you.  If you wish to find the root cause of
                 this message, check your Python scripts to see if they set up
                 sys.path or os.environ['PATH'] and remove that code.  If the
                 scripts do not attempt to configure these variables, you may
                 need to check your Windows PATH variables from windows, as they
                 may have been configured there.
                 """) % (syspath, envpath)
    )


def add_dir_to_path(psse_ver, psse_path):
    """Add psse_path to 'sys.path' and 'os.environ['PATH'].

    This affects the os and sys modules, thus these side-effects are global.
    Adds them to the start of the path variables so that they are always used
    in preference.

    This is all side-effects which is not the prettiest.
    """
    sys.path.insert(0, psse_path)
    os.environ['PATH'] = psse_path + ';' + os.environ['PATH']

    if psse_ver >= 34:
        # Also add the PSSBIN dir
        pssebin_dir = os.path.join(os.path.dirname(psse_path), 'PSSBIN')
        sys.path.insert(0, pssebin_dir)
        os.environ['PATH'] = pssebin_dir + ';' + os.environ['PATH']


def rem_dir_from_path(psse_path):
    """Remove psse_path from 'sys.path' and 'os.environ['PATH'].

    list.remove(bla) will always remove the first instance of bla from the
    list. Thus this will reverse any changes done by add_dir_to_path().
    """

    if psse_path in sys.path:
        sys.path.remove(psse_path)
    if psse_path in os.environ['PATH']:
        sys_paths = os.environ['PATH'].split(';')
        sys_paths.remove(psse_path)
        os.environ['PATH'] = ';'.join(sys_paths)


@memoize
def get_pssbin_paths_dict():
    if is_win64():
        pti_key = _winreg.OpenKey(
            _winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Wow6432Node\\PTI')
    else:
        pti_key = _winreg.OpenKey(
            _winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\PTI')

    pssbin_paths = {}

    sub_key_cnt = _winreg.QueryInfoKey(pti_key)[0]
    for i in range(sub_key_cnt):
        sub_key = _winreg.EnumKey(pti_key, i)
        try:
            ver_key = _winreg.OpenKey(pti_key, sub_key + '\\Product Paths')
        except WindowsError:
            pass
        else:
            # Version num is the last 2 digits of the subkey
            version_num = int(sub_key[-2:])
            path = _winreg.QueryValueEx(ver_key, 'PsseExePath')[0]
            pssbin_paths[version_num] = path

    if not len(pssbin_paths):
        raise PsseImportError('No installs of PSSE found.')

    _winreg.CloseKey(ver_key)
    _winreg.CloseKey(pti_key)
    return pssbin_paths


@memoize
def get_psse_locations_dict():
    psspy_dirs = {}
    pssbin_paths = get_pssbin_paths_dict()
    for psse_ver, pssbin in pssbin_paths.items():
        pyvers_and_psspy_paths = get_required_python_ver_and_paths(psse_ver, pssbin)
        for pyver, psspy_path in pyvers_and_psspy_paths:
            psspy_dirs[(psse_ver, pyver)] = psspy_path
    return psspy_dirs


def get_psspy_paths_for_psse_ver(psse_ver):
    """
    Return a dict of pyver -> psspy path for the supplied psse_ver.
    """
    pyver_to_path = {}
    psspy_paths = get_psse_locations_dict()
    for (psse_ver, pyver), psspy_path in psspy_paths.items():
        pyver_to_path[pyver] = psspy_path
    return pyver_to_path


def check_to_raise_compat_python_error(psse_version):
    if not module_settings['ignore_python_mismatch']:
        psspy_paths = get_psse_locations_dict()
        possible_pyvers = []
        for psse_ver, pyver in psspy_paths.keys():
            if psse_ver == psse_version:
                possible_pyvers.append(pyver)

        if sys.winver not in possible_pyvers:
            pyver_text = ' or '.join(possible_pyvers)
            raise PsseImportError(
                "Current Python and PSSE version "
                "incompatible.\n\n"
                "PSSE %s requires Python %s to run.\n"
                "Current Python is Version %s.\n" % (
                    psse_version, pyver_text, sys.winver)
            )


@check_initialized
def add_pssepath(pref_psse_ver=None):
    """Add the PSSBIN path to the required locations.

    Try to import the requested version of PSSE. If the requested version
    doesn't work, raise an exception. By default, import the latest version.
    """
    psspy_paths = get_psse_locations_dict()
    current_pyver = sys.winver

    if pref_psse_ver:
        available_psse_versions_set = set()
        for psse_ver, pyver in psspy_paths.keys():
            available_psse_versions_set.add(psse_ver)
        available_psse_versions = sorted(available_psse_versions_set)

        if pref_psse_ver in available_psse_versions:
            selected_psse_ver = pref_psse_ver
            check_to_raise_compat_python_error(selected_psse_ver)
        else:
            if len(available_psse_versions) == 1:
                ver_string = (
                    'the installed version: %s' % (available_psse_versions[0],)
                )
            else:
                psses = ', '.join([str(x) for x in available_psse_versions])
                ver_string = 'an installed version: %s' % psses

            raise PsseImportError(
                'Attempted to initialize PSSE version %s but it was not present.\n'
                'Let pssepath select the latest version by not specifying a '
                'version when\n'
                'calling "pssepath.add_pssepath()", or select %s'
                % (pref_psse_ver, ver_string)
            )
    else:
        # automatically select the most recent version.
        rev_sorted_vers = sorted(psspy_paths.keys(), reverse=True)
        selected_psse_ver = None
        # If 'ignore_python_mismatch = True' this will always return
        # the most recent version as check_to_raise_compat_python_error won't
        # raise an error.
        for psse_ver, pyver in rev_sorted_vers:
            if pyver != current_pyver:
                continue
            try:
                check_to_raise_compat_python_error(psse_ver)
            except PsseImportError:
                pass
            else:
                selected_psse_ver = psse_ver
                break
        if not selected_psse_ver:
            raise PsseImportError(
                'No installed PSSE versions are compatible '
                'with the running version of Python (%s)\n' % (sys.winver,)
            )

    selected_path = psspy_paths[(selected_psse_ver, current_pyver)]
    add_dir_to_path(selected_psse_ver, selected_path)
    req_python_exec = os.path.join(
        get_python_locations_dict()[current_pyver], 'python.exe')
    set_status(
        req_python_exec=req_python_exec, psse_version=selected_psse_ver,
        initialized=True)


@check_initialized
def select_pssepath():
    """Produce a prompt to select the version of PSSE"""

    print 'Please select from the available PSSE installs:\n'
    print_psse_selection()
    psspy_paths = get_psse_locations_dict()
    versions = sorted(psspy_paths.keys())
    while True:
        try:
            user_input = int(
                raw_input('Enter a number from the above PSSE installations: ')
            )
        except ValueError:
            continue

        if 0 < user_input <= len(psspy_paths):
            # Less one due to zero based vs 1-based (len)
            break

    selected_key = versions[user_input - 1]
    psse_ver, pyver = selected_key
    selected_path = psspy_paths[selected_key]
    check_to_raise_compat_python_error(versions[user_input - 1])
    add_dir_to_path(psse_ver, selected_path)
    req_python_ver = get_required_python_for_psspy_in(selected_path)
    req_python_exec = os.path.join(
        get_python_locations_dict()[req_python_ver], 'python.exe')
    set_status(
        req_python_exec=req_python_exec, psse_version=versions[user_input - 1],
        initialized=True)


def print_psse_selection():

    psspy_paths = get_psse_locations_dict()
    versions = sorted(psspy_paths.keys())
    for i, (psse_ver, pyver) in enumerate(versions):
        python_str = 'Requires Python %s' % (pyver)
        if pyver == sys.winver:
            python_str += ' (Current running Python)'
        elif pyver in get_python_locations_dict().keys():
            python_str += ' (Installed, not current running version.)'
        print ('  %i. PSSE Version %d\n'
               '      %s' % (i+1, psse_ver, python_str))


# ============== Python version detection
def read_magic_number(fname):
    pyc_file = open(fname, 'rb')
    magic = pyc_file.read(2)
    pyc_file.close()
    return int(magic[::-1].encode('hex'), 16)


def find_file_on_path(fname, dir_checklist=None):
    """Return the first file on the path which matches fname.

    By default, this function will search sys.path for a matching file. This
    can be overridden by passing a list of dirs to be checked in as
    'dir_checklist'.
    """
    if not dir_checklist:
        dir_checklist = sys.path

    for path_dir in dir_checklist:
        potential_file = os.path.join(path_dir, fname)
        if os.path.isfile(potential_file):
            return potential_file


def get_required_python_ver_and_paths(psse_ver, pssbin):
    """
    Return a list of [(pyver, psspy_dir), ...] or [] if no path.
    """
    if psse_ver < 34:
        pyvers_and_paths = get_required_python_ver_psse_33_and_older(pssbin)
    elif psse_ver >= 34:
        pyvers_and_paths = get_required_python_ver_psse_34_and_newer(pssbin)

    return pyvers_and_paths


def get_required_python_for_psspy_in(psspy_dir):
    """
    Returns python version of psspy.pyc or None if no psspy.pyc file in this dir.
    """
    probable_pyc = os.path.join(psspy_dir, 'psspy.pyc')
    if not os.path.isfile(probable_pyc):
        return None

    magic = read_magic_number(probable_pyc)
    # only the first 3 digits are important (2.x etc)
    return PYC_MAGIC_NUMS[magic][:3]


def get_required_python_ver_psse_33_and_older(pssbin):
    """
    Returns a list of possible python versions and the path which has psspy.pyc
    on it (of which 33 and lower only supports one version).
    """
    pyver = get_required_python_for_psspy_in(pssbin)
    if pyver is not None:
        return [(pyver, pssbin)]
    return []


def find_psse_pydirs(psse_base_dir):
    pyvers_and_paths = []
    for fpath in os.listdir(psse_base_dir):
        psspy_dir = os.path.join(psse_base_dir, fpath)
        if (os.path.isdir(psspy_dir) and fpath.upper().startswith('PSSPY')):
            pyver = get_required_python_for_psspy_in(psspy_dir)
            if pyver is not None:
                pyvers_and_paths.append((pyver, psspy_dir))
    return pyvers_and_paths


def get_required_python_ver_psse_34_and_newer(pssbin):
    """
    Returns a list of possible python versions. PSSE34 and above support
    multiple python verions.
    """
    # Get the directory above the PSSBIN dir.
    psse_base_dir = os.path.dirname(pssbin)
    pyvers_and_paths = find_psse_pydirs(psse_base_dir)
    return pyvers_and_paths


@memoize
def get_python_locations_dict():
    if is_win64():
        python_key = _winreg.OpenKey(
            _winreg.HKEY_LOCAL_MACHINE,
            'SOFTWARE\\Wow6432Node\\Python\\PythonCore')
    else:
        python_key = _winreg.OpenKey(
            _winreg.HKEY_LOCAL_MACHINE,
            'SOFTWARE\\Python\\PythonCore')

    python_paths = {}

    sub_key_cnt = _winreg.QueryInfoKey(python_key)[0]
    for i in range(sub_key_cnt):
        sub_key = _winreg.EnumKey(python_key, i)
        try:
            ver_key = _winreg.OpenKey(python_key, sub_key + '\\InstallPath')
        except WindowsError:
            pass
        else:
            # Version num is the last 2 digits of the subkey
            version_num = sub_key
            path = _winreg.QueryValue(ver_key, None)
            python_paths[version_num] = path

    if not len(python_paths):
        raise PsseImportError(
            'No installs of Python found... wait how are you running this...')

    _winreg.CloseKey(ver_key)
    _winreg.CloseKey(python_key)
    return python_paths


@memoize
def get_os_architecture():

    try:
        # If this does not raise an exception, we are on a 64 bit windows.
        _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Wow6432Node')
        os_arch = "Win64"
    except WindowsError:
        os_arch = "Win32"

    return os_arch


def is_win64():
    if get_os_architecture == "Win64":
        return True
    return False


PYC_MAGIC_NUMS = {
    20121: '1.5.x',
    50428: '1.6',
    50823: '2.0.x',
    60202: '2.1.x',
    60717: '2.2',
    62011: '2.3a0',
    62021: '2.3a0',
    62011: '2.3a0',
    62041: '2.4a0',
    62051: '2.4a3',
    62061: '2.4b1',
    62071: '2.5a0',
    62081: '2.5a0',
    62091: '2.5a0',
    62092: '2.5a0',
    62101: '2.5b3',
    62111: '2.5b3',
    62121: '2.5c1',
    62131: '2.5c2',
    62151: '2.6a0',
    62161: '2.6a1',
    62171: '2.7a0',
    62181: '2.7a0',
    62191: '2.7a0',
    62201: '2.7a0',
    62211: '2.7a0',

    3000: '3.0',
    3010: '3.0',
    3020: '3.0',
    3030: '3.0',
    3040: '3.0',
    3050: '3.0',
    3060: '3.0',
    3061: '3.0',
    3071: '3.0',
    3081: '3.0',
    3091: '3.0',
    3101: '3.0',
    3103: '3.0',
    3111: '3.0a4',
    3131: '3.0a5',
    3141: '3.1a0',
    3151: '3.1a0',
    3160: '3.2a0',
    3170: '3.2a1',
    3180: '3.2a2',
    3190: '3.3a0',
    3200: '3.3a0',
    3210: '3.3a0',
    3220: '3.3a1',
    3230: '3.3a4',
    3250: '3.4a1',
    3260: '3.4a1',
    3270: '3.4a1',
    3280: '3.4a1',
    3290: '3.4a4',
    3300: '3.4a4',
    3310: '3.4rc2',
    3320: '3.5a0',
    3330: '3.5b1',
    3340: '3.5b2',
    3350: '3.5b2',
    3360: '3.6a0',
    3361: '3.6a0',
    3370: '3.6a1',
    3371: '3.6a1',
    3372: '3.6a1',
    3373: '3.6b1',
    3375: '3.6b1',
    3376: '3.6b1',
    3377: '3.6b1',
    3378: '3.6b2',
    3379: '3.6rc1',
    3390: '3.7a0',
}


module_settings = {
    'ignore_python_mismatch': False,
}


def set_ignore_python_mismatch(value):
    module_settings['ignore_python_mismatch'] = value


def set_status(**kwargs):

    global psse_version, req_python_exec, initialized

    if 'psse_version' in kwargs:
        psse_version = kwargs['psse_version']

    if 'req_python_exec' in kwargs:
        req_python_exec = kwargs['req_python_exec']

    if 'initialized' in kwargs:
        initialized = kwargs['initialized']


def check_already_present_psse():
    """
    Raises errors if the already present PSSE looks to be misconfigured.
    """
    if not initialized and check_psspy_already_in_path():

        pypaths = set()
        pyvers = []
        # need to find the required python for this version
        for possible_psspy_dir in sys.path:
            if 'PSSBIN' in possible_psspy_dir or 'PSSPY' in possible_psspy_dir:

                pyver = get_required_python_for_psspy_in(possible_psspy_dir)
                if pyver is not None:
                    abs_path = os.path.abspath(possible_psspy_dir)
                    if abs_path not in pypaths:
                        pyvers.append(pyver)

        if len(pyvers) > 1:
            paths = '\n'.join(list(pypaths))
            raise PsseImportError(
                "WARNING: your PATH has been configured to make multiple "
                "psspy.pyc files availble. This may lead to loading the "
                "incorrect copy. Please only have one version available. "
                "\nThe multiple paths are:\n%s" % paths
            )

        req_python = pyvers[0]

        if req_python != sys.winver:
            raise PsseImportError(
                "WARNING: you have started a Python %s session when the\n"
                "version required by the PSSE available in your path is\n"
                "Python %s.\n"
                "Either use the required version of Python or,\n"
                "if you have another version of PSSE installed, change your\n"
                "PATH settings to point at the other install.\n\n"
                "Run '%s -m pssepath' for more info about the versions\n"
                "installed on your system.\n\n" % (
                    sys.winver, req_python, sys.executable)
            )
            try:
                req_python_exec = os.path.join(
                    get_python_locations_dict()[req_python], 'python.exe')
                set_status(req_python_exec=req_python_exec)
            except KeyError:
                # Very unlikely
                # Don't have the required version of python to run this version of
                # psse.  Something is not right...
                raise PsseImportError(
                    "Required version of python (%s) not located in registry.\n"
                    % (req_python,)
                )
        else:
            set_status(req_python_exec=sys.executable)

        # a psse path is already present and looks ok.
        set_status(initialized=True)


if __name__ == "__main__":
    # print the available psse installs.
    logging.basicConfig(format="%(message)s", level=logging.INFO)
    check_already_present_psse()
    print 'Found the following PSSE versions installed:\n'
    print_psse_selection()
    raw_input("Press Enter to continue...")
