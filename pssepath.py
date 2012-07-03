import os
import sys
from textwrap import dedent
import _winreg

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
                print_pathmismatch_warning(syspath, envpath)
        else:
            print_path_noenviron_warning()

    return False

def check_initialized(fn):
    def wrapped(*args, **kwargs):
        if initialized:
            print "psspath has already added PSSBIN to the system, continuing."
        elif check_psspy_already_in_path():
            print "PSSBIN already in path, adding PSSBIN from pssepath skipped."
        else:
            fn(*args, **kwargs)
    return wrapped

def run_once(fn):
    def wrapped(*args, **kwargs):
        if not getattr(fn, 'hasrun', False):
            setattr(fn, 'hasrun', True)
            fn(*args, **kwargs)
    return wrapped

@run_once
def print_path_noenviron_warning():
    print dedent("""\
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

@run_once
def print_pathmismatch_warning(syspath, envpath):
    print (dedent("""\
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
                 """) % (syspath, envpath))

def add_dir_to_path(psse_path):
    """Add psse_path to 'sys.path' and 'os.environ['PATH'].

    This affects the os and sys modules, thus these side-effects are global.

    This is all side-effects which is not the prettiest.
    """
    sys.path.append(psse_path)
    os.environ['PATH'] = os.environ['PATH'] + ';' +  psse_path

def rem_dir_from_path(psse_path):
    """Remove psse_path from 'sys.path' and 'os.environ['PATH']."""

    if psse_path in sys.path:
        sys.path.remove(psse_path)
    if psse_path in os.environ['PATH']:
        sys_paths = os.environ['PATH'].split(';')
        sys_paths.remove(psse_path)
        os.environ.update({'PATH': ';'.join(sys_paths)})

def _get_psse_locations_dict():
    pti_key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\PTI')

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

def check_to_raise_compat_python_error(psse_version):
    if not ignore_python_mismatch:
        selected_path = pssbin_paths[psse_version]
        req_python_ver = get_required_python_ver(selected_path)
        if req_python_ver != sys.winver:
            raise PsseImportError("Current Python and PSSE version "
                "incompatible.\n\n"
                "PSSE %s requires Python %s to run.\n"
                "Current Python is Version %s.\n" % (psse_version,
                    req_python_ver, sys.winver))

@check_initialized
def add_pssepath(pref_ver=None):
    """Add the PSSBIN path to the required locations.

    Try to import the requested version of PSSE. If the requested version
    doesn't work, raise an exception. By default, import the latest version.
    """
    if pref_ver:
        if pref_ver in pssbin_paths.keys():
            selected_ver = pref_ver
            check_to_raise_compat_python_error(selected_ver)
        else:
            if len(pssbin_paths) == 1:
                ver_string = ('the installed version: %s' %
                        (pssbin_paths.keys()[0],))
            else:
                psses = ', '.join([str(x) for x in pssbin_paths.keys()])
                ver_string = 'an installed version: %s' % psses

            raise PsseImportError('Attempted to initialize PSSE version %s but '
                'it was not present.\n'
                'Let pssepath select the latest version by not specifying a '
                'version when\n'
                'calling "pssepath.import_psse()", or select %s'
                % (pref_ver, ver_string))
    else:
        # automatically select the most recent version.
        rev_sorted_vers =  sorted(pssbin_paths.keys(), reverse=True)
        selected_ver = None
        # If 'ignore_python_mismatch = True' this will always return
        # the most recent version as check_to_raise_compat_python_error won't
        # raise an error.
        for ver in rev_sorted_vers:
            try:
                check_to_raise_compat_python_error(ver)
            except PsseImportError:
                pass
            else:
                selected_ver = ver
                break
        if not selected_ver:
            raise PsseImportError('No installed PSSE versions are compatible '
                    'with the running version of Python (%s)\n' % (sys.winver,))


    selected_path = pssbin_paths[selected_ver]
    add_dir_to_path(selected_path)
    global initialized, psse_version, req_python_exec
    psse_version = selected_ver
    req_python_ver = get_required_python_ver(selected_path)
    req_python_exec = os.path.join(python_paths[req_python_ver],'python.exe')
    initialized = True

@check_initialized
def select_pssepath():
    """Produce a prompt to select the version of PSSE"""

    print 'Please select from the available PSSE installs:\n'
    print_psse_selection()
    versions = sorted(pssbin_paths.keys())
    while True:
        try:
            user_input = int(raw_input('Enter a number from the above '
                                    'PSSE installations: '))
        except ValueError:
            continue

        if 0 < user_input <= len(pssbin_paths):
            # Less one due to zero based vs 1-based (len)
            break

    selected_path = pssbin_paths[versions[user_input - 1]]
    check_to_raise_compat_python_error(versions[user_input - 1])
    add_dir_to_path(selected_path)
    global initialized, psse_version, req_python_exec
    psse_version = versions[user_input - 1]
    req_python_ver = get_required_python_ver(selected_path)
    req_python_exec = os.path.join(python_paths[req_python_ver],'python.exe')
    initialized = True

def print_psse_selection():

    versions = sorted(pssbin_paths.keys())
    for i, ver in enumerate(versions):
        req_python_ver = get_required_python_ver(pssbin_paths[ver])
        python_str = 'Requires Python %s' % (req_python_ver)
        if req_python_ver == sys.winver:
            python_str += ' (Current running Python)'
        elif req_python_ver in python_paths.keys():
            python_str += ' (Installed)'
        print ('  %i. PSSE Version %d\n'
               '      %s' %(i+1, ver, python_str))

# ============== Python version detection
def read_magic_number(fname):
    pyc_file = open(fname,'rb')
    magic = pyc_file.read(2)
    pyc_file.close()
    return int(magic[::-1].encode('hex'),16)

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

def get_required_python_ver(pssbin):
    probable_pyc = os.path.join(pssbin,'psspy.pyc')
    if not os.path.isfile(probable_pyc):
        # not in the suspected dir, perhaps abnormal install.
        probable_pyc = find_file_on_path('psspy.pyc')

    magic = read_magic_number(probable_pyc)
    # only the first 3 digits are important (2.x etc)
    return pyc_magic_nums[magic][:3]

def _get_python_locations_dict():
    python_key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE,
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
        raise PsseImportError('No installs of Python found... wait how are you'
                ' running this...')

    _winreg.CloseKey(ver_key)
    _winreg.CloseKey(python_key)
    return python_paths

pyc_magic_nums = {20121: '1.5', 20121: '1.5.1', 20121: '1.5.2', 50428: '1.6',
                  50823: '2.0', 50823: '2.0.1', 60202: '2.1', 60202: '2.1.1',
                  60202: '2.1.2', 60717: '2.2', 62011: '2.3a0', 62021: '2.3a0',
                  62011: '2.3a0', 62041: '2.4a0', 62051: '2.4a3',
                  62061: '2.4b1', 62071: '2.5a0', 62081: '2.5a0',
                  62091: '2.5a0', 62092: '2.5a0', 62101: '2.5b3',
                  62111: '2.5b3', 62121: '2.5c1', 62131: '2.5c2',
                  62151: '2.6a0', 62161: '2.6a1', 62171: '2.7a0',
            }


# scrape pssbin paths from registry
pssbin_paths = _get_psse_locations_dict()
python_paths = _get_python_locations_dict()
psse_version = None
req_python_exec = None
ignore_python_mismatch = False
initialized = False
if check_psspy_already_in_path():
    initialized = True

    # need to find the required python for this version
    for folder in sys.path:
        if 'PSSBIN' in folder:
            # have a guess at the folder we want.
            probable_folder = folder
            break

    if not probable_folder:
        # search the entire path for psspy.pyc
        probable_folder = ''

    req_python = get_required_python_ver(probable_folder)

    if req_python != sys.winver:
        print ("WARNING: you have started a Python %s session when the\n"
                "version required by the PSSE available in your path is\n"
                "Python %s.\n"
                "Either use the required version of Python or,\n"
                "if you have another version of PSSE installed, change your\n"
                "PATH settings to point at the other install.\n\n"
                "Run '%s -m pssepath' for more info about the versions\n"
                "installed on your system.\n\n"% (sys.winver, req_python,
                    sys.executable))

        try:
            req_python_exec = os.path.join(python_paths[req_python],
                    'python.exe')
        except KeyError:
            # Very unlikely
            # Don't have the required version of python to run this version of
            # psse.  Something is not right...
            print ("Required version of python (%s) not located in registry.\n"
                    % (req_python,))
    else:
        req_python_exec = sys.executable


if __name__ == "__main__":
    # print the available psse installs.
    print 'Found the following PSSE versions installed:\n'
    print_psse_selection()
    raw_input("Press Enter to continue...")
