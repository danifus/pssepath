import os
import sys
import _winreg

# python_v

class PsseImportError(Exception):
    pass

def check_psspy_already_in_path():
    """Return boolean if 'import psspy' works when this function is called.
    """
    try:
        import psspy
    except ImportError:
        return False
    return True

def check_initialized(fn):
    def wrapped(*args, **kwargs):
        if check_psspy_already_in_path():
            print "PSSBIN already in path, adding PSSBIN from pssepath skipped."
        elif initialized:
            print "psspath has already been executed, continuing."
        else:
            fn(*args, **kwargs)
    return wrapped

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

@check_initialized
def add_pssepath(pref_ver=None):
    """Add the PSSBIN path to the required locations.

    Try to import the requested version of PSSE. If the requested version
    doesn't work, raise an exception. By default, import the latest version.
    """
    if pref_ver:
        if pref_ver in pssbin_paths.keys():
            selected_ver = pref_ver
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
        selected_ver = sorted(pssbin_paths.keys())[-1]

    selected_path = pssbin_paths[selected_ver]
    add_dir_to_path(selected_path)
    global initialized, psse_version
    psse_version = selected_ver
    initialized = True

@check_initialized
def select_pssepath():
    """Produce a prompt to select the version of PSSE"""

    print 'Please select from the available PSSE installs:\n'
    versions = sorted(pssbin_paths.keys())
    for i, ver in enumerate(versions):
        print '  %i. PSSE Version %d\n' %(i+1, ver)
    while True:
        try:
            user_input = int(raw_input('Enter a number from the above '
                                    'PSSE installations: '))
        except ValueError:
            continue

        if 0 < user_input <= len(pssbin_paths):
            # Less one due to zero based vs 1-based (len)
            break
    add_dir_to_path(pssbin_paths[versions[user_input - 1]])
    global initialized, psse_version
    psse_version = versions[user_input - 1]
    initialized = True


# ============== Python version detection
def read_magic_number(fname):
    pyc_file = open(fname,'rb')
    magic = pyc_file.read(2)
    pyc_file.close()
    return int(magic[::-1].encode('hex'),16)

def find_file_on_path(fname):
    """Return the first file on the path which matches fname."""
    for path_dir in sys.path:
        potential_file = os.path.join(path_dir, fname)
        if os.path.isfile(potential_file):
            return potential_file

def get_required_python_ver(psse_version):
    probable_pyc = os.path.join(pssbin_paths[psse_version],'psspy.pyc')
    if not os.path.isfile(probable_pyc):
        # not in the suspected dir, perhaps abnormal install.
        probable_pyc = find_file_on_path('psspy.pyc')

    magic = read_magic_number(probable_pyc)
    req_py_ver = pyc_magic_nums[magic]

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
psse_version = None
initialized = False

if __name__ == "__main__":
    # print the available psse installs.
    print 'Found the following PSSE versions installed:\n'
    versions = sorted(pssbin_paths.keys())
    for i, ver in enumerate(versions):
        print '  %i. PSSE Version %d\n' %(i+1, ver)
