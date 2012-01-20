import os
import sys
import win32api

PSSPYC_FILENAME = 'psspy.pyc'

def usual_psse_paths():
    # The path is $DRIVE:\Program Files <x86>\PTI\PSSE*
    drives = win32api.GetLogicalDriveStrings()
    drives = [drivestr for drivestr in drives.split('\x00') if drivestr]

    COMMON_PROG_FILES = ('Program Files', 'Program Files (x86)')

    paths = [os.path.join(drive, folder, 'PTI') for drive in drives
                                             for folder in COMMON_PROG_FILES]

    paths = filter(os.path.exists, paths)
    return paths

def is_directory_pssbin(files):
    if PSSPYC_FILENAME in files:
        return True
    else:
        return False

def add_psse_path(psse_path):
    """Add psse_path to 'sys.path' and 'os.environ['PATH'].

    Checks whether there is already a PSSE environment in the path. The path
    that we are adding will probably be later in the path search and thus
    overlooked, so we abort.

    This affects the os and sys modules, thus these side-effects are global.

    This is all side-effects which is not the prettiest.
    """
    try:
        # if this works, then there is already a path that has a psspy module
        # in it and we can't be certain that our setup is doing its job.
        import psspy
    except ImportError:
        # No PSSE environ found, so lets add one!
        sys.path.append(psse_path)
        os.environ['PATH'] = os.environ['PATH'] + ';' +  psse_path
    else:
        path_to_install = os.path.dirname(psspy.__file__)
        # We found a PSSE environ already active. Let's get out of here!
        raise Exception('pssepath found a PSSE environ already loaded: '
                'Aborting\n'
                'Path to install: %s:' % (path_to_install,))

def rem_psse_path(psse_path):
    """Remove psse_path from 'sys.path' and 'os.environ['PATH']."""

    if psse_path in sys.path:
        sys.path.remove(psse_path)
    if psse_path in os.environ['PATH']:
        sys_paths = os.environ['PATH'].split(';')
        sys_paths.remove(psse_path)
        os.environ.update({'PATH': ';'.join(sys_paths)})

def is_working_install(path):
    add_psse_path(path)
    try:
        import psspy
    except ImportError:
        return False
    else:
        # import worked
        return True
    finally:
        rem_psse_path(path)

def get_psse_version(path):
    add_psse_path(path)
    import psspy
    rem_psse_path(path)
    # psspy.psseversion() returns: name,major,minor,modlvl,date,stat
    version = psspy.psseversion()
    return version


def walk_for_pssbin(path_top, depth = None):
    for root, dirs, files in os.walk(path_top):
        if is_directory_pssbin(files):
            yield root
        # We only want to go so deep in the search
        if depth and root.count(os.sep) >= depth:
            # dirs[:] is the list of directories which os.walk will descend into
            del dirs[:]

def select_psse_install(installs):
    print 'Please select from the available PSSE installs:\n'
    for i, path in enumerate(installs):
        name,major,minor,modlvl,date,stat = get_psse_version(path)
        print '  %i. Version %s.%s - %s\n' %(i+1, major, minor, path)
    while True:
        try:
            user_input = int(raw_input('Enter a number from the above '
                                    'PSSE installations: '))
        except ValueError:
            continue

        if 1 <= user_input <= len(installs):
            # Less one due to zero based vs 1-based (len)
            return user_input - 1

def setup_psspy_env():
    # We should look in any configuration files first and if they don't exist,
    # then move onto automatically determining the install location.

    # Reason for configuration file:
    #   - the non-standard location search takes a while to complete.
    #   - if some sort of prompt is needed to determine which is the correct
    #     install to use, this only happens once.

    # Possible locations for config files:
    #   - python site-library or python directory
    #     - allows for install specific options. A 2.5 and 2.7 install
    #       wouldn't conflict.
    #   - user's home
    #     - Should be writable, regardless of corporate setups. May be moved
    #       between computers, which may not work if they have different install
    #       directories. A function to update this entry would be trivial to run
    #       again, but non-savy users may not find it straight away.
    #     - It could be set up to make a unique config name (computer name with
    #       python version) for different computers it the config file is run on

    # It would be preferable to only have one location for it so there is only
    # one file to look at if things go wrong (instead of having: "First look
    # here then if the possible entries in there don't match, look over
    # there.")

    # Look to automatically find the path.
    # Look in the usual suspects
    pssbin_paths = []
    paths = usual_psse_paths()
    for p in paths:
        pssbin_paths.extend(list(walk_for_pssbin(p)))

    for p in pssbin_paths[:]:
        if not is_working_install(p):
            pssbin_paths.remove(p)

    if len(pssbin_paths) == 1:
        # Only one way about it.
        psse_location = pssbin_paths[0]
    elif len(pssbin_paths) > 1:
        # Not tested yet so will
        install_index = select_psse_install(pssbin_paths)
        psse_location = pssbin_paths[install_index]
    else:
        raise ImportError('No PSSE installs found in the usual locations.\n')

    name,major,minor,modlvl,date,stat = get_psse_version(psse_location)
    print ('[pssepath] Using PSSE Version %s.%s\n'
           '[pssepath]   Install Path: %s\n' %(major, minor, psse_location))

    # Time to add the path and get to work.
    add_psse_path(psse_location)

# ========== Python init
# get the ball rolling on import.
setup_psspy_env()

if __name__ == "__main__":
    # do the same thing as just importing. Probably can do better than that.
    # Maybe running the module should produce some diagnostics about what
    # installs are present etc.
    setup_psspy_env()
