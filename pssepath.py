import os
import sys
import win32api

PSSPYC_FILENAME = 'psspy.pyc'

psse32_files = ['pyutils.dll','psspyc.pyd', 'PTIUtils.dll']

def usual_psse_paths():
    """Return a List of usual PTI install locations which exist.

    Returns all permutations of:
        - every drive on the system;
        - both usual "Program Files" dirs used;
        - with 'PTI' appended on the end;
        - only returns paths that exist.

    These are the most common install locations. Specific versions will be
    installed under this dir.
    """
    # The path is $DRIVE:\Program Files <x86>\PTI\PSSE*
    drives = win32api.GetLogicalDriveStrings()
    drives = [drivestr for drivestr in drives.split('\x00') if drivestr]

    COMMON_PROG_FILES = ('Program Files', 'Program Files (x86)')

    paths = [os.path.join(drive, folder, 'PTI') for drive in drives
                                             for folder in COMMON_PROG_FILES]

    paths = filter(os.path.exists, paths)
    return paths


def is_directory_pssbin(files):
    """Check whether the files passed in are indicitive with a PSSBIN dir.

    At the moment, it only checks whether the 'psspy.pyc' is in this folder.
    This function could easily be extended to be more rigorous by checking the
    existance of more files.
    """
    if PSSPYC_FILENAME in files:
        return True
    else:
        return False


def is_working_install(path=None):
    """Check 'psspy' can be imported once 'path' has been add to system.
    
    If 'path' is not given, then it will try it without adding anything to the
    system paths.
    """

    # TODO: This should? fail if the incorrect version of Python is used. need
    # to check that it does.

    # This is just a more robust (and more time consuming) version of
    # 'is_directory_pssbin()'. I'

    if path:
        add_psse_path(path)
    try:
        import psspy
        # Call a function of the API to make sure it isn't just a folder with a
        # file named psspy.py or psspy.pyc
        version = psspy.psseversion()

    except (ImportError, AttributeError):
        # ImportError for when psspy isn't there,
        # AttributeError for when psspy.psseversion() isn't there.
        return False
    else:
        # import worked
        return True
    finally:
        if path:
            rem_psse_path(path)


def walk_for_pssbin(path_top, depth = None):
    """Return a list of all possible PSSBIN dirs under 'path_top'.

    'depth' specifies how many folders deep the search should progress.

    Thus, if you have PSSE32 and PSSE33 installed in the same PTI dir, it will
    find both of these folders.
    """

    for root, dirs, files in os.walk(path_top):
        if is_directory_pssbin(files):
            yield root
        # We only want to go so deep in the search
        if depth and root.count(os.sep) >= depth:
            # dirs[:] is the list of directories which os.walk will descend into
            del dirs[:]

def check_psspy_already_in_path():
    """Return boolean if 'import psspy' works when this function is called.
    """
    try:
        import psspy
    except ImportError:
        return False
    return True

def get_available_psspy_location():
    """Returns a string with the path of the currently accessible PSSE install.
    """

    # Unfortunately, it is not as simple as getting the file name from
    # psspy.__file__ due to it being a pyc file, reported location and the
    # actual location may be different.
    if check_psspy_already_in_path():
        import psspy
        return os.path.normpath(os.path.dirname(psspy.__file__))
        # for directory in sys.path:
        #     if os.path.exists(os.path.join(directory,'psspy.pyc')):
        #         return directory

def add_psse_path(psse_path):
    """Add psse_path to 'sys.path' and 'os.environ['PATH'].

    This affects the os and sys modules, thus these side-effects are global.

    This is all side-effects which is not the prettiest.
    """
    sys.path.append(psse_path)
    os.environ['PATH'] = os.environ['PATH'] + ';' +  psse_path

def rem_psse_path(psse_path):
    """Remove psse_path from 'sys.path' and 'os.environ['PATH']."""

    if psse_path in sys.path:
        sys.path.remove(psse_path)
    if psse_path in os.environ['PATH']:
        sys_paths = os.environ['PATH'].split(';')
        sys_paths.remove(psse_path)
        os.environ.update({'PATH': ';'.join(sys_paths)})


def get_psse_version(path):
    """Return a version tuple: (name,major,minor,modlvl,date,stat)"""

    add_psse_path(path)
    import psspy
    rem_psse_path(path)
    version = psspy.psseversion()
    return version

def select_psse_install(installs):
    """Return selected index from the printed selection menu of installs.

    'installs' is a list of __valid__ installs. Installs should be verified by
        this stage.
    """

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
