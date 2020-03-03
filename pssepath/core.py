from __future__ import with_statement

import logging
import os
import sys
from functools import wraps
from textwrap import dedent

try:
    # Py2
    import _winreg as winreg
except ImportError:
    # Py3
    import winreg

from .compat import compat_input, simple_print, open_hkey_ctxmg
from . import helpers


logger = logging.getLogger(__name__)


PSSE_VERSION = None
INITIALIZED = False


class PsseImportError(Exception):
    pass


def check_psspy_already_in_path():
    """Return True if psspy.pyc in the sys.path and os.environ['PATH'] dirs.

    Otherwise, print a warning message and return False so the paths get
    reconfigured.
    """
    syspath = find_file_on_path("psspy.pyc", sys.path)

    if syspath:
        # file in one of the files on the sys.path (python's path) list.
        envpaths = os.environ["PATH"].split(";")
        envpath = find_file_on_path("psspy.pyc", envpaths)
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
        if INITIALIZED:
            logger.info("psspath has already added PSSBIN to the system, continuing.")
        elif check_psspy_already_in_path():
            check_already_present_psse()
            logger.info("PSSBIN already in path, adding PSSBIN from pssepath skipped.")
            set_status(initialized=True)
        else:
            fn(*args, **kwargs)

    return wrapped


@helpers.run_once
def log_path_noenviron_warning():
    logger.warning(
        dedent(
            """\
       pssepath: Warning - PSSBIN found on sys.path, but not os.environ['PATH'].
                           Running pssepath.add_pssepath() will reconfigure.

                 Running pssepath.add_pssepath() will attempt to reconfigure
                 your paths for you.  If you wish to find the root cause of
                 this message, check your Python scripts to see if they set up
                 sys.path or os.environ['PATH'] and remove that code.  If the
                 scripts do not attempt to configure these variables, you may
                 need to check your Windows PATH variables from windows, as they
                 may have been configured there.
                 """
        )
    )


@helpers.run_once
def log_pathmismatch_warning(syspath, envpath):
    logger.warning(
        dedent(
            """\
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
                 """
        )
        % (syspath, envpath)
    )


def add_dir_to_path(psse_ver, psse_path):
    """Add psse_path to 'sys.path' and 'os.environ['PATH'].

    This affects the os and sys modules, thus these side-effects are global.
    Adds them to the start of the path variables so that they are always used
    in preference.

    This is all side-effects which is not the prettiest.
    """
    sys.path.insert(0, psse_path)
    os.environ["PATH"] = psse_path + ";" + os.environ["PATH"]

    if psse_ver >= 34:
        # Also add the PSSBIN dir
        pssebin_dir = os.path.join(os.path.dirname(psse_path), "PSSBIN")
        sys.path.insert(0, pssebin_dir)
        os.environ["PATH"] = pssebin_dir + ";" + os.environ["PATH"]


def search_pssbin_reg_key(pti_key):
    pssbin_paths = {}
    for sub_key in helpers.enum_reg_keys(pti_key):
        try:
            with open_hkey_ctxmg(pti_key, sub_key + "\\Product Paths") as ver_key:
                # Version num is the last 2 digits of the subkey
                version_num = int(sub_key[-2:])
                pssbin_paths[version_num] = helpers.get_reg_value(
                    ver_key, "PsseExePath"
                )
        except WindowsError:
            pass
    return pssbin_paths


@helpers.memoize
def get_pssbin_paths_dict():
    pssbin_paths = {}
    if helpers.is_win64():
        # Check 32bit install registry
        try:
            with open_hkey_ctxmg(
                winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Wow6432Node\\PTI"
            ) as pti_key:
                pssbin_paths.update(search_pssbin_reg_key(pti_key))
        except WindowsError:
            pass
        # Check 64bit install registry
        try:
            with open_hkey_ctxmg(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\PTI") as pti_key:
                pssbin_paths.update(search_pssbin_reg_key(pti_key))
        except WindowsError:
            pass
    else:
        # Only 32bit install registry
        try:
            with open_hkey_ctxmg(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\PTI") as pti_key:
                pssbin_paths.update(search_pssbin_reg_key(pti_key))
        except WindowsError:
            pass

    if not len(pssbin_paths):
        raise PsseImportError("No installs of PSSE found.")

    return pssbin_paths


@helpers.memoize
def get_psse_locations_dict():
    """Return a dict of {(psse_ver, pyver): psspy_path}"""
    psspy_dirs = {}
    pssbin_paths = get_pssbin_paths_dict()
    for psse_ver, pssbin in pssbin_paths.items():
        pyvers_and_psspy_paths = get_required_python_ver_and_paths(psse_ver, pssbin)
        for pyver, psspy_path in pyvers_and_psspy_paths:
            psspy_dirs[(psse_ver, pyver)] = psspy_path
    return psspy_dirs


def get_psse_arch(psse_version):
    if psse_version < 35:
        return "32bit"
    return "64bit"


def check_to_raise_compat_python_error(psse_and_py_versions):
    psspy_paths = get_psse_locations_dict()
    possible_pyvers = []
    req_psse_ver, req_py_ver = psse_and_py_versions
    for psse_ver, pyver in psspy_paths.keys():
        if psse_ver == req_psse_ver:
            possible_pyvers.append(pyver)

    running_py_ver = helpers.get_python_ver()
    if running_py_ver not in possible_pyvers:
        pyver_text = " or ".join(["-".join(py_ver) for py_ver in possible_pyvers])

        psse_arch = get_psse_arch(req_psse_ver)

        raise PsseImportError(
            "Current Python and PSSE version "
            "incompatible.\n\n"
            "PSSE %s (%s) requires Python %s to run.\n"
            "Currently running Python%s.\n"
            % (req_psse_ver, psse_arch, pyver_text, "-".join(running_py_ver))
        )


@check_initialized
def add_pssepath(pref_psse_ver=None):
    """Add the PSSBIN path to the required locations.

    Try to import the requested version of PSSE. If the requested version
    doesn't work, raise an exception. By default, import the latest version.
    """
    psspy_paths = get_psse_locations_dict()
    current_pyver = helpers.get_python_ver()

    if pref_psse_ver:
        available_psse_versions_set = set()
        for psse_ver, pyver in psspy_paths.keys():
            available_psse_versions_set.add(psse_ver)
        available_psse_versions = sorted(available_psse_versions_set)

        if pref_psse_ver in available_psse_versions:
            selected_psse_ver = pref_psse_ver
            check_to_raise_compat_python_error((selected_psse_ver, current_pyver))
        else:
            if len(available_psse_versions) == 1:
                ver_string = "the installed version: %s" % (available_psse_versions[0],)
            else:
                psses = ", ".join([str(x) for x in available_psse_versions])
                ver_string = "an installed version: %s" % psses

            psse_arch = get_psse_arch(pref_psse_ver)
            raise PsseImportError(
                "Attempted to initialize PSSE version %s (%s) but it was not present.\n"
                "Let pssepath select the latest version by not specifying a "
                "version when\n"
                'calling "pssepath.add_pssepath()", or select %s'
                % (pref_psse_ver, psse_arch, ver_string)
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
                check_to_raise_compat_python_error((psse_ver, pyver))
            except PsseImportError:
                pass
            else:
                selected_psse_ver = psse_ver
                break
        if not selected_psse_ver:
            raise PsseImportError(
                "No installed PSSE versions (%s) are compatible "
                "with the running version of Python (%s)"
                % (
                    ", ".join(
                        ["v%s" % psse_ver for psse_ver, py_ver in sorted(psspy_paths.keys())]
                    ),
                    "-".join(current_pyver),
                )
            )

    selected_path = psspy_paths[(selected_psse_ver, current_pyver)]
    add_dir_to_path(selected_psse_ver, selected_path)
    set_status(psse_version=selected_psse_ver, initialized=True)


@check_initialized
def select_pssepath():
    """Produce a prompt to select the version of PSSE"""

    simple_print("Please select from the available PSSE installs:\n")
    options = print_psse_selection()
    psspy_paths = get_psse_locations_dict()
    while True:
        try:
            user_input = int(
                compat_input("Enter a number from the above PSSE installations: ")
            )
        except ValueError:
            continue

        if 0 < user_input <= len(options):
            # Less one due to zero based vs 1-based (len)
            break

    selected_key = options[user_input]
    psse_ver, pyver = selected_key
    selected_path = psspy_paths[selected_key]
    check_to_raise_compat_python_error(selected_key)
    add_dir_to_path(psse_ver, selected_path)
    set_status(psse_version=selected_key, initialized=True)


def print_psse_selection():

    psspy_paths = get_psse_locations_dict()
    versions = sorted(psspy_paths.keys())
    running_py_ver = helpers.get_python_ver()
    options = dict([(i + 1, version) for i, version in enumerate(versions)])
    installed_py_vers = get_installed_py_vers()
    for i, (psse_ver, pyver) in options.items():
        python_str = "Requires Python%s" % ("-".join(pyver),)
        if pyver == running_py_ver:
            python_str += " (Current running Python)"
        elif running_py_ver in installed_py_vers:
            python_str += " (Installed, not current running version.)"
        simple_print("  %i. PSSE Version %d\n      %s" % (i, psse_ver, python_str))
    return options


def print_python_selection():
    python_by_paths = get_pythons_by_location()
    python_paths = {}
    for path, vals in python_by_paths.items():
        py_ver = (vals[0], vals[2])
        msg = "    %s: %s" % (vals[1], path)
        try:
            python_paths[py_ver].append(msg)
        except KeyError:
            python_paths[py_ver] = [msg]

    parts = []
    running_py_ver = helpers.get_python_ver()
    for version in sorted(python_paths.keys()):
        if running_py_ver == version:
            py_msg = "  " + "-".join(version) + " (currently running):"
        else:
            py_msg = "  " + "-".join(version) + ":"
        parts.append(py_msg)
        for msg in python_paths[version]:
            parts.append(msg)
    simple_print("\n".join(parts))


# ============== Python version detection
def get_pythons_from_reg(python_key, fallback_nbits):
    """
    Returns [(path_to_python, version, company, nbits, fallback_nbits)].
    """
    pythons_by_location = []
    for company in helpers.enum_reg_keys(python_key):
        with open_hkey_ctxmg(python_key, company) as company_key:
            for version_tag in helpers.enum_reg_keys(company_key):

                with open_hkey_ctxmg(company_key, version_tag) as ver_key:
                    arch = helpers.get_reg_value(ver_key, "SysArchitecture")

                with open_hkey_ctxmg(company_key, version_tag) as ver_key:
                    # 2.7 or 3.7 etc. (maybe 3.7.1 for verions that aren't
                    # PythonCore).
                    sys_version = helpers.get_reg_value(ver_key, "SysVersion")

                if sys_version is None:
                    sys_version = version_tag

                # only use 3.7 from a 3.7.1 version.
                sys_version = sys_version[:3]

                installpath = "\\".join([version_tag, "InstallPath"])
                with open_hkey_ctxmg(company_key, installpath) as install_key:
                    path = winreg.QueryValue(install_key, None)

                pythons_by_location.append(
                    (path, sys_version, company, arch, fallback_nbits)
                )
    return pythons_by_location


def get_pythons_by_location():
    """Returns a dictionary of {python install path: (version, company, arch)}

    Searches the appropriate registry keys if running on windows 32bit or 64bit.
    """

    def consolodate(python_infos, python_dict):
        """Update the list of python installs, deduped by install path.

        Gives preference to entries which have a non-default 'arch' value.
        """
        for path, version, company, nbits, fallback_nbits in python_infos:
            if nbits is not None:
                arch = nbits
            else:
                arch = fallback_nbits
            if path in python_dict:
                if python_dict[path][2] != unknown_bits:
                    continue

            python_dict[path] = (version, company, arch)
        return python_dict

    pythons_by_location = {}
    unknown_bits = "?bits"
    try:
        with open_hkey_ctxmg(
            winreg.HKEY_CURRENT_USER, "SOFTWARE\\Python"
        ) as python_key:
            python_infos = get_pythons_from_reg(python_key, unknown_bits)
        pythons_by_location = consolodate(python_infos, pythons_by_location)
    except WindowsError:
        pass

    if helpers.is_win64():
        try:
            with open_hkey_ctxmg(
                winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Wow6432Node\\Python"
            ) as python_key:
                python_infos = get_pythons_from_reg(python_key, "32bit")
            pythons_by_location = consolodate(python_infos, pythons_by_location)
        except WindowsError:
            pass

        try:
            with open_hkey_ctxmg(
                winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Python"
            ) as python_key:
                python_infos = get_pythons_from_reg(python_key, "64bit")
            pythons_by_location = consolodate(python_infos, pythons_by_location)
        except WindowsError:
            pass
    else:
        try:
            with open_hkey_ctxmg(
                winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Python"
            ) as python_key:
                python_infos = get_pythons_from_reg(python_key, "32bit")
            pythons_by_location = consolodate(python_infos, pythons_by_location)
        except WindowsError:
            pass

    return pythons_by_location


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


def get_required_python_for_psspy_in(psspy_dir, psse_ver=None):
    """
    Returns python version of psspy.pyc and psse archicture or None if no psspy.pyc file in this dir.

    returns:
        eg. ("2.7", "32bit")
    """
    probable_pyc = os.path.join(psspy_dir, "psspy.pyc")
    if not os.path.isfile(probable_pyc):
        return None

    py_ver = helpers.read_magic_number(probable_pyc)
    # only the first 3 digits are important (2.x etc)
    py_ver = py_ver[:3]
    if psse_ver is None:
        return py_ver, None
    else:
        return py_ver, get_psse_arch(psse_ver)


def get_required_python_ver_psse_33_and_older(pssbin, psse_ver):
    """
    Returns a list of possible python versions and the path which has psspy.pyc
    on it (of which 33 and lower only supports one version).
    """
    pyver = get_required_python_for_psspy_in(pssbin, psse_ver)
    if pyver is not None:
        return [(pyver, pssbin)]
    return []


def find_psse_pydirs(psse_base_dir, psse_ver):
    pyvers_and_paths = []
    for fpath in os.listdir(psse_base_dir):
        psspy_dir = os.path.join(psse_base_dir, fpath)
        if os.path.isdir(psspy_dir) and fpath.upper().startswith("PSSPY"):
            pyver = get_required_python_for_psspy_in(psspy_dir, psse_ver)
            if pyver is not None:
                pyvers_and_paths.append((pyver, psspy_dir))
    return pyvers_and_paths


def get_required_python_ver_psse_34_and_newer(pssbin, psse_ver):
    """
    Returns a list of possible python versions. PSSE34 and above support
    multiple python verions.
    """
    # Get the directory above the PSSBIN dir.
    psse_base_dir = os.path.dirname(pssbin)
    pyvers_and_paths = find_psse_pydirs(psse_base_dir, psse_ver)
    return pyvers_and_paths


def get_required_python_ver_and_paths(psse_ver, pssbin):
    """
    Return a list of [(pyver, psspy_dir), ...] or [] if no path.
    """
    if psse_ver < 34:
        pyvers_and_paths = get_required_python_ver_psse_33_and_older(pssbin, psse_ver)
    elif psse_ver >= 34:
        pyvers_and_paths = get_required_python_ver_psse_34_and_newer(pssbin, psse_ver)

    return pyvers_and_paths


@helpers.memoize
def get_installed_py_vers():
    """
    Returns a list of (py_ver, nbits) for any detected python paths.

    Will only return one entry per (py_ver, nbits) combo.
    """
    pythons_by_location = get_pythons_by_location()

    python_vers = set()
    for path, vals in pythons_by_location.items():
        python_vers.add((vals[0], vals[2]))

    python_vers = list(python_vers)

    if not len(python_vers):
        raise PsseImportError(
            "No installs of Python found... wait how are you running this..."
        )

    return python_vers


def get_psse_programfiles(psse_version):
    if psse_version < 35:
        return helpers.get_programfiles_32()
    else:
        if helpers.is_win64():
            return helpers.get_programfiles_64()
        else:
            raise Exception(
                "PSSe 35 and greater are 64bit only and you are using a 32bit windows."
            )


def set_status(**kwargs):

    global PSSE_VERSION, INITIALIZED

    if "psse_version" in kwargs:
        PSSE_VERSION = kwargs["psse_version"]

    if "initialized" in kwargs:
        INITIALIZED = kwargs["initialized"]


def check_already_present_psse():
    """
    Raises errors if the already present PSSE looks to be misconfigured.
    """
    if not INITIALIZED and check_psspy_already_in_path():

        pypaths = set()
        pyvers = []
        # need to find the required python for this version
        for possible_psspy_dir in sys.path:
            if "PSSBIN" in possible_psspy_dir or "PSSPY" in possible_psspy_dir:

                pyver = get_required_python_for_psspy_in(possible_psspy_dir)
                if pyver is not None:
                    abs_path = os.path.abspath(possible_psspy_dir)
                    if abs_path not in pypaths:
                        pyvers.append(pyver)

        if len(pyvers) > 1:
            paths = "\n".join(list(pypaths))
            raise PsseImportError(
                "WARNING: your PATH has been configured to make multiple "
                "psspy.pyc files availble. This may lead to loading the "
                "incorrect copy. Please only have one version available. "
                "\nThe multiple paths are:\n%s" % paths
            )

        req_python = pyvers[0]
        running_py_ver = helpers.get_python_ver()

        if req_python[0] != running_py_ver[0]:
            raise PsseImportError(
                "WARNING: you have started a Python %s session when the\n"
                "version required by the PSSE available in your path is\n"
                "Python %s.\n"
                "Either use the required version of Python or,\n"
                "if you have another version of PSSE installed, change your\n"
                "PATH settings to point at the other install.\n\n"
                "Run '%s -m psseutils.pssepathinfo' for more info about the versions\n"
                "installed on your system.\n\n"
                % ("-".join(running_py_ver), req_python, sys.executable)
            )

        # a psse path is already present and looks ok.
        set_status(initialized=True)
