pssepath - Easy PSSE Python coding
====================================

*author*: whit. (whit.com.au)

`pssepath` simplifies the code required to setup the Python environment necessary
to use the PSSE API. Using `pssepath` all you have to do is::

```python
    import pssepath
    pssepath.add_pssepath()

    import psspy
```

Tested and works on:

- PSSE 32
- PSSE 33
- PSSE 34

Supports 32 and 64 bit windows (and provides warnings when using mismatched 64
bit python when PSSE requires 32 bit python).

Using this module makes the PSSE system files available for use while avoiding
making modifications to system paths or hardcoding the location of the PSSE
system folders. This makes PSSE easy to use.

Without `pssepath`, you have to do something like this::

```python
    import os
    import sys

    PSSE_LOCATION = r"C:\Program Files\PTI\PSSE32\PSSBIN"
    sys.path.append(PSSE_LOCATION)
    os.environ['PATH'] = os.environ['PATH'] + ';' +  PSSE_LOCATION

    import psspy
```

Furthermore, by including `pssepath` with any scripts you distribute, others will
be able to use your code without having to edit your code to suit their
varying install paths (such as different versions of PSSE).

It can also provide information about which version of Python to use with a
particular install of PSSE to avoid `ImportError: Bad magic number...`.

Installation
-------------

`pip install pssepath`

or copy the `pssepath` code directory (the dir that contains `core.py`) to your
project's root directory.


Usage
------
`pssepath` provides 3 methods for setting up the PSSE paths:


- `pssepath.add_pssepath()`

    Adds the most recent version of PSSE that works with the currently
    running version of Python.

- `pssepath.add_pssepath(<version>)`

    Adds the requested version of PSSE. Remember that specifying a version
    number may make your code less portable if you or your colleagues ever use a
    different version.  `pssepath.add_pssepath(33)`

- `pssepath.select_pssepath()`

    Produces a menu of all the PSSE and Python installs on your system,
    along with the required version of Python.

If you have set up your system to have the PSSE system files on the system path
at all times, `pssepath` will only use these files.

For information about the PSSE versions installed on your system, either:

- execute the pssepath.py file from windows; or
- run `python -m pssepath.pssepathinfo` You may have to specify the python
  install path: ie. `c:\Python25\python -m pssepath.pssepathinfo` or `py.exe
  -2.5 -m pssepath.pssepathinfo`

This will provide you with a summary similar to the following::

```
Found the following versions of PSSE installed:

    1. PSSE version 32
        Requires Python 2.5-32bit (Current running Python)
    2. PSSE version XX
        Requires Python 2.X-32bit (Installed)
    3. PSSE version XX
        Requires Python 2.X-32bit

Found the following Python installations:
  2.5-32bit (currently running):
    PythonCore: C:\Python25\
  3.7-64bit:
    PythonCore: C:\Users\dan\AppData\Local\Programs\Python\Python37\
```

The status next to the Python version indicates the installation status of the
required Python for the particular PSSE install.

- `Current running Python`

    indicates the Python version used to invoke the script
    (`c:\Python25\python.exe` if invoked as `c:\Python25\python.exe -m pssepath`) is
    appropriate for that version of PSSE.

- `Installed`

  indicates that a Python version different to the one used to invoke the
  script is required for that PSSE version, but that it is already installed
  on your system.

`<nothing>`

  The absence of a status means that a different version of Python is
  required to run that version of PSSE and it is not installed on your
  system. As Python comes bundled with PSSE, this status is unlikely.

Ensuring you use the correct version of Python for the version of PSSE you are
running will avoid seeing `ImportError: Bad magic number...` ever again.

License
--------
This program is released under the very permissive MIT license. You may freely
use it for commercial purposes, without needing to provide modified source.

Read the LICENSE file for more information.

Tips on managing multiple Python versions
-------------------------------------------
I like to use batch files that point to a specific python version.  For
example::

```shell
$ more C:\bin\python25.bat
@C:\Python25\python.exe %*
```

Where the PATH includes `c:\bin`.  Now you can run python scripts with the
command::

```shell
python25 myscript.py args
```

instead of:

```shell
c:\Python25\python.exe myscript.py args
```

Contributers
--------------
Discussion about this module was conducted at the [Python for PSSE Forum](https://psspy.org/psse-help-forum/question/3/how-do-i-import-the-psspy-module-in-a-python>) involving the following members:

- Daniel Hillier
- JervisW
- Chip Webber

Improvements or suggestions
-----------------------------
Visit the [Python for PSSE Forum](https://psspy.org/psse-help-forum/question/3/how-do-i-import-the-psspy-module-in-a-python>)

See also:

- github: https://github.com/danifus/pssepath
- contact: daniel@whit.com.au

For any other questions about Python and PSSE, feel free to raise them on the
[Python for PSSE Forum](https://psspy.org>)
