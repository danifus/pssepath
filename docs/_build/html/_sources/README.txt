====================================
pssepath - Easy PSSE Python coding
====================================

:author: whit. (whit.com.au)

pssepath simplifies the code required to setup the Python environment necessary
to use the PSSE API. The usage of this package (after installation) would be as
follows::

    import pssepath
    pssepath.add_pssepath()
    import psspy

Notice that using this module makes the PSSE system files available for use
while avoiding making modifications to system paths or hardcoding the location
of the PSSE system folders. This makes PSSE easy to use.

Furthermore, by including pssepath with any scripts you distribute, others will
be able to use your code without having to edit your code to suit their
varying install paths (such as different versions of PSSE).

It can also provide information about which version of Python to use with a
particular install of PSSE to avoid "ImportError: Bad magic number...".

Installation
-------------
To install from source::

    python setup.py install

Usage
------
pssepath provides 3 methods for setting up the PSSE paths:

.. method::  pssepath.add_pssepath()

      Adds the most recent version of PSSE that works with the currently
      running version of Python.

.. method:: pssepath.add_pssepath(<version>)

    Adds the requested version of PSSE. Remember that specifying a version
    number may make your code less portable if you or your colleagues ever use a
    different version.  pssepath.add_pssepath(33)

.. method:: pssepath.select_pssepath()

    Produces a menu of all the PSSE installs on your system, along with
    the required version of Python.

If you have set up your system to have the PSSE system files on the system path
at all times, pssepath will only use these files.

For information about the PSSE versions installed on your system, either:

    - execute the pssepath.py file from windows; or
    - run ``python -m pssepath`` You may have to specify the python install
      path: ie. ``c:\Python25\python -m pssepath``

This will provide you with a summary similar to the following::

    Found the following versions of PSSE installed:

        1. PSSE version 32
            Requires Python 2.5 (Currently running)
        2. PSSE version XX
            Requires Python 2.X (Installed)
        3. PSSE version XX
            Requires Python 2.X
        etc.

The status next to the Python version indicates the installation status of the
required Python for the particular PSSE install.

Currently running
    indicates that the Python version used to invoke the script
    ("c:\Python25\python" if invoked as "c:\Python25\python -m pssepath") is
    appropriate for that version.
    
Installed
    indicates that a Python version different to the one used to invoke the
    script is required for that PSSE version, but that it is already installed
    on your system.  

<nothing>    
     The absence of a status means that a different version of Python is
     required to run that version of PSSE and it is not installed on your
     system. As Python comes bundled with PSSE, this status is unlikely.

Ensuring you use the correct version of Python for the version of PSSE you are
running will avoid seeing ``ImportError: Bad magic number...`` ever again.

License
--------
This program is released under the very permissive MIT license. You may freely
use it for commercial purposes, without needing to provide modified source.

Read the LICENSE file for more information.

Contributers
--------------
Discussion about this module was conducted at the `whit psse forum <http://forum.whit.com.au/psse-help-forum/question/3/how-do-i-import-the-psspy-module-in-a-python>`_ involving the following members:

      - Chip Webber 
      - JervisW
      - Daniel Hillier

Improvements or suggestions
-----------------------------
Visit the `whit forum <http://forum.whit.com.au/psse-help-forum/question/3/how-do-i-import-the-psspy-module-in-a-python>`_

See also:

    - github: https://github.com/danaran/pssepath
    - contact: daniel .at. whit.com.au
