#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pytk - Python ToolKit: A thin Pythonic layer on top of TkInter GUI.
"""

# Copyright (c) Riccardo Metere <rick@metere.it>

# ======================================================================
# :: Future Imports
from __future__ import (
    division, absolute_import, print_function, unicode_literals, )

# ======================================================================
# :: Python Standard Library Imports
try:
    import tkinter as tk
    import tkinter.ttk as ttk
    import tkinter.messagebox as messagebox
    import tkinter.filedialog as filedialog
    import tkinter.simpledialog as simpledialog
except ImportError:
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
    import tkFileDialog as filedialog
    import tkSimpleDialog as simpledialog

Window = tk.Toplevel
Style = ttk.Style

# ======================================================================
# :: External Imports
# import flyingcircus as fc  # Everything you always wanted to have in Python*
from flyingcircus import msg, dbg, fmt, fmtm, elapsed, report, pkg_paths
from flyingcircus import VERB_LVL, VERB_LVL_NAMES, D_VERB_LVL

# ======================================================================
# :: Version
from ._version import __version__

# ======================================================================
# :: Project Details
INFO = {
    'name': 'PyTK',
    'author': 'PyTK developers',
    'contrib': (
        'Riccardo Metere <rick@metere.it>',
    ),
    'copyright': 'Copyright (C) 2017-2018',
    'license': 'GNU General Public License version 3 or later (GPLv3+)',
    'notice':
        """
This program is free software and it comes with ABSOLUTELY NO WARRANTY.
It is covered by the GNU General Public License version 3 (GPLv3).
You are welcome to redistribute it under its terms and conditions.
        """,
    'version': __version__
}

# ======================================================================
# :: quick and dirty timing facility
_EVENTS = []

# ======================================================================
# Greetings
MY_GREETINGS = r"""
 ____       _____ _  __
|  _ \ _   |_   _| |/ /
| |_) | | | || | | ' /
|  __/| |_| || | | . \
|_|    \__, ||_| |_|\_\
       |___/
"""
# generated with: pyfiglet 'PyTK' -f standard

# :: Causes the greetings to be printed any time the library is loaded.
# print(MY_GREETINGS)

# ======================================================================
if __name__ == '__main__':
    import doctest  # Test interactive Python examples

    msg(__doc__.strip())
    doctest.testmod()
    msg(report())
