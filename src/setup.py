#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

from distutils.core import setup

import py2exe


setup(name='excel-password-recovery',
      windows=[{'script': 'src/gui-main.py'}],
      options={
               'py2exe': {
                          'includes': ['Tkinter', 'tkFileDialog',
                                       'tkMessageBox', 'tools']
                          }
               },
      )
