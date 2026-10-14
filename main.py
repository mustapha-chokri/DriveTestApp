#!/usr/bin/env python
"""
DriveTestApp - Main Entry Point.
Maintained for backward-compatibility with existing shortcuts and scripts.
Delegates directly to the modularized application structure in src/
"""
import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from run import main

if __name__ == "__main__":
    main()
