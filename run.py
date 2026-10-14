#!/usr/bin/env python
"""Main entry point to launch DriveTestApp."""
import sys
import os

# Ensure the project root directory is on the python search path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from src.ui.app import DriveTestAppUI

def main():
    """Runs the DriveTestApp graphical interface."""
    app = DriveTestAppUI()
    app.mainloop()

if __name__ == "__main__":
    main()
