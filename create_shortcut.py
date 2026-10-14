#!/usr/bin/env python
"""Creates a Windows desktop shortcut for DriveTestApp."""
import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from src.utils.shortcut import create_desktop_shortcut

if __name__ == "__main__":
    try:
        path = create_desktop_shortcut()
        print(f"Shortcut created successfully at: {path}")
    except Exception as e:
        print(f"Error creating shortcut: {e}")
