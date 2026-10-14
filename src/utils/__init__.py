"""Utilities package for DriveTestApp."""
from .logger import setup_logger, AppLogger
from .shortcut import create_desktop_shortcut

__all__ = ["setup_logger", "AppLogger", "create_desktop_shortcut"]
