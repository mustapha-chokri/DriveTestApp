"""Logging utilities for DriveTestApp."""
import os
import logging
from datetime import datetime

DEFAULT_LOG_FILE = "app.log"

def setup_logger(log_file=DEFAULT_LOG_FILE, level=logging.INFO):
    """Sets up standard application logging."""
    logger = logging.getLogger("DriveTestApp")
    logger.setLevel(level)

    # Avoid duplicate handlers on re-initialization
    if not logger.handlers:
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger

class AppLogger:
    """Wrapper that logs to Python logger and optionally streams to a UI callback."""

    def __init__(self, ui_callback=None, log_file=DEFAULT_LOG_FILE):
        self.logger = setup_logger(log_file)
        self.ui_callback = ui_callback

    def set_callback(self, ui_callback):
        self.ui_callback = ui_callback

    def log(self, message, level="INFO"):
        time_str = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{time_str}] {message}"
        
        if self.ui_callback:
            try:
                self.ui_callback(formatted)
            except Exception:
                pass

        if level.upper() == "ERROR":
            self.logger.error(message)
        elif level.upper() == "WARNING":
            self.logger.warning(message)
        else:
            self.logger.info(message)
