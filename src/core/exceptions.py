"""Core package exceptions."""

class DriveTestAppError(Exception):
    """Base exception for all DriveTestApp errors."""
    pass

class FileValidationError(DriveTestAppError):
    """Raised when an input file or directory is missing or invalid."""
    pass

class SheetNotFoundError(DriveTestAppError):
    """Raised when an expected Excel worksheet does not exist."""
    pass

class MacroExecutionError(DriveTestAppError):
    """Raised when an Excel COM macro fails to execute."""
    pass
