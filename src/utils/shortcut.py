"""Windows Desktop shortcut creator utility for DriveTestApp."""
import os
import sys
from pathlib import Path

try:
    import win32com.client
except ImportError:
    win32com = None

def get_desktop_path():
    """Finds the user's active desktop directory, including OneDrive redirected paths."""
    # Check OneDrive Desktop first if present
    onedrive = os.environ.get("OneDrive") or os.environ.get("OneDriveCommercial")
    if onedrive:
        for candidate in ["Bureau", "Desktop"]:
            p = Path(onedrive) / candidate
            if p.exists():
                return str(p)

    # Standard user desktop
    home_desktop = Path.home() / "Desktop"
    if home_desktop.exists():
        return str(home_desktop)

    # Fallback to home Bureau (French Windows)
    home_bureau = Path.home() / "Bureau"
    if home_bureau.exists():
        return str(home_bureau)

    return str(Path.home())

def create_desktop_shortcut(target_script=None, shortcut_name="DriveTestApp.lnk"):
    """Creates a desktop shortcut running the script invisibly via pythonw.exe."""
    if win32com is None:
        raise RuntimeError("pywin32 is required to create a Windows shortcut.")

    desktop = get_desktop_path()
    if not target_script:
        # Default to run.py in root
        project_root = Path(__file__).resolve().parent.parent.parent
        target_script = str(project_root / "run.py")

    working_dir = str(Path(target_script).parent)
    python_exe = sys.executable
    pythonw_exe = python_exe.lower().replace("python.exe", "pythonw.exe")
    launcher = pythonw_exe if os.path.exists(pythonw_exe) else python_exe

    shortcut_path = os.path.join(desktop, shortcut_name)

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = launcher
    shortcut.Arguments = f'"{target_script}"'
    shortcut.WorkingDirectory = working_dir
    shortcut.IconLocation = launcher
    shortcut.Description = "DriveTestApp - Network Automation & RNP Processing"
    shortcut.Save()

    return shortcut_path

if __name__ == "__main__":
    try:
        path = create_desktop_shortcut()
        print(f"Shortcut created successfully at: {path}")
    except Exception as e:
        print(f"Error creating shortcut: {e}")
