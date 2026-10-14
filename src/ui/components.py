"""Reusable UI components for DriveTestApp."""
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText

class HeaderWidget(ttk.Frame):
    """Modern header banner with title and subtitle."""

    def __init__(self, parent, title="DriveTestApp", subtitle="Network Automation Hub"):
        super().__init__(parent)
        self.pack(fill=X, pady=(0, 20))

        title_label = ttk.Label(
            self,
            text=title,
            font=("Segoe UI", 24, "bold"),
            foreground="#00e5ff"
        )
        title_label.pack(side=LEFT)

        subtitle_label = ttk.Label(
            self,
            text=subtitle,
            font=("Segoe UI", 10),
            foreground="#888888"
        )
        subtitle_label.pack(side=LEFT, padx=12, pady=(10, 0))

class FilePickerField:
    """A standard file/folder picker row with label, entry, and browse button."""

    def __init__(self, parent, label_text, string_var, browse_command, row, width=68):
        self.label = ttk.Label(parent, text=label_text)
        self.label.grid(row=row, column=0, sticky=W, pady=6, padx=(0, 10))

        self.entry = ttk.Entry(parent, textvariable=string_var, width=width)
        self.entry.grid(row=row, column=1, sticky=EW, pady=6)

        self.btn = ttk.Button(parent, text="...", width=5, command=browse_command)
        self.btn.grid(row=row, column=2, padx=(8, 0), pady=6)

class LogViewerWidget(ttk.Frame):
    """Scrolled log viewer widget with autoscroll support."""

    def __init__(self, parent, title="Activity Logs", height=12):
        super().__init__(parent)
        self.pack(fill=BOTH, expand=True)

        header = ttk.Label(self, text=title, font=("Segoe UI", 10, "bold"))
        header.pack(anchor=W, pady=(10, 5))

        self.text_widget = ScrolledText(
            self,
            height=height,
            autohide=True,
            font=("Consolas", 9)
        )
        self.text_widget.pack(fill=BOTH, expand=True)

    def write_line(self, message):
        """Appends a new line to the log and auto-scrolls to the end."""
        self.text_widget.insert(END, f"{message}\n")
        self.text_widget.see(END)

    def clear(self):
        """Clears all text from the log viewer."""
        self.text_widget.delete("1.0", END)
