"""Main Application Window for DriveTestApp."""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from ..config import AppConfig
from ..utils.logger import AppLogger
from ..core.excel_processor import ExcelProcessor
from .components import HeaderWidget, FilePickerField, LogViewerWidget

class DriveTestAppUI(ttk.Window):
    """Main application window using ttkbootstrap."""

    def __init__(self, themename="superhero"):
        super().__init__(themename=themename, title="DriveTestApp - Network Automation")
        self.geometry("820x680")
        self.minsize(700, 550)

        # Config & Logger
        self.config = AppConfig()
        self.logger = AppLogger()

        # Tkinter Variables
        self.rnp_path_var = tk.StringVar(value=self.config.rnp_path)
        self.adt_path_var = tk.StringVar(value=self.config.adt_path)
        self.output_dir_var = tk.StringVar(value=self.config.output_dir)
        self.city_var = tk.StringVar(value=self.config.city)

        self.last_nbf_dir = None

        self._build_ui()
        self.logger.set_callback(self.log_viewer.write_line)
        self.logger.log("DriveTestApp initialized and ready.")

    def _build_ui(self):
        container = ttk.Frame(self, padding=20)
        container.pack(fill=BOTH, expand=True)

        # Header
        HeaderWidget(container)

        # Configuration Card
        card = ttk.Labelframe(container, text=" Configuration ", padding=16)
        card.pack(fill=X, pady=(0, 15))
        card.columnconfigure(1, weight=1)

        # File Fields
        FilePickerField(card, "RNP Source File:", self.rnp_path_var, self.browse_rnp, 0)
        FilePickerField(card, "Template ADT:", self.adt_path_var, self.browse_adt, 1)
        FilePickerField(card, "Output Directory:", self.output_dir_var, self.browse_output, 2)

        # Target City Field
        city_label = ttk.Label(card, text="Target City:")
        city_label.grid(row=3, column=0, sticky=W, pady=8, padx=(0, 10))
        city_entry = ttk.Entry(card, textvariable=self.city_var, width=35)
        city_entry.grid(row=3, column=1, sticky=W, pady=8)

        # Action Buttons
        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill=X, pady=(0, 15))

        self.run_btn = ttk.Button(
            btn_frame,
            text="RUN PROCESS",
            style='primary.TButton',
            width=16,
            command=self.start_processing
        )
        self.run_btn.pack(side=LEFT, padx=(0, 8))

        self.nbf_btn = ttk.Button(
            btn_frame,
            text="OPEN NBF FOLDER",
            style='info.Outline.TButton',
            state=DISABLED,
            command=self.open_nbf_folder
        )
        self.nbf_btn.pack(side=LEFT, padx=(0, 8))

        exit_btn = ttk.Button(
            btn_frame,
            text="EXIT",
            style='danger.Outline.TButton',
            command=self.destroy
        )
        exit_btn.pack(side=RIGHT)

        # Log Section
        self.log_viewer = LogViewerWidget(container)

    def browse_rnp(self):
        filename = filedialog.askopenfilename(
            title="Select Source RNP File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if filename:
            self.rnp_path_var.set(filename)

    def browse_adt(self):
        filename = filedialog.askopenfilename(
            title="Select Template ADT File",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All Files", "*.*")]
        )
        if filename:
            self.adt_path_var.set(filename)

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Directory")
        if folder:
            self.output_dir_var.set(folder)

    def open_nbf_folder(self):
        if self.last_nbf_dir and os.path.exists(self.last_nbf_dir):
            os.startfile(self.last_nbf_dir)
        else:
            messagebox.showwarning("Notice", "NBF folder does not exist yet.")

    def start_processing(self):
        """Saves config and launches processing in a daemon background thread."""
        self.config.save(
            rnp_path=self.rnp_path_var.get(),
            adt_path=self.adt_path_var.get(),
            output_dir=self.output_dir_var.get(),
            city=self.city_var.get()
        )

        self.run_btn.config(state=DISABLED)
        self.nbf_btn.config(state=DISABLED)

        processor = ExcelProcessor(
            adt_path=self.adt_path_var.get(),
            rnp_path=self.rnp_path_var.get(),
            output_dir=self.output_dir_var.get(),
            city_name=self.city_var.get(),
            logger=self.logger
        )

        worker = threading.Thread(target=self._run_async, args=(processor,), daemon=True)
        worker.start()

    def _run_async(self, processor):
        try:
            result = processor.run()
            self.last_nbf_dir = result.get("nbf_dir")
            if self.last_nbf_dir and os.path.exists(self.last_nbf_dir):
                self.after(0, lambda: self.nbf_btn.config(state=NORMAL))
            self.after(0, lambda: messagebox.showinfo("Success", "Process completed successfully!"))
        except Exception as e:
            self.logger.log(f"CRITICAL ERROR: {e}", level="ERROR")
            self.after(0, lambda err=str(e): messagebox.showerror("Execution Error", err))
        finally:
            self.after(0, lambda: self.run_btn.config(state=NORMAL))
