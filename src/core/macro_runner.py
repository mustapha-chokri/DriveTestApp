"""Excel COM automation and Macro execution for DriveTestApp."""
import os
from time import sleep
from .exceptions import MacroExecutionError

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

class MacroRunner:
    """Handles Excel COM automation: running VBA macros and extracting NBF files."""

    def __init__(self, logger=None):
        self.logger = logger

    def _log(self, message, level="INFO"):
        if self.logger:
            self.logger.log(message, level=level)

    def is_available(self):
        """Checks if win32com and Excel COM automation are supported on this environment."""
        return win32com is not None and pythoncom is not None

    def execute_and_export_nbf(self, excel_path, macro_name, name_identifier, output_dir):
        """
        Opens the workbook via Excel COM, runs the specified macro,
        extracts range O1:O13 for 2G, 3G, 4G, and writes .NBF files.
        """
        if not self.is_available():
            raise MacroExecutionError(
                "pywin32 / win32com is not installed or not available on this platform."
            )

        self._log("Launching Excel COM automation in background...")
        pythoncom.CoInitialize()

        excel_app = None
        wb = None
        nbf_dir = os.path.join(output_dir, "NBF_Output")
        os.makedirs(nbf_dir, exist_ok=True)

        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            sleep(1)
            abs_path = os.path.abspath(excel_path)
            self._log(f"Opening workbook in Excel: {abs_path}")
            wb = excel_app.Workbooks.Open(abs_path)

            self._log(f"Running VBA macro: {macro_name}...")
            try:
                excel_app.Run(macro_name)
                self._log("VBA macro executed successfully.")
            except Exception as e:
                self._log(f"Warning during macro execution: {e}", level="WARNING")

            sheets_to_process = ["2G", "3G", "4G"]
            for sheet_name in sheets_to_process:
                try:
                    sheet = wb.Sheets(sheet_name)
                    range_vals = sheet.Range("O1:O13").Value
                    filepath = os.path.join(nbf_dir, f"{name_identifier}_{sheet_name}.NBF")
                    with open(filepath, "w", encoding="utf-8") as f:
                        if range_vals:
                            for row in range_vals:
                                val = row[0]
                                f.write(str(val) if val is not None else "")
                                f.write("\n")
                    self._log(f"Generated NBF: {os.path.basename(filepath)}")
                except Exception as ex:
                    self._log(f"Error processing NBF for sheet '{sheet_name}': {ex}", level="WARNING")

            wb.Save()
            self._log("Workbook saved successfully.")
            return nbf_dir
        except Exception as e:
            self._log(f"COM execution failed: {e}", level="ERROR")
            raise MacroExecutionError(f"Failed executing Excel macro or exporting NBF: {e}") from e
        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=True)
                except Exception:
                    pass
            if excel_app:
                try:
                    excel_app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
