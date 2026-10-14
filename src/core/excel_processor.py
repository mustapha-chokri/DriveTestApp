"""Excel processing and data pipeline for DriveTestApp."""
import os
import pathlib
from openpyxl import load_workbook
from .exceptions import FileValidationError, SheetNotFoundError
from .macro_runner import MacroRunner

class ExcelProcessor:
    """Orchestrates data extraction from RNP files, injection into ADT templates, and macro execution."""

    def __init__(self, adt_path, rnp_path, output_dir, city_name, logger=None):
        self.adt_path = adt_path
        self.rnp_path = rnp_path
        self.output_dir = output_dir
        self.city_name = city_name
        self.logger = logger
        self.macro_runner = MacroRunner(logger=self.logger)
        self.last_nbf_dir = None
        self.generated_file_path = None

    def _log(self, message, level="INFO"):
        if self.logger:
            self.logger.log(message, level=level)

    def validate_inputs(self):
        """Validates existence of required template and source files."""
        if not self.adt_path or not os.path.exists(self.adt_path):
            raise FileValidationError(f"Template ADT file not found: {self.adt_path}")
        if not self.rnp_path or not os.path.exists(self.rnp_path):
            raise FileValidationError(f"Source RNP file not found: {self.rnp_path}")
        if not self.city_name or not self.city_name.strip():
            raise FileValidationError("Target city name cannot be empty.")
        if not self.output_dir:
            raise FileValidationError("Output directory cannot be empty.")
        
        os.makedirs(self.output_dir, exist_ok=True)

    def clear_range(self, ws):
        """Clears target rows (15 to 30) in the given worksheet."""
        ws.delete_rows(15, 16)
        self._log(f"Cleared rows 15-30 in sheet: {ws.title}")

    def run(self):
        """Executes the full pipeline: validation -> openpyxl transfer -> COM macro & NBF export."""
        self.validate_inputs()

        self._log("Starting processing pipeline...")
        self._log(f"Loading Template ADT: {os.path.basename(self.adt_path)}")
        wb_adt = load_workbook(self.adt_path, keep_vba=True)

        try:
            ws_adt_input = wb_adt['Input Data']
            ws_adt_2g = wb_adt['2G']
            ws_adt_3g = wb_adt['3G']
            ws_adt_4g = wb_adt['4G']
        except KeyError as e:
            wb_adt.close()
            raise SheetNotFoundError(f"Missing required sheet in Template ADT: {e}")

        self._log(f"Loading Source RNP: {os.path.basename(self.rnp_path)}")
        wb_rnp = load_workbook(self.rnp_path, data_only=True)

        ws_rnp_input = wb_rnp['Input Data'] if 'Input Data' in wb_rnp.sheetnames else None
        required_rnp_sheets = ['2G RNP', '3G RNP', '4G RNP']
        rnp_sheets = {s: wb_rnp[s] if s in wb_rnp.sheetnames else None for s in required_rnp_sheets}

        # Clear ranges in ADT
        self.clear_range(ws_adt_2g)
        self.clear_range(ws_adt_3g)
        self.clear_range(ws_adt_4g)

        # Copy Input Data
        if ws_rnp_input:
            self._log("Copying 'Input Data'...")
            for i in range(1, ws_rnp_input.max_row + 1):
                for j in range(1, ws_rnp_input.max_column + 1):
                    ws_adt_input.cell(row=i + 7, column=j).value = ws_rnp_input.cell(row=i, column=j).value

        # Set target city
        city_formatted = self.city_name.strip().upper()
        ws_adt_input['M2'] = city_formatted
        self._log(f"Target City set to: {city_formatted}")

        # Copy 2G, 3G, 4G sheets
        def copy_sheet_data(source_ws, target_ws, start_row_offset=14, max_col_override=None):
            if not source_ws:
                return
            self._log(f"Transferring data from '{source_ws.title}' to '{target_ws.title}'...")
            max_col = max_col_override if max_col_override else source_ws.max_column
            for i in range(1, source_ws.max_row + 1):
                for j in range(1, max_col + 1):
                    target_ws.cell(row=i + start_row_offset, column=j).value = source_ws.cell(row=i, column=j).value

        copy_sheet_data(rnp_sheets.get('2G RNP'), ws_adt_2g, max_col_override=12)
        copy_sheet_data(rnp_sheets.get('3G RNP'), ws_adt_3g)
        copy_sheet_data(rnp_sheets.get('4G RNP'), ws_adt_4g)

        # File naming
        rnp_stem = pathlib.Path(self.rnp_path).stem
        name_identifier = rnp_stem.replace("RNP_", "")
        new_file_path = os.path.join(self.output_dir, f"applic_DT_{name_identifier}.xlsm")

        self._log(f"Saving preliminary workbook: {new_file_path}")
        wb_adt.save(new_file_path)
        wb_adt.close()
        wb_rnp.close()

        self.generated_file_path = new_file_path

        # Run COM macro & export NBF
        self.last_nbf_dir = self.macro_runner.execute_and_export_nbf(
            excel_path=new_file_path,
            macro_name="PasteValues2G3G4G",
            name_identifier=name_identifier,
            output_dir=self.output_dir
        )

        self._log("Process pipeline completed successfully!")
        return {
            "output_file": self.generated_file_path,
            "nbf_dir": self.last_nbf_dir
        }
