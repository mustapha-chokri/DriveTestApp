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

        # Copy Input Data and normalize Site Azimuths (e.g., convert "40/100/270" -> "40;100;270")
        if ws_rnp_input:
            self._log("Copying 'Input Data'...")
            
            # Detect column for Site Azimuths
            azimuth_col = None
            for col_idx in range(1, ws_rnp_input.max_column + 1):
                header_val = ws_rnp_input.cell(row=1, column=col_idx).value
                if header_val and "AZIMUTH" in str(header_val).strip().upper():
                    azimuth_col = col_idx
                    break

            for i in range(1, ws_rnp_input.max_row + 1):
                for j in range(1, ws_rnp_input.max_column + 1):
                    val = ws_rnp_input.cell(row=i, column=j).value
                    # Format azimuths on data rows (row >= 2) or wherever slashes are present in azimuth column
                    if (j == azimuth_col or j == 5) and i >= 2 and val is not None:
                        val_str = str(val).strip()
                        if "/" in val_str:
                            parts = [p.strip() for p in val_str.split("/") if p.strip()]
                            cleaned_val = ";".join(parts)
                            self._log(f"Normalized Site Azimuths (row {i}): '{val}' -> '{cleaned_val}'")
                            val = cleaned_val
                    ws_adt_input.cell(row=i + 7, column=j).value = val

        # Ensure cell E9 in Input Data is normalized if it still contains slashes
        e9_val = ws_adt_input['E9'].value
        if e9_val and "/" in str(e9_val):
            parts = [p.strip() for p in str(e9_val).split("/") if p.strip()]
            cleaned_e9 = ";".join(parts)
            ws_adt_input['E9'].value = cleaned_e9
            self._log(f"Verified Input Data cell E9: '{e9_val}' -> '{cleaned_e9}'")

        # Set target city
        city_formatted = self.city_name.strip().upper()
        ws_adt_input['M2'] = city_formatted
        self._log(f"Target City set to: {city_formatted}")

        # Helper to copy sheet data
        def copy_sheet_data(source_ws, target_ws, start_row_offset=14, max_col_override=None):
            if not source_ws:
                return
            self._log(f"Transferring data from '{source_ws.title}' to '{target_ws.title}'...")
            max_col = max_col_override if max_col_override else source_ws.max_column
            for i in range(1, source_ws.max_row + 1):
                for j in range(1, max_col + 1):
                    target_ws.cell(row=i + start_row_offset, column=j).value = source_ws.cell(row=i, column=j).value

        # 2G RNP: Detect 'BSC' column and skip 'Sector' column if present
        ws_2g_rnp = rnp_sheets.get('2G RNP')
        if ws_2g_rnp:
            self._log("Processing '2G RNP'...")
            start_col_2g = 1
            for c in range(1, min(10, ws_2g_rnp.max_column + 1)):
                cell_val = ws_2g_rnp.cell(row=1, column=c).value
                if cell_val is not None and str(cell_val).strip().upper() == "BSC":
                    start_col_2g = c
                    break

            if start_col_2g > 1:
                first_col_val = ws_2g_rnp.cell(row=1, column=1).value
                self._log(
                    f"Detected '{first_col_val}' in column 1 of '2G RNP'. "
                    f"Skipping to 'BSC' at column {start_col_2g}."
                )

            # Copy 12 columns starting from 'BSC' into target ADT '2G' columns 1..12
            for i in range(1, ws_2g_rnp.max_row + 1):
                for offset in range(12):
                    src_c = start_col_2g + offset
                    tgt_c = 1 + offset
                    if src_c <= ws_2g_rnp.max_column:
                        ws_adt_2g.cell(row=i + 14, column=tgt_c).value = ws_2g_rnp.cell(row=i, column=src_c).value

        # 3G RNP and 4G RNP standard transfer
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
