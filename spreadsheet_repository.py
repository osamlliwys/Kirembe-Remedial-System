"""SpreadsheetRepository for managing all interactions with the Excel file."""

import logging
from typing import List, Dict, Optional
from src.domain.learner import Learner
from src.data.excel_reader import ExcelReader
from src.data.excel_writer import ExcelWriter
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class SpreadsheetRepository:
    """
    Manages all interactions with the Excel spreadsheet database.
    
    This repository provides methods to load learners from the spreadsheet,
    save learner updates, and validate the file format.
    """
    
    def __init__(self, file_path: str):
        """
        Initialize the repository with a file path.
        
        Args:
            file_path: Path to the Excel spreadsheet file
        """
        self.file_path = file_path
        self.excel_reader = ExcelReader()
        self.excel_writer = ExcelWriter()
    
    def _compute_form2_2025_balances(self, workbook) -> Dict[int, float]:
        """
        Compute FORM 2 2025 balance for each row.

        FORM 3 2026 column D references 'FORM 2 2025'!T{row} (prior-year balance).
        Since openpyxl cannot evaluate formulas, we compute these manually:
            balance = arrears(D) + debt(E) - sum(payments H:S)

        Returns:
            Dict mapping Excel row number -> computed balance (float)
        """
        balances = {}
        if 'FORM 2 2025' not in workbook.sheetnames:
            return balances

        ws = workbook['FORM 2 2025']
        try:
            for row_idx, row in enumerate(
                ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True), start=5
            ):
                # Skip blank or non-learner rows
                if not row or len(row) < 5:
                    continue
                adm = row[1]
                if not adm or not str(adm).strip().isdigit():
                    continue

                arrears_raw = row[3]  # col D — prior-year arrears (numeric or None)
                debt_raw = row[4]     # col E — 2025 debt (numeric)

                arrears = float(arrears_raw) if isinstance(arrears_raw, (int, float)) else 0.0
                debt    = float(debt_raw)    if isinstance(debt_raw,    (int, float)) else 0.0

                # Payments are in columns H-S (indices 7-18)
                total_paid = sum(
                    float(v) for v in row[7:19]
                    if isinstance(v, (int, float)) and v > 0
                )

                balances[row_idx] = arrears + debt - total_paid

        except Exception as e:
            logger.warning(f"Could not compute FORM 2 2025 balances: {e}")

        return balances

    def load_all_learners(self) -> Dict[str, any]:
        """
        Load all learner data from all sheets in the spreadsheet.
        
        When a learner appears in multiple sheets (same admission number),
        their payment records are merged and the most recent total_owed is used.
        
        Returns:
            Dictionary with:
                - 'success': bool indicating if operation succeeded
                - 'learners': List of Learner objects (if successful)
                - 'message': str with success/error message
                - 'duplicates': List of duplicate admission numbers (if any found)
                - 'sheets_processed': List of sheet names that were processed
        """
        try:
            logger.info(f"Loading learners from {self.file_path}")
            
            # Read the workbook
            workbook = self.excel_reader.read_spreadsheet(self.file_path)

            # Pre-compute FORM 2 2025 balances (needed as arrears for FORM 3 2026)
            form2_2025_balances = self._compute_form2_2025_balances(workbook)
            logger.info(f"Pre-computed {len(form2_2025_balances)} FORM 2 2025 row balances")
            
            # Dictionary to track learners by admission number
            # Using a dictionary allows us to efficiently merge records from multiple sheets
            # Key: admission_number (str), Value: Learner object
            learners_dict = {}
            sheets_processed = []
            merged_count = 0
            
            # Process each sheet in the workbook
            # Some spreadsheets have learner data split across multiple sheets
            # (e.g., by grade level or payment period)
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                logger.info(f"Processing sheet: {sheet_name}")
                
                sheet_learner_count = 0
                
                # Try both row 5 and row 6 as starting points
                # Some sheets have data at row 5, others at row 6
                # This handles variations in spreadsheet formatting
                start_row = 5
                
                # Check if row 5 has valid data, otherwise try row 6
                # We look for a numeric admission number in the expected columns
                first_row = None
                for row in worksheet.iter_rows(min_row=5, max_row=5, values_only=True):
                    first_row = row
                    break
                
                # If row 5 doesn't have admission number in expected columns, try row 6
                # This handles sheets where headers take up more rows
                if first_row and len(first_row) >= 3:
                    # Check columns B, C, D for admission number (numeric value)
                    has_valid_data = False
                    for col_idx in [1, 2, 3]:  # Columns B, C, D
                        if first_row[col_idx] and str(first_row[col_idx]).isdigit():
                            has_valid_data = True
                            break
                    
                    if not has_valid_data:
                        # No valid admission number found in row 5, try row 6
                        start_row = 6
                
                # Parse each row starting from determined start_row
                for row_idx, row in enumerate(worksheet.iter_rows(min_row=start_row, values_only=True), start=start_row):
                    # For FORM 3 2026, pass the pre-computed Form 2 2025 balance
                    # as the arrears override so total_owed = arrears + current_debt
                    arrears_override = None
                    if 'FORM 3 2026' in sheet_name and row_idx in form2_2025_balances:
                        arrears_override = form2_2025_balances[row_idx]
                    
                    learner = self.excel_reader.parse_learner_row(
                        row, sheet_name, row_idx, arrears_override=arrears_override
                    )
                    
                    if learner is not None:
                        admission_number = learner.admission_number
                        
                        if admission_number in learners_dict:
                            # Learner already exists from a previous sheet - merge the records
                            # This handles cases where payment data is split across multiple sheets
                            existing_learner = learners_dict[admission_number]
                            
                            # IMPORTANT: Prioritize FORM 3 2026 over FORM 2 2025
                            # Many learners appear in both sheets, but FORM 3 2026 has the correct current data
                            current_is_form3_2026 = "FORM 3 2026" in sheet_name
                            existing_is_form2_2025 = "FORM 2 2025" in existing_learner.source_sheet
                            
                            logger.debug(f"Merging {admission_number}: current={sheet_name}, existing={existing_learner.source_sheet}")
                            
                            # If current sheet is FORM 3 2026 and existing is FORM 2 2025, replace ALL data
                            if current_is_form3_2026 and existing_is_form2_2025:
                                existing_learner.source_sheet = learner.source_sheet
                                existing_learner.total_owed = learner.total_owed
                                existing_learner.name = learner.name
                                # Clear old payment records and use only FORM 3 2026 payments
                                existing_learner.payment_records = learner.payment_records.copy()
                                logger.info(f"Replaced {admission_number} with FORM 3 2026 data: {learner.source_sheet}, owed={learner.total_owed}")
                                merged_count += 1
                                continue  # Skip the payment merge below since we replaced everything
                            
                            # If current sheet is FORM 2 2025 and existing is already FORM 3 2026, skip it
                            current_is_form2_2025 = "FORM 2 2025" in sheet_name
                            existing_is_form3_2026 = "FORM 3 2026" in existing_learner.source_sheet
                            
                            if current_is_form2_2025 and existing_is_form3_2026:
                                logger.info(f"Skipping FORM 2 2025 data for {admission_number} - already have FORM 3 2026 data")
                                merged_count += 1
                                continue  # Don't merge anything - keep FORM 3 2026 data only
                            
                            # For other cases (same sheet, different years), update total_owed
                            existing_learner.total_owed = learner.total_owed
                            
                            # Merge payment records (add new payments to existing list)
                            # This combines payment history from all sheets
                            for payment in learner.payment_records:
                                existing_learner.payment_records.append(payment)
                            
                            # Sort payment records chronologically to maintain proper order
                            # This ensures payment history is always displayed in date order
                            existing_learner.payment_records.sort(key=lambda p: p.timestamp)
                            
                            merged_count += 1
                            logger.info(f"Merged payment records for {admission_number} from {sheet_name}")
                        else:
                            # New learner - add to dictionary
                            learners_dict[admission_number] = learner
                            sheet_learner_count += 1
                
                sheets_processed.append(f"{sheet_name}: {sheet_learner_count} new learners")
                logger.info(f"Loaded {sheet_learner_count} new learners from {sheet_name}")
            
            # Convert dictionary to list
            learners = list(learners_dict.values())
            
            success_msg = (
                f"Successfully imported {len(learners)} unique learners from "
                f"{len(workbook.sheetnames)} sheets"
            )
            if merged_count > 0:
                success_msg += f" ({merged_count} records merged across sheets)"
            
            logger.info(f"Total unique learners: {len(learners)}, Merged records: {merged_count}")
            
            return {
                'success': True,
                'learners': learners,
                'message': success_msg,
                'duplicates': [],
                'sheets_processed': sheets_processed
            }
            
        except FileNotFoundError as e:
            logger.error(f"File not found: {str(e)}")
            return {
                'success': False,
                'learners': [],
                'message': str(e),
                'duplicates': [],
                'sheets_processed': []
            }
        except Exception as e:
            error_msg = (
                f"Error loading spreadsheet: {str(e)}. "
                "Please ensure the file is not corrupted and is in the correct format."
            )
            logger.error(f"{error_msg} - Full error: {repr(e)}")
            return {
                'success': False,
                'learners': [],
                'message': error_msg,
                'duplicates': [],
                'sheets_processed': []
            }
    
    def save_learner(self, learner: Learner) -> Dict[str, any]:
        """
        Save a learner's updated information to the spreadsheet.
        
        Args:
            learner: The Learner object to save
            
        Returns:
            Dictionary with:
                - 'success': bool indicating if operation succeeded
                - 'message': str with success/error message
        """
        try:
            logger.info(f"Saving learner {learner.admission_number} to spreadsheet")
            result = self.excel_writer.write_learner(learner, self.file_path)
            
            if result['success']:
                logger.info(f"Successfully saved learner {learner.admission_number}")
            else:
                logger.error(f"Failed to save learner {learner.admission_number}: {result['message']}")
            
            return result
            
        except Exception as e:
            error_msg = (
                f"Error saving learner {learner.admission_number}: {str(e)}. "
                "An unexpected error occurred while saving to the spreadsheet."
            )
            logger.error(f"{error_msg} - Full error: {repr(e)}")
            return {
                'success': False,
                'message': error_msg
            }
    
    def validate_file_format(self) -> Dict[str, any]:
        """
        Validate that the spreadsheet has the expected structure.
        
        Checks:
        - File exists and can be opened
        - Has at least one sheet
        - Each sheet has data rows
        
        Returns:
            Dictionary with:
                - 'success': bool indicating if format is valid
                - 'message': str with validation result or error details
        """
        try:
            logger.info(f"Validating file format for {self.file_path}")
            
            # Try to read the workbook
            workbook = self.excel_reader.read_spreadsheet(self.file_path)
            
            # Check if workbook has at least one sheet
            if len(workbook.sheetnames) == 0:
                error_msg = (
                    "Payment database file has no sheets. "
                    "The file appears to be empty or corrupted."
                )
                logger.error(error_msg)
                return {
                    'success': False,
                    'message': error_msg
                }
            
            # Basic validation - just check that sheets exist and have some data
            for sheet_name in workbook.sheetnames:
                worksheet = workbook[sheet_name]
                
                # Check if worksheet has any data
                if worksheet.max_row < 5:
                    error_msg = (
                        f"Sheet '{sheet_name}' appears to be empty or has insufficient rows. "
                        f"Expected at least 5 rows, found {worksheet.max_row}."
                    )
                    logger.error(error_msg)
                    return {
                        'success': False,
                        'message': error_msg
                    }
            
            success_msg = f"Payment database file format is valid ({len(workbook.sheetnames)} sheets found)."
            logger.info(success_msg)
            
            return {
                'success': True,
                'message': success_msg
            }
            
        except FileNotFoundError as e:
            logger.error(f"File not found during validation: {str(e)}")
            return {
                'success': False,
                'message': str(e)
            }
        except Exception as e:
            error_msg = (
                f"Error validating file format: {str(e)}. "
                "The file may be corrupted or in an unsupported format."
            )
            logger.error(f"{error_msg} - Full error: {repr(e)}")
            return {
                'success': False,
                'message': error_msg
            }


    def update_learner(self, learner: Learner) -> Dict[str, any]:
        """
        Update a learner's information in the spreadsheet.

        This is an alias for save_learner() to maintain consistency
        with the payment correction service interface.

        Args:
            learner: The Learner object to update

        Returns:
            Dictionary with:
                - 'success': bool indicating if operation succeeded
                - 'message': str with success/error message
        """
        return self.save_learner(learner)


    def get_stream_totals(self) -> Dict[str, any]:
        """
        Get live payment totals for each class stream by summing actual learner records.

        This calculates totals dynamically from the in-memory learner data rather than
        reading fixed cells, so the chart updates immediately whenever a payment is recorded.

        Returns:
            Dictionary with:
                - 'success': bool indicating if operation succeeded
                - 'streams': List of dicts with stream_name, total_collected,
                             learner_count, total_owed
                - 'message': str with success/error message
        """
        try:
            result = self.load_all_learners()
            if not result['success']:
                return {'success': False, 'streams': [], 'message': result['message']}

            learners = result['learners']

            # Map source_sheet stream suffixes → friendly group names
            # Each entry: (stream_key, display_name, group_key)
            # group_key links cohorts that are the same students across years
            stream_config = [
                # Grade 10 / Form 2 cohort
                ('GRADE 10 2026 - Achievers', 'Grade 10A 2026', 'Grade 10 Achievers'),
                ('GRADE 10 2026 - Champions', 'Grade 10C 2026', 'Grade 10 Champions'),
                # Form 2 → Form 3 cohort (same students, different years)
                ('FORM 2 2025 - Achievers',   'Form 2A 2025',   'Form 3A Cohort'),
                ('FORM 3 2026 - Achievers',   'Form 3A 2026',   'Form 3A Cohort'),
                ('FORM 2 2025 - Champions',   'Form 2C 2025',   'Form 3C Cohort'),
                ('FORM 3 2026 - Champions',   'Form 3C 2026',   'Form 3C Cohort'),
                # Form 3 → Form 4 cohort (same students, different years)
                ('FORM 3 2025 - Achievers',   'Form 3A 2025',   'Form 4A Cohort'),
                ('FORM 4 2026 - Achievers',   'Form 4A 2026',   'Form 4A Cohort'),
                ('FORM 3 2025 - Champions',   'Form 3C 2025',   'Form 4C Cohort'),
                ('FORM 4 2026 - Champions',   'Form 4C 2026',   'Form 4C Cohort'),
            ]

            # Build totals per stream key
            stream_totals: Dict[str, dict] = {}
            for sk, display, group in stream_config:
                stream_totals[sk] = {
                    'stream_name': display,
                    'group_key': group,
                    'total_collected': 0.0,
                    'total_owed': 0.0,
                    'learner_count': 0,
                }

            # Also handle base sheet names (without stream suffix) as fallback
            sheet_fallback = {
                'GRADE 10 2026': 'Grade 10A 2026',
                'FORM 2 2025':   'Form 2A 2025',
                'FORM 3 2026':   'Form 3A 2026',
                'FORM 3 2025':   'Form 3A 2025',
                'FORM 4 2026':   'Form 4A 2026',
            }

            for learner in learners:
                source = getattr(learner, 'source_sheet', '') or ''
                total_paid = sum(
                    float(p.amount) for p in learner.payment_records
                    if p.amount and float(p.amount) > 0
                )
                total_owed = float(getattr(learner, 'total_owed', 0) or 0)

                # Match by exact stream key first
                matched = False
                for sk in stream_totals:
                    if sk.lower() in source.lower() or source.lower() in sk.lower():
                        stream_totals[sk]['total_collected'] += total_paid
                        stream_totals[sk]['total_owed'] += total_owed
                        stream_totals[sk]['learner_count'] += 1
                        matched = True
                        break

                if not matched:
                    # Fallback: match by base sheet name only
                    for base, display in sheet_fallback.items():
                        if base.lower() in source.lower():
                            for sk, data in stream_totals.items():
                                if data['stream_name'] == display:
                                    data['total_collected'] += total_paid
                                    data['total_owed'] += total_owed
                                    data['learner_count'] += 1
                                    break
                            break

            streams = [
                {
                    'stream_name': v['stream_name'],
                    'group_key': v['group_key'],
                    'total_collected': round(v['total_collected'], 2),
                    'total_owed': round(v['total_owed'], 2),
                    'learner_count': v['learner_count'],
                }
                for v in stream_totals.values()
                if v['learner_count'] > 0  # only include streams with actual learners
            ]

            # Sort: group cohorts together, then by year within group
            streams.sort(key=lambda x: (x['group_key'], x['stream_name']))

            logger.info(f"Live stream totals calculated for {len(streams)} streams")
            return {
                'success': True,
                'streams': streams,
                'message': f'Successfully calculated {len(streams)} stream totals'
            }

        except Exception as e:
            error_msg = f"Error calculating stream totals: {str(e)}"
            logger.error(f"{error_msg} - Full error: {repr(e)}")
            return {'success': False, 'streams': [], 'message': error_msg}
    
    def _evaluate_sum_formula(self, sheet, cell_ref: str) -> float:
        """
        Helper method to evaluate a SUM formula in a cell.
        
        Args:
            sheet: The worksheet object
            cell_ref: Cell reference (e.g., 'H59')
            
        Returns:
            The calculated sum as a float
        """
        cell = sheet[cell_ref]
        value = cell.value
        
        if value is None:
            return 0.0
        elif isinstance(value, (int, float)):
            return float(value)
        elif isinstance(value, str) and value.startswith('='):
            formula = value[1:]  # Remove '='
            
            # Check if formula has both + and - (complex expression)
            if '+' in formula and '-' in formula:
                return self._evaluate_subtraction(sheet, formula)
            
            # Check if there's only + for adding an initial balance cell
            elif '+' in formula:
                # Split by the last + to separate the main part from the addition
                parts = formula.rsplit('+', 1)
                if len(parts) == 2:
                    main_part = parts[0]
                    add_cell = parts[1].strip()
                    
                    # Evaluate the main part
                    if main_part.startswith('SUM('):
                        main_total = self._evaluate_simple_sum(sheet, main_part)
                    elif '-' in main_part:
                        main_total = self._evaluate_subtraction(sheet, main_part)
                    else:
                        main_total = 0.0
                    
                    # Add the additional cell value
                    try:
                        add_val = sheet[add_cell].value
                        if isinstance(add_val, (int, float)):
                            main_total += float(add_val)
                    except:
                        pass
                    
                    return main_total
            
            # Handle simple SUM formulas (no addition)
            elif formula.startswith('SUM('):
                return self._evaluate_simple_sum(sheet, formula)
            
            # Handle subtraction formulas (e.g., H101-H59)
            elif '-' in formula:
                return self._evaluate_subtraction(sheet, formula)
            
            return 0.0
        else:
            return 0.0
    
    def _evaluate_simple_sum(self, sheet, formula: str) -> float:
        """Evaluate a simple SUM formula."""
        if not formula.startswith('SUM('):
            return 0.0
        
        # Remove 'SUM(' and ')'
        inner = formula[4:-1] if formula.endswith(')') else formula[4:]
        total = 0.0
        
        # Check if it's a range (e.g., H40:H58) or list (e.g., H76,I76,...)
        if ':' in inner:
            # Range format
            parts = inner.split(':')
            if len(parts) == 2:
                start_cell = parts[0]
                end_cell = parts[1]
                
                # Extract column and row
                import re
                start_match = re.match(r'([A-Z]+)(\d+)', start_cell)
                end_match = re.match(r'([A-Z]+)(\d+)', end_cell)
                
                if start_match and end_match:
                    col = start_match.group(1)
                    start_row = int(start_match.group(2))
                    end_row = int(end_match.group(2))
                    
                    # Sum all cells in range
                    for row in range(start_row, end_row + 1):
                        cell_val = sheet[f'{col}{row}'].value
                        if isinstance(cell_val, (int, float)):
                            total += float(cell_val)
        else:
            # List format
            cell_refs = [c.strip() for c in inner.split(',')]
            for ref in cell_refs:
                try:
                    cell_val = sheet[ref].value
                    if isinstance(cell_val, (int, float)):
                        total += float(cell_val)
                except:
                    pass
        
        return total
    
    def _evaluate_subtraction(self, sheet, formula: str) -> float:
        """Evaluate a subtraction formula (e.g., H101-H59 or H101+J103-H59)."""
        # Handle formulas with both + and -
        # Example: H101+J103-H59
        
        total = 0.0
        current_value = 0.0
        current_op = '+'
        current_term = ''
        
        # Parse the formula character by character
        for char in formula + '+':  # Add a + at the end to process the last term
            if char in ['+', '-']:
                if current_term:
                    # Evaluate the current term
                    term = current_term.strip()
                    if term.startswith('SUM('):
                        term_val = self._evaluate_simple_sum(sheet, term)
                    else:
                        # It's a cell reference
                        try:
                            cell_val = sheet[term].value
                            if isinstance(cell_val, (int, float)):
                                term_val = float(cell_val)
                            elif isinstance(cell_val, str) and cell_val.startswith('='):
                                # Recursively evaluate
                                term_val = self._evaluate_sum_formula(sheet, term)
                            else:
                                term_val = 0.0
                        except:
                            term_val = 0.0
                    
                    # Apply the operation
                    if current_op == '+':
                        total += term_val
                    else:  # current_op == '-'
                        total -= term_val
                    
                    current_term = ''
                
                current_op = char
            else:
                current_term += char
        
        return total
