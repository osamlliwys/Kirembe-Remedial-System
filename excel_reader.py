"""ExcelReader for parsing Excel spreadsheet data into domain objects."""

import logging
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import List, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from src.domain.learner import Learner
from src.domain.payment_record import PaymentRecord
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ExcelReader:
    """
    Handles reading and parsing Excel spreadsheet data.
    
    The Excel file is expected to have the following structure:
    - Column A: Learner Name
    - Column B: Admission Number
    - Column C: Total Owed
    - Columns D+: Payment records (amount and date pairs)
    """
    
    def read_spreadsheet(self, file_path: str):
        """
        Load Excel workbook from the specified file path.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            The workbook object containing all sheets
            
        Raises:
            FileNotFoundError: If the file doesn't exist
            Exception: If the file cannot be read or is invalid
        """
        try:
            logger.info(f"Attempting to read spreadsheet: {file_path}")
            workbook = load_workbook(file_path)
            logger.info(f"Successfully loaded spreadsheet with {len(workbook.sheetnames)} sheets")
            return workbook
        except FileNotFoundError:
            error_msg = (
                f"Cannot find payment database file: {file_path}. "
                "Please ensure the file exists in the correct location."
            )
            logger.error(error_msg)
            raise FileNotFoundError(error_msg)
        except PermissionError as e:
            error_msg = (
                f"Cannot access payment database file: Permission denied. "
                f"The file '{file_path}' may be open in another program. "
                "Please close the file and try again."
            )
            logger.error(f"{error_msg} - {str(e)}")
            raise Exception(error_msg)
        except Exception as e:
            error_msg = (
                f"Cannot read payment database file: {str(e)}. "
                "The file may be corrupted or in an unsupported format. "
                "Please ensure it is a valid Excel (.xlsx) file."
            )
            logger.error(error_msg)
            raise Exception(error_msg)
    
    def parse_learner_row(self, row: Tuple, sheet_name: str = "", row_number: int = 0, arrears_override=None) -> Optional[Learner]:
        """
        Convert a spreadsheet row to a Learner object.
        
        The actual spreadsheet structure is:
        - Column A (index 0): NO. (row number)
        - Column B (index 1): ADM (admission number)
        - Column C (index 2): NAME
        - Column D (index 3): ARREAS (not used)
        - Column E (index 4): DEBT (total owed)
        - Columns H+ (index 7+): Payment records
        
        FORM 4 2026 has a different structure:
        - Column C (index 2): NO. (row number)
        - Column D (index 3): ADM (admission number)
        - Column E (index 4): NAME
        - Column F (index 5): ARREAS (not used)
        - Column G (index 6): DEBT (total owed)
        - Columns J+ (index 9+): Payment records
        
        Args:
            row: A tuple representing a row from the spreadsheet
            sheet_name: Name of the sheet being processed (for special handling)
            row_number: The row number in the Excel sheet (for stream detection)
            
        Returns:
            A Learner object if the row is valid, None if the row should be skipped
            
        Note:
            Handles missing or malformed data gracefully by returning None
            for invalid rows. Logs warnings for data quality issues.
            
        Error Handling:
            - Returns None for empty rows or rows with insufficient columns
            - Returns None for rows with missing required fields (name, admission number, total owed)
            - Returns None for rows with negative total owed amounts
            - Logs warnings for all data quality issues to aid debugging
        """
        # Check if this is Form 4 2026 with different structure
        is_form4_2026 = "FORM 4 2026" in sheet_name.upper() or "FORM 4A" in sheet_name.upper()
        
        if is_form4_2026:
            # Form 4 2026 structure: columns shifted by 2
            # Skip empty rows or rows with insufficient data
            if not row or len(row) < 7:
                return None
            
            # Extract fields from Form 4 2026 column positions
            admission_number = row[3]  # Column D: Admission number
            name = row[4]              # Column E: Learner's full name
            
            # Calculate cumulative balance: Column F (2024 arrears) + Column G (2025 debt)
            arrears_2024 = row[5] if row[5] is not None else 0  # Column F: 2024 arrears
            debt_2025 = row[6] if row[6] is not None else 0     # Column G: 2025 debt
            
            # Skip rows where arrears or debt are formulas (these are total rows)
            if isinstance(arrears_2024, str) or isinstance(debt_2025, str):
                return None
            
            total_owed_raw = float(arrears_2024) + float(debt_2025)
            
            payment_start_index = 9    # Payments start at column J
        else:
            # Standard structure
            # Skip empty rows or rows with insufficient data
            # Minimum 5 columns required: NO, ADM, NAME, ARREAS, DEBT
            if not row or len(row) < 5:
                return None
            
            # Extract fields from correct columns based on spreadsheet structure
            admission_number = row[1]  # Column B: Admission number (unique identifier)
            name = row[2]              # Column C: Learner's full name
            
            # For FORM 3 2026, Column D (arrears) contains formulas referencing FORM 2 2025
            # We should use only Column E (current debt) as the total_owed
            # For other sheets, we sum Column D (arrears) + Column E (debt)
            if "FORM 3 2026" in sheet_name:
                # FORM 3 2026: total_owed = prior-year arrears (from FORM 2 2025) + current-year debt
                # Column D holds a formula referencing 'FORM 2 2025'!T{row} which openpyxl
                # cannot evaluate. The caller passes arrears_override (computed from FORM 2 2025).
                debt_2026 = row[4] if row[4] is not None else 0  # Column E: current year debt (2200)
                
                if isinstance(debt_2026, str):
                    return None
                
                prior_arrears = float(arrears_override) if arrears_override is not None else 0.0
                total_owed_raw = prior_arrears + float(debt_2026)
            else:
                # Other sheets: Calculate cumulative balance
                arrears_2024 = row[3] if row[3] is not None else 0  # Column D: 2024 arrears
                debt_2025 = row[4] if row[4] is not None else 0     # Column E: 2025 debt
                
                # Skip rows where arrears or debt are formulas (these are total rows)
                if isinstance(arrears_2024, str) or isinstance(debt_2025, str):
                    return None
                
                total_owed_raw = float(arrears_2024) + float(debt_2025)
            
            payment_start_index = 7    # Payments start at column H
        
        # Validate required fields are present
        # All three fields must have values for a valid learner record
        if not name or not admission_number or total_owed_raw is None:
            return None
        
        # Convert name and admission number to strings and remove whitespace
        # This handles cases where Excel stores numbers as numeric types
        name = str(name).strip()
        admission_number = str(admission_number).strip()
        
        # Skip if name or admission number is empty after stripping whitespace
        # This catches cells that contain only spaces or formatting
        if not name or not admission_number:
            return None
        
        # Parse and validate total_owed amount
        # Must be a valid decimal number and non-negative
        try:
            total_owed = Decimal(str(total_owed_raw))
            if total_owed < 0:
                logger.warning(
                    f"Skipping learner {admission_number} ({name}): "
                    f"Negative total owed amount ({total_owed})"
                )
                return None
        except (InvalidOperation, ValueError, TypeError) as e:
            logger.warning(
                f"Skipping learner {admission_number} ({name}): "
                f"Invalid total owed value '{total_owed_raw}' - {str(e)}"
            )
            return None
        
        # Extract payment history from payment columns
        # Payment data is stored as alternating amount/date pairs in subsequent columns
        payment_records = self.extract_payment_history(row[payment_start_index:], admission_number)
        
        # Determine the stream (A or C) based on row number and sheet
        stream_suffix = self._determine_stream(sheet_name, row_number)
        full_sheet_name = f"{sheet_name} {stream_suffix}" if stream_suffix else sheet_name
        
        logger.debug(
            f"Parsed learner from {full_sheet_name}: {name} ({admission_number}), "
            f"Total owed: {total_owed}, Payments: {len(payment_records)}"
        )
        
        # Create and return the Learner domain object
        return Learner(
            name=name,
            admission_number=admission_number,
            total_owed=total_owed,
            payment_records=payment_records,
            source_sheet=full_sheet_name,
            grade=full_sheet_name
        )
    
    def extract_payment_history(
        self,
        payment_columns: Tuple,
        admission_number: str
    ) -> List[PaymentRecord]:
        """
        Parse payment columns to extract payment records.
        
        The REMEDIAL_PAYMENT_BALANCES file stores amounts only (no dates).
        Dates are stored separately in PAYMENT_DATES_LOG.xlsx.
        
        This method handles both:
        - New format: amount only per column
        - Old format: (amount, date) pairs (for backward compatibility)
        
        Args:
            payment_columns: Tuple of values from payment columns onwards
            admission_number: The admission number to associate with payments
            
        Returns:
            List of PaymentRecord objects parsed from the columns
        """
        payment_records = []
        now = datetime.now()
        
        i = 0
        while i < len(payment_columns):
            amount_raw = payment_columns[i]
            
            # Skip None values
            if amount_raw is None:
                i += 1
                continue
            
            # Try to parse as amount
            try:
                amount = Decimal(str(amount_raw))
                if amount <= 0:
                    i += 1
                    continue
            except (InvalidOperation, ValueError, TypeError):
                i += 1
                continue
            
            # Check if next column is a date (old format) or another amount/None (new format)
            timestamp = now  # default timestamp
            next_is_date = False
            
            if i + 1 < len(payment_columns):
                next_val = payment_columns[i + 1]
                if isinstance(next_val, datetime):
                    timestamp = next_val
                    next_is_date = True
                elif isinstance(next_val, str) and next_val.strip():
                    parsed = self._parse_date_string(next_val.strip())
                    if parsed:
                        timestamp = parsed
                        next_is_date = True
            
            payment_record = PaymentRecord(
                amount=amount,
                timestamp=timestamp,
                admission_number=admission_number
            )
            payment_records.append(payment_record)
            
            # Advance: skip date column if it was a date, otherwise just move one
            i += 2 if next_is_date else 1
        
        return payment_records
    
    def _parse_date_string(self, date_str: str) -> Optional[datetime]:
        """
        Parse a date string into a datetime object.
        
        Tries multiple common date formats to handle various date representations
        that might appear in the spreadsheet.
        
        Args:
            date_str: String representation of a date
            
        Returns:
            datetime object if parsing succeeds, None otherwise
            
        Supported Formats:
            - ISO format: "2024-01-15 14:30:00" or "2024-01-15"
            - DD/MM/YYYY: "15/01/2024"
            - MM/DD/YYYY: "01/15/2024"
            - DD-MM-YYYY: "15-01-2024"
            - MM-DD-YYYY: "01-15-2024"
        """
        # List of date formats to try, in order of likelihood
        date_formats = [
            "%Y-%m-%d %H:%M:%S",  # ISO datetime with time
            "%Y-%m-%d",            # ISO date only
            "%d/%m/%Y",            # Day/Month/Year with slashes
            "%m/%d/%Y",            # Month/Day/Year with slashes
            "%d-%m-%Y",            # Day-Month-Year with dashes
            "%m-%d-%Y",            # Month-Day-Year with dashes
        ]
        
        date_str = date_str.strip()
        
        # Try each format until one succeeds
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                # This format didn't match, try the next one
                continue
        
        # None of the formats matched
        return None

    def _determine_stream(self, sheet_name: str, row_number: int) -> str:
        """
        Determine which stream (A for Achievers or C for Champions) a learner belongs to
        based on their row number and sheet name.
        
        Args:
            sheet_name: Name of the Excel sheet
            row_number: Row number in the Excel sheet
            
        Returns:
            Stream suffix like "- Achievers" or "- Champions", or empty string if not applicable
        """
        # Define row ranges for each sheet and stream
        # Based on complete analysis of Excel structure covering ALL learners
        stream_ranges = {
            'GRADE 10 2026': {
                (6, 58): '- Achievers',     # Grade 10A (rows 6-58, includes all Achievers)
                (73, 100): '- Champions',   # Grade 10C (rows 73-100)
            },
            'FORM 2 2025': {
                (5, 75): '- Achievers',     # Form 2A (rows 5-75, excluding row 39 which is BALANCE B/D)
                (104, 171): '- Champions',  # Form 2C (rows 104-171)
            },
            'FORM 3 2026': {
                (5, 75): '- Achievers',     # Form 3A (rows 5-75, excluding row 39 which is BALANCE B/D)
                (104, 171): '- Champions',  # Form 3C (rows 104-171)
            },
            'FORM 3 2025': {
                (5, 66): '- Achievers',     # Form 3A (rows 5-66)
                (75, 124): '- Champions',   # Form 3C (rows 75-124)
            },
            'FORM 4 2026': {
                (5, 68): '- Achievers',     # Form 4A (rows 5-68)
                (76, 125): '- Champions',   # Form 4C (rows 76-125)
            },
        }
        
        # Check if this sheet has stream divisions
        if sheet_name in stream_ranges:
            ranges = stream_ranges[sheet_name]
            for (start_row, end_row), stream_suffix in ranges.items():
                if start_row <= row_number <= end_row:
                    return stream_suffix
        
        # No stream division found, return empty string
        return ""
