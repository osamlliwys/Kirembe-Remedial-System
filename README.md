# Learner Payment Management System

A comprehensive payment tracking system for Kirembe Secondary School that manages learner payment records across multiple grade levels.

## Features

- **Multi-Sheet Support**: Automatically loads and merges learner data from all grade levels (Grade 10, Form 2-4)
- **Payment Record Merging**: Combines payment history for students progressing through grades
- **Real-Time Updates**: Dashboard refreshes immediately after payment recording
- **Search Functionality**: Quick learner lookup by admission number
- **Payment Tracking**: Complete payment history with chronological ordering
- **Balance Calculation**: Automatic calculation of outstanding balances with credit/overpayment handling
- **Dual Interface**: Modern Windows GUI and traditional CLI options
- **Professional UI**: CustomTkinter-based desktop application with tabbed interface
- **Comprehensive Error Handling**: Clear, actionable error messages
- **Input Validation**: Sanitizes and validates all user inputs

## System Requirements

- Python 3.7 or higher
- Windows, macOS, or Linux
- Excel file with learner payment data

## Installation

1. **Clone or download the project** to your local machine

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

   Required packages:
   - openpyxl (Excel file operations)
   - pandas (data manipulation)
   - pytest (testing framework)
   - hypothesis (property-based testing)
   - customtkinter (modern GUI framework)
   - pillow (image processing)

3. **Prepare your data file**:
   - Place your Excel file in the project root directory
   - Default filename: `REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx`
   - To use a different file, update the `SPREADSHEET_FILE` variable in `app.py`

## Spreadsheet Format Requirements

The system expects an Excel workbook with the following structure:

### Sheet Structure
- Multiple sheets supported (one per grade/form)
- Each sheet should have:
  - Header rows (rows 1-4): School name, title, column headers
  - Data rows (starting row 5 or 6): Learner records

### Column Layout
- **Column A**: Row number (optional)
- **Column B**: Admission Number (required)
- **Column C**: Learner Name (required)
- **Column D**: Arrears (optional)
- **Column E**: Total Owed/Debt (required)
- **Columns H+**: Payment records (amount and date pairs)

### Example Sheet Structure
```
Row 1: KIREMBE SECONDARY SCHOOL
Row 2: REMEDIAL PAYMENTS GRADE 10
Row 3: NO. | ADM | NAME | ARREAS | DEBT
Row 4: (empty or additional headers)
Row 5+: Data rows with learner information
```

## Usage

### Choosing Your Interface

The system provides two interfaces:

1. **GUI Application (Recommended)**: Modern Windows desktop application
   ```bash
   python app_gui.py
   ```

2. **CLI Application**: Traditional command-line interface
   ```bash
   python app.py
   ```

**Not sure which to use?** See `INTERFACE_COMPARISON.md` for a detailed comparison.

### GUI Application Features

The GUI provides a modern, tabbed interface with three main sections:

#### Dashboard Tab
- View all learners in a scrollable table
- Color-coded balances (green for paid/credit, red for high balances)
- Real-time refresh button
- Shows: Name, Admission #, Total Owed, Total Paid, Balance, Payment Count

#### Search Learner Tab
- Enter admission number to search
- View complete learner details
- See full payment history (most recent first)
- Color-coded balance display with credit/overpayment indicators

#### Record Payment Tab
- Lookup learner by admission number
- View current balance before recording payment
- Enter payment amount with validation
- Instant confirmation and dashboard refresh

### Starting the GUI Application

**Option 1: Windows Batch File (Easiest)**

Double-click `start_gui.bat` in the project folder.

**Option 2: Command Line**

Run the GUI from the command line:

```bash
python app_gui.py
```

Or on Windows:

```bash
py app_gui.py
```

The application will:
1. Display the Kirembe Secondary School header
2. Load all learner data from the Excel file
3. Show a confirmation with the number of learners loaded
4. Open the Dashboard tab with all learners displayed

See `GUI_GUIDE.md` for detailed instructions on using the GUI application.

### Starting the CLI Application

**Option 1: Windows Batch File (Easiest)**

Double-click `start_cli.bat` in the project folder.

**Option 2: Command Line**

Run the application from the command line:

```bash
python app.py
```

Or on Windows:

```bash
py app.py
```

### Main Menu Options

The application presents a menu with four options:

```
1. Search learner by admission number
2. Record payment
3. View all learners
4. Exit
```

### 1. Search for a Learner

- Select option `1`
- Enter the admission number (e.g., `1710`)
- View complete learner details:
  - Name and admission number
  - Total owed, total paid, and current balance
  - Complete payment history in chronological order

### 2. Record a Payment

- Select option `2`
- Enter the admission number
- Enter the payment amount (supports formats like `5000`, `KES 5000`, `5,000.50`)
- View confirmation with updated balance and recent payment history

### 3. View All Learners

- Select option `3`
- See a comprehensive table with all learners:
  - Name
  - Admission number
  - Total owed
  - Total paid
  - Current balance
  - Number of payments

### 4. Exit

- Select option `4` to close the application

## Features in Detail

### Multi-Sheet Data Loading

The system automatically:
- Loads data from all sheets in the Excel workbook
- Merges payment records for students appearing in multiple sheets
- Uses the most recent `total_owed` value for each student
- Maintains complete payment history across all grade levels

Example: A student in both "FORM 2 2025" and "FORM 3 2026" will have:
- Payment records from both sheets combined
- Total owed from the most recent sheet (FORM 3 2026)
- Complete chronological payment history

### Currency Handling

- All amounts displayed in KES (Kenya Shillings)
- Supports flexible input formats:
  - `5000`
  - `KES 5000`
  - `5,000.50`
  - `Ksh 5000`
- Automatic sanitization removes currency symbols and commas

### Input Validation

The system validates:
- **Admission numbers**: Must be non-empty and valid format
- **Payment amounts**: 
  - Must be positive (> 0)
  - Maximum 1,000,000 KES
  - Valid decimal format
  - Rejects zero or negative values

### Error Handling

Comprehensive error messages for:
- File not found or inaccessible
- Invalid spreadsheet format
- Duplicate admission numbers
- Learner not found
- Invalid payment amounts
- File permission issues

## Project Structure

```
learner-payment-management/
├── app.py                          # CLI application entry point
├── app_gui.py                      # GUI application entry point (NEW)
├── start_gui.bat                   # Windows launcher for GUI (NEW)
├── start_cli.bat                   # Windows launcher for CLI (NEW)
├── requirements.txt                # Python dependencies
├── pytest.ini                      # Test configuration
├── README.md                       # This file
├── GUI_GUIDE.md                   # GUI application guide (NEW)
├── TESTING_GUIDE.md               # Testing instructions
├── src/                           # Source code
│   ├── __init__.py
│   ├── domain/                    # Business logic
│   │   ├── learner.py            # Learner entity
│   │   ├── payment_record.py     # Payment record entity
│   │   ├── payment_calculator.py # Calculation utilities
│   │   ├── payment_validator.py  # Payment validation
│   │   └── admission_number_validator.py
│   ├── data/                      # Data access layer
│   │   ├── excel_reader.py       # Excel reading
│   │   ├── excel_writer.py       # Excel writing
│   │   └── spreadsheet_repository.py
│   ├── application/               # Application services
│   │   ├── learner_service.py    # Learner operations
│   │   └── payment_service.py    # Payment operations
│   ├── presentation/              # User interface
│   │   └── cli.py                # Command-line interface
│   └── utils/                     # Utilities
│       └── logger.py             # Logging configuration
├── tests/                         # Test suite
│   ├── test_learner.py
│   ├── test_payment_record.py
│   ├── test_payment_calculator.py
│   ├── test_excel_reader.py
│   ├── test_excel_writer.py
│   ├── test_spreadsheet_repository.py
│   ├── test_learner_service.py
│   ├── test_payment_service.py
│   └── test_cli.py
└── data/                          # Data directory
    └── (Excel files go here)
```

## Testing

Run the test suite:

```bash
pytest
```

Run with coverage:

```bash
pytest --cov=src --cov-report=html
```

The test suite includes:
- Unit tests for all components
- Property-based tests for correctness properties
- Integration tests for complete workflows
- 150+ tests with 84%+ pass rate

## Troubleshooting

### "Cannot find payment database file"

- Verify the Excel file exists in the project root directory
- Check the filename matches `SPREADSHEET_FILE` in `app.py`
- Ensure you have read permissions for the file

### "Python not found" or "py not found"

- Install Python from https://python.org/downloads/
- During installation, check "Add Python to PATH"
- Restart your terminal after installation

### "Invalid spreadsheet format"

- Verify your Excel file has the required columns (B: ADM, C: NAME, E: DEBT)
- Ensure data starts at row 5 or 6
- Check that at least one sheet has valid data

### "Duplicate admission numbers found"

- Review the error message for specific duplicate entries
- Check your Excel file for duplicate admission numbers
- Decide if duplicates should be merged or corrected

## Data Backup

**Important**: The system writes payment records back to the Excel file. Always:
- Keep a backup of your original Excel file
- Test with a copy before using with production data
- Regularly backup your data file

## Support

For issues or questions:
1. Check the TESTING_GUIDE.md for common scenarios
2. Review error messages for specific guidance
3. Verify your Excel file format matches requirements

## License

This project is created for Kirembe Secondary School by Pambu A.W.
copyright reserved

## Version

Version 1.0.0 - March 2026
