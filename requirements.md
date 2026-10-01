# Requirements Document

## Introduction

This document specifies the requirements for adding date picker functionality to the Learner Payment Management System. The system currently has both a GUI application (using CustomTkinter) and a web application (using Flask). Currently, payment dates are automatically set to the current timestamp when recording payments. This feature will allow users to manually select payment dates when recording historical payments or backdating transactions.

## Glossary

- **Date_Picker**: A UI component that allows users to select dates from a calendar interface
- **GUI_Application**: The desktop application built with CustomTkinter (app_gui.py)
- **Web_Application**: The browser-based application built with Flask (app_web.py)
- **Payment_Record**: A domain entity representing a payment transaction with amount, timestamp, and admission number
- **Payment_Service**: The application service responsible for recording payments
- **Historical_Payment**: A payment that occurred on a date other than the current date

## Requirements

### Requirement 1: GUI Date Picker Component

**User Story:** As a school administrator using the desktop application, I want to select a payment date from a calendar widget, so that I can record historical payments accurately.

#### Acceptance Criteria

1. WHEN the payment recording form loads, THE GUI_Application SHALL display a date picker widget before the payment amount field
2. WHEN the date picker is displayed, THE GUI_Application SHALL show day, month, and year components
3. WHEN a user clicks the date picker widget, THE GUI_Application SHALL display a calendar interface for date selection
4. WHEN a user selects a date from the calendar, THE GUI_Application SHALL populate the date picker field with the selected date
5. WHEN no date is selected, THE GUI_Application SHALL default to the current date
6. WHEN a user selects a future date, THE GUI_Application SHALL reject the selection and display an error message

### Requirement 2: Web Date Picker Component

**User Story:** As a school administrator using the web application, I want to select a payment date using an HTML5 date input, so that I can record historical payments from any browser.

#### Acceptance Criteria

1. WHEN the payment recording form loads, THE Web_Application SHALL display an HTML5 date input field before the payment amount field
2. WHEN the date input is displayed, THE Web_Application SHALL show day, month, and year in the browser's native format
3. WHEN a user clicks the date input field, THE Web_Application SHALL display the browser's native date picker
4. WHEN a user selects a date, THE Web_Application SHALL capture the selected date value
5. WHEN no date is selected, THE Web_Application SHALL default to the current date
6. WHEN a user selects a future date, THE Web_Application SHALL reject the selection and display an error message

### Requirement 3: Payment Recording with Custom Dates

**User Story:** As a school administrator, I want to record payments with custom dates, so that I can accurately track when payments were actually made.

#### Acceptance Criteria

1. WHEN a user submits a payment with a selected date, THE Payment_Service SHALL create a Payment_Record with the specified timestamp
2. WHEN a user submits a payment without selecting a date, THE Payment_Service SHALL create a Payment_Record with the current timestamp
3. WHEN a payment is recorded with a custom date, THE Payment_Service SHALL validate that the date is not in the future
4. WHEN a payment is recorded with a custom date, THE Payment_Service SHALL persist the payment with the correct timestamp to the spreadsheet

### Requirement 4: Date Validation

**User Story:** As a system, I want to validate payment dates, so that data integrity is maintained.

#### Acceptance Criteria

1. WHEN a payment date is provided, THE Payment_Service SHALL validate that the date is not in the future
2. WHEN a payment date is in the future, THE Payment_Service SHALL reject the payment and return an error message
3. WHEN a payment date is valid, THE Payment_Service SHALL accept the payment for processing
4. WHEN a payment date is not provided, THE Payment_Service SHALL treat it as the current date

### Requirement 5: Display Payment Dates

**User Story:** As a school administrator, I want to see the exact date and time of each payment, so that I can verify payment history.

#### Acceptance Criteria

1. WHEN displaying payment history in the GUI, THE GUI_Application SHALL show each payment's timestamp in the format "YYYY-MM-DD HH:MM:SS"
2. WHEN displaying payment history in the web interface, THE Web_Application SHALL show each payment's timestamp in the format "YYYY-MM-DD HH:MM:SS"
3. WHEN displaying the dashboard, THE GUI_Application SHALL continue to show payment counts without individual timestamps
4. WHEN displaying the dashboard, THE Web_Application SHALL continue to show payment counts without individual timestamps

### Requirement 6: Backward Compatibility

**User Story:** As a system maintainer, I want the date picker feature to be backward compatible, so that existing functionality is not disrupted.

#### Acceptance Criteria

1. WHEN the date picker is not used, THE Payment_Service SHALL record payments with the current timestamp as before
2. WHEN existing payment records are loaded, THE GUI_Application SHALL display them correctly with their existing timestamps
3. WHEN existing payment records are loaded, THE Web_Application SHALL display them correctly with their existing timestamps
4. WHEN the spreadsheet is saved, THE Payment_Service SHALL maintain the existing timestamp format for all payment records
