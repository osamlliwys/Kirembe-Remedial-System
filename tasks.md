# Implementation Plan: Date Picker Integration

## Overview

This implementation plan adds date picker functionality to both the GUI (CustomTkinter) and web (Flask) applications. The implementation follows a layered approach, starting with the application layer (Payment Service), then the presentation layers (GUI and Web), and finally integration testing. Each task builds incrementally to ensure the feature works end-to-end.

## Tasks

- [ ] 1. Enhance Payment Service to accept optional payment date
  - [ ] 1.1 Modify `record_payment()` method signature to accept optional `payment_date` parameter
    - Add `payment_date: Optional[datetime] = None` parameter to method
    - Update method to use provided date or default to `datetime.now()`
    - Add validation to reject future dates
    - Return appropriate error message for future dates
    - _Requirements: 3.1, 3.2, 3.3, 4.1, 4.2, 4.3, 4.4_
  
  - [ ]* 1.2 Write property test for payment timestamp assignment
    - **Property 3: Payment Created with Specified Timestamp**
    - **Validates: Requirements 3.1**
  
  - [ ]* 1.3 Write property test for future date validation
    - **Property 4: Future Date Validation**
    - **Validates: Requirements 3.3, 4.1**
  
  - [ ]* 1.4 Write unit tests for Payment Service date handling
    - Test payment recording with custom date
    - Test payment recording without date (defaults to current)
    - Test future date rejection
    - Test error message for future dates
    - _Requirements: 3.1, 3.2, 3.3, 4.2_

- [ ] 2. Add date picker to GUI application
  - [ ] 2.1 Install tkcalendar library for DateEntry widget
    - Add `tkcalendar` to project dependencies
    - Import `DateEntry` in `app_gui.py`
    - _Requirements: 1.1_
  
  - [ ] 2.2 Add date picker widget to payment form
    - Position date picker before payment amount field
    - Configure DateEntry with `dd/mm/yyyy` format to show day, month, year
    - Set `maxdate` to current date to prevent future selection
    - Update form grid row numbers for payment amount field
    - _Requirements: 1.1, 1.2, 1.4_
  
  - [ ] 2.3 Modify `record_payment()` method in GUI to use selected date
    - Get selected date from `payment_date_entry.get_date()`
    - Combine date with current time to create datetime
    - Validate date is not in future (client-side check)
    - Pass `payment_datetime` to `payment_service.record_payment()`
    - Display error message if future date detected
    - _Requirements: 1.3, 1.5, 1.6, 3.1_
  
  - [ ]* 2.4 Write property test for GUI date selection
    - **Property 1: GUI Date Selection Populates Field**
    - **Validates: Requirements 1.3**
  
  - [ ]* 2.5 Write unit tests for GUI date picker integration
    - Test date picker widget initialization
    - Test date selection updates field
    - Test default date is current date
    - Test future date rejection in GUI
    - _Requirements: 1.1, 1.3, 1.4, 1.5_

- [ ] 3. Checkpoint - Test GUI date picker functionality
  - Ensure all tests pass, ask the user if questions arise.

- [ ] 4. Add date picker to web application
  - [ ] 4.1 Add HTML5 date input field to payment form
    - Position date input before payment amount field in HTML template
    - Set `type="date"` for native browser date picker
    - Set `max` attribute dynamically to current date
    - Add proper styling to match existing form fields
    - _Requirements: 2.1, 2.2_
  
  - [ ] 4.2 Update JavaScript `recordPayment()` function
    - Get date value from `payment-date` input field
    - Validate date is not in future (client-side check)
    - Include `payment_date` in API request body
    - Display error message if future date detected
    - _Requirements: 2.3, 2.4, 2.5_
  
  - [ ] 4.3 Update `/api/record-payment` endpoint to accept date parameter
    - Extract `payment_date` from request JSON
    - Parse date string in `YYYY-MM-DD` format
    - Combine with current time to create datetime
    - Validate date is not in future (server-side check)
    - Pass `payment_date` to `payment_service.record_payment()`
    - Return appropriate error for invalid or future dates
    - _Requirements: 2.3, 2.5, 3.1, 4.1, 4.2_
  
  - [ ]* 4.4 Write property test for web date selection
    - **Property 2: Web Date Selection Captured**
    - **Validates: Requirements 2.3**
  
  - [ ]* 4.5 Write unit tests for web date picker integration
    - Test date input field rendering
    - Test date value capture from form
    - Test API endpoint with custom date
    - Test API endpoint without date (defaults to current)
    - Test API endpoint with future date (rejected)
    - Test API endpoint with invalid date format
    - _Requirements: 2.1, 2.3, 2.4, 2.5, 3.1_

- [ ] 5. Checkpoint - Test web date picker functionality
  - Ensure all tests pass, ask the user if questions arise.

- [ ] 6. Verify timestamp persistence and display
  - [ ] 6.1 Verify payment history displays timestamps correctly
    - Confirm GUI payment history shows "YYYY-MM-DD HH:MM:SS" format
    - Confirm web payment history shows "YYYY-MM-DD HH:MM:SS" format
    - Verify existing code already handles this correctly
    - _Requirements: 5.1, 5.2_
  
  - [ ]* 6.2 Write property test for timestamp persistence round-trip
    - **Property 5: Timestamp Persistence Round-Trip**
    - **Validates: Requirements 3.4, 6.4**
  
  - [ ]* 6.3 Write property test for timestamp display format
    - **Property 6: Timestamp Display Format**
    - **Validates: Requirements 5.1, 5.2**
  
  - [ ]* 6.4 Write property test for existing records display
    - **Property 7: Existing Records Display Correctly**
    - **Validates: Requirements 6.2, 6.3**

- [ ] 7. Integration testing and backward compatibility verification
  - [ ] 7.1 Test end-to-end payment recording with custom dates
    - Test GUI: record payment with past date
    - Test Web: record payment with past date
    - Verify payment appears in history with correct timestamp
    - Verify payment persists to spreadsheet correctly
    - _Requirements: 3.1, 3.4, 5.1, 5.2_
  
  - [ ] 7.2 Test backward compatibility
    - Test recording payment without selecting date (should use current time)
    - Verify existing payment records load and display correctly
    - Verify dashboard continues to show payment counts
    - Verify no breaking changes to existing functionality
    - _Requirements: 6.1, 6.2, 6.3, 6.4_
  
  - [ ]* 7.3 Write integration tests
    - Test complete payment flow with custom date in GUI
    - Test complete payment flow with custom date in web
    - Test loading and displaying existing payments
    - _Requirements: 3.1, 3.4, 5.1, 5.2, 6.2, 6.3_

- [ ] 8. Final checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- The `tkcalendar` library provides the DateEntry widget for CustomTkinter
- HTML5 date input provides native date picker in modern browsers
- Date format in GUI is `dd/mm/yyyy` to show day, month, year clearly
- Date format in web API is `YYYY-MM-DD` (ISO 8601 standard)
- All date validation happens both client-side (UX) and server-side (security)
- Property tests should run minimum 100 iterations each
- Backward compatibility is critical - existing code must continue to work
