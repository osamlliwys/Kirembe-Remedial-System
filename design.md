# Design Document: Date Picker Integration

## Overview

This design adds date picker functionality to both the GUI and web applications of the Learner Payment Management System. The implementation will allow users to select custom dates when recording payments, enabling accurate tracking of historical payments while maintaining backward compatibility with existing functionality.

The design follows a layered architecture approach:
- **Presentation Layer**: Date picker UI components in both GUI (CustomTkinter) and web (HTML5) interfaces
- **Application Layer**: Enhanced Payment_Service to accept optional date parameters
- **Domain Layer**: No changes required - PaymentRecord already supports custom timestamps
- **Data Layer**: No changes required - SpreadsheetRepository already persists timestamps

## Architecture

### Component Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    Presentation Layer                        │
│  ┌──────────────────────┐    ┌──────────────────────────┐  │
│  │   GUI Application    │    │   Web Application        │  │
│  │   (CustomTkinter)    │    │   (Flask + HTML5)        │  │
│  │                      │    │                          │  │
│  │  - DateEntry widget  │    │  - <input type="date">   │  │
│  │  - Date validation   │    │  - JavaScript validation │  │
│  └──────────┬───────────┘    └──────────┬───────────────┘  │
└─────────────┼──────────────────────────┼───────────────────┘
              │                          │
              └──────────┬───────────────┘
                         │
┌────────────────────────▼─────────────────────────────────────┐
│                  Application Layer                           │
│  ┌──────────────────────────────────────────────────────┐   │
│  │           PaymentService (Enhanced)                   │   │
│  │                                                        │   │
│  │  record_payment(learner, amount, date=None)          │   │
│  │    - Accepts optional date parameter                  │   │
│  │    - Validates date is not in future                  │   │
│  │    - Defaults to current datetime if None             │   │
│  └────────────────────────┬─────────────────────────────┘   │
└─────────────────────────────┼───────────────────────────────┘
                              │
┌─────────────────────────────▼───────────────────────────────┐
│                     Domain Layer                             │
│  ┌──────────────────────────────────────────────────────┐   │
│  │         PaymentRecord (No changes needed)            │   │
│  │                                                        │   │
│  │  __init__(amount, timestamp, admission_number)       │   │
│  └────────────────────────┬─────────────────────────────┘   │
└─────────────────────────────┼───────────────────────────────┘
                              │
┌─────────────────────────────▼───────────────────────────────┐
│                      Data Layer                              │
│  ┌──────────────────────────────────────────────────────┐   │
│  │    SpreadsheetRepository (No changes needed)         │   │
│  │                                                        │   │
│  │  - Persists timestamps to Excel                       │   │
│  └──────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────┘
```

### Design Decisions

1. **Minimal Changes**: The domain and data layers already support custom timestamps, so changes are only needed in presentation and application layers.

2. **Optional Parameter**: The date parameter is optional in `record_payment()` to maintain backward compatibility. When not provided, it defaults to the current datetime.

3. **Client-Side Validation**: Both GUI and web interfaces perform initial date validation (no future dates) before sending to the service layer.

4. **Server-Side Validation**: The Payment_Service performs final validation to ensure data integrity regardless of client behavior.

5. **Native Components**: Use CustomTkinter's DateEntry for GUI and HTML5 date input for web to leverage platform-native date pickers.

## Components and Interfaces

### 1. GUI Application (app_gui.py)

#### New Component: Date Picker Widget

The date picker will be positioned before the payment amount field and will display day, month, and year components.

```python
# In setup_payment_tab() method
# Add date picker BEFORE payment amount field

# Payment date label
ctk.CTkLabel(
    form_frame,
    text="Payment Date:",
    font=ctk.CTkFont(size=14, weight="bold")
).grid(row=2, column=0, padx=20, pady=15, sticky="w")

# Date picker widget with day, month, year display
self.payment_date_entry = DateEntry(
    form_frame,
    width=300,
    date_pattern='dd/mm/yyyy',  # Shows day, month, year
    maxdate=datetime.now().date(),  # Prevent future dates
    showweeknumbers=False,
    showothermonthdays=False
)
self.payment_date_entry.grid(row=2, column=1, padx=20, pady=15, sticky="ew")

# Payment amount (now row 3 instead of row 2)
ctk.CTkLabel(
    form_frame,
    text="Payment Amount (KES):",
    font=ctk.CTkFont(size=14, weight="bold")
).grid(row=3, column=0, padx=20, pady=15, sticky="w")

self.payment_amount_entry = ctk.CTkEntry(
    form_frame,
    placeholder_text="Enter amount (e.g., 1000.00)",
    width=300
)
self.payment_amount_entry.grid(row=3, column=1, padx=20, pady=15, sticky="ew")
```

#### Modified Method: record_payment()

```python
def record_payment(self):
    """Record a payment for a learner with optional custom date."""
    # ... existing validation code ...
    
    # Get selected date from date picker
    selected_date = self.payment_date_entry.get_date()
    
    # Combine date with current time to create datetime
    payment_datetime = datetime.combine(
        selected_date,
        datetime.now().time()
    )
    
    # Validate date is not in future
    if payment_datetime > datetime.now():
        messagebox.showerror(
            "Invalid Date",
            "Payment date cannot be in the future."
        )
        return
    
    # Record payment with custom date
    result = self.payment_service.record_payment(
        learner,
        amount,
        payment_date=payment_datetime
    )
    # ... rest of existing code ...
```

### 2. Web Application (app_web.py)

#### Modified HTML Template

Add date input field in the payment form BEFORE the payment amount field. The HTML5 date input will display day, month, and year in the browser's native format:

```html
<!-- Add this BEFORE the payment amount field -->
<div class="form-group">
    <label for="payment-date">Payment Date:</label>
    <input 
        type="date" 
        id="payment-date" 
        max="" 
        style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
</div>

<!-- Payment amount field comes after -->
<div class="form-group">
    <label for="payment-amount">Payment Amount (KES):</label>
    <input type="number" id="payment-amount" placeholder="Enter amount (e.g., 1000.00)" step="0.01" min="0">
</div>
```

#### Modified JavaScript: recordPayment()

```javascript
async function recordPayment() {
    const admissionNumber = document.getElementById('payment-admission').value.trim();
    const amount = document.getElementById('payment-amount').value.trim();
    const paymentDate = document.getElementById('payment-date').value;
    
    // ... existing validation ...
    
    // Validate date if provided
    if (paymentDate) {
        const selectedDate = new Date(paymentDate);
        const today = new Date();
        today.setHours(23, 59, 59, 999); // End of today
        
        if (selectedDate > today) {
            resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Payment date cannot be in the future</div>';
            return;
        }
    }
    
    const response = await fetch('/api/record-payment', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            admission_number: admissionNumber,
            amount: parseFloat(amount),
            payment_date: paymentDate || null
        })
    });
    // ... rest of existing code ...
}
```

#### Modified API Endpoint: /api/record-payment

```python
@app.route('/api/record-payment', methods=['POST'])
def record_payment():
    """API endpoint to record a payment with optional date."""
    try:
        data = request.get_json()
        admission_number = data.get('admission_number', '').strip()
        amount_str = data.get('amount', '')
        payment_date_str = data.get('payment_date')
        
        # ... existing validation ...
        
        # Parse payment date if provided
        payment_date = None
        if payment_date_str:
            try:
                # Parse date string (YYYY-MM-DD format)
                date_obj = datetime.strptime(payment_date_str, '%Y-%m-%d').date()
                # Combine with current time
                payment_date = datetime.combine(date_obj, datetime.now().time())
                
                # Validate not in future
                if payment_date > datetime.now():
                    return jsonify({
                        'success': False,
                        'message': 'Payment date cannot be in the future'
                    }), 400
            except ValueError:
                return jsonify({
                    'success': False,
                    'message': 'Invalid date format'
                }), 400
        
        # Record payment with optional date
        result = payment_service.record_payment(learner, amount, payment_date)
        # ... rest of existing code ...
```

### 3. Payment Service (src/application/payment_service.py)

#### Modified Method: record_payment()

```python
def record_payment(
    self,
    learner: Learner,
    amount: Decimal,
    payment_date: Optional[datetime] = None
) -> dict:
    """
    Record a payment for a learner with optional custom date.
    
    Args:
        learner: The learner making the payment
        amount: Payment amount
        payment_date: Optional payment date (defaults to current datetime)
    
    Returns:
        Dictionary with success status and message
    """
    # Use provided date or default to current datetime
    timestamp = payment_date if payment_date is not None else datetime.now()
    
    # Validate date is not in future
    if timestamp > datetime.now():
        return {
            'success': False,
            'message': 'Payment date cannot be in the future'
        }
    
    # Create payment record with specified timestamp
    payment = PaymentRecord(
        amount=amount,
        timestamp=timestamp,
        admission_number=learner.admission_number
    )
    
    # ... rest of existing validation and persistence code ...
```

## Data Models

No changes required to existing data models. The `PaymentRecord` class already accepts a `timestamp` parameter in its constructor and stores it appropriately.

### Existing PaymentRecord Structure

```python
class PaymentRecord:
    def __init__(
        self,
        amount: Decimal,
        timestamp: datetime,
        admission_number: str
    ):
        self.amount = amount
        self.timestamp = timestamp
        self.admission_number = admission_number
```

This structure already supports custom timestamps, so no modifications are needed.


## Correctness Properties

A property is a characteristic or behavior that should hold true across all valid executions of a system—essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.

### Property 1: GUI Date Selection Populates Field

*For any* valid date selected in the GUI date picker widget, the date picker field should be populated with that selected date value.

**Validates: Requirements 1.3**

### Property 2: Web Date Selection Captured

*For any* valid date selected in the web application's date input field, the application should capture and store that date value for submission.

**Validates: Requirements 2.3**

### Property 3: Payment Created with Specified Timestamp

*For any* valid payment date provided to the Payment_Service, the created Payment_Record should have a timestamp matching the provided date.

**Validates: Requirements 3.1**

### Property 4: Future Date Validation

*For any* date provided to the Payment_Service, if the date is in the future (after the current datetime), the service should reject the payment and return an error.

**Validates: Requirements 3.3, 4.1**

### Property 5: Timestamp Persistence Round-Trip

*For any* Payment_Record with a custom timestamp, persisting it to the spreadsheet and then reading it back should produce a Payment_Record with an equivalent timestamp.

**Validates: Requirements 3.4, 6.4**

### Property 6: Timestamp Display Format

*For any* Payment_Record displayed in either the GUI or web application's payment history, the timestamp should be formatted as "YYYY-MM-DD HH:MM:SS".

**Validates: Requirements 5.1, 5.2**

### Property 7: Existing Records Display Correctly

*For any* existing Payment_Record loaded from the spreadsheet, both the GUI and web applications should display the payment with its original timestamp preserved.

**Validates: Requirements 6.2, 6.3**

## Error Handling

### Date Validation Errors

**Future Date Error:**
- **Trigger**: User selects or submits a date in the future
- **Response**: Display error message "Payment date cannot be in the future"
- **Recovery**: Allow user to select a different date

**Invalid Date Format Error:**
- **Trigger**: Web API receives malformed date string
- **Response**: Return HTTP 400 with error message "Invalid date format"
- **Recovery**: Client should validate date format before submission

### UI Error Handling

**GUI Application:**
- Use `messagebox.showerror()` to display date validation errors
- Prevent form submission when date is invalid
- Keep form data intact so user can correct the date

**Web Application:**
- Display error messages in the `alert alert-error` div
- Use client-side JavaScript validation to catch errors before API call
- Server-side validation as final safeguard

### Backward Compatibility

**Missing Date Parameter:**
- **Behavior**: When `payment_date` parameter is `None` or not provided
- **Response**: Use `datetime.now()` as the timestamp
- **Rationale**: Maintains existing behavior for code that doesn't use the new feature

## Testing Strategy

### Dual Testing Approach

This feature requires both unit tests and property-based tests to ensure comprehensive coverage:

**Unit Tests** focus on:
- Specific examples of date selection and validation
- Edge cases (current date, dates at boundaries)
- Error conditions (future dates, invalid formats)
- Integration between UI components and services

**Property-Based Tests** focus on:
- Universal properties that hold for all valid dates
- Round-trip persistence of timestamps
- Format validation across many generated dates
- Validation logic across a range of inputs

### Property-Based Testing Configuration

**Library Selection:**
- **Python**: Use `hypothesis` library for property-based testing
- **Minimum iterations**: 100 test cases per property
- **Test tagging**: Each property test must reference its design property

**Tag Format:**
```python
# Feature: date-picker-integration, Property 3: Payment Created with Specified Timestamp
```

### Test Coverage

#### Unit Tests

1. **GUI Date Picker Tests**
   - Test date picker widget initialization
   - Test date selection updates field value
   - Test default date is current date
   - Test future date rejection
   - Test calendar display on click

2. **Web Date Input Tests**
   - Test date input field rendering
   - Test date value capture from form
   - Test default date behavior
   - Test client-side future date validation
   - Test date format validation

3. **Payment Service Tests**
   - Test payment recording with custom date
   - Test payment recording without date (default behavior)
   - Test future date rejection
   - Test date validation logic

4. **Integration Tests**
   - Test end-to-end payment recording with custom date in GUI
   - Test end-to-end payment recording with custom date in web
   - Test payment history display with custom dates

#### Property-Based Tests

1. **Property 1: GUI Date Selection** (Feature: date-picker-integration, Property 1)
   - Generate random valid dates
   - Verify field population for each date

2. **Property 2: Web Date Selection** (Feature: date-picker-integration, Property 2)
   - Generate random valid dates
   - Verify date capture for each date

3. **Property 3: Payment Timestamp** (Feature: date-picker-integration, Property 3)
   - Generate random valid dates
   - Create payments with those dates
   - Verify Payment_Record timestamps match

4. **Property 4: Future Date Validation** (Feature: date-picker-integration, Property 4)
   - Generate random future dates
   - Verify all are rejected with error

5. **Property 5: Timestamp Round-Trip** (Feature: date-picker-integration, Property 5)
   - Generate random Payment_Records with various timestamps
   - Persist and reload
   - Verify timestamps are preserved

6. **Property 6: Timestamp Format** (Feature: date-picker-integration, Property 6)
   - Generate random Payment_Records
   - Format timestamps for display
   - Verify all match "YYYY-MM-DD HH:MM:SS" pattern

7. **Property 7: Existing Records Display** (Feature: date-picker-integration, Property 7)
   - Generate random existing Payment_Records
   - Load and display in both interfaces
   - Verify timestamps are preserved and displayed correctly

### Edge Cases to Test

- **Current date/time**: Selecting today's date should work
- **Past dates**: Historical dates should be accepted
- **Boundary dates**: Dates at the edge of valid ranges
- **No date selected**: Should default to current datetime
- **Midnight timestamps**: Dates with time component at 00:00:00
- **Existing payments**: Loading payments created before this feature

### Testing Dependencies

**Required Libraries:**
- `hypothesis`: Property-based testing framework
- `pytest`: Test runner
- `unittest.mock`: For mocking datetime.now() in tests
- `freezegun`: For time-based testing (optional but recommended)

**Test Data:**
- Sample spreadsheet with existing payment records
- Test learner data with various payment histories
- Date ranges covering past, present, and future
