# Web Application Guide

## Kirembe Secondary School - Payment Management System

This guide will help you use the web-based interface for the payment management system.

## What is the Web App?

The web application provides a browser-based interface that:
- Works completely **offline** (no internet required)
- Can be accessed from **any device** on your local network
- Has the **same features** as the desktop GUI
- Uses a modern, responsive design
- Works on computers, tablets, and phones

## Starting the Web Application

### Option 1: Windows Batch File (Easiest)

1. Double-click `start_web.bat` in the project folder
2. Wait for the server to start
3. Your browser will show the URL: `http://localhost:5000`
4. Open that URL in any browser

### Option 2: Command Line

```bash
python app_web.py
```

Or on Windows:

```bash
py app_web.py
```

### What You'll See

When the server starts, you'll see:

```
============================================================
Kirembe Secondary School - Payment Management System
Web Application
============================================================

✓ Successfully loaded 165 learners

============================================================
Starting web server...
============================================================

🌐 Open your browser and go to:

   http://localhost:5000

   or

   http://127.0.0.1:5000

============================================================
Press Ctrl+C to stop the server
============================================================
```

## Accessing from Other Devices

### On the Same Network

You can access the web app from other devices (phones, tablets, other computers) on the same network:

1. Find your computer's IP address:
   - Windows: Open Command Prompt and type `ipconfig`
   - Look for "IPv4 Address" (e.g., 192.168.1.100)

2. On the other device, open a browser and go to:
   ```
   http://YOUR-IP-ADDRESS:5000
   ```
   Example: `http://192.168.1.100:5000`

## Using the Web Interface

The web app has three main tabs:

### 1. Dashboard Tab 📊

**What it shows:**
- All 165 learners in a scrollable table
- Color-coded balances:
  - 🔴 Red: Outstanding balance
  - 🟢 Green: Fully paid or credit
- Total owed, total paid, and payment count for each learner

**How to use:**
1. Click the "Dashboard" tab (it opens by default)
2. Scroll through the table to view all learners
3. Click "🔄 Refresh" to reload the latest data

**Features:**
- Sticky header (stays visible when scrolling)
- Alternating row colors for easy reading
- Hover effect on rows
- Real-time data from Excel file

### 2. Search Learner Tab 🔍

**What it does:**
- Search for a specific learner by admission number
- View complete learner details
- See full payment history

**How to use:**
1. Click the "Search Learner" tab
2. Enter the admission number in the search box
3. Click "🔍 Search" or press Enter
4. View the learner's complete information

**What you'll see:**
- Name and admission number
- Total owed, total paid, and balance
- Complete payment history (most recent first)
- Color-coded balance with credit indicator

**Example:**
```
Enter: 1710
Result: Shows all details for student with admission number 1710
```

### 3. Record Payment Tab 💰

**What it does:**
- Record new payments for learners
- Validate learner exists before payment
- Update the system immediately

**How to use:**
1. Click the "Record Payment" tab
2. Enter the admission number
3. Click "Lookup" to verify the learner
4. See the learner's name and current balance
5. Enter the payment amount
6. Click "💰 Record Payment"
7. See confirmation with new balance

**Payment amount formats:**
- `5000`
- `5000.50`
- `1500.75`

**Validation:**
- Amount must be greater than zero
- Maximum: 1,000,000 KES
- Learner must exist in the system

## Features in Detail

### Color-Coded Balances

The system uses colors to help you quickly identify payment status:

- **Red text**: Student owes money
  - Example: "KES 5,000.00"
  
- **Green text with "Credit"**: Student has overpaid
  - Example: "KES 500.00 (Credit)"
  
- **Green text**: Student is fully paid
  - Example: "KES 0.00"

### Real-Time Updates

When you record a payment:
1. The payment is immediately saved to the Excel file
2. The dashboard automatically refreshes
3. You see the updated balance instantly
4. All other users see the update when they refresh

### Responsive Design

The web app works on:
- Desktop computers (Windows, Mac, Linux)
- Tablets (iPad, Android tablets)
- Smartphones (iPhone, Android phones)
- Any device with a modern web browser

### Offline Operation

The web app:
- Runs completely on your local computer
- Doesn't require internet connection
- All data stays on your computer
- No external services or cloud storage

## Common Tasks

### Task 1: Check All Learner Balances

1. Open the web app
2. Go to Dashboard tab (default)
3. Scroll through the table
4. Look for red balances (students who owe money)

### Task 2: Find a Specific Student

1. Go to Search Learner tab
2. Enter admission number
3. Click Search
4. View complete details and payment history

### Task 3: Record a Payment

1. Go to Record Payment tab
2. Enter admission number
3. Click Lookup to verify
4. Enter payment amount
5. Click Record Payment
6. See confirmation message

### Task 4: Access from Phone/Tablet

1. Make sure your computer is running the web app
2. Find your computer's IP address
3. On your phone/tablet, open browser
4. Go to `http://YOUR-IP:5000`
5. Use the app normally

## Troubleshooting

### "This site can't be reached"

**Problem**: Browser can't connect to the server

**Solutions**:
1. Check that `app_web.py` is running
2. Look for the "Starting web server..." message
3. Try `http://127.0.0.1:5000` instead
4. Check if another program is using port 5000

### "Cannot find payment database file"

**Problem**: Excel file is missing

**Solutions**:
1. Check that `REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx` exists
2. Place the file in the same folder as `app_web.py`
3. Restart the web server

### "Error loading data"

**Problem**: Can't read from Excel file

**Solutions**:
1. Close Excel if the file is open
2. Check file permissions
3. Restart the web server
4. Check the terminal for error messages

### Can't access from other devices

**Problem**: Other devices can't connect

**Solutions**:
1. Check that both devices are on the same network
2. Verify your computer's IP address
3. Check Windows Firewall settings
4. Make sure the web server is running

### Page is slow or unresponsive

**Problem**: Web app is slow

**Solutions**:
1. Refresh the page (F5)
2. Clear browser cache
3. Close other browser tabs
4. Restart the web server

## Stopping the Web Server

To stop the web server:

1. Go to the terminal/command prompt where it's running
2. Press `Ctrl+C`
3. Wait for the server to shut down
4. You'll see "Shutting down..." message

## Security Notes

### Local Network Only

The web server is configured to:
- Accept connections from your local network only
- Not be accessible from the internet
- Run on your computer without external access

### Data Privacy

- All data stays on your computer
- No data is sent to external servers
- Excel file is read/written locally
- No cloud storage or external services

### Recommended Practices

1. **Don't expose to internet**: Keep the server on your local network only
2. **Backup your data**: Regularly backup the Excel file
3. **Close when not in use**: Stop the server when you're done
4. **Monitor access**: Only share the URL with authorized users

## Advantages of Web App

### Compared to CLI

✅ Visual interface (no typing commands)  
✅ Works on any device with a browser  
✅ Multiple users can access simultaneously  
✅ Modern, colorful design  
✅ Touch-friendly for tablets  

### Compared to Desktop GUI

✅ No installation needed on client devices  
✅ Access from phones and tablets  
✅ Multiple users can use at once  
✅ Works on Mac, Linux, Windows  
✅ Easier to share with others  

### Unique Benefits

✅ Access from anywhere on your network  
✅ No software installation on other devices  
✅ Works on any operating system  
✅ Easy to demonstrate to others  
✅ Professional web interface  

## Technical Details

### Technology Stack

- **Backend**: Flask (Python web framework)
- **Frontend**: HTML5, CSS3, JavaScript
- **Data**: Excel file (openpyxl)
- **Server**: Built-in Flask development server

### Port Configuration

- Default port: 5000
- Can be changed in `app_web.py` if needed
- Accessible on all network interfaces (0.0.0.0)

### Browser Compatibility

Works with:
- Google Chrome (recommended)
- Mozilla Firefox
- Microsoft Edge
- Safari
- Opera
- Any modern browser

### Performance

- Handles 165 learners easily
- Fast page loads (< 1 second)
- Instant search results
- Real-time updates
- Minimal memory usage

## Getting Help

If you encounter issues:

1. Check this guide first
2. Look at the terminal for error messages
3. Try restarting the web server
4. Check the main README.md
5. Verify Excel file is accessible

## Version

Web App Version 1.0.0 - March 2026

---

**Enjoy using the Kirembe Secondary School Payment Management Web App!** 🎓
