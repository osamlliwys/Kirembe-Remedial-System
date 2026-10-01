#!/usr/bin/env python3
"""
Kirembe Secondary School - Payment Management System
Copyright (c) 2026 Kirembe Secondary School. All Rights Reserved.

Web application entry point for the Learner Payment Management System.

This software is proprietary and confidential. Unauthorized copying,
modification, distribution, or use is strictly prohibited.

This module provides a Flask-based web interface that works completely offline.
All HTML, CSS, and JavaScript are embedded for offline use.

Version: 3.0
Date: March 2026
"""

import os
import sys
from flask import Flask, render_template_string, jsonify, request
from decimal import Decimal, InvalidOperation
from datetime import datetime
from src.data.spreadsheet_repository import SpreadsheetRepository
from src.application.learner_service import LearnerService
from src.application.payment_service import PaymentService
from src.application.payment_correction_service import PaymentCorrectionService
from src.teacher_payroll.teacher_manager import TeacherManager
from src.teacher_payroll.attendance_tracker import AttendanceTracker
from src.teacher_payroll.payroll_calculator import PayrollCalculator
from src.teacher_payroll.financial_dashboard import FinancialDashboard

# Configuration
SPREADSHEET_FILE = "REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx"
TEACHER_PAYROLL_FILE = "TEACHER_PAYROLL.xlsx"

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = 'kirembe-secondary-school-2026'

# Global services
learner_service = None
payment_service = None
payment_correction_service = None
teacher_manager = None
attendance_tracker = None
payroll_calculator = None
financial_dashboard = None

# HTML Template with embedded CSS and JavaScript
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Kirembe Secondary School - Payment Management</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #2d5016 0%, #4a7c2c 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2d5016 0%, #3d6b1f 100%);
            color: white;
            padding: 20px 30px;
            text-align: center;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 20px;
        }
        
        .header-logo {
            width: 80px;
            height: 80px;
            object-fit: contain;
            object-position: center;
            border-radius: 8px;
            background: white;
            padding: 5px;
        }
        
        .header-text {
            flex: 1;
        }
        
        .header h1 {
            font-size: 32px;
            margin-bottom: 5px;
            text-transform: uppercase;
            letter-spacing: 2px;
            color: #FFD700;
        }
        
        .header p {
            font-size: 16px;
            opacity: 0.95;
            color: #f0f0f0;
        }
        
        .tabs {
            display: flex;
            background: #f5f5f5;
            border-bottom: 2px solid #ddd;
        }
        
        .tab {
            flex: 1;
            padding: 20px;
            text-align: center;
            cursor: pointer;
            background: #f5f5f5;
            border: none;
            font-size: 16px;
            font-weight: 600;
            color: #666;
            transition: all 0.3s;
            position: relative;
        }
        
        .tab:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        
        /* Dashboard Tab - Blue */
        .tab:nth-child(1) {
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            color: #1565c0;
        }
        
        .tab:nth-child(1):hover {
            background: linear-gradient(135deg, #bbdefb 0%, #90caf9 100%);
        }
        
        .tab:nth-child(1).active {
            background: white;
            color: #1565c0;
            border-bottom: 4px solid #1565c0;
            box-shadow: 0 -2px 10px rgba(21, 101, 192, 0.3);
        }
        
        /* Search Tab - Purple */
        .tab:nth-child(2) {
            background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%);
            color: #6a1b9a;
        }
        
        .tab:nth-child(2):hover {
            background: linear-gradient(135deg, #e1bee7 0%, #ce93d8 100%);
        }
        
        .tab:nth-child(2).active {
            background: white;
            color: #6a1b9a;
            border-bottom: 4px solid #6a1b9a;
            box-shadow: 0 -2px 10px rgba(106, 27, 154, 0.3);
        }
        
        /* Payment Tab - Green */
        .tab:nth-child(3) {
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
            color: #2e7d32;
        }
        
        .tab:nth-child(3):hover {
            background: linear-gradient(135deg, #c8e6c9 0%, #a5d6a7 100%);
        }
        
        .tab:nth-child(3).active {
            background: white;
            color: #2e7d32;
            border-bottom: 4px solid #2e7d32;
            box-shadow: 0 -2px 10px rgba(46, 125, 50, 0.3);
        }
        
        /* Add Learner Tab - Orange */
        .tab:nth-child(4) {
            background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
            color: #e65100;
        }
        
        .tab:nth-child(4):hover {
            background: linear-gradient(135deg, #ffe0b2 0%, #ffcc80 100%);
        }
        
        .tab:nth-child(4).active {
            background: white;
            color: #e65100;
            border-bottom: 4px solid #e65100;
            box-shadow: 0 -2px 10px rgba(230, 81, 0, 0.3);
        }
        
        /* Exit Learner Tab - Red */
        .tab:nth-child(5) {
            background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
            color: #c62828;
        }
        
        .tab:nth-child(5):hover {
            background: linear-gradient(135deg, #ffcdd2 0%, #ef9a9a 100%);
        }
        
        .tab:nth-child(5).active {
            background: white;
            color: #c62828;
            border-bottom: 4px solid #c62828;
            box-shadow: 0 -2px 10px rgba(198, 40, 40, 0.3);
        }
        
        /* Promote Students Tab - Teal */
        .tab:nth-child(6) {
            background: linear-gradient(135deg, #e0f2f1 0%, #b2dfdb 100%);
            color: #00695c;
        }
        
        .tab:nth-child(6):hover {
            background: linear-gradient(135deg, #b2dfdb 0%, #80cbc4 100%);
        }
        
        .tab:nth-child(6).active {
            background: white;
            color: #00695c;
            border-bottom: 4px solid #00695c;
            box-shadow: 0 -2px 10px rgba(0, 105, 92, 0.3);
        }
        
        /* Teacher Payroll Tab - Gold/Yellow */
        .tab:nth-child(7) {
            background: linear-gradient(135deg, #fff9c4 0%, #fff59d 100%);
            color: #f57f17;
        }
        
        .tab:nth-child(7):hover {
            background: linear-gradient(135deg, #fff59d 0%, #fff176 100%);
        }
        
        .tab:nth-child(7).active {
            background: white;
            color: #f57f17;
            border-bottom: 4px solid #f57f17;
            box-shadow: 0 -2px 10px rgba(245, 127, 23, 0.3);
        }
        
        .tab-content {
            display: none;
            padding: 30px;
            animation: fadeIn 0.3s;
        }
        
        .tab-content.active {
            display: block;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .stats-bar {
            background: #f8f9fa;
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .stats-text {
            font-size: 16px;
            color: #333;
            font-weight: 600;
        }
        
        .btn {
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .btn-primary {
            background: #2d5016;
            color: white;
        }
        
        .btn-primary:hover {
            background: #1f3a0f;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(45, 80, 22, 0.3);
        }
        
        .btn-success {
            background: #28a745;
            color: white;
        }
        
        .btn-success:hover {
            background: #218838;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(40, 167, 69, 0.3);
        }
        
        .table-container {
            overflow-x: auto;
            max-height: 500px;
            border: 1px solid #ddd;
            border-radius: 8px;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
        }
        
        thead {
            position: sticky;
            top: 0;
            background: #2b2b2b;
            color: white;
            z-index: 10;
        }
        
        th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
            font-size: 14px;
        }
        
        tbody tr {
            border-bottom: 1px solid #eee;
            transition: background 0.2s;
        }
        
        tbody tr:nth-child(even) {
            background: #f9f9f9;
        }
        
        tbody tr:hover {
            background: #f0f0f0;
        }
        
        td {
            padding: 12px 15px;
            font-size: 13px;
        }
        
        .balance-positive {
            color: #dc3545;
            font-weight: 600;
        }
        
        .balance-zero {
            color: #28a745;
            font-weight: 600;
        }
        
        .balance-credit {
            color: #28a745;
            font-weight: 600;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #333;
            font-size: 14px;
        }
        
        input[type="text"],
        input[type="number"] {
            width: 100%;
            padding: 12px;
            border: 2px solid #ddd;
            border-radius: 6px;
            font-size: 14px;
            transition: border 0.3s;
        }
        
        input:focus {
            outline: none;
            border-color: #2d5016;
        }
        
        .search-box {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
        }
        
        .search-box input {
            flex: 1;
        }
        
        .result-card {
            background: #f8f9fa;
            border: 2px solid #ddd;
            border-radius: 8px;
            padding: 20px;
            margin-top: 20px;
        }
        
        .result-card h3 {
            color: #2d5016;
            margin-bottom: 15px;
            font-size: 20px;
        }
        
        .detail-row {
            display: flex;
            justify-content: space-between;
            padding: 10px 0;
            border-bottom: 1px solid #ddd;
        }
        
        .detail-row:last-child {
            border-bottom: none;
        }
        
        .detail-label {
            font-weight: 600;
            color: #666;
        }
        
        .detail-value {
            color: #333;
            font-weight: 500;
        }
        
        .payment-history {
            margin-top: 20px;
            max-height: 300px;
            overflow-y: auto;
        }
        
        .payment-item {
            display: flex;
            justify-content: space-between;
            padding: 10px;
            background: white;
            border-radius: 4px;
            margin-bottom: 8px;
            border-left: 4px solid #28a745;
        }
        
        .alert {
            padding: 15px 20px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-weight: 500;
        }
        
        .alert-success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .alert-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .alert-info {
            background: #d1ecf1;
            color: #0c5460;
            border: 1px solid #bee5eb;
        }
        
        .learner-info {
            background: #e7f3ff;
            padding: 12px;
            border-radius: 6px;
            margin-top: 10px;
            font-size: 14px;
            color: #004085;
        }
        
        .loading {
            text-align: center;
            padding: 40px;
            color: #666;
        }
        
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #2d5016;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        /* Modal styles */
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.5);
        }
        
        .modal-content {
            background-color: white;
            margin: 10% auto;
            padding: 30px;
            border-radius: 10px;
            width: 90%;
            max-width: 500px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        }
        
        .modal-header {
            font-size: 20px;
            font-weight: 600;
            margin-bottom: 20px;
            color: #2d5016;
        }
        
        .modal-buttons {
            display: flex;
            gap: 10px;
            justify-content: flex-end;
            margin-top: 20px;
        }
        
        .btn-cancel {
            background: #6c757d;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
        }
        
        .btn-cancel:hover {
            background: #5a6268;
        }
        
        /* Print Receipt Styles */
        .receipt {
            display: none;
            background: white;
            padding: 40px;
            max-width: 800px;
            margin: 0 auto;
        }
        
        .receipt-header {
            text-align: center;
            border-bottom: 2px solid #2d5016;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        
        .receipt-header h1 {
            color: #2d5016;
            margin: 10px 0;
            font-size: 28px;
        }
        
        .receipt-header p {
            color: #666;
            margin: 5px 0;
        }
        
        .receipt-title {
            text-align: center;
            font-size: 24px;
            font-weight: 600;
            color: #2d5016;
            margin: 20px 0;
            text-transform: uppercase;
        }
        
        .receipt-info {
            margin: 20px 0;
        }
        
        .receipt-row {
            display: flex;
            justify-content: space-between;
            padding: 10px 0;
            border-bottom: 1px solid #eee;
        }
        
        .receipt-label {
            font-weight: 600;
            color: #333;
        }
        
        .receipt-value {
            color: #666;
        }
        
        .receipt-amount {
            background: #f0f0f0;
            padding: 20px;
            margin: 20px 0;
            border-radius: 8px;
            text-align: center;
        }
        
        .receipt-amount-label {
            font-size: 14px;
            color: #666;
            margin-bottom: 10px;
        }
        
        .receipt-amount-value {
            font-size: 32px;
            font-weight: 700;
            color: #2d5016;
        }
        
        .receipt-footer {
            margin-top: 40px;
            padding-top: 20px;
            border-top: 2px solid #2d5016;
            text-align: center;
            color: #666;
            font-size: 12px;
        }
        
        @media print {
            /* Hide everything by default */
            body > * { display: none !important; }
            
            /* PAYROLL PRINT: show only payroll-print-doc */
            body.printing-payroll > #payroll-print-doc {
                display: block !important;
            }
            body.printing-payroll > #payroll-print-doc * {
                display: revert !important;
                visibility: visible !important;
            }
            
            /* RECEIPT PRINT: show only the receipt div */
            body.printing-receipt > .receipt {
                display: block !important;
                position: fixed;
                left: 0; top: 0;
                width: 100%;
            }
            body.printing-receipt > .receipt * {
                display: revert !important;
                visibility: visible !important;
            }
            
            @page { margin: 12mm; }
            body.printing-payroll { @page { size: landscape; margin: 15mm; } }
            
            .btn, .no-print { display: none !important; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <img src="data:image/png;base64,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" alt="Kirembe Secondary School Logo" class="header-logo">
            <div class="header-text">
                <h1>Kirembe Secondary School</h1>
                <p>Learner Payment Management System</p>
            </div>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="showTab('dashboard')">📊 Dashboard</button>
            <button class="tab" onclick="showTab('search')">🔍 Search Learner</button>
            <button class="tab" onclick="showTab('payment')">💰 Record Payment</button>
            <button class="tab" onclick="showTab('addlearner')">➕ Add New Learner</button>
            <button class="tab" onclick="showTab('exitlearner')">🚪 Exit Learner</button>
            <button class="tab" onclick="showTab('promote')">🎓 Promote Students</button>
            <button class="tab" onclick="showTab('payroll')">💵 Teacher Payroll</button>
        </div>
        
        <!-- Dashboard Tab -->
        <div id="dashboard" class="tab-content active">
            <div class="stats-bar">
                <span class="stats-text" id="stats-text">Loading...</span>
                <button class="btn btn-primary" onclick="loadDashboard()">🔄 Refresh</button>
            </div>
            
            <!-- Payment Collections Chart -->
            <div style="background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <h3 style="color: #2d5016; margin-bottom: 20px; font-size: 18px;">📊 Payment Collections by Class</h3>
                <div id="class-chart" style="min-height: 300px;">
                    <div class="loading">
                        <div class="spinner"></div>
                        Loading statistics...
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Search Tab -->
        <div id="search" class="tab-content">
            <div class="search-box">
                <input type="text" id="search-admission" placeholder="Enter admission number" onkeypress="if(event.key==='Enter') searchLearner()">
                <button class="btn btn-primary" onclick="searchLearner()">🔍 Search</button>
            </div>
            <div id="search-results"></div>
        </div>
        
        <!-- Payment Tab -->
        <div id="payment" class="tab-content">
            <div class="form-group">
                <label for="payment-admission">Admission Number:</label>
                <div style="display: flex; gap: 10px;">
                    <input type="text" id="payment-admission" placeholder="Enter admission number">
                    <button class="btn btn-primary" onclick="lookupLearner()">Lookup</button>
                </div>
                <div id="learner-info"></div>
            </div>
            
            <div class="form-group">
                <label for="payment-amount">Payment Amount (KES):</label>
                <input type="number" id="payment-amount" placeholder="Enter amount (e.g., 1000.00)" step="0.01" min="0">
            </div>
            
            <div class="form-group">
                <label for="payment-date">Payment Date (Optional):</label>
                <input type="date" id="payment-date" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">Leave blank to use current date and time</small>
            </div>
            
            <button class="btn btn-success" onclick="recordPayment()">💰 Record Payment</button>
            
            <div id="payment-result"></div>
        </div>
        
        <!-- Add New Learner Tab -->
        <div id="addlearner" class="tab-content">
            <h2 style="color: #2d5016; margin-bottom: 20px;">➕ Add New Learner</h2>
            
            <div class="form-group">
                <label for="new-admission">Admission Number: *</label>
                <input type="text" id="new-admission" placeholder="Enter admission number (e.g., 2001)" required>
            </div>
            
            <div class="form-group">
                <label for="new-name">Full Name: *</label>
                <input type="text" id="new-name" placeholder="Enter learner's full name" required>
            </div>
            
            <div class="form-group">
                <label for="new-grade">Grade/Class: *</label>
                <select id="new-grade" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                    <option value="">-- Select Grade --</option>
                    <option value="GRADE 10 2026">Grade 10 (Form 2) 2026</option>
                    <option value="FORM 3 2025">Form 3 2025</option>
                    <option value="FORM 3 2026">Form 3 2026</option>
                    <option value="FORM 4 2026">Form 4 2026</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="new-arrears">Previous Arrears (if any):</label>
                <input type="number" id="new-arrears" placeholder="Enter arrears amount (default: 0)" step="0.01" min="0" value="0">
                <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">Enter any outstanding balance from previous school</small>
            </div>
            
            <div class="form-group">
                <label for="new-term">Admission Term: *</label>
                <select id="new-term" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                    <option value="">-- Select Term --</option>
                    <option value="term1">Term 1 (January - March) - Full Fee: 2,200</option>
                    <option value="term2">Term 2 (April - July) - Prorated: 1,400</option>
                    <option value="term3">Term 3 (August - December) - Prorated: 600</option>
                </select>
                <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">Fee is prorated based on admission term</small>
            </div>
            
            <div class="form-group">
                <label for="new-debt">Current Year Debt:</label>
                <input type="number" id="new-debt" placeholder="Fee amount (auto-calculated based on term)" step="0.01" min="0" readonly>
                <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">
                    <strong>Fee Structure:</strong><br>
                    • Term 1 (Jan-Mar): 2,200 (Full year)<br>
                    • Term 2 (Apr-Jul): 1,400 (Remaining terms)<br>
                    • Term 3 (Aug-Dec): 600 (Final term)
                </small>
            </div>
            
            <div class="form-group">
                <label for="new-admission-date">Admission Date:</label>
                <input type="date" id="new-admission-date" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">Leave blank to use today's date</small>
            </div>
            
            <button class="btn btn-success" onclick="addNewLearner()">➕ Add Learner</button>
            
            <div id="add-learner-result"></div>
        </div>
        
        <!-- Exit Learner Tab -->
        <div id="exitlearner" class="tab-content">
            <h2 style="color: #2d5016; margin-bottom: 20px;">🚪 Exit Learner</h2>
            
            <div class="alert alert-info">
                <strong>Note:</strong> Exiting a learner will mark them as inactive but preserve all their payment history and records for future reference.
            </div>
            
            <div class="form-group">
                <label for="exit-admission">Admission Number: *</label>
                <div style="display: flex; gap: 10px;">
                    <input type="text" id="exit-admission" placeholder="Enter admission number">
                    <button class="btn btn-primary" onclick="lookupLearnerForExit()">Lookup</button>
                </div>
                <div id="exit-learner-info"></div>
            </div>
            
            <div class="form-group">
                <label for="exit-date">Exit Date:</label>
                <input type="date" id="exit-date" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">Leave blank to use today's date</small>
            </div>
            
            <div class="form-group">
                <label for="exit-reason">Reason for Exit:</label>
                <select id="exit-reason" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                    <option value="Graduated">Graduated</option>
                    <option value="Transferred">Transferred to Another School</option>
                    <option value="Withdrawn">Withdrawn by Parent/Guardian</option>
                    <option value="Expelled">Expelled</option>
                    <option value="Other">Other</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="exit-notes">Additional Notes (Optional):</label>
                <textarea id="exit-notes" placeholder="Enter any additional notes about the exit" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px; min-height: 100px; resize: vertical;"></textarea>
            </div>
            
            <button class="btn btn-success" onclick="exitLearner()" style="background: #dc3545;">🚪 Exit Learner</button>
            
            <div id="exit-learner-result"></div>
        </div>
        
        <!-- Promote Students Tab -->
        <div id="promote" class="tab-content">
            <h2 style="color: #2d5016; margin-bottom: 20px;">🎓 Promote Students to Next Grade</h2>
            
            <div class="alert alert-info">
                <strong>Year-End Processing:</strong> This feature promotes all students to the next grade level and carries forward unpaid balances as arrears.
            </div>
            
            <div class="form-group">
                <label for="promote-year">New Academic Year: *</label>
                <input type="number" id="promote-year" placeholder="e.g., 2027" min="2026" max="2050" value="2027" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">Enter the year for the new academic session</small>
            </div>
            
            <div class="form-group">
                <label>Select Classes to Promote:</label>
                <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; border: 2px solid #ddd;">
                    <div style="margin-bottom: 10px;">
                        <input type="checkbox" id="promote-grade10" checked>
                        <label for="promote-grade10" style="display: inline; margin-left: 8px; font-weight: normal;">
                            Grade 10 (Form 2) 2026 → Form 3 2027
                        </label>
                    </div>
                    <div style="margin-bottom: 10px;">
                        <input type="checkbox" id="promote-form3" checked>
                        <label for="promote-form3" style="display: inline; margin-left: 8px; font-weight: normal;">
                            Form 3 2026 → Form 4 2027
                        </label>
                    </div>
                    <div>
                        <input type="checkbox" id="promote-form4">
                        <label for="promote-form4" style="display: inline; margin-left: 8px; font-weight: normal;">
                            Form 4 2026 → Graduate/Exit (Mark as completed)
                        </label>
                    </div>
                </div>
            </div>
            
            <div class="alert" style="background: #fff3cd; border-color: #ffc107; color: #856404; margin-bottom: 20px;">
                <strong>⚠️ Important:</strong> This action will:
                <ul style="margin: 10px 0 0 20px;">
                    <li>Create new sheets for the next academic year</li>
                    <li>Move all students to their next grade level</li>
                    <li>Calculate unpaid balances and carry them forward as arrears</li>
                    <li>Add new annual fee (2200) for each student</li>
                    <li>Preserve all payment history</li>
                    <li>Create a backup before making changes</li>
                </ul>
            </div>
            
            <div style="display: flex; gap: 10px;">
                <button class="btn btn-primary" onclick="previewPromotion()">👁️ Preview Changes</button>
                <button class="btn btn-success" onclick="promoteStudents()" style="background: #28a745;">🎓 Promote Students</button>
            </div>
            
            <div id="promote-preview" style="margin-top: 20px;"></div>
            <div id="promote-result"></div>
        </div>
        
        <!-- Teacher Payroll Tab -->
        <div id="payroll" class="tab-content">
            <h2 style="color: #f57f17; margin-bottom: 20px;">💵 Teacher Payroll</h2>
            
            <!-- Financial Summary -->
            <div style="background: linear-gradient(135deg, #fff9c4 0%, #fff59d 100%); padding: 20px; border-radius: 8px; margin-bottom: 20px;">
                <h3 style="color: #f57f17; margin-bottom: 15px;">📊 Financial Summary</h3>
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px;">
                    <div style="background: white; padding: 15px; border-radius: 6px; text-align: center;">
                        <div style="font-size: 12px; color: #666; margin-bottom: 5px;">NET REVENUE</div>
                        <div id="student-revenue" style="font-size: 24px; font-weight: 700; color: #2e7d32;">KES 0.00</div>
                        <div style="font-size: 11px; color: #999; margin-top: 4px;">Student Payments Total</div>
                    </div>
                    <div style="background: white; padding: 15px; border-radius: 6px; text-align: center;">
                        <div style="font-size: 12px; color: #666; margin-bottom: 5px;">TEACHER PAYROL</div>
                        <div id="teacher-expenses" style="font-size: 24px; font-weight: 700; color: #f57f17;">KES 0.00</div>
                        <div style="font-size: 11px; color: #999; margin-top: 4px;">All Weekly Summaries Total</div>
                    </div>
                    <div style="background: white; padding: 15px; border-radius: 6px; text-align: center;">
                        <div style="font-size: 12px; color: #666; margin-bottom: 5px;">NET BALANCE</div>
                        <div id="net-balance" style="font-size: 24px; font-weight: 700; color: #1565c0;">KES 0.00</div>
                        <div style="font-size: 11px; color: #999; margin-top: 4px;">Revenue − Payrol</div>
                    </div>
                </div>
            </div>
            
            <!-- ===== PAYROLL PERIOD OPEN / CLOSE (between Financial Summary and Record Attendance) ===== -->
            <div style="background: #fff8e1; border: 2px solid #ffd54f; padding: 16px 18px; border-radius: 8px; margin-bottom: 20px;">
                <div style="display: flex; align-items: center; gap: 16px; flex-wrap: wrap;">
                    <strong style="color: #e65100; white-space: nowrap;">⏱ Payroll Period:</strong>
                    
                    <!-- Term selector — always accessible, not locked -->
                    <select id="period-term" style="padding: 8px 12px; border: 2px solid #ffc107; border-radius: 6px; font-size: 14px; background: white; min-width: 100px;">
                        <option value="1">Term 1</option>
                        <option value="2">Term 2</option>
                        <option value="3">Term 3</option>
                    </select>
                    
                    <!-- Week selector — always accessible, not locked -->
                    <select id="period-week" style="padding: 8px 12px; border: 2px solid #ffc107; border-radius: 6px; font-size: 14px; background: white; min-width: 100px;">
                        <option value="">-- Week --</option>
                        <option value="1">Week 1</option>
                        <option value="2">Week 2</option>
                        <option value="3">Week 3</option>
                        <option value="4">Week 4</option>
                        <option value="5">Week 5</option>
                        <option value="6">Week 6</option>
                        <option value="7">Week 7</option>
                        <option value="8">Week 8</option>
                        <option value="9">Week 9</option>
                        <option value="10">Week 10</option>
                        <option value="11">Week 11</option>
                        <option value="12">Week 12</option>
                    </select>
                    
                    <span id="payroll-period-status" style="font-weight: 600; color: #888; flex: 1; min-width: 200px;">No active period — select Term &amp; Week, then open</span>
                    <button class="btn" onclick="openPayrollPeriod()" id="btn-open-period" style="background: #43a047; color: white; padding: 8px 16px; white-space: nowrap;">▶ Open Payroll Period</button>
                    <button class="btn" onclick="closePayrollPeriod()" id="btn-close-period" style="background: #e53935; color: white; padding: 8px 16px; white-space: nowrap;" disabled>⏹ Close &amp; Tally</button>
                </div>
                <div id="payroll-period-result" style="font-size: 13px; color: #555; margin-top: 8px;"></div>
            </div>
            
            <!-- Record Attendance Form -->
            <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px;" id="attendance-form-section">
                <h3 style="color: #f57f17; margin-bottom: 5px;">📝 Record Attendance</h3>
                <div id="active-period-display" style="font-size: 13px; color: #555; margin-bottom: 15px; font-weight: 600;"></div>
                <div class="form-group">
                    <label for="payroll-teacher">Select Teacher: *</label>
                    <select id="payroll-teacher" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                        <option value="">-- Select Teacher --</option>
                    </select>
                </div>
                
                <!-- Hidden fields that mirror the active period term/week for form submission -->
                <input type="hidden" id="payroll-term" value="">
                <input type="hidden" id="payroll-week" value="">
                
                <div class="form-group">
                    <label>Weekday Lessons (200 KES each):</label>
                    <div style="display: flex; gap: 15px; flex-wrap: wrap; margin-top: 10px;">
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="day-monday"> Monday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="day-tuesday"> Tuesday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="day-wednesday"> Wednesday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="day-thursday"> Thursday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="day-friday"> Friday
                        </label>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Weekend Lesson (500 KES):</label>
                    <div style="margin-top: 10px;">
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="day-saturday"> Saturday
                        </label>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Weekend Lesson (300 KES):</label>
                    <div style="margin-top: 10px;">
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="day-saturdaytp"> Saturday
                        </label>
                    </div>
                </div>
                
                <button class="btn btn-success" onclick="recordAttendance()" style="background: #f57f17;">💵 Record Attendance</button>
                
                <div id="attendance-result"></div>
            </div>
            
            <!-- Support Staff Attendance Form -->
            <div style="background: #f3e5f5; padding: 20px; border-radius: 8px; margin-bottom: 20px;" id="support-staff-section">
                <h3 style="color: #6a1b9a; margin-bottom: 5px;">👥 Record Support Staff Attendance</h3>
                <div id="active-period-display-staff" style="font-size: 13px; color: #555; margin-bottom: 15px; font-weight: 600;"></div>
                
                <!-- Hidden fields mirroring the active payroll period -->
                <input type="hidden" id="support-term" value="">
                <input type="hidden" id="support-week" value="">
                
                <div class="form-group">
                    <label for="support-staff-name">Select Support Staff: *</label>
                    <select id="support-staff-name" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                        <option value="">-- Select Staff Member --</option>
                        <option value="Mr. Jaoko">Mr. Jaoko</option>
                        <option value="Mr. Morris">Mr. Morris</option>
                        <option value="Mr. Collins">Mr. Collins</option>
                        <option value="Mad Rose">Mad Rose</option>
                        <option value="Madame Pamela">Madame Pamela</option>
                        <option value="Mr. Amon">Mr. Amon</option>
                        <option value="Madame Afflin">Madame Afflin</option>
                        <option value="Madame Filgona">Madame Filgona</option>
                        <option value="Madam Gladys">Madam Gladys</option>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>Weekday Attendance (200 KES each):</label>
                    <div style="display: flex; gap: 15px; flex-wrap: wrap; margin-top: 10px;">
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="support-monday"> Monday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="support-tuesday"> Tuesday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="support-wednesday"> Wednesday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="support-thursday"> Thursday
                        </label>
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="support-friday"> Friday
                        </label>
                    </div>
                </div>
                
                <div class="form-group">
                    <label>Weekend Duty (max 200 KES per weekend):</label>
                    <div style="margin-top: 10px;">
                        <label style="display: flex; align-items: center; gap: 5px; font-weight: normal;">
                            <input type="checkbox" id="support-saturday"> Saturday (200 KES — capped at max)
                        </label>
                    </div>
                    <small style="color: #666; font-size: 12px; margin-top: 5px; display: block;">Weekend duty pay is fixed at a maximum of 200 KES per weekend.</small>
                </div>
                
                <button class="btn btn-success" onclick="recordSupportStaffAttendance()" style="background: #6a1b9a;">👥 Record Support Staff Attendance</button>
                
                <div id="support-attendance-result"></div>
            </div>
            
            <!-- Print Weekly Summary Section -->
            <div style="background: white; padding: 20px; border-radius: 8px; margin-top: 20px;">
                <h3 style="color: #f57f17; margin-bottom: 15px;">🖨️ Print Weekly Summary</h3>
                <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-bottom: 15px;">
                    <div class="form-group">
                        <label for="print-term">Term:</label>
                        <select id="print-term" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                            <option value="1">Term 1</option>
                            <option value="2">Term 2</option>
                            <option value="3">Term 3</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label for="print-week">Week:</label>
                        <select id="print-week" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                            <option value="">-- Select Week --</option>
                            <option value="1">Week 1</option>
                            <option value="2">Week 2</option>
                            <option value="3">Week 3</option>
                            <option value="4">Week 4</option>
                            <option value="5">Week 5</option>
                            <option value="6">Week 6</option>
                            <option value="7">Week 7</option>
                            <option value="8">Week 8</option>
                            <option value="9">Week 9</option>
                            <option value="10">Week 10</option>
                            <option value="11">Week 11</option>
                            <option value="12">Week 12</option>
                        </select>
                    </div>
                    <div style="display: flex; align-items: flex-end;">
                        <button class="btn btn-primary" onclick="printWeeklySummary()" style="width: 100%; background: #f57f17;">🖨️ Print Summary</button>
                    </div>
                </div>
            </div>
            
            <!-- Add/Remove Teachers Section -->
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 20px;">
                <!-- Add Teacher -->
                <div style="background: #e8f5e9; padding: 20px; border-radius: 8px;">
                    <h3 style="color: #2e7d32; margin-bottom: 15px;">➕ Add New Teacher</h3>
                    <div class="form-group">
                        <label for="new-teacher-name">Teacher Name: *</label>
                        <input type="text" id="new-teacher-name" placeholder="e.g., Jane Doe" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px;">
                    </div>
                    <div class="form-group">
                        <label for="new-teacher-subject">Subject: *</label>
                        <input type="text" id="new-teacher-subject" placeholder="e.g., Chemistry" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px;">
                    </div>
                    <div class="form-group">
                        <label for="new-teacher-phone">Phone:</label>
                        <input type="text" id="new-teacher-phone" placeholder="e.g., 0712345678" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px;">
                    </div>
                    <div class="form-group">
                        <label for="new-teacher-email">Email:</label>
                        <input type="text" id="new-teacher-email" placeholder="e.g., teacher@school.com" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px;">
                    </div>
                    <button class="btn btn-success" onclick="addNewTeacher()">➕ Add Teacher</button>
                    <div id="add-teacher-result"></div>
                </div>
                
                <!-- Deactivate Teacher -->
                <div style="background: #ffebee; padding: 20px; border-radius: 8px;">
                    <h3 style="color: #c62828; margin-bottom: 15px;">🚪 Remove Teacher</h3>
                    <p style="color: #666; margin-bottom: 15px; font-size: 14px;">
                        Mark a teacher as inactive when they leave. Their records will be preserved.
                    </p>
                    <div class="form-group">
                        <label for="remove-teacher">Select Teacher: *</label>
                        <select id="remove-teacher" style="width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px;">
                            <option value="">-- Select Teacher --</option>
                        </select>
                    </div>
                    <button class="btn" onclick="deactivateTeacher()" style="background: #c62828; color: white;">🚪 Deactivate Teacher</button>
                    <div id="remove-teacher-result"></div>
                </div>
            </div>
        </div>
        
        <!-- ===== OLD PRINT TEMPLATE (kept for internal reference, hidden) ===== -->
        <div id="print-template" style="display: none;"></div>
        </div>
        
    <!-- ===== PAYROLL PRINT DOCUMENT (shown only when printing) ===== -->
    <div id="payroll-print-doc" style="display:none; font-family: Arial, sans-serif; font-size: 11pt; color: #000;">
        <!-- Header with Logo -->
        <div style="display: flex; align-items: center; border-bottom: 3px double #2d5016; padding-bottom: 10px; margin-bottom: 14px;">
            <img id="ppr-logo" src="" alt="Logo" style="height: 70px; width: 70px; object-fit: contain; margin-right: 18px; border: 1px solid #ccc; padding: 3px;">
            <div style="flex: 1; text-align: center;">
                <div style="font-size: 17pt; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">KIREMBE SECONDARY SCHOOL</div>
                <div style="font-size: 10pt; color: #444; margin-top: 2px;">REMEDIAL PROGRAMME — TEACHER &amp; SUPPORT STAFF PAYROLL</div>
                <div style="font-size: 9pt; color: #666; margin-top: 2px;" id="ppr-meta"></div>
            </div>
        </div>
        
        <!-- SECTION 1: Teacher Payroll -->
        <div style="font-size: 11pt; font-weight: bold; background: #2d5016; color: white; padding: 5px 10px; margin-bottom: 0;">SECTION A — TEACHER ATTENDANCE &amp; PAY</div>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;" id="ppr-teacher-table">
            <thead>
                <tr style="background: #e8f5e9;">
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: left; width: 22%;">Teacher Name</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: center; width: 14%;">Day(s)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: right; width: 12%;">Weekday (KES)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: right; width: 12%;">Weekend (KES)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: right; width: 12%;">Total (KES)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: center; width: 16%;">Timestamp</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: center; width: 12%;">Signature</th>
                </tr>
            </thead>
            <tbody id="ppr-teacher-tbody"></tbody>
            <tfoot>
                <tr style="background: #f1f8e9; font-weight: bold;">
                    <td colspan="2" style="border: 1px solid #999; padding: 5px 8px;">TEACHER TOTAL</td>
                    <td style="border: 1px solid #999; padding: 5px 8px; text-align: right;" id="ppr-t-weekday"></td>
                    <td style="border: 1px solid #999; padding: 5px 8px; text-align: right;" id="ppr-t-weekend"></td>
                    <td style="border: 1px solid #999; padding: 5px 8px; text-align: right;" id="ppr-t-total"></td>
                    <td colspan="2" style="border: 1px solid #999;"></td>
                </tr>
            </tfoot>
        </table>
        
        <!-- SECTION 2: Support Staff Payroll -->
        <div style="font-size: 11pt; font-weight: bold; background: #4a148c; color: white; padding: 5px 10px; margin-top: 12px; margin-bottom: 0;">SECTION B — SUPPORT STAFF ATTENDANCE &amp; PAY</div>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;" id="ppr-staff-table">
            <thead>
                <tr style="background: #f3e5f5;">
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: left; width: 22%;">Staff Name</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: center; width: 14%;">Day(s)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: right; width: 12%;">Weekday (KES)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: right; width: 12%;">Weekend (KES)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: right; width: 12%;">Total (KES)</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: center; width: 16%;">Timestamp</th>
                    <th style="border: 1px solid #999; padding: 5px 8px; text-align: center; width: 12%;">Signature</th>
                </tr>
            </thead>
            <tbody id="ppr-staff-tbody"></tbody>
            <tfoot>
                <tr style="background: #f8e8ff; font-weight: bold;">
                    <td colspan="2" style="border: 1px solid #999; padding: 5px 8px;">SUPPORT STAFF TOTAL</td>
                    <td style="border: 1px solid #999; padding: 5px 8px; text-align: right;" id="ppr-s-weekday"></td>
                    <td style="border: 1px solid #999; padding: 5px 8px; text-align: right;" id="ppr-s-weekend"></td>
                    <td style="border: 1px solid #999; padding: 5px 8px; text-align: right;" id="ppr-s-total"></td>
                    <td colspan="2" style="border: 1px solid #999;"></td>
                </tr>
            </tfoot>
        </table>
        
        <!-- Grand Total -->
        <table style="width: 40%; margin-left: auto; border-collapse: collapse; margin-top: 6px;">
            <tr style="background: #2d5016; color: white; font-weight: bold; font-size: 12pt;">
                <td style="border: 2px solid #2d5016; padding: 6px 12px;">GRAND TOTAL PAYROLL</td>
                <td style="border: 2px solid #2d5016; padding: 6px 12px; text-align: right;" id="ppr-grand-total"></td>
            </tr>
        </table>
        
        <!-- Approval Section -->
        <div style="margin-top: 28px; display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 30px; padding-top: 10px; border-top: 2px solid #ccc;">
            <div>
                <div style="font-size: 9pt; color: #555; margin-bottom: 32px;">Prepared by (Payroll Officer):</div>
                <div style="border-top: 1px solid #000; padding-top: 4px; font-size: 9pt;">Name &amp; Signature ________________</div>
                <div style="font-size: 9pt; margin-top: 4px;">Date: _______________</div>
            </div>
            <div>
                <div style="font-size: 9pt; color: #555; margin-bottom: 32px;">Verified by (Deputy Principal):</div>
                <div style="border-top: 1px solid #000; padding-top: 4px; font-size: 9pt;">Name &amp; Signature ________________</div>
                <div style="font-size: 9pt; margin-top: 4px;">Date: _______________</div>
            </div>
            <div>
                <div style="font-size: 9pt; color: #555; margin-bottom: 32px;">Approved by (Principal):</div>
                <div style="border-top: 1px solid #000; padding-top: 4px; font-size: 9pt;">Name &amp; Signature ________________</div>
                <div style="font-size: 9pt; margin-top: 4px;">Date: _______________</div>
            </div>
        </div>
        <div style="text-align: center; margin-top: 10px; font-size: 8pt; color: #777; border-top: 1px solid #ddd; padding-top: 6px;">
            Kirembe Secondary School — Remedial Programme Payroll | Generated: <span id="ppr-generated"></span>
        </div>
    </div>
    
    <script>
        // Global state
        let currentLearner = null;
        // Safe store for learner data passed to printReceipt (avoids apostrophe/quote issues in onclick)
        const _printDataStore = {};
        let exitLearnerData = null;
        let promotionPreviewData = null;
        
        // Auto-fill debt based on grade selection
        document.addEventListener('DOMContentLoaded', function() {
            const gradeSelect = document.getElementById('new-grade');
            const termSelect = document.getElementById('new-term');
            const debtInput = document.getElementById('new-debt');
            
            // Fee structure based on admission term
            const termFees = {
                'term1': 2200,  // Full year (Jan-Mar)
                'term2': 1400,  // 2 terms remaining (Apr-Jul)
                'term3': 600    // 1 term remaining (Aug-Dec)
            };
            
            if (termSelect && debtInput) {
                termSelect.addEventListener('change', function() {
                    const selectedTerm = this.value;
                    if (selectedTerm && termFees[selectedTerm]) {
                        debtInput.value = termFees[selectedTerm].toFixed(2);
                    } else {
                        debtInput.value = '';
                    }
                });
            }
        });
        
        // Tab switching
        function showTab(tabName) {
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.remove('active');
            });
            document.querySelectorAll('.tab').forEach(tab => {
                tab.classList.remove('active');
            });
            
            // Show selected tab
            document.getElementById(tabName).classList.add('active');
            event.target.classList.add('active');
            
            // Load data if needed
            if (tabName === 'dashboard') {
                loadDashboard();
            }
        }
        
        // Load dashboard data
        async function loadDashboard() {
            // Load class statistics chart
            loadClassChart();
            
            // Update stats bar with learner count (lightweight — no table rendering)
            try {
                const response = await fetch('/api/learners');
                const data = await response.json();
                if (data.success) {
                    document.getElementById('stats-text').textContent = `✓ ${data.learners.length} learners loaded`;
                } else {
                    document.getElementById('stats-text').textContent = 'Error loading learner count';
                }
            } catch (error) {
                document.getElementById('stats-text').textContent = 'Could not load stats';
            }
        }
        
        // Load class statistics chart
        // Load class statistics chart — live progression bars grouped by cohort
        async function loadClassChart() {
            const chartDiv = document.getElementById('class-chart');
            chartDiv.innerHTML = '<div class="loading"><div class="spinner"></div>Loading statistics...</div>';

            try {
                const response = await fetch('/api/class-statistics');
                const data = await response.json();

                if (!data.success || !data.statistics || data.statistics.length === 0) {
                    chartDiv.innerHTML = '<p style="color:#666;text-align:center;padding:40px;">No payment data available</p>';
                    return;
                }

                const stats = data.statistics;

                // --- Group cohorts together ---
                const groups = {};
                stats.forEach(s => {
                    const g = s.group_key || s.class_name;
                    if (!groups[g]) groups[g] = [];
                    groups[g].push(s);
                });

                // Colour palette — each group gets a distinct hue pair (darker=older year, brighter=newer year)
                const palette = [
                    ['#1a5276', '#2e86c1'],  // deep blue → bright blue
                    ['#1d6a2e', '#27ae60'],  // dark green → bright green
                    ['#7d6608', '#f1c40f'],  // dark gold → bright yellow
                    ['#6e2f1a', '#e67e22'],  // dark orange → bright orange
                    ['#4a235a', '#8e44ad'],  // dark purple → bright purple
                    ['#1b4f72', '#17a589'],  // navy → teal
                ];
                const fmt = n => 'KES ' + (n||0).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',');

                let html = '<div style="display:flex;flex-direction:column;gap:22px;">';
                let pi = 0;

                Object.entries(groups).forEach(([groupName, members]) => {
                    const [dark, bright] = palette[pi % palette.length]; pi++;
                    const maxInGroup = Math.max(...members.map(m => m.total_owed || m.total_collected || 1));
                    const isCohort = members.length > 1;

                    html += `<div style="background:#f9f9f9;border-radius:10px;padding:14px 18px;border-left:5px solid ${dark};">`;
                    if (isCohort) {
                        html += `<div style="font-size:12px;font-weight:700;color:${dark};letter-spacing:0.5px;margin-bottom:10px;text-transform:uppercase;">📈 ${groupName} — Cohort Progression</div>`;
                    }

                    members.forEach((s, idx) => {
                        const isNewer = idx === members.length - 1;
                        const barColor = isCohort ? (isNewer ? bright : dark) : bright;
                        const collected = s.total_collected || 0;
                        const owed = s.total_owed || collected || 1;
                        const pct = Math.min(100, owed > 0 ? (collected / owed) * 100 : 0);
                        const count = s.learner_count || '';

                        html += `
                        <div style="margin-bottom:${idx < members.length-1 ? '10px' : '0'};">
                            <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px;">
                                <span style="font-weight:600;font-size:14px;color:#333;">
                                    ${isNewer && isCohort ? '↳ ' : ''}${s.class_name}
                                    ${count ? `<span style="font-size:11px;color:#888;font-weight:normal;"> (${count} learners)</span>` : ''}
                                </span>
                                <span style="font-size:13px;font-weight:700;color:${barColor};">${fmt(collected)}</span>
                            </div>
                            <div style="background:#e9ecef;border-radius:6px;height:28px;overflow:hidden;position:relative;">
                                <div style="background:${barColor};height:100%;width:${pct.toFixed(1)}%;transition:width 0.8s ease;display:flex;align-items:center;padding:0 10px;min-width:${pct>0?'2px':'0'};">
                                    ${pct >= 12 ? `<span style="color:white;font-size:12px;font-weight:600;">${pct.toFixed(0)}%</span>` : ''}
                                </div>
                                ${pct < 12 && pct > 0 ? `<span style="position:absolute;left:8px;top:50%;transform:translateY(-50%);font-size:12px;font-weight:600;color:#555;">${pct.toFixed(0)}%</span>` : ''}
                            </div>
                            ${owed > collected ? `<div style="font-size:11px;color:#999;margin-top:2px;text-align:right;">Outstanding: ${fmt(owed - collected)}</div>` : ''}
                        </div>`;
                    });

                    html += '</div>';
                });

                html += '</div>';
                chartDiv.innerHTML = html;

            } catch (error) {
                chartDiv.innerHTML = `<p style="color:#dc3545;text-align:center;padding:40px;">Error loading statistics: ${error.message}</p>`;
            }
        }
        
        // Search learner
        async function searchLearner() {
            const admissionNumber = document.getElementById('search-admission').value.trim();
            const resultsDiv = document.getElementById('search-results');
            
            if (!admissionNumber) {
                resultsDiv.innerHTML = '<div class="alert alert-error">Please enter an admission number</div>';
                return;
            }
            
            resultsDiv.innerHTML = '<div class="loading"><div class="spinner"></div>Searching...</div>';
            
            try {
                const response = await fetch(`/api/learner/${admissionNumber}`);
                const data = await response.json();
                
                if (data.success) {
                    const learner = data.details;
                    let balanceClass = 'balance-positive';
                    let balanceText = `KES ${Math.abs(learner.balance).toFixed(2)}`;
                    
                    if (learner.balance < 0) {
                        balanceClass = 'balance-credit';
                        balanceText += ' (Credit/Overpayment)';
                    } else if (learner.balance === 0) {
                        balanceClass = 'balance-zero';
                    }
                    
                    let historyHTML = '';
                    if (learner.payment_history.length > 0) {
                        historyHTML = '<div class="payment-history">';
                        learner.payment_history.reverse().forEach((payment, index) => {
                            const originalIndex = learner.payment_history.length - 1 - index;
                            // Store learner safely to avoid apostrophe issues in onclick attribute
                            const storeKey = `${learner.admission_number}_${originalIndex}`;
                            _printDataStore[storeKey] = {
                                learner: learner,
                                prevBalance: learner.balance + payment.amount,
                                currentBalance: learner.balance
                            };
                            historyHTML += `
                                <div class="payment-item" style="display: flex; justify-content: space-between; align-items: center; padding: 12px; background: white; border-radius: 4px; margin-bottom: 8px; border-left: 4px solid #28a745;">
                                    <div>
                                        <div style="font-size: 13px; color: #666;">${payment.timestamp}</div>
                                        <div style="color: #28a745; font-weight: 600; font-size: 15px;">KES ${payment.amount.toFixed(2)}</div>
                                    </div>
                                    <div style="display: flex; gap: 8px;">
                                        <button onclick="printReceiptFromStore('${storeKey}', ${payment.amount}, '${payment.timestamp}')" 
                                                style="padding: 6px 12px; background: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">
                                            🖨️ Print
                                        </button>
                                        <button onclick="showEditPaymentModal('${learner.admission_number}', ${originalIndex}, ${payment.amount}, '${payment.timestamp}')" 
                                                style="padding: 6px 12px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">
                                            ✏️ Edit
                                        </button>
                                        <button onclick="showDeletePaymentModal('${learner.admission_number}', ${originalIndex}, ${payment.amount}, '${payment.timestamp}')" 
                                                style="padding: 6px 12px; background: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">
                                            🗑️ Delete
                                        </button>
                                    </div>
                                </div>
                            `;
                        });
                        historyHTML += '</div>';
                    } else {
                        historyHTML = '<p style="color: #666; margin-top: 10px;">No payments recorded</p>';
                    }
                    
                    resultsDiv.innerHTML = `
                        <div class="result-card">
                            <h3>✓ Learner Found</h3>
                            <div class="detail-row">
                                <span class="detail-label">Name:</span>
                                <span class="detail-value">${learner.name}</span>
                            </div>
                            <div class="detail-row">
                                <span class="detail-label">Admission Number:</span>
                                <span class="detail-value">${learner.admission_number}</span>
                            </div>
                            <div class="detail-row">
                                <span class="detail-label">Class/Grade:</span>
                                <span class="detail-value">${learner.source_sheet || 'Unknown'}</span>
                            </div>
                            <div class="detail-row">
                                <span class="detail-label">Total Owed:</span>
                                <span class="detail-value">KES ${learner.total_owed.toFixed(2)}</span>
                            </div>
                            <div class="detail-row">
                                <span class="detail-label">Total Paid:</span>
                                <span class="detail-value">KES ${learner.total_paid.toFixed(2)}</span>
                            </div>
                            <div class="detail-row">
                                <span class="detail-label">Balance:</span>
                                <span class="detail-value ${balanceClass}">${balanceText}</span>
                            </div>
                            <h3 style="margin-top: 20px;">Payment History (${learner.payment_history.length} payments)</h3>
                            ${historyHTML}
                        </div>
                    `;
                } else {
                    resultsDiv.innerHTML = `<div class="alert alert-error">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultsDiv.innerHTML = `<div class="alert alert-error">Error: ${error.message}</div>`;
            }
        }
        
        // Lookup learner for payment
        async function lookupLearner() {
            const admissionNumber = document.getElementById('payment-admission').value.trim();
            const infoDiv = document.getElementById('learner-info');
            
            if (!admissionNumber) {
                infoDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 10px;">Please enter an admission number</div>';
                return;
            }
            
            try {
                const response = await fetch(`/api/learner/${admissionNumber}`);
                const data = await response.json();
                
                if (data.success) {
                    currentLearner = data.details;
                    let balanceText = `KES ${Math.abs(currentLearner.balance).toFixed(2)}`;
                    if (currentLearner.balance < 0) {
                        balanceText += ' (Credit)';
                    }
                    
                    infoDiv.innerHTML = `
                        <div class="learner-info">
                            ✓ ${currentLearner.name} | Current Balance: ${balanceText}
                        </div>
                    `;
                } else {
                    currentLearner = null;
                    infoDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 10px;">❌ ${data.message}</div>`;
                }
            } catch (error) {
                currentLearner = null;
                infoDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 10px;">Error: ${error.message}</div>`;
            }
        }
        
        // Record payment
        async function recordPayment() {
            const admissionNumber = document.getElementById('payment-admission').value.trim();
            const amount = document.getElementById('payment-amount').value.trim();
            const paymentDate = document.getElementById('payment-date').value;
            const resultDiv = document.getElementById('payment-result');
            
            if (!admissionNumber) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please enter an admission number</div>';
                return;
            }
            
            if (!amount || parseFloat(amount) <= 0) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please enter a valid payment amount</div>';
                return;
            }
            
            resultDiv.innerHTML = '<div class="loading" style="margin-top: 20px;"><div class="spinner"></div>Recording payment...</div>';
            
            try {
                const requestBody = {
                    admission_number: admissionNumber,
                    amount: parseFloat(amount)
                };
                
                // Add payment date if provided
                if (paymentDate) {
                    requestBody.payment_date = paymentDate;
                }
                
                const response = await fetch('/api/record-payment', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });
                
                const data = await response.json();
                
                if (data.success) {
                    // Use data from API response for receipt (more reliable than currentLearner)
                    const learnerForReceipt = {
                        name: data.learner_name || (currentLearner ? currentLearner.name : admissionNumber),
                        admission_number: data.admission_number || admissionNumber,
                        source_sheet: data.source_sheet || (currentLearner ? currentLearner.source_sheet : 'N/A')
                    };
                    const prevBalance = data.previous_balance !== undefined ? data.previous_balance : (currentLearner ? currentLearner.balance : 0);
                    const payDate = data.payment_date || paymentDate || new Date().toISOString();
                    
                    // Store safely to avoid apostrophe/quote issues in onclick attribute
                    const rpStoreKey = 'rp_' + Date.now();
                    _printDataStore[rpStoreKey] = {
                        learner: learnerForReceipt,
                        prevBalance: prevBalance,
                        currentBalance: data.new_balance
                    };
                    
                    resultDiv.innerHTML = `
                        <div class="alert alert-success" style="margin-top: 20px;">
                            ✓ ${data.message}<br>
                            <strong>New Balance:</strong> KES ${data.new_balance.toFixed(2)}<br>
                            <button class="btn btn-primary" style="margin-top: 15px;" onclick="printReceiptFromStore('${rpStoreKey}', ${data.payment_amount || parseFloat(amount)}, '${payDate}')">
                                🖨️ Print Receipt
                            </button>
                        </div>
                    `;
                    
                    // Clear form
                    document.getElementById('payment-amount').value = '';
                    document.getElementById('payment-date').value = '';
                    document.getElementById('learner-info').innerHTML = '';
                    currentLearner = null;
                    
                    // Refresh dashboard and financial summary
                    loadDashboard();
                    loadFinancialSummary();
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 20px;">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 20px;">Error: ${error.message}</div>`;
            }
        }
        
        // Edit Payment Modal Functions
        let editPaymentData = {};
        
        function showEditPaymentModal(admissionNumber, paymentIndex, currentAmount, currentDate) {
            editPaymentData = { admissionNumber, paymentIndex };
            document.getElementById('edit-amount').value = currentAmount;
            
            // Parse and format date
            const date = new Date(currentDate);
            const dateStr = date.toISOString().split('T')[0];
            document.getElementById('edit-date').value = dateStr;
            
            document.getElementById('editPaymentModal').style.display = 'block';
            document.getElementById('edit-result').innerHTML = '';
        }
        
        function closeEditModal() {
            document.getElementById('editPaymentModal').style.display = 'none';
            editPaymentData = {};
        }
        
        async function saveEditPayment() {
            const newAmount = document.getElementById('edit-amount').value;
            const newDate = document.getElementById('edit-date').value;
            const resultDiv = document.getElementById('edit-result');
            
            if (!newAmount || parseFloat(newAmount) <= 0) {
                resultDiv.innerHTML = '<div class="alert alert-error">Please enter a valid amount</div>';
                return;
            }
            
            resultDiv.innerHTML = '<div style="color: #666;">Saving changes...</div>';
            
            try {
                const response = await fetch('/api/edit-payment', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        admission_number: editPaymentData.admissionNumber,
                        payment_index: editPaymentData.paymentIndex,
                        new_amount: parseFloat(newAmount),
                        new_date: newDate
                    })
                });
                
                const data = await response.json();
                
                if (data.success) {
                    resultDiv.innerHTML = '<div class="alert alert-success">✓ Payment updated successfully!</div>';
                    setTimeout(() => {
                        closeEditModal();
                        searchLearner(); // Refresh the search results
                        loadFinancialSummary(); // Update NET REVENUE
                    }, 1500);
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error">Error: ${error.message}</div>`;
            }
        }
        
        // Delete Payment Modal Functions
        let deletePaymentData = {};
        
        function showDeletePaymentModal(admissionNumber, paymentIndex, amount, date) {
            deletePaymentData = { admissionNumber, paymentIndex };
            document.getElementById('delete-message').innerHTML = 
                `Are you sure you want to delete this payment?<br><br>
                <strong>Amount:</strong> KES ${parseFloat(amount).toFixed(2)}<br>
                <strong>Date:</strong> ${date}`;
            document.getElementById('deletePaymentModal').style.display = 'block';
            document.getElementById('delete-result').innerHTML = '';
        }
        
        function closeDeleteModal() {
            document.getElementById('deletePaymentModal').style.display = 'none';
            deletePaymentData = {};
        }
        
        async function confirmDeletePayment() {
            const resultDiv = document.getElementById('delete-result');
            resultDiv.innerHTML = '<div style="color: #666;">Deleting payment...</div>';
            
            try {
                const response = await fetch('/api/delete-payment', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        admission_number: deletePaymentData.admissionNumber,
                        payment_index: deletePaymentData.paymentIndex
                    })
                });
                
                const data = await response.json();
                
                if (data.success) {
                    resultDiv.innerHTML = '<div class="alert alert-success">✓ Payment deleted successfully!</div>';
                    setTimeout(() => {
                        closeDeleteModal();
                        searchLearner(); // Refresh the search results
                        loadFinancialSummary(); // Update NET REVENUE
                    }, 1500);
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error">Error: ${error.message}</div>`;
            }
        }
        
        // Close modals when clicking outside
        window.onclick = function(event) {
            const editModal = document.getElementById('editPaymentModal');
            const deleteModal = document.getElementById('deletePaymentModal');
            if (event.target == editModal) {
                closeEditModal();
            }
            if (event.target == deleteModal) {
                closeDeleteModal();
            }
        }
        
        // Print Receipt Function
        function printReceipt(learnerData, paymentAmount, paymentDate, previousBalance, currentBalance) {
            // Generate receipt number (timestamp-based)
            const receiptNumber = 'RCP-' + Date.now();
            
            // Format date
            const formattedDate = new Date(paymentDate).toLocaleDateString('en-GB', {
                day: '2-digit',
                month: 'short',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });
            
            // Fill in receipt details
            document.getElementById('receipt-number').textContent = receiptNumber;
            document.getElementById('receipt-date').textContent = formattedDate;
            document.getElementById('receipt-student-name').textContent = learnerData.name;
            document.getElementById('receipt-admission').textContent = learnerData.admission_number;
            document.getElementById('receipt-class').textContent = learnerData.source_sheet || 'N/A';
            document.getElementById('receipt-amount').textContent = 'KES ' + parseFloat(paymentAmount).toFixed(2);
            document.getElementById('receipt-prev-balance').textContent = 'KES ' + parseFloat(previousBalance).toFixed(2);
            document.getElementById('receipt-current-balance').textContent = 'KES ' + parseFloat(currentBalance).toFixed(2);
            document.getElementById('receipt-print-date').textContent = new Date().toLocaleString('en-GB');
            
            // Ensure payroll print doc is hidden and body class is set correctly
            document.getElementById('payroll-print-doc').style.display = 'none';
            document.body.classList.remove('printing-payroll');
            
            // Show receipt
            const receiptEl = document.getElementById('receipt');
            receiptEl.style.display = 'block';
            
            // Tag body for receipt print CSS
            document.body.classList.add('printing-receipt');
            
            setTimeout(() => {
                window.print();
                // Clean up after print
                receiptEl.style.display = 'none';
                document.body.classList.remove('printing-receipt');
            }, 150);
        }
        
        // Wrapper: retrieve stored learner data and call printReceipt safely
        // Avoids apostrophe/single-quote issues when embedding JSON in onclick attributes
        function printReceiptFromStore(storeKey, paymentAmount, paymentDate) {
            const stored = _printDataStore[storeKey];
            if (!stored) {
                alert('Print data not available. Please search the learner again.');
                return;
            }
            printReceipt(stored.learner, paymentAmount, paymentDate, stored.prevBalance, stored.currentBalance);
        }
        
        // Add New Learner Function
        async function addNewLearner() {
            const admission = document.getElementById('new-admission').value.trim();
            const name = document.getElementById('new-name').value.trim();
            const grade = document.getElementById('new-grade').value;
            const term = document.getElementById('new-term').value;
            const arrears = document.getElementById('new-arrears').value || '0';
            const debt = document.getElementById('new-debt').value;
            const admissionDate = document.getElementById('new-admission-date').value;
            const resultDiv = document.getElementById('add-learner-result');
            
            // Validation
            if (!admission) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please enter an admission number</div>';
                return;
            }
            
            if (!name) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please enter learner name</div>';
                return;
            }
            
            if (!grade) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please select a grade</div>';
                return;
            }
            
            if (!term) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please select admission term</div>';
                return;
            }
            
            if (!debt) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Fee amount not calculated. Please select a term.</div>';
                return;
            }
            
            resultDiv.innerHTML = '<div class="loading" style="margin-top: 20px;"><div class="spinner"></div>Adding learner...</div>';
            
            try {
                const requestBody = {
                    admission_number: admission,
                    name: name,
                    grade: grade,
                    term: term,
                    arrears: parseFloat(arrears),
                    debt: parseFloat(debt),
                    admission_date: admissionDate || new Date().toISOString().split('T')[0]
                };
                
                const response = await fetch('/api/add-learner', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });
                
                const data = await response.json();
                
                if (data.success) {
                    const termNames = {
                        'term1': 'Term 1 (Jan-Mar)',
                        'term2': 'Term 2 (Apr-Jul)',
                        'term3': 'Term 3 (Aug-Dec)'
                    };
                    
                    resultDiv.innerHTML = `
                        <div class="alert alert-success" style="margin-top: 20px;">
                            ✓ ${data.message}<br>
                            <strong>Admission:</strong> ${admission}<br>
                            <strong>Name:</strong> ${name}<br>
                            <strong>Grade:</strong> ${grade}<br>
                            <strong>Admission Term:</strong> ${termNames[term]}<br>
                            <strong>Fee Amount:</strong> KES ${parseFloat(debt).toFixed(2)}<br>
                            <strong>Previous Arrears:</strong> KES ${parseFloat(arrears).toFixed(2)}<br>
                            <strong>Total Owed:</strong> KES ${(parseFloat(arrears) + parseFloat(debt)).toFixed(2)}
                        </div>
                    `;
                    
                    // Clear form
                    document.getElementById('new-admission').value = '';
                    document.getElementById('new-name').value = '';
                    document.getElementById('new-grade').value = '';
                    document.getElementById('new-term').value = '';
                    document.getElementById('new-arrears').value = '0';
                    document.getElementById('new-debt').value = '';
                    document.getElementById('new-admission-date').value = '';
                    
                    // Refresh dashboard
                    loadDashboard();
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 20px;">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 20px;">Error: ${error.message}</div>`;
            }
        }
        
        // Lookup Learner for Exit
        async function lookupLearnerForExit() {
            const admissionNumber = document.getElementById('exit-admission').value.trim();
            const infoDiv = document.getElementById('exit-learner-info');
            
            if (!admissionNumber) {
                infoDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 10px;">Please enter an admission number</div>';
                return;
            }
            
            infoDiv.innerHTML = '<div class="loading" style="margin-top: 10px;"><div class="spinner"></div>Looking up learner...</div>';
            
            try {
                const response = await fetch(`/api/learner/${admissionNumber}`);
                const data = await response.json();
                
                if (data.success) {
                    exitLearnerData = data.details;
                    infoDiv.innerHTML = `
                        <div class="learner-info" style="margin-top: 10px;">
                            <strong>Name:</strong> ${data.details.name}<br>
                            <strong>Grade:</strong> ${data.details.source_sheet}<br>
                            <strong>Total Owed:</strong> KES ${data.details.total_owed.toFixed(2)}<br>
                            <strong>Total Paid:</strong> KES ${data.details.total_paid.toFixed(2)}<br>
                            <strong>Balance:</strong> KES ${Math.abs(data.details.balance).toFixed(2)} ${data.details.balance < 0 ? '(Credit)' : ''}
                        </div>
                    `;
                } else {
                    exitLearnerData = null;
                    infoDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 10px;">❌ ${data.message}</div>`;
                }
            } catch (error) {
                exitLearnerData = null;
                infoDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 10px;">Error: ${error.message}</div>`;
            }
        }
        
        // Exit Learner Function
        async function exitLearner() {
            const admission = document.getElementById('exit-admission').value.trim();
            const exitDate = document.getElementById('exit-date').value;
            const reason = document.getElementById('exit-reason').value;
            const notes = document.getElementById('exit-notes').value.trim();
            const resultDiv = document.getElementById('exit-learner-result');
            
            if (!admission) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please enter an admission number</div>';
                return;
            }
            
            if (!exitLearnerData) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 20px;">Please lookup the learner first</div>';
                return;
            }
            
            // Confirm exit
            if (!confirm(`Are you sure you want to exit ${exitLearnerData.name} (${admission})? This will mark them as inactive but preserve all records.`)) {
                return;
            }
            
            resultDiv.innerHTML = '<div class="loading" style="margin-top: 20px;"><div class="spinner"></div>Exiting learner...</div>';
            
            try {
                const requestBody = {
                    admission_number: admission,
                    exit_date: exitDate || new Date().toISOString().split('T')[0],
                    reason: reason,
                    notes: notes
                };
                
                const response = await fetch('/api/exit-learner', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });
                
                const data = await response.json();
                
                if (data.success) {
                    resultDiv.innerHTML = `
                        <div class="alert alert-success" style="margin-top: 20px;">
                            ✓ ${data.message}<br>
                            <strong>Learner:</strong> ${exitLearnerData.name}<br>
                            <strong>Exit Date:</strong> ${requestBody.exit_date}<br>
                            <strong>Reason:</strong> ${reason}<br>
                            <strong>Final Balance:</strong> KES ${Math.abs(exitLearnerData.balance).toFixed(2)} ${exitLearnerData.balance < 0 ? '(Credit)' : '(Owed)'}
                        </div>
                    `;
                    
                    // Clear form
                    document.getElementById('exit-admission').value = '';
                    document.getElementById('exit-date').value = '';
                    document.getElementById('exit-reason').value = 'Graduated';
                    document.getElementById('exit-notes').value = '';
                    document.getElementById('exit-learner-info').innerHTML = '';
                    exitLearnerData = null;
                    
                    // Refresh dashboard
                    loadDashboard();
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 20px;">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 20px;">Error: ${error.message}</div>`;
            }
        }
        
        // Preview Promotion
        async function previewPromotion() {
            const year = document.getElementById('promote-year').value;
            const promoteGrade10 = document.getElementById('promote-grade10').checked;
            const promoteForm3 = document.getElementById('promote-form3').checked;
            const promoteForm4 = document.getElementById('promote-form4').checked;
            const previewDiv = document.getElementById('promote-preview');
            
            if (!year) {
                previewDiv.innerHTML = '<div class="alert alert-error">Please enter the new academic year</div>';
                return;
            }
            
            if (!promoteGrade10 && !promoteForm3 && !promoteForm4) {
                previewDiv.innerHTML = '<div class="alert alert-error">Please select at least one class to promote</div>';
                return;
            }
            
            previewDiv.innerHTML = '<div class="loading"><div class="spinner"></div>Calculating promotion preview...</div>';
            
            try {
                const requestBody = {
                    year: parseInt(year),
                    promote_grade10: promoteGrade10,
                    promote_form3: promoteForm3,
                    promote_form4: promoteForm4,
                    preview_only: true
                };
                
                const response = await fetch('/api/promote-students', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });
                
                const data = await response.json();
                
                if (data.success) {
                    promotionPreviewData = data.preview;
                    let previewHTML = '<div class="result-card" style="background: #e7f3ff;">';
                    previewHTML += '<h3 style="color: #004085;">📊 Promotion Preview</h3>';
                    
                    if (data.preview.grade10) {
                        previewHTML += `
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 6px;">
                                <h4 style="color: #2d5016;">Grade 10 → Form 3 ${year}</h4>
                                <p><strong>Students:</strong> ${data.preview.grade10.count}</p>
                                <p><strong>Total Arrears to Carry:</strong> KES ${data.preview.grade10.total_arrears.toFixed(2)}</p>
                                <p><strong>New Fees:</strong> KES ${data.preview.grade10.new_fees.toFixed(2)}</p>
                                <p><strong>Total Owed (New Year):</strong> KES ${data.preview.grade10.total_new_owed.toFixed(2)}</p>
                            </div>
                        `;
                    }
                    
                    if (data.preview.form3) {
                        previewHTML += `
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 6px;">
                                <h4 style="color: #2d5016;">Form 3 → Form 4 ${year}</h4>
                                <p><strong>Students:</strong> ${data.preview.form3.count}</p>
                                <p><strong>Total Arrears to Carry:</strong> KES ${data.preview.form3.total_arrears.toFixed(2)}</p>
                                <p><strong>New Fees:</strong> KES ${data.preview.form3.new_fees.toFixed(2)}</p>
                                <p><strong>Total Owed (New Year):</strong> KES ${data.preview.form3.total_new_owed.toFixed(2)}</p>
                            </div>
                        `;
                    }
                    
                    if (data.preview.form4) {
                        previewHTML += `
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 6px;">
                                <h4 style="color: #2d5016;">Form 4 → Graduate</h4>
                                <p><strong>Students:</strong> ${data.preview.form4.count}</p>
                                <p><strong>Final Balances:</strong> KES ${data.preview.form4.total_final_balance.toFixed(2)}</p>
                                <p style="color: #856404;"><em>These students will be marked as graduated</em></p>
                            </div>
                        `;
                    }
                    
                    previewHTML += '</div>';
                    previewDiv.innerHTML = previewHTML;
                } else {
                    previewDiv.innerHTML = `<div class="alert alert-error">❌ ${data.message}</div>`;
                }
            } catch (error) {
                previewDiv.innerHTML = `<div class="alert alert-error">Error: ${error.message}</div>`;
            }
        }
        
        // Promote Students
        async function promoteStudents() {
            const year = document.getElementById('promote-year').value;
            const promoteGrade10 = document.getElementById('promote-grade10').checked;
            const promoteForm3 = document.getElementById('promote-form3').checked;
            const promoteForm4 = document.getElementById('promote-form4').checked;
            const resultDiv = document.getElementById('promote-result');
            
            if (!year) {
                resultDiv.innerHTML = '<div class="alert alert-error">Please enter the new academic year</div>';
                return;
            }
            
            if (!promoteGrade10 && !promoteForm3 && !promoteForm4) {
                resultDiv.innerHTML = '<div class="alert alert-error">Please select at least one class to promote</div>';
                return;
            }
            
            // Confirm action
            if (!confirm(`⚠️ IMPORTANT: This will promote students to the next grade and create new sheets for ${year}.\n\nA backup will be created automatically.\n\nAre you sure you want to proceed?`)) {
                return;
            }
            
            resultDiv.innerHTML = '<div class="loading"><div class="spinner"></div>Promoting students... This may take a moment...</div>';
            
            try {
                const requestBody = {
                    year: parseInt(year),
                    promote_grade10: promoteGrade10,
                    promote_form3: promoteForm3,
                    promote_form4: promoteForm4,
                    preview_only: false
                };
                
                const response = await fetch('/api/promote-students', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });
                
                const data = await response.json();
                
                if (data.success) {
                    let successHTML = '<div class="alert alert-success">';
                    successHTML += `<h3>✓ ${data.message}</h3>`;
                    successHTML += `<p><strong>Backup Created:</strong> ${data.backup_file}</p>`;
                    
                    if (data.results.grade10) {
                        successHTML += `<p><strong>Grade 10:</strong> ${data.results.grade10.promoted} students promoted to Form 3 ${year}</p>`;
                    }
                    
                    if (data.results.form3) {
                        successHTML += `<p><strong>Form 3:</strong> ${data.results.form3.promoted} students promoted to Form 4 ${year}</p>`;
                    }
                    
                    if (data.results.form4) {
                        successHTML += `<p><strong>Form 4:</strong> ${data.results.form4.graduated} students marked as graduated</p>`;
                    }
                    
                    successHTML += '<p style="margin-top: 15px;"><em>Please restart the application to see the new sheets.</em></p>';
                    successHTML += '</div>';
                    
                    resultDiv.innerHTML = successHTML;
                    
                    // Clear preview
                    document.getElementById('promote-preview').innerHTML = '';
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error">Error: ${error.message}</div>`;
            }
        }
        
        // Load dashboard on page load
        window.addEventListener('load', () => {
            loadDashboard();
        });
        
        // Teacher Payroll Functions
        
        // Print Weekly Payroll Summary
        async function printWeeklySummary() {
            const term = document.getElementById('print-term').value;
            const week = document.getElementById('print-week').value;
            
            if (!week) {
                alert('Please select a week to print');
                return;
            }
            
            try {
                // Fetch teacher attendance records
                const response = await fetch(`/api/payroll-records?term=${term}&week=${week}`);
                const data = await response.json();
                
                if (!data.success) {
                    alert('Error loading payroll data: ' + data.message);
                    return;
                }
                
                // Fetch support staff attendance records
                let staffRecords = [];
                let staffError = null;
                try {
                    const sResp = await fetch(`/api/support-staff-records?term=${term}&week=${week}`);
                    const sData = await sResp.json();
                    if (sData.success) {
                        staffRecords = sData.records || [];
                    } else {
                        staffError = sData.message || 'Unknown error fetching staff records';
                    }
                } catch(e) {
                    staffError = e.message;
                }
                
                const teacherRecords = (data.records || []).filter(r =>
                    String(r.term) === String(term) && String(r.week) === String(week)
                );
                
                // A week is only valid if teacher records exist for it.
                // Support staff entries without teacher data are phantom/unentered weeks —
                // block printing to avoid phantom KES 1,800 totals.
                if (teacherRecords.length === 0) {
                    alert(`No teacher records found for Term ${term}, Week ${week}.\n\nThis week has not been entered yet. Support staff records cannot be printed without corresponding teacher attendance.`);
                    return;
                }
                
                // Set logo
                const logoImg = document.querySelector('.header-logo');
                document.getElementById('ppr-logo').src = logoImg ? logoImg.src : '';
                
                // Meta info
                document.getElementById('ppr-meta').textContent =
                    `Term ${term}  |  Week ${week}  |  Date Printed: ${new Date().toLocaleDateString('en-GB', {day:'2-digit',month:'short',year:'numeric'})}`;
                document.getElementById('ppr-generated').textContent = new Date().toLocaleString('en-GB');
                
                // --- Teacher section ---
                let tWkday = 0, tWkend = 0, tTotal = 0;
                const tBody = document.getElementById('ppr-teacher-tbody');
                tBody.innerHTML = '';
                teacherRecords.forEach(r => {
                    const wkd = parseFloat(r.weekday_pay || 0);
                    const wke = parseFloat(r.weekend_pay || 0);
                    const tot = parseFloat(r.total_pay || 0);
                    tWkday += wkd; tWkend += wke; tTotal += tot;
                    const days = typeof r.days_attended === 'object'
                        ? Object.keys(r.days_attended).filter(d => r.days_attended[d]).map(d => d.charAt(0).toUpperCase()+d.slice(1)).join(', ')
                        : (r.days_attended || 'N/A');
                    // Format timestamp
                    let ts = r.date_recorded || r.timestamp || '';
                    try { ts = new Date(ts).toLocaleString('en-GB', {day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}); } catch(e){}
                    tBody.innerHTML += `<tr>
                        <td style="border:1px solid #bbb;padding:4px 7px;">${r.teacher_name||'—'}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:center;">${days}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:right;">${wkd>0?wkd.toFixed(2):''}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:right;">${wke>0?wke.toFixed(2):''}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:right;font-weight:bold;">${tot.toFixed(2)}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:center;font-size:9pt;color:#555;">${ts}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;"></td>
                    </tr>`;
                });
                document.getElementById('ppr-t-weekday').textContent = tWkday > 0 ? tWkday.toFixed(2) : '';
                document.getElementById('ppr-t-weekend').textContent = tWkend > 0 ? tWkend.toFixed(2) : '';
                document.getElementById('ppr-t-total').textContent = tTotal.toFixed(2);
                
                // --- Support staff section ---
                let sWkday = 0, sWkend = 0, sTotal = 0;
                const sBody = document.getElementById('ppr-staff-tbody');
                sBody.innerHTML = '';
                staffRecords.forEach(r => {
                    const wkd = parseFloat(r.weekday_pay || 0);
                    const wke = parseFloat(r.weekend_pay || 0);
                    const tot = parseFloat(r.total_pay || 0);
                    sWkday += wkd; sWkend += wke; sTotal += tot;
                    let ts = r.date_recorded || r.timestamp || '';
                    try { ts = new Date(ts).toLocaleString('en-GB', {day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'}); } catch(e){}
                    sBody.innerHTML += `<tr>
                        <td style="border:1px solid #bbb;padding:4px 7px;">${r.staff_name||'—'}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:center;">${r.days_attended||'N/A'}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:right;">${wkd>0?wkd.toFixed(2):''}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:right;">${wke>0?wke.toFixed(2):''}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:right;font-weight:bold;">${tot.toFixed(2)}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;text-align:center;font-size:9pt;color:#555;">${ts}</td>
                        <td style="border:1px solid #bbb;padding:4px 7px;"></td>
                    </tr>`;
                });
                if (staffRecords.length === 0) {
                    const msg = staffError 
                        ? `Error loading staff records: ${staffError}` 
                        : 'No support staff records for this period';
                    sBody.innerHTML = `<tr><td colspan="7" style="border:1px solid #bbb;padding:6px;text-align:center;color:#999;">${msg}</td></tr>`;
                }
                document.getElementById('ppr-s-weekday').textContent = sWkday > 0 ? sWkday.toFixed(2) : '';
                document.getElementById('ppr-s-weekend').textContent = sWkend > 0 ? sWkend.toFixed(2) : '';
                document.getElementById('ppr-s-total').textContent = sTotal.toFixed(2);
                
                // Grand total
                document.getElementById('ppr-grand-total').textContent = 'KES ' + (tTotal + sTotal).toFixed(2);
                
                // Show and print
                const printDoc = document.getElementById('payroll-print-doc');
                
                // Ensure receipt is hidden and body is tagged for payroll printing
                document.getElementById('receipt').style.display = 'none';
                document.body.classList.remove('printing-receipt');
                
                printDoc.style.display = 'block';
                document.body.classList.add('printing-payroll');
                
                setTimeout(() => {
                    window.print();
                    // Clean up
                    printDoc.style.display = 'none';
                    document.body.classList.remove('printing-payroll');
                }, 200);
                
            } catch (error) {
                alert('Error printing payroll: ' + error.message);
            }
        }
        
        // ===== PAYROLL PERIOD OPEN / CLOSE =====
        let activePayrollPeriod = null;
        
        // Initially lock attendance form until a period is opened
        function setAttendanceLocked(locked) {
            const section = document.getElementById('attendance-form-section');
            if (!section) return;
            if (locked) {
                section.style.opacity = '0.4';
                section.style.pointerEvents = 'none';
                section.title = 'Open a payroll period first before recording attendance';
            } else {
                section.style.opacity = '1';
                section.style.pointerEvents = '';
                section.title = '';
            }
            // Same for support staff section
            const staffSection = document.querySelector('#payroll .tab-content div[style*="f3e5f5"]') ||
                                  document.getElementById('support-staff-section');
            if (staffSection) {
                staffSection.style.opacity = locked ? '0.4' : '1';
                staffSection.style.pointerEvents = locked ? 'none' : '';
            }
        }
        
        // Lock attendance on page load
        document.addEventListener('DOMContentLoaded', function() {
            setAttendanceLocked(true);
        });
        
        async function openPayrollPeriod() {
            const term = document.getElementById('period-term').value;
            const week = document.getElementById('period-week').value;
            const resultDiv = document.getElementById('payroll-period-result');
            
            if (!week) {
                resultDiv.innerHTML = '<span style="color:#c62828;">⚠ Please select a Week first</span>';
                return;
            }
            
            const now = new Date().toLocaleString('en-GB');
            activePayrollPeriod = { term, week, openedAt: now };
            
            // Sync hidden term/week fields in both attendance forms
            document.getElementById('payroll-term').value = term;
            document.getElementById('payroll-week').value = week;
            document.getElementById('support-term').value = term;
            document.getElementById('support-week').value = week;
            
            // Show active period info inside the attendance forms
            const periodLabel = `Term ${term}, Week ${week}`;
            const display = document.getElementById('active-period-display');
            if (display) display.textContent = `Active period: ${periodLabel}`;
            const displayStaff = document.getElementById('active-period-display-staff');
            if (displayStaff) displayStaff.textContent = `Active period: ${periodLabel}`;
            
            document.getElementById('payroll-period-status').textContent =
                `OPEN — Term ${term}, Week ${week} (started ${now})`;
            document.getElementById('payroll-period-status').style.color = '#2e7d32';
            document.getElementById('btn-open-period').disabled = true;
            document.getElementById('btn-close-period').disabled = false;
            // Lock the period selectors while open
            document.getElementById('period-term').disabled = true;
            document.getElementById('period-week').disabled = true;
            resultDiv.innerHTML = `<span style="color:#2e7d32;">✓ Payroll period opened: Term ${term}, Week ${week}. You may now record attendance.</span>`;
            
            // Unlock attendance recording
            setAttendanceLocked(false);
            
            try {
                await fetch('/api/payroll-period/open', {
                    method: 'POST',
                    headers: {'Content-Type':'application/json'},
                    body: JSON.stringify({term: parseInt(term), week: parseInt(week), opened_at: now})
                });
            } catch(e) { /* non-critical */ }
        }
        
        async function closePayrollPeriod() {
            const resultDiv = document.getElementById('payroll-period-result');
            if (!activePayrollPeriod) {
                resultDiv.innerHTML = '<span style="color:#c62828;">⚠ No open payroll period to close</span>';
                return;
            }
            
            if (!confirm(`Close payroll for Term ${activePayrollPeriod.term}, Week ${activePayrollPeriod.week}?\n\nThis will tally totals and prevent further attendance entry for this period.`)) {
                return;
            }
            
            const { term, week } = activePayrollPeriod;
            const now = new Date().toLocaleString('en-GB');
            
            try {
                const resp = await fetch('/api/payroll-period/close', {
                    method: 'POST',
                    headers: {'Content-Type':'application/json'},
                    body: JSON.stringify({term: parseInt(term), week: parseInt(week), closed_at: now})
                });
                const data = await resp.json();
                
                if (data.success) {
                    resultDiv.innerHTML = `<strong style="color:#c62828;">⏹ Payroll CLOSED — Term ${term} Week ${week}</strong><br>
                        Teacher total: KES ${(data.teacher_total||0).toFixed(2)} &nbsp;|&nbsp; 
                        Staff total: KES ${(data.staff_total||0).toFixed(2)} &nbsp;|&nbsp; 
                        <strong>Grand total: KES ${(data.grand_total||0).toFixed(2)}</strong>`;
                } else {
                    resultDiv.innerHTML = `<strong style="color:#c62828;">⏹ Payroll closed: Term ${term}, Week ${week}</strong>`;
                }
            } catch(e) {
                resultDiv.innerHTML = `<strong style="color:#c62828;">⏹ Payroll closed: Term ${term}, Week ${week}</strong>`;
            }
            
            document.getElementById('payroll-period-status').textContent =
                `CLOSED — Term ${term}, Week ${week} (closed ${now})`;
            document.getElementById('payroll-period-status').style.color = '#c62828';
            document.getElementById('btn-open-period').disabled = false;
            document.getElementById('btn-close-period').disabled = true;
            // Re-enable period selectors for next period
            document.getElementById('period-term').disabled = false;
            document.getElementById('period-week').disabled = false;
            document.getElementById('period-week').value = '';
            
            // Clear active period display in forms
            const display = document.getElementById('active-period-display');
            if (display) display.textContent = '';
            const displayStaff = document.getElementById('active-period-display-staff');
            if (displayStaff) displayStaff.textContent = '';
            
            // Lock attendance again
            setAttendanceLocked(true);
            activePayrollPeriod = null;
            
            // Update financial summary — closing payroll affects TEACHER PAYROL and NET BALANCE
            loadFinancialSummary();
        }
        
        async function loadTeachers() {
            try {
                const response = await fetch('/api/teachers');
                const data = await response.json();
                
                if (data.success && data.teachers.length > 0) {
                    // Populate attendance dropdown
                    const select = document.getElementById('payroll-teacher');
                    select.innerHTML = '<option value="">-- Select Teacher --</option>';
                    
                    // Also populate remove-teacher dropdown
                    const removeSelect = document.getElementById('remove-teacher');
                    if (removeSelect) {
                        removeSelect.innerHTML = '<option value="">-- Select Teacher --</option>';
                    }
                    
                    data.teachers.forEach(teacher => {
                        const option = document.createElement('option');
                        option.value = teacher.id;
                        option.textContent = teacher.name;
                        option.dataset.name = teacher.name;
                        select.appendChild(option);
                        
                        if (removeSelect) {
                            const opt2 = document.createElement('option');
                            opt2.value = teacher.id;
                            opt2.textContent = teacher.name;
                            removeSelect.appendChild(opt2);
                        }
                    });
                } else if (data.success && data.teachers.length === 0) {
                    const select = document.getElementById('payroll-teacher');
                    select.innerHTML = '<option value="">-- No teachers found --</option>';
                } else {
                    console.error('Failed to load teachers:', data.message);
                }
            } catch (error) {
                console.error('Error loading teachers:', error);
            }
        }
        
        async function loadFinancialSummary() {
            try {
                const response = await fetch('/api/financial-summary');
                const data = await response.json();
                
                if (data.success) {
                    document.getElementById('student-revenue').textContent = 
                        `KES ${data.summary.student_revenue.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',')}`;
                    document.getElementById('teacher-expenses').textContent = 
                        `KES ${data.summary.teacher_expenses.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',')}`;
                    document.getElementById('net-balance').textContent = 
                        `KES ${data.summary.net_balance.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',')}`;
                }
            } catch (error) {
                console.error('Error loading financial summary:', error);
            }
        }
        
        async function loadPayrollRecords() {
            const tbody = document.getElementById('payroll-tbody');
            tbody.innerHTML = '<tr><td colspan="8" class="loading"><div class="spinner"></div>Loading payroll records...</td></tr>';
            
            try {
                const response = await fetch('/api/payroll-records');
                const data = await response.json();
                
                if (data.success) {
                    tbody.innerHTML = '';
                    
                    if (data.records.length === 0) {
                        tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px; color: #666;">No payroll records found</td></tr>';
                        return;
                    }
                    
                    data.records.forEach(record => {
                        const row = document.createElement('tr');
                        
                        // Format days attended
                        const days = [];
                        if (record.monday === 'Yes') days.push('M');
                        if (record.tuesday === 'Yes') days.push('T');
                        if (record.wednesday === 'Yes') days.push('W');
                        if (record.thursday === 'Yes') days.push('Th');
                        if (record.friday === 'Yes') days.push('F');
                        if (record.saturday === 'Yes') days.push('Sat');
                        
                        row.innerHTML = `
                            <td>T${record.term}</td>
                            <td>W${record.week}</td>
                            <td>${record.teacher_name}</td>
                            <td>${days.join(', ')}</td>
                            <td>KES ${record.weekday_pay.toFixed(2)}</td>
                            <td>KES ${record.weekend_pay.toFixed(2)}</td>
                            <td style="font-weight: 600; color: #f57f17;">KES ${record.total_pay.toFixed(2)}</td>
                            <td>${record.date_recorded}</td>
                        `;
                        
                        tbody.appendChild(row);
                    });
                } else {
                    tbody.innerHTML = `<tr><td colspan="8" class="alert alert-error">${data.message}</td></tr>`;
                }
            } catch (error) {
                tbody.innerHTML = `<tr><td colspan="8" class="alert alert-error">Error loading payroll records: ${error.message}</td></tr>`;
            }
        }
        
        async function recordAttendance() {
            const teacherSelect = document.getElementById('payroll-teacher');
            const teacherId = teacherSelect.value;
            const teacherName = teacherSelect.options[teacherSelect.selectedIndex]?.dataset.name;
            const term = document.getElementById('payroll-term').value;
            const week = document.getElementById('payroll-week').value;
            const resultDiv = document.getElementById('attendance-result');
            
            // Validation
            if (!teacherId) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 15px;">Please select a teacher</div>';
                return;
            }
            
            if (!week) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 15px;">No active payroll period — please open a period first</div>';
                return;
            }
            
            // Get selected days
            const days = {
                monday: document.getElementById('day-monday').checked,
                tuesday: document.getElementById('day-tuesday').checked,
                wednesday: document.getElementById('day-wednesday').checked,
                thursday: document.getElementById('day-thursday').checked,
                friday: document.getElementById('day-friday').checked,
                saturday: document.getElementById('day-saturday').checked,
                saturdaytp: document.getElementById('day-saturdaytp').checked
            };
            
            // Check if at least one day is selected
            if (!Object.values(days).some(d => d)) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 15px;">Please select at least one day</div>';
                return;
            }
            
            resultDiv.innerHTML = '<div class="loading" style="margin-top: 15px;"><div class="spinner"></div>Recording attendance...</div>';
            
            try {
                const response = await fetch('/api/record-attendance', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        teacher_id: parseInt(teacherId),
                        teacher_name: teacherName,
                        term: parseInt(term),
                        week: parseInt(week),
                        days: days
                    })
                });
                
                const data = await response.json();
                
                if (data.success) {
                    resultDiv.innerHTML = `
                        <div class="alert alert-success" style="margin-top: 15px;">
                            <strong>✓ ${data.message}</strong><br>
                            Weekday Pay: KES ${data.weekday_pay.toFixed(2)}<br>
                            Weekend Pay: KES ${data.weekend_pay.toFixed(2)}<br>
                            <strong>Total Pay: KES ${data.total_pay.toFixed(2)}</strong>
                        </div>
                    `;
                    
                    // Clear form — only reset teacher and day selections, NOT the hidden
                    // payroll-term/payroll-week fields which stay locked to the active period
                    document.getElementById('payroll-teacher').value = '';
                    document.getElementById('day-monday').checked = false;
                    document.getElementById('day-tuesday').checked = false;
                    document.getElementById('day-wednesday').checked = false;
                    document.getElementById('day-thursday').checked = false;
                    document.getElementById('day-friday').checked = false;
                    document.getElementById('day-saturday').checked = false;
                    document.getElementById('day-saturdaytp').checked = false;
                    
                    // Reload data
                    loadPayrollRecords();
                    loadFinancialSummary();
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 15px;">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 15px;">Error: ${error.message}</div>`;
            }
        }
        
        // Support Staff Attendance Function
        async function recordSupportStaffAttendance() {
            const staffName = document.getElementById('support-staff-name').value;
            const term = document.getElementById('support-term').value;
            const week = document.getElementById('support-week').value;
            const resultDiv = document.getElementById('support-attendance-result');
            
            if (!staffName) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 15px;">Please select a staff member</div>';
                return;
            }
            if (!week) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 15px;">No active payroll period — please open a period first</div>';
                return;
            }
            
            const days = {
                monday: document.getElementById('support-monday').checked,
                tuesday: document.getElementById('support-tuesday').checked,
                wednesday: document.getElementById('support-wednesday').checked,
                thursday: document.getElementById('support-thursday').checked,
                friday: document.getElementById('support-friday').checked,
                saturday: document.getElementById('support-saturday').checked
            };
            
            if (!Object.values(days).some(d => d)) {
                resultDiv.innerHTML = '<div class="alert alert-error" style="margin-top: 15px;">Please select at least one day</div>';
                return;
            }
            
            const weekdayCount = ['monday','tuesday','wednesday','thursday','friday'].filter(d => days[d]).length;
            const weekdayPay = weekdayCount * 200;
            const weekendPay = days.saturday ? 200 : 0;  // Max 200 KES for weekend duty
            const totalPay = weekdayPay + weekendPay;
            
            resultDiv.innerHTML = '<div class="loading" style="margin-top: 15px;"><div class="spinner"></div>Recording attendance...</div>';
            
            try {
                const response = await fetch('/api/record-support-staff-attendance', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        staff_name: staffName,
                        term: parseInt(term),
                        week: parseInt(week),
                        days: days,
                        weekday_pay: weekdayPay,
                        weekend_pay: weekendPay,
                        total_pay: totalPay
                    })
                });
                
                const data = await response.json();
                
                if (data.success) {
                    resultDiv.innerHTML = `
                        <div class="alert alert-success" style="margin-top: 15px;">
                            <strong>✓ Attendance recorded for ${staffName}</strong><br>
                            Weekday Pay: KES ${weekdayPay.toFixed(2)}<br>
                            Weekend Duty Pay: KES ${weekendPay.toFixed(2)}<br>
                            <strong>Total Pay: KES ${totalPay.toFixed(2)}</strong>
                        </div>
                    `;
                    document.getElementById('support-staff-name').value = '';
                    // Don't reset support-term/support-week — they stay synced to the active period
                    ['support-monday','support-tuesday','support-wednesday','support-thursday','support-friday','support-saturday'].forEach(id => {
                        document.getElementById(id).checked = false;
                    });
                    loadPayrollRecords();
                    loadFinancialSummary();
                } else {
                    resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 15px;">❌ ${data.message}</div>`;
                }
            } catch (error) {
                resultDiv.innerHTML = `<div class="alert alert-error" style="margin-top: 15px;">Error: ${error.message}</div>`;
            }
        }
        
        // Load teacher payroll data when tab is shown
        const originalShowTab = showTab;
        showTab = function(tabName) {
            originalShowTab.call(this, tabName);
            
            if (tabName === 'payroll') {
                loadTeachers();
                loadFinancialSummary();
                loadPayrollRecords();
            }
        };
    </script>
    
    <!-- Edit Payment Modal -->
    <div id="editPaymentModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">Edit Payment</div>
            <div class="form-group">
                <label for="edit-amount">Payment Amount (KES):</label>
                <input type="number" id="edit-amount" step="0.01" min="0" style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 6px;">
            </div>
            <div class="form-group">
                <label for="edit-date">Payment Date:</label>
                <input type="date" id="edit-date" style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 6px;">
            </div>
            <div class="modal-buttons">
                <button class="btn-cancel" onclick="closeEditModal()">Cancel</button>
                <button class="btn btn-success" onclick="saveEditPayment()">Save Changes</button>
            </div>
            <div id="edit-result" style="margin-top: 15px;"></div>
        </div>
    </div>
    
    <!-- Delete Payment Modal -->
    <div id="deletePaymentModal" class="modal">
        <div class="modal-content">
            <div class="modal-header" style="color: #dc3545;">⚠️ Confirm Deletion</div>
            <p id="delete-message" style="margin: 20px 0; font-size: 15px;"></p>
            <p style="color: #666; font-size: 14px;">This action cannot be undone.</p>
            <div class="modal-buttons">
                <button class="btn-cancel" onclick="closeDeleteModal()">Cancel</button>
                <button class="btn" style="background: #dc3545;" onclick="confirmDeletePayment()">Delete Payment</button>
            </div>
            <div id="delete-result" style="margin-top: 15px;"></div>
        </div>
    </div>
    
    <!-- Receipt Container (Hidden, shown only when printing) -->
    <div id="receipt" class="receipt">
        <div class="receipt-header">
            <img src="data:image/png;base64,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" alt="Kirembe Secondary School Logo" style="width: 100px; height: 100px; object-fit: contain; margin: 0 auto 15px; display: block;">
            <h1>KIREMBE SECONDARY SCHOOL</h1>
            <p>P.O. Box [Address] | Tel: [Phone] | Email: [Email]</p>
        </div>
        
        <div class="receipt-title">PAYMENT RECEIPT</div>
        
        <div class="receipt-info">
            <div class="receipt-row">
                <span class="receipt-label">Receipt No:</span>
                <span class="receipt-value" id="receipt-number"></span>
            </div>
            <div class="receipt-row">
                <span class="receipt-label">Date:</span>
                <span class="receipt-value" id="receipt-date"></span>
            </div>
            <div class="receipt-row">
                <span class="receipt-label">Student Name:</span>
                <span class="receipt-value" id="receipt-student-name"></span>
            </div>
            <div class="receipt-row">
                <span class="receipt-label">Admission Number:</span>
                <span class="receipt-value" id="receipt-admission"></span>
            </div>
            <div class="receipt-row">
                <span class="receipt-label">Class/Grade:</span>
                <span class="receipt-value" id="receipt-class"></span>
            </div>
        </div>
        
        <div class="receipt-amount">
            <div class="receipt-amount-label">AMOUNT PAID</div>
            <div class="receipt-amount-value" id="receipt-amount"></div>
        </div>
        
        <div class="receipt-info">
            <div class="receipt-row">
                <span class="receipt-label">Previous Balance:</span>
                <span class="receipt-value" id="receipt-prev-balance"></span>
            </div>
            <div class="receipt-row">
                <span class="receipt-label">Current Balance:</span>
                <span class="receipt-value" id="receipt-current-balance"></span>
            </div>
        </div>
        
        <div class="receipt-footer">
            <p>Thank you for your payment</p>
            <p>This is an official receipt from Kirembe Secondary School</p>
            <p>Printed on: <span id="receipt-print-date"></span></p>
        </div>
    </div>
</body>
</html>
"""


def initialize_services():
    """Initialize the learner and payment services."""
    global learner_service, payment_service, payment_correction_service
    global teacher_manager, attendance_tracker, payroll_calculator, financial_dashboard
    
    if not os.path.exists(SPREADSHEET_FILE):
        print(f"ERROR: Cannot find {SPREADSHEET_FILE}")
        sys.exit(1)
    
    try:
        repository = SpreadsheetRepository(SPREADSHEET_FILE)
        learner_service = LearnerService(repository)
        payment_service = PaymentService(repository)
        payment_correction_service = PaymentCorrectionService(repository)
        
        result = learner_service.initialize_system()
        
        if not result['success']:
            print(f"ERROR: {result['message']}")
            sys.exit(1)
        
        print(f"✓ Successfully loaded {result['learner_count']} learners")
        
        # Initialize teacher payroll services
        if os.path.exists(TEACHER_PAYROLL_FILE):
            teacher_manager = TeacherManager(TEACHER_PAYROLL_FILE)
            attendance_tracker = AttendanceTracker(TEACHER_PAYROLL_FILE)
            payroll_calculator = PayrollCalculator(TEACHER_PAYROLL_FILE)
            financial_dashboard = FinancialDashboard(SPREADSHEET_FILE, TEACHER_PAYROLL_FILE)
            print(f"✓ Teacher payroll system initialized")
        else:
            print(f"⚠ Teacher payroll file not found - payroll features disabled")
        
        return True
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        sys.exit(1)


@app.route('/')
def index():
    """Serve the main page."""
    from flask import make_response
    response = make_response(render_template_string(HTML_TEMPLATE))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response


@app.route('/logo')
def serve_logo():
    """Serve the Kirembe logo."""
    from flask import send_file
    logo_path = 'kirembe_logo.png'
    if os.path.exists(logo_path):
        return send_file(logo_path, mimetype='image/png')
    return '', 404


@app.route('/edit-learner')
def edit_learner_page():
    """Serve the edit learner page."""
    from flask import render_template
    return render_template('edit_learner.html')


@app.route('/api/reload')
def reload_data():
    """API endpoint to reload data from Excel file."""
    global learner_service
    try:
        result = learner_service.initialize_system()
        return jsonify(result)
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error reloading data: {str(e)}'
        }), 500


@app.route('/api/learners')
def get_all_learners():
    """API endpoint to get all learners."""
    try:
        learners = learner_service.get_all_learners()
        
        learners_data = []
        for learner in learners:
            learners_data.append({
                'name': learner.name,
                'admission_number': learner.admission_number,
                'total_owed': float(learner.total_owed),
                'total_paid': float(learner.calculate_total_paid()),
                'balance': float(learner.calculate_balance()),
                'payment_count': len(learner.payment_records)
            })
        
        return jsonify({
            'success': True,
            'learners': learners_data
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        }), 500


@app.route('/api/class-statistics')
def get_class_statistics():
    """API endpoint to get live payment statistics by class stream — computed from in-memory cache."""
    try:
        if learner_service is None:
            return jsonify({'success': False, 'message': 'System not initialised'}), 503

        # Use the already-loaded in-memory learner cache — zero file I/O
        all_learners = list(learner_service._learners_cache.values())

        # Stream config: (match_key, display_name, group_key)
        # match_key is matched against learner.source_sheet (case-insensitive substring)
        stream_config = [
            ('GRADE 10 2026 - Achievers', 'Grade 10A 2026',  'Grade 10 Achievers'),
            ('GRADE 10 2026 - Champions', 'Grade 10C 2026',  'Grade 10 Champions'),
            ('FORM 2 2025 - Achievers',   'Form 2A 2025',    'Form 3A Cohort'),
            ('FORM 3 2026 - Achievers',   'Form 3A 2026',    'Form 3A Cohort'),
            ('FORM 2 2025 - Champions',   'Form 2C 2025',    'Form 3C Cohort'),
            ('FORM 3 2026 - Champions',   'Form 3C 2026',    'Form 3C Cohort'),
            ('FORM 3 2025 - Achievers',   'Form 3A 2025',    'Form 4A Cohort'),
            ('FORM 4 2026 - Achievers',   'Form 4A 2026',    'Form 4A Cohort'),
            ('FORM 3 2025 - Champions',   'Form 3C 2025',    'Form 4C Cohort'),
            ('FORM 4 2026 - Champions',   'Form 4C 2026',    'Form 4C Cohort'),
        ]

        # Fallback: base sheet name only (no stream suffix stored)
        base_fallback = {
            'GRADE 10 2026': 'Grade 10A 2026',
            'FORM 2 2025':   'Form 2A 2025',
            'FORM 3 2026':   'Form 3A 2026',
            'FORM 3 2025':   'Form 3A 2025',
            'FORM 4 2026':   'Form 4A 2026',
        }

        # Build totals dict keyed by display_name
        totals = {
            cfg[1]: {'stream_name': cfg[1], 'group_key': cfg[2],
                     'total_collected': 0.0, 'total_owed': 0.0, 'learner_count': 0}
            for cfg in stream_config
        }
        display_by_match = {cfg[0].upper(): cfg[1] for cfg in stream_config}

        for learner in all_learners:
            source = (getattr(learner, 'source_sheet', '') or '').upper()
            total_paid = sum(
                float(p.amount) for p in learner.payment_records
                if p.amount and float(p.amount) > 0
            )
            total_owed = float(getattr(learner, 'total_owed', 0) or 0)

            # Match by stream key
            matched_display = None
            for match_key_upper, display in display_by_match.items():
                if match_key_upper in source:
                    matched_display = display
                    break

            # Fallback: base sheet name
            if not matched_display:
                for base, display in base_fallback.items():
                    if base.upper() in source:
                        matched_display = display
                        break

            if matched_display and matched_display in totals:
                totals[matched_display]['total_collected'] += total_paid
                totals[matched_display]['total_owed']      += total_owed
                totals[matched_display]['learner_count']   += 1

        statistics = [
            {
                'class_name':      v['stream_name'],
                'group_key':       v['group_key'],
                'total_collected': round(v['total_collected'], 2),
                'total_owed':      round(v['total_owed'], 2),
                'learner_count':   v['learner_count'],
            }
            for v in totals.values()
            if v['learner_count'] > 0
        ]

        # Sort: group cohorts together, then by year within group
        statistics.sort(key=lambda x: (x['group_key'], x['class_name']))

        return jsonify({'success': True, 'statistics': statistics})

    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/learner/<admission_number>')
def get_learner(admission_number):
    """API endpoint to get a specific learner."""
    try:
        result = learner_service.get_learner_details(admission_number)
        
        if result['success']:
            details = result['details']
            
            # Load actual payment dates from PAYMENT_DATES_LOG.xlsx
            # This replaces the default datetime.now() timestamps with real recorded dates
            import openpyxl as _opxl
            date_log = {}  # key: (adm, str(amount)) -> list of timestamp strings in order
            log_file = 'PAYMENT_DATES_LOG.xlsx'
            try:
                if os.path.exists(log_file):
                    _wb = _opxl.load_workbook(log_file, read_only=True)
                    _ws = _wb.active
                    for _row in _ws.iter_rows(min_row=2, values_only=True):
                        if _row[1] and _row[3] and _row[6]:
                            _key = (str(_row[1]).strip(), str(float(_row[3])))
                            if _key not in date_log:
                                date_log[_key] = []
                            date_log[_key].append(str(_row[6]))
                    _wb.close()
            except Exception:
                pass  # Fall back to default timestamps if log unavailable
            
            # Track how many times each (adm, amount) key has been used to handle duplicates
            date_log_pos = {}
            
            # Format payment history using real dates from log
            payment_history = []
            for payment in details['payment_history']:
                _key = (str(admission_number).strip(), str(float(payment.amount)))
                _pos = date_log_pos.get(_key, 0)
                _dates = date_log.get(_key, [])
                if _pos < len(_dates):
                    ts_str = _dates[_pos]
                    date_log_pos[_key] = _pos + 1
                else:
                    # Fallback: use whatever timestamp was parsed (may be today's date if not in log)
                    ts_str = payment.timestamp.strftime('%Y-%m-%d %H:%M:%S') if hasattr(payment.timestamp, 'strftime') else str(payment.timestamp)
                payment_history.append({
                    'amount': float(payment.amount),
                    'timestamp': ts_str
                })
            
            return jsonify({
                'success': True,
                'details': {
                    'name': details['name'],
                    'admission_number': details['admission_number'],
                    'source_sheet': details.get('source_sheet', 'Unknown'),
                    'total_owed': float(details['total_owed']),
                    'total_paid': float(details['total_paid']),
                    'balance': float(details['balance']),
                    'payment_history': payment_history
                }
            })
        else:
            return jsonify({
                'success': False,
                'message': result['message']
            }), 404
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        }), 500


@app.route('/api/record-payment', methods=['POST'])
def record_payment():
    """API endpoint to record a payment."""
    try:
        data = request.get_json()
        admission_number = data.get('admission_number', '').strip()
        amount_str = data.get('amount', '')
        payment_date_str = data.get('payment_date')
        
        if not admission_number:
            return jsonify({
                'success': False,
                'message': 'Admission number is required'
            }), 400
        
        # Find learner
        learner = learner_service.search_by_admission_number(admission_number)
        
        if learner is None:
            return jsonify({
                'success': False,
                'message': f'No learner found with admission number {admission_number}'
            }), 404
        
        # Parse amount
        try:
            amount = Decimal(str(amount_str))
            
            if      amount <= 0:
                return jsonify({
                    'success': False,
                    'message': 'Payment amount must be greater than zero'
                }), 400
            
            if amount > Decimal('1000000'):
                return jsonify({
                    'success': False,
                    'message': 'Payment amount exceeds maximum allowed (KES 1,000,000)'
                }), 400
                
        except (InvalidOperation, ValueError):
            return jsonify({
                'success': False,
                'message': 'Invalid payment amount'
            }), 400
        
        # Parse payment date if provided
        payment_timestamp = None
        if payment_date_str:
            try:
                # Parse date string (YYYY-MM-DD) and set time to current time
                payment_date = datetime.fromisoformat(payment_date_str)
                # Combine the date with current time
                now = datetime.now()
                payment_timestamp = datetime(
                    payment_date.year,
                    payment_date.month,
                    payment_date.day,
                    now.hour,
                    now.minute,
                    now.second
                )
            except ValueError:
                return jsonify({
                    'success': False,
                    'message': 'Invalid date format. Use YYYY-MM-DD format.'
                }), 400
        
        # Capture previous balance before recording payment
        previous_balance = float(learner.calculate_balance())
        
        # Record payment with optional timestamp
        result = payment_service.record_payment(learner, amount, payment_timestamp)
        
        if result['success']:
            return jsonify({
                'success': True,
                'message': result['message'],
                'new_balance': float(learner.calculate_balance()),
                'previous_balance': previous_balance,
                'learner_name': learner.name,
                'admission_number': learner.admission_number,
                'source_sheet': learner.source_sheet or 'N/A',
                'payment_amount': float(amount),
                'payment_date': (payment_timestamp or datetime.now()).strftime('%Y-%m-%d %H:%M:%S')
            })
        else:
            return jsonify({
                'success': False,
                'message': result['message']
            }), 400
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        }), 500


@app.route('/api/edit-payment', methods=['POST'])
def edit_payment_endpoint():
    """API endpoint to edit a payment record."""
    try:
        data = request.get_json()
        admission_number = data.get('admission_number', '').strip()
        payment_index = data.get('payment_index')
        new_amount = data.get('new_amount')
        new_date_str = data.get('new_date')
        
        if not admission_number:
            return jsonify({
                'success': False,
                'message': 'Admission number is required'
            }), 400
        
        if payment_index is None:
            return jsonify({
                'success': False,
                'message': 'Payment index is required'
            }), 400
        
        # Find learner
        learner = learner_service.search_by_admission_number(admission_number)
        
        if learner is None:
            return jsonify({
                'success': False,
                'message': f'No learner found with admission number {admission_number}'
            }), 404
        
        # Parse new amount if provided
        parsed_amount = None
        if new_amount is not None:
            try:
                parsed_amount = Decimal(str(new_amount))
            except (InvalidOperation, ValueError):
                return jsonify({
                    'success': False,
                    'message': 'Invalid payment amount'
                }), 400
        
        # Parse new date if provided
        parsed_date = None
        if new_date_str:
            try:
                parsed_date = datetime.fromisoformat(new_date_str.replace('Z', '+00:00'))
            except ValueError:
                return jsonify({
                    'success': False,
                    'message': 'Invalid date format. Use ISO format (YYYY-MM-DD)'
                }), 400
        
        # Edit payment
        result = payment_correction_service.edit_payment(
            learner,
            int(payment_index),
            parsed_amount,
            parsed_date
        )
        
        if result['success']:
            return jsonify({
                'success': True,
                'message': result['message'],
                'new_balance': float(learner.calculate_balance())
            })
        else:
            return jsonify({
                'success': False,
                'message': result['message']
            }), 400
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        }), 500


@app.route('/api/delete-payment', methods=['POST'])
def delete_payment_endpoint():
    """API endpoint to delete a payment record."""
    try:
        data = request.get_json()
        admission_number = data.get('admission_number', '').strip()
        payment_index = data.get('payment_index')
        
        if not admission_number:
            return jsonify({
                'success': False,
                'message': 'Admission number is required'
            }), 400
        
        if payment_index is None:
            return jsonify({
                'success': False,
                'message': 'Payment index is required'
            }), 400
        
        # Find learner
        learner = learner_service.search_by_admission_number(admission_number)
        
        if learner is None:
            return jsonify({
                'success': False,
                'message': f'No learner found with admission number {admission_number}'
            }), 404
        
        # Delete payment
        result = payment_correction_service.delete_payment(
            learner,
            int(payment_index)
        )
        
        if result['success']:
            return jsonify({
                'success': True,
                'message': result['message'],
                'new_balance': float(result['new_balance'])
            })
        else:
            return jsonify({
                'success': False,
                'message': result['message']
            }), 400
            
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        }), 500


@app.route('/api/add-learner', methods=['POST'])
def add_learner():
    """API endpoint to add a new learner."""
    try:
        data = request.get_json()
        admission_number = data.get('admission_number', '').strip()
        name = data.get('name', '').strip()
        grade = data.get('grade', '').strip()
        arrears = data.get('arrears', 0)
        debt = data.get('debt', 2200)
        admission_date = data.get('admission_date', datetime.now().strftime('%Y-%m-%d'))
        
        # Validation
        if not admission_number:
            return jsonify({
                'success': False,
                'message': 'Admission number is required'
            }), 400
        
        if not name:
            return jsonify({
                'success': False,
                'message': 'Learner name is required'
            }), 400
        
        if not grade:
            return jsonify({
                'success': False,
                'message': 'Grade/Class is required'
            }), 400
        
        # Check if learner already exists
        existing_learner = learner_service.search_by_admission_number(admission_number)
        if existing_learner:
            return jsonify({
                'success': False,
                'message': f'Learner with admission number {admission_number} already exists'
            }), 400
        
        # Add learner to Excel file
        import openpyxl
        wb = openpyxl.load_workbook(SPREADSHEET_FILE)
        
        # Find the appropriate sheet
        sheet_name = grade
        if sheet_name not in wb.sheetnames:
            return jsonify({
                'success': False,
                'message': f'Grade/Class "{grade}" not found in the system'
            }), 400
        
        ws = wb[sheet_name]
        
        # Find the next empty row
        next_row = ws.max_row + 1
        
        # Determine column structure based on sheet
        if "GRADE 10" in sheet_name.upper():
            # Grade 10 structure: A=NO, B=ADM, C=NAME, D=ARREARS, E=DEBT
            ws.cell(next_row, 1).value = next_row - 5  # Row number
            ws.cell(next_row, 2).value = admission_number
            ws.cell(next_row, 3).value = name
            ws.cell(next_row, 4).value = float(arrears)
            ws.cell(next_row, 5).value = float(debt)
        else:
            # Standard structure: B=ADM, C=NAME, D=ARREARS, E=DEBT
            ws.cell(next_row, 2).value = admission_number
            ws.cell(next_row, 3).value = name
            ws.cell(next_row, 4).value = float(arrears)
            ws.cell(next_row, 5).value = float(debt)
        
        # Save the workbook
        wb.save(SPREADSHEET_FILE)
        
        # Reload services to pick up the new learner
        initialize_services()
        
        return jsonify({
            'success': True,
            'message': f'Learner {name} (Admission: {admission_number}) added successfully to {grade}'
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error adding learner: {str(e)}'
        }), 500


@app.route('/api/learner/update', methods=['POST'])
def update_learner():
    """API endpoint to update learner details."""
    try:
        data = request.get_json()
        admission_number = data.get('admission_number', '').strip()
        new_name = data.get('name', '').strip()
        new_balance = data.get('total_owed')
        
        # Validation
        if not admission_number:
            return jsonify({
                'success': False,
                'message': 'Admission number is required'
            }), 400
        
        if not new_name:
            return jsonify({
                'success': False,
                'message': 'Name cannot be empty'
            }), 400
        
        if new_balance is None or float(new_balance) < 0:
            return jsonify({
                'success': False,
                'message': 'Invalid balance amount'
            }), 400
        
        # Find the learner
        learner = learner_service.search_by_admission_number(admission_number)
        if not learner:
            return jsonify({
                'success': False,
                'message': f'Learner with admission number {admission_number} not found'
            }), 404
        
        # Update learner details
        learner.name = new_name
        learner.total_owed = Decimal(str(new_balance))
        
        # Save to Excel
        result = learner_service.save_learner(learner)
        
        if result['success']:
            # Reload services to pick up changes
            initialize_services()
            
            return jsonify({
                'success': True,
                'message': f'Learner details updated successfully'
            })
        else:
            return jsonify({
                'success': False,
                'message': result['message']
            }), 500
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error updating learner: {str(e)}'
        }), 500


@app.route('/api/exit-learner', methods=['POST'])
def exit_learner():
    """API endpoint to exit a learner (mark as inactive)."""
    try:
        data = request.get_json()
        admission_number = data.get('admission_number', '').strip()
        exit_date = data.get('exit_date', datetime.now().strftime('%Y-%m-%d'))
        reason = data.get('reason', 'Unknown')
        notes = data.get('notes', '').strip()
        
        # Validation
        if not admission_number:
            return jsonify({
                'success': False,
                'message': 'Admission number is required'
            }), 400
        
        # Check if learner exists
        learner = learner_service.search_by_admission_number(admission_number)
        if not learner:
            return jsonify({
                'success': False,
                'message': f'No learner found with admission number {admission_number}'
            }), 404
        
        # Create exited learners log file if it doesn't exist
        import json
        import os
        
        exit_log_file = 'exited_learners.json'
        
        # Load existing exits
        if os.path.exists(exit_log_file):
            with open(exit_log_file, 'r') as f:
                exits = json.load(f)
        else:
            exits = []
        
        # Add new exit record
        exit_record = {
            'admission_number': admission_number,
            'name': learner.name,
            'grade': learner.source_sheet,
            'exit_date': exit_date,
            'reason': reason,
            'notes': notes,
            'final_balance': float(learner.calculate_balance()),
            'total_owed': float(learner.total_owed),
            'total_paid': float(learner.total_paid),
            'payment_history': [
                {
                    'amount': float(p.amount),
                    'timestamp': p.timestamp.isoformat() if hasattr(p.timestamp, 'isoformat') else str(p.timestamp)
                }
                for p in learner.payment_history
            ]
        }
        
        exits.append(exit_record)
        
        # Save exit log
        with open(exit_log_file, 'w') as f:
            json.dump(exits, f, indent=2)
        
        # Mark learner as exited in Excel by adding a note/comment
        # For now, we'll just log it - you can extend this to modify the Excel file
        
        return jsonify({
            'success': True,
            'message': f'Learner {learner.name} (Admission: {admission_number}) has been exited. All records preserved in exit log.'
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error exiting learner: {str(e)}'
        }), 500


@app.route('/api/promote-students', methods=['POST'])
def promote_students():
    """API endpoint to promote students to the next grade."""
    try:
        data = request.get_json()
        year = data.get('year')
        promote_grade10 = data.get('promote_grade10', False)
        promote_form3 = data.get('promote_form3', False)
        promote_form4 = data.get('promote_form4', False)
        preview_only = data.get('preview_only', False)
        
        if not year:
            return jsonify({
                'success': False,
                'message': 'Academic year is required'
            }), 400
        
        import openpyxl
        from datetime import datetime
        import shutil
        
        # Create backup
        backup_file = f"REMEDIAL_PAYMENT_BALANCES_BACKUP_PROMOTION_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        if not preview_only:
            shutil.copy2(SPREADSHEET_FILE, backup_file)
        
        # Load workbook
        wb = openpyxl.load_workbook(SPREADSHEET_FILE)
        
        preview = {}
        results = {}
        
        # Promote Grade 10 to Form 3
        if promote_grade10:
            grade10_sheet = "GRADE 10 2026"
            form3_new_sheet = f"FORM 3 {year}"
            
            if grade10_sheet in wb.sheetnames:
                ws_source = wb[grade10_sheet]
                
                # Calculate preview data
                students = []
                total_arrears = 0
                
                for row_idx in range(6, ws_source.max_row + 1):
                    adm = ws_source.cell(row_idx, 2).value
                    name = ws_source.cell(row_idx, 3).value
                    arrears = ws_source.cell(row_idx, 4).value or 0
                    debt = ws_source.cell(row_idx, 5).value or 0
                    
                    if not adm or not name:
                        continue
                    
                    # Skip total rows
                    if isinstance(debt, str) and '=' in str(debt):
                        continue
                    
                    # Find learner to get current balance
                    learner = learner_service.search_by_admission_number(str(adm))
                    if learner:
                        current_balance = float(learner.calculate_balance())
                        students.append({
                            'admission': adm,
                            'name': name,
                            'current_balance': current_balance,
                            'new_arrears': current_balance,
                            'new_debt': 2200,
                            'new_total': current_balance + 2200
                        })
                        total_arrears += current_balance
                
                preview['grade10'] = {
                    'count': len(students),
                    'total_arrears': total_arrears,
                    'new_fees': len(students) * 2200,
                    'total_new_owed': total_arrears + (len(students) * 2200),
                    'students': students
                }
                
                # If not preview, create new sheet and move students
                if not preview_only:
                    # Create new Form 3 sheet or use existing
                    if form3_new_sheet not in wb.sheetnames:
                        ws_new = wb.create_sheet(form3_new_sheet)
                        # Copy headers from existing Form 3 sheet
                        if "FORM 3 2025" in wb.sheetnames:
                            ws_template = wb["FORM 3 2025"]
                            for row_idx in range(1, 6):
                                for col_idx in range(1, 20):
                                    ws_new.cell(row_idx, col_idx).value = ws_template.cell(row_idx, col_idx).value
                    else:
                        ws_new = wb[form3_new_sheet]
                    
                    # Add students to new sheet
                    next_row = ws_new.max_row + 1
                    promoted_count = 0
                    
                    for student in students:
                        ws_new.cell(next_row, 2).value = student['admission']
                        ws_new.cell(next_row, 3).value = student['name']
                        ws_new.cell(next_row, 4).value = student['new_arrears']
                        ws_new.cell(next_row, 5).value = student['new_debt']
                        next_row += 1
                        promoted_count += 1
                    
                    results['grade10'] = {'promoted': promoted_count}
        
        # Promote Form 3 to Form 4
        if promote_form3:
            form3_sheet = "FORM 3 2026"
            form4_new_sheet = f"FORM 4 {year}"
            
            if form3_sheet in wb.sheetnames:
                ws_source = wb[form3_sheet]
                
                # Calculate preview data
                students = []
                total_arrears = 0
                
                for row_idx in range(2, ws_source.max_row + 1):
                    adm = ws_source.cell(row_idx, 2).value
                    name = ws_source.cell(row_idx, 3).value
                    arrears = ws_source.cell(row_idx, 4).value or 0
                    debt = ws_source.cell(row_idx, 5).value or 0
                    
                    if not adm or not name:
                        continue
                    
                    # Skip total rows
                    if isinstance(debt, str) and '=' in str(debt):
                        continue
                    
                    # Find learner to get current balance
                    learner = learner_service.search_by_admission_number(str(adm))
                    if learner:
                        current_balance = float(learner.calculate_balance())
                        students.append({
                            'admission': adm,
                            'name': name,
                            'current_balance': current_balance,
                            'new_arrears': current_balance,
                            'new_debt': 2200,
                            'new_total': current_balance + 2200
                        })
                        total_arrears += current_balance
                
                preview['form3'] = {
                    'count': len(students),
                    'total_arrears': total_arrears,
                    'new_fees': len(students) * 2200,
                    'total_new_owed': total_arrears + (len(students) * 2200),
                    'students': students
                }
                
                # If not preview, create new sheet and move students
                if not preview_only:
                    # Create new Form 4 sheet or use existing
                    if form4_new_sheet not in wb.sheetnames:
                        ws_new = wb.create_sheet(form4_new_sheet)
                        # Copy headers from existing Form 4 sheet
                        if "FORM 4 2026" in wb.sheetnames:
                            ws_template = wb["FORM 4 2026"]
                            for row_idx in range(1, 6):
                                for col_idx in range(1, 20):
                                    ws_new.cell(row_idx, col_idx).value = ws_template.cell(row_idx, col_idx).value
                    else:
                        ws_new = wb[form4_new_sheet]
                    
                    # Add students to new sheet
                    next_row = ws_new.max_row + 1
                    promoted_count = 0
                    
                    for student in students:
                        ws_new.cell(next_row, 4).value = student['admission']
                        ws_new.cell(next_row, 5).value = student['name']
                        ws_new.cell(next_row, 6).value = student['new_arrears']
                        ws_new.cell(next_row, 7).value = student['new_debt']
                        next_row += 1
                        promoted_count += 1
                    
                    results['form3'] = {'promoted': promoted_count}
        
        # Graduate Form 4
        if promote_form4:
            form4_sheet = "FORM 4 2026"
            
            if form4_sheet in wb.sheetnames:
                ws_source = wb[form4_sheet]
                
                # Calculate preview data
                students = []
                total_final_balance = 0
                
                for row_idx in range(6, ws_source.max_row + 1):
                    adm = ws_source.cell(row_idx, 4).value
                    name = ws_source.cell(row_idx, 5).value
                    
                    if not adm or not name:
                        continue
                    
                    # Find learner to get current balance
                    learner = learner_service.search_by_admission_number(str(adm))
                    if learner:
                        current_balance = float(learner.calculate_balance())
                        students.append({
                            'admission': adm,
                            'name': name,
                            'final_balance': current_balance
                        })
                        total_final_balance += current_balance
                
                preview['form4'] = {
                    'count': len(students),
                    'total_final_balance': total_final_balance,
                    'students': students
                }
                
                # If not preview, mark as graduated (exit them)
                if not preview_only:
                    import json
                    import os
                    
                    exit_log_file = 'exited_learners.json'
                    
                    if os.path.exists(exit_log_file):
                        with open(exit_log_file, 'r') as f:
                            exits = json.load(f)
                    else:
                        exits = []
                    
                    graduated_count = 0
                    for student in students:
                        learner = learner_service.search_by_admission_number(str(student['admission']))
                        if learner:
                            exit_record = {
                                'admission_number': student['admission'],
                                'name': student['name'],
                                'grade': 'FORM 4 2026',
                                'exit_date': datetime.now().strftime('%Y-%m-%d'),
                                'reason': 'Graduated',
                                'notes': f'Graduated in {year}',
                                'final_balance': student['final_balance'],
                                'total_owed': float(learner.total_owed),
                                'total_paid': float(learner.total_paid),
                                'payment_history': [
                                    {
                                        'amount': float(p.amount),
                                        'timestamp': p.timestamp.isoformat() if hasattr(p.timestamp, 'isoformat') else str(p.timestamp)
                                    }
                                    for p in learner.payment_history
                                ]
                            }
                            exits.append(exit_record)
                            graduated_count += 1
                    
                    with open(exit_log_file, 'w') as f:
                        json.dump(exits, f, indent=2)
                    
                    results['form4'] = {'graduated': graduated_count}
        
        # Save workbook if not preview
        if not preview_only:
            wb.save(SPREADSHEET_FILE)
            
            return jsonify({
                'success': True,
                'message': f'Students promoted successfully to {year}',
                'backup_file': backup_file,
                'results': results
            })
        else:
            return jsonify({
                'success': True,
                'preview': preview
            })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error promoting students: {str(e)}'
        }), 500


@app.route('/api/teachers')
def get_teachers():
    """API endpoint to get all teachers."""
    try:
        if teacher_manager is None:
            return jsonify({
                'success': False,
                'message': 'Teacher payroll system not initialized'
            }), 503
        
        teachers = teacher_manager.get_all_teachers()
        
        return jsonify({
            'success': True,
            'teachers': teachers
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error loading teachers: {str(e)}'
        }), 500


@app.route('/api/record-attendance', methods=['POST'])
def record_attendance_endpoint():
    """API endpoint to record teacher attendance."""
    try:
        if attendance_tracker is None or payroll_calculator is None:
            return jsonify({
                'success': False,
                'message': 'Teacher payroll system not initialized'
            }), 503
        
        data = request.get_json()
        teacher_id = data.get('teacher_id')
        teacher_name = data.get('teacher_name', '')
        term = data.get('term')
        week = data.get('week')
        days = data.get('days', {})
        
        if not all([teacher_id, term, week]):
            return jsonify({
                'success': False,
                'message': 'Teacher ID, term, and week are required'
            }), 400
        
        # Record attendance (this returns the calculated pay)
        result = attendance_tracker.record_attendance(
            teacher_id=int(teacher_id),
            teacher_name=teacher_name,
            term=int(term),
            week=int(week),
            days=days
        )
        
        if result['success']:
            return jsonify({
                'success': True,
                'message': result['message'],
                'weekday_pay': result['weekday_pay'],
                'weekend_pay': result['weekend_pay'],
                'total_pay': result['total_pay']
            })
        else:
            return jsonify({
                'success': False,
                'message': result.get('message', 'Failed to record attendance')
            }), 400

    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error recording attendance: {str(e)}'
        }), 500


@app.route('/api/record-support-staff-attendance', methods=['POST'])
def record_support_staff_attendance():
    """API endpoint to record support staff attendance."""
    try:
        import openpyxl
        from datetime import datetime as dt
        import os
        
        data = request.get_json()
        staff_name = data.get('staff_name', '').strip()
        term = data.get('term')
        week = data.get('week')
        days = data.get('days', {})
        weekday_pay = data.get('weekday_pay', 0)
        weekend_pay = data.get('weekend_pay', 0)
        total_pay = data.get('total_pay', 0)
        
        if not staff_name:
            return jsonify({'success': False, 'message': 'Staff name is required'}), 400
        if not term or not week:
            return jsonify({'success': False, 'message': 'Term and week are required'}), 400
        
        # Enforce 200 KES cap on weekend duty
        if weekend_pay > 200:
            weekend_pay = 200
            total_pay = weekday_pay + weekend_pay
        
        days_attended = [d.capitalize() for d, v in days.items() if v]
        
        # Log to TEACHER_PAYROLL.xlsx under Staff Attendance sheet
        if os.path.exists(TEACHER_PAYROLL_FILE):
            try:
                wb = openpyxl.load_workbook(TEACHER_PAYROLL_FILE)
                sheet_name = 'Staff Attendance'
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                else:
                    ws = wb.create_sheet(sheet_name)
                    ws.append(['Date Recorded', 'Staff Name', 'Term', 'Week',
                               'Days Attended', 'Weekday Pay (KES)', 'Weekend Pay (KES)', 'Total Pay (KES)'])
                
                ws.append([
                    dt.now().strftime('%Y-%m-%d %H:%M:%S'),
                    staff_name,
                    f'Term {term}',
                    f'Week {week}',
                    ', '.join(days_attended),
                    float(weekday_pay),
                    float(weekend_pay),
                    float(total_pay)
                ])
                wb.save(TEACHER_PAYROLL_FILE)
            except Exception as e:
                # Log to a fallback file if payroll file is locked
                pass
        
        return jsonify({
            'success': True,
            'message': f'Attendance recorded for {staff_name} - Term {term}, Week {week}',
            'staff_name': staff_name,
            'weekday_pay': float(weekday_pay),
            'weekend_pay': float(weekend_pay),
            'total_pay': float(total_pay),
            'days_attended': days_attended
        })
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error recording support staff attendance: {str(e)}'}), 500


@app.route('/api/support-staff-records', methods=['GET'])
def get_support_staff_records():
    """Return support staff attendance records filtered by term and week."""
    try:
        import openpyxl as _opxl
        term = request.args.get('term', '')
        week = request.args.get('week', '')
        records = []

        if os.path.exists(TEACHER_PAYROLL_FILE):
            wb = _opxl.load_workbook(TEACHER_PAYROLL_FILE, data_only=True)
            sheet_name = 'Staff Attendance'
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                headers = None
                for row in ws.iter_rows(values_only=True):
                    # Capture header row
                    if headers is None:
                        headers = [str(h).strip() if h else '' for h in row]
                        continue
                    if not row or not any(row):
                        continue

                    rec = dict(zip(headers, row))

                    # --- Detect row format ---
                    # NEW format rows (written by this app) use positional columns:
                    #   col0=Date Recorded, col1=Staff Name, col2=Term('Term X'),
                    #   col3=Week('Week Y'), col4=Days, col5=Weekday Pay, col6=Weekend Pay, col7=Total Pay
                    # OLD format rows (pre-existing) use:
                    #   col0=ID, col1=Name, col2=Weekend, col3=Teacher ID, ...
                    #   — these have no Term/Week and should be skipped.

                    # Try header-based lookup first (works if sheet was recreated with new headers)
                    r_term_raw = str(rec.get('Term', rec.get('term', '')) or '').strip()
                    r_week_raw = str(rec.get('Week', rec.get('week', '')) or '').strip()

                    # If headers don't have Term/Week, fall back to positional detection:
                    # A new-format row has a date-like string in col0 and 'Term X' pattern in col2
                    if (not r_term_raw or not r_week_raw) and len(row) >= 4:
                        col2 = str(row[2] or '').strip()
                        col3 = str(row[3] or '').strip()
                        if col2.lower().startswith('term ') or col3.lower().startswith('week '):
                            r_term_raw = col2
                            r_week_raw = col3

                    # Normalise: strip 'Term '/'Week ' prefix
                    r_term = r_term_raw.replace('Term ', '').replace('term ', '').strip()
                    r_week = r_week_raw.replace('Week ', '').replace('week ', '').strip()

                    # Skip rows with no recognisable term or week (old-format / phantom rows)
                    if not r_term or not r_week:
                        continue
                    # Must be numeric
                    if not r_term.isdigit() or not r_week.isdigit():
                        continue

                    # Filter by requested term and week
                    if term and r_term != str(term):
                        continue
                    if week and r_week != str(week):
                        continue

                    # Extract values — prefer header-mapped, fall back to positional
                    staff_name = (rec.get('Staff Name') or rec.get('Name') or
                                  (row[1] if len(row) > 1 else '')) or ''
                    days_attended = (rec.get('Days Attended') or
                                     (row[4] if len(row) > 4 else '')) or ''
                    try:
                        weekday_pay = float(rec.get('Weekday Pay (KES)', rec.get('Weekday Pay',
                                            row[5] if len(row) > 5 else 0)) or 0)
                    except (TypeError, ValueError):
                        weekday_pay = 0.0
                    try:
                        weekend_pay = float(rec.get('Weekend Pay (KES)', rec.get('Weekend Pay',
                                            row[6] if len(row) > 6 else 0)) or 0)
                    except (TypeError, ValueError):
                        weekend_pay = 0.0
                    try:
                        total_pay = float(rec.get('Total Pay (KES)', rec.get('Total Pay',
                                          row[7] if len(row) > 7 else 0)) or 0)
                    except (TypeError, ValueError):
                        total_pay = weekday_pay + weekend_pay

                    date_recorded = str(rec.get('Date Recorded', row[0] if len(row) > 0 else '') or '')

                    records.append({
                        'staff_name': str(staff_name),
                        'term': r_term,
                        'week': r_week,
                        'days_attended': str(days_attended),
                        'weekday_pay': weekday_pay,
                        'weekend_pay': weekend_pay,
                        'total_pay': total_pay,
                        'date_recorded': date_recorded,
                    })
            wb.close()

        return jsonify({'success': True, 'records': records})

    except Exception as e:
        return jsonify({'success': False, 'records': [], 'message': str(e)}), 500


@app.route('/api/payroll-period/open', methods=['POST'])
def open_payroll_period():
    """Mark a payroll period as open."""
    try:
        import json as _json
        data = request.get_json()
        term = data.get('term')
        week = data.get('week')
        opened_at = data.get('opened_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

        period_file = 'payroll_periods.json'
        periods = []
        if os.path.exists(period_file):
            with open(period_file, 'r') as f:
                periods = _json.load(f)

        periods.append({'term': term, 'week': week, 'opened_at': opened_at, 'closed_at': None, 'status': 'open'})
        with open(period_file, 'w') as f:
            _json.dump(periods, f, indent=2)

        return jsonify({'success': True, 'message': f'Payroll period opened: Term {term}, Week {week}'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/payroll-period/close', methods=['POST'])
def close_payroll_period():
    """Close a payroll period and tally totals."""
    try:
        import json as _json
        import openpyxl as _opxl
        data = request.get_json()
        term = data.get('term')
        week = data.get('week')
        closed_at = data.get('closed_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

        teacher_total = 0.0
        staff_total = 0.0

        if os.path.exists(TEACHER_PAYROLL_FILE):
            wb = _opxl.load_workbook(TEACHER_PAYROLL_FILE, data_only=True)
            if 'Attendance' in wb.sheetnames:
                ws = wb['Attendance']
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[1] and str(row[1]) == str(term) and row[2] and str(row[2]) == str(week):
                        teacher_total += float(row[8] or 0)
            if 'Staff Attendance' in wb.sheetnames:
                ws2 = wb['Staff Attendance']
                headers = None
                for row in ws2.iter_rows(min_row=1, values_only=True):
                    if headers is None:
                        headers = list(row)
                        continue
                    if not row or not any(row):
                        continue
                    try:
                        rec = dict(zip(headers, row))
                        # Normalise term/week — handle both header-based and positional
                        r_term_raw = str(rec.get('Term', rec.get('term', '')) or '').strip()
                        r_week_raw = str(rec.get('Week', rec.get('week', '')) or '').strip()

                        # Positional fallback for rows appended with new schema
                        if (not r_term_raw or not r_week_raw) and len(row) >= 4:
                            col2 = str(row[2] or '').strip()
                            col3 = str(row[3] or '').strip()
                            if col2.lower().startswith('term ') or col3.lower().startswith('week '):
                                r_term_raw = col2
                                r_week_raw = col3

                        r_term = r_term_raw.replace('Term ', '').replace('term ', '').strip()
                        r_week = r_week_raw.replace('Week ', '').replace('week ', '').strip()

                        if not r_term or not r_week or not r_term.isdigit() or not r_week.isdigit():
                            continue
                        if r_term != str(term) or r_week != str(week):
                            continue

                        # Get total pay — header-based then positional
                        try:
                            pay = float(rec.get('Total Pay (KES)', rec.get('Total Pay',
                                        row[7] if len(row) > 7 else 0)) or 0)
                        except (TypeError, ValueError):
                            pay = 0.0
                        staff_total += pay
                    except (ValueError, IndexError, TypeError):
                        pass
            wb.close()

        period_file = 'payroll_periods.json'
        periods = []
        if os.path.exists(period_file):
            with open(period_file, 'r') as f:
                periods = _json.load(f)

        for p in periods:
            if p['term'] == term and p['week'] == week and p['status'] == 'open':
                p['closed_at'] = closed_at
                p['status'] = 'closed'
                p['teacher_total'] = teacher_total
                p['staff_total'] = staff_total
                p['grand_total'] = teacher_total + staff_total

        with open(period_file, 'w') as f:
            _json.dump(periods, f, indent=2)

        return jsonify({
            'success': True,
            'message': f'Payroll closed: Term {term}, Week {week}',
            'teacher_total': teacher_total,
            'staff_total': staff_total,
            'grand_total': teacher_total + staff_total
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/delete-attendance', methods=['POST'])
def delete_attendance_endpoint():
    """API endpoint to delete an attendance record."""
    try:
        if attendance_tracker is None:
            return jsonify({
                'success': False,
                'message': 'Teacher payroll system not initialized'
            }), 503
        
        data = request.get_json()
        record_id = data.get('record_id')
        
        if not record_id:
            return jsonify({
                'success': False,
                'message': 'Record ID is required'
            }), 400
        
        # Delete the attendance record
        result = attendance_tracker.delete_attendance_record(record_id)
        
        if result['success']:
            return jsonify({
                'success': True,
                'message': result['message']
            })
        else:
            return jsonify({
                'success': False,
                'message': result['message']
            }), 400
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error deleting attendance: {str(e)}'
        }), 500


@app.route('/api/payroll-records')
def get_payroll_records():
    """API endpoint to get payroll records."""
    try:
        if attendance_tracker is None:
            return jsonify({
                'success': False,
                'message': 'Teacher payroll system not initialized'
            }), 503
        
        term = request.args.get('term', type=int)
        
        records = attendance_tracker.get_all_attendance(term=term)
        
        return jsonify({
            'success': True,
            'records': records
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error loading payroll records: {str(e)}'
        }), 500


@app.route('/api/financial-summary')
def get_financial_summary():
    """API endpoint to get financial summary."""
    try:
        if financial_dashboard is None:
            return jsonify({
                'success': False,
                'message': 'Teacher payroll system not initialized'
            }), 503
        
        # Get financial summary for ALL terms (Term 1 Week 1 to current)
        summary = financial_dashboard.get_financial_summary(term=None)
        
        return jsonify({
            'success': True,
            'summary': {
                'student_revenue': float(summary['student_revenue']),
                'teacher_expenses': float(summary['teacher_expenses']),
                'net_balance': float(summary['net_balance']),
                'term': 'All Terms'
            }
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error loading financial summary: {str(e)}'
        }), 500


@app.route('/api/add-teacher', methods=['POST'])
def add_teacher_endpoint():
    """API endpoint to add a new teacher."""
    try:
        if teacher_manager is None:
            return jsonify({
                'success': False,
                'message': 'Teacher payroll system not initialized'
            }), 503
        
        data = request.get_json()
        name = data.get('name', '').strip()
        subject = data.get('subject', '').strip()
        phone = data.get('phone', '').strip()
        email = data.get('email', '').strip()
        
        if not name or not subject:
            return jsonify({
                'success': False,
                'message': 'Teacher name and subject are required'
            }), 400
        
        # Add the teacher
        new_teacher = teacher_manager.add_teacher(
            name=name,
            subject=subject,
            phone=phone,
            email=email
        )
        
        return jsonify({
            'success': True,
            'message': f'Teacher {name} added successfully',
            'teacher': new_teacher
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error adding teacher: {str(e)}'
        }), 500


@app.route('/api/deactivate-teacher', methods=['POST'])
def deactivate_teacher_endpoint():
    """API endpoint to deactivate a teacher."""
    try:
        if teacher_manager is None:
            return jsonify({
                'success': False,
                'message': 'Teacher payroll system not initialized'
            }), 503
        
        data = request.get_json()
        teacher_id = data.get('teacher_id')
        
        if not teacher_id:
            return jsonify({
                'success': False,
                'message': 'Teacher ID is required'
            }), 400
        
        # Get teacher info before deactivating
        teacher = teacher_manager.get_teacher_by_id(int(teacher_id))
        if not teacher:
            return jsonify({
                'success': False,
                'message': 'Teacher not found'
            }), 404
        
        # Deactivate the teacher
        success = teacher_manager.deactivate_teacher(int(teacher_id))
        
        if success:
            return jsonify({
                'success': True,
                'message': f'Teacher {teacher["name"]} has been deactivated'
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Failed to deactivate teacher'
            }), 500
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error deactivating teacher: {str(e)}'
        }), 500


def main():
    """Main entry point for the web application."""
    print("=" * 60)
    print("Kirembe Secondary School - Payment Management System")
    print("Web Application")
    print("=" * 60)
    print()
    
    # Initialize services
    if initialize_services():
        print()
        print("=" * 60)
        print("Starting web server...")
        print("=" * 60)
        print()
        print("🌐 Open your browser and go to:")
        print()
        print("   http://localhost:5002")
        print()
        print("   or")
        print()
        print("   http://127.0.0.1:5002")
        print()
        print("=" * 60)
        print("Press Ctrl+C to stop the server")
        print("=" * 60)
        print()
        
        # Run the Flask app
        app.run(host='0.0.0.0', port=5002, debug=False)


if __name__ == "__main__":
    main()
