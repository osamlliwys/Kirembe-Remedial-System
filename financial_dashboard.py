"""
Financial Dashboard - Calculates revenue vs expenses
© 2026 Kirembe Secondary School
"""
import openpyxl
from decimal import Decimal
from typing import Dict
from .payroll_calculator import PayrollCalculator


class FinancialDashboard:
    """Calculates financial summary: student revenue vs teacher payroll"""
    
    def __init__(self, student_excel: str = "REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx",
                 teacher_excel: str = "TEACHER_PAYROLL.xlsx"):
        self.student_excel = student_excel
        self.teacher_excel = teacher_excel
        self.payroll_calculator = PayrollCalculator(teacher_excel)
    
    def calculate_student_revenue(self) -> Decimal:
        """
        Calculate total student payments (NET REVENUE) from PAYMENT_DATES_LOG.xlsx.
        
        Sums all values in the Amount column (column D) of the Payment Log sheet.
        This updates dynamically as new payments are recorded.
        
        Returns:
            Total amount collected from students
        """
        try:
            import os
            
            log_file = 'PAYMENT_DATES_LOG.xlsx'
            if not os.path.exists(log_file):
                print(f"Payment dates log not found: {log_file}")
                return Decimal('0')
            
            wb = openpyxl.load_workbook(log_file, data_only=True)
            ws = wb['Payment Log'] if 'Payment Log' in wb.sheetnames else wb.active
            
            total_revenue = Decimal('0')
            # Sum ALL amount values in column D starting from row 2 (row 1 is the header)
            for row in ws.iter_rows(min_row=2, values_only=True):
                amount = row[3] if len(row) > 3 else None  # Column D = index 3
                if amount and isinstance(amount, (int, float)) and amount > 0:
                    total_revenue += Decimal(str(amount))
            
            wb.close()
            return total_revenue
            
        except FileNotFoundError:
            print("PAYMENT_DATES_LOG.xlsx not found")
            return Decimal('0')
        except Exception as e:
            print(f"Error calculating student revenue: {e}")
            return Decimal('0')
    
    def calculate_teacher_expenses(self, term: int = None) -> Decimal:
        """
        Calculate total teacher payroll (expenses) by summing all weekly summaries.
        
        Gets the grand total of weekly summaries from Term 1 Week 1 to the last existing
        payroll summary. This uses the "Close & Tally" totals from payroll_periods.json.
        
        Args:
            term: Optional term number to filter by (None = all terms from Term 1 Week 1 to current)
        
        Returns:
            Total teacher payroll expenses
        """
        try:
            import json
            import os
            
            # Check if payroll_periods.json exists (contains Close & Tally summaries)
            periods_file = 'payroll_periods.json'
            if os.path.exists(periods_file):
                with open(periods_file, 'r') as f:
                    periods = json.load(f)
                
                total_expenses = Decimal('0')
                
                # Sum all closed period grand_totals (teacher + staff)
                for period in periods:
                    if period.get('status') == 'closed':
                        if term is None or period.get('term') == term:
                            grand_total = period.get('grand_total', 0)
                            total_expenses += Decimal(str(grand_total))
                
                # If we found closed periods with tallies, use that total
                if total_expenses > 0:
                    return total_expenses
            
            # Fallback: Calculate from attendance records directly
            if term:
                summary = self.payroll_calculator.calculate_term_payroll(term)
                return Decimal(str(summary['total_payroll']))
            else:
                # Calculate for all terms from Term 1 Week 1 onwards
                records = self.payroll_calculator.get_all_payroll_records()
                total_expenses = sum(r['total_pay'] for r in records)
                return Decimal(str(total_expenses))
                
        except Exception as e:
            print(f"Error calculating teacher expenses: {e}")
            return Decimal('0')
    
    def get_financial_summary(self, term: int = None) -> Dict:
        """
        Get complete financial summary.
        
        Args:
            term: Optional term number to filter expenses by
        
        Returns:
            Dictionary with revenue, expenses, and net balance
        """
        revenue = self.calculate_student_revenue()
        expenses = self.calculate_teacher_expenses(term)
        net_balance = revenue - expenses
        
        return {
            'student_revenue': float(revenue),
            'teacher_expenses': float(expenses),
            'net_balance': float(net_balance),
            'term': term if term else 'All Terms'
        }
