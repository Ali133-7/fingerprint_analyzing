"""
Report Generator Module
Handles creating a comprehensive, multi-sheet Excel report.
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import os

class ReportGenerator:
    def __init__(self):
        self.output_dir = "output"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        
        # Mapping of English column names to Arabic
        self.column_mapping = {
            # Employee Summary
            'Name': 'اسم الموظف',
            'Department': 'القسم',
            'Total Working Days': 'عدد أيام الدوام',
            'Complete Days': 'الأيام المستوفية',
            'Incomplete Days': 'أيام النقص',
            'Absent Days': 'أيام الغياب',
            'LateCount': 'عدد مرات التأخير',
            'LateMinutes': 'مجموع دقائق التأخير',
            'Actual Checks': 'البصمات الفعلية',
            'Required Checks': 'البصمات المطلوبة',
            'Missing Checks': 'البصمات المفقودة',
            'Compliance Rate': 'نسبة الالتزام %',
            'FinalStatus': 'الحالة النهائية',
            # Daily Details
            'Date': 'التاريخ',
            'Day Status': 'الحالة اليومية',
            'Attendance Status': 'تفاصيل المطابقة',
            # Late Analysis
            'Required Time': 'الوقت المطلوب',
            'Actual Time': 'الوقت الفعلي',
            'Delay (minutes)': 'التأخير (بالدقائق)',
            # Matching Log
            'Tolerance Window': 'نافذة التسامح',
            'Match Result': 'نتيجة المطابقة',
            'Matched Punch': 'البصمة المطابقة',
            'Failure Reason': 'سبب الفشل'
        }

    def generate_full_report(self, employee_summary_data, daily_details_data, filename=None, tolerance_minutes=30):
        """
        Generate full attendance report with validation
        
        Args:
            employee_summary_data: DataFrame with employee summary
            daily_details_data: DataFrame with daily details
            filename: Output filename (optional)
            tolerance_minutes: Tolerance in minutes for late calculation (default: 30)
        """
        self.tolerance_minutes = tolerance_minutes
        # Validate employee summary data
        if employee_summary_data is None or employee_summary_data.empty:
            raise ValueError("بيانات ملخص الموظفين فارغة أو غير صحيحة")
        
        required_summary_cols = ['Name', 'Department', 'Total Working Days', 'Complete Days', 
                                'Incomplete Days', 'Absent Days', 'Actual Checks', 'Required Checks']
        missing_summary_cols = [col for col in required_summary_cols if col not in employee_summary_data.columns]
        if missing_summary_cols:
            raise ValueError(f"بيانات ملخص الموظفين غير مكتملة. أعمدة مفقودة: {', '.join(missing_summary_cols)}")
        
        # Validate daily details data
        if daily_details_data is None or daily_details_data.empty:
            raise ValueError("بيانات التفاصيل اليومية فارغة أو غير صحيحة")
        
        required_daily_cols = ['Name', 'Department', 'Date', 'Day Status', 'Actual Checks', 'Required Checks']
        missing_daily_cols = [col for col in required_daily_cols if col not in daily_details_data.columns]
        if missing_daily_cols:
            raise ValueError(f"بيانات التفاصيل اليومية غير مكتملة. أعمدة مفقودة: {', '.join(missing_daily_cols)}")
        
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"full_attendance_report_{timestamp}.xlsx"
        
        filepath = os.path.join(self.output_dir, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Sheet 1: Employee Summary
            self._write_employee_summary_sheet(writer, employee_summary_data)
            
            # Sheet 2: Daily Attendance Details
            self._write_daily_details_sheet(writer, daily_details_data)
            
            # Sheet 3: Absences (الغيابات)
            self._write_absences_sheet(writer, employee_summary_data)
            
            # Sheet 4: Missing Punches (البصمات المتروكة)
            self._write_missing_punches_sheet(writer, daily_details_data)
            
            # Sheet 5: Raw Fingerprint Matching Log
            self._write_matching_log_sheet(writer, daily_details_data)
        
        return filepath

    def _write_employee_summary_sheet(self, writer, data):
        # Create a copy and rename columns to Arabic
        data_arabic = data.copy()
        data_arabic.columns = [self.column_mapping.get(col, col) for col in data_arabic.columns]
        
        data_arabic.to_excel(writer, sheet_name='ملخص الموظفين', index=False)
        ws = writer.sheets['ملخص الموظفين']
        self._format_sheet(ws, {
            'أيام الغياب': 'FFC7CE',      # Light Red
            'أيام النقص': 'FFEB9C', # Yellow
            'الأيام المستوفية': 'C6EFCE',    # Light Green
        })

    def _write_daily_details_sheet(self, writer, data):
        """
        Create daily details sheet - one row per day
        Missing punches details are available in the separate "البصمات المتروكة" sheet
        """
        # Select columns that exist in data
        available_cols = ['Name', 'Department', 'Date', 'Day Status', 'LateCount', 'LateMinutes', 
                         'Actual Checks', 'Required Checks', 'Missing Checks', 'Compliance Rate']
        cols_to_use = [col for col in available_cols if col in data.columns]
        daily_data = data[cols_to_use].copy()
        
        # Rename columns to Arabic
        daily_data.columns = [self.column_mapping.get(col, col) for col in daily_data.columns]
        
        # Sort by employee name, then by date
        if 'اسم الموظف' in daily_data.columns and 'التاريخ' in daily_data.columns:
            daily_data = daily_data.sort_values(['اسم الموظف', 'التاريخ'])
        
        daily_data.to_excel(writer, sheet_name='تفاصيل الحضور اليومي', index=False)
        ws = writer.sheets['تفاصيل الحضور اليومي']
        
        # Format sheet with dynamic status values
        self._format_sheet(ws, {
            'الحالة اليومية': {
                'غائب': 'FFC7CE',
                'مستوفي': 'C6EFCE',
            }
        })
        
        # Apply yellow background for rows with "نقص بصمة" in status
        from openpyxl.styles import PatternFill
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        
        # Find the status column index
        status_col_idx = None
        for idx, header in enumerate(ws[1], 1):
            if header.value == 'الحالة اليومية':
                status_col_idx = idx
                break
        
        # Apply yellow to rows with "نقص بصمة" in status
        if status_col_idx:
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=status_col_idx)
                if cell.value and 'نقص بصمة' in str(cell.value):
                    cell.fill = yellow_fill

    def _write_late_analysis_sheet(self, writer, data):
        """
        Create late analysis sheet showing all late arrivals
        Only includes delays that exceed tolerance_minutes
        """
        late_data = []
        
        # Check if data has required columns
        if data is None or data.empty:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['اسم الموظف', 'التاريخ', 'الوقت المطلوب', 'الوقت الفعلي', 'التأخير (بالدقائق)'])
            empty_df.to_excel(writer, sheet_name='تحليل التأخير', index=False)
            ws = writer.sheets['تحليل التأخير']
            self._format_sheet(ws, {})
            return
        
        # Check if Attendance Status column exists
        if 'Attendance Status' not in data.columns:
            # Try to get from original data if it was filtered
            # For now, create empty sheet
            empty_df = pd.DataFrame(columns=['اسم الموظف', 'التاريخ', 'الوقت المطلوب', 'الوقت الفعلي', 'التأخير (بالدقائق)'])
            empty_df.to_excel(writer, sheet_name='تحليل التأخير', index=False)
            ws = writer.sheets['تحليل التأخير']
            self._format_sheet(ws, {})
            return
        
        for _, row in data.iterrows():
            # Check if LateCount exists and is greater than 0
            late_count = row.get('LateCount', 0)
            try:
                late_count = float(late_count) if not pd.isna(late_count) else 0
            except (ValueError, TypeError):
                late_count = 0
                
            if late_count <= 0:
                continue
                
            # Check if Attendance Status exists and is a list
            attendance_status = row.get('Attendance Status', [])
            
            # Handle case where Attendance Status might be stored as string
            if isinstance(attendance_status, str):
                try:
                    import ast
                    attendance_status = ast.literal_eval(attendance_status)
                except:
                    continue
                    
            if not isinstance(attendance_status, list) or len(attendance_status) == 0:
                continue
                
            for status in attendance_status:
                if not isinstance(status, dict):
                    continue
                if not status.get('matched', False):
                    continue
                    
                actual_time = status.get('actual_time')
                required_time = status.get('required_time')
                if actual_time is None or required_time is None:
                    continue
                    
                try:
                    # Calculate delay in minutes
                    if hasattr(actual_time, 'strftime') and hasattr(required_time, 'strftime'):
                        delay = (actual_time - required_time).total_seconds() / 60
                    else:
                        # Try to convert to datetime if they're strings
                        try:
                            from datetime import datetime
                            if isinstance(actual_time, str):
                                # Try different formats
                                for fmt in ['%H:%M:%S', '%H:%M', '%Y-%m-%d %H:%M:%S']:
                                    try:
                                        actual_time = datetime.strptime(actual_time, fmt)
                                        break
                                    except:
                                        continue
                            if isinstance(required_time, str):
                                for fmt in ['%H:%M:%S', '%H:%M', '%Y-%m-%d %H:%M:%S']:
                                    try:
                                        required_time = datetime.strptime(required_time, fmt)
                                        break
                                    except:
                                        continue
                            if hasattr(actual_time, 'strftime') and hasattr(required_time, 'strftime'):
                                delay = (actual_time - required_time).total_seconds() / 60
                            else:
                                continue
                        except:
                            continue
                    
                    # Only count as late if delay exceeds tolerance_minutes
                    # If delay is within tolerance (0 to tolerance_minutes), it's not considered late
                    if delay > self.tolerance_minutes:
                        late_data.append({
                            'اسم الموظف': row.get('Name', 'N/A'),
                            'التاريخ': row.get('Date', 'N/A'),
                            'الوقت المطلوب': status.get('required_time_str', 'N/A'),
                            'الوقت الفعلي': actual_time.strftime('%H:%M:%S') if hasattr(actual_time, 'strftime') else str(actual_time),
                            'التأخير (بالدقائق)': int(delay - self.tolerance_minutes)  # Only count minutes beyond tolerance
                        })
                except (AttributeError, TypeError, ValueError) as e:
                    # Skip this entry if there's an error calculating delay
                    continue
        
        late_df = pd.DataFrame(late_data)
        if not late_df.empty:
            late_df.to_excel(writer, sheet_name='تحليل التأخير', index=False)
            ws = writer.sheets['تحليل التأخير']
            self._format_sheet(ws, {})
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['اسم الموظف', 'التاريخ', 'الوقت المطلوب', 'الوقت الفعلي', 'التأخير (بالدقائق)'])
            empty_df.to_excel(writer, sheet_name='تحليل التأخير', index=False)
            ws = writer.sheets['تحليل التأخير']
            self._format_sheet(ws, {})

    def _write_matching_log_sheet(self, writer, data):
        log_data = []
        for _, row in data.iterrows():
            # Check if Attendance Status exists and is a list
            attendance_status = row.get('Attendance Status', [])
            if not isinstance(attendance_status, list):
                continue
                
            for status in attendance_status:
                if not isinstance(status, dict):
                    continue
                    
                # Safely get values with defaults
                tolerance_start = status.get('tolerance_start')
                tolerance_end = status.get('tolerance_end')
                actual_time = status.get('actual_time')
                matched = status.get('matched', False)
                
                # Format tolerance window safely
                if tolerance_start is not None and hasattr(tolerance_start, 'strftime'):
                    tolerance_start_str = tolerance_start.strftime('%H:%M')
                else:
                    tolerance_start_str = 'N/A'
                    
                if tolerance_end is not None and hasattr(tolerance_end, 'strftime'):
                    tolerance_end_str = tolerance_end.strftime('%H:%M')
                else:
                    tolerance_end_str = 'N/A'
                
                # Format actual time safely
                if matched and actual_time is not None and hasattr(actual_time, 'strftime'):
                    actual_time_str = actual_time.strftime('%H:%M:%S')
                else:
                    actual_time_str = 'N/A'
                
                log_data.append({
                    'اسم الموظف': row['Name'],
                    'التاريخ': row['Date'],
                    'الوقت المطلوب': status.get('required_time_str', 'N/A'),
                    'نافذة التسامح': f"{tolerance_start_str} - {tolerance_end_str}",
                    'نتيجة المطابقة': 'مطابق' if matched else 'غير مطابق',
                    'البصمة المطابقة': actual_time_str,
                    'سبب الفشل': '' if matched else 'لم يتم العثور على بصمة ضمن نافذة التسامح'
                })
        
        log_df = pd.DataFrame(log_data)
        if not log_df.empty:
            log_df.to_excel(writer, sheet_name='سجل مطابقة البصمات', index=False)
            ws = writer.sheets['سجل مطابقة البصمات']
            self._format_sheet(ws, {})
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['اسم الموظف', 'التاريخ', 'الوقت المطلوب', 'نافذة التسامح', 
                                            'نتيجة المطابقة', 'البصمة المطابقة', 'سبب الفشل'])
            empty_df.to_excel(writer, sheet_name='سجل مطابقة البصمات', index=False)
            ws = writer.sheets['سجل مطابقة البصمات']
            self._format_sheet(ws, {})
    
    def _write_missing_punches_sheet(self, writer, data):
        """
        Create a sheet showing missing punches sorted by employee name
        """
        missing_punches_data = []
        
        for _, row in data.iterrows():
            # Check if Attendance Status exists and is a list
            attendance_status = row.get('Attendance Status', [])
            if not isinstance(attendance_status, list):
                continue
            
            # Get employee info
            emp_name = row.get('Name', 'N/A')
            emp_dept = row.get('Department', 'N/A')
            date = row.get('Date', 'N/A')
            
            # Find all unmatched required times (missing punches)
            for status in attendance_status:
                if not isinstance(status, dict):
                    continue
                
                matched = status.get('matched', False)
                if not matched:  # This is a missing punch
                    required_time_str = status.get('required_time_str', 'N/A')
                    description = status.get('description', 'N/A')
                    tolerance_start = status.get('tolerance_start')
                    tolerance_end = status.get('tolerance_end')
                    
                    # Format tolerance window
                    if tolerance_start is not None and hasattr(tolerance_start, 'strftime'):
                        tolerance_start_str = tolerance_start.strftime('%H:%M')
                    else:
                        tolerance_start_str = 'N/A'
                        
                    if tolerance_end is not None and hasattr(tolerance_end, 'strftime'):
                        tolerance_end_str = tolerance_end.strftime('%H:%M')
                    else:
                        tolerance_end_str = 'N/A'
                    
                    tolerance_window = f"{tolerance_start_str} - {tolerance_end_str}"
                    
                    missing_punches_data.append({
                        'اسم الموظف': emp_name,
                        'القسم': emp_dept,
                        'التاريخ': date,
                        'الوقت المطلوب': required_time_str,
                        'الوصف': description,
                        'نافذة التسامح': tolerance_window,
                        'الحالة': 'لم يبصم'
                    })
        
        # Create DataFrame and sort by employee name
        missing_df = pd.DataFrame(missing_punches_data)
        
        if not missing_df.empty:
            # Sort by employee name, then by date, then by required time
            missing_df = missing_df.sort_values(['اسم الموظف', 'التاريخ', 'الوقت المطلوب'])
            missing_df.to_excel(writer, sheet_name='البصمات المتروكة', index=False)
            ws = writer.sheets['البصمات المتروكة']
            self._format_sheet(ws, {})
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['اسم الموظف', 'القسم', 'التاريخ', 'الوقت المطلوب', 
                                            'الوصف', 'نافذة التسامح', 'الحالة'])
            empty_df.to_excel(writer, sheet_name='البصمات المتروكة', index=False)
            ws = writer.sheets['البصمات المتروكة']
            self._format_sheet(ws, {})
    
    def _write_absences_sheet(self, writer, data):
        """
        Create a sheet showing absences with reasons, sorted by employee name
        """
        absences_data = []
        
        for _, row in data.iterrows():
            absent_days = row.get('Absent Days', 0)
            if absent_days > 0:
                absences_data.append({
                    'اسم الموظف': row.get('Name', 'N/A'),
                    'القسم': row.get('Department', 'N/A'),
                    'عدد أيام الغياب': absent_days,
                    'عدد البصمات المتروكة': row.get('Missing Checks', 0),
                    'سبب الغياب': row.get('Absence Reason', 'N/A'),
                    'عدد أيام الدوام': row.get('Total Working Days', 0),
                    'نسبة الالتزام %': row.get('Compliance Rate', 0)
                })
        
        # Create DataFrame and sort by employee name
        absences_df = pd.DataFrame(absences_data)
        
        if not absences_df.empty:
            # Sort by employee name
            absences_df = absences_df.sort_values('اسم الموظف')
            absences_df.to_excel(writer, sheet_name='الغيابات', index=False)
            ws = writer.sheets['الغيابات']
            self._format_sheet(ws, {
                'عدد أيام الغياب': 'FFC7CE',  # Light Red for absence days
            })
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['اسم الموظف', 'القسم', 'عدد أيام الغياب', 
                                            'عدد البصمات المتروكة', 'سبب الغياب', 
                                            'عدد أيام الدوام', 'نسبة الالتزام %'])
            empty_df.to_excel(writer, sheet_name='الغيابات', index=False)
            ws = writer.sheets['الغيابات']
            self._format_sheet(ws, {})


    def _format_sheet(self, ws, color_map):
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx, column in enumerate(ws.columns):
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 40)
            ws.column_dimensions[column_letter].width = adjusted_width

            # Apply conditional formatting
            if column[0].value in color_map:
                rules = color_map[column[0].value]
                if isinstance(rules, dict): # Rule per value
                    for cell in column[1:]:
                        if cell.value in rules:
                            cell.fill = PatternFill(start_color=rules[cell.value], end_color=rules[cell.value], fill_type="solid")
                else: # Rule for any non-zero value
                    color = rules
                    for cell in column[1:]:
                        if cell.value and int(cell.value) > 0:
                             cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")


def main():
    """
    Example usage of the ReportGenerator
    """
    # Create dummy data for testing
    employee_summary = pd.DataFrame({
        'Name': ['Ahmad', 'Fatima'],
        'Department': ['Admin', 'Accounting'],
        'Complete Days': [20, 18],
        'Incomplete Days': [2, 3],
        'Absent Days': [0, 1],
    })

    daily_details = pd.DataFrame({
        'Name': ['Ahmad', 'Ahmad', 'Fatima'],
        'Department': ['Admin', 'Admin', 'Accounting'],
        'Date': [datetime(2023, 1, 1), datetime(2023, 1, 2), datetime(2023, 1, 1)],
        'Day Status': ['مستوفي', 'نقص بصمة', 'غائب'],
        'LateCount': [0, 1, 0],
        'LateMinutes': [0, 15, 0],
        'Attendance Status': [[{'matched': True, 'required_time': datetime(2023,1,1,8,0), 'required_time_str': '08:00', 'actual_time': datetime(2023,1,1,8,5), 'tolerance_start': datetime(2023,1,1,7,30), 'tolerance_end': datetime(2023,1,1,8,30)}]] * 3
    })
    
    generator = ReportGenerator()
    filepath = generator.generate_full_report(employee_summary, daily_details)
    print(f"Full report generated at: {filepath}")

if __name__ == '__main__':
    main()