"""
Import Widget for Attendance System
Handles importing fingerprint reports and shift schedules
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
import pandas as pd
from styles import StyleSheets # Import the new styles
from data_validator import DataValidator # Import data validator


class ImportWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.fingerprint_data = None
        self.shift_data = None
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Title
        title_label = QLabel("استيراد البيانات")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 14, QFont.Bold))
        title_label.setStyleSheet(StyleSheets.TitleLabel) # Apply title style
        layout.addWidget(title_label)
        
        # Create import groups
        self.create_fingerprint_import_group(layout)
        self.create_shift_import_group(layout)
        
        # Data preview table
        self.create_preview_table(layout)
        
        # Import and validate button
        import_btn = QPushButton("استيراد وتحليل البيانات")
        import_btn.setStyleSheet(StyleSheets.PrimaryButton) # Apply primary button style
        import_btn.clicked.connect(self.import_and_validate)
        layout.addWidget(import_btn)
        
        # Add stretch to push everything up
        layout.addStretch()
        
    def create_fingerprint_import_group(self, parent_layout):
        """Create the fingerprint report import group"""
        group_box = QGroupBox("تقرير البصمات")
        group_box.setStyleSheet(StyleSheets.GroupBoxTitle) # Apply group box title style
        group_layout = QHBoxLayout(group_box)
        
        # Fingerprint file path label
        self.fingerprint_path_label = QLabel("لم يتم اختيار ملف")
        self.fingerprint_path_label.setWordWrap(True)
        group_layout.addWidget(self.fingerprint_path_label)
        
        # Browse button for fingerprint report
        browse_fingerprint_btn = QPushButton("اختيار ملف")
        browse_fingerprint_btn.setStyleSheet(StyleSheets.SecondaryButton) # Apply secondary button style
        browse_fingerprint_btn.clicked.connect(self.browse_fingerprint_file)
        group_layout.addWidget(browse_fingerprint_btn)
        
        parent_layout.addWidget(group_box)
        
    def create_shift_import_group(self, parent_layout):
        """Create the shift schedule import group"""
        group_box = QGroupBox("جدول المناوبات")
        group_box.setStyleSheet(StyleSheets.GroupBoxTitle) # Apply group box title style
        group_layout = QHBoxLayout(group_box)
        
        # Shift file path label
        self.shift_path_label = QLabel("لم يتم اختيار ملف")
        self.shift_path_label.setWordWrap(True)
        group_layout.addWidget(self.shift_path_label)
        
        # Browse button for shift schedule
        browse_shift_btn = QPushButton("اختيار ملف")
        browse_shift_btn.setStyleSheet(StyleSheets.SecondaryButton) # Apply secondary button style
        browse_shift_btn.clicked.connect(self.browse_shift_file)
        group_layout.addWidget(browse_shift_btn)
        
        parent_layout.addWidget(group_box)
        
    def create_preview_table(self, parent_layout):
        """Create table to preview imported data"""
        group_box = QGroupBox("معاينة البيانات")
        group_box.setStyleSheet(StyleSheets.GroupBoxTitle) # Apply group box title style
        group_layout = QVBoxLayout(group_box)
        
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(5)
        self.preview_table.setHorizontalHeaderLabels(["اسم الموظف", "القسم", "التاريخ", "الوقت", "الحالة"])
        self.preview_table.horizontalHeader().setStyleSheet(StyleSheets.TableHeader) # Apply table header style
        self.preview_table.setStyleSheet(StyleSheets.TableView) # Apply table view style
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        group_layout.addWidget(self.preview_table)
        parent_layout.addWidget(group_box)
        
    def browse_fingerprint_file(self):
        """Browse and select fingerprint report file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "اختر ملف تقرير البصمات",
            "",
            "Files (*.xlsx *.xls *.csv)"
        )
        
        if file_path:
            self.fingerprint_path_label.setText(file_path)
            self.fingerprint_file_path = file_path
            
    def browse_shift_file(self):
        """Browse and select shift schedule file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "اختر ملف جدول المناوبات",
            "",
            "Files (*.xlsx *.xls *.csv)"
        )
        
        if file_path:
            self.shift_path_label.setText(file_path)
            self.shift_file_path = file_path
            
    def import_and_validate(self):
        """Import and validate both fingerprint and shift data"""
        if not hasattr(self, 'fingerprint_file_path') or not hasattr(self, 'shift_file_path'):
            QMessageBox.warning(self, "تحذير", "يرجى اختيار كلا الملفين أولاً")
            return
            
        try:
            # Load fingerprint data
            try:
                if self.fingerprint_file_path.endswith('.csv'):
                    self.fingerprint_data = pd.read_csv(self.fingerprint_file_path, encoding='utf-8')
                else:
                    self.fingerprint_data = pd.read_excel(self.fingerprint_file_path)
                    
                # Check if the first row contains column headers (common Excel issue)
                # If columns are named Column1, Column2, etc., the headers are in the first data row
                if len(self.fingerprint_data.columns) > 0 and self.fingerprint_data.columns[0].startswith('Column'):
                    # First row contains headers, use it as column names
                    if len(self.fingerprint_data) > 0:
                        new_columns = self.fingerprint_data.iloc[0].tolist()
                        self.fingerprint_data = self.fingerprint_data.iloc[1:].copy()
                        self.fingerprint_data.columns = new_columns
                        # Reset index after dropping first row
                        self.fingerprint_data.reset_index(drop=True, inplace=True)
            except FileNotFoundError:
                QMessageBox.critical(self, "خطأ", f"لم يتم العثور على ملف البصمات:\n{self.fingerprint_file_path}")
                return
            except pd.errors.EmptyDataError:
                QMessageBox.critical(self, "خطأ", "ملف البصمات فارغ أو لا يحتوي على بيانات")
                return
            except Exception as e:
                QMessageBox.critical(self, "خطأ", f"فشل قراءة ملف البصمات:\n{str(e)}")
                return
                
            # Load shift data
            try:
                if self.shift_file_path.endswith('.csv'):
                    self.shift_data = pd.read_csv(self.shift_file_path, encoding='utf-8')
                else:
                    self.shift_data = pd.read_excel(self.shift_file_path)
                    
                # Check if the first row contains column headers (common Excel issue)
                # If columns are named Column1, Column2, etc., the headers are in the first data row
                if len(self.shift_data.columns) > 0 and self.shift_data.columns[0].startswith('Column'):
                    # First row contains headers, use it as column names
                    if len(self.shift_data) > 0:
                        new_columns = self.shift_data.iloc[0].tolist()
                        self.shift_data = self.shift_data.iloc[1:].copy()
                        self.shift_data.columns = new_columns
                        # Reset index after dropping first row
                        self.shift_data.reset_index(drop=True, inplace=True)
            except FileNotFoundError:
                QMessageBox.critical(self, "خطأ", f"لم يتم العثور على ملف المناوبات:\n{self.shift_file_path}")
                return
            except pd.errors.EmptyDataError:
                QMessageBox.critical(self, "خطأ", "ملف المناوبات فارغ أو لا يحتوي على بيانات")
                return
            except Exception as e:
                QMessageBox.critical(self, "خطأ", f"فشل قراءة ملف المناوبات:\n{str(e)}")
                return
                
            # Validate data structure first
            fingerprint_valid = self.validate_fingerprint_data()
            shift_valid = self.validate_shift_data()
            
            if not fingerprint_valid or not shift_valid:
                error_msg = "هناك مشاكل في البيانات:\n\n"
                if not fingerprint_valid:
                    error_msg += "ملف البصمات:\n"
                    if self.fingerprint_data is not None and not self.fingerprint_data.empty:
                        error_msg += f"الأعمدة الموجودة: {', '.join(self.fingerprint_data.columns.tolist())}\n"
                        error_msg += "الأعمدة المطلوبة: Name, Department, Date, Time\n"
                    else:
                        error_msg += "الملف فارغ أو لا يحتوي على بيانات\n"
                    error_msg += "\n"
                if not shift_valid:
                    error_msg += "ملف المناوبات:\n"
                    if self.shift_data is not None and not self.shift_data.empty:
                        error_msg += f"الأعمدة الموجودة: {', '.join(self.shift_data.columns.tolist())}\n"
                        error_msg += "الأعمدة المطلوبة: Name, Shift Date\n"
                    else:
                        error_msg += "الملف فارغ أو لا يحتوي على بيانات\n"
                QMessageBox.warning(self, "تحذير", error_msg)
                return
            
            # Advanced validation using DataValidator
            fp_valid, fp_errors, fp_warnings = DataValidator.validate_fingerprint_data(self.fingerprint_data)
            sh_valid, sh_errors, sh_warnings = DataValidator.validate_shift_data(self.shift_data)
            consistency_valid, consistency_errors, consistency_warnings = DataValidator.validate_data_consistency(
                self.fingerprint_data, self.shift_data
            )
            
            # Collect all errors and warnings
            all_errors = fp_errors + sh_errors + consistency_errors
            all_warnings = fp_warnings + sh_warnings + consistency_warnings
            
            if all_errors:
                error_msg = "❌ أخطاء في البيانات تمنع الاستيراد:\n\n"
                for error in all_errors:
                    error_msg += f"• {error}\n"
                QMessageBox.critical(self, "خطأ", error_msg)
                return
            
            # Show warnings if any
            if all_warnings:
                warning_msg = "⚠️ تحذيرات (البيانات قابلة للاستخدام ولكن قد تحتوي على مشاكل):\n\n"
                for warning in all_warnings:
                    warning_msg += f"• {warning}\n"
                reply = QMessageBox.question(
                    self, 
                    "تحذير", 
                    warning_msg + "\nهل تريد المتابعة رغم التحذيرات؟",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return
            
            QMessageBox.information(self, "نجاح", "تم استيراد البيانات بنجاح وتم التحقق من صحتها")
            self.update_preview_table()
                
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ غير متوقع أثناء استيراد البيانات:\n{str(e)}")
            
    def validate_fingerprint_data(self):
        """Validate fingerprint data structure"""
        # Check if there's data to validate
        if self.fingerprint_data is None:
            return True  # No data to validate, return True
        
        # Check if it's an empty DataFrame
        if hasattr(self.fingerprint_data, 'empty') and self.fingerprint_data.empty:
            return True  # No data to validate, return True
        
        # Check if it has columns to validate
        if not hasattr(self.fingerprint_data, 'columns'):
            return True  # No columns attribute, return True
        
        required_columns = ['Name', 'Department', 'Date', 'Time']
        
        # Map expected English column names to what the calculator expects
        column_mapping = {
            'Name': ['Name', 'اسم الموظف', 'Employee Name', 'EmployeeName', 'Employee_Name'],
            'Department': ['Department', 'القسم', 'القسم\t', 'Dept'],
            'Date': ['Date', 'التاريخ', 'Fingerprint_Date'],
            'Time': ['Time', 'الوقت', 'Fingerprint_Time']
        }
        
        # First, strip any whitespace (including tabs) from column names
        self.fingerprint_data.columns = self.fingerprint_data.columns.str.strip()
        
        # Rename columns to English if they exist in Arabic or other formats
        for eng_col, possible_names in column_mapping.items():
            for col_name in possible_names:
                if col_name in self.fingerprint_data.columns and eng_col != col_name:
                    self.fingerprint_data.rename(columns={col_name: eng_col}, inplace=True)
                    break
            
        # Check if all required columns exist
        missing_cols = []
        for col in required_columns:
            if col not in self.fingerprint_data.columns:
                missing_cols.append(col)
        
        if missing_cols:
            print(f"Missing required column(s) in fingerprint data: {', '.join(missing_cols)}")
            print(f"Available columns in fingerprint data: {', '.join(self.fingerprint_data.columns.tolist())}")
            return False
                    
        return True
        
    def validate_shift_data(self):
        """Validate shift schedule data structure"""
        # Check if there's data to validate
        if self.shift_data is None:
            return True  # No data to validate, return True
        
        # Check if it's an empty DataFrame
        if hasattr(self.shift_data, 'empty') and self.shift_data.empty:
            return True  # No data to validate, return True
        
        # Check if it has columns to validate
        if not hasattr(self.shift_data, 'columns'):
            return True  # No columns attribute, return True
        
        # For shift data, the system expects Name and Shift Date
        # We'll accept both 'Date' and 'Shift Date' as the date column
        
        # Map expected English column names to what the calculator expects
        column_mapping = {
            'Name': ['Name', 'اسم الموظف', 'Employee Name', 'EmployeeName', 'Employee_Name'],
            'Shift Date': ['Shift Date', 'تاريخ المناوبة', 'تاريخ المناوبة\t', 'ShiftDate', 'Date', 'التاريخ', 'Shift_Date']
        }
        
        # First, strip any whitespace (including tabs) from column names
        self.shift_data.columns = self.shift_data.columns.str.strip()
        
        # Rename columns to English if they exist in Arabic or other formats
        for eng_col, possible_names in column_mapping.items():
            for col_name in possible_names:
                if col_name in self.shift_data.columns and eng_col != col_name:
                    self.shift_data.rename(columns={col_name: eng_col}, inplace=True)
                    break
            
        # Check if required columns exist - accept either 'Shift Date' or 'Date'
        name_exists = 'Name' in self.shift_data.columns
        shift_date_exists = 'Shift Date' in self.shift_data.columns or 'Date' in self.shift_data.columns
        
        if not name_exists:
            print("Missing required column: Name in shift data")
            print(f"Available columns in shift data: {', '.join(self.shift_data.columns.tolist())}")
            return False
        
        if not shift_date_exists:
            print("Missing required column: Shift Date in shift data")
            print(f"Available columns in shift data: {', '.join(self.shift_data.columns.tolist())}")
            return False
            
        # If 'Date' exists but 'Shift Date' doesn't, rename it
        if 'Date' in self.shift_data.columns and 'Shift Date' not in self.shift_data.columns:
            self.shift_data.rename(columns={'Date': 'Shift Date'}, inplace=True)
        
        return True
        
    def update_preview_table(self):
        """Update the preview table with imported data"""
        if self.fingerprint_data is not None:
            # Limit to first 10 rows for preview
            preview_data = self.fingerprint_data.head(10)
            
            # Update column count to 5 (Name, Department, Date, Time, Status)
            self.preview_table.setColumnCount(5)
            self.preview_table.setHorizontalHeaderLabels(["اسم الموظف", "القسم", "التاريخ", "الوقت", "الحالة"])
            
            self.preview_table.setRowCount(len(preview_data))
            
            for row_idx, (_, row) in enumerate(preview_data.iterrows()):
                for col_idx, col_name in enumerate(['Name', 'Department', 'Date', 'Time']):
                    if col_name in row:
                        item = QTableWidgetItem(str(row[col_name]))
                        self.preview_table.setItem(row_idx, col_idx, item)
                        
            # Add a status column
            for row_idx in range(len(preview_data)):
                item = QTableWidgetItem("معلق")
                self.preview_table.setItem(row_idx, 4, item)


if __name__ == '__main__':
    # Test the widget
    import sys
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    widget = ImportWidget()
    widget.show()
    sys.exit(app.exec_())