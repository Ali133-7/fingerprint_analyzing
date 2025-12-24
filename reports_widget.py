"""
Reports Widget for Attendance System
Handles generating attendance reports with advanced filtering
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QPushButton, 
                             QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QComboBox, 
                             QDateEdit, QFileDialog, QMessageBox, QTabWidget, QSpinBox, 
                             QDoubleSpinBox, QCheckBox, QLineEdit)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor
import pandas as pd
from styles import StyleSheets, ColorPalette
from report_generator import ReportGenerator


class ReportsWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.attendance_results = None
        self.daily_results = None
        self.filtered_attendance_results = None
        self.filtered_daily_results = None
        self.tolerance_minutes = 30  # Default tolerance, will be updated from calculator
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        # Title
        title_label = QLabel("التقارير والفلترة")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet(StyleSheets.TitleLabel)
        layout.addWidget(title_label)
        
        # Filters Section
        self.create_filters_section(layout)
        
        # Action Buttons
        buttons_layout = QHBoxLayout()
        
        apply_filter_btn = QPushButton("تطبيق الفلاتر")
        apply_filter_btn.setStyleSheet(StyleSheets.PrimaryButton)
        apply_filter_btn.clicked.connect(self.apply_filters)
        buttons_layout.addWidget(apply_filter_btn)
        
        clear_filter_btn = QPushButton("إزالة الفلاتر")
        clear_filter_btn.setStyleSheet(StyleSheets.SecondaryButton)
        clear_filter_btn.clicked.connect(self.clear_filters)
        buttons_layout.addWidget(clear_filter_btn)
        
        export_btn = QPushButton("تصدير التقرير المفلتر")
        export_btn.setStyleSheet(StyleSheets.ExportButton)
        export_btn.clicked.connect(self.export_filtered_report)
        buttons_layout.addWidget(export_btn)
        
        layout.addLayout(buttons_layout)
        
        # Create tab widget for results
        results_tab_widget = QTabWidget()
        results_tab_widget.setStyleSheet(StyleSheets.TabWidget)
        
        # Aggregated results table
        self.create_results_table()
        aggregated_group = QGroupBox("النتائج المجمعة")
        aggregated_group.setStyleSheet(StyleSheets.GroupBoxTitle)
        aggregated_layout = QVBoxLayout(aggregated_group)
        aggregated_layout.addWidget(self.results_table)
        results_tab_widget.addTab(aggregated_group, "النتائج المجمعة")
        
        # Daily results tab with a sub-tab for breakdown
        daily_tabs = QTabWidget()
        daily_tabs.setStyleSheet(StyleSheets.TabWidget)
        
        # Daily results table
        self.create_daily_results_table()
        daily_summary_group = QWidget()
        daily_summary_layout = QVBoxLayout(daily_summary_group)
        daily_summary_layout.addWidget(self.daily_results_table)
        daily_tabs.addTab(daily_summary_group, "ملخص الحضور اليومي")
        
        # Daily breakdown table
        self.create_daily_breakdown_table()
        daily_breakdown_group = QWidget()
        daily_breakdown_layout = QVBoxLayout(daily_breakdown_group)
        daily_breakdown_layout.addWidget(self.daily_breakdown_table)
        daily_tabs.addTab(daily_breakdown_group, "تفاصيل الحضور اليومي")
        
        results_tab_widget.addTab(daily_tabs, "التفاصيل اليومية")
        
        layout.addWidget(results_tab_widget)
        
        # Connect daily results table click to update breakdown
        self.daily_results_table.itemClicked.connect(self.display_breakdown_for_selected_day)
        
        layout.addStretch()
        
    def create_filters_section(self, parent_layout):
        """Create comprehensive filters section"""
        filters_group = QGroupBox("فلاتر البحث")
        filters_group.setStyleSheet(StyleSheets.GroupBoxTitle)
        filters_layout = QVBoxLayout(filters_group)
        
        # Row 1: Date Range
        date_row = QHBoxLayout()
        date_row.addWidget(QLabel("نطاق التاريخ:"))
        
        date_row.addWidget(QLabel("من:"))
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-1))
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setCalendarPopup(True)
        date_row.addWidget(self.start_date_edit)
        
        date_row.addWidget(QLabel("إلى:"))
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setCalendarPopup(True)
        date_row.addWidget(self.end_date_edit)
        
        self.date_filter_check = QCheckBox("تفعيل فلتر التاريخ")
        self.date_filter_check.setChecked(True)
        date_row.addWidget(self.date_filter_check)
        date_row.addStretch()
        
        filters_layout.addLayout(date_row)
        
        # Row 2: Employee and Department
        emp_dept_row = QHBoxLayout()
        
        emp_dept_row.addWidget(QLabel("الموظف:"))
        self.employee_combo = QComboBox()
        self.employee_combo.setEditable(True)
        self.employee_combo.addItem("الكل")
        self.employee_combo.setCurrentText("الكل")
        emp_dept_row.addWidget(self.employee_combo)
        
        emp_dept_row.addWidget(QLabel("القسم:"))
        self.department_combo = QComboBox()
        self.department_combo.setEditable(True)
        self.department_combo.addItem("الكل")
        self.department_combo.setCurrentText("الكل")
        emp_dept_row.addWidget(self.department_combo)
        
        filters_layout.addLayout(emp_dept_row)
        
        # Row 3: Status and Compliance Rate
        status_rate_row = QHBoxLayout()
        
        status_rate_row.addWidget(QLabel("الحالة اليومية:"))
        self.status_combo = QComboBox()
        self.status_combo.addItems(["الكل", "مستوفي", "نقص بصمة", "غائب"])
        self.status_combo.setCurrentText("الكل")
        status_rate_row.addWidget(self.status_combo)
        
        status_rate_row.addWidget(QLabel("نسبة الالتزام:"))
        self.compliance_min_spin = QDoubleSpinBox()
        self.compliance_min_spin.setRange(0, 100)
        self.compliance_min_spin.setValue(0)
        self.compliance_min_spin.setSuffix("%")
        status_rate_row.addWidget(QLabel("من"))
        status_rate_row.addWidget(self.compliance_min_spin)
        
        self.compliance_max_spin = QDoubleSpinBox()
        self.compliance_max_spin.setRange(0, 100)
        self.compliance_max_spin.setValue(100)
        self.compliance_max_spin.setSuffix("%")
        status_rate_row.addWidget(QLabel("إلى"))
        status_rate_row.addWidget(self.compliance_max_spin)
        
        self.compliance_filter_check = QCheckBox("تفعيل فلتر النسبة")
        status_rate_row.addWidget(self.compliance_filter_check)
        status_rate_row.addStretch()
        
        filters_layout.addLayout(status_rate_row)
        
        # Row 4: Additional filters
        additional_row = QHBoxLayout()
        
        additional_row.addWidget(QLabel("أيام الغياب:"))
        self.absent_days_min_spin = QSpinBox()
        self.absent_days_min_spin.setRange(0, 999)
        self.absent_days_min_spin.setValue(0)
        additional_row.addWidget(QLabel("من"))
        additional_row.addWidget(self.absent_days_min_spin)
        
        self.absent_days_max_spin = QSpinBox()
        self.absent_days_max_spin.setRange(0, 999)
        self.absent_days_max_spin.setValue(999)
        additional_row.addWidget(QLabel("إلى"))
        additional_row.addWidget(self.absent_days_max_spin)
        
        self.absent_filter_check = QCheckBox("تفعيل فلتر الغياب")
        additional_row.addWidget(self.absent_filter_check)
        additional_row.addStretch()
        
        filters_layout.addLayout(additional_row)
        
        parent_layout.addWidget(filters_group)
        
    def create_results_table(self):
        """Create table to show aggregated report results"""
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(13)
        self.results_table.setHorizontalHeaderLabels([
            "اسم الموظف", 
            "القسم", 
            "عدد أيام الدوام", 
            "الأيام المستوفية", 
            "أيام النقص", 
            "أيام الغياب", 
            "سبب الغياب",
            "عدد مرات التأخير", 
            "مجموع دقائق التأخير",
            "البصمات الفعلية",
            "البصمات المطلوبة",
            "البصمات المفقودة",
            "نسبة الالتزام %",
            "الحالة النهائية"
        ])
        self.results_table.horizontalHeader().setStyleSheet(StyleSheets.TableHeader)
        self.results_table.setStyleSheet(StyleSheets.TableView)
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
    def create_daily_results_table(self):
        """Create table to show daily attendance details"""
        self.daily_results_table = QTableWidget()
        self.daily_results_table.setColumnCount(12)
        self.daily_results_table.setHorizontalHeaderLabels([
            "اسم الموظف", 
            "القسم", 
            "التاريخ",
            "الحالة اليومية",
            "عدد البصمات المطابقة",
            "عدد البصمات المطلوبة",
            "عدد البصمات المفقودة",
            "عدد مرات التأخير",
            "مجموع دقائق التأخير",
            "نسبة الالتزام %",
            "تفاصيل البصمات المتروكة",
            "تفاصيل المطابقة"
        ])
        self.daily_results_table.horizontalHeader().setStyleSheet(StyleSheets.TableHeader)
        self.daily_results_table.setStyleSheet(StyleSheets.TableView)
        self.daily_results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
    def create_daily_breakdown_table(self):
        """Create table to show the breakdown of daily attendance"""
        self.daily_breakdown_table = QTableWidget()
        self.daily_breakdown_table.setColumnCount(6)
        self.daily_breakdown_table.setHorizontalHeaderLabels([
            "الوقت المطلوب",
            "الوصف",
            "نافذة التسامح",
            "البصمة المطابقة",
            "نتيجة المطابقة",
            "سبب الفشل"
        ])
        self.daily_breakdown_table.horizontalHeader().setStyleSheet(StyleSheets.TableHeader)
        self.daily_breakdown_table.setStyleSheet(StyleSheets.TableView)
        self.daily_breakdown_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
    def populate_dropdowns(self):
        """Populate employee and department dropdowns from actual data"""
        if self.attendance_results is not None and not self.attendance_results.empty:
            # Clear and populate employee combo
            self.employee_combo.clear()
            self.employee_combo.addItem("الكل")
            unique_employees = sorted(self.attendance_results['Name'].unique())
            for emp in unique_employees:
                self.employee_combo.addItem(str(emp))
            self.employee_combo.setCurrentText("الكل")
            
            # Clear and populate department combo
            self.department_combo.clear()
            self.department_combo.addItem("الكل")
            unique_departments = sorted(self.attendance_results['Department'].unique())
            for dept in unique_departments:
                self.department_combo.addItem(str(dept))
            self.department_combo.setCurrentText("الكل")
        
    def apply_filters(self):
        """Apply all selected filters and display results"""
        if self.attendance_results is None or self.attendance_results.empty or \
           self.daily_results is None or self.daily_results.empty:
            QMessageBox.warning(self, "تحذير", "لا توجد بيانات لتصفيتها. يرجى حساب الحضور أولاً.")
            return
        
        # Start with all results
        filtered_agg_results = self.attendance_results.copy()
        filtered_daily_results = self.daily_results.copy()
        
        # Filter by Employee
        selected_employee = self.employee_combo.currentText()
        if selected_employee and selected_employee != "الكل":
            filtered_agg_results = filtered_agg_results[filtered_agg_results['Name'] == selected_employee]
            filtered_daily_results = filtered_daily_results[filtered_daily_results['Name'] == selected_employee]
        
        # Filter by Department
        selected_department = self.department_combo.currentText()
        if selected_department and selected_department != "الكل":
            filtered_agg_results = filtered_agg_results[filtered_agg_results['Department'] == selected_department]
            filtered_daily_results = filtered_daily_results[filtered_daily_results['Department'] == selected_department]
        
        # Filter by Date Range
        if self.date_filter_check.isChecked():
            start_date = self.start_date_edit.date().toString("yyyy-MM-dd")
            end_date = self.end_date_edit.date().toString("yyyy-MM-dd")
            
            # Convert daily_results 'Date' column to datetime for filtering
            if 'Date' in filtered_daily_results.columns:
                if not pd.api.types.is_datetime64_any_dtype(filtered_daily_results['Date']):
                    filtered_daily_results['Date'] = pd.to_datetime(filtered_daily_results['Date'])
            
            filtered_daily_results = filtered_daily_results[
                (filtered_daily_results['Date'] >= start_date) & 
                (filtered_daily_results['Date'] <= end_date)
            ]
            
            # Re-aggregate filtered daily results for aggregated view
            if not filtered_daily_results.empty:
                agg_functions = {
                    'Complete Days': 'sum', 
                    'Incomplete Days': 'sum',
                    'Absent Days': 'sum',
                    'LateCount': 'sum',
                    'LateMinutes': 'sum',
                    'Required Checks': 'sum',
                    'Actual Checks': 'sum',
                    'Missing Checks': 'sum'
                }
                
                re_aggregated_results = filtered_daily_results.groupby(['Name', 'Department']).agg(agg_functions).reset_index()
                
                # Recalculate Compliance Rate
                re_aggregated_results['Compliance Rate'] = 0.0
                mask = re_aggregated_results['Required Checks'] > 0
                re_aggregated_results.loc[mask, 'Compliance Rate'] = round(
                    (re_aggregated_results.loc[mask, 'Actual Checks'] / 
                     re_aggregated_results.loc[mask, 'Required Checks']) * 100, 2
                )
                
                # Recalculate Total Working Days
                total_working_days_df = filtered_daily_results.groupby(['Name', 'Department'])['Date'].nunique().reset_index(name='Total Working Days')
                re_aggregated_results = re_aggregated_results.merge(total_working_days_df, on=['Name', 'Department'], how='left')
                
                # Calculate FinalStatus
                re_aggregated_results['FinalStatus'] = re_aggregated_results.apply(
                    lambda row: 'ملتزم' if row['Incomplete Days'] == 0 and row['Absent Days'] == 0 else 'غير ملتزم',
                    axis=1
                )
                
                # Ensure column order
                column_order = [
                    'Name', 'Department', 'Total Working Days',
                    'Complete Days', 'Incomplete Days', 'Absent Days', 
                    'LateCount', 'LateMinutes',
                    'Actual Checks', 'Required Checks', 'Missing Checks',
                    'Compliance Rate', 'FinalStatus'
                ]
                
                filtered_agg_results = re_aggregated_results[[col for col in column_order if col in re_aggregated_results.columns]]
            else:
                filtered_agg_results = pd.DataFrame()
        
        # Filter by Day Status
        selected_status = self.status_combo.currentText()
        if selected_status and selected_status != "الكل":
            filtered_daily_results = filtered_daily_results[filtered_daily_results['Day Status'] == selected_status]
            
            # Re-aggregate after status filter
            if not filtered_daily_results.empty:
                agg_functions = {
                    'Complete Days': 'sum', 
                    'Incomplete Days': 'sum',
                    'Absent Days': 'sum',
                    'LateCount': 'sum',
                    'LateMinutes': 'sum',
                    'Required Checks': 'sum',
                    'Actual Checks': 'sum',
                    'Missing Checks': 'sum'
                }
                
                re_aggregated_results = filtered_daily_results.groupby(['Name', 'Department']).agg(agg_functions).reset_index()
                
                # Recalculate metrics
                re_aggregated_results['Compliance Rate'] = 0.0
                mask = re_aggregated_results['Required Checks'] > 0
                re_aggregated_results.loc[mask, 'Compliance Rate'] = round(
                    (re_aggregated_results.loc[mask, 'Actual Checks'] / 
                     re_aggregated_results.loc[mask, 'Required Checks']) * 100, 2
                )
                
                total_working_days_df = filtered_daily_results.groupby(['Name', 'Department'])['Date'].nunique().reset_index(name='Total Working Days')
                re_aggregated_results = re_aggregated_results.merge(total_working_days_df, on=['Name', 'Department'], how='left')
                
                re_aggregated_results['FinalStatus'] = re_aggregated_results.apply(
                    lambda row: 'ملتزم' if row['Incomplete Days'] == 0 and row['Absent Days'] == 0 else 'غير ملتزم',
                    axis=1
                )
                
                column_order = [
                    'Name', 'Department', 'Total Working Days',
                'Complete Days', 'Incomplete Days', 'Absent Days', 
                'LateCount', 'LateMinutes',
                    'Actual Checks', 'Required Checks', 'Missing Checks',
                'Compliance Rate', 'FinalStatus'
                ]
                
                filtered_agg_results = re_aggregated_results[[col for col in column_order if col in re_aggregated_results.columns]]
        
        # Filter by Compliance Rate
        if self.compliance_filter_check.isChecked():
            min_rate = self.compliance_min_spin.value()
            max_rate = self.compliance_max_spin.value()
            filtered_agg_results = filtered_agg_results[
                (filtered_agg_results['Compliance Rate'] >= min_rate) & 
                (filtered_agg_results['Compliance Rate'] <= max_rate)
            ]
            
            # Filter daily results based on aggregated compliance rate
            if not filtered_agg_results.empty:
                valid_names = set(filtered_agg_results['Name'].unique())
                filtered_daily_results = filtered_daily_results[filtered_daily_results['Name'].isin(valid_names)]
        
        # Filter by Absent Days
        if self.absent_filter_check.isChecked():
            min_absent = self.absent_days_min_spin.value()
            max_absent = self.absent_days_max_spin.value()
            filtered_agg_results = filtered_agg_results[
                (filtered_agg_results['Absent Days'] >= min_absent) & 
                (filtered_agg_results['Absent Days'] <= max_absent)
            ]
            
            # Filter daily results based on aggregated absent days
            if not filtered_agg_results.empty:
                valid_names = set(filtered_agg_results['Name'].unique())
                filtered_daily_results = filtered_daily_results[filtered_daily_results['Name'].isin(valid_names)]
        
        # Store filtered results
        self.filtered_attendance_results = filtered_agg_results
        self.filtered_daily_results = filtered_daily_results
        
        # Display results
        self.display_filtered_results()
        
        QMessageBox.information(self, "نجاح", f"تم تطبيق الفلاتر. عدد النتائج: {len(filtered_agg_results)} موظف")
        
    def display_filtered_results(self):
        """Display filtered results in tables"""
        # Display aggregated results
        if self.filtered_attendance_results is not None and not self.filtered_attendance_results.empty:
            self.results_table.setRowCount(len(self.filtered_attendance_results))
            
            for row_idx, (_, row) in enumerate(self.filtered_attendance_results.iterrows()):
                for col_idx, col_name in enumerate([
                    'Name', 'Department', 'Total Working Days',
                    'Complete Days', 'Incomplete Days', 'Absent Days', 'Absence Reason',
                    'LateCount', 'LateMinutes', 'Actual Checks', 'Required Checks', 'Missing Checks', 'Compliance Rate', 'FinalStatus'
                ]):
                    if col_name in row:
                        value = str(row[col_name])
                        if col_name == 'Compliance Rate':
                            value = f"{value}%"
                        elif col_name == 'LateMinutes':
                            value = str(int(float(value)))
                        item = QTableWidgetItem(value)
                        
                        # Apply color coding
                        if col_name == 'Absent Days' and int(row[col_name]) > 0:
                            item.setBackground(QColor(220, 20, 60))
                        elif col_name == 'Incomplete Days' and int(row[col_name]) > 0:
                            item.setBackground(QColor(255, 165, 0))
                        elif col_name == 'Complete Days' and int(row[col_name]) > 0:
                            item.setBackground(QColor(0, 100, 0))
                        
                        self.results_table.setItem(row_idx, col_idx, item)
                    else:
                        self.results_table.setItem(row_idx, col_idx, QTableWidgetItem(""))
        else:
            self.results_table.setRowCount(0)
        
        # Display daily results
        if self.filtered_daily_results is not None and not self.filtered_daily_results.empty:
            self.daily_results_table.setRowCount(len(self.filtered_daily_results))
            
            for row_idx, (_, row) in enumerate(self.filtered_daily_results.iterrows()):
                # Get missing punches details
                attendance_status = row.get('Attendance Status', [])
                missing_punches_details = []
                
                if isinstance(attendance_status, list):
                    for status in attendance_status:
                        if not isinstance(status, dict):
                            continue
                        if not status.get('matched', False):
                            # This is a missing punch
                            required_time_str = status.get('required_time_str', 'N/A')
                            description = status.get('description', 'N/A')
                            required_time = status.get('required_time')
                            
                            # Format date and time
                            if required_time and hasattr(required_time, 'strftime'):
                                date_str = required_time.strftime('%Y-%m-%d')
                                time_str = required_time.strftime('%H:%M')
                                missing_punches_details.append(f"{date_str} {time_str} ({description})")
                            else:
                                missing_punches_details.append(f"{required_time_str} ({description})")
                
                missing_punches_str = "؛ ".join(missing_punches_details) if missing_punches_details else "لا توجد"
                
                for col_idx, col_name in enumerate([
                    'Name', 'Department', 'Date',
                    'Day Status', 'Actual Checks', 'Required Checks', 'Missing Checks',
                    'LateCount', 'LateMinutes', 'Compliance Rate', 'Missing Punches Details', 'Attendance Status'
                ]):
                    if col_name == 'Missing Punches Details':
                        # Add missing punches details column
                        value = missing_punches_str
                        item = QTableWidgetItem(value)
                    elif col_name in row:
                        value = str(row[col_name])
                        if col_name == 'Compliance Rate':
                            value = f"{value}%"
                        elif col_name in ['LateMinutes', 'LateCount', 'Actual Checks', 'Required Checks', 'Missing Checks']:
                            value = str(int(float(value)))
                        elif col_name == 'Attendance Status':
                            value = "اضغط لعرض التفاصيل"
                        else:
                            value = str(row[col_name])
                            
                        item = QTableWidgetItem(value)
                        
                        if col_name == 'Day Status':
                            status = str(row[col_name])
                            if 'غائب' in status:
                                item.setBackground(QColor(220, 20, 60))
                            elif 'نقص بصمة' in status:
                                item.setBackground(QColor(255, 165, 0))
                            elif 'مستوفي' in status:
                                item.setBackground(QColor(0, 100, 0))
                        
                        self.daily_results_table.setItem(row_idx, col_idx, item)
                    else:
                        self.daily_results_table.setItem(row_idx, col_idx, QTableWidgetItem(""))
        else:
            self.daily_results_table.setRowCount(0)
            self.clear_daily_breakdown_table()
    
    def clear_filters(self):
        """Clear all filters and display all results"""
        # Reset filter controls
        self.employee_combo.setCurrentText("الكل")
        self.department_combo.setCurrentText("الكل")
        self.status_combo.setCurrentText("الكل")
        self.compliance_min_spin.setValue(0)
        self.compliance_max_spin.setValue(100)
        self.compliance_filter_check.setChecked(False)
        self.absent_days_min_spin.setValue(0)
        self.absent_days_max_spin.setValue(999)
        self.absent_filter_check.setChecked(False)
        self.date_filter_check.setChecked(True)
        
        # Display all results
        self.filtered_attendance_results = self.attendance_results
        self.filtered_daily_results = self.daily_results
        self.display_actual_results()
        self.display_daily_results()
        
        QMessageBox.information(self, "نجاح", "تم إزالة جميع الفلاتر")
    
    def display_actual_results(self):
        """Display the actual attendance calculation results"""
        if self.attendance_results is None or self.attendance_results.empty:
            self.clear_results_table()
            return
        
        # Validate results before displaying
        required_cols = ['Name', 'Department', 'Total Working Days', 'Complete Days', 
                         'Incomplete Days', 'Absent Days', 'Actual Checks', 'Required Checks']
        missing_cols = [col for col in required_cols if col not in self.attendance_results.columns]
        if missing_cols:
            QMessageBox.critical(self, "خطأ", f"النتائج غير مكتملة. أعمدة مفقودة: {', '.join(missing_cols)}")
            self.clear_results_table()
            return
        
        self.results_table.setRowCount(len(self.attendance_results))
        
        for row_idx, (_, row) in enumerate(self.attendance_results.iterrows()):
            for col_idx, col_name in enumerate([
                'Name', 'Department', 'Total Working Days',
                'Complete Days', 'Incomplete Days', 'Absent Days', 'Absence Reason',
                'LateCount', 'LateMinutes', 'Actual Checks', 'Required Checks', 'Missing Checks', 'Compliance Rate', 'FinalStatus'
            ]):
                if col_name in row:
                    value = str(row[col_name])
                    if col_name == 'Compliance Rate':
                        value = f"{value}%"
                    elif col_name == 'LateMinutes':
                        value = str(int(float(value)))
                    item = QTableWidgetItem(value)
                    
                    # Apply color coding
                    if col_name == 'Absent Days' and int(row[col_name]) > 0:
                        item.setBackground(QColor(220, 20, 60))
                    elif col_name == 'Incomplete Days' and int(row[col_name]) > 0:
                        item.setBackground(QColor(255, 165, 0))
                    elif col_name == 'Complete Days' and int(row[col_name]) > 0:
                        item.setBackground(QColor(0, 100, 0))
                    
                    self.results_table.setItem(row_idx, col_idx, item)
                else:
                    self.results_table.setItem(row_idx, col_idx, QTableWidgetItem(""))
    
    def display_daily_results(self):
        """Display the daily attendance calculation results"""
        if self.daily_results is None or self.daily_results.empty:
            self.clear_daily_results_table()
            self.clear_daily_breakdown_table()
            return
        
        # Validate daily results before displaying
        required_cols = ['Name', 'Department', 'Date', 'Day Status', 'Actual Checks', 'Required Checks']
        missing_cols = [col for col in required_cols if col not in self.daily_results.columns]
        if missing_cols:
            QMessageBox.critical(self, "خطأ", f"النتائج اليومية غير مكتملة. أعمدة مفقودة: {', '.join(missing_cols)}")
            self.clear_daily_results_table()
            self.clear_daily_breakdown_table()
            return
        
        self.daily_results_table.setRowCount(len(self.daily_results))
        self.clear_daily_breakdown_table()
        
        for row_idx, (_, row) in enumerate(self.daily_results.iterrows()):
            # Get missing punches details
            attendance_status = row.get('Attendance Status', [])
            missing_punches_details = []
            
            if isinstance(attendance_status, list):
                for status in attendance_status:
                    if not isinstance(status, dict):
                        continue
                    if not status.get('matched', False):
                        # This is a missing punch
                        required_time_str = status.get('required_time_str', 'N/A')
                        description = status.get('description', 'N/A')
                        required_time = status.get('required_time')
                        
                        # Format date and time
                        if required_time and hasattr(required_time, 'strftime'):
                            date_str = required_time.strftime('%Y-%m-%d')
                            time_str = required_time.strftime('%H:%M')
                            missing_punches_details.append(f"{date_str} {time_str} ({description})")
                        else:
                            missing_punches_details.append(f"{required_time_str} ({description})")
            
            missing_punches_str = "؛ ".join(missing_punches_details) if missing_punches_details else "لا توجد"
            
            for col_idx, col_name in enumerate([
                'Name', 'Department', 'Date',
                'Day Status', 'Actual Checks', 'Required Checks', 'Missing Checks',
                'LateCount', 'LateMinutes', 'Compliance Rate', 'Missing Punches Details', 'Attendance Status'
            ]):
                if col_name == 'Missing Punches Details':
                    # Add missing punches details column
                    value = missing_punches_str
                    item = QTableWidgetItem(value)
                elif col_name in row:
                    value = str(row[col_name])
                    if col_name == 'Compliance Rate':
                        value = f"{value}%"
                    elif col_name in ['LateMinutes', 'LateCount', 'Actual Checks', 'Required Checks', 'Missing Checks']:
                        value = str(int(float(value)))
                    elif col_name == 'Attendance Status':
                        value = "اضغط لعرض التفاصيل"
                    else:
                        value = str(row[col_name])
                        
                    item = QTableWidgetItem(value)
                    
                    if col_name == 'Day Status':
                        status = str(row[col_name])
                        if 'غائب' in status:
                            item.setBackground(QColor(220, 20, 60))
                        elif 'نقص بصمة' in status:
                            item.setBackground(QColor(255, 165, 0))
                        elif 'مستوفي' in status:
                            item.setBackground(QColor(0, 100, 0))
                    
                    self.daily_results_table.setItem(row_idx, col_idx, item)
                else:
                    self.daily_results_table.setItem(row_idx, col_idx, QTableWidgetItem(""))

    def display_breakdown_for_selected_day(self, item):
        """Display the attendance breakdown for the selected day"""
        data_to_use = getattr(self, 'filtered_daily_results', None)
        if data_to_use is None or data_to_use.empty:
            data_to_use = self.daily_results
            
        if data_to_use is None or data_to_use.empty:
            self.clear_daily_breakdown_table()
            return
            
        selected_row = item.row()
        if selected_row < len(data_to_use):
            day_data = data_to_use.iloc[selected_row]
            attendance_status_list = day_data.get('Attendance Status', [])
            
            if not isinstance(attendance_status_list, list):
                attendance_status_list = []
            
            self.daily_breakdown_table.setRowCount(len(attendance_status_list))
            
            for i, status in enumerate(attendance_status_list):
                req_time = status.get('required_time_str', 'N/A')
                desc = status.get('description', 'N/A')
                tolerance_start = status.get('tolerance_start', None)
                tolerance_end = status.get('tolerance_end', None)
                actual_time = status.get('actual_time', None)
                matched = status.get('matched', False)
                
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
                
                # Set table items
                self.daily_breakdown_table.setItem(i, 0, QTableWidgetItem(req_time))
                self.daily_breakdown_table.setItem(i, 1, QTableWidgetItem(desc))
                self.daily_breakdown_table.setItem(i, 2, QTableWidgetItem(tolerance_window))
                
                if matched and actual_time is not None and hasattr(actual_time, 'strftime'):
                    actual_time_str = actual_time.strftime('%H:%M:%S')
                    self.daily_breakdown_table.setItem(i, 3, QTableWidgetItem(actual_time_str))
                    
                    match_result_item = QTableWidgetItem("مطابق")
                    match_result_item.setBackground(QColor(0, 100, 0))
                    self.daily_breakdown_table.setItem(i, 4, match_result_item)
                    self.daily_breakdown_table.setItem(i, 5, QTableWidgetItem(""))
                else:
                    self.daily_breakdown_table.setItem(i, 3, QTableWidgetItem(""))
                    match_result_item = QTableWidgetItem("غير مطابق")
                    match_result_item.setBackground(QColor(220, 20, 60))
                    self.daily_breakdown_table.setItem(i, 4, match_result_item)
                    failure_reason = "لم يتم العثور على بصمة ضمن نافذة التسامح"
                    self.daily_breakdown_table.setItem(i, 5, QTableWidgetItem(failure_reason))

    def clear_results_table(self):
        """Clear all data from the results table"""
        self.results_table.setRowCount(0)
    
    def clear_daily_results_table(self):
        """Clear all data from the daily results table"""
        self.daily_results_table.setRowCount(0)
    
    def clear_daily_breakdown_table(self):
        """Clear all data from the daily breakdown table"""
        self.daily_breakdown_table.setRowCount(0)
    
    def export_filtered_report(self):
        """Export the filtered report to Excel with Arabic column headers"""
        # Use filtered results if available, otherwise use full results
        data_to_export_agg = self.filtered_attendance_results if self.filtered_attendance_results is not None and not self.filtered_attendance_results.empty else self.attendance_results
        data_to_export_daily = self.filtered_daily_results if self.filtered_daily_results is not None and not self.filtered_daily_results.empty else self.daily_results
        
        # Ensure Attendance Status is preserved in daily data for late analysis
        if data_to_export_daily is not None and not data_to_export_daily.empty:
            # If Attendance Status is missing, try to get it from original daily_results
            if 'Attendance Status' not in data_to_export_daily.columns and self.daily_results is not None and not self.daily_results.empty:
                if 'Attendance Status' in self.daily_results.columns:
                    # Merge Attendance Status back from original data
                    merge_cols = ['Name', 'Date']
                    if all(col in data_to_export_daily.columns and col in self.daily_results.columns for col in merge_cols):
                        attendance_status_df = self.daily_results[merge_cols + ['Attendance Status']].drop_duplicates(subset=merge_cols)
                        data_to_export_daily = data_to_export_daily.merge(
                            attendance_status_df,
                            on=merge_cols,
                            how='left'
                        )
        
        if data_to_export_agg is None or data_to_export_agg.empty:
            QMessageBox.warning(self, "تحذير", "لا توجد بيانات لتصديرها. يرجى حساب الحضور أولاً.")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "حفظ التقرير",
            "",
            "Excel Files (*.xlsx)"
        )
        
        if file_path:
            try:
                generator = ReportGenerator()
                generator.generate_full_report(
                    data_to_export_agg,
                    data_to_export_daily,
                    file_path,
                    tolerance_minutes=self.tolerance_minutes
                )
                
                QMessageBox.information(self, "نجاح", f"تم تصدير التقرير بنجاح إلى:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير التقرير: {str(e)}")


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    widget = ReportsWidget()
    widget.show()
    sys.exit(app.exec_())
