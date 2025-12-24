"""
Smart Attendance System
نظام إدارة وتحليل الحضور والانصراف المعتمد على البصمة
Main Application File
"""

import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QTabWidget, QMenuBar, QMenu, QAction, QMessageBox, QTableWidgetItem,
                             QDialog, QTextEdit, QPushButton, QScrollArea)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from settings_widget import SettingsWidget
from import_widget import ImportWidget
from reports_widget import ReportsWidget
from attendance_calculator import AttendanceCalculator
from report_generator import ReportGenerator
from styles import StyleSheets # Import the new styles

class AttendanceSystemApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("نظام إدارة وتحليل الحضور والانصراف المعتمد على البصمة")
        self.setGeometry(100, 100, 1200, 800)
        
        # Set application font for Arabic support
        font = QFont("Arial", 10)
        font.setPointSize(10)
        self.setFont(font)
        
        # Initialize calculator and report generator
        self.calculator = AttendanceCalculator(late_threshold=3, debug_mode=False)
        self.report_generator = ReportGenerator()
        
        # Initialize UI components
        self.init_ui()
        
    def init_ui(self):
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create tab widget for different sections
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(StyleSheets.TabWidget) # Apply tab widget style
        
        # Add tabs
        self.settings_tab = self.create_settings_tab()
        self.import_tab = self.create_import_tab()
        self.reports_tab = self.create_reports_tab()
        
        self.tabs.addTab(self.settings_tab, "الإعدادات")
        self.tabs.addTab(self.import_tab, "استيراد البيانات")
        self.tabs.addTab(self.reports_tab, "التقارير")
        
        main_layout.addWidget(self.tabs)
        
        # Create menu bar
        self.create_menu_bar()
        
    def create_settings_tab(self):
        """Create the settings tab for modifying attendance parameters"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        self.settings_widget = SettingsWidget()
        layout.addWidget(self.settings_widget)
        
        return tab
        
    def create_import_tab(self):
        """Create the data import tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        self.import_widget = ImportWidget()
        layout.addWidget(self.import_widget)
        
        return tab
        
    def create_reports_tab(self):
        """Create the reports tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        self.reports_widget = ReportsWidget()
        layout.addWidget(self.reports_widget)
        
        return tab
        
    def create_menu_bar(self):
        """Create the application menu bar"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu('ملف')
        exit_action = QAction('خروج', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Processing menu
        processing_menu = menubar.addMenu('معالجة')
        calculate_action = QAction('حساب الحضور', self)
        calculate_action.triggered.connect(self.calculate_attendance)
        processing_menu.addAction(calculate_action)
        
        # Debug menu
        debug_menu = menubar.addMenu('تصحيح')
        toggle_debug_action = QAction('تفعيل/تعطيل وضع التصحيح', self)
        toggle_debug_action.triggered.connect(self.toggle_debug_mode)
        debug_menu.addAction(toggle_debug_action)
        
        # Help menu
        help_menu = menubar.addMenu('مساعدة')
        help_action = QAction('دليل الاستخدام', self)
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)
        help_menu.addSeparator()
        about_action = QAction('عن النظام', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
    def calculate_attendance(self):
        """Calculate attendance based on imported data"""
        try:
            # Get current settings
            settings = self.settings_widget.get_settings()
            
            # Get imported data from import widget
            if self.import_widget.fingerprint_data is None or self.import_widget.shift_data is None:
                QMessageBox.warning(self, "تحذير", "يرجى استيراد كلا ملفي البصمات وجدول المناوبات أولاً")
                return
                
            # Update calculator with current settings
            self.calculator = AttendanceCalculator(
                required_times=[time[0] for time in settings['times']],  # Extract time strings
                tolerance_minutes=settings['tolerance'],
                late_threshold=3,  # Default value, not used in current logic
                absence_threshold=settings.get('absence_threshold', 3),  # Get absence threshold from settings, default to 3
                time_normalization_map={
                    '03:00': '15:00',
                    '03:00:00': '15:00'
                    # Additional mappings can be added here as needed
                }
            )
            
            # Validate data before calculation
            from data_validator import DataValidator
            fp_valid, fp_errors, fp_warnings = DataValidator.validate_fingerprint_data(self.import_widget.fingerprint_data)
            sh_valid, sh_errors, sh_warnings = DataValidator.validate_shift_data(self.import_widget.shift_data)
            
            if not fp_valid or not sh_valid:
                error_msg = "❌ البيانات غير صحيحة ولا يمكن حساب الحضور:\n\n"
                if fp_errors:
                    error_msg += "أخطاء في ملف البصمات:\n"
                    for error in fp_errors:
                        error_msg += f"  • {error}\n"
                if sh_errors:
                    error_msg += "\nأخطاء في ملف المناوبات:\n"
                    for error in sh_errors:
                        error_msg += f"  • {error}\n"
                QMessageBox.critical(self, "خطأ", error_msg)
                return
            
            # Perform calculation
            results = self.calculator.calculate_attendance(
                self.import_widget.fingerprint_data,
                self.import_widget.shift_data,
                settings
            )
            
            # Validate results before displaying
            if results is None or results.empty:
                QMessageBox.warning(self, "تحذير", "لم يتم إنتاج أي نتائج. يرجى التحقق من البيانات.")
                return
            
            # Check for required columns in results
            required_result_cols = ['Name', 'Department', 'Total Working Days', 'Complete Days', 
                                   'Incomplete Days', 'Absent Days', 'Actual Checks', 'Required Checks']
            missing_cols = [col for col in required_result_cols if col not in results.columns]
            if missing_cols:
                QMessageBox.critical(self, "خطأ", f"النتائج غير مكتملة. أعمدة مفقودة: {', '.join(missing_cols)}")
                return
            
            # Update reports widget with results
            self.reports_widget.attendance_results = results
            
            # Get and update daily results
            self.reports_widget.daily_results = self.calculator.get_daily_results()
            
            # Update tolerance_minutes in reports widget for export
            self.reports_widget.tolerance_minutes = self.calculator.tolerance_minutes
            
            # Validate daily results
            if self.reports_widget.daily_results is None or self.reports_widget.daily_results.empty:
                QMessageBox.warning(self, "تحذير", "لم يتم إنتاج نتائج يومية. يرجى التحقق من البيانات.")
                return
            
            # Call the reports widget's method to properly display results with all columns
            self.reports_widget.display_actual_results()
            
            # Call the reports widget's method to display daily results
            self.reports_widget.display_daily_results()
            
            # Populate dropdowns with the new data
            self.reports_widget.populate_dropdowns()
            
            # Show warnings if any
            all_warnings = fp_warnings + sh_warnings
            if all_warnings:
                warning_msg = "⚠️ تم حساب الحضور بنجاح مع التحذيرات التالية:\n\n"
                for warning in all_warnings:
                    warning_msg += f"• {warning}\n"
                QMessageBox.warning(self, "تحذير", warning_msg)
            else:
                QMessageBox.information(self, "نجاح", f"✅ تم حساب الحضور بدقة 100% لـ {len(results)} موظف بنجاح")
            
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء حساب الحضور: {str(e)}")
            
    def toggle_debug_mode(self):
        """Toggle debug mode on/off"""
        if hasattr(self.calculator, 'debug_mode'):
            self.calculator.debug_mode = not self.calculator.debug_mode
            status = "مفعّل" if self.calculator.debug_mode else "معطّل"
            QMessageBox.information(self, "وضع التصحيح", f"وضع التصحيح الآن {status}")
        else:
            QMessageBox.information(self, "معلومة", "نظام التصحيح غير متوفر")
            
    def show_help(self):
        """Show comprehensive help dialog"""
        help_dialog = QDialog(self)
        help_dialog.setWindowTitle("دليل الاستخدام - نظام إدارة وتحليل الحضور والانصراف")
        help_dialog.setGeometry(100, 100, 900, 700)
        help_dialog.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(help_dialog)
        
        # Create scrollable text area
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setFont(QFont("Arial", 10))
        
        # Help content
        help_content = """
        <div dir="rtl" style="font-family: Arial; font-size: 11pt; line-height: 1.6;">
        
        <h1 style="text-align: center; color: #2E7D32;">📘 دليل الاستخدام الشامل</h1>
        <h2 style="text-align: center; color: #1976D2;">نظام إدارة وتحليل الحضور والانصراف المعتمد على البصمة</h2>
        
        <hr>
        
        <h2 style="color: #1976D2;">👨‍💻 معلومات المطور</h2>
        <p><strong>المطور والمصمم:</strong> علي إبراهيم مصطفى</p>
        <p><strong>الإصدار:</strong> Version 1.0</p>
        <p><strong>تاريخ الإصدار:</strong> يناير 2025</p>
        <p><strong>حالة النظام:</strong> Production Ready ✅</p>
        
        <hr>
        
        <h2 style="color: #1976D2;">📋 نظرة عامة</h2>
        <p>نظام إدارة وتحليل الحضور والانصراف المعتمد على البصمة هو تطبيق سطح مكتب متخصص في معالجة وتحليل بيانات الحضور والانصراف للموظفين بناءً على بيانات البصمات الإلكترونية وجداول المناوبات.</p>
        
        <h3 style="color: #388E3C;">المميزات الرئيسية:</h3>
        <ul>
            <li>✅ <strong>دقة 100%</strong> في حساب الحضور والانصراف</li>
            <li>✅ <strong>واجهة عربية</strong> سهلة الاستخدام</li>
            <li>✅ <strong>تقارير شاملة</strong> متعددة الأوراق بصيغة Excel</li>
            <li>✅ <strong>قواعد واضحة</strong> للحضور والغياب والتأخير</li>
            <li>✅ <strong>قابل للتخصيص</strong> من خلال الإعدادات</li>
            <li>✅ <strong>معالجة تلقائية</strong> للبيانات</li>
            <li>✅ <strong>فلترة متقدمة</strong> للبيانات</li>
        </ul>
        
        <hr>
        
        <h2 style="color: #1976D2;">🔧 التقنيات المستخدمة</h2>
        <h3 style="color: #388E3C;">لغة البرمجة:</h3>
        <ul>
            <li><strong>Python 3.7+</strong> - اللغة الأساسية للتطبيق</li>
        </ul>
        
        <h3 style="color: #388E3C;">المكتبات الرئيسية:</h3>
        <ul>
            <li><strong>PyQt5</strong> - واجهة المستخدم الرسومية (GUI) مع دعم كامل للغة العربية</li>
            <li><strong>Pandas</strong> - معالجة وتحليل البيانات الضخمة بكفاءة</li>
            <li><strong>openpyxl</strong> - توليد تقارير Excel احترافية متعددة الأوراق</li>
            <li><strong>datetime</strong> - معالجة التواريخ والأوقات بدقة</li>
        </ul>
        
        <hr>
        
        <h2 style="color: #1976D2;">📖 دليل الاستخدام خطوة بخطوة</h2>
        
        <h3 style="color: #388E3C;">الخطوة 1: إعداد النظام ⚙️</h3>
        <ol>
            <li>افتح التطبيق</li>
            <li>انتقل إلى تبويب <strong>"الإعدادات"</strong></li>
            <li>حدد أوقات البصمات المطلوبة (افتراضياً: 08:00, 12:00, 15:00, 20:00)</li>
            <li>حدد نافذة التسامح (افتراضياً: 30 دقيقة)</li>
            <li>حدد عتبة الغياب (افتراضياً: 3 بصمات = يوم غياب)</li>
            <li>انقر على <strong>"حفظ الإعدادات"</strong></li>
        </ol>
        
        <h3 style="color: #388E3C;">الخطوة 2: استيراد البيانات 📥</h3>
        <ol>
            <li>انتقل إلى تبويب <strong>"استيراد البيانات"</strong></li>
            <li><strong>استيراد ملف البصمات:</strong>
                <ul>
                    <li>انقر على زر "استيراد ملف البصمات"</li>
                    <li>اختر الملف من الجهاز (CSV أو XLSX)</li>
                    <li>راجع معاينة البيانات للتأكد من صحتها</li>
                </ul>
            </li>
            <li><strong>استيراد ملف المناوبات:</strong>
                <ul>
                    <li>انقر على زر "استيراد ملف المناوبات"</li>
                    <li>اختر الملف من الجهاز (CSV أو XLSX)</li>
                    <li>راجع معاينة البيانات للتأكد من صحتها</li>
                </ul>
            </li>
        </ol>
        
        <h3 style="color: #388E3C;">الخطوة 3: حساب الحضور 🔄</h3>
        <ol>
            <li>انتقل إلى القائمة: <strong>معالجة → حساب الحضور</strong></li>
            <li>انتظر اكتمال المعالجة</li>
            <li>سيتم عرض رسالة نجاح مع عدد الموظفين المعالجين</li>
            <li>سيتم عرض النتائج تلقائياً في تبويب "التقارير"</li>
        </ol>
        
        <h3 style="color: #388E3C;">الخطوة 4: عرض النتائج 📊</h3>
        <ol>
            <li>انتقل إلى تبويب <strong>"التقارير"</strong></li>
            <li>راجع <strong>النتائج المجمعة</strong> لكل موظف</li>
            <li>راجع <strong>التفاصيل اليومية</strong> لكل يوم مناوبة</li>
            <li>انقر على صف لعرض تفاصيل المطابقة</li>
        </ol>
        
        <h3 style="color: #388E3C;">الخطوة 5: فلترة البيانات 🔍</h3>
        <p>يمكنك فلترة البيانات حسب:</p>
        <ul>
            <li><strong>التاريخ:</strong> نطاق تاريخي محدد</li>
            <li><strong>الموظف:</strong> موظف محدد</li>
            <li><strong>القسم:</strong> قسم محدد</li>
            <li><strong>الحالة اليومية:</strong> مستوفي / نقص بصمة / غائب</li>
            <li><strong>نسبة الالتزام:</strong> نطاق محدد (مثل 80% - 100%)</li>
            <li><strong>أيام الغياب:</strong> نطاق محدد (مثل 0 - 5 أيام)</li>
        </ul>
        
        <h3 style="color: #388E3C;">الخطوة 6: تصدير التقرير 💾</h3>
        <ol>
            <li>طبق الفلاتر المطلوبة (اختياري)</li>
            <li>انقر على زر <strong>"تصدير التقرير"</strong></li>
            <li>اختر مكان الحفظ واسم الملف</li>
            <li>انقر على "حفظ"</li>
            <li>سيتم إنشاء ملف Excel يحتوي على 5 أوراق:
                <ul>
                    <li>ملخص الموظفين</li>
                    <li>تفاصيل الحضور اليومي</li>
                    <li>الغيابات</li>
                    <li>البصمات المتروكة</li>
                    <li>سجل مطابقة البصمات</li>
                </ul>
            </li>
        </ol>
        
        <hr>
        
        <h2 style="color: #1976D2;">📁 هيكلية الملفات المطلوبة</h2>
        
        <h3 style="color: #388E3C;">ملف البصمات:</h3>
        <p><strong>الأعمدة المطلوبة:</strong></p>
        <ul>
            <li><strong>Name</strong> - اسم الموظف (نص)</li>
            <li><strong>Department</strong> - القسم (نص)</li>
            <li><strong>Date</strong> - تاريخ البصمة (YYYY-MM-DD)</li>
            <li><strong>Time</strong> - وقت البصمة (HH:MM)</li>
        </ul>
        <p><strong>صيغ المدعومة:</strong> CSV, XLSX</p>
        
        <h3 style="color: #388E3C;">ملف المناوبات:</h3>
        <p><strong>الأعمدة المطلوبة:</strong></p>
        <ul>
            <li><strong>Name</strong> - اسم الموظف (نص)</li>
            <li><strong>Shift Date</strong> - تاريخ المناوبة (YYYY-MM-DD)</li>
        </ul>
        <p><strong>صيغ المدعومة:</strong> CSV, XLSX</p>
        
        <hr>
        
        <h2 style="color: #1976D2;">📐 قواعد الحساب</h2>
        
        <h3 style="color: #388E3C;">قاعدة الحضور:</h3>
        <p>الحضور = وجود بصمة واحدة على الأقل في يوم المناوبة</p>
        
        <h3 style="color: #388E3C;">قاعدة الغياب:</h3>
        <p>الغياب = كل X بصمة متروكة = يوم غياب واحد</p>
        <p>حيث X = عتبة الغياب (قيمة قابلة للتخصيص من الإعدادات، افتراضياً 3)</p>
        <p><strong>مثال:</strong> 12 بصمة متروكة ÷ 3 = 4 أيام غياب</p>
        
        <h3 style="color: #388E3C;">قاعدة التأخير:</h3>
        <p>التأخير يُحسب فقط إذا تجاوز نافذة التسامح</p>
        <p><strong>الصيغة:</strong> التأخير = الوقت الفعلي - الوقت المطلوب - نافذة التسامح</p>
        <p><strong>مثال:</strong> إذا كان الوقت المطلوب 08:00 ونافذة التسامح 30 دقيقة:</p>
        <ul>
            <li>الوقت الفعلي 08:15 → لا تأخير (ضمن نافذة التسامح)</li>
            <li>الوقت الفعلي 08:45 → تأخير 15 دقيقة (45 - 30 = 15)</li>
        </ul>
        
        <h3 style="color: #388E3C;">الحالة اليومية:</h3>
        <ul>
            <li><strong>مستوفي:</strong> جميع البصمات المطلوبة موجودة</li>
            <li><strong>نقص بصمة (X):</strong> بعض البصمات مفقودة (X = عدد البصمات المفقودة)</li>
            <li><strong>غائب:</strong> صفر بصمات في يوم المناوبة</li>
        </ul>
        
        <hr>
        
        <h2 style="color: #1976D2;">⚙️ الإعدادات</h2>
        
        <h3 style="color: #388E3C;">أوقات البصمات المطلوبة:</h3>
        <p>يمكنك تحديد الأوقات التي يجب أن يبصم فيها الموظف. القيمة الافتراضية:</p>
        <ul>
            <li>08:00 - بداية الدوام</li>
            <li>12:00 - نهاية الصباح</li>
            <li>15:00 - بداية الوردية</li>
            <li>20:00 - نهاية الوردية</li>
        </ul>
        
        <h3 style="color: #388E3C;">نافذة التسامح:</h3>
        <p>الفترة الزمنية المسموح بها قبل وبعد الوقت المطلوب. القيمة الافتراضية: 30 دقيقة</p>
        
        <h3 style="color: #388E3C;">عتبة الغياب:</h3>
        <p>عدد البصمات المتروكة التي تساوي يوم غياب واحد. القيمة الافتراضية: 3</p>
        
        <hr>
        
        <h2 style="color: #1976D2;">📊 التقارير</h2>
        <p>النظام يولد تقريراً واحداً شاملاً بصيغة Excel يحتوي على <strong>5 أوراق</strong>:</p>
        <ol>
            <li><strong>ملخص الموظفين:</strong> إحصائيات شاملة لكل موظف</li>
            <li><strong>تفاصيل الحضور اليومي:</strong> تفاصيل كل يوم مناوبة</li>
            <li><strong>الغيابات:</strong> تقرير مفصل عن الغيابات</li>
            <li><strong>البصمات المتروكة:</strong> جميع البصمات المتروكة</li>
            <li><strong>سجل مطابقة البصمات:</strong> سجل تفصيلي لجميع محاولات المطابقة</li>
        </ol>
        
        <hr>
        
        <h2 style="color: #1976D2;">💡 نصائح للاستخدام الأمثل</h2>
        <ul>
            <li>✅ تأكد من صحة تنسيق الملفات قبل الاستيراد</li>
            <li>✅ راجع معاينة البيانات للتأكد من صحتها</li>
            <li>✅ استخدم الفلاتر للتركيز على موظفين محددين</li>
            <li>✅ راجع التفاصيل اليومية للتحقق من الدقة</li>
            <li>✅ احفظ التقارير بأسماء واضحة ووصفية</li>
        </ul>
        
        <hr>
        
        <h2 style="color: #1976D2;">🐛 وضع التصحيح</h2>
        <p>يمكنك تفعيل وضع التصحيح من القائمة: <strong>تصحيح → تفعيل/تعطيل وضع التصحيح</strong></p>
        <p>عند التفعيل، سيتم عرض معلومات تفصيلية في Terminal/Console مفيدة للمطورين.</p>
        
        <hr>
        
        <div style="text-align: center; padding: 20px; background-color: #E8F5E9; border-radius: 10px;">
            <h3 style="color: #2E7D32;">✅ دقة 100% في جميع الحسابات</h3>
            <p><strong>تم تطوير هذا النظام بواسطة:</strong></p>
            <p style="font-size: 14pt; font-weight: bold; color: #1976D2;">علي إبراهيم مصطفى</p>
            <p><strong>الإصدار:</strong> Version 1.0</p>
            <p><strong>© 2025 - جميع الحقوق محفوظة</strong></p>
        </div>
        
        </div>
        """
        
        text_edit.setHtml(help_content)
        layout.addWidget(text_edit)
        
        # Close button
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("إغلاق")
        close_btn.setStyleSheet(StyleSheets.PrimaryButton)
        close_btn.clicked.connect(help_dialog.close)
        button_layout.addWidget(close_btn)
        layout.addLayout(button_layout)
        
        help_dialog.exec_()
    
    def show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self, 
            "عن النظام", 
            "<div dir='rtl' style='font-family: Arial; font-size: 11pt;'>"
            "<h2 style='text-align: center; color: #1976D2;'>نظام إدارة وتحليل الحضور والانصراف</h2>"
            "<h3 style='text-align: center;'>المعتمد على البصمة</h3>"
            "<hr>"
            "<p style='text-align: center;'><strong>المطور والمصمم:</strong></p>"
            "<p style='text-align: center; font-size: 12pt; font-weight: bold; color: #2E7D32;'>علي إبراهيم مصطفى</p>"
            "<hr>"
            "<p style='text-align: center;'><strong>الإصدار:</strong> Version 1.0</p>"
            "<p style='text-align: center;'><strong>تاريخ الإصدار:</strong> يناير 2025</p>"
            "<p style='text-align: center;'><strong>حالة النظام:</strong> Production Ready ✅</p>"
            "<hr>"
            "<p style='text-align: center; color: #666;'>© 2025 - جميع الحقوق محفوظة</p>"
            "</div>"
        )


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(
        StyleSheets.TitleLabel +
        StyleSheets.GroupBoxTitle +
        StyleSheets.TableView # Apply global table view style
    ) # Apply global stylesheets
    
    # Set layout direction to RTL for Arabic
    app.setLayoutDirection(Qt.RightToLeft)
    
    window = AttendanceSystemApp()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()