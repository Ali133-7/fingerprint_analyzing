"""
Settings Widget for Attendance System
Allows modification of attendance times and tolerance
"""

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QGroupBox, QFormLayout, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QHBoxLayout, QLabel, QSpinBox, QDoubleSpinBox, QTimeEdit
from PyQt5.QtCore import Qt, QTime
from PyQt5.QtGui import QFont
from styles import StyleSheets, ColorPalette # Import the new styles


class SettingsWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Title
        title_label = QLabel("إعدادات النظام")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 14, QFont.Bold))
        title_label.setStyleSheet(StyleSheets.TitleLabel) # Apply title style
        layout.addWidget(title_label)
        
        # Create settings groups
        self.create_times_settings_group(layout)
        self.create_tolerance_settings_group(layout)
        
        # Save button
        save_btn = QPushButton("حفظ الإعدادات")
        save_btn.setStyleSheet(StyleSheets.PrimaryButton) # Apply primary button style
        save_btn.clicked.connect(self.save_settings)
        layout.addWidget(save_btn)
        
        # Add stretch to push everything up
        layout.addStretch()
        
    def create_times_settings_group(self, parent_layout):
        """Create the attendance times settings group"""
        group_box = QGroupBox("أوقات البصمات المطلوبة")
        group_box.setStyleSheet(StyleSheets.GroupBoxTitle) # Apply group box title style
        group_layout = QVBoxLayout(group_box)
        
        # Table for attendance times
        self.times_table = QTableWidget(6, 2)
        self.times_table.setHorizontalHeaderLabels(["الوقت", "الوصف"])
        self.times_table.horizontalHeader().setStyleSheet(StyleSheets.TableHeader) # Apply table header style
        self.times_table.setStyleSheet(StyleSheets.TableView) # Apply table view style
        self.times_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        # Set default times according to requirements (08:00, 12:00, 15:00, 20:00, 23:00, 08:00 next day)
        default_times = [
            ("08:00", "بداية الدوام"),
            ("12:00", "نهاية الصباح"),
            ("15:00", "بداية الوردية"),
            ("20:00", "نهاية الوردية"),
            ("23:00", "نهاية الليل"),
            ("08:00", "نهاية المناوبة (اليوم التالي)")
        ]
        
        for row, (time, description) in enumerate(default_times):
            time_edit = QTimeEdit()
            time_edit.setTime(QTime.fromString(time, "HH:mm"))
            time_edit.setDisplayFormat("HH:mm")
            
            self.times_table.setCellWidget(row, 0, time_edit)
            self.times_table.setItem(row, 1, QTableWidgetItem(description))
        
        group_layout.addWidget(self.times_table)
        parent_layout.addWidget(group_box)
        
    def create_tolerance_settings_group(self, parent_layout):
        """Create the tolerance settings group"""
        group_box = QGroupBox("إعدادات السماح")
        group_box.setStyleSheet(StyleSheets.GroupBoxTitle) # Apply group box title style
        group_layout = QFormLayout(group_box)
        
        # Tolerance spinbox (in minutes)
        self.tolerance_spinbox = QSpinBox()
        self.tolerance_spinbox.setRange(1, 120)  # 1 to 120 minutes
        self.tolerance_spinbox.setValue(30)  # Default 30 minutes
        self.tolerance_spinbox.setSuffix(" دقيقة")
        
        # Absence threshold spinbox (number of missing punches = 1 absence day)
        self.absence_threshold_spinbox = QSpinBox()
        self.absence_threshold_spinbox.setRange(1, 20)  # 1 to 20 missing punches
        self.absence_threshold_spinbox.setValue(3)  # Default: every 3 missing punches = 1 absence day
        
        group_layout.addRow("وقت السماح المسموح:", self.tolerance_spinbox)
        group_layout.addRow("عتبة الغياب (عدد البصمات المتروكة = يوم غياب واحد):", self.absence_threshold_spinbox)
        parent_layout.addWidget(group_box)
        
    def save_settings(self):
        """Save the current settings"""
        print("Settings saved!")
        # In a real implementation, this would save to a config file
        self.get_settings()
        
    def get_settings(self):
        """Get current settings values"""
        settings = {
            'times': [],
            'tolerance': self.tolerance_spinbox.value(),
            'absence_threshold': self.absence_threshold_spinbox.value()  # Number of missing punches = 1 absence day
        }
        
        # Get the times from the table
        for row in range(self.times_table.rowCount()):
            time_widget = self.times_table.cellWidget(row, 0)
            time_str = time_widget.time().toString("HH:mm")
            description_item = self.times_table.item(row, 1)
            description = description_item.text() if description_item else ""
            settings['times'].append((time_str, description))
            
        return settings


if __name__ == '__main__':
    # Test the widget
    import sys
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    widget = SettingsWidget()
    widget.show()
    sys.exit(app.exec_())