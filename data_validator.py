"""
أداة للتحقق من دقة البيانات قبل الحساب
"""
import pandas as pd
from datetime import datetime

class DataValidator:
    """فئة للتحقق من دقة البيانات"""
    
    @staticmethod
    def validate_fingerprint_data(fingerprint_data):
        """
        التحقق من دقة بيانات البصمات
        Returns: (is_valid, errors_list, warnings_list)
        """
        errors = []
        warnings = []
        
        if fingerprint_data is None or fingerprint_data.empty:
            errors.append("بيانات البصمات فارغة")
            return False, errors, warnings
        
        # التحقق من الأعمدة المطلوبة
        required_cols = ['Name', 'Department', 'Date', 'Time']
        missing_cols = [col for col in required_cols if col not in fingerprint_data.columns]
        if missing_cols:
            errors.append(f"أعمدة مفقودة في بيانات البصمات: {', '.join(missing_cols)}")
            return False, errors, warnings
        
        # التحقق من صيغة التاريخ
        try:
            dates = pd.to_datetime(fingerprint_data['Date'], format='%Y-%m-%d', errors='coerce')
            invalid_dates = dates.isna().sum()
            if invalid_dates > 0:
                errors.append(f"عدد {invalid_dates} سجل يحتوي على تاريخ غير صحيح")
        except Exception as e:
            errors.append(f"خطأ في التحقق من صيغة التاريخ: {str(e)}")
        
        # التحقق من صيغة الوقت
        try:
            for time_str in fingerprint_data['Time']:
                try:
                    datetime.strptime(str(time_str), '%H:%M')
                except ValueError:
                    try:
                        datetime.strptime(str(time_str), '%H:%M:%S')
                    except ValueError:
                        errors.append(f"صيغة وقت غير صحيحة: {time_str}")
                        break
        except Exception as e:
            errors.append(f"خطأ في التحقق من صيغة الوقت: {str(e)}")
        
        # التحقق من القيم الفارغة
        for col in required_cols:
            null_count = fingerprint_data[col].isna().sum()
            if null_count > 0:
                warnings.append(f"العمود '{col}' يحتوي على {null_count} قيمة فارغة")
        
        # Note: Duplicate fingerprints are automatically removed during processing
        # No need to warn about them here
        
        is_valid = len(errors) == 0
        return is_valid, errors, warnings
    
    @staticmethod
    def validate_shift_data(shift_data):
        """
        التحقق من دقة بيانات المناوبات
        Returns: (is_valid, errors_list, warnings_list)
        """
        errors = []
        warnings = []
        
        if shift_data is None or shift_data.empty:
            errors.append("بيانات المناوبات فارغة")
            return False, errors, warnings
        
        # التحقق من الأعمدة المطلوبة
        required_cols = ['Name', 'Shift Date']
        missing_cols = [col for col in required_cols if col not in shift_data.columns]
        if missing_cols:
            errors.append(f"أعمدة مفقودة في بيانات المناوبات: {', '.join(missing_cols)}")
            return False, errors, warnings
        
        # التحقق من صيغة التاريخ
        try:
            dates = pd.to_datetime(shift_data['Shift Date'], format='%Y-%m-%d', errors='coerce')
            invalid_dates = dates.isna().sum()
            if invalid_dates > 0:
                errors.append(f"عدد {invalid_dates} سجل يحتوي على تاريخ غير صحيح")
        except Exception as e:
            errors.append(f"خطأ في التحقق من صيغة التاريخ: {str(e)}")
        
        # التحقق من القيم الفارغة
        for col in required_cols:
            null_count = shift_data[col].isna().sum()
            if null_count > 0:
                warnings.append(f"العمود '{col}' يحتوي على {null_count} قيمة فارغة")
        
        # التحقق من المناوبات المكررة (نفس الموظف، نفس التاريخ)
        duplicates = shift_data.duplicated(subset=['Name', 'Shift Date']).sum()
        if duplicates > 0:
            warnings.append(f"عدد {duplicates} مناوبة مكررة (نفس الموظف، نفس التاريخ)")
        
        is_valid = len(errors) == 0
        return is_valid, errors, warnings
    
    @staticmethod
    def validate_data_consistency(fingerprint_data, shift_data):
        """
        التحقق من اتساق البيانات بين ملف البصمات وملف المناوبات
        Returns: (is_valid, errors_list, warnings_list)
        """
        errors = []
        warnings = []
        
        if fingerprint_data is None or shift_data is None:
            return True, errors, warnings
        
        # التحقق من أن جميع الموظفين في جدول المناوبات لديهم بصمات
        shift_employees = set(shift_data['Name'].unique())
        fingerprint_employees = set(fingerprint_data['Name'].unique())
        
        employees_with_shifts_no_fingerprints = shift_employees - fingerprint_employees
        if employees_with_shifts_no_fingerprints:
            warnings.append(f"موظفون لديهم مناوبات ولكن لا توجد لهم بصمات: {', '.join(employees_with_shifts_no_fingerprints)}")
        
        # التحقق من أن جميع الموظفين في البصمات لديهم مناوبات
        employees_with_fingerprints_no_shifts = fingerprint_employees - shift_employees
        if employees_with_fingerprints_no_shifts:
            warnings.append(f"موظفون لديهم بصمات ولكن لا توجد لهم مناوبات: {', '.join(employees_with_fingerprints_no_shifts)}")
        
        # التحقق من أن أسماء الموظفين متطابقة (حساسة لحالة الأحرف)
        # هذا يتم التحقق منه في التحقق السابق
        
        is_valid = len(errors) == 0
        return is_valid, errors, warnings

