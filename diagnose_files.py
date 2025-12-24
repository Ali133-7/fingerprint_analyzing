"""
سكريبت تشخيصي للتحقق من هيكلية ملفات البصمات والمناوبات
"""
import pandas as pd
import sys
import os

def diagnose_file(file_path, file_type):
    """تشخيص ملف واحد"""
    print(f"\n{'='*60}")
    print(f"تشخيص ملف: {file_path}")
    print(f"نوع الملف: {file_type}")
    print(f"{'='*60}")
    
    if not os.path.exists(file_path):
        print(f"❌ الملف غير موجود: {file_path}")
        return False
    
    try:
        # قراءة الملف
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, encoding='utf-8')
        else:
            df = pd.read_excel(file_path)
        
        print(f"\n✅ تم قراءة الملف بنجاح")
        print(f"عدد الصفوف: {len(df)}")
        print(f"عدد الأعمدة: {len(df.columns)}")
        
        # عرض الأعمدة
        print(f"\n📋 الأعمدة الموجودة في الملف:")
        for i, col in enumerate(df.columns, 1):
            # عرض التمثيل الدقيق للعمود (للكشف عن مسافات أو أحرف خفية)
            col_repr = repr(col)
            print(f"  {i}. '{col}' (repr: {col_repr})")
        
        # التحقق من الأعمدة المطلوبة
        if file_type == 'fingerprints':
            required = ['Name', 'Department', 'Date', 'Time']
            column_mapping = {
                'Name': ['Name', 'اسم الموظف', 'Employee Name', 'EmployeeName', 'Employee_Name'],
                'Department': ['Department', 'القسم', 'القسم\t', 'Dept'],
                'Date': ['Date', 'التاريخ', 'Fingerprint_Date'],
                'Time': ['Time', 'الوقت', 'Fingerprint_Time']
            }
        else:  # shifts
            required = ['Name', 'Shift Date']
            column_mapping = {
                'Name': ['Name', 'اسم الموظف', 'Employee Name', 'EmployeeName', 'Employee_Name'],
                'Shift Date': ['Shift Date', 'تاريخ المناوبة', 'تاريخ المناوبة\t', 'ShiftDate', 'Date', 'التاريخ', 'Shift_Date']
            }
        
        # تنظيف أسماء الأعمدة
        df.columns = df.columns.str.strip()
        
        # محاولة إعادة تسمية الأعمدة
        print(f"\n🔄 محاولة إعادة تسمية الأعمدة...")
        renamed = {}
        for eng_col, possible_names in column_mapping.items():
            for col_name in possible_names:
                if col_name in df.columns and eng_col != col_name:
                    df.rename(columns={col_name: eng_col}, inplace=True)
                    renamed[col_name] = eng_col
                    print(f"  ✅ '{col_name}' → '{eng_col}'")
                    break
        
        # التحقق من الأعمدة المطلوبة
        print(f"\n✅ التحقق من الأعمدة المطلوبة:")
        missing = []
        for col in required:
            if col in df.columns:
                print(f"  ✅ '{col}' موجود")
            else:
                print(f"  ❌ '{col}' مفقود")
                missing.append(col)
        
        if missing:
            print(f"\n❌ الأعمدة المفقودة: {', '.join(missing)}")
            print(f"\n💡 الحلول المقترحة:")
            print(f"   1. تأكد من أن الملف يحتوي على الأعمدة التالية:")
            for col in missing:
                print(f"      - {col}")
            print(f"   2. يمكنك استخدام الأسماء العربية التالية:")
            for col in missing:
                if col in column_mapping:
                    print(f"      - {col}: {', '.join(column_mapping[col])}")
            return False
        else:
            print(f"\n✅ جميع الأعمدة المطلوبة موجودة!")
            return True
        
    except Exception as e:
        print(f"\n❌ خطأ في قراءة الملف: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """الدالة الرئيسية"""
    print("="*60)
    print("أداة تشخيص ملفات البصمات والمناوبات")
    print("="*60)
    
    # البحث عن الملفات
    fingerprint_files = [
        'بصمات الحضور.xlsx',
        'fingerprints.csv',
        'قالب_البصمات.csv'
    ]
    
    shift_files = [
        'المناوبات.xlsx',
        'shift_schedule.csv',
        'قالب_المناوبات.csv'
    ]
    
    # تشخيص ملفات البصمات
    fingerprint_found = False
    for file in fingerprint_files:
        if os.path.exists(file):
            diagnose_file(file, 'fingerprints')
            fingerprint_found = True
            break
    
    if not fingerprint_found:
        print("\n⚠️  لم يتم العثور على ملف البصمات")
        print("   الملفات المطلوبة:")
        for file in fingerprint_files:
            print(f"   - {file}")
    
    # تشخيص ملفات المناوبات
    shift_found = False
    for file in shift_files:
        if os.path.exists(file):
            diagnose_file(file, 'shifts')
            shift_found = True
            break
    
    if not shift_found:
        print("\n⚠️  لم يتم العثور على ملف المناوبات")
        print("   الملفات المطلوبة:")
        for file in shift_files:
            print(f"   - {file}")
    
    print("\n" + "="*60)
    print("انتهى التشخيص")
    print("="*60)

if __name__ == "__main__":
    main()

