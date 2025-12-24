"""
سكريبت للتحقق من صحة النتائج المحسوبة
"""
import pandas as pd
from attendance_calculator import AttendanceCalculator

def verify_results():
    """التحقق من النتائج"""
    print("="*60)
    print("التحقق من صحة النتائج المحسوبة")
    print("="*60)
    
    # قراءة البيانات
    print("\n📂 قراءة البيانات...")
    fingerprint_data = pd.read_csv('البصمات_من_الصور.csv', encoding='utf-8')
    shift_data = pd.read_csv('المناوبات_من_الصور.csv', encoding='utf-8')
    
    print(f"✅ ملف البصمات: {len(fingerprint_data)} صف")
    print(f"✅ ملف المناوبات: {len(shift_data)} صف")
    
    # إنشاء الآلة الحاسبة
    calculator = AttendanceCalculator(
        required_times=["08:00", "12:00", "15:00", "20:00", "23:00", "08:00"],
        tolerance_minutes=30,
        late_threshold=3,
        debug_mode=False
    )
    
    # حساب الحضور
    print("\n🔄 حساب الحضور...")
    results = calculator.calculate_attendance(fingerprint_data, shift_data)
    
    if results.empty:
        print("❌ لم يتم إنتاج أي نتائج!")
        return
    
    print(f"\n✅ تم حساب الحضور لـ {len(results)} موظف")
    
    # عرض النتائج
    print("\n" + "="*60)
    print("النتائج المحسوبة:")
    print("="*60)
    
    # عرض الأعمدة المهمة فقط
    display_cols = ['Name', 'Department', 'Total Working Days', 'Complete Days', 
                   'Incomplete Days', 'Absent Days', 'LateCount', 'LateMinutes',
                   'Required Checks', 'Actual Checks', 'Missing Checks', 
                   'Compliance Rate', 'FinalStatus']
    
    for col in display_cols:
        if col not in results.columns:
            print(f"⚠️  العمود '{col}' غير موجود في النتائج")
    
    # عرض النتائج بشكل منسق
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 20)
    
    print("\n" + results[display_cols].to_string(index=False))
    
    # مقارنة مع النتائج المتوقعة من الصورة
    print("\n" + "="*60)
    print("مقارنة مع النتائج المتوقعة:")
    print("="*60)
    
    expected_results = {
        'ابراهيم محجوب': {
            'Total Working Days': 3,
            'Complete Days': 0,
            'Incomplete Days': 3,
            'Absent Days': 0,
            'LateCount': 6,
            'LateMinutes': 74,
            'Required Checks': 18,
            'Actual Checks': 9,
            'Missing Checks': 9,
            'Compliance Rate': 50.0,
            'FinalStatus': 'غير ملتزم'
        },
        'حسن علي': {
            'Total Working Days': 3,
            'Complete Days': 1,
            'Incomplete Days': 2,
            'Absent Days': 0,
            'LateCount': 9,
            'LateMinutes': 340,
            'Required Checks': 18,
            'Actual Checks': 12,
            'Missing Checks': 6,
            'Compliance Rate': 66.67,
            'FinalStatus': 'غير ملتزم'
        },
        'رياض ياسين': {
            'Total Working Days': 3,
            'Complete Days': 2,
            'Incomplete Days': 1,
            'Absent Days': 0,
            'LateCount': 9,
            'LateMinutes': 105,
            'Required Checks': 18,
            'Actual Checks': 13,
            'Missing Checks': 5,
            'Compliance Rate': 72.22,
            'FinalStatus': 'غير ملتزم'
        },
        'صفاء طالب': {
            'Total Working Days': 3,
            'Complete Days': 0,
            'Incomplete Days': 3,
            'Absent Days': 0,
            'LateCount': 5,
            'LateMinutes': 146,
            'Required Checks': 18,
            'Actual Checks': 7,
            'Missing Checks': 11,
            'Compliance Rate': 38.89,
            'FinalStatus': 'غير ملتزم'
        },
        'عبيدة عامر': {
            'Total Working Days': 3,
            'Complete Days': 1,
            'Incomplete Days': 2,
            'Absent Days': 0,
            'LateCount': 9,
            'LateMinutes': 45,
            'Required Checks': 18,
            'Actual Checks': 12,
            'Missing Checks': 6,
            'Compliance Rate': 66.67,
            'FinalStatus': 'غير ملتزم'
        }
    }
    
    all_match = True
    for name, expected in expected_results.items():
        emp_result = results[results['Name'] == name]
        if emp_result.empty:
            print(f"\n❌ {name}: غير موجود في النتائج")
            all_match = False
            continue
        
        emp_row = emp_result.iloc[0]
        print(f"\n📊 {name}:")
        matches = True
        
        for col, expected_val in expected.items():
            if col not in emp_row:
                print(f"  ⚠️  {col}: غير موجود")
                matches = False
                continue
            
            actual_val = emp_row[col]
            
            # معالجة خاصة للقيم العشرية
            if isinstance(expected_val, float):
                if abs(actual_val - expected_val) < 0.01:
                    print(f"  ✅ {col}: {actual_val} (متوقع: {expected_val})")
                else:
                    print(f"  ❌ {col}: {actual_val} (متوقع: {expected_val})")
                    matches = False
            else:
                if actual_val == expected_val:
                    print(f"  ✅ {col}: {actual_val}")
                else:
                    print(f"  ❌ {col}: {actual_val} (متوقع: {expected_val})")
                    matches = False
        
        if not matches:
            all_match = False
    
    print("\n" + "="*60)
    if all_match:
        print("✅ جميع النتائج متطابقة مع النتائج المتوقعة!")
    else:
        print("⚠️  هناك بعض الاختلافات في النتائج")
    print("="*60)
    
    # حفظ النتائج للمقارنة
    results.to_csv('النتائج_المحسوبة.csv', index=False, encoding='utf-8-sig')
    print("\n💾 تم حفظ النتائج في: النتائج_المحسوبة.csv")

if __name__ == "__main__":
    verify_results()

