"""
تحليل تفصيلي لبيانات موظف واحد للتحقق من الحسابات
"""
import pandas as pd
from datetime import datetime, timedelta
from attendance_calculator import AttendanceCalculator

def analyze_employee(name):
    """تحليل تفصيلي لموظف واحد"""
    print(f"\n{'='*60}")
    print(f"تحليل تفصيلي لـ: {name}")
    print(f"{'='*60}")
    
    # قراءة البيانات
    fingerprint_data = pd.read_csv('البصمات_من_الصور.csv', encoding='utf-8')
    shift_data = pd.read_csv('المناوبات_من_الصور.csv', encoding='utf-8')
    
    # تصفية بيانات الموظف
    emp_fingerprints = fingerprint_data[fingerprint_data['Name'] == name]
    emp_shifts = shift_data[shift_data['Name'] == name]
    
    print(f"\n📊 بيانات الموظف:")
    print(f"  عدد البصمات: {len(emp_fingerprints)}")
    print(f"  عدد أيام المناوبة: {len(emp_shifts)}")
    
    print(f"\n📅 أيام المناوبة:")
    for _, shift in emp_shifts.iterrows():
        print(f"  - {shift['Shift Date']}")
    
    print(f"\n🕐 البصمات حسب التاريخ:")
    for date in emp_shifts['Shift Date'].unique():
        day_fingerprints = emp_fingerprints[emp_fingerprints['Date'] == date]
        print(f"\n  📆 {date}:")
        if len(day_fingerprints) == 0:
            print("    ⚠️  لا توجد بصمات!")
        else:
            for _, fp in day_fingerprints.iterrows():
                print(f"    - {fp['Time']}")
        
        # البصمات في اليوم التالي (للبصمة السادسة)
        next_date = (pd.to_datetime(date) + timedelta(days=1)).strftime('%Y-%m-%d')
        next_day_fingerprints = emp_fingerprints[emp_fingerprints['Date'] == next_date]
        if len(next_day_fingerprints) > 0:
            print(f"\n  📆 {next_date} (اليوم التالي - للبصمة السادسة):")
            for _, fp in next_day_fingerprints.iterrows():
                print(f"    - {fp['Time']}")
    
    # حساب الحضور
    calculator = AttendanceCalculator(
        required_times=["08:00", "12:00", "15:00", "20:00", "23:00", "08:00"],
        tolerance_minutes=30,
        late_threshold=3,
        debug_mode=True  # تفعيل وضع التصحيح
    )
    
    print(f"\n{'='*60}")
    print("حساب الحضور (مع تفاصيل التصحيح):")
    print(f"{'='*60}")
    
    results = calculator.calculate_attendance(emp_fingerprints, emp_shifts)
    
    if not results.empty:
        print(f"\n✅ النتائج:")
        print(results.to_string(index=False))
        
        # الحصول على النتائج اليومية
        daily_results = calculator.get_daily_results()
        if not daily_results.empty:
            emp_daily = daily_results[daily_results['Name'] == name]
            print(f"\n📋 النتائج اليومية:")
            for _, day in emp_daily.iterrows():
                print(f"\n  📅 {day['Date']}:")
                print(f"    الحالة: {day['Day Status']}")
                print(f"    البصمات المطابقة: {day['Actual Checks']}")
                print(f"    البصمات المطلوبة: {day['Required Checks']}")
                print(f"    البصمات المفقودة: {day['Missing Checks']}")
                print(f"    عدد التأخيرات: {day['LateCount']}")
                print(f"    دقائق التأخير: {day['LateMinutes']}")

# تحليل موظفين مختلفين
if __name__ == "__main__":
    # تحليل ابراهيم محجوب (أكبر اختلاف)
    analyze_employee('ابراهيم محجوب')
    
    print("\n\n")
    
    # تحليل عبيدة عامر (متطابق تماماً)
    analyze_employee('عبيدة عامر')

