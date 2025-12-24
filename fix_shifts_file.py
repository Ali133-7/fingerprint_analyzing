"""
سكريبت لإصلاح ملف المناوبات
"""
import pandas as pd

def fix_shifts_file():
    """إصلاح ملف المناوبات"""
    file_path = 'المناوبات.xlsx'
    
    print("محاولة قراءة الملف...")
    try:
        # قراءة الملف
        df = pd.read_excel(file_path)
        
        print(f"الأعمدة الحالية: {df.columns.tolist()}")
        print(f"عدد الصفوف: {len(df)}")
        
        # التحقق من أن الأعمدة هي Column1, Column2
        if len(df.columns) == 2 and df.columns[0].startswith('Column'):
            print("\nتم اكتشاف أن رؤوس الأعمدة في الصف الأول من البيانات")
            
            # استخدام الصف الأول كرؤوس أعمدة
            new_columns = df.iloc[0].tolist()
            print(f"رؤوس الأعمدة الجديدة: {new_columns}")
            
            # إزالة الصف الأول واستخدامه كرؤوس
            df = df.iloc[1:].copy()
            df.columns = new_columns
            
            # إعادة تعيين الفهرس
            df.reset_index(drop=True, inplace=True)
            
            print(f"\nبعد الإصلاح:")
            print(f"الأعمدة: {df.columns.tolist()}")
            print(f"عدد الصفوف: {len(df)}")
            print(f"\nأول 5 صفوف:")
            print(df.head())
            
            # حفظ الملف
            output_file = 'المناوبات_مصحح.xlsx'
            df.to_excel(output_file, index=False)
            print(f"\n✅ تم حفظ الملف المصحح في: {output_file}")
            print(f"   يمكنك استخدام هذا الملف بدلاً من الملف الأصلي")
            
        else:
            print("\nالملف يبدو صحيحاً بالفعل")
            print(f"الأعمدة: {df.columns.tolist()}")
            
    except PermissionError:
        print("\n❌ خطأ: الملف مفتوح في برنامج آخر")
        print("   يرجى إغلاق ملف Excel ثم المحاولة مرة أخرى")
    except Exception as e:
        print(f"\n❌ خطأ: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    fix_shifts_file()

