@echo off
chcp 65001 >nul
cls
echo ========================================
echo    بناء تطبيق نظام الحضور والانصراف
echo    المطور: علي إبراهيم مصطفى
echo ========================================
echo.

REM التحقق من وجود Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ خطأ: Python غير مثبت أو غير موجود في PATH
    pause
    exit /b 1
)

REM التحقق من وجود PyInstaller
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  PyInstaller غير مثبت. جاري التثبيت...
    pip install pyinstaller
    if errorlevel 1 (
        echo ❌ فشل تثبيت PyInstaller
        pause
        exit /b 1
    )
)

REM التحقق من وجود الأيقونة
if not exist "icon.ico" (
    echo ⚠️  تحذير: ملف icon.ico غير موجود
    echo    سيتم بناء التطبيق بدون أيقونة
    echo.
)

echo تنظيف المجلدات السابقة...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"
for /d /r . %%d in (__pycache__) do @if exist "%%d" rmdir /s /q "%%d"
echo ✅ تم التنظيف
echo.

echo جاري بناء التطبيق...
echo هذا قد يستغرق بضع دقائق...
echo.

REM بناء التطبيق باستخدام PyInstaller
pyinstaller --clean --noconfirm build_exe.spec

echo.
echo ========================================
if exist "dist\نظام_الحضور_والانصراف.exe" (
    echo.
    echo ✅✅✅ تم بناء التطبيق بنجاح! ✅✅✅
    echo.
    echo 📁 الملف النهائي موجود في:
    echo    dist\نظام_الحضور_والانصراف.exe
    echo.
    echo 💡 يمكنك الآن:
    echo    - نسخ هذا الملف إلى أي جهاز Windows
    echo    - تشغيله مباشرة بدون الحاجة إلى Python
    echo    - مشاركته مع المستخدمين الآخرين
    echo.
    echo 📊 حجم الملف: 
    for %%A in ("dist\نظام_الحضور_والانصراف.exe") do echo    %%~zA بايت
    echo.
) else (
    echo.
    echo ❌ فشل بناء التطبيق
    echo.
    echo يرجى التحقق من:
    echo    - وجود جميع المكتبات المطلوبة
    echo    - وجود جميع ملفات Python في المشروع
    echo    - رسائل الخطأ أعلاه
    echo.
)
echo ========================================
pause

