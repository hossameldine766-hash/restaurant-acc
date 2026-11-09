@echo off
chcp 65001 >nul
title بناء برنامج RestoBooks
color 0A

echo ============================================================
echo             بناء برنامج RestoBooks
echo ============================================================
echo.

cd /d "%~dp0"

REM ═══════════════ تنظيف الملفات القديمة ═══════════════
echo [تنظيف] حذف المجلدات القديمة...

if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "RestoBooks.spec" del /q "RestoBooks.spec"

echo ✅ تم التنظيف
echo.

REM ═══════════════ فحص الملفات ═══════════════
echo [فحص] التأكد من وجود الملفات المطلوبة...

if not exist "res1.py" (
    echo ❌ خطأ: res1.py غير موجود!
    pause
    exit /b 1
)
if not exist ".venv" (
    echo ❌ خطأ: .venv غير موجود!
    pause
    exit /b 1
)

echo ✅ كل الملفات موجودة
echo.

REM ═══════════════ تفعيل البيئة ═══════════════
echo [1/3] تفعيل البيئة الافتراضية...
call .venv\Scripts\activate.bat
echo ✅ تم التفعيل
echo.

REM ═══════════════ البناء ═══════════════
echo [2/3] بدء البناء (قد يستغرق 3-7 دقائق)...
echo.

pyinstaller --onefile --windowed ^
    --icon=app.ico ^
    --name="RestoBooks" ^
    --splash splash.png ^
    --add-data "m.png;." ^
    --add-data "app.ico;." ^
    --add-data "splash.png;." ^
    --hidden-import=tkcalendar ^
    --hidden-import=PIL._tkinter_finder ^
    --hidden-import=bidi ^
    --hidden-import=arabic_reshaper ^
    --clean ^
    res1.py

if errorlevel 1 (
    echo.
    echo ❌ حدث خطأ أثناء البناء!
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  ✅ تم البناء بنجاح!
echo  الملف الجاهز: dist\RestoBooks.exe
echo ============================================================
echo.

echo [3/3] فتح مجلد dist...
start "" "dist"

echo.
echo اكتملت العملية! اضغط أي زر للإغلاق...
pause >nul