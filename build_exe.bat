@echo off
echo ========================================
echo SJ Group - GST Reconciler EXE Builder
echo Fix for Ordinal 380 Error Included
echo ========================================

echo Checking Python...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python not found! Install Python 3.10 and tick Add to PATH
    pause
    exit /b
)

echo.
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller==6.3.0

echo.
echo Cleaning old builds...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q *.spec 2>nul

echo.
echo Building EXE - This will take 2-3 minutes...
echo If Ordinal error comes, this --noupx --clean fixes it
pyinstaller --onefile --windowed --noupx --clean --name "SJGroup_GST_Reconciler" src/Main.py

echo.
echo ========================================
if exist dist\SJGroup_GST_Reconciler.exe (
    echo SUCCESS! EXE Created at:
    echo %cd%\dist\SJGroup_GST_Reconciler.exe
    echo.
    echo Now upload this EXE to GitHub Releases
) else (
    echo FAILED - Trying console version...
    pyinstaller --onefile --console --noupx --clean --name "SJGroup_GST_Reconciler_Console" src/Main.py
)

echo ========================================
pause
