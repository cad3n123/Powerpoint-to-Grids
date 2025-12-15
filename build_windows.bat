@echo off
setlocal enabledelayedexpansion

echo ==== Grid Maker Builder (UI Update) ====

set "STARTDIR=%cd%"
set "PYTHON_INSTALL_DIR=%LocalAppData%\Programs\GridMakerPython"

if not exist "grid_maker.py" (
    echo Error: grid_maker.py not found.
    pause
    exit /b
)

REM --- STEP 1: PYTHON SETUP ---
for /f "delims=" %%P in ('where python 2^>nul') do (
    echo %%P | find /I "WindowsApps" >nul
    if errorlevel 1 set "USE_PYTHON=%%P" & goto :found_python
)
echo Python not found. Installing standalone...
mkdir "%PYTHON_INSTALL_DIR%"
curl -o python-installer.exe https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
start /wait python-installer.exe /quiet InstallAllUsers=0 TargetDir="%PYTHON_INSTALL_DIR%" Include_pip=1 PrependPath=0
del python-installer.exe
set "USE_PYTHON=%PYTHON_INSTALL_DIR%\python.exe"
:found_python
echo Using Python: %USE_PYTHON%

REM --- STEP 2: BUILD ENVIRONMENT ---
set "TEMPDIR=%STARTDIR%\build_temp"
if exist "%TEMPDIR%" rmdir /s /q "%TEMPDIR%"
mkdir "%TEMPDIR%"
cd /d "%TEMPDIR%"

"%USE_PYTHON%" -m venv venv
call venv\Scripts\activate.bat

echo Installing libraries...
pip install --upgrade pip
pip install pyinstaller python-pptx pdf2image PyQt5

REM --- STEP 3: DOWNLOAD POPPLER ---
if not exist "%STARTDIR%\poppler.zip" (
    echo Downloading Poppler...
    curl -L -o poppler.zip https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip
) else (
    copy "%STARTDIR%\poppler.zip" .
)

tar -xf poppler.zip
for /d %%D in (poppler-*) do set "POPPLER_ROOT=%%D"
set "POPPLER_BIN=%CD%\%POPPLER_ROOT%\Library\bin"

REM --- STEP 4: BUILD EXE ---
echo Building executable...
copy "%STARTDIR%\grid_maker.py" .

REM --hidden-import is crucial here because we import pptx inside the thread!
pyinstaller --onefile --windowed ^
    --add-binary "%POPPLER_BIN%;poppler_bin" ^
    --hidden-import pptx ^
    --hidden-import pdf2image ^
    --exclude-module PyQt5.QtWebEngine ^
    --exclude-module PyQt5.QtMultimedia ^
    --exclude-module PyQt5.QtSvg ^
    grid_maker.py

REM --- STEP 5: CLEANUP ---
move /Y dist\grid_maker.exe "%STARTDIR%\grid_maker.exe"
cd /d "%STARTDIR%"
rmdir /s /q "%TEMPDIR%"

echo.
echo ========================================================
echo BUILD COMPLETE: grid_maker.exe
echo ========================================================
pause