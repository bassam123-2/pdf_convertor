@echo off
rem Build the Windows installer on this PC.
rem
rem   packaging\build.bat
rem
rem Needs, once: Python 3.12 with "pip install -r requirements.txt pyinstaller",
rem Tesseract in C:\Program Files\Tesseract-OCR, and Inno Setup 6.
rem Makes dist\PDFConverter-Setup-<version>.exe after the packaged app has
rem passed its own self test.

setlocal
cd /d "%~dp0\.."

for /f "usebackq delims=" %%v in (`python -c "from engine.version import VERSION; print(VERSION)"`) do set VERSION=%%v
echo Building PDF Converter %VERSION%

python -m PyInstaller PDFTool.spec --noconfirm --clean || goto :failed

echo Checking the packaged app...
"dist\PDF Converter\PDF Converter.exe" --selftest "dist\selftest.json"
if errorlevel 1 (
  type "dist\selftest.json"
  goto :failed
)

set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="%ProgramFiles%\Inno Setup 6\ISCC.exe"
%ISCC% /DAppVersion=%VERSION% packaging\PDFTool.iss || goto :failed

echo.
echo Done: dist\PDFConverter-Setup-%VERSION%.exe
exit /b 0

:failed
echo.
echo The build failed. See the messages above.
exit /b 1
