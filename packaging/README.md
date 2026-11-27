# PDF Converter: the Windows app

## Get the installer
Every change pushed to `main` builds a new installer on GitHub (workflow
"Windows app"). The packaged app checks itself first: every module,
Tesseract with Arabic, and a real PDF to Word conversion. An installer is
made only when all of that passes.

The newest installer is always on the **windows-latest** release:
GitHub → Releases → "PDF Converter … (latest build)" →
`PDFConverter-Setup-<version>.exe`.

## Install
1. Run `PDFConverter-Setup-<version>.exe`.
2. Windows may say the program is not signed: **More info → Run anyway**.
3. No administrator rights and no internet are needed.

A newer installer goes over the old one; settings and learned corrections
stay in `%LOCALAPPDATA%\PDFTool`.

## What the app is
`desktop_app.pyw` shows a loading card, starts the same page as
`streamlit run app.py` on this computer only (127.0.0.1), and opens it in a
Microsoft Edge app window. Closing the window stops it. The app window shows
the Word converter only (`PDF_TOOL_HIDE_EXCEL=1`).

If something goes wrong, the logs are in `%LOCALAPPDATA%\PDFTool\logs`
(`desktop.log`, `server.log`).

## Build on a Windows PC instead
Once: Python 3.12, `pip install -r requirements.txt pyinstaller`, Tesseract
in `C:\Program Files\Tesseract-OCR`, Inno Setup 6. Then:

    packaging\build.bat

## Check a packaged copy by hand

    "PDF Converter.exe" --selftest report.json
