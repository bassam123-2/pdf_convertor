; Inno Setup script for PDF Converter.
;
;   iscc packaging\PDFTool.iss
;
; Turns dist\PDF Converter (made by PyInstaller from PDFTool.spec) into one
; installer: dist\PDFConverter-Setup-<version>.exe. Installing a newer build
; over an older one replaces the program and keeps the user's settings,
; which live in %LOCALAPPDATA%\PDFTool, not in the install folder.
; No administrator rights are needed: it installs for the current user.

#define AppName "PDF Converter"
#ifndef AppVersion
  #define AppVersion "1.1.0"
#endif
#define AppPublisher "PDF Converter"
#define AppExe "PDF Converter.exe"

[Setup]
; Same AppId as the earlier "PDF Tool" builds, so this one replaces them.
AppId={{B7E4C2A1-6D58-4F0A-9C31-2A6E8B4D91F0}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\PDF Converter
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=PDFConverter-Setup-{#AppVersion}
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
VersionInfoVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
VersionInfoProductName={#AppName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
SetupLogging=yes

[Languages]
Name: "arabic"; MessagesFile: "compiler:Languages\Arabic.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "..\dist\PDF Converter\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; Files of the earlier "PDF Tool" build, if it was installed here.
Type: filesandordirs; Name: "{app}\_internal"

[Icons]
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
