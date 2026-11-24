; Inno Setup script. Installing a newer build of this AppId replaces the old one.
; User data stays in %LOCALAPPDATA%\PDFTool and is not touched.

#define AppName "PDF Tool"
#define AppVersion "1.0.0"
#define AppPublisher "PDF Tool"
#define AppExe "PDF Tool.exe"

[Setup]
AppId={{B7E4C2A1-6D58-4F0A-9C31-2A6E8B4D91F0}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\PDF Tool
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=PDFTool-Setup-{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
UninstallDisplayIcon={app}\{#AppExe}
VersionInfoVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
SetupLogging=yes
InfoAfterFile=UPDATE.txt

[Languages]
Name: "arabic"; MessagesFile: "compiler:Languages\Arabic.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "اختصار على سطح المكتب"; Flags: checkedonce

[Files]
Source: "..\dist\PDF Tool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "UPDATE.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\محوّل PDF"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
Name: "{autoprograms}\محوّل PDF"; Filename: "{app}\{#AppExe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "فتح محوّل PDF"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
