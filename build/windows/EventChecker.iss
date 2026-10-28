; Inno Setup script for EventChecker

#define MyAppName "Event Inspector"
#ifndef MyAppVersion
#define MyAppVersion "2.5.0.88"
#endif
#define MyAppPublisher ""
#define MyAppURL ""
#define MyAppExeName "EventInspector.exe"

[Setup]
AppId={{B7B1A29B-6C7B-4D28-9E12-3A4E1D3E1F00}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=EventInspector-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; Uncomment and set icon if you have one
SetupIconFile=..\..\assets\app.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons"; Flags: unchecked

[Files]
Source: "..\..\dist\EventInspector\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Optional params file packaged alongside the app if it exists
#ifexist "..\..\Default event + Default Params.xlsx"
Source: "..\..\Default event + Default Params.xlsx"; DestDir: "{app}"; Flags: ignoreversion
#endif

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
