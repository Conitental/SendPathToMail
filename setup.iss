#define MyAppName "SendPathToMail"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Constantin Heinzler"
#define MyAppURL "https://github.com/Conitental/SendPathToMail"

[Setup]
AppId={{DAC9996F-1E9E-498B-A871-CB9AF9861B38}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=LICENSE
OutputBaseFilename=SendPathToMail_Setup
Compression=lzma
SolidCompression=yes

SetupIconFile=icon.ico
UninstallDisplayIcon={app}\icon.ico

[Files]
Source: "sendPathToMail.vbs"; DestDir: "{app}"
Source: "icon.ico"; DestDir: "{app}"

[Registry]
; add the context menu item to files
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\SendPathToMail"; Flags: uninsdeletekey; ValueType: string; ValueData: "Send path as mail"
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\SendPathToMail"; Flags: uninsdeletekey; ValueType: string; ValueData: "{app}\icon.ico"; ValueName: "Icon"
Root: HKLM; Subkey: "SOFTWARE\Classes\*\shell\SendPathToMail\command"; Flags: uninsdeletekey; ValueType: string; ValueData: "wscript ""{app}\sendPathToMail.vbs"" ""%1"""
; add the context menu item to folders
Root: HKLM; Subkey: "SOFTWARE\Classes\Directory\shell\SendPathToMail"; Flags: uninsdeletekey; ValueType: string; ValueData: "Send path as mail"
Root: HKLM; Subkey: "SOFTWARE\Classes\Directory\shell\SendPathToMail"; Flags: uninsdeletekey; ValueType: string; ValueData: "{app}\icon.ico"; ValueName: "Icon"
Root: HKLM; Subkey: "SOFTWARE\Classes\Directory\shell\SendPathToMail\command"; Flags: uninsdeletekey; ValueType: string; ValueData: "wscript ""{app}\sendPathToMail.vbs"" ""%1"""