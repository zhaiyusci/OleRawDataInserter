#define AppName "Figure Package Word Add-in"
#define AppVersion "1.0.0"
#define AppPublisher "OleRawDataInserter"
#define AddinFile "OleRawDataInserter.dotm"
#define ZipToolFile "FigurePackageZipTool.exe"

[Setup]
AppId={{9B96DE61-3B3F-4E60-A534-F9020D993DF2}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={userappdata}\Microsoft\Word\STARTUP
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..\..\release
OutputBaseFilename=FigurePackageWordAddinSetup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\..\dist\OleRawDataInserter.dotm"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; DestName: "{#AddinFile}"; Flags: ignoreversion
Source: "..\..\dist\FigurePackageZipTool.exe"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; DestName: "{#ZipToolFile}"; Flags: ignoreversion

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\{#AddinFile}"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\{#ZipToolFile}"
