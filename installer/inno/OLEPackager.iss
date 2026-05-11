#define AppName "OLE Packager"
#define AppVersion "1.0.0"
#define AppPublisher "OLEPackager"
#define AddinFile "OLEPackager.dotm"
#define ZipToolFile "OLEPackagerZipTool.exe"

[Setup]
AppId={{9B96DE61-3B3F-4E60-A534-F9020D993DF2}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={userappdata}\Microsoft\Word\STARTUP
UninstallFilesDir={userappdata}\OLE Packager\Uninstall
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..\..\release
OutputBaseFilename=OLEPackagerSetup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\..\dist\OLEPackager.dotm"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; DestName: "{#AddinFile}"; Flags: ignoreversion
Source: "..\..\dist\OLEPackagerZipTool.exe"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; DestName: "{#ZipToolFile}"; Flags: ignoreversion

[InstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\OleRawDataInserter.dotm"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\FigurePackageZipTool.exe"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\unins000.exe"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\unins000.dat"

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\{#AddinFile}"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\{#ZipToolFile}"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\OleRawDataInserter.dotm"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\FigurePackageZipTool.exe"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\unins000.exe"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\unins000.dat"
