#define AppName "OLE Packager"
#define AppVersion "0.6.0"
#define AppPublisher "OLEPackager"
#define WordAddinName "OLEPackager.WordAddIn"
#define WordAddinFile "OLEPackager.WordAddIn.vsto"
#define PowerPointAddinName "OLEPackager.PowerPointAddIn"
#define PowerPointAddinFile "OLEPackager.PowerPointAddIn.vsto"

[Setup]
AppId={{40E6FA86-2E63-4298-905B-1A78BEC8DEE2}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={userappdata}\OLE Packager\VSTO
UninstallFilesDir={userappdata}\OLE Packager\VSTO\Uninstall
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..\..\release
OutputBaseFilename=OLEPackagerVstoSetup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "full"; Description: "Microsoft Word and PowerPoint"
Name: "wordonly"; Description: "Microsoft Word only"
Name: "powerpointonly"; Description: "Microsoft PowerPoint only"
Name: "custom"; Description: "Custom"; Flags: iscustom

[Components]
Name: "word"; Description: "Microsoft Word add-in"; Types: full wordonly custom
Name: "powerpoint"; Description: "Microsoft PowerPoint add-in"; Types: full powerpointonly custom

[Files]
Source: "..\..\src-csharp\OLEPackager.WordAddIn\bin\Release\OLEPackager.WordAddIn.dll"; DestDir: "{app}\Word"; Flags: ignoreversion; Components: word
Source: "..\..\src-csharp\OLEPackager.WordAddIn\bin\Release\OLEPackager.WordAddIn.dll.manifest"; DestDir: "{app}\Word"; Flags: ignoreversion; Components: word
Source: "..\..\src-csharp\OLEPackager.WordAddIn\bin\Release\OLEPackager.WordAddIn.vsto"; DestDir: "{app}\Word"; Flags: ignoreversion; Components: word
Source: "..\..\src-csharp\OLEPackager.WordAddIn\bin\Release\OLEPackager.Core.dll"; DestDir: "{app}\Word"; Flags: ignoreversion; Components: word
Source: "..\..\src-csharp\OLEPackager.WordAddIn\bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\Word"; Flags: ignoreversion; Components: word
Source: "..\..\src-csharp\OLEPackager.PowerPointAddIn\bin\Release\OLEPackager.PowerPointAddIn.dll"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion; Components: powerpoint
Source: "..\..\src-csharp\OLEPackager.PowerPointAddIn\bin\Release\OLEPackager.PowerPointAddIn.dll.manifest"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion; Components: powerpoint
Source: "..\..\src-csharp\OLEPackager.PowerPointAddIn\bin\Release\OLEPackager.PowerPointAddIn.vsto"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion; Components: powerpoint
Source: "..\..\src-csharp\OLEPackager.PowerPointAddIn\bin\Release\OLEPackager.Core.dll"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion; Components: powerpoint
Source: "..\..\src-csharp\OLEPackager.PowerPointAddIn\bin\Release\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion; Components: powerpoint

[InstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\OLEPackager.dotm"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\OLEPackagerZipTool.exe"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\OleRawDataInserter.dotm"
Type: files; Name: "{userappdata}\Microsoft\Word\STARTUP\FigurePackageZipTool.exe"
Type: filesandordirs; Name: "{app}\Assets"
Type: files; Name: "{app}\OLEPackager.WordAddIn.dll"
Type: files; Name: "{app}\OLEPackager.WordAddIn.dll.manifest"
Type: files; Name: "{app}\OLEPackager.WordAddIn.vsto"
Type: files; Name: "{app}\OLEPackager.Core.dll"
Type: files; Name: "{app}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"

[Registry]
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: string; ValueName: "Description"; ValueData: "OLE Packager C# VSTO add-in for Word"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OLE Packager"; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetWordManifestUrl}"; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: string; ValueName: "Description"; ValueData: "OLE Packager C# VSTO add-in for Word"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OLE Packager"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordAddinName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetWordManifestUrl}"; Check: IsWin64; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: string; ValueName: "Description"; ValueData: "OLE Packager C# VSTO add-in for PowerPoint"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OLE Packager"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetPowerPointManifestUrl}"; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: string; ValueName: "Description"; ValueData: "OLE Packager C# VSTO add-in for PowerPoint"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OLE Packager"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PowerPointAddinName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetPowerPointManifestUrl}"; Check: IsWin64; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\d9bb28a6-5d45-4ca3-88e2-9f4277799011"; ValueType: string; ValueName: "Url"; ValueData: "{code:GetWordTrustUrl}"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\d9bb28a6-5d45-4ca3-88e2-9f4277799011"; ValueType: string; ValueName: "PublicKey"; ValueData: "{code:GetWordPublicKey}"; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\d9bb28a6-5d45-4ca3-88e2-9f4277799011"; ValueType: string; ValueName: "Name"; ValueData: "OLE Packager for Word"; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\d9bb28a6-5d45-4ca3-88e2-9f4277799011"; ValueType: string; ValueName: "Url"; ValueData: "{code:GetWordTrustUrl}"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\d9bb28a6-5d45-4ca3-88e2-9f4277799011"; ValueType: string; ValueName: "PublicKey"; ValueData: "{code:GetWordPublicKey}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\d9bb28a6-5d45-4ca3-88e2-9f4277799011"; ValueType: string; ValueName: "Name"; ValueData: "OLE Packager for Word"; Check: IsWin64; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\a3d47cdf-9fdc-4de0-b865-a31873670f22"; ValueType: string; ValueName: "Url"; ValueData: "{code:GetPowerPointTrustUrl}"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\a3d47cdf-9fdc-4de0-b865-a31873670f22"; ValueType: string; ValueName: "PublicKey"; ValueData: "{code:GetPowerPointPublicKey}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\a3d47cdf-9fdc-4de0-b865-a31873670f22"; ValueType: string; ValueName: "Name"; ValueData: "OLE Packager for PowerPoint"; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\a3d47cdf-9fdc-4de0-b865-a31873670f22"; ValueType: string; ValueName: "Url"; ValueData: "{code:GetPowerPointTrustUrl}"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\a3d47cdf-9fdc-4de0-b865-a31873670f22"; ValueType: string; ValueName: "PublicKey"; ValueData: "{code:GetPowerPointPublicKey}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\VSTO\Security\Inclusion\a3d47cdf-9fdc-4de0-b865-a31873670f22"; ValueType: string; ValueName: "Name"; ValueData: "OLE Packager for PowerPoint"; Check: IsWin64; Components: powerpoint

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
function BuildFileUrl(ManifestPath: String): String;
begin
  StringChangeEx(ManifestPath, '\', '/', True);
  Result := 'file:///' + ManifestPath;
end;

function BuildManifestUrl(ManifestPath: String): String;
begin
  Result := BuildFileUrl(ManifestPath) + '|vstolocal';
end;

function GetWordManifestUrl(Param: String): String;
var
  ManifestPath: String;
begin
  ManifestPath := ExpandConstant('{app}\Word\{#WordAddinFile}');
  Result := BuildManifestUrl(ManifestPath);
end;

function GetPowerPointManifestUrl(Param: String): String;
var
  ManifestPath: String;
begin
  ManifestPath := ExpandConstant('{app}\PowerPoint\{#PowerPointAddinFile}');
  Result := BuildManifestUrl(ManifestPath);
end;

function GetWordTrustUrl(Param: String): String;
begin
  Result := BuildFileUrl(ExpandConstant('{app}\Word\{#WordAddinFile}'));
end;

function GetPowerPointTrustUrl(Param: String): String;
begin
  Result := BuildFileUrl(ExpandConstant('{app}\PowerPoint\{#PowerPointAddinFile}'));
end;

function ExtractPublicKey(ManifestPath: String): String;
var
  Contents: AnsiString;
  StartPosition: Integer;
  EndPosition: Integer;
  StartTag: String;
  EndTag: String;
begin
  if not LoadStringFromFile(ManifestPath, Contents) then
    RaiseException('Unable to read VSTO manifest: ' + ManifestPath);

  StartTag := '<RSAKeyValue>';
  EndTag := '</RSAKeyValue>';
  StartPosition := Pos(StartTag, String(Contents));
  EndPosition := Pos(EndTag, String(Contents));
  if (StartPosition = 0) or (EndPosition < StartPosition) then
    RaiseException('Unable to read VSTO public key: ' + ManifestPath);

  Result := Copy(String(Contents), StartPosition, EndPosition - StartPosition + Length(EndTag));
end;

function GetWordPublicKey(Param: String): String;
begin
  Result := ExtractPublicKey(ExpandConstant('{app}\Word\{#WordAddinFile}'));
end;

function GetPowerPointPublicKey(Param: String): String;
begin
  Result := ExtractPublicKey(ExpandConstant('{app}\PowerPoint\{#PowerPointAddinFile}'));
end;
