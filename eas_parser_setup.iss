[Setup]
AppName=EAS XML Parser
AppVersion=1.0
DefaultDirName={autopf}\EASParser
DefaultGroupName=EAS XML Parser
OutputDir=.\Output
OutputBaseFilename=EASParserSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=.\icon.ico
UninstallDisplayIcon={app}\EASParser.exe
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: ".\dist\EASParser\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\EAS XML Parser"; Filename: "{app}\EASParser.exe"
Name: "{autodesktop}\EAS XML Parser"; Filename: "{app}\EASParser.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Run]
Filename: "{app}\EASParser.exe"; Description: "{cm:LaunchProgram,EAS XML Parser}"; Flags: nowait postinstall skipifsilent
