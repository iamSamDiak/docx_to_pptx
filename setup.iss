[Setup]
AppName=EcoDim Converter
AppVersion=1.0.0
DefaultDirName={pf}\EcoDim
DefaultGroupName=EcoDim
OutputDir=installer
OutputBaseFilename=EcoDim_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=assets\app.ico

[Files]
Source: "dist\EcoDim.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\EcoDim"; Filename: "{app}\EcoDim.exe"
Name: "{commondesktop}\EcoDim"; Filename: "{app}\EcoDim.exe"

[Run]
Filename: "{app}\EcoDim.exe"; Description: "Lancer l'application"; Flags: nowait postinstall skipifsilent
