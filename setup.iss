[Setup]
AppName=DOCX to PPTX Converter
AppVersion=1.0.0
DefaultDirName={pf}\DOCX to PPTX
DefaultGroupName=DOCX to PPTX
OutputDir=installer
OutputBaseFilename=DOCX_to_PPTX_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=icons\app.ico

[Files]
Source: "dist\DOCX to PPTX.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "icons\*"; DestDir: "{app}\icons"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\DOCX to PPTX"; Filename: "{app}\DOCX to PPTX.exe"
Name: "{commondesktop}\DOCX to PPTX"; Filename: "{app}\DOCX to PPTX.exe"

[Run]
Filename: "{app}\DOCX to PPTX.exe"; Description: "Lancer l'application"; Flags: nowait postinstall skipifsilent
