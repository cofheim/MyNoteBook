[Setup]
AppName=MyNoteBook
AppVersion=1.0
DefaultDirName={autopf}\MyNoteBook
DefaultGroupName=MyNoteBook
OutputBaseFilename=MyNoteBookSetup
Compression=lzma
SolidCompression=yes
OutputDir=Installer

[Files]
Source: "Release\NBapp.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "Release\contacts.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\MyNoteBook"; Filename: "{app}\NBapp.exe"
Name: "{commondesktop}\MyNoteBook"; Filename: "{app}\NBapp.exe" 