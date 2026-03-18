#define AppVersion GetFileVersion("dist\Gerador de Certificados\Gerador de Certificados.exe")

[Setup]
AppName=Gerador de Certificados CAHIS
AppVersion={#AppVersion}
AppPublisher=Yann Dezedias
DefaultDirName={autopf}\Gerador CAHIS
DefaultGroupName=Gerador CAHIS
UninstallDisplayIcon={app}\Gerador de Certificados.exe
Compression=lzma2
SolidCompression=yes
OutputDir=.\Instalador
OutputBaseFilename=Instalador_Gerador_v{#AppVersion}
SetupIconFile=.\icone.ico
ArchitecturesInstallIn64BitMode=x64

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\Gerador de Certificados\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Gerador de Certificados"; Filename: "{app}\Gerador de Certificados.exe"
Name: "{autodesktop}\Gerador de Certificados"; Filename: "{app}\Gerador de Certificados.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Gerador de Certificados.exe"; Description: "{cm:LaunchProgram,Gerador de Certificados}"; Flags: nowait postinstall skipifsilent