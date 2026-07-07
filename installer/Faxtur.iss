; ============================================================================
; Faxtur
; Copyright © 2026 Frédéric Brouard
;
; This Source Code Form is subject to the terms of the
; Mozilla Public License, v. 2.0.
; If a copy of the MPL was not distributed with this file,
; You can obtain one at https://mozilla.org/MPL/2.0/
; ============================================================================
#define MyAppName "Faxtur"
#define MyAppVersion "1.0.1"
#define MyAppExeName "Faxtur.exe"

[Setup]
AppId={{8D4F0D75-9F0A-4C33-9B1C-FAXTUR}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\Faxtur
DefaultGroupName={#MyAppName}
OutputBaseFilename=Faxtur_Setup_{#MyAppVersion}
SetupIconFile=resources\Faxtur.ico
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64

[Files]
Source: "dist\Faxtur\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Faxtur"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\resources\Faxtur.ico"
Name: "{commondesktop}\Faxtur"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\resources\Faxtur.ico"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le Bureau"; GroupDescription: "Raccourcis :"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Lancer Faxtur"; Flags: nowait postinstall skipifsilent
