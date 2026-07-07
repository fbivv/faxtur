# ============================================================================
# Faxtur
# Copyright © 2026 Frédéric Brouard
#
# This Source Code Form is subject to the terms of the
# Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file,
# You can obtain one at https://mozilla.org/MPL/2.0/
# ============================================================================
$ErrorActionPreference = "Stop"

function Write-Step($msg) { Write-Host "`n=== $msg ===" -ForegroundColor Cyan }
function Write-Ok($msg) { Write-Host "OK  $msg" -ForegroundColor Green }
function Write-Warn($msg) { Write-Host "!!  $msg" -ForegroundColor Yellow }

# Installation sans droits administrateur : dossier utilisateur
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceRoot = Split-Path -Parent $ScriptDir
$InstallDir = Join-Path $env:LOCALAPPDATA "Faxtur"
$RuntimeDir = Join-Path $InstallDir "runtime"
$Tmp = Join-Path $env:TEMP "FaxturInstall"
New-Item -ItemType Directory -Force -Path $Tmp | Out-Null

function Get-UserDesktopPath {
    try {
        $desktop = [Environment]::GetFolderPath("Desktop")
        if ($desktop -and (Test-Path $desktop)) { return $desktop }
    } catch {}
    return (Join-Path $env:USERPROFILE "Desktop")
}

function Select-InstallFolder($Title, $DefaultPath) {
    New-Item -ItemType Directory -Force -Path $DefaultPath | Out-Null
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        Add-Type -AssemblyName System.Drawing | Out-Null

        # Fenetre proprietaire invisible et toujours au premier plan.
        # Sans proprietaire, Windows peut placer la boite de choix de dossier
        # derriere la fenetre CMD/PowerShell de l'installateur.
        $owner = New-Object System.Windows.Forms.Form
        $owner.Text = "Faxtur - installation"
        $owner.StartPosition = "CenterScreen"
        $owner.Size = New-Object System.Drawing.Size(1, 1)
        $owner.ShowInTaskbar = $false
        $owner.TopMost = $true
        $owner.Opacity = 0
        $owner.Show()
        $owner.Activate()

        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Title
        $dialog.SelectedPath = $DefaultPath
        $dialog.ShowNewFolderButton = $true
        $result = $dialog.ShowDialog($owner)

        $owner.Close()
        $owner.Dispose()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $dialog.SelectedPath) {
            New-Item -ItemType Directory -Force -Path $dialog.SelectedPath | Out-Null
            return $dialog.SelectedPath
        }
    } catch {
        Write-Warn "Selection graphique indisponible, chemin par defaut utilise : $DefaultPath"
    }
    return $DefaultPath
}

function Choose-UserFolders {
    Write-Step "Choix des dossiers de travail"
    $desktop = Get-UserDesktopPath
    $defaultInput = Join-Path $desktop "Factures"
    $defaultOutput = Join-Path $desktop "Facture-X"
    $defaultTodo = Join-Path $desktop "A traiter"

    Write-Host "Vous allez choisir les dossiers utilises par Faxtur."
    Write-Host "Les valeurs proposees sont basees sur le Bureau de l'utilisateur courant : $desktop"

    $inputDir = Select-InstallFolder "Dossier des factures PDF a convertir" $defaultInput
    $outputDir = Select-InstallFolder "Dossier des Factur-X generes" $defaultOutput
    $todoDir = Select-InstallFolder "Dossier des factures a traiter manuellement" $defaultTodo

    New-Item -ItemType Directory -Force -Path $inputDir | Out-Null
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
    New-Item -ItemType Directory -Force -Path $todoDir | Out-Null

    Write-Ok "Factures : $inputDir"
    Write-Ok "Factur-X : $outputDir"
    Write-Ok "A traiter : $todoDir"

    return @{
        input = $inputDir
        output = $outputDir
        todo = $todoDir
    }
}

Write-Step "Installation de Faxtur"
Write-Host "Source      : $SourceRoot"
Write-Host "Destination : $InstallDir"

function Invoke-Download($Url, $OutFile) {
    Write-Host "Telechargement : $Url"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
}

function Get-PythonExe {
    $candidates = @(
        (Get-Command python.exe -ErrorAction SilentlyContinue),
        (Get-Command python -ErrorAction SilentlyContinue),
        (Get-Command py.exe -ErrorAction SilentlyContinue),
        (Get-Command py -ErrorAction SilentlyContinue),
        (Get-Command python3.exe -ErrorAction SilentlyContinue),
        (Get-Command python3 -ErrorAction SilentlyContinue)
    )
    foreach ($cmd in $candidates) {
        if ($cmd -and $cmd.Source) {
            try {
                & $cmd.Source --version | Out-Null
                if ($LASTEXITCODE -eq 0) { return $cmd.Source }
            } catch {}
        }
    }
    $paths = @(
        "$env:ProgramFiles\Python313\python.exe",
        "$env:ProgramFiles\Python312\python.exe",
        "$env:LocalAppData\Programs\Python\Python313\python.exe",
        "$env:LocalAppData\Programs\Python\Python312\python.exe"
    )
    foreach ($x in $paths) {
        if (Test-Path $x) { return $x }
    }
    return $null
}

function Install-PythonIfNeeded {
    Write-Step "Verification Python"
    $python = Get-PythonExe
    if ($python) { Write-Ok "Python trouve : $python"; return $python }

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) { throw "Python absent et winget indisponible. Installe Python puis relance." }

    Write-Warn "Python absent. Installation Python via winget..."
    winget install -e --id Python.Python.3.13 --scope user --accept-source-agreements --accept-package-agreements

    $python = Get-PythonExe
    if (-not $python) {
        throw "Python installe mais non detecte. Ferme/reouvre la session puis relance."
    }
    return $python
}

function Find-GhostscriptExe {
    $candidates = @(
        (Join-Path $RuntimeDir "ghostscript\bin\gswin64c.exe"),
        (Join-Path $RuntimeDir "ghostscript\gswin64c.exe")
    )
    foreach ($p in $candidates) { if (Test-Path $p) { return $p } }

    $gsRoot = Join-Path $env:ProgramFiles "gs"
    if (Test-Path $gsRoot) {
        $found = Get-ChildItem $gsRoot -Recurse -Filter gswin64c.exe -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    $cmd = Get-Command gswin64c.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Install-Ghostscript {
    Write-Step "Ghostscript"
    $existing = Find-GhostscriptExe
    if ($existing) { Write-Ok "Ghostscript trouve : $existing"; return $existing }

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($winget) {
        try {
            Write-Warn "Ghostscript absent. Installation via winget..."
            winget install -e --id ArtifexSoftware.Ghostscript --scope user --accept-source-agreements --accept-package-agreements
            $existing = Find-GhostscriptExe
            if ($existing) { Write-Ok "Ghostscript installe : $existing"; return $existing }
        } catch { Write-Warn "Installation winget Ghostscript echouee, tentative par telechargement." }
    }

    $url = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10071/gs10071w64.exe"
    $exe = Join-Path $Tmp "ghostscript.exe"
    Invoke-Download $url $exe
    $dest = Join-Path $RuntimeDir "ghostscript"
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Start-Process -FilePath $exe -ArgumentList @("/S", "/D=$dest") -Wait
    Start-Sleep -Seconds 2
    $gs = Find-GhostscriptExe
    if (-not $gs) { throw "Ghostscript non detecte apres installation." }
    Write-Ok "Ghostscript installe : $gs"
    return $gs
}

function Find-JavaExe {
    $found = Get-ChildItem (Join-Path $RuntimeDir "java") -Recurse -Filter java.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) { return $found.FullName }
    $cmd = Get-Command java.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Install-JavaRuntime {
    Write-Step "Java Runtime pour veraPDF"
    $java = Find-JavaExe
    if ($java) { Write-Ok "Java trouve : $java"; return $java }
    $url = "https://api.adoptium.net/v3/binary/latest/21/ga/windows/x64/jre/hotspot/normal/eclipse"
    $zip = Join-Path $Tmp "jre.zip"
    Invoke-Download $url $zip
    $javaDir = Join-Path $RuntimeDir "java"
    if (Test-Path $javaDir) { Remove-Item $javaDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $javaDir | Out-Null
    Expand-Archive -Path $zip -DestinationPath $javaDir -Force
    $java = Find-JavaExe
    if (-not $java) { throw "Java non detecte apres extraction." }
    Write-Ok "Java installe : $java"
    return $java
}

function Find-VeraPdf {
    $candidates = @(
        (Join-Path $RuntimeDir "veraPDF\verapdf.bat"),
        (Join-Path $RuntimeDir "veraPDF\bin\verapdf.bat"),
        (Join-Path $RuntimeDir "veraPDF\verapdf.exe"),
        (Join-Path $RuntimeDir "veraPDF\bin\verapdf.exe"),
        (Join-Path $env:ProgramFiles "veraPDF\verapdf.bat"),
        (Join-Path $env:ProgramFiles "vera\verapdf.bat")
    )
    foreach ($p in $candidates) { if (Test-Path $p) { return $p } }
    $cmd = Get-Command verapdf.bat -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Install-VeraPDF($JavaExe) {
    Write-Step "veraPDF"
    $existing = Find-VeraPdf
    if ($existing) { Write-Ok "veraPDF trouve : $existing"; return $existing }

    $url = "https://software.verapdf.org/releases/verapdf-installer.zip"
    $zip = Join-Path $Tmp "verapdf-installer.zip"
    $extract = Join-Path $Tmp "verapdf-installer"
    Invoke-Download $url $zip
    if (Test-Path $extract) { Remove-Item $extract -Recurse -Force }
    Expand-Archive -Path $zip -DestinationPath $extract -Force

    $bat = Get-ChildItem $extract -Recurse -Include "vera-install.bat","verapdf-install.bat" | Select-Object -First 1
    if (-not $bat) { throw "Installateur veraPDF introuvable dans le ZIP." }

    $autoTemplate = Join-Path $InstallDir "installer\verapdf-auto-install.xml.template"
    $auto = Join-Path $Tmp "verapdf-auto-install.xml"
    $installPath = (Join-Path $RuntimeDir "veraPDF") -replace "\\", "/"
    (Get-Content $autoTemplate -Raw).Replace("__INSTALL_PATH__", $installPath) | Set-Content -Path $auto -Encoding UTF8

    $env:JAVA_HOME = Split-Path -Parent (Split-Path -Parent $JavaExe)
    $env:PATH = (Split-Path -Parent $JavaExe) + ";" + $env:PATH
    Start-Process -FilePath $bat.FullName -ArgumentList @($auto) -WorkingDirectory $bat.Directory.FullName -Wait
    Start-Sleep -Seconds 2
    $vp = Find-VeraPdf
    if (-not $vp) { throw "veraPDF non detecte apres installation." }
    Write-Ok "veraPDF installe : $vp"
    return $vp
}

function Configure-AppSettings($GsPath, $VeraPath, $JavaPath, $UserFolders) {
    Write-Step "Configuration"
    $settingsPath = Join-Path $InstallDir "config\settings.json"
    New-Item -ItemType Directory -Force -Path (Join-Path $InstallDir "runtime\icc") | Out-Null
    $iccPath = Join-Path $InstallDir "runtime\icc\sRGB.icc"
    if (-not (Test-Path $iccPath)) {
        $engineIcc = Join-Path $InstallDir "engine\icc\sRGB.icc"
        if (Test-Path $engineIcc) { Copy-Item $engineIcc $iccPath -Force }
    }
    $settings = @{
        last_company = ""
        default_input_dir = $UserFolders.input
        default_output_dir = $UserFolders.output
        default_todo_dir = $UserFolders.todo
        verapdf_path = $VeraPath
        ghostscript_path = $GsPath
        icc_profile_path = $iccPath
        java_path = $JavaPath
        prefer_embedded_tools = $true
    }
    New-Item -ItemType Directory -Force -Path (Split-Path $settingsPath) | Out-Null
    $settings | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsPath -Encoding UTF8
    Write-Ok "Chemins enregistres dans config\settings.json"
}

function Create-Shortcuts {
    Write-Step "Raccourcis"
    $launcher = Join-Path $InstallDir "Faxtur.vbs"
    if (-not (Test-Path $launcher)) { $launcher = Join-Path $InstallDir "Faxtur.bat" }
    $shell = New-Object -ComObject WScript.Shell
    $desktop = [Environment]::GetFolderPath("Desktop")
    $shortcut = $shell.CreateShortcut((Join-Path $desktop "Faxtur.lnk"))
    $shortcut.TargetPath = $launcher
    $shortcut.WorkingDirectory = $InstallDir
    $icon = Join-Path $InstallDir "resources\Faxtur.ico"
    if (Test-Path $icon) { $shortcut.IconLocation = $icon }
    else { $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,70" }
    $shortcut.Save()
    $startDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Faxtur"
    New-Item -ItemType Directory -Force -Path $startDir | Out-Null
    $shortcut2 = $shell.CreateShortcut((Join-Path $startDir "Faxtur.lnk"))
    $shortcut2.TargetPath = $launcher
    $shortcut2.WorkingDirectory = $InstallDir
    if (Test-Path $icon) { $shortcut2.IconLocation = $icon }
    else { $shortcut2.IconLocation = "$env:SystemRoot\System32\shell32.dll,70" }
    $shortcut2.Save()
    Write-Ok "Raccourcis crees"
}

try {
    Write-Step "Copie de l'application"
    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
    Get-ChildItem $SourceRoot -Force | ForEach-Object {
        if ($_.Name -notin @("logs", "temp", "__pycache__")) {
            Copy-Item $_.FullName -Destination $InstallDir -Recurse -Force
        }
    }
    New-Item -ItemType Directory -Force -Path (Join-Path $InstallDir "logs") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $InstallDir "temp") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $RuntimeDir "ghostscript") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $RuntimeDir "veraPDF") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $RuntimeDir "java") | Out-Null
    Write-Ok "Application copiee"

    $py = Install-PythonIfNeeded
    Write-Step "Dependances Python"
    Push-Location $InstallDir
    if (-not (Test-Path ".venv\Scripts\python.exe")) {
        & $py -m venv .venv
    }
    & ".\.venv\Scripts\python.exe" -m pip install --upgrade pip
    & ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
    Pop-Location
    Write-Ok "Dependances installees avec py/venv"

    $gs = Install-Ghostscript
    $java = Install-JavaRuntime
    $vp = Install-VeraPDF $java
    $userFolders = Choose-UserFolders
    Configure-AppSettings $gs $vp $java $userFolders
    Create-Shortcuts

    Write-Step "Installation terminee"
    Write-Host "Faxtur : $InstallDir" -ForegroundColor Green
    Write-Host "Ghostscript    : $gs"
    Write-Host "veraPDF        : $vp"
    Write-Host "Java           : $java"
    Write-Host "Factures       : $($userFolders.input)"
    Write-Host "Factur-X       : $($userFolders.output)"
    Write-Host "A traiter      : $($userFolders.todo)"
    Write-Host "Lanceur        : Bureau > Faxtur"
    exit 0
}
catch {
    Write-Host "`nERREUR INSTALLATION : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
    exit 1
}
