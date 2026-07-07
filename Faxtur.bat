@echo off
REM ============================================================================
REM Faxtur
REM Copyright © 2026 Frédéric Brouard
REM
REM This Source Code Form is subject to the terms of the
REM Mozilla Public License, v. 2.0.
REM If a copy of the MPL was not distributed with this file,
REM You can obtain one at https://mozilla.org/MPL/2.0/
REM ============================================================================
setlocal
cd /d "%~dp0"

REM Toujours utiliser le venv installe si present
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0main.py"
  exit /b 0
)
if exist "%~dp0.venv\Scripts\python.exe" (
  "%~dp0.venv\Scripts\python.exe" "%~dp0main.py"
  exit /b %ERRORLEVEL%
)

REM Secours : detection Windows sans supposer py/python/python3
set PYEXE=
where python >nul 2>nul && set PYEXE=python
if "%PYEXE%"=="" where py >nul 2>nul && set PYEXE=py
if "%PYEXE%"=="" where python3 >nul 2>nul && set PYEXE=python3

if "%PYEXE%"=="" (
  echo Python introuvable. Relancez INSTALLER_FACTURX_STUDIO.bat.
  pause
  exit /b 1
)

%PYEXE% -c "import pypdf" >nul 2>nul
if errorlevel 1 (
  echo Installation de pypdf pour l'utilisateur courant...
  %PYEXE% -m pip install --user pypdf
)

%PYEXE% "%~dp0main.py"
exit /b %ERRORLEVEL%
