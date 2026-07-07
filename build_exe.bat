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
set PYEXE=
where python >nul 2>nul && set PYEXE=python
if "%PYEXE%"=="" where py >nul 2>nul && set PYEXE=py
if "%PYEXE%"=="" where python3 >nul 2>nul && set PYEXE=python3
if "%PYEXE%"=="" (
  echo Python introuvable.
  pause
  exit /b 1
)
%PYEXE% -m pip install -r requirements.txt
%PYEXE% -m pip install pyinstaller
%PYEXE% -m PyInstaller --noconfirm --windowed --name Faxtur --icon "resources\Faxtur.ico" --add-data "resources;resources" --add-data "engine;engine" --add-data "companies;companies" --add-data "config;config" --add-data "runtime;runtime" main.py
pause
