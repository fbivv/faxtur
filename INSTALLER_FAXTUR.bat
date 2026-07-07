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
chcp 65001 >nul
title Installation Faxtur

echo.
echo ============================================
echo        Installation Faxtur
echo ============================================
echo.
echo Cet installateur va copier Faxtur dans votre profil utilisateur,
echo detecter Python automatiquement puis installer les dependances,
echo puis installer/configurer Ghostscript et veraPDF sans droits administrateur.
echo.
echo Vous choisirez aussi les dossiers de travail :
echo - Factures
echo - Facture-X
echo - A traiter
echo Les valeurs par defaut seront proposees sur le Bureau de l'utilisateur courant.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\install_faxtur.ps1"
set ERR=%ERRORLEVEL%
echo.
if not "%ERR%"=="0" (
  echo Installation terminee avec erreur. Voir les messages ci-dessus.
) else (
  echo Installation terminee.
)
pause
exit /b %ERR%
