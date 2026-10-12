@echo off
chcp 65001 >nul
title PowerPoint Jump Points Add-In V2.0 Installer
echo ====================================================================
echo  PowerPoint Jump Points Add-In V2.0 - Saubere Installation
echo ====================================================================
echo.
echo 1. Schließe eventuell im Hintergrund laufende PowerPoint-Prozesse...
taskkill /f /im POWERPNT.EXE 2>nul
timeout /t 1 /nobreak >nul

echo 2. Bereinige veraltete Registrierungseinträge und VSTO-Metadaten...
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointJumpPointsAddIn" /f 2>nul
powershell -NoProfile -Command "Get-ChildItem 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata' -ErrorAction SilentlyContinue | ForEach-Object { if ((Get-ItemProperty $_.PsPath).friendlyName -like '*JumpPoint*' -or (Get-ItemProperty $_.PsPath).friendlyName -like '*Jump*Point*') { Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue } }" 2>nul
powershell -NoProfile -Command "Get-ChildItem 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion' -ErrorAction SilentlyContinue | ForEach-Object { if ((Get-ItemProperty $_.PsPath).Url -like '*JumpPoint*') { Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue } }" 2>nul

echo.
echo 3. Starte Installation für Version 2.0...
start "" "%~dp0setup.exe"
