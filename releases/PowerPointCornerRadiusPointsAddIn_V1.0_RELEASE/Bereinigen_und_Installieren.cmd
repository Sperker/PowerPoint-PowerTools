@echo off
chcp 65001 >nul
title PowerPoint Corner Radius Points Add-In V1.0 Installer
echo ====================================================================
echo  PowerPoint Corner Radius Points Add-In V1.0 - Saubere Installation
echo ====================================================================
echo.
echo 1. Schließe eventuell im Hintergrund laufende PowerPoint-Prozesse...
taskkill /f /im POWERPNT.EXE 2>nul
timeout /t 1 /nobreak >nul

echo 2. Bereinige veraltete Registrierungseinträge und VSTO-Metadaten...
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointCornerRadiusPointsAddIn" /f 2>nul
powershell -NoProfile -Command "Get-ChildItem 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata' -ErrorAction SilentlyContinue | ForEach-Object { if ((Get-ItemProperty $_.PsPath).friendlyName -like '*CornerRadius*') { Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue } }" 2>nul
powershell -NoProfile -Command "Get-ChildItem 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion' -ErrorAction SilentlyContinue | ForEach-Object { if ((Get-ItemProperty $_.PsPath).Url -like '*CornerRadius*') { Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue } }" 2>nul

echo.
echo 3. Starte Installation für Version 1.0...
start "" "%~dp0setup.exe"
