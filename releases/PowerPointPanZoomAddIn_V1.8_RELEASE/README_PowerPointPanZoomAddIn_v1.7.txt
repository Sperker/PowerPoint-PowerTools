PowerPoint Pan & Zoom Add-in v1.7
===================================

Purpose
-------
PowerPoint Pan & Zoom Add-in adds practical pan and cursor-based zoom support to Microsoft PowerPoint Desktop on Windows.

What is new in v1.7
-------------------
v1.7 is a packaging and release-readiness update. It keeps the proven v1.6 functionality stable and adds a clearer distribution package structure, installation guide, troubleshooting guide, release checklist and license notice for the embedded Google Material Symbols icons.

Feature overview
----------------
- Pan Mode: Middle mouse button + drag to pan around a zoomed slide.
- Zoom to Cursor: Ctrl + Shift + MouseWheel to zoom toward the cursor.
- Settings: Configure zoom levels and pan sensitivity per user.
- About: Shows version, shortcuts, current settings and settings file path.
- Custom embedded Material Icons: No dependency on Office imageMso icon names.

Installation
------------
1. Close PowerPoint completely.
2. Extract the ZIP package to a local folder.
3. Run PowerPointPanZoomAddIn.vsto or setup.exe.
4. Confirm the installation prompts.
5. Open PowerPoint.
6. The add-in appears under the Add-Ins tab as "Pan & Zoom v1.7".

Recommended package contents
----------------------------
A complete v1.7 install package should contain:

- setup.exe
- PowerPointPanZoomAddIn.vsto
- Application Files\
- README_PowerPointPanZoomAddIn_v1.7.txt
- INSTALL_GUIDE_PowerPointPanZoomAddIn_v1.7.txt
- TROUBLESHOOTING_PowerPointPanZoomAddIn_v1.7.txt
- RELEASE_CHECKLIST_PowerPointPanZoomAddIn_v1.7.txt
- LICENSE_NOTICE_MaterialIcons.txt
- CHANGELOG_PowerPointPanZoomAddIn.txt

Settings location
-----------------
Settings are stored per user under:

%APPDATA%\PowerPointPanZoomAddIn\settings.ini

Icon attribution
----------------
This add-in uses Google Material Symbols / Material Icons, licensed under the Apache License 2.0.
https://fonts.google.com/icons

Notes
-----
- Windows PowerPoint Desktop only.
- Installed per user when deployed with ClickOnce.
- No admin rights are required if all required prerequisites are already available.
- Temporary invisible zoom anchors are cleaned up automatically when Zoom to Cursor is disabled or when PowerPoint closes.
