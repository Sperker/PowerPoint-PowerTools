## PowerPoint Jump Points Add-in v1.9.0

### Purpose
PowerPoint Jump Points Add-in allows users to define invisible jump points on slides and quickly navigate between them in large Microsoft PowerPoint Desktop decks on Windows.

### What is new in v1.9.0
v1.9.0 is a release-readiness update and a Navigator usability cleanup.

Changes:
- Replaced the previous `Current Jump Point` Navigator shortcut with `Previous Jump Point`.
- The Navigator now has two clearly separated rows:
  - `Jump Point navigation:` with `Previous Jump Point` and `Next Jump Point`.
  - `Navigator actions:` with `Go To`, `Rename`, `Delete`, `Delete All`, and `Refresh List`.
- The Navigator dialog is narrower and more side-panel friendly.
- The Jump Point table keeps the important slide columns visible first and exposes the remaining details through horizontal scrolling.
- `Keep this window on top` remains visible in the top area of the dialog.
- Release documentation files were added for installation, troubleshooting, release checks, icon licensing, and changelog.

### Feature overview
- Toggle Jump Point: Adds or removes a jump point on the current slide.
- Previous Jump Point: Jumps to the previous defined jump point.
- Next Jump Point: Jumps to the next defined jump point.
- List / Navigator: Opens a modeless Navigator window.
- Navigator actions: Go To, Rename, Delete, Delete All, Refresh List.
- Settings: Configure toggle messages, toggle sound, and Navigator confirmations.
- About: Shows version, jump point count, and settings status.
- Custom embedded Material Icons: No dependency on Office imageMso icon names.

### How jump points are stored
Jump points are stored directly in the PowerPoint deck using slide tags:
- `PPT_JUMPPOINT = TRUE`
- `PPT_JUMPPOINT_LABEL` for optional custom Navigator names
- `PPT_JUMPPOINT_CREATED`
- `PPT_JUMPPOINT_VERSION`

### Installation
- Close PowerPoint completely.
- Extract the full ZIP package to a local folder.
- Run `PowerPointJumpPointsAddIn.vsto` or `setup.exe` from the ClickOnce publish package.
- Confirm the installation prompts.
- Open PowerPoint.
- The add-in appears under the Add-Ins tab as `Jump Points v1.9.0`.

### Recommended package contents
A complete v1.9.0 install package should contain:
- setup.exe
- PowerPointJumpPointsAddIn.vsto
- Application Files\
- README_PowerPointJumpPointsAddIn_v1.9.0.txt
- INSTALL_GUIDE_PowerPointJumpPointsAddIn_v1.9.0.txt
- TROUBLESHOOTING_PowerPointJumpPointsAddIn_v1.9.0.txt
- RELEASE_CHECKLIST_PowerPointJumpPointsAddIn_v1.9.0.txt
- LICENSE_NOTICE_MaterialIcons.txt
- CHANGELOG_PowerPointJumpPointsAddIn.txt

### Settings location
Settings are stored per user under:
`HKCU\Software\PowerPointJumpPointsAddIn`

### Icon attribution
This add-in uses Google Material Symbols / Material Icons, licensed under Apache License 2.0.
Source: https://fonts.google.com/icons

### Notes
- Windows PowerPoint Desktop only.
- Installed per user when deployed with ClickOnce.
- No admin rights are required if all required prerequisites are already available.
- Jump points stay inside the deck because they are stored as slide tags.
