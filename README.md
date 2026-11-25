# PowerPoint PowerTools 🚀
> A curated suite of modern VSTO add-ins for Microsoft PowerPoint Desktop designed to improve productivity, layout precision, slide navigation, and canvas control.

---

## 📦 Overview of the 5 PowerTools

All 5 add-ins integrate seamlessly side-by-side into PowerPoint's **Add-Ins** tab and follow a unified design standard (Google Material Symbols Outlined, standardized 30×30 px icon dimensions, high-DPI ready, and Quick Access Toolbar compatible).

| Add-In | Version | Key Features | Ribbon Commands |
| :--- | :---: | :--- | :--- |
| **Corner Radius** | `v1.0` | Copies and applies an exact absolute corner radius (in points) across shapes, pictures, text boxes, and placeholders regardless of dimensions. | `Copy Radius`, `Apply Radius`, `Settings`, `About` |
| **Jump Points** | `v1.9.0` | IDE-style bookmark navigation: jump quickly between marked slides via the ribbon, navigator list, or QAT shortcuts. | `Jump Points` (Split), `Previous Jump Point`, `Next Jump Point`, `List`, `Settings`, `About` |
| **Pan & Cursor Zoom** | `v1.8` | Smooth canvas panning (hand tool / Pan Mode) and cursor-centered zooming directly toward the mouse pointer. | `Pan Mode`, `Zoom to Cursor`, `Settings`, `About` |
| **Stay on Slide** | `v1.1` | Prevents accidental slide changes while scrolling (hard stop at slide edges), preserves object selection, and provides an extended workspace canvas. | `Stay on Slide`, `Settings`, `About` |
| **Missing Edge Align** | `v1.2` | Precision alignment of shapes to slide edges and neighboring objects. | `EdgeAlign`, `Settings`, `About` |

---

## 🛠️ Quick Installation

Each add-in includes its own ClickOnce installer package in the [`releases/`](releases/) directory:

1. Open the desired add-in folder under `releases/` (e.g. `PowerPointStayOnSlidePointsAddIn_V1.1_RELEASE`).
2. Double-click **`Bereinigen_und_Installieren.cmd`**.
   * This cleanly closes PowerPoint, clears outdated VSTO cache entries, and runs `setup.exe`.
3. Launch PowerPoint – the add-in appears immediately in the **Add-Ins** tab.

Alternatively, ready-to-share **`.zip` archives** for every add-in are available in [`releases/`](releases/).

---

## 💻 Building from Source

The suite is structured as a monorepo and can be opened and built from a single solution:

1. **Prerequisites:**
   * Visual Studio 2022 (Community, Professional, or Enterprise)
   * Workload: *Office/SharePoint development* (VSTO)
   * .NET Framework 4.8 Developer Pack
2. **Build:**
   * Open [`PowerPointPowerTools.sln`](PowerPointPowerTools.sln) in Visual Studio 2022.
   * Press `Ctrl + Shift + B` (or choose *Build -> Build Solution*).
   * All 5 projects compile cleanly against the VSTO and Office Interop assemblies.

```powershell
# Or via Developer Command Prompt / MSBuild:
msbuild PowerPointPowerTools.sln /p:Configuration=Release
```

---

## 📂 Repository Structure

```text
PowerPoint-PowerTools/
├── .gitignore                          # Visual Studio & VSTO ignore rules
├── LICENSE                             # MIT License
├── README.md                           # Documentation & overview
├── PowerPointPowerTools.sln            # Central solution for all 5 add-ins
├── Build-Releases.ps1                  # Automated build, publish & ZIP script
├── src/
│   ├── PowerPointCornerRadiusPointsAddIn/   # Source: Corner Radius v1.0
│   ├── PowerPointJumpPointsAddIn/           # Source: Jump Points v1.9.0
│   ├── PowerPointMissingEdgeAlignAddIn/     # Source: Missing Edge Align v1.2
│   ├── PowerPointPanZoomAddIn/              # Source: Pan and Cursor Zoom v1.8
│   └── PowerPointStayOnSlidePointsAddIn/    # Source: Stay on Slide v1.1
└── releases/
    ├── PowerPointCornerRadiusPointsAddIn_V1.0_RELEASE/ (and .zip)
    ├── PowerPointJumpPointsAddIn_V2.0_RELEASE/         (and .zip)
    ├── PowerPointMissingEdgeAlignAddIn_V1.2_RELEASE/   (and .zip)
    ├── PowerPointPanZoomAddIn_V1.8_RELEASE/            (and .zip)
    └── PowerPointStayOnSlidePointsAddIn_V1.1_RELEASE/  (and .zip)
```

---

## 🎨 Design & UX Principles

* **Google Material Symbols:** Clean, modern icons with consistent visual weight.
* **Standardized Spacing:** All large buttons use a 30×30 px core icon size with 9 px optical padding on a 48×48 canvas.
* **QAT-Friendly:** All primary commands can be right-clicked and added to the *Quick Access Toolbar* for instant `Alt + [Number]` keyboard access.
* **Non-Intrusive:** Built for demanding professional workflows without interfering with native PowerPoint mechanics.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).  
© 2026 **Sperker.Training**. All rights reserved.
