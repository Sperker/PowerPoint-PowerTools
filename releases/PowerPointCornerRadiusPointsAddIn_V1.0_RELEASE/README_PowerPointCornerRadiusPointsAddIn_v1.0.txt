PowerPoint Corner Radius Points Add-In v1.0
===================================================

Overview:
PowerPoint Corner Radius Points AddIn solves a key design problem in Microsoft PowerPoint:
By default, PowerPoint calculates corner rounding proportionally relative to the shorter side of a shape.
When shapes are resized or have different aspect ratios, their corner radii look inconsistent and mismatched.

With this Add-In:
- Every shape gets the exact same visual corner radius in points (pt), regardless of width or height.
- Copy Radius: Reads the absolute corner radius of any selected rounded shape.
- Apply Radius: Computes and applies the required relative adjustment to all selected shapes simultaneously.
- Automatic Conversion: Plain right-angled rectangles are automatically converted into rounded rectangles upon applying.
- Quick Access Toolbar (QAT): Right-click either button to add it to your Quick Access Toolbar for one-click access.
- Smart State: The "Apply Radius" button is disabled (grayed out) until a radius is copied or configured in Settings.

Ribbon Controls:
- Tab: Add-Ins
- Group: Corner Radius v1.0
- Button: Copy Radius (Icon: Google Material Symbol "Rounded Corner")
- Button: Apply Radius (Icon: Google Material Symbol "Content Paste Go")
- Button: Settings (Configure audio/visual feedback and manual radius)
- Button: About (Product information, version details, and active status)

Feedback on Copy:
- Sound (System Sound): Enabled by default.
- Notification (Message Box): Disabled by default.
- Both can be configured independently in Settings.

Manual Radius:
- You can enter a custom radius directly in Settings.
- Click "Set Radius" to activate it immediately.
- No default radius is preset on clean install.

System Requirements:
- Microsoft Windows 10 / 11
- Microsoft PowerPoint 2016 / 2019 / 2021 / Microsoft 365 (Desktop version)
- .NET Framework 4.8
- Visual Studio 2010 Tools for Office Runtime (VSTO Runtime)
