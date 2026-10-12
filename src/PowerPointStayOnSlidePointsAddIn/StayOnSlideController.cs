using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal sealed class StayOnSlideController : IDisposable
    {
        private readonly PowerPoint.Application _app;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly MouseWheelHook _mouseWheelHook;
        private readonly KeyboardHook _keyboardHook;
        private readonly CanvasWorkspaceManager _workspaceManager;
        private StayOnSlideSettings _settings;

        private int _lockedSlideIndex = 1;
        private Point _lastLeftClickScreenPoint;
        private DateTime _lastLeftClickTime = DateTime.MinValue;
        private bool _isInternalNavigating;
        private bool _isSavingOrPrinting;
        private List<string> _lastSelectedShapeNames = new List<string>();

        public bool StayOnSlideEnabled
        {
            get => _settings.StayOnSlideEnabled;
            set
            {
                if (_settings.StayOnSlideEnabled != value)
                {
                    _settings.StayOnSlideEnabled = value;
                    _settings.Save();

                    if (value)
                    {
                        SyncCurrentSlideIndex();
                        EnsureCurrentSlideAnchors();
                    }
                    else
                    {
                        CleanupAllAnchors();
                    }
                }
            }
        }

        public StayOnSlideSettings Settings => _settings;

        public StayOnSlideController(PowerPoint.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _settings = StayOnSlideSettings.Load();
            _workspaceManager = new CanvasWorkspaceManager();

            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 250;
            _timer.Tick += Timer_Tick;

            _mouseWheelHook = new MouseWheelHook();
            _mouseWheelHook.MouseWheel += MouseWheelHook_MouseWheel;
            _mouseWheelHook.LeftButtonDown += MouseWheelHook_LeftButtonDown;

            _keyboardHook = new KeyboardHook();
            _keyboardHook.KeyDown += KeyboardHook_KeyDown;

            AttachAppEvents();
        }

        public void Start()
        {
            _mouseWheelHook.Start();
            _keyboardHook.Start();
            _timer.Start();

            SyncCurrentSlideIndex();
            if (StayOnSlideEnabled && _settings.EnableVirtualCanvas)
            {
                EnsureCurrentSlideAnchors();
            }
        }

        public void Dispose()
        {
            try
            {
                DetachAppEvents();
            }
            catch { }

            try
            {
                _timer.Stop();
                _timer.Tick -= Timer_Tick;
                _timer.Dispose();
            }
            catch { }

            try
            {
                _mouseWheelHook.MouseWheel -= MouseWheelHook_MouseWheel;
                _mouseWheelHook.LeftButtonDown -= MouseWheelHook_LeftButtonDown;
                _mouseWheelHook.Dispose();
            }
            catch { }

            try
            {
                _keyboardHook.KeyDown -= KeyboardHook_KeyDown;
                _keyboardHook.Dispose();
            }
            catch { }

            try
            {
                CleanupAllAnchors();
            }
            catch { }
        }

        private void AttachAppEvents()
        {
            try
            {
                _app.SlideSelectionChanged += App_SlideSelectionChanged;
                _app.WindowSelectionChange += App_WindowSelectionChange;
                _app.PresentationBeforeSave += App_PresentationBeforeSave;
                _app.PresentationBeforeClose += App_PresentationBeforeClose;
                _app.PresentationPrint += App_PresentationPrint;
            }
            catch { }
        }

        private void DetachAppEvents()
        {
            try
            {
                _app.SlideSelectionChanged -= App_SlideSelectionChanged;
                _app.WindowSelectionChange -= App_WindowSelectionChange;
                _app.PresentationBeforeSave -= App_PresentationBeforeSave;
                _app.PresentationBeforeClose -= App_PresentationBeforeClose;
                _app.PresentationPrint -= App_PresentationPrint;
            }
            catch { }
        }

        private void MouseWheelHook_LeftButtonDown(object sender, GlobalMouseClickEventArgs e)
        {
            _lastLeftClickScreenPoint = e.ScreenPoint;
            _lastLeftClickTime = DateTime.UtcNow;
        }

        private void MouseWheelHook_MouseWheel(object sender, GlobalMouseWheelEventArgs e)
        {
            try
            {
                if (!StayOnSlideEnabled)
                    return;

                // If Ctrl is held down, the user is zooming (native PowerPoint zoom or PanZoomAddIn)
                // Pass through completely so zoom works smoothly!
                if (NativeMethods.IsCtrlDown())
                    return;

                // If Shift is held down, the user is scrolling horizontally
                // Pass through so horizontal scrolling works natively!
                if (NativeMethods.IsShiftDown())
                    return;

                if (!IsPowerPointForeground() && !IsCursorOverPowerPoint(e.ScreenPoint))
                    return;

                var wnd = GetActiveWindow();
                if (wnd == null)
                    return;

                // If cursor is over the Left Thumbnails Sidebar, let PowerPoint handle it natively!
                if (IsCursorOverThumbnailsPane(wnd, e.ScreenPoint))
                {
                    return;
                }

                // If cursor is over the Slide Canvas, intercept and control the scroll
                if (IsCursorOverSlideCanvas(wnd, e.ScreenPoint))
                {
                    e.Handled = true;
                    HandleCanvasScroll(wnd, e.Delta);
                }
            }
            catch
            {
            }
        }

        private void HandleCanvasScroll(PowerPoint.DocumentWindow wnd, int delta)
        {
            try
            {
                var pres = wnd.Presentation;
                if (pres == null) return;

                IntPtr docHwnd = new IntPtr(wnd.HWND);
                NativeMethods.RECT mdiRect;
                if (!NativeMethods.GetWindowRect(docHwnd, out mdiRect))
                {
                    IntPtr pptHwnd = new IntPtr(_app.HWND);
                    NativeMethods.TryGetMdiClassRect(pptHwnd, out mdiRect);
                }

                int padding = _settings.EnableVirtualCanvas ? _settings.CanvasPaddingPoints : 0;
                float slideHeight = pres.PageSetup.SlideHeight;
                float topCanvasY = -padding;
                float bottomCanvasY = slideHeight + padding;

                float topScreen = wnd.PointsToScreenPixelsY(topCanvasY);
                float bottomScreen = wnd.PointsToScreenPixelsY(bottomCanvasY);

                int lines = _settings.WheelScrollStepLines <= 0 ? 3 : _settings.WheelScrollStepLines;

                const int bottomChromeMargin = 160;
                const int topChromeMargin = 90;

                if (delta < 0)
                {
                    // Scrolling Down
                    if (bottomScreen <= mdiRect.Bottom - bottomChromeMargin)
                    {
                        // Clean hard stop at the bottom of the canvas!
                        return;
                    }

                    wnd.SmallScroll(Down: lines, Up: 0, ToRight: 0, ToLeft: 0);
                }
                else if (delta > 0)
                {
                    // Scrolling Up
                    if (topScreen >= mdiRect.Top + topChromeMargin)
                    {
                        // Clean hard stop at the top of the canvas!
                        return;
                    }

                    wnd.SmallScroll(Down: 0, Up: lines, ToRight: 0, ToLeft: 0);
                }

                // Guard check: did an overshoot jump slides?
                int currentSlideIndex = GetCurrentSlideIndex(wnd);
                if (currentSlideIndex != _lockedSlideIndex && currentSlideIndex > 0 && _lockedSlideIndex > 0)
                {
                    RestoreLockedSlide(wnd, currentSlideIndex > _lockedSlideIndex);
                }
            }
            catch
            {
            }
        }

        private void KeyboardHook_KeyDown(object sender, GlobalKeyEventArgs e)
        {
            try
            {
                if (!StayOnSlideEnabled || !_settings.BlockKeyboardSlideNavigation)
                    return;

                if (!IsPowerPointForeground())
                    return;

                var wnd = GetActiveWindow();
                if (wnd == null)
                    return;

                Point cursorPos = NativeMethods.GetCursorPosition();
                if (IsCursorOverThumbnailsPane(wnd, cursorPos))
                {
                    // User is actively in thumbnail pane
                    return;
                }

                if (e.VirtualKeyCode == NativeMethods.VK_NEXT) // Page Down
                {
                    e.Handled = true;
                    HandleCanvasScroll(wnd, -120);
                }
                else if (e.VirtualKeyCode == NativeMethods.VK_PRIOR) // Page Up
                {
                    e.Handled = true;
                    HandleCanvasScroll(wnd, +120);
                }
            }
            catch
            {
            }
        }

        private void App_SlideSelectionChanged(PowerPoint.SlideRange range)
        {
            if (_isInternalNavigating || _isSavingOrPrinting)
                return;

            try
            {
                var wnd = GetActiveWindow();
                if (wnd == null)
                    return;

                int newSlideIndex = (range != null && range.Count >= 1) ? range[1].SlideIndex : GetCurrentSlideIndex(wnd);
                if (newSlideIndex <= 0)
                    return;

                if (!StayOnSlideEnabled)
                {
                    _lockedSlideIndex = newSlideIndex;
                    return;
                }

                // Check if user intentionally clicked in thumbnail sidebar
                Point currentCursor = NativeMethods.GetCursorPosition();
                bool recentThumbnailClick = (DateTime.UtcNow - _lastLeftClickTime < TimeSpan.FromMilliseconds(900))
                    && IsCursorOverThumbnailsPane(wnd, _lastLeftClickScreenPoint);

                bool cursorCurrentlyInThumbnails = IsCursorOverThumbnailsPane(wnd, currentCursor);

                if (recentThumbnailClick || cursorCurrentlyInThumbnails)
                {
                    // Intentional slide change via sidebar!
                    _lockedSlideIndex = newSlideIndex;
                    if (_settings.EnableVirtualCanvas)
                    {
                        EnsureCurrentSlideAnchors();
                    }
                    return;
                }

                // Unintended slide change (e.g. scroll boundary cross, key press) -> Revert!
                if (_lockedSlideIndex > 0 && newSlideIndex != _lockedSlideIndex)
                {
                    RestoreLockedSlide(wnd, newSlideIndex > _lockedSlideIndex);
                }
            }
            catch
            {
            }
        }

        private void App_WindowSelectionChange(PowerPoint.Selection sel)
        {
            if (_isInternalNavigating)
                return;

            try
            {
                if (sel != null && (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes || sel.Type == PowerPoint.PpSelectionType.ppSelectionText))
                {
                    PowerPoint.ShapeRange range = null;
                    try { range = sel.ShapeRange; } catch { }

                    if (range != null && range.Count > 0)
                    {
                        var list = new List<string>();
                        for (int i = 1; i <= range.Count; i++)
                        {
                            try
                            {
                                string name = range[i].Name;
                                if (!string.IsNullOrEmpty(name))
                                    list.Add(name);
                            }
                            catch { }
                        }

                        if (list.Count > 0)
                        {
                            _lastSelectedShapeNames = list;
                            return;
                        }
                    }
                }

                if (sel != null && (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone || sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides))
                {
                    _lastSelectedShapeNames.Clear();
                }
            }
            catch { }
        }

        private void RestoreSelection(PowerPoint.DocumentWindow wnd)
        {
            try
            {
                if (_lastSelectedShapeNames == null || _lastSelectedShapeNames.Count == 0)
                    return;

                var slide = GetCurrentSlide(wnd);
                if (slide == null)
                    return;

                bool first = true;
                foreach (string name in _lastSelectedShapeNames)
                {
                    try
                    {
                        var shape = slide.Shapes[name];
                        if (shape != null)
                        {
                            shape.Select(first ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                            first = false;
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void RestoreLockedSlide(PowerPoint.DocumentWindow wnd, bool wasMovingDown = false)
        {
            if (_isInternalNavigating || _lockedSlideIndex <= 0)
                return;

            _isInternalNavigating = true;
            try
            {
                wnd.View.GotoSlide(_lockedSlideIndex);

                if (wasMovingDown)
                {
                    try
                    {
                        var pres = wnd.Presentation;
                        if (pres != null)
                        {
                            int padding = _settings.EnableVirtualCanvas ? _settings.CanvasPaddingPoints : 0;
                            float slideHeight = pres.PageSetup.SlideHeight;
                            float bottomCanvasY = slideHeight + padding;
                            wnd.ScrollIntoView(0, bottomCanvasY - 50, 50, 50, Office.MsoTriState.msoFalse);
                        }
                    }
                    catch { }
                }

                RestoreSelection(wnd);
            }
            catch { }
            finally
            {
                _isInternalNavigating = false;
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!StayOnSlideEnabled || _isSavingOrPrinting)
                    return;

                var wnd = GetActiveWindow();
                if (wnd == null)
                    return;

                if (_settings.EnableVirtualCanvas)
                {
                    var slide = GetCurrentSlide(wnd);
                    if (slide != null)
                    {
                        _workspaceManager.EnsureAnchors(slide, _settings.CanvasPaddingPoints);
                    }
                }
            }
            catch
            {
            }
        }

        private void App_PresentationBeforeSave(PowerPoint.Presentation pres, ref bool cancel)
        {
            _isSavingOrPrinting = true;
            try
            {
                _workspaceManager.RemoveAllAnchors(pres);
            }
            catch { }
            finally
            {
                _isSavingOrPrinting = false;
            }
        }

        private void App_PresentationBeforeClose(PowerPoint.Presentation pres, ref bool cancel)
        {
            try
            {
                _workspaceManager.RemoveAllAnchors(pres);
            }
            catch { }
        }

        private void App_PresentationPrint(PowerPoint.Presentation pres)
        {
            _isSavingOrPrinting = true;
            try
            {
                _workspaceManager.RemoveAllAnchors(pres);
            }
            catch { }
            finally
            {
                _isSavingOrPrinting = false;
            }
        }

        public void CleanupAllAnchors()
        {
            try
            {
                var pres = _app.ActivePresentation;
                if (pres != null)
                {
                    _workspaceManager.RemoveAllAnchors(pres);
                }
            }
            catch { }
        }

        public void EnsureCurrentSlideAnchors()
        {
            try
            {
                var wnd = GetActiveWindow();
                if (wnd == null) return;
                var slide = GetCurrentSlide(wnd);
                if (slide != null && _settings.EnableVirtualCanvas)
                {
                    _workspaceManager.EnsureAnchors(slide, _settings.CanvasPaddingPoints);
                }
            }
            catch { }
        }

        public void SyncCurrentSlideIndex()
        {
            try
            {
                var wnd = GetActiveWindow();
                if (wnd != null)
                {
                    int index = GetCurrentSlideIndex(wnd);
                    if (index > 0)
                        _lockedSlideIndex = index;
                }
            }
            catch { }
        }

        private PowerPoint.DocumentWindow GetActiveWindow()
        {
            try
            {
                if (_app == null) return null;
                return _app.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }

        private PowerPoint.Slide GetCurrentSlide(PowerPoint.DocumentWindow wnd)
        {
            try
            {
                if (wnd.Selection != null && wnd.Selection.SlideRange != null && wnd.Selection.SlideRange.Count >= 1)
                    return wnd.Selection.SlideRange[1];
            }
            catch { }

            try
            {
                return wnd.View.Slide as PowerPoint.Slide;
            }
            catch
            {
                return null;
            }
        }

        private int GetCurrentSlideIndex(PowerPoint.DocumentWindow wnd)
        {
            try
            {
                var slide = GetCurrentSlide(wnd);
                return slide != null ? slide.SlideIndex : 0;
            }
            catch
            {
                return 0;
            }
        }

        private bool IsCursorOverThumbnailsPane(PowerPoint.DocumentWindow wnd, Point cursorScreen)
        {
            try
            {
                if (wnd == null || wnd.Application == null) return false;
                IntPtr pptHwnd = new IntPtr(wnd.Application.HWND);
                NativeMethods.RECT rect;
                if (!NativeMethods.TryGetMdiClassRect(pptHwnd, out rect)) return false;

                if (cursorScreen.X < rect.Left || cursorScreen.X > rect.Right ||
                    cursorScreen.Y < rect.Top || cursorScreen.Y > rect.Bottom)
                {
                    return false;
                }

                int splitPercent = wnd.SplitHorizontal;
                int splitWidth = (splitPercent > 0 && splitPercent < 50)
                    ? (int)((rect.Right - rect.Left) * (splitPercent / 100.0f))
                    : 240;

                int splitX = rect.Left + splitWidth;
                return cursorScreen.X < splitX;
            }
            catch
            {
                return false;
            }
        }

        private bool IsCursorOverSlideCanvas(PowerPoint.DocumentWindow wnd, Point cursorScreen)
        {
            try
            {
                if (wnd == null || wnd.Application == null) return false;
                IntPtr pptHwnd = new IntPtr(wnd.Application.HWND);
                NativeMethods.RECT rect;
                if (!NativeMethods.TryGetMdiClassRect(pptHwnd, out rect)) return false;

                if (cursorScreen.X < rect.Left || cursorScreen.X > rect.Right ||
                    cursorScreen.Y < rect.Top || cursorScreen.Y > rect.Bottom)
                {
                    return false;
                }

                int splitPercent = wnd.SplitHorizontal;
                int splitWidth = (splitPercent > 0 && splitPercent < 50)
                    ? (int)((rect.Right - rect.Left) * (splitPercent / 100.0f))
                    : 240;

                int splitX = rect.Left + splitWidth;
                return cursorScreen.X >= splitX;
            }
            catch
            {
                return false;
            }
        }

        private bool IsPowerPointForeground()
        {
            try
            {
                IntPtr fg = NativeMethods.GetForegroundWindow();
                if (fg == IntPtr.Zero) return false;

                uint pid;
                NativeMethods.GetWindowThreadProcessId(fg, out pid);

                var proc = System.Diagnostics.Process.GetProcessById((int)pid);
                return proc.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private bool IsCursorOverPowerPoint(Point screenPoint)
        {
            try
            {
                NativeMethods.POINT p = new NativeMethods.POINT { X = screenPoint.X, Y = screenPoint.Y };
                IntPtr hwnd = NativeMethods.WindowFromPoint(p);
                if (hwnd == IntPtr.Zero) return false;

                IntPtr root = NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT);
                if (root == IntPtr.Zero) root = hwnd;

                uint pid;
                NativeMethods.GetWindowThreadProcessId(root, out pid);

                var proc = System.Diagnostics.Process.GetProcessById((int)pid);
                return proc.ProcessName.Equals("POWERPNT", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        public void ShowSettings()
        {
            try
            {
                StayOnSlideSettings editable = new StayOnSlideSettings
                {
                    StayOnSlideEnabled = _settings.StayOnSlideEnabled,
                    EnableVirtualCanvas = _settings.EnableVirtualCanvas,
                    CanvasPaddingPoints = _settings.CanvasPaddingPoints,
                    BlockKeyboardSlideNavigation = _settings.BlockKeyboardSlideNavigation,
                    WheelScrollStepLines = _settings.WheelScrollStepLines
                };

                using (SettingsForm form = new SettingsForm(editable))
                {
                    DialogResult result = form.ShowDialog();
                    if (result != DialogResult.OK)
                        return;

                    _settings = form.Settings;
                    _settings.Save();

                    if (_settings.StayOnSlideEnabled && _settings.EnableVirtualCanvas)
                    {
                        EnsureCurrentSlideAnchors();
                    }
                    else if (!_settings.EnableVirtualCanvas)
                    {
                        CleanupAllAnchors();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Einstellungen fehlgeschlagen",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        public void ShowAbout()
        {
            string message = AddInInfo.ProductName + " v" + AddInInfo.Version + Environment.NewLine + Environment.NewLine
                + "Funktionen & Verhalten:" + Environment.NewLine
                + "- Bleibt im Bearbeitungsmodus fest auf der aktuellen Folie fixiert." + Environment.NewLine
                + "- Kein versehentlicher Folienwechsel durch Mausrad oder Tastatur auf der Folie." + Environment.NewLine
                + "- Folienwechsel erfolgt komfortabel über die linke Miniatur-Sidebar." + Environment.NewLine
                + "- Automatischer Montagebereich (Pasteboard) über den oberen und unteren Rand hinaus." + Environment.NewLine + Environment.NewLine
                + "Aktuelle Einstellungen:" + Environment.NewLine
                + "- Modus aktiv: " + (_settings.StayOnSlideEnabled ? "Ja" : "Nein") + Environment.NewLine
                + "- Montagebereich erweitert: " + (_settings.EnableVirtualCanvas ? "Ja (" + _settings.CanvasPaddingPoints + " pt)" : "Nein") + Environment.NewLine
                + "- Tastaturnavigation sperren: " + (_settings.BlockKeyboardSlideNavigation ? "Ja" : "Nein") + Environment.NewLine
                + "- Mausrad-Schrittweite: " + _settings.WheelScrollStepLines + " Zeilen" + Environment.NewLine + Environment.NewLine
                + "Einstellungsdatei:" + Environment.NewLine
                + StayOnSlideSettings.GetSettingsFilePath() + Environment.NewLine + Environment.NewLine
                + "Hinweis:" + Environment.NewLine
                + "Temporäre Montage-Anker werden vor dem Speichern oder Drucken automatisch entfernt." + Environment.NewLine + Environment.NewLine
                + "Open Source (MIT-Lizenz) © Sperker.Training" + Environment.NewLine
                + "https://github.com/Sperker/PowerPoint-PowerTools";

            MessageBox.Show(
                message,
                AddInInfo.ProductName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }
}
