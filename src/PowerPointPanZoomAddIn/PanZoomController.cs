using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointPanZoomAddIn
{
    internal sealed class PanZoomController : IDisposable
    {
        private const string AnchorNamePrefix = "__PPT_PANZOOM_TEMP_ANCHOR__";
        private const string AnchorTagName = "PPT_PANZOOM_TEMP_ANCHOR";
        private const string AnchorTagValue = "TRUE";

        private readonly PowerPoint.Application _app;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly MouseWheelHook _mouseWheelHook;
        private PanZoomSettings _settings;

        private bool _lastMiddleButtonDown;
        private bool _isPanning;
        private Point _dragStartScreen;

        private int _pendingZoomDirection;
        private Point _pendingZoomPoint;

        private bool _anchorSessionActive;
        private PowerPoint.Shape _anchorSessionAnchor;

        public bool PanEnabled { get; set; } = true;
        public bool ZoomToCursorEnabled { get; set; } = true;
        public bool RequireCtrlShiftForWheelZoom { get; set; } = true;

        public PanZoomController(PowerPoint.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _settings = PanZoomSettings.Load();

            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = _settings.TimerIntervalMilliseconds;
            _timer.Tick += Timer_Tick;

            _mouseWheelHook = new MouseWheelHook();
            _mouseWheelHook.MouseWheel += MouseWheelHook_MouseWheel;
        }

        public void Start()
        {
            _mouseWheelHook.Start();
            _timer.Start();
        }

        public void Dispose()
        {
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
                _mouseWheelHook.Dispose();
            }
            catch { }
        }

        public void ShowSettings()
        {
            try
            {
                PanZoomSettings editableSettings = new PanZoomSettings();
                editableSettings.ZoomLevels = (int[])_settings.ZoomLevels.Clone();
                editableSettings.PanThresholdPixels = _settings.PanThresholdPixels;
                editableSettings.PanStepDivisor = _settings.PanStepDivisor;
                editableSettings.TimerIntervalMilliseconds = _settings.TimerIntervalMilliseconds;

                using (SettingsForm form = new SettingsForm(editableSettings))
                {
                    DialogResult result = form.ShowDialog();
                    if (result != DialogResult.OK)
                        return;

                    _settings = form.Settings;
                    _settings.Save();
                    _timer.Interval = _settings.TimerIntervalMilliseconds;

                    EndAnchorSession();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Settings failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        public void ShowAbout()
        {
            string message = AddInInfo.ProductName + " v" + AddInInfo.Version + Environment.NewLine + Environment.NewLine
                + "Shortcuts" + Environment.NewLine
                + "- Ctrl + Shift + MouseWheel: Zoom to cursor" + Environment.NewLine
                + "- Middle Mouse Button + Drag: Pan" + Environment.NewLine + Environment.NewLine
                + "Current zoom levels" + Environment.NewLine
                + ZoomConfiguration.FormatZoomLevels(_settings.ZoomLevels) + Environment.NewLine + Environment.NewLine
                + "Current pan settings" + Environment.NewLine
                + "- Pan threshold pixels: " + _settings.PanThresholdPixels + Environment.NewLine
                + "- Pan step divisor: " + _settings.PanStepDivisor + Environment.NewLine
                + "- Timer interval: " + _settings.TimerIntervalMilliseconds + " ms" + Environment.NewLine + Environment.NewLine
                + "Settings file" + Environment.NewLine
                + PanZoomSettings.GetSettingsFilePath() + Environment.NewLine + Environment.NewLine
                + "Notes" + Environment.NewLine
                + "- PowerPoint Desktop on Windows only." + Environment.NewLine
                + "- Temporary invisible zoom anchors are cleaned automatically." + Environment.NewLine
                + "- Installed per user when deployed via ClickOnce." + Environment.NewLine + Environment.NewLine
                + "Open Source (MIT-Lizenz) © Sperker.Training" + Environment.NewLine
                + "https://github.com/Sperker/PowerPoint-PowerTools";

            MessageBox.Show(
                message,
                AddInInfo.ProductName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        public void CleanupAllZoomAnchors()
        {
            try
            {
                EndAnchorSession();

                var pres = _app.ActivePresentation;
                if (pres == null)
                    return;

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    DeleteAnchorsOnSlide(slide);
                }
            }
            catch { }
        }

        public void EndAnchorSession()
        {
            if (!_anchorSessionActive)
                return;

            try
            {
                DeselectCurrentSelection();
            }
            catch { }

            _anchorSessionActive = false;
            _anchorSessionAnchor = null;
        }

        private void MouseWheelHook_MouseWheel(object sender, GlobalMouseWheelEventArgs e)
        {
            try
            {
                bool cursorOverPpt = IsCursorOverPowerPoint(e.ScreenPoint);
                bool ctrlDown = IsCtrlDown();
                bool shiftDown = IsShiftDown();
                bool correctGesture = ctrlDown && shiftDown;

                if (!ZoomToCursorEnabled) return;
                if (!cursorOverPpt) return;
                if (RequireCtrlShiftForWheelZoom && !correctGesture) return;

                if (!_anchorSessionActive)
                {
                    _pendingZoomPoint = e.ScreenPoint;
                }

                int direction = e.Delta > 0 ? +1 : -1;
                Interlocked.Add(ref _pendingZoomDirection, direction);

                e.Handled = true;
            }
            catch { }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                bool ctrlShiftDown = ZoomToCursorEnabled
                    && IsCtrlDown()
                    && IsShiftDown();

                if (!ctrlShiftDown && _anchorSessionActive)
                {
                    EndAnchorSession();
                }

                int direction = Interlocked.Exchange(ref _pendingZoomDirection, 0);
                if (direction != 0)
                {
                    direction = direction > 0 ? +1 : -1;
                    ApplyZoomWithFixedAnchor(direction);
                }

                HandlePan();
            }
            catch { }
        }

        private void ApplyZoomWithFixedAnchor(int direction)
        {
            var wnd = GetActiveWindow();
            if (wnd == null)
                return;

            if (!_anchorSessionActive || _anchorSessionAnchor == null)
            {
                Point startPoint = _pendingZoomPoint.IsEmpty ? NativeMethods.GetCursorPosition() : _pendingZoomPoint;
                EnsureAnchorSession(startPoint);
            }

            if (_anchorSessionAnchor == null)
                return;

            try
            {
                _anchorSessionAnchor.Select(MsoTriState.msoTrue);
            }
            catch { }

            int currentZoom = wnd.View.Zoom;
            int targetZoom = GetNextZoomLevel(currentZoom, direction);

            if (targetZoom == currentZoom) return;

            wnd.View.ZoomToFit = MsoTriState.msoFalse;
            wnd.View.Zoom = targetZoom;
        }

        private void EnsureAnchorSession(Point screenPoint)
        {
            try
            {
                var wnd = GetActiveWindow();
                if (wnd == null)
                    return;

                var slide = GetCurrentSlide(wnd);
                if (slide == null)
                    return;

                PointF slidePoint = ScreenPointToSlidePoint(wnd, screenPoint);
                slidePoint = ClampToSlide(slidePoint);

                PowerPoint.Shape anchor = GetOrCreateAnchor(slide);
                MoveAnchorIfNeeded(anchor, slidePoint);

                try
                {
                    anchor.Select(MsoTriState.msoTrue);
                }
                catch { }

                _anchorSessionAnchor = anchor;
                _anchorSessionActive = true;
            }
            catch { }
        }

        private int GetNextZoomLevel(int currentZoom, int direction)
        {
            int[] zoomLevels = _settings.ZoomLevels;
            if (zoomLevels == null || zoomLevels.Length == 0)
                zoomLevels = ZoomConfiguration.DefaultZoomLevels;

            if (direction > 0)
            {
                for (int i = 0; i < zoomLevels.Length; i++)
                {
                    if (currentZoom < zoomLevels[i])
                        return zoomLevels[i];
                }

                return zoomLevels[zoomLevels.Length - 1];
            }

            if (direction < 0)
            {
                for (int i = zoomLevels.Length - 1; i >= 0; i--)
                {
                    if (currentZoom > zoomLevels[i])
                        return zoomLevels[i];
                }

                return zoomLevels[0];
            }

            return currentZoom;
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
                return _app.ActiveWindow.View.Slide as PowerPoint.Slide;
            }
            catch
            {
                return null;
            }
        }

        private PowerPoint.Shape GetOrCreateAnchor(PowerPoint.Slide slide)
        {
            PowerPoint.Shape existing = FindAnchor(slide);
            if (existing != null)
                return existing;

            float x = Math.Max(1, _app.ActivePresentation.PageSetup.SlideWidth / 2);
            float y = Math.Max(1, _app.ActivePresentation.PageSetup.SlideHeight / 2);

            PowerPoint.Shape anchor = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                x,
                y,
                2f,
                2f
            );

            anchor.Name = AnchorNamePrefix + slide.SlideID;
            anchor.Tags.Add(AnchorTagName, AnchorTagValue);

            anchor.Fill.Visible = MsoTriState.msoTrue;
            anchor.Fill.Transparency = 1.0f;
            anchor.Line.Visible = MsoTriState.msoFalse;

            try { anchor.ZOrder(Office.MsoZOrderCmd.msoSendToBack); } catch { }

            return anchor;
        }

        private PowerPoint.Shape FindAnchor(PowerPoint.Slide slide)
        {
            try
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    bool nameMatches = false;
                    bool tagMatches = false;

                    try { nameMatches = shape.Name != null && shape.Name.StartsWith(AnchorNamePrefix, StringComparison.OrdinalIgnoreCase); } catch { }
                    try { tagMatches = string.Equals(shape.Tags[AnchorTagName], AnchorTagValue, StringComparison.OrdinalIgnoreCase); } catch { }

                    if (nameMatches || tagMatches)
                        return shape;
                }
            }
            catch { }

            return null;
        }

        private int DeleteAnchorsOnSlide(PowerPoint.Slide slide)
        {
            int deleted = 0;

            try
            {
                for (int i = slide.Shapes.Count; i >= 1; i--)
                {
                    PowerPoint.Shape shape = slide.Shapes[i];
                    bool nameMatches = false;
                    bool tagMatches = false;

                    try { nameMatches = shape.Name != null && shape.Name.StartsWith(AnchorNamePrefix, StringComparison.OrdinalIgnoreCase); } catch { }
                    try { tagMatches = string.Equals(shape.Tags[AnchorTagName], AnchorTagValue, StringComparison.OrdinalIgnoreCase); } catch { }

                    if (nameMatches || tagMatches)
                    {
                        shape.Delete();
                        deleted++;
                    }
                }
            }
            catch { }

            return deleted;
        }

        private void MoveAnchorIfNeeded(PowerPoint.Shape anchor, PointF slidePoint)
        {
            try
            {
                float newLeft = slidePoint.X;
                float newTop = slidePoint.Y;

                if (Math.Abs(anchor.Left - newLeft) < 1.0f && Math.Abs(anchor.Top - newTop) < 1.0f)
                    return;

                anchor.Left = newLeft;
                anchor.Top = newTop;
            }
            catch { }
        }

        private PointF ScreenPointToSlidePoint(PowerPoint.DocumentWindow wnd, Point screenPoint)
        {
            try
            {
                float sx0 = wnd.PointsToScreenPixelsX(0f);
                float sx100 = wnd.PointsToScreenPixelsX(100f);
                float sy0 = wnd.PointsToScreenPixelsY(0f);
                float sy100 = wnd.PointsToScreenPixelsY(100f);

                float scaleX = (sx100 - sx0) / 100f;
                float scaleY = (sy100 - sy0) / 100f;

                if (Math.Abs(scaleX) < 0.0001f || Math.Abs(scaleY) < 0.0001f)
                    return GetSlideCenterPoint();

                float x = (screenPoint.X - sx0) / scaleX;
                float y = (screenPoint.Y - sy0) / scaleY;

                return new PointF(x, y);
            }
            catch
            {
                return GetSlideCenterPoint();
            }
        }

        private PointF GetSlideCenterPoint()
        {
            try
            {
                return new PointF(
                    _app.ActivePresentation.PageSetup.SlideWidth / 2f,
                    _app.ActivePresentation.PageSetup.SlideHeight / 2f
                );
            }
            catch
            {
                return new PointF(100f, 100f);
            }
        }

        private PointF ClampToSlide(PointF point)
        {
            try
            {
                float width = _app.ActivePresentation.PageSetup.SlideWidth;
                float height = _app.ActivePresentation.PageSetup.SlideHeight;

                float x = point.X;
                float y = point.Y;

                if (x < 0) x = 0;
                if (y < 0) y = 0;
                if (x > width - 2f) x = width - 2f;
                if (y > height - 2f) y = height - 2f;

                return new PointF(x, y);
            }
            catch
            {
                return point;
            }
        }

        private void DeselectCurrentSelection()
        {
            try
            {
                var wnd = GetActiveWindow();
                if (wnd == null)
                    return;

                if (wnd.Selection != null)
                    wnd.Selection.Unselect();
            }
            catch { }
        }

        private void HandlePan()
        {
            if (!PanEnabled)
            {
                ResetPanState();
                return;
            }

            if (!IsPowerPointForeground() && !IsCursorOverPowerPoint())
            {
                ResetPanState();
                return;
            }

            bool middleDown = NativeMethods.IsKeyDown(NativeMethods.VK_MBUTTON);

            if (middleDown && !_lastMiddleButtonDown)
            {
                _dragStartScreen = NativeMethods.GetCursorPosition();
                _isPanning = true;
            }

            if (!middleDown && _lastMiddleButtonDown)
            {
                ResetPanState();
            }

            if (middleDown && _isPanning)
            {
                Point current = NativeMethods.GetCursorPosition();

                int dx = current.X - _dragStartScreen.X;
                int dy = current.Y - _dragStartScreen.Y;

                if (Math.Abs(dx) >= _settings.PanThresholdPixels || Math.Abs(dy) >= _settings.PanThresholdPixels)
                {
                    PanByDelta(dx, dy);
                    _dragStartScreen = current;
                }
            }

            _lastMiddleButtonDown = middleDown;
        }

        private void PanByDelta(int dx, int dy)
        {
            try
            {
                var wnd = GetActiveWindow();
                if (wnd == null)
                    return;

                int divisor = _settings.PanStepDivisor <= 0 ? 8 : _settings.PanStepDivisor;
                int horizontalSteps = Math.Abs(dx) / divisor;
                int verticalSteps = Math.Abs(dy) / divisor;

                if (horizontalSteps == 0 && verticalSteps == 0)
                    return;

                int down = 0;
                int up = 0;
                int toRight = 0;
                int toLeft = 0;

                if (dx > 0)
                    toLeft = horizontalSteps;
                else if (dx < 0)
                    toRight = horizontalSteps;

                if (dy > 0)
                    up = verticalSteps;
                else if (dy < 0)
                    down = verticalSteps;

                wnd.SmallScroll(
                    Down: down,
                    Up: up,
                    ToRight: toRight,
                    ToLeft: toLeft
                );
            }
            catch { }
        }

        private void ResetPanState()
        {
            _lastMiddleButtonDown = false;
            _isPanning = false;
        }

        private PowerPoint.DocumentWindow GetActiveWindow()
        {
            try
            {
                if (_app == null)
                    return null;

                return _app.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }

        private bool IsPowerPointForeground()
        {
            try
            {
                IntPtr fg = NativeMethods.GetForegroundWindow();
                if (fg == IntPtr.Zero)
                    return false;

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

        private bool IsCursorOverPowerPoint()
        {
            return IsCursorOverPowerPoint(NativeMethods.GetCursorPosition());
        }

        private bool IsCursorOverPowerPoint(Point screenPoint)
        {
            try
            {
                NativeMethods.POINT p = new NativeMethods.POINT { X = screenPoint.X, Y = screenPoint.Y };
                IntPtr hwnd = NativeMethods.WindowFromPoint(p);
                if (hwnd == IntPtr.Zero)
                    return false;

                IntPtr root = NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT);
                if (root == IntPtr.Zero)
                    root = hwnd;

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

        private bool IsCtrlDown()
        {
            return NativeMethods.IsKeyDown(NativeMethods.VK_CONTROL)
                || NativeMethods.IsKeyDown(NativeMethods.VK_LCONTROL)
                || NativeMethods.IsKeyDown(NativeMethods.VK_RCONTROL);
        }

        private bool IsShiftDown()
        {
            return NativeMethods.IsKeyDown(NativeMethods.VK_SHIFT)
                || NativeMethods.IsKeyDown(NativeMethods.VK_LSHIFT)
                || NativeMethods.IsKeyDown(NativeMethods.VK_RSHIFT);
        }
    }
}
