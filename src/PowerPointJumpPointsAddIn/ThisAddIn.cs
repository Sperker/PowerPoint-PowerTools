using System;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointJumpPointsAddIn
{
    public partial class ThisAddIn
    {
        internal static ThisAddIn Instance { get; private set; }

        internal JumpPointController Controller { get; private set; }
        internal JumpPointRibbon Ribbon { get; private set; }
        internal JumpPointSettings Settings { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;
            EnsureInitialized();
            HookPowerPointEvents();

            try
            {
                Controller?.RememberSlideFromActiveWindow();
                Controller?.NotifyCurrentSlideChanged();
                Ribbon?.Invalidate();
            }
            catch
            {
                // Ignore startup capture issues.
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            UnhookPowerPointEvents();

            Ribbon = null;
            Controller = null;
            Settings = null;
            Instance = null;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            EnsureInitialized();
            Ribbon = new JumpPointRibbon(Controller);
            return Ribbon;
        }

        private void EnsureInitialized()
        {
            if (Settings == null)
            {
                Settings = new JumpPointSettings();
                Settings.Load();
            }

            if (Controller == null)
            {
                Controller = new JumpPointController(this.Application, Settings);
            }
            else
            {
                Controller.SetApplication(this.Application);
            }
        }

        private void HookPowerPointEvents()
        {
            try { this.Application.WindowSelectionChange += Application_WindowSelectionChange; } catch { }
            try { this.Application.WindowActivate += Application_WindowActivate; } catch { }
        }

        private void UnhookPowerPointEvents()
        {
            try { this.Application.WindowSelectionChange -= Application_WindowSelectionChange; } catch { }
            try { this.Application.WindowActivate -= Application_WindowActivate; } catch { }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
                EnsureInitialized();
                Controller?.RememberSlideFromSelection(Sel);
                Controller?.NotifyCurrentSlideChanged();
                Ribbon?.Invalidate();
            }
            catch { }
        }

        private void Application_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            try
            {
                EnsureInitialized();
                Controller?.RememberSlideFromWindow(Wn);
                Controller?.NotifyCurrentSlideChanged();
                Ribbon?.Invalidate();
            }
            catch { }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
