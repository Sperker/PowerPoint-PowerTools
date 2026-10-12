using System;
using Office = Microsoft.Office.Core;

namespace PowerPointStayOnSlidePointsAddIn
{
    public partial class ThisAddIn
    {
        internal StayOnSlideController Controller { get; private set; }
        private StayOnSlideRibbon _ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Controller = new StayOnSlideController(this.Application);
            Controller.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (Controller != null)
                {
                    Controller.CleanupAllAnchors();
                    Controller.Dispose();
                    Controller = null;
                }
            }
            catch
            {
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new StayOnSlideRibbon();
            return _ribbon;
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
