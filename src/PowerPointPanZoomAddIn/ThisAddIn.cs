using System;
using Office = Microsoft.Office.Core;

namespace PowerPointPanZoomAddIn
{
    public partial class ThisAddIn
    {
        internal PanZoomController Controller { get; private set; }
        private PanZoomRibbon _ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {

            Controller = new PanZoomController(this.Application);
            Controller.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (Controller != null)
                {
                    Controller.EndAnchorSession();
                    Controller.CleanupAllZoomAnchors();
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
            _ribbon = new PanZoomRibbon();
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
