using System;
using Office = Microsoft.Office.Core;

namespace PowerPointCornerRadiusPointsAddIn
{
    public partial class ThisAddIn
    {
        internal CornerRadiusController Controller { get; private set; }
        private CornerRadiusRibbon _ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Controller = new CornerRadiusController(this.Application);
            if (_ribbon != null)
            {
                Controller.RegisterRibbon(_ribbon);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Controller = null;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new CornerRadiusRibbon();
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
