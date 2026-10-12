using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace PowerPointPanZoomAddIn
{
    [ComVisible(true)]
    public class PanZoomRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'><ribbon><tabs><tab idMso='TabAddIns'><group id='grpPanZoom' label='" + AddInInfo.RibbonGroupLabelXml + @"'><toggleButton id='tglPan' label='Pan Mode' size='large' getImage='GetRibbonImage' screentip='Pan Mode' supertip='Enable or disable panning in PowerPoint edit view. When enabled, hold the middle mouse button and drag to pan around the zoomed slide.' onAction='OnTogglePan' getPressed='GetPanPressed'/><toggleButton id='tglZoomCursor' label='Zoom to Cursor' size='large' getImage='GetRibbonImage' screentip='Zoom to Cursor' supertip='Enable or disable cursor-based zooming. Hold Ctrl + Shift and use the mouse wheel over the slide to zoom toward the cursor using predefined zoom levels.' onAction='OnToggleZoomToCursor' getPressed='GetZoomToCursorPressed'/><box id='boxPanZoomUtilities' boxStyle='vertical'><button id='btnSettings' label='Settings' getImage='GetRibbonImage' screentip='Pan &amp; Zoom Settings' supertip='Configure zoom levels and pan sensitivity.' onAction='OnSettings'/><button id='btnAbout' label='About' getImage='GetRibbonImage' screentip='About Pan &amp; Zoom' supertip='Show version, shortcuts, zoom levels, current settings and settings file path.' onAction='OnAbout'/></box></group></tab></tabs></ribbon></customUI>";
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void Invalidate()
        {
            if (_ribbon != null)
            {
                _ribbon.Invalidate();
            }
        }

        public object GetRibbonImage(Office.IRibbonControl control)
        {
            return RibbonIconProvider.GetImage(control.Id);
        }

        public bool GetPanPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn != null
                && Globals.ThisAddIn.Controller != null
                && Globals.ThisAddIn.Controller.PanEnabled;
        }

        public bool GetZoomToCursorPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn != null
                && Globals.ThisAddIn.Controller != null
                && Globals.ThisAddIn.Controller.ZoomToCursorEnabled;
        }

        public void OnTogglePan(Office.IRibbonControl control, bool pressed)
        {
            if (Globals.ThisAddIn == null || Globals.ThisAddIn.Controller == null)
                return;

            Globals.ThisAddIn.Controller.PanEnabled = pressed;
            Invalidate();
        }

        public void OnToggleZoomToCursor(Office.IRibbonControl control, bool pressed)
        {
            if (Globals.ThisAddIn == null || Globals.ThisAddIn.Controller == null)
                return;

            Globals.ThisAddIn.Controller.ZoomToCursorEnabled = pressed;

            if (!pressed)
            {
                Globals.ThisAddIn.Controller.EndAnchorSession();
                Globals.ThisAddIn.Controller.CleanupAllZoomAnchors();
            }

            Invalidate();
        }

        public void OnSettings(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn == null || Globals.ThisAddIn.Controller == null)
                return;

            Globals.ThisAddIn.Controller.ShowSettings();
        }

        public void OnAbout(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn == null || Globals.ThisAddIn.Controller == null)
                return;

            Globals.ThisAddIn.Controller.ShowAbout();
        }
    }
}
