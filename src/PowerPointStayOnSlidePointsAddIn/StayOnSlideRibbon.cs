using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace PowerPointStayOnSlidePointsAddIn
{
    [ComVisible(true)]
    public class StayOnSlideRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabAddIns'>
        <group id='grpStayOnSlide' label='" + AddInInfo.RibbonGroupLabelXml + @"'>
          <toggleButton id='tglStayOnSlide' label='Stay on Slide' size='large' getImage='GetRibbonImage' screentip='Stay on Slide Mode' supertip='Enables or disables Stay on Slide mode. When active, scrolling on the slide canvas will not accidentally switch to the next or previous slide. Slide switching remains available via the left thumbnail sidebar.' onAction='OnToggleStayOnSlide' getPressed='GetStayOnSlidePressed'/>
          <box id='boxStayOnSlideUtilities' boxStyle='vertical'>
            <button id='btnSettings' label='Settings' getImage='GetRibbonImage' screentip='Stay on Slide Settings' supertip='Configure pasteboard padding size and keyboard navigation lock.' onAction='OnSettings'/>
            <button id='btnAbout' label='About' getImage='GetRibbonImage' screentip='About Stay on Slide' supertip='Show version details, features, current settings, and file paths.' onAction='OnAbout'/>
          </box>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
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

        public bool GetStayOnSlidePressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn != null
                && Globals.ThisAddIn.Controller != null
                && Globals.ThisAddIn.Controller.StayOnSlideEnabled;
        }

        public void OnToggleStayOnSlide(Office.IRibbonControl control, bool pressed)
        {
            if (Globals.ThisAddIn == null || Globals.ThisAddIn.Controller == null)
                return;

            Globals.ThisAddIn.Controller.StayOnSlideEnabled = pressed;
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
