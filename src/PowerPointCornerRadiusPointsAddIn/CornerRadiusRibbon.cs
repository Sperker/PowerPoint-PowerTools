using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace PowerPointCornerRadiusPointsAddIn
{
    [ComVisible(true)]
    public class CornerRadiusRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
  <ribbon>
    <tabs>
      <tab idMso='TabAddIns'>
        <group id='grpCornerRadius' label='" + AddInInfo.RibbonGroupLabelXml + @"'>
          <button id='btnCopyRadius' label='Copy Radius' size='large' getImage='GetRibbonImage' screentip='Copy Corner Radius' supertip='Copies the absolute corner radius of the currently selected shape in points (pt).' onAction='OnCopyRadius'/>
          <button id='btnApplyRadius' label='Apply Radius' size='large' getImage='GetRibbonImage' screentip='Apply Corner Radius' supertip='Applies the stored absolute corner radius to all selected shapes.' onAction='OnApplyRadius' getEnabled='GetApplyRadiusEnabled'/>
          <box id='boxCornerRadiusUtilities' boxStyle='vertical'>
            <button id='btnSettings' label='Settings' getImage='GetRibbonImage' screentip='Corner Radius Settings' supertip='Configure manual radius and feedback options.' onAction='OnSettings'/>
            <button id='btnAbout' label='About' getImage='GetRibbonImage' screentip='About Corner Radius' supertip='View product information and current status.' onAction='OnAbout'/>
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
            if (Globals.ThisAddIn != null && Globals.ThisAddIn.Controller != null)
            {
                Globals.ThisAddIn.Controller.RegisterRibbon(this);
            }
        }

        public void Invalidate()
        {
            if (_ribbon != null)
            {
                _ribbon.Invalidate();
            }
        }

        public void InvalidateControl(string controlId)
        {
            if (_ribbon != null)
            {
                try
                {
                    _ribbon.InvalidateControl(controlId);
                }
                catch
                {
                    _ribbon.Invalidate();
                }
            }
        }

        public object GetRibbonImage(Office.IRibbonControl control)
        {
            return RibbonIconProvider.GetImage(control.Id);
        }

        public bool GetApplyRadiusEnabled(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn != null
                && Globals.ThisAddIn.Controller != null
                && Globals.ThisAddIn.Controller.CanApplyRadius;
        }

        public void OnCopyRadius(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn == null || Globals.ThisAddIn.Controller == null)
                return;

            Globals.ThisAddIn.Controller.CopyRadius();
        }

        public void OnApplyRadius(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn == null || Globals.ThisAddIn.Controller == null)
                return;

            Globals.ThisAddIn.Controller.ApplyRadius();
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
