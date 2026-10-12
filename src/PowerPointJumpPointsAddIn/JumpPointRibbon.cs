using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using stdole;

namespace PowerPointJumpPointsAddIn
{
    [ComVisible(true)]
    public sealed class JumpPointRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private readonly JumpPointController _controller;
        private readonly RibbonIconProvider _iconProvider;

        public JumpPointRibbon(JumpPointController controller)
        {
            _controller = controller;
            _iconProvider = new RibbonIconProvider();
        }

        public string GetCustomUI(string ribbonID)
        {
            return $@"
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon>
    <tabs>
      <tab idMso=""TabAddIns"">
        <group id=""jumpPointsGroup"" label=""{AddInInfo.RibbonGroupLabelXml}"">

          <splitButton id=""jumpPointsSplitButton"" size=""large"" getKeytip=""GetKeytip"">

            <button id=""addJumpPointButton""
                    label=""Jump Points""
                    getKeytip=""GetKeytip""
                    screentip=""Add Jump Point""
                    supertip=""Adds a jump point to the current slide. If the current slide already has a jump point, nothing is removed.""
                    getImage=""GetRibbonImage""
                    onAction=""OnAddJumpPoint"" />

            <menu id=""jumpPointsMenu"" getKeytip=""GetKeytip"">

              <button id=""addJumpPointMenuButton""
                      label=""Add Jump Point""
                      getKeytip=""GetKeytip""
                      screentip=""Add Jump Point""
                      supertip=""Adds a jump point to the current slide. If the current slide already has a jump point, nothing is removed.""
                      getImage=""GetRibbonImage""
                      onAction=""OnAddJumpPoint"" />

              <button id=""previousJumpPointMenuButton""
                      label=""Previous Jump Point""
                      getKeytip=""GetKeytip""
                      screentip=""Previous Jump Point""
                      supertip=""Jumps to the previous defined jump point in the deck.""
                      getImage=""GetRibbonImage""
                      onAction=""OnPreviousJumpPoint"" />

              <button id=""nextJumpPointMenuButton""
                      label=""Next Jump Point""
                      getKeytip=""GetKeytip""
                      screentip=""Next Jump Point""
                      supertip=""Jumps to the next defined jump point in the deck.""
                      getImage=""GetRibbonImage""
                      onAction=""OnNextJumpPoint"" />

              <menuSeparator id=""jumpPointsMenuSeparator1"" />

              <button id=""clearJumpPointsButton""
                      label=""Clear All""
                      getKeytip=""GetKeytip""
                      screentip=""Clear All Jump Points""
                      supertip=""Removes all jump points from the current presentation after confirmation.""
                      getImage=""GetRibbonImage""
                      onAction=""OnClearAllJumpPoints"" />

            </menu>

          </splitButton>

          <button id=""previousJumpPointButton""
                  label=""Previous Jump Point""
                  getKeytip=""GetKeytip""
                  screentip=""Previous Jump Point""
                  supertip=""Jumps to the previous defined jump point in the deck.""
                  size=""large""
                  getImage=""GetRibbonImage""
                  onAction=""OnPreviousJumpPoint"" />

          <button id=""nextJumpPointButton""
                  label=""Next Jump Point""
                  getKeytip=""GetKeytip""
                  screentip=""Next Jump Point""
                  supertip=""Jumps to the next defined jump point in the deck.""
                  size=""large""
                  getImage=""GetRibbonImage""
                  onAction=""OnNextJumpPoint"" />

          <separator id=""jumpPointsSeparatorAfterNavigation"" />

          <button id=""listJumpPointsButton""
                  label=""List""
                  getKeytip=""GetKeytip""
                  screentip=""List Jump Points""
                  supertip=""Shows all jump points in the current presentation and allows deleting selected jump points.""
                  getImage=""GetRibbonImage""
                  onAction=""OnListJumpPoints"" />

          <button id=""settingsJumpPointsButton""
                  label=""Settings""
                  getKeytip=""GetKeytip""
                  screentip=""Jump Points Settings""
                  supertip=""Configure messages, sound feedback, and navigator confirmations.""
                  getImage=""GetRibbonImage""
                  onAction=""OnSettings"" />

          <button id=""aboutJumpPointsButton""
                  label=""About""
                  getKeytip=""GetKeytip""
                  screentip=""About Jump Points""
                  supertip=""Shows information about the PowerPoint Jump Points Add-in.""
                  getImage=""GetRibbonImage""
                  onAction=""OnAbout"" />

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
            try
            {
                _ribbon?.Invalidate();
            }
            catch
            {
                // Ignore ribbon invalidation issues.
            }
        }

        public string GetKeytip(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "jumpPointsSplitButton":
                        return "J";

                    case "jumpPointsMenu":
                        return "J";

                    case "addJumpPointButton":
                    case "addJumpPointMenuButton":
                        return "A";

                    case "previousJumpPointButton":
                    case "previousJumpPointMenuButton":
                        return "P";

                    case "nextJumpPointButton":
                    case "nextJumpPointMenuButton":
                        return "N";

                    case "clearJumpPointsButton":
                        return "C";

                    case "listJumpPointsButton":
                        return "L";

                    case "settingsJumpPointsButton":
                        return "S";

                    case "aboutJumpPointsButton":
                        return "I";

                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public void OnAddJumpPoint(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide slide = _controller.GetCurrentSlide();

                if (slide == null)
                {
                    MessageBox.Show(
                        "No active slide found.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    return;
                }

                if (!_controller.HasJumpPoint(slide))
                {
                    /*
                     * Important:
                     * Do not call SetJumpPoint(slide) directly here.
                     * ToggleJumpPointOnCurrentSlide() uses the existing feedback logic,
                     * including sound, refresh of the open navigator and current-slide highlighting.
                     */
                    _controller.ToggleJumpPointOnCurrentSlide();
                }
                else
                {
                    /*
                     * Add-only behavior:
                     * If the current slide already has a jump point,
                     * do not remove it.
                     */
                    _controller.NotifyCurrentSlideChanged();
                }

                Invalidate();
            }
            catch
            {
                MessageBox.Show(
                    "Could not add jump point.",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        public void OnToggleJumpPoint(Office.IRibbonControl control, bool pressed)
        {
            _controller.ToggleJumpPointOnCurrentSlide();
            Invalidate();
        }

        public bool GetToggleJumpPointPressed(Office.IRibbonControl control)
        {
            try
            {
                return _controller.HasJumpPointOnCurrentSlide();
            }
            catch
            {
                return false;
            }
        }

        public void OnPreviousJumpPoint(Office.IRibbonControl control)
        {
            _controller.GoToPreviousJumpPoint();
            Invalidate();
        }

        public void OnNextJumpPoint(Office.IRibbonControl control)
        {
            _controller.GoToNextJumpPoint();
            Invalidate();
        }

        public void OnListJumpPoints(Office.IRibbonControl control)
        {
            _controller.ShowJumpPointListDialog();
            Invalidate();
        }

        public void OnClearAllJumpPoints(Office.IRibbonControl control)
        {
            _controller.ClearAllJumpPoints();
            Invalidate();
        }

        public void OnSettings(Office.IRibbonControl control)
        {
            _controller.ShowSettingsDialog();
            Invalidate();
        }

        public void OnAbout(Office.IRibbonControl control)
        {
            int jumpPointCount = _controller.GetJumpPointCount();

            MessageBox.Show(
                $"{AddInInfo.ProductName}\n" +
                $"Version {AddInInfo.Version}\n\n" +
                $"Jump points in current presentation: {jumpPointCount}\n" +
                $"Wrap around: {(_controller.IsWrapAroundEnabled() ? "Enabled" : "Disabled")}\n" +
                $"Hidden slides: {(_controller.AreHiddenSlidesIncluded() ? "Included" : "Excluded")}\n" +
                $"Toggle messages: {(_controller.IsToggleResultMessageEnabled() ? "Enabled" : "Disabled")}\n" +
                $"Toggle sound: {(_controller.IsToggleResultSoundEnabled() ? "Enabled" : "Disabled")}\n" +
                $"Navigator confirmations: {(_controller.IsNavigatorConfirmationEnabled() ? "Enabled" : "Disabled")}\n\n" +
                "Define invisible jump points on PowerPoint slides and quickly navigate between them.\n\n" +
                "Tip: Right-click the Jump Points command and add it to the Quick Access Toolbar for fast Alt-key access.\n\n" +
                "Open Source (MIT-Lizenz) © Sperker.Training\n" +
                "https://github.com/Sperker/PowerPoint-PowerTools",
                AddInInfo.ProductName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public IPictureDisp GetRibbonImage(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "addJumpPointButton":
                    case "addJumpPointMenuButton":
                        return _iconProvider.GetIcon("add_circle") ?? _iconProvider.GetIcon("bookmark");

                    case "toggleJumpPointButton":
                        return _iconProvider.GetIcon("bookmark");

                    case "previousJumpPointButton":
                    case "previousJumpPointMenuButton":
                        return _iconProvider.GetIcon("arrow_upward");

                    case "nextJumpPointButton":
                    case "nextJumpPointMenuButton":
                        return _iconProvider.GetIcon("arrow_downward");

                    case "listJumpPointsButton":
                        return _iconProvider.GetIcon("list");

                    case "clearJumpPointsButton":
                        return _iconProvider.GetIcon("delete");

                    case "settingsJumpPointsButton":
                        return _iconProvider.GetIcon("settings");

                    case "aboutJumpPointsButton":
                        return _iconProvider.GetIcon("info");

                    default:
                        return null;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}