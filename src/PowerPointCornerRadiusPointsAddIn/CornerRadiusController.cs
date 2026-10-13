using System;
using System.Drawing;
using System.Media;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointCornerRadiusPointsAddIn
{
    internal sealed class CornerRadiusController
    {
        private readonly PowerPoint.Application _app;
        private CornerRadiusSettings _settings;
        private float? _copiedRadiusPoints;
        private CornerRadiusRibbon _ribbon;

        public CornerRadiusSettings Settings => _settings;
        public float? CopiedRadiusPoints => _copiedRadiusPoints;

        public bool CanApplyRadius
        {
            get
            {
                if (_copiedRadiusPoints.HasValue && _copiedRadiusPoints.Value > 0)
                    return true;

                if (_settings.ManualRadiusPoints.HasValue && _settings.ManualRadiusPoints.Value > 0)
                    return true;

                return false;
            }
        }

        public float EffectiveRadius
        {
            get
            {
                if (_copiedRadiusPoints.HasValue && _copiedRadiusPoints.Value > 0)
                    return _copiedRadiusPoints.Value;

                if (_settings.ManualRadiusPoints.HasValue && _settings.ManualRadiusPoints.Value > 0)
                    return _settings.ManualRadiusPoints.Value;

                return 0f;
            }
        }

        public CornerRadiusController(PowerPoint.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _settings = CornerRadiusSettings.Load();
        }

        public void RegisterRibbon(CornerRadiusRibbon ribbon)
        {
            _ribbon = ribbon;
        }

        public void CopyRadius()
        {
            try
            {
                var wnd = _app.ActiveWindow;
                if (wnd == null || wnd.Selection == null)
                {
                    MessageBox.Show(
                        "Please select a shape with rounded corners on the slide first.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                var sel = wnd.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange == null || sel.ShapeRange.Count == 0)
                {
                    MessageBox.Show(
                        "Please select a shape with rounded corners on the slide first.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                PowerPoint.Shape sourceShape = sel.ShapeRange[1];
                float radius;

                if (!TryGetCornerRadius(sourceShape, out radius))
                {
                    MessageBox.Show(
                        "The selected shape does not have an adjustable corner radius.\n\nPlease select e.g. a rounded rectangle.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                _copiedRadiusPoints = radius;

                // Feedback
                if (_settings.PlaySoundOnCopy)
                {
                    try { SystemSounds.Asterisk.Play(); } catch { }
                }

                if (_settings.ShowMessageOnCopy)
                {
                    MessageBox.Show(
                        $"Corner radius of {radius:0.##} pt copied successfully.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }

                _ribbon?.InvalidateControl("btnApplyRadius");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error copying corner radius: " + ex.Message,
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        public void ApplyRadius()
        {
            try
            {
                if (!CanApplyRadius)
                {
                    MessageBox.Show(
                        "No corner radius available yet.\n\nPlease copy a radius from a shape first or enter a value in Settings.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                float targetRadius = EffectiveRadius;
                if (targetRadius <= 0)
                    return;

                var wnd = _app.ActiveWindow;
                if (wnd == null || wnd.Selection == null)
                {
                    MessageBox.Show(
                        "Please select at least one target shape on the slide.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                var sel = wnd.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange == null || sel.ShapeRange.Count == 0)
                {
                    MessageBox.Show(
                        "Please select at least one target shape on the slide.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                int updatedCount = 0;
                for (int i = 1; i <= sel.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape shape = sel.ShapeRange[i];
                    if (ApplyCornerRadiusToShape(shape, targetRadius))
                    {
                        updatedCount++;
                    }
                }

                if (updatedCount == 0)
                {
                    MessageBox.Show(
                        "None of the selected shapes support rounded corners.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error applying corner radius: " + ex.Message,
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private bool TryGetCornerRadius(PowerPoint.Shape shape, out float radius)
        {
            radius = 0f;
            if (shape == null) return false;

            try
            {
                // Group check
                if (shape.Type == MsoShapeType.msoGroup && shape.GroupItems != null && shape.GroupItems.Count > 0)
                {
                    return TryGetCornerRadius(shape.GroupItems[1], out radius);
                }

                if (shape.Type != MsoShapeType.msoAutoShape)
                    return false;

                if (shape.Adjustments == null || shape.Adjustments.Count == 0)
                    return false;

                float minDim = Math.Min(shape.Width, shape.Height);
                if (minDim <= 0.01f)
                    return false;

                // For rounded rectangle shapes, adjustment 1 controls corner radius
                float adj = shape.Adjustments[1];
                radius = adj * minDim;
                return radius >= 0;
            }
            catch
            {
                return false;
            }
        }

        private bool ApplyCornerRadiusToShape(PowerPoint.Shape shape, float targetRadiusPoints)
        {
            if (shape == null) return false;

            try
            {
                // If shape is a group, apply recursively to each group item
                if (shape.Type == MsoShapeType.msoGroup && shape.GroupItems != null)
                {
                    bool anyUpdated = false;
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        if (ApplyCornerRadiusToShape(shape.GroupItems[i], targetRadiusPoints))
                            anyUpdated = true;
                    }
                    return anyUpdated;
                }

                // If user selected a standard rectangle, automatically convert to rounded rectangle!
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRectangle)
                {
                    shape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle;
                }

                if (shape.Adjustments == null || shape.Adjustments.Count == 0)
                    return false;

                float minDim = Math.Min(shape.Width, shape.Height);
                if (minDim <= 0.01f)
                    return false;

                // Calculate required relative adjustment: adj = radius / minDim
                float adj = targetRadiusPoints / minDim;
                adj = Math.Max(0.0f, Math.Min(0.5f, adj));

                shape.Adjustments[1] = adj;
                return true;
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
                using (SettingsForm form = new SettingsForm(_settings, _copiedRadiusPoints))
                {
                    DialogResult result = form.ShowDialog();
                    if (result != DialogResult.OK)
                        return;

                    _settings = form.Settings;
                    _settings.Save();

                    if (form.ApplyManualNowRequested || form.SelectedRadiusValue.HasValue)
                    {
                        _copiedRadiusPoints = form.SelectedRadiusValue;
                    }

                    _ribbon?.InvalidateControl("btnApplyRadius");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Settings failed: " + ex.Message,
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        public void ShowAbout()
        {
            string currentRadiusStr = _copiedRadiusPoints.HasValue
                ? _copiedRadiusPoints.Value.ToString("0.##") + " pt (copied)"
                : (_settings.ManualRadiusPoints.HasValue
                    ? _settings.ManualRadiusPoints.Value.ToString("0.##") + " pt (manual)"
                    : "None set");

            string message = AddInInfo.ProductName + " v" + AddInInfo.Version + Environment.NewLine + Environment.NewLine
                + "Features:" + Environment.NewLine
                + "- Copies the absolute corner radius of a shape in points (pt)." + Environment.NewLine
                + "- Applies the exact same visual corner radius across any number of shapes regardless of dimensions." + Environment.NewLine
                + "- Automatically converts standard rectangles into rounded rectangles upon applying." + Environment.NewLine
                + "- Fully compatible with the Quick Access Toolbar (QAT)." + Environment.NewLine + Environment.NewLine
                + "Status:" + Environment.NewLine
                + "- Active radius: " + currentRadiusStr + Environment.NewLine
                + "- Sound on copy: " + (_settings.PlaySoundOnCopy ? "Enabled" : "Disabled") + Environment.NewLine
                + "- Notification on copy: " + (_settings.ShowMessageOnCopy ? "Enabled" : "Disabled") + Environment.NewLine + Environment.NewLine
                + "Settings file:" + Environment.NewLine
                + CornerRadiusSettings.GetSettingsFilePath() + Environment.NewLine + Environment.NewLine
                + "Open Source (MIT License) © Sperker.Training" + Environment.NewLine
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
