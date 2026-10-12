using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointPanZoomAddIn
{
    internal sealed class SettingsForm : Form
    {
        private readonly TextBox _txtZoomLevels;
        private readonly NumericUpDown _numPanThreshold;
        private readonly NumericUpDown _numPanStepDivisor;
        private readonly NumericUpDown _numTimerInterval;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;
        private readonly Button _btnReset;

        public PanZoomSettings Settings { get; private set; }

        public SettingsForm(PanZoomSettings settings)
        {
            Settings = settings ?? new PanZoomSettings();

            Text = AddInInfo.ProductName + " Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 600;
            Height = 475;

            Label lblTitle = new Label();
            lblTitle.Text = AddInInfo.ProductName + " v" + AddInInfo.Version;
            lblTitle.Font = new Font(lblTitle.Font, FontStyle.Bold);
            lblTitle.Left = 16;
            lblTitle.Top = 14;
            lblTitle.Width = 540;
            lblTitle.Height = 24;
            Controls.Add(lblTitle);

            GroupBox grpZoom = new GroupBox();
            grpZoom.Text = "Zoom settings";
            grpZoom.Left = 16;
            grpZoom.Top = 48;
            grpZoom.Width = 550;
            grpZoom.Height = 120;
            Controls.Add(grpZoom);

            Label lblZoom = new Label();
            lblZoom.Text = "Zoom levels (comma-separated, 10-400, at least two values):";
            lblZoom.Left = 14;
            lblZoom.Top = 26;
            lblZoom.Width = 510;
            lblZoom.Height = 20;
            grpZoom.Controls.Add(lblZoom);

            _txtZoomLevels = new TextBox();
            _txtZoomLevels.Left = 14;
            _txtZoomLevels.Top = 50;
            _txtZoomLevels.Width = 515;
            _txtZoomLevels.Height = 24;
            grpZoom.Controls.Add(_txtZoomLevels);

            Label lblZoomHint = new Label();
            lblZoomHint.Text = "Tip: Include 100% as a stable base zoom level. Values are sorted and duplicates are removed automatically.";
            lblZoomHint.Left = 14;
            lblZoomHint.Top = 78;
            lblZoomHint.Width = 515;
            lblZoomHint.Height = 30;
            grpZoom.Controls.Add(lblZoomHint);

            GroupBox grpPan = new GroupBox();
            grpPan.Text = "Pan settings";
            grpPan.Left = 16;
            grpPan.Top = 180;
            grpPan.Width = 550;
            grpPan.Height = 110;
            Controls.Add(grpPan);

            Label lblPanThreshold = new Label();
            lblPanThreshold.Text = "Pan threshold pixels (lower = reacts earlier):";
            lblPanThreshold.Left = 14;
            lblPanThreshold.Top = 30;
            lblPanThreshold.Width = 360;
            lblPanThreshold.Height = 20;
            grpPan.Controls.Add(lblPanThreshold);

            _numPanThreshold = new NumericUpDown();
            _numPanThreshold.Left = 395;
            _numPanThreshold.Top = 27;
            _numPanThreshold.Width = 120;
            _numPanThreshold.Minimum = 1;
            _numPanThreshold.Maximum = 100;
            grpPan.Controls.Add(_numPanThreshold);

            Label lblPanDivisor = new Label();
            lblPanDivisor.Text = "Pan step divisor (higher = slower/finer pan):";
            lblPanDivisor.Left = 14;
            lblPanDivisor.Top = 65;
            lblPanDivisor.Width = 360;
            lblPanDivisor.Height = 20;
            grpPan.Controls.Add(lblPanDivisor);

            _numPanStepDivisor = new NumericUpDown();
            _numPanStepDivisor.Left = 395;
            _numPanStepDivisor.Top = 62;
            _numPanStepDivisor.Width = 120;
            _numPanStepDivisor.Minimum = 1;
            _numPanStepDivisor.Maximum = 100;
            grpPan.Controls.Add(_numPanStepDivisor);

            GroupBox grpAdvanced = new GroupBox();
            grpAdvanced.Text = "Advanced";
            grpAdvanced.Left = 16;
            grpAdvanced.Top = 302;
            grpAdvanced.Width = 550;
            grpAdvanced.Height = 70;
            Controls.Add(grpAdvanced);

            Label lblTimer = new Label();
            lblTimer.Text = "Timer interval milliseconds:";
            lblTimer.Left = 14;
            lblTimer.Top = 30;
            lblTimer.Width = 360;
            lblTimer.Height = 20;
            grpAdvanced.Controls.Add(lblTimer);

            _numTimerInterval = new NumericUpDown();
            _numTimerInterval.Left = 395;
            _numTimerInterval.Top = 27;
            _numTimerInterval.Width = 120;
            _numTimerInterval.Minimum = 10;
            _numTimerInterval.Maximum = 250;
            grpAdvanced.Controls.Add(_numTimerInterval);

            _btnReset = new Button();
            _btnReset.Text = "Reset defaults";
            _btnReset.Left = 16;
            _btnReset.Top = 392;
            _btnReset.Width = 120;
            _btnReset.Click += BtnReset_Click;
            Controls.Add(_btnReset);

            _btnSave = new Button();
            _btnSave.Text = "Save";
            _btnSave.Left = 370;
            _btnSave.Top = 392;
            _btnSave.Width = 90;
            _btnSave.Click += BtnSave_Click;
            Controls.Add(_btnSave);

            _btnCancel = new Button();
            _btnCancel.Text = "Cancel";
            _btnCancel.Left = 475;
            _btnCancel.Top = 392;
            _btnCancel.Width = 90;
            _btnCancel.Click += BtnCancel_Click;
            Controls.Add(_btnCancel);

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;

            LoadSettingsIntoControls();
        }

        private void LoadSettingsIntoControls()
        {
            _txtZoomLevels.Text = ZoomConfiguration.FormatZoomLevels(Settings.ZoomLevels);
            _numPanThreshold.Value = Settings.PanThresholdPixels;
            _numPanStepDivisor.Value = Settings.PanStepDivisor;
            _numTimerInterval.Value = Settings.TimerIntervalMilliseconds;
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            Settings.ResetToDefaults();
            LoadSettingsIntoControls();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            int[] parsedZoomLevels;
            string error;
            if (!ZoomConfiguration.TryParseZoomLevels(_txtZoomLevels.Text, out parsedZoomLevels, out error))
            {
                MessageBox.Show(error, "Invalid settings", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!ZoomConfiguration.ContainsRecommendedBaseZoom(parsedZoomLevels))
            {
                DialogResult answer = MessageBox.Show(
                    "The zoom levels do not contain 100%. This is allowed, but 100% is recommended as a stable base zoom level. Save anyway?",
                    "Recommended zoom level missing",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (answer != DialogResult.Yes)
                    return;
            }

            Settings.ZoomLevels = parsedZoomLevels;
            Settings.PanThresholdPixels = (int)_numPanThreshold.Value;
            Settings.PanStepDivisor = (int)_numPanStepDivisor.Value;
            Settings.TimerIntervalMilliseconds = (int)_numTimerInterval.Value;

            _txtZoomLevels.Text = ZoomConfiguration.FormatZoomLevels(Settings.ZoomLevels);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
