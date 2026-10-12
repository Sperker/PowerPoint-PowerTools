using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointCornerRadiusPointsAddIn
{
    internal sealed class SettingsForm : Form
    {
        private readonly CheckBox _chkPlaySound;
        private readonly CheckBox _chkShowMessage;
        private readonly CheckBox _chkEnableManualRadius;
        private readonly NumericUpDown _numManualRadius;
        private readonly Button _btnApplyManualNow;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;
        private readonly Button _btnReset;

        public CornerRadiusSettings Settings { get; private set; }
        public bool ApplyManualNowRequested { get; private set; }
        public float? SelectedRadiusValue { get; private set; }

        public SettingsForm(CornerRadiusSettings settings, float? currentActiveRadius)
        {
            Settings = settings ?? new CornerRadiusSettings();

            Text = AddInInfo.ProductName + " Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 540;
            Height = 390;

            Label lblTitle = new Label();
            lblTitle.Text = AddInInfo.ProductName + " v" + AddInInfo.Version;
            lblTitle.Font = new Font(lblTitle.Font, FontStyle.Bold);
            lblTitle.Left = 16;
            lblTitle.Top = 14;
            lblTitle.Width = 490;
            lblTitle.Height = 22;
            Controls.Add(lblTitle);

            // GroupBox 1: Feedback on Copy
            GroupBox grpFeedback = new GroupBox();
            grpFeedback.Text = "Feedback on Copying Radius";
            grpFeedback.Left = 16;
            grpFeedback.Top = 42;
            grpFeedback.Width = 490;
            grpFeedback.Height = 90;
            Controls.Add(grpFeedback);

            _chkPlaySound = new CheckBox();
            _chkPlaySound.Text = "Play sound when a corner radius is copied";
            _chkPlaySound.Left = 14;
            _chkPlaySound.Top = 24;
            _chkPlaySound.Width = 460;
            _chkPlaySound.Height = 22;
            grpFeedback.Controls.Add(_chkPlaySound);

            _chkShowMessage = new CheckBox();
            _chkShowMessage.Text = "Show notification message when a corner radius is copied";
            _chkShowMessage.Left = 14;
            _chkShowMessage.Top = 52;
            _chkShowMessage.Width = 460;
            _chkShowMessage.Height = 22;
            grpFeedback.Controls.Add(_chkShowMessage);

            // GroupBox 2: Manual Corner Radius
            GroupBox grpManual = new GroupBox();
            grpManual.Text = "Manual Corner Radius (Points / pt)";
            grpManual.Left = 16;
            grpManual.Top = 140;
            grpManual.Width = 490;
            grpManual.Height = 140;
            Controls.Add(grpManual);

            _chkEnableManualRadius = new CheckBox();
            _chkEnableManualRadius.Text = "Preset manual corner radius (used if none copied):";
            _chkEnableManualRadius.Left = 14;
            _chkEnableManualRadius.Top = 24;
            _chkEnableManualRadius.Width = 460;
            _chkEnableManualRadius.Height = 22;
            _chkEnableManualRadius.CheckedChanged += ChkEnableManualRadius_CheckedChanged;
            grpManual.Controls.Add(_chkEnableManualRadius);

            Label lblRadius = new Label();
            lblRadius.Text = "Radius value (pt):";
            lblRadius.Left = 34;
            lblRadius.Top = 56;
            lblRadius.Width = 120;
            lblRadius.Height = 22;
            grpManual.Controls.Add(lblRadius);

            _numManualRadius = new NumericUpDown();
            _numManualRadius.Left = 160;
            _numManualRadius.Top = 54;
            _numManualRadius.Width = 100;
            _numManualRadius.DecimalPlaces = 1;
            _numManualRadius.Minimum = 1;
            _numManualRadius.Maximum = 1000;
            _numManualRadius.Value = 20;
            grpManual.Controls.Add(_numManualRadius);

            _btnApplyManualNow = new Button();
            _btnApplyManualNow.Text = "Set Radius";
            _btnApplyManualNow.Left = 280;
            _btnApplyManualNow.Top = 52;
            _btnApplyManualNow.Width = 120;
            _btnApplyManualNow.Height = 26;
            _btnApplyManualNow.Click += BtnApplyManualNow_Click;
            grpManual.Controls.Add(_btnApplyManualNow);

            Label lblStatus = new Label();
            string activeStr = currentActiveRadius.HasValue ? (currentActiveRadius.Value.ToString("0.##") + " pt") : "None";
            lblStatus.Text = "Currently active radius in memory: " + activeStr;
            lblStatus.ForeColor = Color.DimGray;
            lblStatus.Left = 14;
            lblStatus.Top = 95;
            lblStatus.Width = 460;
            lblStatus.Height = 25;
            grpManual.Controls.Add(lblStatus);

            // Bottom Buttons
            _btnReset = new Button();
            _btnReset.Text = "Default";
            _btnReset.Left = 16;
            _btnReset.Top = 300;
            _btnReset.Width = 90;
            _btnReset.Height = 30;
            _btnReset.Click += BtnReset_Click;
            Controls.Add(_btnReset);

            _btnSave = new Button();
            _btnSave.Text = "Save";
            _btnSave.Left = 296;
            _btnSave.Top = 300;
            _btnSave.Width = 100;
            _btnSave.Height = 30;
            _btnSave.Click += BtnSave_Click;
            Controls.Add(_btnSave);

            _btnCancel = new Button();
            _btnCancel.Text = "Cancel";
            _btnCancel.Left = 406;
            _btnCancel.Top = 300;
            _btnCancel.Width = 100;
            _btnCancel.Height = 30;
            _btnCancel.DialogResult = DialogResult.Cancel;
            Controls.Add(_btnCancel);

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;

            LoadValuesToUi();
        }

        private void ChkEnableManualRadius_CheckedChanged(object sender, EventArgs e)
        {
            _numManualRadius.Enabled = _chkEnableManualRadius.Checked;
            _btnApplyManualNow.Enabled = _chkEnableManualRadius.Checked;
        }

        private void LoadValuesToUi()
        {
            _chkPlaySound.Checked = Settings.PlaySoundOnCopy;
            _chkShowMessage.Checked = Settings.ShowMessageOnCopy;

            if (Settings.ManualRadiusPoints.HasValue && Settings.ManualRadiusPoints.Value > 0)
            {
                _chkEnableManualRadius.Checked = true;
                _numManualRadius.Value = Math.Max(_numManualRadius.Minimum, Math.Min(_numManualRadius.Maximum, (decimal)Settings.ManualRadiusPoints.Value));
            }
            else
            {
                _chkEnableManualRadius.Checked = false;
                _numManualRadius.Value = 20;
            }

            _numManualRadius.Enabled = _chkEnableManualRadius.Checked;
            _btnApplyManualNow.Enabled = _chkEnableManualRadius.Checked;
        }

        private void BtnApplyManualNow_Click(object sender, EventArgs e)
        {
            ApplyManualNowRequested = true;
            SelectedRadiusValue = (float)_numManualRadius.Value;
            SaveUiToSettings();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            var def = new CornerRadiusSettings();
            _chkPlaySound.Checked = def.PlaySoundOnCopy;
            _chkShowMessage.Checked = def.ShowMessageOnCopy;
            _chkEnableManualRadius.Checked = false;
            _numManualRadius.Value = 20;
            _numManualRadius.Enabled = false;
            _btnApplyManualNow.Enabled = false;
        }

        private void SaveUiToSettings()
        {
            Settings.PlaySoundOnCopy = _chkPlaySound.Checked;
            Settings.ShowMessageOnCopy = _chkShowMessage.Checked;

            if (_chkEnableManualRadius.Checked)
            {
                Settings.ManualRadiusPoints = (float)_numManualRadius.Value;
                SelectedRadiusValue = (float)_numManualRadius.Value;
            }
            else
            {
                Settings.ManualRadiusPoints = null;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            SaveUiToSettings();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
