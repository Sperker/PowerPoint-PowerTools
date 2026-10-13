using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal sealed class SettingsForm : Form
    {
        private readonly CheckBox _chkEnableVirtualCanvas;
        private readonly NumericUpDown _numCanvasPadding;
        private readonly CheckBox _chkBlockKeyboard;
        private readonly NumericUpDown _numWheelLines;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;
        private readonly Button _btnReset;

        public StayOnSlideSettings Settings { get; private set; }

        public SettingsForm(StayOnSlideSettings settings)
        {
            Settings = settings ?? new StayOnSlideSettings();

            Text = AddInInfo.ProductName + " Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 580;
            Height = 440;

            Label lblTitle = new Label();
            lblTitle.Text = AddInInfo.ProductName + " v" + AddInInfo.Version;
            lblTitle.Font = new Font(lblTitle.Font, FontStyle.Bold);
            lblTitle.Left = 16;
            lblTitle.Top = 14;
            lblTitle.Width = 530;
            lblTitle.Height = 24;
            Controls.Add(lblTitle);

            // GroupBox 1: Pasteboard
            GroupBox grpCanvas = new GroupBox();
            grpCanvas.Text = "Pasteboard (Virtual Canvas) & Slide Boundaries";
            grpCanvas.Left = 16;
            grpCanvas.Top = 45;
            grpCanvas.Width = 530;
            grpCanvas.Height = 140;
            Controls.Add(grpCanvas);

            _chkEnableVirtualCanvas = new CheckBox();
            _chkEnableVirtualCanvas.Text = "Enable extended pasteboard (work beyond slide boundaries)";
            _chkEnableVirtualCanvas.Left = 14;
            _chkEnableVirtualCanvas.Top = 24;
            _chkEnableVirtualCanvas.Width = 500;
            _chkEnableVirtualCanvas.Height = 22;
            grpCanvas.Controls.Add(_chkEnableVirtualCanvas);

            Label lblPadding = new Label();
            lblPadding.Text = "Top & bottom padding size (points, 200 - 5000):";
            lblPadding.Left = 14;
            lblPadding.Top = 55;
            lblPadding.Width = 350;
            lblPadding.Height = 20;
            grpCanvas.Controls.Add(lblPadding);

            _numCanvasPadding = new NumericUpDown();
            _numCanvasPadding.Left = 380;
            _numCanvasPadding.Top = 52;
            _numCanvasPadding.Width = 120;
            _numCanvasPadding.Minimum = 200;
            _numCanvasPadding.Maximum = 5000;
            _numCanvasPadding.Increment = 100;
            grpCanvas.Controls.Add(_numCanvasPadding);

            Label lblCanvasHint = new Label();
            lblCanvasHint.Text = "Note: Allows free scrolling and placing objects far beyond the top and bottom slide edges without needing manual helper shapes.";
            lblCanvasHint.ForeColor = Color.DimGray;
            lblCanvasHint.Left = 14;
            lblCanvasHint.Top = 85;
            lblCanvasHint.Width = 500;
            lblCanvasHint.Height = 45;
            grpCanvas.Controls.Add(lblCanvasHint);

            // GroupBox 2: Navigation
            GroupBox grpNav = new GroupBox();
            grpNav.Text = "Navigation & Keyboard Control";
            grpNav.Left = 16;
            grpNav.Top = 195;
            grpNav.Width = 530;
            grpNav.Height = 145;
            Controls.Add(grpNav);

            _chkBlockKeyboard = new CheckBox();
            _chkBlockKeyboard.Text = "Block slide switching via keyboard (PageUp / PageDown) on slide";
            _chkBlockKeyboard.Left = 14;
            _chkBlockKeyboard.Top = 24;
            _chkBlockKeyboard.Width = 500;
            _chkBlockKeyboard.Height = 22;
            grpNav.Controls.Add(_chkBlockKeyboard);

            Label lblWheelLines = new Label();
            lblWheelLines.Text = "Mouse wheel scroll step on slide (lines):";
            lblWheelLines.Left = 14;
            lblWheelLines.Top = 55;
            lblWheelLines.Width = 350;
            lblWheelLines.Height = 20;
            grpNav.Controls.Add(lblWheelLines);

            _numWheelLines = new NumericUpDown();
            _numWheelLines.Left = 380;
            _numWheelLines.Top = 52;
            _numWheelLines.Width = 120;
            _numWheelLines.Minimum = 1;
            _numWheelLines.Maximum = 20;
            grpNav.Controls.Add(_numWheelLines);

            Label lblNavHint = new Label();
            lblNavHint.Text = "Tip: When Stay on Slide is active, switch slides intentionally using the left thumbnail sidebar.";
            lblNavHint.ForeColor = Color.DimGray;
            lblNavHint.Left = 14;
            lblNavHint.Top = 85;
            lblNavHint.Width = 500;
            lblNavHint.Height = 45;
            grpNav.Controls.Add(lblNavHint);

            // Bottom Buttons
            _btnReset = new Button();
            _btnReset.Text = "Reset";
            _btnReset.Left = 16;
            _btnReset.Top = 355;
            _btnReset.Width = 100;
            _btnReset.Height = 30;
            _btnReset.Click += BtnReset_Click;
            Controls.Add(_btnReset);

            _btnSave = new Button();
            _btnSave.Text = "Save";
            _btnSave.Left = 336;
            _btnSave.Top = 355;
            _btnSave.Width = 100;
            _btnSave.Height = 30;
            _btnSave.DialogResult = DialogResult.OK;
            _btnSave.Click += BtnSave_Click;
            Controls.Add(_btnSave);

            _btnCancel = new Button();
            _btnCancel.Text = "Cancel";
            _btnCancel.Left = 446;
            _btnCancel.Top = 355;
            _btnCancel.Width = 100;
            _btnCancel.Height = 30;
            _btnCancel.DialogResult = DialogResult.Cancel;
            Controls.Add(_btnCancel);

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;

            LoadValuesToUi();
        }

        private void LoadValuesToUi()
        {
            _chkEnableVirtualCanvas.Checked = Settings.EnableVirtualCanvas;
            _numCanvasPadding.Value = Math.Max(_numCanvasPadding.Minimum, Math.Min(_numCanvasPadding.Maximum, Settings.CanvasPaddingPoints));
            _chkBlockKeyboard.Checked = Settings.BlockKeyboardSlideNavigation;
            _numWheelLines.Value = Math.Max(_numWheelLines.Minimum, Math.Min(_numWheelLines.Maximum, Settings.WheelScrollStepLines));
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            var def = new StayOnSlideSettings();
            _chkEnableVirtualCanvas.Checked = def.EnableVirtualCanvas;
            _numCanvasPadding.Value = def.CanvasPaddingPoints;
            _chkBlockKeyboard.Checked = def.BlockKeyboardSlideNavigation;
            _numWheelLines.Value = def.WheelScrollStepLines;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            Settings.EnableVirtualCanvas = _chkEnableVirtualCanvas.Checked;
            Settings.CanvasPaddingPoints = (int)_numCanvasPadding.Value;
            Settings.BlockKeyboardSlideNavigation = _chkBlockKeyboard.Checked;
            Settings.WheelScrollStepLines = (int)_numWheelLines.Value;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
