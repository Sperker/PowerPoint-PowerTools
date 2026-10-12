using System.Drawing;
using System.Windows.Forms;

namespace PowerPointJumpPointsAddIn
{
    public sealed class JumpPointRenameForm : Form
    {
        private readonly TextBox _nameTextBox;
        private readonly Button _okButton;
        private readonly Button _cancelButton;

        public string JumpPointName
        {
            get { return _nameTextBox.Text.Trim(); }
        }

        public JumpPointRenameForm(string currentName, int slideIndex)
        {
            Text = "Rename Jump Point";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(440, 150);

            Label infoLabel = new Label
            {
                Text = "Jump point name for slide " + slideIndex + ":",
                AutoSize = true,
                Location = new Point(16, 18)
            };

            _nameTextBox = new TextBox
            {
                Location = new Point(19, 46),
                Size = new Size(402, 22),
                Text = currentName ?? string.Empty
            };

            Label hintLabel = new Label
            {
                Text = "Leave empty to use the detected slide title / first text again.",
                AutoSize = true,
                Location = new Point(19, 76),
                ForeColor = SystemColors.GrayText
            };

            _okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(255, 108),
                Size = new Size(80, 26)
            };

            _cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(341, 108),
                Size = new Size(80, 26)
            };

            Controls.Add(infoLabel);
            Controls.Add(_nameTextBox);
            Controls.Add(hintLabel);
            Controls.Add(_okButton);
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;

            Load += delegate
            {
                _nameTextBox.SelectAll();
                _nameTextBox.Focus();
            };
        }
    }
}
