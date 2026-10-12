using System.Drawing;
using System.Windows.Forms;

namespace PowerPointJumpPointsAddIn
{
    public sealed class JumpPointSettingsForm : Form
    {
        private readonly CheckBox _showToggleMessagesCheckBox;
        private readonly CheckBox _playToggleSoundCheckBox;
        private readonly CheckBox _showNavigatorConfirmationsCheckBox;
        private readonly Button _okButton;
        private readonly Button _cancelButton;

        public bool ShowToggleMessages
        {
            get { return _showToggleMessagesCheckBox.Checked; }
        }

        public bool PlayToggleSound
        {
            get { return _playToggleSoundCheckBox.Checked; }
        }

        public bool ShowNavigatorConfirmations
        {
            get { return _showNavigatorConfirmationsCheckBox.Checked; }
        }

        public JumpPointSettingsForm(JumpPointSettings settings)
        {
            Text = "Jump Points Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(485, 235);

            Label titleLabel = new Label
            {
                Text = "Jump Point feedback and confirmations",
                AutoSize = true,
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                Location = new Point(16, 16)
            };

            Label infoLabel = new Label
            {
                Text = "Choose how the add-in should confirm actions and provide feedback.",
                AutoSize = false,
                Location = new Point(16, 42),
                Size = new Size(445, 34)
            };

            _showToggleMessagesCheckBox = new CheckBox
            {
                Text = "Show messages when setting/removing jump points",
                AutoSize = true,
                Location = new Point(19, 86),
                Checked = settings != null && settings.ShowToggleResultMessages
            };

            _playToggleSoundCheckBox = new CheckBox
            {
                Text = "Play sound when setting/removing jump points",
                AutoSize = true,
                Location = new Point(19, 114),
                Checked = settings == null || settings.PlayToggleResultSound
            };

            _showNavigatorConfirmationsCheckBox = new CheckBox
            {
                Text = "Show confirmation dialogs in Jump Points Navigator",
                AutoSize = true,
                Location = new Point(19, 142),
                Checked = settings == null || settings.ShowNavigatorConfirmations
            };

            _okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(303, 190),
                Size = new Size(80, 26)
            };

            _cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(389, 190),
                Size = new Size(80, 26)
            };

            Controls.Add(titleLabel);
            Controls.Add(infoLabel);
            Controls.Add(_showToggleMessagesCheckBox);
            Controls.Add(_playToggleSoundCheckBox);
            Controls.Add(_showNavigatorConfirmationsCheckBox);
            Controls.Add(_okButton);
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;
        }
    }
}
