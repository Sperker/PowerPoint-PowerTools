using System;
using System.Windows.Forms;

namespace PPTMissingEdgeAlignAddIn
{
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
            // Load current settings
            numGap.Value = Properties.Settings.Default.DefaultGap;
            chkSoundArmed.Checked = Properties.Settings.Default.SoundOnArmed;
            chkSoundMissed.Checked = Properties.Settings.Default.SoundOnMissedClick;
            chkSoundError.Checked = Properties.Settings.Default.SoundOnError;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.DefaultGap = (int)numGap.Value;
            Properties.Settings.Default.SoundOnArmed = chkSoundArmed.Checked;
            Properties.Settings.Default.SoundOnMissedClick = chkSoundMissed.Checked;
            Properties.Settings.Default.SoundOnError = chkSoundError.Checked;
            Properties.Settings.Default.Save();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
