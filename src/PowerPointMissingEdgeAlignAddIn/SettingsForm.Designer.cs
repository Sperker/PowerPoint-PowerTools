namespace PPTMissingEdgeAlignAddIn
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpAlignment = new System.Windows.Forms.GroupBox();
            this.lblInstruction = new System.Windows.Forms.Label();
            this.numGap = new System.Windows.Forms.NumericUpDown();
            this.chkSoundArmed = new System.Windows.Forms.CheckBox();
            this.chkSoundMissed = new System.Windows.Forms.CheckBox();
            this.chkSoundError = new System.Windows.Forms.CheckBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.grpAlignment.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numGap)).BeginInit();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(16, 14);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(183, 13);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "PPTMissingEdgeAlignAddIn v.1";
            // 
            // grpAlignment
            // 
            this.grpAlignment.Controls.Add(this.lblInstruction);
            this.grpAlignment.Controls.Add(this.numGap);
            this.grpAlignment.Controls.Add(this.chkSoundArmed);
            this.grpAlignment.Controls.Add(this.chkSoundMissed);
            this.grpAlignment.Controls.Add(this.chkSoundError);
            this.grpAlignment.Location = new System.Drawing.Point(16, 42);
            this.grpAlignment.Name = "grpAlignment";
            this.grpAlignment.Size = new System.Drawing.Size(350, 165);
            this.grpAlignment.TabIndex = 1;
            this.grpAlignment.TabStop = false;
            this.grpAlignment.Text = "Alignment settings";
            // 
            // lblInstruction
            // 
            this.lblInstruction.AutoSize = true;
            this.lblInstruction.Location = new System.Drawing.Point(14, 26);
            this.lblInstruction.Name = "lblInstruction";
            this.lblInstruction.Size = new System.Drawing.Size(182, 13);
            this.lblInstruction.TabIndex = 0;
            this.lblInstruction.Text = "Default gap for CTRL modifier in pt:";
            // 
            // numGap
            // 
            this.numGap.Location = new System.Drawing.Point(14, 50);
            this.numGap.Maximum = new decimal(new int[] {
            200,
            0,
            0,
            0});
            this.numGap.Name = "numGap";
            this.numGap.Size = new System.Drawing.Size(120, 20);
            this.numGap.TabIndex = 1;
            // 
            // chkSoundArmed
            // 
            this.chkSoundArmed.AutoSize = true;
            this.chkSoundArmed.Location = new System.Drawing.Point(14, 80);
            this.chkSoundArmed.Name = "chkSoundArmed";
            this.chkSoundArmed.Size = new System.Drawing.Size(135, 17);
            this.chkSoundArmed.TabIndex = 2;
            this.chkSoundArmed.Text = "Play sound when armed";
            this.chkSoundArmed.UseVisualStyleBackColor = true;
            // 
            // chkSoundMissed
            // 
            this.chkSoundMissed.AutoSize = true;
            this.chkSoundMissed.Location = new System.Drawing.Point(14, 105);
            this.chkSoundMissed.Name = "chkSoundMissed";
            this.chkSoundMissed.Size = new System.Drawing.Size(184, 17);
            this.chkSoundMissed.TabIndex = 3;
            this.chkSoundMissed.Text = "Play sound on missed target click";
            this.chkSoundMissed.UseVisualStyleBackColor = true;
            // 
            // chkSoundError
            // 
            this.chkSoundError.AutoSize = true;
            this.chkSoundError.Location = new System.Drawing.Point(14, 130);
            this.chkSoundError.Name = "chkSoundError";
            this.chkSoundError.Size = new System.Drawing.Size(181, 17);
            this.chkSoundError.TabIndex = 4;
            this.chkSoundError.Text = "Play sound on error / invalid action";
            this.chkSoundError.UseVisualStyleBackColor = true;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(180, 220);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(90, 24);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(276, 220);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 24);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.btnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(384, 260);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.grpAlignment);
            this.Controls.Add(this.lblTitle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "PPTMissingEdgeAlignAddIn Settings";
            this.grpAlignment.ResumeLayout(false);
            this.grpAlignment.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numGap)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.GroupBox grpAlignment;
        private System.Windows.Forms.Label lblInstruction;
        private System.Windows.Forms.NumericUpDown numGap;
        private System.Windows.Forms.CheckBox chkSoundArmed;
        private System.Windows.Forms.CheckBox chkSoundMissed;
        private System.Windows.Forms.CheckBox chkSoundError;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
    }
}
