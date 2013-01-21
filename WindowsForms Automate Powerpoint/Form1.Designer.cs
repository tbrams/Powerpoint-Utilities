namespace WindowsForms_Automate_Powerpoint
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.buttonStrip = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.buttonBackup = new System.Windows.Forms.Button();
            this.labelComments = new System.Windows.Forms.Label();
            this.textboxComments = new System.Windows.Forms.RichTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonQuit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(34, 14);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(110, 43);
            this.openFileButton.TabIndex = 0;
            this.openFileButton.Text = "Open";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.openFileButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // buttonStrip
            // 
            this.buttonStrip.Location = new System.Drawing.Point(34, 141);
            this.buttonStrip.Name = "buttonStrip";
            this.buttonStrip.Size = new System.Drawing.Size(110, 43);
            this.buttonStrip.TabIndex = 3;
            this.buttonStrip.Text = "Strip";
            this.buttonStrip.UseVisualStyleBackColor = true;
            this.buttonStrip.Click += new System.EventHandler(this.buttonStrip_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // buttonBackup
            // 
            this.buttonBackup.Location = new System.Drawing.Point(34, 92);
            this.buttonBackup.Name = "buttonBackup";
            this.buttonBackup.Size = new System.Drawing.Size(110, 43);
            this.buttonBackup.TabIndex = 6;
            this.buttonBackup.Text = "Backup";
            this.buttonBackup.UseVisualStyleBackColor = true;
            this.buttonBackup.Click += new System.EventHandler(this.buttonBackup_Click);
            // 
            // labelComments
            // 
            this.labelComments.AutoSize = true;
            this.labelComments.Location = new System.Drawing.Point(544, 120);
            this.labelComments.Name = "labelComments";
            this.labelComments.Size = new System.Drawing.Size(78, 17);
            this.labelComments.TabIndex = 7;
            this.labelComments.Text = "Comments:";
            // 
            // textboxComments
            // 
            this.textboxComments.Location = new System.Drawing.Point(164, 14);
            this.textboxComments.Name = "textboxComments";
            this.textboxComments.Size = new System.Drawing.Size(693, 229);
            this.textboxComments.TabIndex = 8;
            this.textboxComments.Text = "";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(38, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 17);
            this.label1.TabIndex = 9;
            this.label1.Text = "Comments:";
            // 
            // buttonQuit
            // 
            this.buttonQuit.Location = new System.Drawing.Point(34, 200);
            this.buttonQuit.Name = "buttonQuit";
            this.buttonQuit.Size = new System.Drawing.Size(110, 43);
            this.buttonQuit.TabIndex = 10;
            this.buttonQuit.Text = "Quit";
            this.buttonQuit.UseVisualStyleBackColor = true;
            this.buttonQuit.Click += new System.EventHandler(this.buttonQuit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(291, 261);
            this.Controls.Add(this.buttonQuit);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textboxComments);
            this.Controls.Add(this.labelComments);
            this.Controls.Add(this.buttonBackup);
            this.Controls.Add(this.buttonStrip);
            this.Controls.Add(this.openFileButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "PowerPoint Cleanup";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button buttonStrip;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button buttonBackup;
        private System.Windows.Forms.RichTextBox textboxComments;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelComments;
        private System.Windows.Forms.Button buttonQuit;
    }
}

