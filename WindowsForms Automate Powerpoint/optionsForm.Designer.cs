namespace WindowsForms_Automate_Powerpoint
{
    partial class optionsForm
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
            this.buttonOK = new System.Windows.Forms.Button();
            this.txtFirstDelay = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtFirstDuration = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtDuration = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtDelay = new System.Windows.Forms.TextBox();
            this.comboBoxAnimation = new System.Windows.Forms.ComboBox();
            this.comboBoxStart = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.checkBoxFirstImage = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(521, 164);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(74, 35);
            this.buttonOK.TabIndex = 0;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtFirstDelay
            // 
            this.txtFirstDelay.Location = new System.Drawing.Point(156, 55);
            this.txtFirstDelay.Name = "txtFirstDelay";
            this.txtFirstDelay.Size = new System.Drawing.Size(79, 20);
            this.txtFirstDelay.TabIndex = 1;
            this.txtFirstDelay.Text = "0.0";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 58);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Delay (seconds)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(34, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "First Slide";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(34, 83);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(96, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Duration (seconds)";
            // 
            // txtFirstDuration
            // 
            this.txtFirstDuration.Location = new System.Drawing.Point(156, 80);
            this.txtFirstDuration.Name = "txtFirstDuration";
            this.txtFirstDuration.Size = new System.Drawing.Size(79, 20);
            this.txtFirstDuration.TabIndex = 4;
            this.txtFirstDuration.Text = "0.5";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(34, 152);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Duration (seconds)";
            // 
            // txtDuration
            // 
            this.txtDuration.Location = new System.Drawing.Point(156, 149);
            this.txtDuration.Name = "txtDuration";
            this.txtDuration.Size = new System.Drawing.Size(79, 20);
            this.txtDuration.TabIndex = 9;
            this.txtDuration.Text = "1.75";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(34, 111);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Following Slides";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(34, 127);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(83, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "Delay (seconds)";
            // 
            // txtDelay
            // 
            this.txtDelay.Location = new System.Drawing.Point(156, 124);
            this.txtDelay.Name = "txtDelay";
            this.txtDelay.Size = new System.Drawing.Size(79, 20);
            this.txtDelay.TabIndex = 6;
            this.txtDelay.Text = "2.0";
            // 
            // comboBoxAnimation
            // 
            this.comboBoxAnimation.FormattingEnabled = true;
            this.comboBoxAnimation.Location = new System.Drawing.Point(370, 55);
            this.comboBoxAnimation.Name = "comboBoxAnimation";
            this.comboBoxAnimation.Size = new System.Drawing.Size(225, 21);
            this.comboBoxAnimation.TabIndex = 11;
            this.comboBoxAnimation.SelectedIndexChanged += new System.EventHandler(this.comboBoxAnimation_SelectedIndexChanged);
            // 
            // comboBoxStart
            // 
            this.comboBoxStart.FormattingEnabled = true;
            this.comboBoxStart.Location = new System.Drawing.Point(370, 80);
            this.comboBoxStart.Name = "comboBoxStart";
            this.comboBoxStart.Size = new System.Drawing.Size(225, 21);
            this.comboBoxStart.TabIndex = 13;
            this.comboBoxStart.SelectedIndexChanged += new System.EventHandler(this.comboBoxStart_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(300, 58);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Animation";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(304, 83);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(44, 13);
            this.label8.TabIndex = 15;
            this.label8.Text = "Start on";
            // 
            // checkBoxFirstImage
            // 
            this.checkBoxFirstImage.AutoSize = true;
            this.checkBoxFirstImage.Checked = true;
            this.checkBoxFirstImage.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxFirstImage.Location = new System.Drawing.Point(307, 123);
            this.checkBoxFirstImage.Name = "checkBoxFirstImage";
            this.checkBoxFirstImage.Size = new System.Drawing.Size(174, 17);
            this.checkBoxFirstImage.TabIndex = 16;
            this.checkBoxFirstImage.Text = "First Image Appear autimatically";
            this.checkBoxFirstImage.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(300, 42);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(59, 13);
            this.label9.TabIndex = 17;
            this.label9.Text = "All Slides";
            // 
            // optionsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(625, 221);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.checkBoxFirstImage);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.comboBoxStart);
            this.Controls.Add(this.comboBoxAnimation);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtDuration);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtDelay);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtFirstDuration);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtFirstDelay);
            this.Controls.Add(this.buttonOK);
            this.Name = "optionsForm";
            this.Text = "Slideshow Timing  Options";
            this.Load += new System.EventHandler(this.optionsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.TextBox txtFirstDelay;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtFirstDuration;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtDuration;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtDelay;
        private System.Windows.Forms.ComboBox comboBoxAnimation;
        private System.Windows.Forms.ComboBox comboBoxStart;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.CheckBox checkBoxFirstImage;
        private System.Windows.Forms.Label label9;
    }
}