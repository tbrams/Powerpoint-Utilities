﻿namespace WindowsForms_Automate_Powerpoint
{
    partial class Form2
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
            this.quitButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openGraphicsButton = new System.Windows.Forms.Button();
            this.InsertPictureButton = new System.Windows.Forms.Button();
            this.pictureListBox = new System.Windows.Forms.ListBox();
            this.btnMoveUp = new System.Windows.Forms.Button();
            this.btnMoveDown = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // quitButton
            // 
            this.quitButton.Location = new System.Drawing.Point(12, 332);
            this.quitButton.Name = "quitButton";
            this.quitButton.Size = new System.Drawing.Size(79, 38);
            this.quitButton.TabIndex = 0;
            this.quitButton.Text = "Quit";
            this.quitButton.UseVisualStyleBackColor = true;
            this.quitButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // openGraphicsButton
            // 
            this.openGraphicsButton.Location = new System.Drawing.Point(12, 13);
            this.openGraphicsButton.Name = "openGraphicsButton";
            this.openGraphicsButton.Size = new System.Drawing.Size(79, 38);
            this.openGraphicsButton.TabIndex = 2;
            this.openGraphicsButton.Text = "Pictures";
            this.openGraphicsButton.UseVisualStyleBackColor = true;
            this.openGraphicsButton.Click += new System.EventHandler(this.openGraphicsButton_Click);
            // 
            // InsertPictureButton
            // 
            this.InsertPictureButton.Location = new System.Drawing.Point(12, 57);
            this.InsertPictureButton.Name = "InsertPictureButton";
            this.InsertPictureButton.Size = new System.Drawing.Size(79, 38);
            this.InsertPictureButton.TabIndex = 4;
            this.InsertPictureButton.Text = "InsertPictures";
            this.InsertPictureButton.UseVisualStyleBackColor = true;
            this.InsertPictureButton.Click += new System.EventHandler(this.InsertPictureButton_Click);
            // 
            // pictureListBox
            // 
            this.pictureListBox.FormattingEnabled = true;
            this.pictureListBox.Location = new System.Drawing.Point(113, 15);
            this.pictureListBox.Name = "pictureListBox";
            this.pictureListBox.Size = new System.Drawing.Size(609, 355);
            this.pictureListBox.TabIndex = 5;
            this.pictureListBox.SelectedIndexChanged += new System.EventHandler(this.pictureListBox_SelectedIndexChanged);
            this.pictureListBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.pictureListBox_MouseDoubleClick);
            // 
            // btnMoveUp
            // 
            this.btnMoveUp.Location = new System.Drawing.Point(740, 15);
            this.btnMoveUp.Name = "btnMoveUp";
            this.btnMoveUp.Size = new System.Drawing.Size(75, 23);
            this.btnMoveUp.TabIndex = 6;
            this.btnMoveUp.Text = "Move up";
            this.btnMoveUp.UseVisualStyleBackColor = true;
            this.btnMoveUp.Click += new System.EventHandler(this.btnMoveUp_Click);
            // 
            // btnMoveDown
            // 
            this.btnMoveDown.Location = new System.Drawing.Point(740, 44);
            this.btnMoveDown.Name = "btnMoveDown";
            this.btnMoveDown.Size = new System.Drawing.Size(75, 23);
            this.btnMoveDown.TabIndex = 7;
            this.btnMoveDown.Text = "Move Down";
            this.btnMoveDown.UseVisualStyleBackColor = true;
            this.btnMoveDown.Click += new System.EventHandler(this.btnMoveDown_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(840, 382);
            this.Controls.Add(this.btnMoveDown);
            this.Controls.Add(this.btnMoveUp);
            this.Controls.Add(this.pictureListBox);
            this.Controls.Add(this.InsertPictureButton);
            this.Controls.Add(this.openGraphicsButton);
            this.Controls.Add(this.quitButton);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button quitButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button openGraphicsButton;
        private System.Windows.Forms.Button InsertPictureButton;
        private System.Windows.Forms.ListBox pictureListBox;
        private System.Windows.Forms.Button btnMoveUp;
        private System.Windows.Forms.Button btnMoveDown;
    }
}