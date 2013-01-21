using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Reflection;


namespace WindowsForms_Automate_Powerpoint
{
    public partial class Form1 : Form
    {
        string fileName;
        List<string> myComments = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }


        private void openFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Powerpoint|*.ppt;*.pptx|All files|*.*";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Title = "Please select a presentation";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                fileName = dialog.FileName;

                this.Text = this.Text + ":" + System.IO.Path.GetFileName(fileName);
                this.Size = new Size(887,306);

                listPowerPointNotes();
            }
            else
                Console.WriteLine("Something went wrong reading the filename from the Dialog");

            buttonQuit.Visible = true;
        }

        private void listPowerPointNotes()
        {
            PowerPoint.Application oPowerPoint = null;
            PowerPoint.Presentations oPres = null;
            PowerPoint.Presentation oPre = null;

            //
            try
            {
                oPowerPoint = new PowerPoint.Application();

                // By default PowerPoint is invisible, till you make it visible.
                // oPowerPoint.Visible = Office.MsoTriState.msoTrue;

                oPres = oPowerPoint.Presentations;
                oPre = oPres.Open(fileName, MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);

                int slidenum = 0;

                // Set cursor to indicate this may take time...
                Cursor.Current = Cursors.WaitCursor;
                foreach (PowerPoint.Slide mySlide in oPre.Slides)
                {
                    slidenum++;
                    string s = mySlide.NotesPage.Shapes[2].TextFrame.TextRange.Text;

                    if (s != "")
                    {
                        myComments.Add("Slide: " + slidenum);
                        myComments.Add(s);
                        myComments.Add("");
                        myComments.Add("");
                    }

                    // this is for removing comments:
                    // mySlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "removed ....";
                }


                if (slidenum!=0) {
                    foreach (string myString in myComments)
                    {
                        textboxComments.Text += myString + Environment.NewLine;
                    }

                    label1.Visible = true;
                    textboxComments.Visible = true;
                    buttonBackup.Visible = true;
                    buttonStrip.Visible=true;
                }

                // reset cursor 
                Cursor.Current = Cursors.Default;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Trying to open Powerpoint throws the error: {0}",
                    ex.Message);
            } 

        }


 
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            // Get file name.
            string name = saveFileDialog1.FileName;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textboxComments.Visible = false;
            label1.Visible = false;
            buttonBackup.Visible = false;
            buttonStrip.Visible = false;
            buttonQuit.Visible = false;

        }

        private void buttonBackup_Click(object sender, EventArgs e)
        {
            // FIRST: Get file name from dialog

            // Set the properties on SaveFileDialog1 so the user is  
            // prompted to overwrite the file if it does exist.
            // -- SaveFileDialog1.CreatePrompt = true;
            saveFileDialog1.OverwritePrompt = true;

            // Set the file name to new presentation.pptx, set the type filter 
            // to PowerPoint files, and set the initial directory to the  
            // Desktop folder.
            saveFileDialog1.FileName = "slide notes backup";
            // DefaultExt is only used when "All files" is selected from  
            // the filter box and no extension is specified by the user.
            saveFileDialog1.DefaultExt = "txt";
            saveFileDialog1.Filter =
                "Text (*.txt)|*.txt;|All files|*.*";
            saveFileDialog1.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Call ShowDialog and check for a return value of DialogResult.OK, 
            // which indicates that the file was saved. 
            DialogResult result = saveFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {       
                // save content of comment array to file

                StreamWriter file = new System.IO.StreamWriter(saveFileDialog1.FileName);
                foreach (string line in myComments)
                    file.WriteLine(line);
                file.Close();
            }
            buttonBackup.Enabled = false;
         }

        private void buttonStrip_Click(object sender, EventArgs e)
        {
            // remove all slide notes and save clean copy named: "..._stripped.pptx"

            PowerPoint.Application oPowerPoint = null;
            PowerPoint.Presentations oPres = null;
            PowerPoint.Presentation oPre = null;

            //
            try
            {
                oPowerPoint = new PowerPoint.Application();

                // By default PowerPoint is invisible, till you make it visible.
                // oPowerPoint.Visible = Office.MsoTriState.msoTrue;

                oPres = oPowerPoint.Presentations;
                oPre = oPres.Open(fileName, ReadOnly:MsoTriState.msoFalse, WithWindow:MsoTriState.msoFalse);

                // Set cursor to indicate this may take time...
                Cursor.Current = Cursors.WaitCursor;

                foreach (PowerPoint.Slide mySlide in oPre.Slides)
                {
                    string s = mySlide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                    if (s != "")
                    {
                        mySlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "";
                    }
                }

                string shortFileName= System.IO.Path.GetFileName(fileName);
                string saveDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                int decimalPoint = shortFileName.LastIndexOf(".");
                string fileType = shortFileName.Substring(decimalPoint);
                string sFileName = shortFileName.Substring(0, decimalPoint);
                string newFileName = saveDirectory + "\\" + sFileName + "_stripped" + fileType;

                oPre.SaveAs(newFileName);

                // reset cursor 
                Cursor.Current = Cursors.Default;
                buttonStrip.Enabled = false;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Trying to open Powerpoint throws the error: {0}",
                    ex.Message);
            }

        }

        private void clearPowerPointNotes()
        {
            PowerPoint.Application oPowerPoint = null;
            PowerPoint.Presentations oPres = null;
            PowerPoint.Presentation oPre = null;

            //
            try
            {
                oPowerPoint = new PowerPoint.Application();

                // By default PowerPoint is invisible, till you make it visible.
                // oPowerPoint.Visible = Office.MsoTriState.msoTrue;

                oPres = oPowerPoint.Presentations;
                oPre = oPres.Open(fileName, MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);

                foreach (PowerPoint.Slide mySlide in oPre.Slides)
                {
                    string s = mySlide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                    if (s != "")
                    {
                        mySlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "removed ....";
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Trying to open Powerpoint throws the error: {0}", ex.Message);
            }

            oPre.Save();

            //
        }

        private void buttonQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
