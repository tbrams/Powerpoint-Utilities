using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace WindowsForms_Automate_Powerpoint
{

    public partial class FormAutomateImageAnimation : Form
    {
        string fileName;

        // Animation Settings
        // Timing 
        public float myFirstDelay = 0;
        public float myFirstDuration = 0.5F;
        public float myDuration = 1.75F;
        public float myDelay = 2.0F;

        // myTrigger is used between pictures - change to onClick if you want control here
        public PowerPoint.MsoAnimTriggerType myTrigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;

        // Animations are kept simple, normal effeft is fading, feel free to select an alternative
        public PowerPoint.MsoAnimEffect myAnimation = PowerPoint.MsoAnimEffect.msoAnimEffectFade;

        // Normally first image will just appear. If not, set the following to false
        public Boolean FirstImageAppearWithPrevious = true;

        public FormAutomateImageAnimation()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void Form2_Load(object sender, EventArgs e)
        {
            quitButton.Visible = false;
            InsertPictureButton.Visible = false;
            optionButton.Visible = false;
            pictureListBox.Visible = false;
            btnMoveDown.Visible = false;
            btnMoveUp.Visible = false;
            lblRemove.Visible = false;

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form2_DragEnter);
            this.DragDrop += new DragEventHandler(Form2_DragDrop);

        }

        void Form2_DragEnter(object sender, DragEventArgs e)
        {
            if (pictureListBox.Visible)
               if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Form2_DragDrop(object sender, DragEventArgs e)
        {
            if (pictureListBox.Visible)
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    string dragFileType = System.IO.Path.GetExtension(file);
                    if (dragFileType.Equals(".png", StringComparison.OrdinalIgnoreCase))
                    {
                        Debug.WriteLine("Dragged this into listbox: " + file);
                        pictureListBox.Items.Add(file);
                    }
                    else
                    {
                        // Msgbox with filetype support flag goes here...
                        MessageBox.Show("This image type is currently not supported", "PNG Images only");
                    }
                }
            }
        }

        private void openGraphicsButton_Click(object sender, EventArgs e)
        {
            createPowerPoint();

            // Clean up the unmanaged PowerPoint COM resources by forcing a  
            // garbage collection as soon as the calling function is off the  
            // stack (at which point these objects are no longer rooted). 

            GC.Collect();
            GC.WaitForPendingFinalizers();
            // GC needs to be called twice in order to get the Finalizers called  
            // - the first time in, it simply makes a list of what is to be  
            // finalized, the second time in, it actually is finalizing. Only  
            // then will the object do its automatic ReleaseComObject. 
            GC.Collect();
            GC.WaitForPendingFinalizers(); 

        }

        private void createPowerPoint()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Graphics|*.png;*.jpg;*.gif;*.bmp|All files|*.*";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            dialog.Title = "Please select first graphics file";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                fileName = dialog.FileName;

                this.Text = this.Text + ":" + System.IO.Path.GetFileName(fileName);
                // this.Size = new Size(900, 500);

                string localFile = System.IO.Path.GetFileNameWithoutExtension(fileName);
                localFile=localFile.Replace("1", ""); // Just in case this is the first in a series
               // pictureBox1.Image = Image.FromFile(fileName);

                // check if there are more than one picture available
                string folder = System.IO.Path.GetDirectoryName(fileName);
                string[] extraPics = Directory.GetFiles(folder, localFile+"*.png");
                pictureListBox.Items.AddRange(extraPics);
                pictureListBox.Visible = true;
                btnMoveUp.Visible = true;
                btnMoveUp.Enabled = false;
                btnMoveDown.Visible = true;
                btnMoveDown.Enabled = false;
                lblRemove.Visible = true;
            }
            else
                Debug.WriteLine("Something went wrong reading the filename from the Dialog");

            quitButton.Visible = true;
            optionButton.Visible = true;
            InsertPictureButton.Visible = true;
        }

        private void InsertPictureButton_Click(object sender, EventArgs e)
        {
            // Insert pictures from listbox one by one

            try
            {
                // Create an instance of Microsoft PowerPoint and make it  
                // invisible. (Not allowed???)

                PowerPoint.Application oPowerPoint = new PowerPoint.Application();
                oPowerPoint.Visible = Office.MsoTriState.msoFalse; 

                // Create a new Presentation. 
                PowerPoint.Presentation oPre = oPowerPoint.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                Debug.WriteLine("New presentation created");

                // Insert a new Slide with the picture name as title and a short explanation. 
                string slideTitle = pictureListBox.Items[0].ToString();
                slideTitle = System.IO.Path.GetFileNameWithoutExtension(slideTitle);

                PowerPoint.Slide oSlide = oPre.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutText);
                oSlide.Shapes[1].TextFrame.TextRange.Text = "Slides about "+slideTitle;
                oSlide.Shapes[2].TextFrame.TextRange.Text = "Explanation about the build nature of the slide...";



                // Make the slide hidden, so it will not appear in the slide show
                oSlide.SlideShowTransition.Hidden = MsoTriState.msoTrue;
                

                // Insert a new blank Slide for the picture. 
                oSlide = oPre.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
                Debug.WriteLine("Blank slide for images inserted");



                Boolean FirstImage = true;  // Used to handle special effect for very first picture
                foreach (string p in pictureListBox.Items)
                {                    
                    string imageFileName = p;

                    mcImageDetails imageDetails = new mcImageDetails(imageFileName);
                    float dx = imageDetails.dx;
                    float dy = imageDetails.dy;
                    double imageWidth = imageDetails.imageWidth;
                    double imageHeight = imageDetails.imageHeight;

                    PowerPoint.Shape myPic;
                    myPic = oSlide.Shapes.AddPicture(imageFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 1, 1, -1, -1);
                 
                    const float PixelPerPoint=(float)4/3;
                    double slideWidth = oPowerPoint.ActivePresentation.PageSetup.SlideWidth * PixelPerPoint;
                    double slideHeight = oPowerPoint.ActivePresentation.PageSetup.SlideHeight * PixelPerPoint;


                    // Scale to fit slide area
                    double xfactor = 1.0;
                    double ppp = dx / 96;

                    if (imageWidth> imageHeight) {
                          // X Scaling
                          xfactor=(slideWidth/imageWidth)*ppp;
                          myPic.ScaleWidth((float)xfactor, Office.MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    }
                    else
                    {
                          // Y Scaling
                          xfactor = (slideHeight / imageHeight)*ppp;
                          myPic.ScaleHeight((float)xfactor, Office.MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    }


                    Debug.WriteLine("Image w x h:" + imageWidth + ", " + imageHeight);
                    Debug.WriteLine("Slide  w x h: " + slideWidth + ", " + slideHeight);
                    Debug.WriteLine("Xfactor : " + xfactor);
                   
                    myPic.Left = oPowerPoint.ActivePresentation.PageSetup.SlideWidth / 2 - myPic.Width / 2;
                    myPic.Top = oPowerPoint.ActivePresentation.PageSetup.SlideHeight / 2 - myPic.Height / 2;

                    // Add animation to the screenshot
                    // Remove the trigger for first image, so that it will appear automatically
                    PowerPoint.MsoAnimTriggerType thisTrigger; 
                    thisTrigger=myTrigger;
                    if (FirstImage && FirstImageAppearWithPrevious) {
                        thisTrigger=PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    } 
                    PowerPoint.Effect myEffect;
                    myEffect = oSlide.TimeLine.MainSequence.AddEffect(myPic, myAnimation, trigger: myTrigger);

                    // Include timing etc
                    myEffect.Timing.Duration = myDuration;
                    myEffect.Timing.TriggerDelayTime = myDelay;

                    if (myTrigger == PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious) {
                        if (FirstImage) {
                            if (myAnimation != PowerPoint.MsoAnimEffect.msoAnimEffectAppear) {
                                myEffect.Timing.Duration=myFirstDuration;
                                myEffect.Timing.TriggerDelayTime=myFirstDelay;
                            }                        
                        }   else {
                                myEffect.Timing.Duration=myDuration;
                                myEffect.Timing.TriggerDelayTime=myDelay;                        
                        }                 
                    } 

                    FirstImage=false;
                }

                // Save the presentation as a pptx file and close it. 
                Debug.WriteLine("Save and close the presentation");

                string shortFileName = System.IO.Path.GetFileName(fileName);
                string saveDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                int decimalPoint = shortFileName.LastIndexOf(".");
                string fileType = shortFileName.Substring(decimalPoint);
                string sFileName = shortFileName.Substring(0, decimalPoint);
                string newFileName = saveDirectory + "\\" + sFileName + ".pptx";

                oPre.SaveAs(newFileName,
                    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    Office.MsoTriState.msoTriStateMixed);
                oPre.Close();

                // Quit the PowerPoint application. 
                Debug.WriteLine("Quit the PowerPoint application");
                oPowerPoint.Quit();

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Solution1.AutomatePowerPoint throws the error: {0}",
                    ex.Message);
            }
        }

 
        private void pictureListBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Debug.WriteLine("Doubleclick on: " + pictureListBox.SelectedItem);
            pictureListBox.Items.RemoveAt(pictureListBox.SelectedIndex);
            btnMoveDown.Enabled = false;
            btnMoveUp.Enabled = false;
        }

        private void pictureListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (pictureListBox.SelectedItem !=null) {
                pictureBox.ImageLocation = pictureListBox.SelectedItem.ToString();         
                btnMoveUp.Enabled = true;
                btnMoveDown.Enabled = true;
            }
        }

        private void btnMoveUp_Click(object sender, EventArgs e)
        {
            MoveItem(-1);
        }

        private void btnMoveDown_Click(object sender, EventArgs e)
        {
            MoveItem(1);
        }

        public void MoveItem(int direction)
        {
            // Checking selected item
            if (pictureListBox.SelectedItem == null || pictureListBox.SelectedIndex < 0)
                return; // No selected item - nothing to do

            // Checking bounds of the range
            if (pictureListBox.SelectedIndex + direction < 0 || 
                pictureListBox.SelectedIndex + direction+ pictureListBox.SelectedItems.Count > pictureListBox.Items.Count)
                return; // Index out of range - nothing to do

            int start=pictureListBox.SelectedIndex;
            int howMany=pictureListBox.SelectedItems.Count;
            string moveThisItem="";

            if (direction > 0)
            {
                // backup and delete item at position [start+howmany], then insert it at [start]
                moveThisItem = pictureListBox.Items[start+howMany].ToString(); ;
                pictureListBox.Items.RemoveAt(start+howMany);
                pictureListBox.Items.Insert(start, moveThisItem);
                // Restore selection in new position
                //pictureListBox.SetSelected(start, true);

            }
            else
            {
                // backup and delete item at position [start-1], then insert at [start+howmany-1]
                moveThisItem = pictureListBox.Items[start -1].ToString(); ;
                pictureListBox.Items.RemoveAt(start-1);
                pictureListBox.Items.Insert(start+howMany-1, moveThisItem);
                // Restore selection in new position
                //pictureListBox.SetSelected(start+howMany-1, true);

            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void optionButton_Click(object sender, EventArgs e)
        {
            optionsForm options = new optionsForm(this);

            options.Show();
        }


    }
}
