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
    public partial class Form2 : Form
    {
        string fileName;

        public Form2()
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
            pictureListBox.Visible = false;
            btnMoveDown.Visible = false;
            btnMoveUp.Visible = false;
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
                this.Size = new Size(900, 500);

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
            }
            else
                Debug.WriteLine("Something went wrong reading the filename from the Dialog");

            quitButton.Visible = true;
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
               //oPowerPoint.Visible = Office.MsoTriState.msoFalse; 

                // Create a new Presentation. 
                PowerPoint.Presentation oPre = oPowerPoint.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                Debug.WriteLine("A new presentation is created");

                // Insert a new Slide and add some text to it. 
                Debug.WriteLine("Text Slide inserted");
                PowerPoint.Slide oSlide = oPre.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutText);
                oSlide.Shapes[1].TextFrame.TextRange.Text = "Fist slide inserted";

                // Insert a new blank Slide for the picture. 
                oSlide = oPre.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
                Console.WriteLine("Blank slide for images inserted");

                // Animation settings
                // myTrigger is used between pictures normally, myAlternativeTrigger can be used to click through the images
                PowerPoint.MsoAnimTriggerType myTrigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;
          //      PowerPoint.MsoAnimTriggerType myAlternativeTrigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;

                // Animations are kept simple, normal effeft is fading, alternative is appearing
                PowerPoint.MsoAnimEffect myAnimation = PowerPoint.MsoAnimEffect.msoAnimEffectFade;
          //      PowerPoint.MsoAnimEffect myAlternativeAnimation = PowerPoint.MsoAnimEffect.msoAnimEffectAppear;

                // Timing for animation - CHANGE LATER !!!
                float myFirstDelay=0;
                float myFirstDuration = 0.5F;
                float myDuration = 1.75F;
                float myDelay = 2.0F;

                Boolean FirstImage = true;  // Used to handle special effect for very first picture
                foreach (string p in pictureListBox.Items)
                {                    
                    string imageFileName = p;

                    Bitmap myBitmap = new Bitmap(@imageFileName);
                    Graphics g = Graphics.FromImage(myBitmap);

                    float dx, dy;
                    try
                    {
                        dx = g.DpiX;
                        dy = g.DpiY;
                    }
                    finally
                    {
                        g.Dispose();
                    }
                    Debug.WriteLine("DPI Info for " + imageFileName);
                    Debug.WriteLine(" dpix:" + dx);
                    Debug.WriteLine(" dpiy:" + dy);


                    PowerPoint.Shape myPic;
                    myPic = oSlide.Shapes.AddPicture(imageFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 1, 1, -1, -1);

                    // find png file dimensions
                    var buff = new byte[32];
                    using (var d = File.OpenRead(imageFileName))
                    {
                        d.Read(buff, 0, 32);
                    }
                    const int wOff = 16;
                    const int hOff = 20;
                    double imageWidth = (double)BitConverter.ToInt32(new[] { buff[wOff + 3], buff[wOff + 2], buff[wOff + 1], buff[wOff + 0], }, 0);
                    double imageHeight = (double)BitConverter.ToInt32(new[] { buff[hOff + 3], buff[hOff + 2], buff[hOff + 1], buff[hOff + 0], }, 0);

                    // Scale to fit slide area
                    double xfactor = 1.0;

                    Rectangle bounds = Screen.GetBounds(Point.Empty);
                    Debug.WriteLine("Device Resolution:");
                    Debug.WriteLine("Y scaling: " +bounds.Height);
                    Debug.WriteLine("X scaling: " + bounds.Width);
                 
                    const float PixelPerPoint=(float)4/3;
                    double slideWidth = oPowerPoint.ActivePresentation.PageSetup.SlideWidth * PixelPerPoint;
                    double slideHeight = oPowerPoint.ActivePresentation.PageSetup.SlideHeight * PixelPerPoint;

                    //double aspectRatioScreen = (double)bounds.Width / (double)bounds.Height;
                    //double aspectRatioSlide = slideWidth / slideHeight;

                    //if (aspectRatioSlide>=aspectRatioScreen) {
                    //   ppp = bounds.Width/slideWidth;
                    //} else {
                    //   ppp = bounds.Height/slideHeight;
                    //}

                    double ppp = dx/96;

                    if (imageWidth> imageHeight) {
                          // X Scaling
                          xfactor=(slideWidth/imageWidth)*ppp;
                       // xfactor=(float)960/575;
                          myPic.ScaleWidth((float)xfactor, Office.MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    }
                    else
                    {
                          // Y Scaling
                          // Something is wrong... if there is a DPI statement in the file things start going wrong.
                          // I can get a perfect fit for 120 DPI and this ppp below, but otherwise it does not work
                          // unfortunately
                          xfactor = (slideHeight / imageHeight)*ppp;
                          // if DPI == 96 no need for ppp
                          // if DPI == 120 - need 1.25 more size, confirm with Picture1.png
                        
//                          xfactor = slideHeight / imageHeight;
                          myPic.ScaleHeight((float)xfactor, Office.MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    }


                    Debug.WriteLine("myPic w x h:" + imageWidth + ", " + imageHeight);
                    Debug.WriteLine("Page  w x h: " + slideWidth + ", " + slideHeight);
                    Debug.WriteLine("Xfactor : " + xfactor);
                   
                    myPic.Left = oPowerPoint.ActivePresentation.PageSetup.SlideWidth / 2 - myPic.Width / 2;
                    myPic.Top = oPowerPoint.ActivePresentation.PageSetup.SlideHeight / 2 - myPic.Height / 2;

                    // Add animation to the screenshot
                    // Remove the trigger for first image, so that it will appear automatically
                    PowerPoint.MsoAnimTriggerType thisTrigger; 
                    thisTrigger=myTrigger;
                    if (FirstImage) {
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
                Console.WriteLine("Save and close the presentation");

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
                Console.WriteLine("Quit the PowerPoint application");
                oPowerPoint.Quit();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Solution1.AutomatePowerPoint throws the error: {0}",
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
            btnMoveUp.Enabled = true;
            btnMoveDown.Enabled = true;
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

            // Calculate new index using move direction
            int newIndex = pictureListBox.SelectedIndex + direction;

            // Checking bounds of the range
            if (newIndex < 0 || newIndex >= pictureListBox.Items.Count)
                return; // Index out of range - nothing to do

            object selected = pictureListBox.SelectedItem;

            // Removing removable element
            pictureListBox.Items.Remove(selected);
            // Insert it in new position
            pictureListBox.Items.Insert(newIndex, selected);
            // Restore selection
            pictureListBox.SetSelected(newIndex, true);
        }


    }
}
