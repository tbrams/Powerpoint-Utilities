using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WindowsForms_Automate_Powerpoint
{
    public partial class optionsForm : Form
    {
        private FormAutomateImageAnimation mainForm = null;

        public optionsForm()
        {
            InitializeComponent();
        }

        public optionsForm(Form callingForm)
        {
            mainForm = callingForm as FormAutomateImageAnimation;
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            float myFirstDelay=0, myFirstDuration=0, myDelay=0, myDuration=0;
            if (float.TryParse(txtFirstDelay.Text, out myFirstDelay)) this.mainForm.myFirstDelay = myFirstDelay;
            if (float.TryParse(txtFirstDuration.Text, out myFirstDuration)) this.mainForm.myFirstDuration = myFirstDuration;
            if (float.TryParse(txtDelay.Text, out myDelay)) this.mainForm.myDelay = myDelay;
            if (float.TryParse(txtDuration.Text, out myDuration)) this.mainForm.myDuration = myDuration;

            this.mainForm.myAnimation = (PowerPoint.MsoAnimEffect)Enum.Parse(typeof(PowerPoint.MsoAnimEffect), comboBoxAnimation.Text);
            this.mainForm.myTrigger = (PowerPoint.MsoAnimTriggerType)Enum.Parse(typeof(PowerPoint.MsoAnimTriggerType), comboBoxStart.Text);
            this.Close();
        }

        private void optionsForm_Load(object sender, EventArgs e)
        {
            txtFirstDelay.Text = "0,0";
            txtFirstDuration.Text = "0,5";
            txtDelay.Text = "2,0";
            txtDuration.Text = "1,75";
            comboBoxAnimation.DataSource = System.Enum.GetValues(typeof(PowerPoint.MsoAnimEffect));
            comboBoxAnimation.Text = System.Enum.GetName(typeof(PowerPoint.MsoAnimEffect), PowerPoint.MsoAnimEffect.msoAnimEffectFade);
            comboBoxStart.DataSource = System.Enum.GetValues(typeof(PowerPoint.MsoAnimTriggerType));
            comboBoxStart.Text = System.Enum.GetName(typeof(PowerPoint.MsoAnimTriggerType), PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
        }

        private void comboBoxStart_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBoxAnimation_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
