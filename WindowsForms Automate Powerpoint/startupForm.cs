using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace WindowsForms_Automate_Powerpoint
{
    public partial class startupForm : Form
    {
        public startupForm()
        {
            InitializeComponent();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
           // Form2 imageAnimationBuilder = new Form2();
           // imageAnimationBuilder.Show();
           // this.Close();
            System.Threading.Thread t = new System.Threading.Thread(new System.Threading.ThreadStart(ThreadProcForm2));
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            Application.Exit();
        }

        public static void ThreadProcForm2()
        {
            Application.Run(new FormAutomateImageAnimation());
        }

        private void buttonComments_Click(object sender, EventArgs e)
        {
            Thread t = new System.Threading.Thread(new System.Threading.ThreadStart(ThreadProcForm1));
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            Application.Exit();
        }
        public static void ThreadProcForm1()
        {
            Application.Run(new FormSecureComments());
        }


        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutForm about = new AboutForm();
            about.ShowDialog();
        }


    }
}
