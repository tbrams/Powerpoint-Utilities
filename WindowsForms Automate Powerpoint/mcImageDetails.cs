using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace WindowsForms_Automate_Powerpoint
{
    class mcImageDetails
    {
        public float dx, dy;
        public double imageWidth, imageHeight;
        public double screenWidth, screenHeight;

        public mcImageDetails(){}
        public mcImageDetails(string imageFileName){
            Bitmap myBitmap = new Bitmap(@imageFileName);
            Graphics g = Graphics.FromImage(myBitmap);

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

            // Find png dimensions
            var buff = new byte[32];
            using (var d = File.OpenRead(imageFileName))
            {
                d.Read(buff, 0, 32);
            }
            const int wOff = 16;
            const int hOff = 20;
            imageWidth = (double)BitConverter.ToInt32(new[] { buff[wOff + 3], buff[wOff + 2], buff[wOff + 1], buff[wOff + 0], }, 0);
            imageHeight = (double)BitConverter.ToInt32(new[] { buff[hOff + 3], buff[hOff + 2], buff[hOff + 1], buff[hOff + 0], }, 0);

            Rectangle bounds = Screen.GetBounds(Point.Empty);
            screenWidth = bounds.Width;
            screenHeight = bounds.Height;

            Debug.WriteLine("Device Resolution:");
            Debug.WriteLine("Width: " + screenWidth);
            Debug.WriteLine("Height: " + screenHeight);

        }


    }
}
