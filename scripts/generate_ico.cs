using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

public class ValidIconGenerator
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    extern static bool DestroyIcon(IntPtr handle);

    public static void Generate()
    {
        string srcPath = @"E:\Code\QSBar\Release\Installinfo\logo_icon.png";
        string dstPath = @"E:\Code\QSBar\Installinfo\qsbar.ico";
        
        using (Bitmap bmpOriginal = new Bitmap(srcPath))
        {
            // Resize to 32x32 exactly
            using (Bitmap bmp32 = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {
                using (Graphics g = Graphics.FromImage(bmp32))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(bmpOriginal, 0, 0, 32, 32);
                }
                
                IntPtr hIcon = bmp32.GetHicon();
                using (Icon icon = Icon.FromHandle(hIcon))
                {
                    using (FileStream fs = new FileStream(dstPath, FileMode.Create))
                    {
                        icon.Save(fs);
                    }
                }
                DestroyIcon(hIcon);
            }
        }
        Console.WriteLine("32x32 Icon generated using Icon.Save().");
    }
}
