using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal static class RibbonIconProvider
    {
        public static object GetImage(string controlId)
        {
            string fileName = GetIconFileName(controlId);
            Image image = LoadEmbeddedIcon(fileName);

            if (image == null)
            {
                image = CreateFallbackIcon(controlId);
            }

            object picture = PictureConverter.ImageToPictureDisp(image);
            image.Dispose();
            return picture;
        }

        private static string GetIconFileName(string controlId)
        {
            if (string.Equals(controlId, "tglStayOnSlide", StringComparison.OrdinalIgnoreCase))
                return "frame_inspect.png";

            if (string.Equals(controlId, "btnSettings", StringComparison.OrdinalIgnoreCase))
                return "settings.png";

            if (string.Equals(controlId, "btnAbout", StringComparison.OrdinalIgnoreCase))
                return "info.png";

            return "info.png";
        }

        private static Image LoadEmbeddedIcon(string fileName)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string[] resourceNames = assembly.GetManifestResourceNames();

                foreach (string resourceName in resourceNames)
                {
                    if (resourceName.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase)
                        || resourceName.EndsWith(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                        {
                            if (stream == null)
                                return null;

                            using (Image original = Image.FromStream(stream))
                            {
                                return new Bitmap(original, new Size(32, 32));
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static Image CreateFallbackIcon(string controlId)
        {
            Bitmap bitmap = new Bitmap(32, 32);

            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (SolidBrush background = new SolidBrush(Color.FromArgb(26, 115, 232))) // Blue accent
            using (SolidBrush foreground = new SolidBrush(Color.White))
            using (Font font = new Font("Segoe UI", 16, FontStyle.Bold, GraphicsUnit.Pixel))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillEllipse(background, 5, 5, 22, 22);

                string letter = "S";

                if (string.Equals(controlId, "tglStayOnSlide", StringComparison.OrdinalIgnoreCase))
                    letter = "S";
                else if (string.Equals(controlId, "btnSettings", StringComparison.OrdinalIgnoreCase))
                    letter = "⚙";
                else if (string.Equals(controlId, "btnAbout", StringComparison.OrdinalIgnoreCase))
                    letter = "i";

                graphics.DrawString(letter, font, foreground, 9, 6);
            }

            return bitmap;
        }

        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base(null)
            {
            }

            public static object ImageToPictureDisp(Image image)
            {
                return GetIPictureDispFromPicture(image);
            }
        }
    }
}
