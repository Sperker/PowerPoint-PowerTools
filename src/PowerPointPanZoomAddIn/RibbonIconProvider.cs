using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace PowerPointPanZoomAddIn
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
            if (string.Equals(controlId, "tglPan", StringComparison.OrdinalIgnoreCase))
                return "pan_tool.png";

            if (string.Equals(controlId, "tglZoomCursor", StringComparison.OrdinalIgnoreCase))
                return "ads_click.png";

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
            using (SolidBrush background = new SolidBrush(Color.FromArgb(64, 126, 201)))
            using (SolidBrush foreground = new SolidBrush(Color.White))
            using (Font font = new Font("Segoe UI", 16, FontStyle.Bold, GraphicsUnit.Pixel))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillEllipse(background, 5, 5, 22, 22);

                string letter = "i";

                if (string.Equals(controlId, "tglPan", StringComparison.OrdinalIgnoreCase))
                    letter = "P";
                else if (string.Equals(controlId, "tglZoomCursor", StringComparison.OrdinalIgnoreCase))
                    letter = "Z";
                else if (string.Equals(controlId, "btnSettings", StringComparison.OrdinalIgnoreCase))
                    letter = "S";

                graphics.DrawString(letter, font, foreground, 10, 7);
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
