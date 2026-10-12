using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using stdole;

namespace PowerPointJumpPointsAddIn
{
    internal sealed class RibbonIconProvider
    {
        private readonly Assembly _assembly;

        public RibbonIconProvider()
        {
            _assembly = Assembly.GetExecutingAssembly();
        }

        public IPictureDisp GetIcon(string iconFileNameWithoutExtension)
        {
            try
            {
                string resourceName = FindResourceName(iconFileNameWithoutExtension);

                if (string.IsNullOrEmpty(resourceName))
                {
                    return null;
                }

                using (Stream stream = _assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        return null;
                    }

                    using (Bitmap bitmap = new Bitmap(stream))
                    {
                        return PictureConverter.ImageToPictureDisp(new Bitmap(bitmap));
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        private string FindResourceName(string iconFileNameWithoutExtension)
        {
            if (string.IsNullOrWhiteSpace(iconFileNameWithoutExtension))
            {
                return null;
            }

            string expectedSuffix = ".Icons." + iconFileNameWithoutExtension + ".png";

            return _assembly
                .GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith(expectedSuffix, StringComparison.OrdinalIgnoreCase));
        }

        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base(null) { }

            public static IPictureDisp ImageToPictureDisp(Image image)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
    }
}
