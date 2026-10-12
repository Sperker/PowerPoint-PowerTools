using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace PPTMissingEdgeAlignAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public static Office.IRibbonUI RibbonUI { get; private set; }

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PPTMissingEdgeAlignAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            RibbonUI = ribbonUI;
        }

        public System.Drawing.Bitmap GetImage(Office.IRibbonControl control)
        {
            try
            {
                string imageName = null;
                switch (control.Id)
                {
                    case "btnEdgeAlign":
                        imageName = "auto_awesome_mosaic.png";
                        break;
                    case "btnSettings":
                        imageName = "settings.png";
                        break;
                    case "btnAbout":
                    case "btnInfo":
                        imageName = "info.png";
                        break;
                }

                if (imageName != null)
                {
                    Assembly asm = Assembly.GetExecutingAssembly();
                    string resourceName = "PPTMissingEdgeAlignAddIn.Icons." + imageName;
                    using (Stream stream = asm.GetManifestResourceStream(resourceName))
                    {
                        if (stream != null)
                        {
                            return new System.Drawing.Bitmap(stream);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error loading embedded icon: " + ex.Message);
            }

            // Fallback to geometric generation
            return GetFallbackBitmap(control.Id);
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            return AlignEngine.State == AlignState.Armed;
        }

        public void OnToggleAction(Office.IRibbonControl control, bool pressed)
        {
            if (pressed)
            {
                AlignEngine.ArmEngine();
            }
            else
            {
                AlignEngine.CancelEngine();
            }
        }

        public void OnAction(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnSettings":
                    using (var settingsForm = new SettingsForm())
                    {
                        settingsForm.ShowDialog(GetPowerPointWindow());
                    }
                    break;
                case "btnAbout":
                case "btnInfo":
                    string infoMessage = "PPTMissingEdgeAlignAddIn v1.2\n\n"
                        + "Instructions:\n"
                        + "1. Select the shape you want to move (source shape).\n"
                        + "2. Click 'EdgeAlign' (or use your Quick Access Toolbar shortcut). A beep sound will confirm the action is armed.\n"
                        + "3. Click the reference shape (target shape).\n\n"
                        + "The source shape will be mathematically aligned edge-to-edge to the target shape.\n\n"
                        + "Pro-Tips (Modifiers):\n"
                        + "- Hold the CTRL key when clicking the target shape to apply the gap defined in Settings.\n"
                        + "- Hold the SHIFT key when clicking the target shape to override the automatic alignment axis (vertical vs. horizontal) and force the opposite direction.";

                    System.Windows.Forms.MessageBox.Show(
                        GetPowerPointWindow(),
                        infoMessage,
                        "About PPTMissingEdgeAlignAddIn",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information
                    );
                    break;
            }
        }

        public static System.Windows.Forms.IWin32Window GetPowerPointWindow()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app != null && app.ActiveWindow != null)
                {
                    return new Win32WindowWrapper(new IntPtr(app.ActiveWindow.HWND));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error getting PowerPoint HWND: " + ex.Message);
            }
            return null;
        }

        private class Win32WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private readonly IntPtr _hwnd;
            public Win32WindowWrapper(IntPtr hwnd)
            {
                _hwnd = hwnd;
            }
            public IntPtr Handle
            {
                get { return _hwnd; }
            }
        }

        #endregion

        #region Helpers

        private System.Drawing.Bitmap GetFallbackBitmap(string id)
        {
            int size = (id == "btnEdgeAlign") ? 32 : 16;
            System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(size, size);
            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Transparent);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                if (id == "btnEdgeAlign")
                {
                    using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.DimGray))
                    {
                        g.FillRectangle(brush, 4, 4, 10, 24);
                    }
                    using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.DeepSkyBlue))
                    {
                        g.FillRectangle(brush, 18, 10, 10, 12);
                    }
                    using (var pen = new System.Drawing.Pen(System.Drawing.Color.Crimson, 2f))
                    {
                        g.DrawLine(pen, 15, 2, 15, 30);
                    }
                }
                else if (id == "btnSettings")
                {
                    using (var pen = new System.Drawing.Pen(System.Drawing.Color.DimGray, 3f))
                    {
                        g.DrawEllipse(pen, 3, 3, 10, 10);
                    }
                    using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.DimGray))
                    {
                        g.FillEllipse(brush, 6, 6, 4, 4);
                        g.FillRectangle(brush, 7, 1, 2, 4);
                        g.FillRectangle(brush, 7, 11, 2, 4);
                        g.FillRectangle(brush, 1, 7, 4, 2);
                        g.FillRectangle(brush, 11, 7, 4, 2);
                    }
                }
                else if (id == "btnAbout" || id == "btnInfo")
                {
                    using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.RoyalBlue))
                    {
                        g.FillEllipse(brush, 1, 1, 14, 14);
                    }
                    using (var brush = new System.Drawing.SolidBrush(System.Drawing.Color.White))
                    {
                        g.FillEllipse(brush, 7, 3, 2, 2);
                        g.FillRectangle(brush, 7, 6, 2, 6);
                        g.FillRectangle(brush, 6, 6, 2, 1);
                        g.FillRectangle(brush, 6, 11, 4, 1);
                    }
                }
            }
            return bmp;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] names = asm.GetManifestResourceNames();
            foreach (string name in names)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
