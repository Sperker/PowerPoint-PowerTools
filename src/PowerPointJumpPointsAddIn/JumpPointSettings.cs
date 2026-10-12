using System;
using Microsoft.Win32;

namespace PowerPointJumpPointsAddIn
{
    public sealed class JumpPointSettings
    {
        private const string RegistryPath = @"Software\PowerPointJumpPointsAddIn";

        public bool WrapAroundEnabled { get; set; } = true;
        public bool IncludeHiddenSlides { get; set; } = true;
        public bool ShowMessages { get; set; } = true;
        public bool ShowToggleResultMessages { get; set; } = false;
        public bool PlayToggleResultSound { get; set; } = true;
        public bool ShowNavigatorConfirmations { get; set; } = true;

        public void Load()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    if (key == null)
                    {
                        return;
                    }

                    ShowToggleResultMessages = ReadBool(key, "ShowToggleResultMessages", ShowToggleResultMessages);
                    PlayToggleResultSound = ReadBool(key, "PlayToggleResultSound", PlayToggleResultSound);
                    ShowNavigatorConfirmations = ReadBool(key, "ShowNavigatorConfirmations", ShowNavigatorConfirmations);
                    WrapAroundEnabled = ReadBool(key, "WrapAroundEnabled", WrapAroundEnabled);
                    IncludeHiddenSlides = ReadBool(key, "IncludeHiddenSlides", IncludeHiddenSlides);
                    ShowMessages = ReadBool(key, "ShowMessages", ShowMessages);
                }
            }
            catch
            {
                // If settings cannot be loaded, keep defaults.
            }
        }

        public void Save()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
                {
                    if (key == null)
                    {
                        return;
                    }

                    key.SetValue("ShowToggleResultMessages", ShowToggleResultMessages ? 1 : 0);
                    key.SetValue("PlayToggleResultSound", PlayToggleResultSound ? 1 : 0);
                    key.SetValue("ShowNavigatorConfirmations", ShowNavigatorConfirmations ? 1 : 0);
                    key.SetValue("WrapAroundEnabled", WrapAroundEnabled ? 1 : 0);
                    key.SetValue("IncludeHiddenSlides", IncludeHiddenSlides ? 1 : 0);
                    key.SetValue("ShowMessages", ShowMessages ? 1 : 0);
                }
            }
            catch
            {
                // Ignore save errors.
            }
        }

        private bool ReadBool(RegistryKey key, string name, bool defaultValue)
        {
            try
            {
                object value = key.GetValue(name);

                if (value == null)
                {
                    return defaultValue;
                }

                return Convert.ToInt32(value) != 0;
            }
            catch
            {
                return defaultValue;
            }
        }
    }
}
