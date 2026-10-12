using System;
using System.IO;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal sealed class StayOnSlideSettings
    {
        public bool StayOnSlideEnabled { get; set; }
        public bool EnableVirtualCanvas { get; set; }
        public int CanvasPaddingPoints { get; set; }
        public bool BlockKeyboardSlideNavigation { get; set; }
        public int WheelScrollStepLines { get; set; }

        public StayOnSlideSettings()
        {
            ResetToDefaults();
        }

        public void ResetToDefaults()
        {
            StayOnSlideEnabled = true;
            EnableVirtualCanvas = true;
            CanvasPaddingPoints = 800;
            BlockKeyboardSlideNavigation = true;
            WheelScrollStepLines = 3;
        }

        public static StayOnSlideSettings Load()
        {
            StayOnSlideSettings settings = new StayOnSlideSettings();
            string path = GetSettingsFilePath();

            try
            {
                if (!File.Exists(path))
                    return settings;

                string[] lines = File.ReadAllLines(path);

                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    int index = line.IndexOf('=');
                    if (index <= 0)
                        continue;

                    string key = line.Substring(0, index).Trim();
                    string value = line.Substring(index + 1).Trim();

                    if (key.Equals("StayOnSlideEnabled", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                            settings.StayOnSlideEnabled = parsed;
                    }
                    else if (key.Equals("EnableVirtualCanvas", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                            settings.EnableVirtualCanvas = parsed;
                    }
                    else if (key.Equals("CanvasPaddingPoints", StringComparison.OrdinalIgnoreCase))
                    {
                        int parsed;
                        if (int.TryParse(value, out parsed) && parsed >= 100 && parsed <= 10000)
                            settings.CanvasPaddingPoints = parsed;
                    }
                    else if (key.Equals("BlockKeyboardSlideNavigation", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                            settings.BlockKeyboardSlideNavigation = parsed;
                    }
                    else if (key.Equals("WheelScrollStepLines", StringComparison.OrdinalIgnoreCase))
                    {
                        int parsed;
                        if (int.TryParse(value, out parsed) && parsed >= 1 && parsed <= 50)
                            settings.WheelScrollStepLines = parsed;
                    }
                }
            }
            catch
            {
                settings.ResetToDefaults();
            }

            return settings;
        }

        public void Save()
        {
            try
            {
                string directory = GetSettingsDirectory();
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                string[] lines = new string[]
                {
                    "StayOnSlideEnabled=" + StayOnSlideEnabled,
                    "EnableVirtualCanvas=" + EnableVirtualCanvas,
                    "CanvasPaddingPoints=" + CanvasPaddingPoints,
                    "BlockKeyboardSlideNavigation=" + BlockKeyboardSlideNavigation,
                    "WheelScrollStepLines=" + WheelScrollStepLines
                };

                File.WriteAllLines(GetSettingsFilePath(), lines);
            }
            catch
            {
            }
        }

        public static string GetSettingsDirectory()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "PowerPointStayOnSlidePointsAddIn");
        }

        public static string GetSettingsFilePath()
        {
            return Path.Combine(GetSettingsDirectory(), "settings.ini");
        }
    }
}
