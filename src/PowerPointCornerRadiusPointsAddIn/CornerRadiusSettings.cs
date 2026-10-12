using System;
using System.Globalization;
using System.IO;

namespace PowerPointCornerRadiusPointsAddIn
{
    internal sealed class CornerRadiusSettings
    {
        public bool PlaySoundOnCopy { get; set; }
        public bool ShowMessageOnCopy { get; set; }
        public float? ManualRadiusPoints { get; set; }

        public CornerRadiusSettings()
        {
            ResetToDefaults();
        }

        public void ResetToDefaults()
        {
            PlaySoundOnCopy = true;
            ShowMessageOnCopy = false;
            ManualRadiusPoints = null; // No default radius
        }

        public static CornerRadiusSettings Load()
        {
            CornerRadiusSettings settings = new CornerRadiusSettings();
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

                    if (key.Equals("PlaySoundOnCopy", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                            settings.PlaySoundOnCopy = parsed;
                    }
                    else if (key.Equals("ShowMessageOnCopy", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                            settings.ShowMessageOnCopy = parsed;
                    }
                    else if (key.Equals("ManualRadiusPoints", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrWhiteSpace(value) || value.Equals("null", StringComparison.OrdinalIgnoreCase))
                        {
                            settings.ManualRadiusPoints = null;
                        }
                        else
                        {
                            float parsed;
                            if (float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) && parsed >= 0)
                            {
                                settings.ManualRadiusPoints = parsed > 0 ? (float?)parsed : null;
                            }
                        }
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

                string manualRadStr = ManualRadiusPoints.HasValue
                    ? ManualRadiusPoints.Value.ToString("0.##", CultureInfo.InvariantCulture)
                    : "";

                string[] lines = new string[]
                {
                    "PlaySoundOnCopy=" + PlaySoundOnCopy,
                    "ShowMessageOnCopy=" + ShowMessageOnCopy,
                    "ManualRadiusPoints=" + manualRadStr
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
            return Path.Combine(appData, "PowerPointCornerRadiusPointsAddIn");
        }

        public static string GetSettingsFilePath()
        {
            return Path.Combine(GetSettingsDirectory(), "settings.ini");
        }
    }
}
