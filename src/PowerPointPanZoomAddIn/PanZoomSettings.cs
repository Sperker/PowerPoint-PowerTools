using System;
using System.IO;

namespace PowerPointPanZoomAddIn
{
    internal sealed class PanZoomSettings
    {
        public int[] ZoomLevels { get; set; }
        public int PanThresholdPixels { get; set; }
        public int PanStepDivisor { get; set; }
        public int TimerIntervalMilliseconds { get; set; }

        public PanZoomSettings()
        {
            ResetToDefaults();
        }

        public void ResetToDefaults()
        {
            ZoomLevels = (int[])ZoomConfiguration.DefaultZoomLevels.Clone();
            PanThresholdPixels = 8;
            PanStepDivisor = 8;
            TimerIntervalMilliseconds = 20;
        }

        public static PanZoomSettings Load()
        {
            PanZoomSettings settings = new PanZoomSettings();
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

                    if (key.Equals("ZoomLevels", StringComparison.OrdinalIgnoreCase))
                    {
                        int[] parsedZoomLevels;
                        string error;
                        if (ZoomConfiguration.TryParseZoomLevels(value, out parsedZoomLevels, out error))
                            settings.ZoomLevels = parsedZoomLevels;
                    }
                    else if (key.Equals("PanThresholdPixels", StringComparison.OrdinalIgnoreCase))
                    {
                        int parsed;
                        if (int.TryParse(value, out parsed) && parsed >= 1 && parsed <= 100)
                            settings.PanThresholdPixels = parsed;
                    }
                    else if (key.Equals("PanStepDivisor", StringComparison.OrdinalIgnoreCase))
                    {
                        int parsed;
                        if (int.TryParse(value, out parsed) && parsed >= 1 && parsed <= 100)
                            settings.PanStepDivisor = parsed;
                    }
                    else if (key.Equals("TimerIntervalMilliseconds", StringComparison.OrdinalIgnoreCase))
                    {
                        int parsed;
                        if (int.TryParse(value, out parsed) && parsed >= 10 && parsed <= 250)
                            settings.TimerIntervalMilliseconds = parsed;
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
                    "ZoomLevels=" + ZoomConfiguration.FormatZoomLevels(ZoomLevels),
                    "PanThresholdPixels=" + PanThresholdPixels,
                    "PanStepDivisor=" + PanStepDivisor,
                    "TimerIntervalMilliseconds=" + TimerIntervalMilliseconds
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
            return Path.Combine(appData, "PowerPointPanZoomAddIn");
        }

        public static string GetSettingsFilePath()
        {
            return Path.Combine(GetSettingsDirectory(), "settings.ini");
        }
    }
}
