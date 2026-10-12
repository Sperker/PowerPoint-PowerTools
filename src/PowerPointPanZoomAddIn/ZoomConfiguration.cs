using System;
using System.Collections.Generic;
using System.Linq;

namespace PowerPointPanZoomAddIn
{
    internal static class ZoomConfiguration
    {
        public static readonly int[] DefaultZoomLevels = new int[]
        {
            10,
            25,
            33,
            50,
            66,
            75,
            100,
            125,
            150,
            200,
            250,
            300,
            350,
            400
        };

        public const int MinimumZoom = 10;
        public const int MaximumZoom = 400;

        public static string GetZoomLevelsText()
        {
            return FormatZoomLevels(DefaultZoomLevels);
        }

        public static string FormatZoomLevels(IEnumerable<int> zoomLevels)
        {
            if (zoomLevels == null)
                return string.Empty;

            return string.Join(", ", zoomLevels);
        }

        public static bool TryParseZoomLevels(string text, out int[] zoomLevels, out string errorMessage)
        {
            zoomLevels = null;
            errorMessage = null;

            if (string.IsNullOrWhiteSpace(text))
            {
                errorMessage = "Please enter at least two zoom levels.";
                return false;
            }

            string normalized = text
                .Replace(";", ",")
                .Replace(" ", ",")
                .Replace("\r", ",")
                .Replace("\n", ",")
                .Replace("\t", ",");

            string[] parts = normalized.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            List<int> values = new List<int>();

            foreach (string rawPart in parts)
            {
                string part = rawPart.Trim().Replace("%", "");
                int value;

                if (!int.TryParse(part, out value))
                {
                    errorMessage = "Invalid zoom level: " + rawPart;
                    return false;
                }

                if (value < MinimumZoom || value > MaximumZoom)
                {
                    errorMessage = "Zoom level " + value + " is outside the allowed range " + MinimumZoom + "-" + MaximumZoom + ".";
                    return false;
                }

                if (!values.Contains(value))
                    values.Add(value);
            }

            if (values.Count < 2)
            {
                errorMessage = "Please enter at least two different zoom levels.";
                return false;
            }

            values.Sort();
            zoomLevels = values.ToArray();
            return true;
        }

        public static bool ContainsRecommendedBaseZoom(int[] zoomLevels)
        {
            if (zoomLevels == null)
                return false;

            return zoomLevels.Contains(100);
        }
    }
}
