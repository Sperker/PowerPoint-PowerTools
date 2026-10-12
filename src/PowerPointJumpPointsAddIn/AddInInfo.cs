namespace PowerPointJumpPointsAddIn
{
    internal static class AddInInfo
    {
        public const string Version = "1.9.0";
        public const string ProductName = "PowerPoint Jump Points Add-in";

        public static string RibbonGroupLabelXml
        {
            get
            {
                return "Jump Points v" + Version;
            }
        }

        public const string TagJumpPoint = "PPT_JUMPPOINT";
        public const string TagJumpPointValue = "TRUE";

        public const string TagJumpPointLabel = "PPT_JUMPPOINT_LABEL";
        public const string TagJumpPointCreated = "PPT_JUMPPOINT_CREATED";
        public const string TagJumpPointVersion = "PPT_JUMPPOINT_VERSION";
    }
}
