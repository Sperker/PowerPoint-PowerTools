using System;
using Microsoft.Office.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal sealed class CanvasWorkspaceManager
    {
        public const string AnchorNamePrefix = "__PPT_STAYONSLIDE_ANCHOR__";
        public const string AnchorTagName = "PPT_STAYONSLIDE_ANCHOR";
        public const string AnchorTagValue = "TRUE";

        public void EnsureAnchors(PowerPoint.Slide slide, int paddingPoints)
        {
            if (slide == null || paddingPoints <= 0)
                return;

            try
            {
                var pres = slide.Parent as PowerPoint.Presentation;
                if (pres == null) return;

                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;
                float centerX = Math.Max(1f, slideWidth / 2f);

                float topY = -paddingPoints;
                float bottomY = slideHeight + paddingPoints;

                PowerPoint.Shape topAnchor = FindAnchor(slide, true);
                if (topAnchor == null)
                {
                    topAnchor = CreateAnchorShape(slide, centerX, topY, true);
                }
                else
                {
                    if (Math.Abs(topAnchor.Top - topY) > 2f || Math.Abs(topAnchor.Left - centerX) > 2f)
                    {
                        topAnchor.Top = topY;
                        topAnchor.Left = centerX;
                    }
                }

                PowerPoint.Shape bottomAnchor = FindAnchor(slide, false);
                if (bottomAnchor == null)
                {
                    bottomAnchor = CreateAnchorShape(slide, centerX, bottomY, false);
                }
                else
                {
                    if (Math.Abs(bottomAnchor.Top - bottomY) > 2f || Math.Abs(bottomAnchor.Left - centerX) > 2f)
                    {
                        bottomAnchor.Top = bottomY;
                        bottomAnchor.Left = centerX;
                    }
                }
            }
            catch
            {
            }
        }

        private PowerPoint.Shape CreateAnchorShape(PowerPoint.Slide slide, float x, float y, bool isTop)
        {
            PowerPoint.Shape anchor = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                x,
                y,
                2f,
                2f
            );

            string suffix = isTop ? "_TOP" : "_BOTTOM";
            anchor.Name = AnchorNamePrefix + slide.SlideID + suffix;
            anchor.Tags.Add(AnchorTagName, AnchorTagValue);

            anchor.Fill.Visible = MsoTriState.msoTrue;
            anchor.Fill.Transparency = 1.0f;
            anchor.Line.Visible = MsoTriState.msoFalse;

            try { anchor.ZOrder(Office.MsoZOrderCmd.msoSendToBack); } catch { }

            return anchor;
        }

        private PowerPoint.Shape FindAnchor(PowerPoint.Slide slide, bool isTop)
        {
            string suffix = isTop ? "_TOP" : "_BOTTOM";
            try
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    bool nameMatches = false;
                    bool tagMatches = false;

                    try
                    {
                        nameMatches = shape.Name != null
                            && shape.Name.StartsWith(AnchorNamePrefix, StringComparison.OrdinalIgnoreCase)
                            && shape.Name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase);
                    }
                    catch { }

                    try
                    {
                        tagMatches = string.Equals(shape.Tags[AnchorTagName], AnchorTagValue, StringComparison.OrdinalIgnoreCase)
                            && (shape.Name != null && shape.Name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
                    }
                    catch { }

                    if (nameMatches || tagMatches)
                        return shape;
                }
            }
            catch { }

            return null;
        }

        public int RemoveAnchorsFromSlide(PowerPoint.Slide slide)
        {
            if (slide == null) return 0;
            int deleted = 0;

            try
            {
                for (int i = slide.Shapes.Count; i >= 1; i--)
                {
                    PowerPoint.Shape shape = slide.Shapes[i];
                    bool nameMatches = false;
                    bool tagMatches = false;

                    try
                    {
                        nameMatches = shape.Name != null && shape.Name.StartsWith(AnchorNamePrefix, StringComparison.OrdinalIgnoreCase);
                    }
                    catch { }

                    try
                    {
                        tagMatches = string.Equals(shape.Tags[AnchorTagName], AnchorTagValue, StringComparison.OrdinalIgnoreCase);
                    }
                    catch { }

                    if (nameMatches || tagMatches)
                    {
                        shape.Delete();
                        deleted++;
                    }
                }
            }
            catch { }

            return deleted;
        }

        public int RemoveAllAnchors(PowerPoint.Presentation pres)
        {
            if (pres == null) return 0;
            int total = 0;

            try
            {
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    total += RemoveAnchorsFromSlide(slide);
                }
            }
            catch { }

            return total;
        }
    }
}
