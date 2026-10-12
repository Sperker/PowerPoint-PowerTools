using System;
using System.Collections.Generic;
using System.Linq;
using System.Media;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointJumpPointsAddIn
{
    public sealed class JumpPointController
    {
        private PowerPoint.Application _application;
        private readonly JumpPointSettings _settings;
        private int _lastKnownSlideId = -1;
        private JumpPointListForm _jumpPointListForm;

        public JumpPointController(PowerPoint.Application application, JumpPointSettings settings)
        {
            _application = application;
            _settings = settings ?? new JumpPointSettings();
        }

        public void SetApplication(PowerPoint.Application application)
        {
            if (application != null)
            {
                _application = application;
            }
        }

        public bool IsWrapAroundEnabled() { return _settings.WrapAroundEnabled; }
        public bool AreHiddenSlidesIncluded() { return _settings.IncludeHiddenSlides; }
        public bool IsToggleResultMessageEnabled() { return _settings.ShowToggleResultMessages; }
        public bool IsToggleResultSoundEnabled() { return _settings.PlayToggleResultSound; }
        public bool IsNavigatorConfirmationEnabled() { return _settings.ShowNavigatorConfirmations; }

        public int GetJumpPointCount()
        {
            try { return GetJumpPointSlides().Count; } catch { return 0; }
        }

        public void NotifyCurrentSlideChanged()
        {
            try
            {
                if (_jumpPointListForm == null || _jumpPointListForm.IsDisposed)
                {
                    return;
                }

                PowerPoint.Slide currentSlide = GetCurrentSlide();

                if (currentSlide == null)
                {
                    _jumpPointListForm.ClearCurrentSlideSelection();
                    return;
                }

                _jumpPointListForm.HighlightSlideById(currentSlide.SlideID);
            }
            catch
            {
                // Ignore live highlight issues.
            }
        }

        public void ShowSettingsDialog()
        {
            try
            {
                using (JumpPointSettingsForm form = new JumpPointSettingsForm(_settings))
                {
                    DialogResult result = form.ShowDialog();
                    if (result != DialogResult.OK)
                    {
                        return;
                    }

                    _settings.ShowToggleResultMessages = form.ShowToggleMessages;
                    _settings.PlayToggleResultSound = form.PlayToggleSound;
                    _settings.ShowNavigatorConfirmations = form.ShowNavigatorConfirmations;
                    _settings.Save();
                    ShowMessage("Jump Points settings saved.");
                }
            }
            catch
            {
                ShowMessage("Could not open Jump Points settings.");
            }
        }

        public void ShowJumpPointListDialog()
        {
            try
            {
                if (_jumpPointListForm == null || _jumpPointListForm.IsDisposed)
                {
                    _jumpPointListForm = new JumpPointListForm(this);
                    _jumpPointListForm.FormClosed += JumpPointListForm_FormClosed;
                    _jumpPointListForm.Show();
                    NotifyCurrentSlideChanged();
                    return;
                }

                _jumpPointListForm.RefreshJumpPointList();

                if (_jumpPointListForm.WindowState == FormWindowState.Minimized)
                {
                    _jumpPointListForm.WindowState = FormWindowState.Normal;
                }

                _jumpPointListForm.Show();
                _jumpPointListForm.BringToFront();
                _jumpPointListForm.Activate();
                NotifyCurrentSlideChanged();
            }
            catch
            {
                ShowMessage("Could not open Jump Points list.");
            }
        }

        private void JumpPointListForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if (_jumpPointListForm != null)
                {
                    _jumpPointListForm.FormClosed -= JumpPointListForm_FormClosed;
                }
            }
            catch { }

            _jumpPointListForm = null;
        }

        public List<JumpPointListEntry> GetJumpPointListEntries()
        {
            List<JumpPointListEntry> result = new List<JumpPointListEntry>();
            IReadOnlyList<PowerPoint.Slide> jumpPointSlides = GetJumpPointSlides();

            foreach (PowerPoint.Slide slide in jumpPointSlides)
            {
                try
                {
                    result.Add(new JumpPointListEntry
                    {
                        SlideId = slide.SlideID,
                        SlideIndex = slide.SlideIndex,
                        Label = GetTagValue(slide, AddInInfo.TagJumpPointLabel),
                        Title = GetSlideTitleOrFirstText(slide),
                        Created = GetTagValue(slide, AddInInfo.TagJumpPointCreated),
                        Version = GetTagValue(slide, AddInInfo.TagJumpPointVersion)
                    });
                }
                catch { }
            }

            return result.OrderBy(e => e.SlideIndex).ToList();
        }

        public bool RemoveJumpPointBySlideId(int slideId)
        {
            try
            {
                PowerPoint.Slide slide = FindSlideById(slideId);
                if (slide == null || !HasJumpPoint(slide))
                {
                    return false;
                }

                RemoveJumpPoint(slide);
                RefreshOpenJumpPointListForm();
                NotifyCurrentSlideChanged();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool SetJumpPointLabelBySlideId(int slideId, string label)
        {
            try
            {
                PowerPoint.Slide slide = FindSlideById(slideId);

                if (slide == null || !HasJumpPoint(slide))
                {
                    return false;
                }

                RemoveTagIfExists(slide, AddInInfo.TagJumpPointLabel);

                if (!string.IsNullOrWhiteSpace(label))
                {
                    slide.Tags.Add(AddInInfo.TagJumpPointLabel, label.Trim());
                }

                RefreshOpenJumpPointListForm();
                NotifyCurrentSlideChanged();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public int ClearAllJumpPointsFromNavigator()
        {
            try
            {
                IReadOnlyList<PowerPoint.Slide> jumpPoints = GetJumpPointSlides();
                int removedCount = 0;

                foreach (PowerPoint.Slide slide in jumpPoints)
                {
                    try
                    {
                        RemoveJumpPoint(slide);
                        removedCount++;
                    }
                    catch
                    {
                        // Continue with remaining slides.
                    }
                }

                RefreshOpenJumpPointListForm();
                NotifyCurrentSlideChanged();
                return removedCount;
            }
            catch
            {
                return 0;
            }
        }

        public bool NavigateToSlideById(int slideId)
        {
            try
            {
                PowerPoint.Slide slide = FindSlideById(slideId);
                if (slide == null)
                {
                    return false;
                }

                NavigateToSlide(slide);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void ToggleJumpPointOnCurrentSlide()
        {
            PowerPoint.Slide slide = GetCurrentSlide();
            if (slide == null)
            {
                ShowMessage("No active slide found.");
                return;
            }

            RememberSlide(slide);

            if (HasJumpPoint(slide))
            {
                RemoveJumpPoint(slide);
                NotifyToggleResult($"Jump point removed from slide {slide.SlideIndex}.");
            }
            else
            {
                SetJumpPoint(slide);
                NotifyToggleResult($"Jump point added to slide {slide.SlideIndex}.");
            }

            RefreshOpenJumpPointListForm();
            NotifyCurrentSlideChanged();
        }

        public bool HasJumpPointOnCurrentSlide()
        {
            PowerPoint.Slide slide = GetCurrentSlide();
            return slide != null && HasJumpPoint(slide);
        }

        public bool HasJumpPoint(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return false;
            }

            string value = GetTagValue(slide, AddInInfo.TagJumpPoint);
            return string.Equals(value, AddInInfo.TagJumpPointValue, StringComparison.OrdinalIgnoreCase);
        }

        public void SetJumpPoint(PowerPoint.Slide slide)
        {
            if (slide == null) return;

            RemoveTagIfExists(slide, AddInInfo.TagJumpPoint);
            RemoveTagIfExists(slide, AddInInfo.TagJumpPointCreated);
            RemoveTagIfExists(slide, AddInInfo.TagJumpPointVersion);

            slide.Tags.Add(AddInInfo.TagJumpPoint, AddInInfo.TagJumpPointValue);
            slide.Tags.Add(AddInInfo.TagJumpPointCreated, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            slide.Tags.Add(AddInInfo.TagJumpPointVersion, AddInInfo.Version);
        }

        public void RemoveJumpPoint(PowerPoint.Slide slide)
        {
            if (slide == null) return;

            RemoveTagIfExists(slide, AddInInfo.TagJumpPoint);
            RemoveTagIfExists(slide, AddInInfo.TagJumpPointLabel);
            RemoveTagIfExists(slide, AddInInfo.TagJumpPointCreated);
            RemoveTagIfExists(slide, AddInInfo.TagJumpPointVersion);
        }

        public void ClearAllJumpPoints()
        {
            PowerPoint.Presentation presentation = GetActivePresentation();
            if (presentation == null)
            {
                ShowMessage("No active presentation found.");
                return;
            }

            IReadOnlyList<PowerPoint.Slide> jumpPoints = GetJumpPointSlides();
            if (jumpPoints.Count == 0)
            {
                ShowMessage("No jump points defined in this presentation.");
                return;
            }

            DialogResult result = MessageBox.Show(
                $"Remove all {jumpPoints.Count} jump point(s) from this presentation?",
                AddInInfo.ProductName,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
            {
                return;
            }

            int removedCount = 0;
            foreach (PowerPoint.Slide slide in jumpPoints)
            {
                try
                {
                    RemoveJumpPoint(slide);
                    removedCount++;
                }
                catch { }
            }

            RefreshOpenJumpPointListForm();
            NotifyCurrentSlideChanged();
            ShowMessage($"Removed {removedCount} jump point(s).");
        }

        public IReadOnlyList<PowerPoint.Slide> GetJumpPointSlides()
        {
            PowerPoint.Presentation presentation = GetActivePresentation();
            if (presentation == null)
            {
                return new List<PowerPoint.Slide>();
            }

            List<PowerPoint.Slide> result = new List<PowerPoint.Slide>();
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (!_settings.IncludeHiddenSlides && IsHidden(slide))
                {
                    continue;
                }

                if (HasJumpPoint(slide))
                {
                    result.Add(slide);
                }
            }

            return result.OrderBy(s => s.SlideIndex).ToList();
        }

        public void GoToNextJumpPoint() { GoToJumpPoint(Direction.Next); }
        public void GoToPreviousJumpPoint() { GoToJumpPoint(Direction.Previous); }

        private void GoToJumpPoint(Direction direction)
        {
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            if (currentSlide == null)
            {
                ShowMessage("No active slide found.");
                return;
            }

            RememberSlide(currentSlide);
            IReadOnlyList<PowerPoint.Slide> jumpPoints = GetJumpPointSlides();

            if (jumpPoints.Count == 0)
            {
                ShowMessage("No jump points defined in this presentation. Use Toggle Jump Point to add one.");
                return;
            }

            int currentIndex = currentSlide.SlideIndex;
            PowerPoint.Slide targetSlide = null;

            if (direction == Direction.Next)
            {
                targetSlide = jumpPoints.Where(s => s.SlideIndex > currentIndex).OrderBy(s => s.SlideIndex).FirstOrDefault();
                if (targetSlide == null && _settings.WrapAroundEnabled)
                {
                    targetSlide = jumpPoints.OrderBy(s => s.SlideIndex).FirstOrDefault();
                }
            }
            else
            {
                targetSlide = jumpPoints.Where(s => s.SlideIndex < currentIndex).OrderByDescending(s => s.SlideIndex).FirstOrDefault();
                if (targetSlide == null && _settings.WrapAroundEnabled)
                {
                    targetSlide = jumpPoints.OrderByDescending(s => s.SlideIndex).FirstOrDefault();
                }
            }

            if (targetSlide == null)
            {
                ShowMessage(direction == Direction.Next ? "No next jump point found." : "No previous jump point found.");
                return;
            }

            NavigateToSlide(targetSlide);
        }

        public PowerPoint.Slide GetCurrentSlide()
        {
            PowerPoint.Slide liveSlide = TryGetCurrentSlideLive();
            if (liveSlide != null)
            {
                RememberSlide(liveSlide);
                return liveSlide;
            }

            PowerPoint.Slide rememberedSlide = GetRememberedSlide();
            if (rememberedSlide != null)
            {
                return rememberedSlide;
            }

            return null;
        }

        private PowerPoint.Slide TryGetCurrentSlideLive()
        {
            try
            {
                if (_application == null || _application.Presentations.Count == 0)
                {
                    return null;
                }

                try
                {
                    if (_application.SlideShowWindows.Count > 0)
                    {
                        PowerPoint.SlideShowWindow slideShowWindow = _application.SlideShowWindows[1];
                        if (slideShowWindow != null && slideShowWindow.View != null)
                        {
                            PowerPoint.Slide slideShowSlide = TryResolveSlideFromObject(slideShowWindow.View.Slide);
                            if (slideShowSlide != null) return slideShowSlide;
                        }
                    }
                }
                catch { }

                PowerPoint.DocumentWindow window = null;
                try { window = _application.ActiveWindow; } catch { }
                if (window == null) return null;

                try
                {
                    PowerPoint.Slide slideFromView = TryResolveSlideFromObject(window.View.Slide);
                    if (slideFromView != null) return slideFromView;
                }
                catch { }

                try
                {
                    PowerPoint.Slide slideFromSelection = TryGetSlideFromSelection(window.Selection);
                    if (slideFromSelection != null) return slideFromSelection;
                }
                catch { }

                return null;
            }
            catch
            {
                return null;
            }
        }

        public void RememberSlideFromActiveWindow()
        {
            try
            {
                if (_application == null || _application.Presentations.Count == 0) return;
                PowerPoint.DocumentWindow window = null;
                try { window = _application.ActiveWindow; } catch { }
                RememberSlideFromWindow(window);
            }
            catch { }
        }

        public void RememberSlideFromWindow(PowerPoint.DocumentWindow window)
        {
            if (window == null) return;

            try
            {
                PowerPoint.Slide slideFromView = TryResolveSlideFromObject(window.View.Slide);
                if (slideFromView != null)
                {
                    RememberSlide(slideFromView);
                    return;
                }
            }
            catch { }

            try { RememberSlideFromSelection(window.Selection); } catch { }
        }

        public void RememberSlideFromSelection(PowerPoint.Selection selection)
        {
            try
            {
                PowerPoint.Slide slide = TryGetSlideFromSelection(selection);
                if (slide != null) RememberSlide(slide);
            }
            catch { }
        }

        private PowerPoint.Slide TryGetSlideFromSelection(PowerPoint.Selection selection)
        {
            if (selection == null) return null;

            try
            {
                if (selection.SlideRange != null && selection.SlideRange.Count >= 1)
                {
                    PowerPoint.Slide slide = TryResolveSlideFromObject(selection.SlideRange[1]);
                    if (slide != null) return slide;
                }
            }
            catch { }

            try
            {
                if (selection.ShapeRange != null && selection.ShapeRange.Count >= 1)
                {
                    PowerPoint.Slide slide = TryResolveSlideFromObject(selection.ShapeRange[1]);
                    if (slide != null) return slide;
                }
            }
            catch { }

            try
            {
                if (selection.TextRange != null)
                {
                    PowerPoint.Slide slide = TryResolveSlideFromObject(selection.TextRange);
                    if (slide != null) return slide;
                }
            }
            catch { }

            return null;
        }

        private PowerPoint.Slide TryResolveSlideFromObject(object startObject)
        {
            if (startObject == null) return null;

            object currentObject = startObject;
            for (int i = 0; i < 12; i++)
            {
                if (currentObject == null) return null;

                try
                {
                    PowerPoint.Slide directSlide = currentObject as PowerPoint.Slide;
                    if (directSlide != null) return directSlide;
                }
                catch { }

                try
                {
                    dynamic dynamicObject = currentObject;
                    int slideId = dynamicObject.SlideID;
                    PowerPoint.Slide slideById = FindSlideById(slideId);
                    if (slideById != null) return slideById;
                }
                catch { }

                try
                {
                    dynamic dynamicObject = currentObject;
                    int slideIndex = dynamicObject.SlideIndex;
                    PowerPoint.Presentation presentation = GetActivePresentation();
                    if (presentation != null && slideIndex >= 1 && slideIndex <= presentation.Slides.Count)
                    {
                        return presentation.Slides[slideIndex];
                    }
                }
                catch { }

                try
                {
                    dynamic dynamicObject = currentObject;
                    currentObject = dynamicObject.Parent;
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }

        private void RememberSlide(PowerPoint.Slide slide)
        {
            if (slide == null) return;
            try { _lastKnownSlideId = slide.SlideID; } catch { }
        }

        private PowerPoint.Slide GetRememberedSlide()
        {
            try
            {
                if (_lastKnownSlideId <= 0) return null;
                return FindSlideById(_lastKnownSlideId);
            }
            catch
            {
                return null;
            }
        }

        private PowerPoint.Slide FindSlideById(int slideId)
        {
            try
            {
                PowerPoint.Presentation presentation = GetActivePresentation();
                if (presentation == null || slideId <= 0) return null;

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    try
                    {
                        if (slide.SlideID == slideId) return slide;
                    }
                    catch { }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        public void NavigateToSlide(PowerPoint.Slide slide)
        {
            if (slide == null) return;
            RememberSlide(slide);

            try
            {
                if (_application.SlideShowWindows.Count > 0)
                {
                    _application.SlideShowWindows[1].View.GotoSlide(slide.SlideIndex, Office.MsoTriState.msoFalse);
                    NotifyCurrentSlideChanged();
                    return;
                }
            }
            catch { }

            try
            {
                _application.ActiveWindow.View.GotoSlide(slide.SlideIndex);
                RememberSlide(slide);
                NotifyCurrentSlideChanged();
            }
            catch
            {
                ShowMessage($"Could not navigate to slide {slide.SlideIndex}.");
            }
        }

        private PowerPoint.Presentation GetActivePresentation()
        {
            try
            {
                if (_application == null || _application.Presentations.Count == 0) return null;
                return _application.ActivePresentation;
            }
            catch
            {
                return null;
            }
        }

        private string GetTagValue(PowerPoint.Slide slide, string tagName)
        {
            if (slide == null || string.IsNullOrWhiteSpace(tagName)) return string.Empty;

            try
            {
                PowerPoint.Tags tags = slide.Tags;
                for (int i = 1; i <= tags.Count; i++)
                {
                    string currentName = tags.Name(i);
                    if (string.Equals(currentName, tagName, StringComparison.OrdinalIgnoreCase))
                    {
                        return tags.Value(i);
                    }
                }
            }
            catch { }

            return string.Empty;
        }

        private void RemoveTagIfExists(PowerPoint.Slide slide, string tagName)
        {
            if (slide == null || string.IsNullOrWhiteSpace(tagName)) return;

            try
            {
                PowerPoint.Tags tags = slide.Tags;
                for (int i = tags.Count; i >= 1; i--)
                {
                    string currentName = tags.Name(i);
                    if (string.Equals(currentName, tagName, StringComparison.OrdinalIgnoreCase))
                    {
                        tags.Delete(currentName);
                    }
                }
            }
            catch { }
        }

        private string GetSlideTitleOrFirstText(PowerPoint.Slide slide)
        {
            if (slide == null) return string.Empty;

            try
            {
                string firstText = string.Empty;
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame != null && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            string text = shape.TextFrame.TextRange.Text;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                text = text.Replace("\r", " ").Replace("\n", " ").Trim();
                                if (string.IsNullOrWhiteSpace(firstText)) firstText = text;

                                try
                                {
                                    if (shape.Type == Office.MsoShapeType.msoPlaceholder) return ShortenText(text, 90);
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }

                return ShortenText(firstText, 90);
            }
            catch
            {
                return string.Empty;
            }
        }

        private string ShortenText(string text, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            text = text.Trim();
            if (text.Length <= maxLength) return text;
            return text.Substring(0, maxLength - 3) + "...";
        }

        private bool IsHidden(PowerPoint.Slide slide)
        {
            try { return slide.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue; } catch { return false; }
        }

        private void RefreshOpenJumpPointListForm()
        {
            try
            {
                if (_jumpPointListForm != null && !_jumpPointListForm.IsDisposed)
                {
                    _jumpPointListForm.RefreshJumpPointList();
                }
            }
            catch { }
        }

        private void NotifyToggleResult(string message)
        {
            try
            {
                if (_settings.PlayToggleResultSound) SystemSounds.Asterisk.Play();
            }
            catch { }

            if (_settings.ShowToggleResultMessages) ShowMessage(message);
        }

        private void ShowMessage(string message)
        {
            if (!_settings.ShowMessages) return;
            MessageBox.Show(message, AddInInfo.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private enum Direction
        {
            Previous,
            Next
        }
    }
}
