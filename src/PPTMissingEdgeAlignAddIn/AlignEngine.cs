using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTMissingEdgeAlignAddIn
{
    public enum AlignState
    {
        Idle,
        Armed
    }

    public enum SoundType
    {
        Beep,         // Arming
        Exclamation,  // Missed click
        Hand          // Error / Self-snap
    }

    public static class AlignEngine
    {
        private static AlignState _state = AlignState.Idle;
        private static List<PowerPoint.Shape> _sourceShapes = new List<PowerPoint.Shape>();
        private static string _currentStatus = string.Empty;

        public static AlignState State
        {
            get { return _state; }
        }

        private static void SetStatusBar(PowerPoint.Application app, string text)
        {
            _currentStatus = text;
            System.Diagnostics.Debug.WriteLine("StatusBar: " + text);
        }

        public static void ArmEngine()
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;

            // Pre-Flight Validation
            PowerPoint.Selection sel = null;
            try
            {
                sel = app.ActiveWindow.Selection;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error accessing selection: " + ex.Message);
            }

            if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 1)
            {
                PlaySound(SoundType.Hand);
                try
                {
                    SetStatusBar(app, "⚠️ MissingEdgeAlign: Please select a source shape first.");
                    ClearStatusBarAsync(app);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error setting StatusBar: " + ex.Message);
                }
                InvalidateRibbon();
                return;
            }

            // Arming the Engine
            try
            {
                _sourceShapes.Clear();
                for (int i = 1; i <= sel.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape topShape = GetTopLevelMovingShape(sel.ShapeRange[i]);
                    if (!ContainsShapeId(_sourceShapes, topShape.Id))
                    {
                        _sourceShapes.Add(topShape);
                    }
                }

                _state = AlignState.Armed;
                PlaySound(SoundType.Beep);
                SetStatusBar(app, "⚡ EdgeAlign ARMED: Now click the target shape (hold CTRL for gap)...");

                // Unsubscribe first just in case to avoid double subscription
                app.WindowSelectionChange -= OnTargetSelected;
                app.WindowSelectionChange += OnTargetSelected;

                InvalidateRibbon();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error arming engine: " + ex.Message);
                ResetToIdle(app);
            }
        }

        public static void CancelEngine()
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            try
            {
                app.WindowSelectionChange -= OnTargetSelected;
            }
            catch { }
            ResetToIdle(app);
        }

        private static void OnTargetSelected(PowerPoint.Selection Sel)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;

            if (_state != AlignState.Armed)
            {
                try
                {
                    app.WindowSelectionChange -= OnTargetSelected;
                }
                catch { }
                return;
            }

            try
            {
                app.WindowSelectionChange -= OnTargetSelected;

                if (Sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Sel.ShapeRange.Count < 1)
                {
                    PlaySound(SoundType.Exclamation);
                    ResetToIdle(app);
                    return;
                }

                // Look for target shape in selection (first shape that is not in the source selection)
                PowerPoint.Shape targetShape = null;
                for (int i = 1; i <= Sel.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape topTarget = GetTopLevelMovingShape(Sel.ShapeRange[i]);
                    if (!IsSourceShape(topTarget))
                    {
                        targetShape = topTarget;
                        break;
                    }
                }

                if (targetShape == null)
                {
                    targetShape = GetTopLevelMovingShape(Sel.ShapeRange[1]);
                }

                // Self-Snap Guard: If target shape is actually in the source shapes, play Hand sound and abort
                if (IsSourceShape(targetShape))
                {
                    PlaySound(SoundType.Hand);
                    ResetToIdle(app);
                    return;
                }

                if (_sourceShapes.Count > 0)
                {
                    // Vector Calculation & Alignment
                    bool useModifier = (Control.ModifierKeys & Keys.Control) == Keys.Control;
                    int gap = useModifier ? Properties.Settings.Default.DefaultGap : 0;

                    // Calculate combined bounding box of all source shapes
                    float minLeft = float.MaxValue;
                    float maxRight = float.MinValue;
                    float minTop = float.MaxValue;
                    float maxBottom = float.MinValue;

                    foreach (var s in _sourceShapes)
                    {
                        minLeft = Math.Min(minLeft, s.Left);
                        maxRight = Math.Max(maxRight, s.Left + s.Width);
                        minTop = Math.Min(minTop, s.Top);
                        maxBottom = Math.Max(maxBottom, s.Top + s.Height);
                    }

                    float selectionWidth = maxRight - minLeft;
                    float selectionHeight = maxBottom - minTop;

                    // Calculate projection overlaps to naturally determine the dominant alignment axis
                    bool overlapsHorizontally = Math.Max(minLeft, targetShape.Left) < 
                                                 Math.Min(maxRight, targetShape.Left + targetShape.Width);
                    bool overlapsVertically = Math.Max(minTop, targetShape.Top) < 
                                               Math.Min(maxBottom, targetShape.Top + targetShape.Height);

                    bool isHorizontalOperation = false;

                    if (overlapsHorizontally && !overlapsVertically)
                    {
                        // Staggered vertically -> align vertically
                        isHorizontalOperation = false;
                    }
                    else if (overlapsVertically && !overlapsHorizontally)
                    {
                        // Staggered horizontally -> align horizontally
                        isHorizontalOperation = true;
                    }
                    else
                    {
                        // Fallback to center-to-center delta comparison
                        float centerXSource = minLeft + selectionWidth / 2f;
                        float centerYSource = minTop + selectionHeight / 2f;
                        float centerXTarget = targetShape.Left + targetShape.Width / 2f;
                        float centerYTarget = targetShape.Top + targetShape.Height / 2f;

                        float deltaX = centerXTarget - centerXSource;
                        float deltaY = centerYTarget - centerYSource;

                        isHorizontalOperation = Math.Abs(deltaX) >= Math.Abs(deltaY);
                    }

                    // Shift-Key Modifier: force opposite alignment axis
                    bool forceOppositeAxis = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
                    if (forceOppositeAxis)
                    {
                        isHorizontalOperation = !isHorizontalOperation;
                    }

                    // Align each source shape independently on the determined dominant axis
                    foreach (var s in _sourceShapes)
                    {
                        if (isHorizontalOperation)
                        {
                            // HORIZONTAL OPERATION
                            if (targetShape.Left >= s.Left)
                            {
                                s.Left = targetShape.Left - s.Width - gap;
                            }
                            else
                            {
                                s.Left = targetShape.Left + targetShape.Width + gap;
                            }
                        }
                        else
                        {
                            // VERTICAL OPERATION
                            if (targetShape.Top >= s.Top)
                            {
                                s.Top = targetShape.Top - s.Height - gap;
                            }
                            else
                            {
                                s.Top = targetShape.Top + targetShape.Height + gap;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("OnTargetSelected exception: " + ex.ToString());
                PlaySound(SoundType.Hand);
            }
            finally
            {
                ResetToIdle(app);
            }
        }

        private static bool IsSourceShape(PowerPoint.Shape shape)
        {
            foreach (var s in _sourceShapes)
            {
                if (s != null && s.Id == shape.Id)
                {
                    return true;
                }
            }
            return false;
        }

        private static void ResetToIdle(PowerPoint.Application app)
        {
            _state = AlignState.Idle;
            _sourceShapes.Clear();
            try
            {
                SetStatusBar(app, "");
            }
            catch { }
            InvalidateRibbon();
        }

        private static async void ClearStatusBarAsync(PowerPoint.Application app)
        {
            await Task.Delay(2500);
            try
            {
                if (_currentStatus != null && _currentStatus.Contains("MissingEdgeAlign"))
                {
                    SetStatusBar(app, "");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error in ClearStatusBarAsync: " + ex.Message);
            }
        }

        private static void InvalidateRibbon()
        {
            try
            {
                if (Ribbon.RibbonUI != null)
                {
                    Ribbon.RibbonUI.InvalidateControl("btnEdgeAlign");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error invalidating ribbon: " + ex.Message);
            }
        }

        private static void PlaySound(SoundType type)
        {
            try
            {
                switch (type)
                {
                    case SoundType.Beep:
                        if (Properties.Settings.Default.SoundOnArmed)
                        {
                            System.Media.SystemSounds.Beep.Play();
                        }
                        break;

                    case SoundType.Exclamation:
                        if (Properties.Settings.Default.SoundOnMissedClick)
                        {
                            Task.Run(() =>
                            {
                                try
                                {
                                    int sampleRate = 8000;
                                    int samples = sampleRate * 250 / 1000;
                                    short[] sampleBuffer = new short[samples];
                                    for (int i = 0; i < samples; i++)
                                    {
                                        double t = (double)i / sampleRate;
                                        sampleBuffer[i] = (short)(Math.Sin(2 * Math.PI * 320 * t) * short.MaxValue * 0.4);
                                    }

                                    byte[] wavBytes = CreateWavBytes(sampleBuffer, sampleRate);
                                    using (var ms = new System.IO.MemoryStream(wavBytes))
                                    {
                                        using (var player = new System.Media.SoundPlayer(ms))
                                        {
                                            player.PlaySync();
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine("Error playing Exclamation WAV: " + ex.Message);
                                    System.Media.SystemSounds.Exclamation.Play();
                                }
                            });
                        }
                        break;

                    case SoundType.Hand:
                        if (Properties.Settings.Default.SoundOnError)
                        {
                            Task.Run(() =>
                            {
                                try
                                {
                                    int sampleRate = 8000;
                                    int samples1 = sampleRate * 100 / 1000;
                                    int silence = sampleRate * 40 / 1000;
                                    int samples2 = sampleRate * 150 / 1000;
                                    short[] sampleBuffer = new short[samples1 + silence + samples2];

                                    for (int i = 0; i < samples1; i++)
                                    {
                                        double t = (double)i / sampleRate;
                                        sampleBuffer[i] = (short)(Math.Sin(2 * Math.PI * 240 * t) * short.MaxValue * 0.4);
                                    }
                                    for (int i = 0; i < samples2; i++)
                                    {
                                        double t = (double)i / sampleRate;
                                        sampleBuffer[samples1 + silence + i] = (short)(Math.Sin(2 * Math.PI * 240 * t) * short.MaxValue * 0.4);
                                    }

                                    byte[] wavBytes = CreateWavBytes(sampleBuffer, sampleRate);
                                    using (var ms = new System.IO.MemoryStream(wavBytes))
                                    {
                                        using (var player = new System.Media.SoundPlayer(ms))
                                        {
                                            player.PlaySync();
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine("Error playing Hand WAV: " + ex.Message);
                                    System.Media.SystemSounds.Hand.Play();
                                }
                            });
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("PlaySound outer exception: " + ex.Message);
            }
        }

        private static byte[] CreateWavBytes(short[] sampleBuffer, int sampleRate)
        {
            int dataLength = sampleBuffer.Length * 2;
            int fileLength = 36 + dataLength;
            byte[] wav = new byte[44 + dataLength];

            // RIFF
            wav[0] = 0x52; wav[1] = 0x49; wav[2] = 0x46; wav[3] = 0x46;
            byte[] temp = BitConverter.GetBytes(fileLength);
            Array.Copy(temp, 0, wav, 4, 4);

            // WAVE
            wav[8] = 0x57; wav[9] = 0x41; wav[10] = 0x56; wav[11] = 0x45;

            // fmt
            wav[12] = 0x66; wav[13] = 0x6d; wav[14] = 0x74; wav[15] = 0x20;
            temp = BitConverter.GetBytes(16);
            Array.Copy(temp, 0, wav, 16, 4);

            // PCM (1), Channels (1)
            temp = BitConverter.GetBytes((short)1);
            Array.Copy(temp, 0, wav, 20, 2);
            Array.Copy(temp, 0, wav, 22, 2);

            // Sample rate
            temp = BitConverter.GetBytes(sampleRate);
            Array.Copy(temp, 0, wav, 24, 4);

            // Byte rate
            temp = BitConverter.GetBytes(sampleRate * 2);
            Array.Copy(temp, 0, wav, 28, 4);

            // Block align (2)
            temp = BitConverter.GetBytes((short)2);
            Array.Copy(temp, 0, wav, 32, 2);

            // Bits per sample (16)
            temp = BitConverter.GetBytes((short)16);
            Array.Copy(temp, 0, wav, 34, 2);

            // data
            wav[36] = 0x64; wav[37] = 0x61; wav[38] = 0x74; wav[39] = 0x61;
            temp = BitConverter.GetBytes(dataLength);
            Array.Copy(temp, 0, wav, 40, 4);

            // Data payload
            for (int i = 0; i < sampleBuffer.Length; i++)
            {
                byte[] val = BitConverter.GetBytes(sampleBuffer[i]);
                wav[44 + i * 2] = val[0];
                wav[44 + i * 2 + 1] = val[1];
            }

            return wav;
        }

        private static bool ContainsShapeId(List<PowerPoint.Shape> list, int id)
        {
            foreach (var s in list)
            {
                if (s != null && s.Id == id) return true;
            }
            return false;
        }

        private static PowerPoint.Shape GetTopLevelMovingShape(PowerPoint.Shape shape)
        {
            PowerPoint.Shape current = shape;
            try
            {
                while (current.ParentGroup != null)
                {
                    current = current.ParentGroup;
                }
            }
            catch
            {
                // Ignore parent group lookup errors
            }
            return current;
        }
    }
}
