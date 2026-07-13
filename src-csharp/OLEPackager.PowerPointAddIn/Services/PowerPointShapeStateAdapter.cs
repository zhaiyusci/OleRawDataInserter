using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OLEPackager.PowerPointAddIn
{
    internal static class PowerPointShapeStateAdapter
    {
        public static PowerPointShapeState Capture(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            OfficeShapeState common = new OfficeShapeState
            {
                Width = shape.Width,
                Height = shape.Height,
                Left = shape.Left,
                Top = shape.Top,
                Rotation = TryRead(delegate { return shape.Rotation; }, 0f),
                Visible = TryRead(delegate { return (int)shape.Visible; }, (int)Office.MsoTriState.msoTrue),
                HorizontalFlip = TryRead(delegate { return (int)shape.HorizontalFlip; }, (int)Office.MsoTriState.msoFalse),
                VerticalFlip = TryRead(delegate { return (int)shape.VerticalFlip; }, (int)Office.MsoTriState.msoFalse),
                LockAspectRatio = TryRead(delegate { return (int)shape.LockAspectRatio; }, (int)Office.MsoTriState.msoFalse),
                ZOrderPosition = TryRead(delegate { return shape.ZOrderPosition; }, 0),
                Name = TryRead(delegate { return shape.Name; }, null),
                AlternativeText = TryRead(delegate { return shape.AlternativeText; }, null),
                Title = TryRead(delegate { return shape.Title; }, null),
                Picture = TryRead(delegate { return CapturePicture(shape.PictureFormat); }, null),
                Line = TryRead(delegate { return CaptureLine(shape.Line); }, null),
                Shadow = TryRead(delegate { return CaptureShadow(shape.Shadow); }, null)
            };

            return new PowerPointShapeState
            {
                Common = common,
                BlackWhiteMode = TryRead(delegate { return (int)shape.BlackWhiteMode; }, (int)Office.MsoBlackWhiteMode.msoBlackWhiteAutomatic),
                Tags = CaptureTags(shape.Tags),
                ClickAction = CaptureAction(shape, PowerPoint.PpMouseActivation.ppMouseClick),
                MouseOverAction = CaptureAction(shape, PowerPoint.PpMouseActivation.ppMouseOver)
            };
        }

        public static void Apply(PowerPoint.Shape shape, PowerPointShapeState state, bool includeIdentity)
        {
            if (shape == null || state == null || state.Common == null)
            {
                return;
            }

            OfficeShapeState common = state.Common;
            TryWrite(delegate { shape.LockAspectRatio = Office.MsoTriState.msoFalse; });
            ApplyPicture(shape.PictureFormat, common.Picture);
            ApplyLine(shape.Line, common.Line);
            ApplyShadow(shape.Shadow, common.Shadow);
            TryWrite(delegate { shape.Width = common.Width; });
            TryWrite(delegate { shape.Height = common.Height; });
            TryWrite(delegate { shape.Rotation = common.Rotation; });
            ApplyFlip(shape, Office.MsoFlipCmd.msoFlipHorizontal, common.HorizontalFlip, true);
            ApplyFlip(shape, Office.MsoFlipCmd.msoFlipVertical, common.VerticalFlip, false);
            TryWrite(delegate { shape.Left = common.Left; });
            TryWrite(delegate { shape.Top = common.Top; });
            TryWrite(delegate { shape.Visible = (Office.MsoTriState)common.Visible; });
            TryWrite(delegate { shape.BlackWhiteMode = (Office.MsoBlackWhiteMode)state.BlackWhiteMode; });
            TryWrite(delegate { shape.LockAspectRatio = (Office.MsoTriState)common.LockAspectRatio; });

            ApplyTags(shape.Tags, state.Tags);
            ApplyAction(shape, PowerPoint.PpMouseActivation.ppMouseClick, state.ClickAction);
            ApplyAction(shape, PowerPoint.PpMouseActivation.ppMouseOver, state.MouseOverAction);

            if (includeIdentity)
            {
                ApplyIdentity(shape, common);
            }
        }

        public static void ApplyIdentity(PowerPoint.Shape shape, OfficeShapeState state)
        {
            if (shape == null || state == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(state.Name))
            {
                TryWrite(delegate { shape.Name = state.Name; });
            }

            TryWrite(delegate { shape.AlternativeText = state.AlternativeText ?? string.Empty; });
            TryWrite(delegate { shape.Title = state.Title ?? string.Empty; });
        }

        public static void MoveToZOrder(PowerPoint.Shape shape, int targetPosition)
        {
            if (shape == null || targetPosition <= 0)
            {
                return;
            }

            for (int attempt = 0; attempt < 1024; attempt++)
            {
                int currentPosition = TryRead(delegate { return shape.ZOrderPosition; }, targetPosition);
                if (currentPosition == targetPosition)
                {
                    return;
                }

                try
                {
                    shape.ZOrder(currentPosition > targetPosition
                        ? Office.MsoZOrderCmd.msoSendBackward
                        : Office.MsoZOrderCmd.msoBringForward);
                }
                catch (COMException)
                {
                    return;
                }
            }
        }

        public static void TransferAnimations(PowerPoint.Slide slide, PowerPoint.Shape original, PowerPoint.Shape replacement)
        {
            if (slide == null || original == null || replacement == null)
            {
                return;
            }

            int originalId = original.Id;
            TryTransferSequence(slide.TimeLine.MainSequence, originalId, replacement);

            try
            {
                PowerPoint.Sequences sequences = slide.TimeLine.InteractiveSequences;
                for (int index = 1; index <= sequences.Count; index++)
                {
                    TryTransferSequence(sequences[index], originalId, replacement);
                }
            }
            catch (COMException)
            {
            }
        }

        private static void TryTransferSequence(PowerPoint.Sequence sequence, int originalId, PowerPoint.Shape replacement)
        {
            if (sequence == null)
            {
                return;
            }

            try
            {
                for (int index = 1; index <= sequence.Count; index++)
                {
                    PowerPoint.Effect effect = sequence[index];
                    try
                    {
                        if (effect.Shape != null && effect.Shape.Id == originalId)
                        {
                            effect.Shape = replacement;
                        }
                    }
                    catch (COMException)
                    {
                    }

                    try
                    {
                        PowerPoint.Shape triggerShape = effect.Timing.TriggerShape;
                        if (triggerShape != null && triggerShape.Id == originalId)
                        {
                            effect.Timing.TriggerShape = replacement;
                        }
                    }
                    catch (COMException)
                    {
                    }
                }
            }
            catch (COMException)
            {
            }
        }

        private static OfficePictureState CapturePicture(PowerPoint.PictureFormat picture)
        {
            return new OfficePictureState
            {
                Brightness = TryRead(delegate { return picture.Brightness; }, 0.5f),
                Contrast = TryRead(delegate { return picture.Contrast; }, 0.5f),
                CropBottom = TryRead(delegate { return picture.CropBottom; }, 0f),
                CropLeft = TryRead(delegate { return picture.CropLeft; }, 0f),
                CropRight = TryRead(delegate { return picture.CropRight; }, 0f),
                CropTop = TryRead(delegate { return picture.CropTop; }, 0f),
                ColorType = TryRead(delegate { return (int)picture.ColorType; }, (int)Office.MsoPictureColorType.msoPictureAutomatic),
                TransparencyColor = TryRead(delegate { return picture.TransparencyColor; }, 0),
                TransparentBackground = TryRead(delegate { return (int)picture.TransparentBackground; }, (int)Office.MsoTriState.msoFalse)
            };
        }

        private static OfficeLineState CaptureLine(PowerPoint.LineFormat line)
        {
            return new OfficeLineState
            {
                Visible = TryRead(delegate { return (int)line.Visible; }, (int)Office.MsoTriState.msoFalse),
                ForeColorRgb = TryRead(delegate { return line.ForeColor.RGB; }, 0),
                BackColorRgb = TryRead(delegate { return line.BackColor.RGB; }, 0),
                Transparency = TryRead(delegate { return line.Transparency; }, 0f),
                Weight = TryRead(delegate { return line.Weight; }, 1f),
                DashStyle = TryRead(delegate { return (int)line.DashStyle; }, (int)Office.MsoLineDashStyle.msoLineSolid),
                Style = TryRead(delegate { return (int)line.Style; }, (int)Office.MsoLineStyle.msoLineSingle),
                InsetPen = TryRead(delegate { return (int)line.InsetPen; }, (int)Office.MsoTriState.msoFalse),
                Pattern = TryRead(delegate { return (int)line.Pattern; }, (int)Office.MsoPatternType.msoPatternMixed),
                BeginArrowheadLength = TryRead(delegate { return (int)line.BeginArrowheadLength; }, (int)Office.MsoArrowheadLength.msoArrowheadLengthMedium),
                BeginArrowheadStyle = TryRead(delegate { return (int)line.BeginArrowheadStyle; }, (int)Office.MsoArrowheadStyle.msoArrowheadNone),
                BeginArrowheadWidth = TryRead(delegate { return (int)line.BeginArrowheadWidth; }, (int)Office.MsoArrowheadWidth.msoArrowheadWidthMedium),
                EndArrowheadLength = TryRead(delegate { return (int)line.EndArrowheadLength; }, (int)Office.MsoArrowheadLength.msoArrowheadLengthMedium),
                EndArrowheadStyle = TryRead(delegate { return (int)line.EndArrowheadStyle; }, (int)Office.MsoArrowheadStyle.msoArrowheadNone),
                EndArrowheadWidth = TryRead(delegate { return (int)line.EndArrowheadWidth; }, (int)Office.MsoArrowheadWidth.msoArrowheadWidthMedium)
            };
        }

        private static OfficeShadowState CaptureShadow(PowerPoint.ShadowFormat shadow)
        {
            return new OfficeShadowState
            {
                Visible = TryRead(delegate { return (int)shadow.Visible; }, (int)Office.MsoTriState.msoFalse),
                ForeColorRgb = TryRead(delegate { return shadow.ForeColor.RGB; }, 0),
                Transparency = TryRead(delegate { return shadow.Transparency; }, 0f),
                Type = TryRead(delegate { return (int)shadow.Type; }, (int)Office.MsoShadowType.msoShadow1),
                Style = TryRead(delegate { return (int)shadow.Style; }, (int)Office.MsoShadowStyle.msoShadowStyleOuterShadow),
                OffsetX = TryRead(delegate { return shadow.OffsetX; }, 0f),
                OffsetY = TryRead(delegate { return shadow.OffsetY; }, 0f),
                Blur = TryRead(delegate { return shadow.Blur; }, 0f),
                Size = TryRead(delegate { return shadow.Size; }, 100f),
                RotateWithShape = TryRead(delegate { return (int)shadow.RotateWithShape; }, (int)Office.MsoTriState.msoFalse),
                Obscured = TryRead(delegate { return (int)shadow.Obscured; }, (int)Office.MsoTriState.msoFalse)
            };
        }

        private static Dictionary<string, string> CaptureTags(PowerPoint.Tags tags)
        {
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                for (int index = 1; index <= tags.Count; index++)
                {
                    string name = tags.Name(index);
                    result[name] = tags[name];
                }
            }
            catch (COMException)
            {
            }

            return result;
        }

        private static PowerPointActionState CaptureAction(PowerPoint.Shape shape, PowerPoint.PpMouseActivation activation)
        {
            try
            {
                PowerPoint.ActionSetting action = shape.ActionSettings[activation];
                return new PowerPointActionState
                {
                    Action = (int)action.Action,
                    ActionVerb = action.ActionVerb,
                    AnimateAction = (int)action.AnimateAction,
                    Run = action.Run,
                    ShowAndReturn = (int)action.ShowAndReturn,
                    SlideShowName = action.SlideShowName,
                    HyperlinkAddress = action.Hyperlink.Address,
                    HyperlinkSubAddress = action.Hyperlink.SubAddress,
                    HyperlinkScreenTip = action.Hyperlink.ScreenTip,
                    SoundName = action.SoundEffect.Name,
                    SoundType = (int)action.SoundEffect.Type
                };
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static void ApplyPicture(PowerPoint.PictureFormat picture, OfficePictureState state)
        {
            if (picture == null || state == null) { return; }
            TryWrite(delegate { picture.Brightness = state.Brightness; });
            TryWrite(delegate { picture.Contrast = state.Contrast; });
            TryWrite(delegate { picture.ColorType = (Office.MsoPictureColorType)state.ColorType; });
            TryWrite(delegate { picture.TransparencyColor = state.TransparencyColor; });
            TryWrite(delegate { picture.TransparentBackground = (Office.MsoTriState)state.TransparentBackground; });
            TryWrite(delegate { picture.CropBottom = state.CropBottom; });
            TryWrite(delegate { picture.CropLeft = state.CropLeft; });
            TryWrite(delegate { picture.CropRight = state.CropRight; });
            TryWrite(delegate { picture.CropTop = state.CropTop; });
        }

        private static void ApplyLine(PowerPoint.LineFormat line, OfficeLineState state)
        {
            if (line == null || state == null) { return; }
            TryWrite(delegate { line.ForeColor.RGB = state.ForeColorRgb; });
            TryWrite(delegate { line.BackColor.RGB = state.BackColorRgb; });
            TryWrite(delegate { line.Transparency = state.Transparency; });
            TryWrite(delegate { line.Weight = state.Weight; });
            TryWrite(delegate { line.DashStyle = (Office.MsoLineDashStyle)state.DashStyle; });
            TryWrite(delegate { line.Style = (Office.MsoLineStyle)state.Style; });
            TryWrite(delegate { line.InsetPen = (Office.MsoTriState)state.InsetPen; });
            if (state.Pattern > 0)
            {
                TryWrite(delegate { line.Pattern = (Office.MsoPatternType)state.Pattern; });
            }
            TryWrite(delegate { line.BeginArrowheadLength = (Office.MsoArrowheadLength)state.BeginArrowheadLength; });
            TryWrite(delegate { line.BeginArrowheadStyle = (Office.MsoArrowheadStyle)state.BeginArrowheadStyle; });
            TryWrite(delegate { line.BeginArrowheadWidth = (Office.MsoArrowheadWidth)state.BeginArrowheadWidth; });
            TryWrite(delegate { line.EndArrowheadLength = (Office.MsoArrowheadLength)state.EndArrowheadLength; });
            TryWrite(delegate { line.EndArrowheadStyle = (Office.MsoArrowheadStyle)state.EndArrowheadStyle; });
            TryWrite(delegate { line.EndArrowheadWidth = (Office.MsoArrowheadWidth)state.EndArrowheadWidth; });
            TryWrite(delegate { line.Visible = (Office.MsoTriState)state.Visible; });
        }

        private static void ApplyShadow(PowerPoint.ShadowFormat shadow, OfficeShadowState state)
        {
            if (shadow == null || state == null) { return; }
            TryWrite(delegate { shadow.ForeColor.RGB = state.ForeColorRgb; });
            TryWrite(delegate { shadow.Transparency = state.Transparency; });
            TryWrite(delegate { shadow.Type = (Office.MsoShadowType)state.Type; });
            TryWrite(delegate { shadow.Style = (Office.MsoShadowStyle)state.Style; });
            TryWrite(delegate { shadow.OffsetX = state.OffsetX; });
            TryWrite(delegate { shadow.OffsetY = state.OffsetY; });
            TryWrite(delegate { shadow.Blur = state.Blur; });
            TryWrite(delegate { shadow.Size = state.Size; });
            TryWrite(delegate { shadow.RotateWithShape = (Office.MsoTriState)state.RotateWithShape; });
            TryWrite(delegate { shadow.Obscured = (Office.MsoTriState)state.Obscured; });
            TryWrite(delegate { shadow.Visible = (Office.MsoTriState)state.Visible; });
        }

        private static void ApplyTags(PowerPoint.Tags tags, Dictionary<string, string> values)
        {
            if (tags == null || values == null) { return; }
            foreach (KeyValuePair<string, string> item in values)
            {
                TryWrite(delegate { tags.Add(item.Key, item.Value ?? string.Empty); });
            }
        }

        private static void ApplyAction(PowerPoint.Shape shape, PowerPoint.PpMouseActivation activation, PowerPointActionState state)
        {
            if (state == null) { return; }
            try
            {
                PowerPoint.ActionSetting action = shape.ActionSettings[activation];
                TryWrite(delegate { action.ActionVerb = state.ActionVerb ?? string.Empty; });
                TryWrite(delegate { action.AnimateAction = (Office.MsoTriState)state.AnimateAction; });
                TryWrite(delegate { action.Run = state.Run ?? string.Empty; });
                TryWrite(delegate { action.ShowAndReturn = (Office.MsoTriState)state.ShowAndReturn; });
                TryWrite(delegate { action.SlideShowName = state.SlideShowName ?? string.Empty; });
                TryWrite(delegate { action.Hyperlink.Address = state.HyperlinkAddress ?? string.Empty; });
                TryWrite(delegate { action.Hyperlink.SubAddress = state.HyperlinkSubAddress ?? string.Empty; });
                TryWrite(delegate { action.Hyperlink.ScreenTip = state.HyperlinkScreenTip ?? string.Empty; });
                TryWrite(delegate { action.SoundEffect.Name = state.SoundName ?? string.Empty; });
                TryWrite(delegate { action.SoundEffect.Type = (PowerPoint.PpSoundEffectType)state.SoundType; });
                TryWrite(delegate { action.Action = (PowerPoint.PpActionType)state.Action; });
            }
            catch (COMException)
            {
            }
        }

        private static void ApplyFlip(PowerPoint.Shape shape, Office.MsoFlipCmd command, int desiredState, bool horizontal)
        {
            if (desiredState != (int)Office.MsoTriState.msoTrue && desiredState != (int)Office.MsoTriState.msoFalse) { return; }
            int currentState = TryRead(horizontal ? (Func<int>)delegate { return (int)shape.HorizontalFlip; } : delegate { return (int)shape.VerticalFlip; }, desiredState);
            if (currentState != desiredState) { TryWrite(delegate { shape.Flip(command); }); }
        }

        private static T TryRead<T>(Func<T> reader, T fallback)
        {
            try { return reader(); }
            catch (COMException) { return fallback; }
            catch (UnauthorizedAccessException) { return fallback; }
            catch (OutOfMemoryException) { return fallback; }
            catch (NotImplementedException) { return fallback; }
            catch (NotSupportedException) { return fallback; }
        }

        private static void TryWrite(Action writer)
        {
            try { writer(); }
            catch (COMException) { }
            catch (ArgumentException) { }
            catch (UnauthorizedAccessException) { }
            catch (OutOfMemoryException) { }
            catch (NotImplementedException) { }
            catch (NotSupportedException) { }
        }
    }

    internal sealed class PowerPointShapeState
    {
        public OfficeShapeState Common { get; set; }
        public int BlackWhiteMode { get; set; }
        public Dictionary<string, string> Tags { get; set; }
        public PowerPointActionState ClickAction { get; set; }
        public PowerPointActionState MouseOverAction { get; set; }

        public bool RequiresRenderedDisplay
        {
            get
            {
                return Common != null
                    && (Math.Abs(Common.Rotation) > 0.01f
                        || Common.HorizontalFlip == (int)Office.MsoTriState.msoTrue
                        || Common.VerticalFlip == (int)Office.MsoTriState.msoTrue);
            }
        }

        public void PrepareForRenderedDisplay()
        {
            RenderedImageService.PrepareStateForRenderedDisplay(Common);
        }
    }

    internal sealed class PowerPointActionState
    {
        public int Action { get; set; }
        public string ActionVerb { get; set; }
        public int AnimateAction { get; set; }
        public string Run { get; set; }
        public int ShowAndReturn { get; set; }
        public string SlideShowName { get; set; }
        public string HyperlinkAddress { get; set; }
        public string HyperlinkSubAddress { get; set; }
        public string HyperlinkScreenTip { get; set; }
        public string SoundName { get; set; }
        public int SoundType { get; set; }
    }
}
