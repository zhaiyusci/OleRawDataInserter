using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OLEPackager.WordAddIn
{
    internal static class WordShapeStateAdapter
    {
        public static OfficeShapeState Capture(Word.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            return new OfficeShapeState
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
                ZOrderPosition = CaptureEffectiveZOrder(shape),
                Name = TryRead(delegate { return shape.Name; }, null),
                AlternativeText = TryRead(delegate { return shape.AlternativeText; }, null),
                Title = TryRead(delegate { return shape.Title; }, null),
                Picture = TryRead(delegate { return CapturePicture(shape.PictureFormat); }, null),
                Line = TryRead(delegate { return CaptureLine(shape.Line); }, null),
                Shadow = TryRead(delegate { return CaptureShadow(shape.Shadow); }, null)
            };
        }

        public static OfficeShapeState Capture(Word.InlineShape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            return new OfficeShapeState
            {
                Width = shape.Width,
                Height = shape.Height,
                LockAspectRatio = TryRead(delegate { return (int)shape.LockAspectRatio; }, (int)Office.MsoTriState.msoFalse),
                AlternativeText = TryRead(delegate { return shape.AlternativeText; }, null),
                Title = TryRead(delegate { return shape.Title; }, null),
                Picture = TryRead(delegate { return CapturePicture(shape.PictureFormat); }, null),
                Line = TryRead(delegate { return CaptureLine(shape.Line); }, null),
                Shadow = TryRead(delegate { return CaptureShadow(shape.Shadow); }, null)
            };
        }

        public static void Apply(Word.Shape shape, OfficeShapeState state, bool includeIdentity)
        {
            if (shape == null || state == null)
            {
                return;
            }

            TryWrite(delegate { shape.LockAspectRatio = Office.MsoTriState.msoFalse; });
            ApplyPicture(shape.PictureFormat, state.Picture);
            ApplyLine(shape.Line, state.Line);
            ApplyShadow(shape.Shadow, state.Shadow);
            TryWrite(delegate { shape.Width = state.Width; });
            TryWrite(delegate { shape.Height = state.Height; });
            TryWrite(delegate { shape.Rotation = state.Rotation; });
            ApplyFlip(shape, Office.MsoFlipCmd.msoFlipHorizontal, state.HorizontalFlip, true);
            ApplyFlip(shape, Office.MsoFlipCmd.msoFlipVertical, state.VerticalFlip, false);
            TryWrite(delegate { shape.Left = state.Left; });
            TryWrite(delegate { shape.Top = state.Top; });
            TryWrite(delegate { shape.Visible = (Office.MsoTriState)state.Visible; });
            TryWrite(delegate { shape.LockAspectRatio = (Office.MsoTriState)state.LockAspectRatio; });

            if (includeIdentity)
            {
                ApplyIdentity(shape, state);
            }
        }

        public static void Apply(Word.InlineShape shape, OfficeShapeState state)
        {
            if (shape == null || state == null)
            {
                return;
            }

            TryWrite(delegate { shape.LockAspectRatio = Office.MsoTriState.msoFalse; });
            ApplyPicture(shape.PictureFormat, state.Picture);
            ApplyLine(shape.Line, state.Line);
            ApplyShadow(shape.Shadow, state.Shadow);
            TryWrite(delegate { shape.Width = state.Width; });
            TryWrite(delegate { shape.Height = state.Height; });
            TryWrite(delegate { shape.LockAspectRatio = (Office.MsoTriState)state.LockAspectRatio; });
            TryWrite(delegate { shape.AlternativeText = state.AlternativeText ?? string.Empty; });
            TryWrite(delegate { shape.Title = state.Title ?? string.Empty; });
        }

        public static void ApplyIdentity(Word.Shape shape, OfficeShapeState state)
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

        public static void MoveToZOrder(Word.Shape shape, int targetPosition)
        {
            if (shape == null || targetPosition <= 0)
            {
                return;
            }

            // A newly converted OLE shape starts at the top of Word's drawing layer.
            // OLE shapes report a sentinel ZOrderPosition, so count backward moves.
            try
            {
                int shapeCount = shape.Anchor.Document.Shapes.Count;
                int backwardMoves = Math.Max(0, shapeCount - targetPosition);
                for (int move = 0; move < backwardMoves; move++)
                {
                    shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                }
            }
            catch (COMException)
            {
            }
        }

        private static OfficePictureState CapturePicture(Word.PictureFormat picture)
        {
            if (picture == null)
            {
                return null;
            }

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

        private static int CaptureEffectiveZOrder(Word.Shape shape)
        {
            int rawPosition = TryRead(delegate { return shape.ZOrderPosition; }, 0);
            Word.Shapes shapes;
            try
            {
                shapes = shape.Anchor.Document.Shapes;
            }
            catch (Exception)
            {
                return rawPosition > 0 && rawPosition < 1000000 ? rawPosition : 1;
            }

            int count = shapes.Count;
            if (rawPosition >= 1 && rawPosition <= count)
            {
                return rawPosition;
            }

            HashSet<int> occupied = new HashSet<int>();
            List<int> sentinelShapeIds = new List<int>();
            for (int index = 1; index <= count; index++)
            {
                Word.Shape candidate = shapes[index];
                int position = TryRead(delegate { return candidate.ZOrderPosition; }, 0);
                if (position >= 1 && position <= count)
                {
                    occupied.Add(position);
                }
                else
                {
                    sentinelShapeIds.Add(TryRead(delegate { return candidate.ID; }, 0));
                }
            }

            List<int> missing = new List<int>();
            for (int position = 1; position <= count; position++)
            {
                if (!occupied.Contains(position))
                {
                    missing.Add(position);
                }
            }

            int shapeId = TryRead(delegate { return shape.ID; }, 0);
            int sentinelIndex = sentinelShapeIds.IndexOf(shapeId);
            if (sentinelIndex >= 0 && sentinelIndex < missing.Count)
            {
                return missing[sentinelIndex];
            }

            return missing.Count > 0 ? missing[0] : Math.Max(1, count);
        }

        private static OfficeLineState CaptureLine(Word.LineFormat line)
        {
            if (line == null)
            {
                return null;
            }

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

        private static OfficeShadowState CaptureShadow(Word.ShadowFormat shadow)
        {
            if (shadow == null)
            {
                return null;
            }

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

        private static void ApplyPicture(Word.PictureFormat picture, OfficePictureState state)
        {
            if (picture == null || state == null)
            {
                return;
            }

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

        private static void ApplyLine(Word.LineFormat line, OfficeLineState state)
        {
            if (line == null || state == null)
            {
                return;
            }

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

        private static void ApplyShadow(Word.ShadowFormat shadow, OfficeShadowState state)
        {
            if (shadow == null || state == null)
            {
                return;
            }

            TryWrite(delegate { shadow.ForeColor.RGB = state.ForeColorRgb; });
            TryWrite(delegate { shadow.Transparency = state.Transparency; });
            TryWrite(delegate { shadow.Type = (Office.MsoShadowType)state.Type; });
            TryWrite(delegate { shadow.OffsetX = state.OffsetX; });
            TryWrite(delegate { shadow.OffsetY = state.OffsetY; });
            TryWrite(delegate { shadow.Blur = state.Blur; });
            TryWrite(delegate { shadow.Size = state.Size; });
            TryWrite(delegate { shadow.RotateWithShape = (Office.MsoTriState)state.RotateWithShape; });
            TryWrite(delegate { shadow.Obscured = (Office.MsoTriState)state.Obscured; });
            TryWrite(delegate { shadow.Visible = (Office.MsoTriState)state.Visible; });
        }

        private static void ApplyFlip(Word.Shape shape, Office.MsoFlipCmd command, int desiredState, bool horizontal)
        {
            if (desiredState != (int)Office.MsoTriState.msoTrue && desiredState != (int)Office.MsoTriState.msoFalse)
            {
                return;
            }

            int currentState = TryRead(
                horizontal
                    ? (Func<int>)delegate { return (int)shape.HorizontalFlip; }
                    : delegate { return (int)shape.VerticalFlip; },
                desiredState);
            if (currentState != desiredState)
            {
                TryWrite(delegate { shape.Flip(command); });
            }
        }

        private static T TryRead<T>(Func<T> reader, T fallback)
        {
            try
            {
                return reader();
            }
            catch (COMException)
            {
                return fallback;
            }
            catch (UnauthorizedAccessException)
            {
                return fallback;
            }
            catch (OutOfMemoryException)
            {
                return fallback;
            }
            catch (NotImplementedException)
            {
                return fallback;
            }
            catch (NotSupportedException)
            {
                return fallback;
            }
        }

        private static void TryWrite(Action writer)
        {
            try
            {
                writer();
            }
            catch (COMException)
            {
            }
            catch (ArgumentException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            catch (OutOfMemoryException)
            {
            }
            catch (NotImplementedException)
            {
            }
            catch (NotSupportedException)
            {
            }
        }
    }
}
