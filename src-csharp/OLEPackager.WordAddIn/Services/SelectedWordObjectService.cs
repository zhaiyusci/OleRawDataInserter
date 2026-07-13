using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OLEPackager.WordAddIn
{
    internal sealed class SelectedWordObjectService
    {
        private readonly Word.Application application;
        private readonly WordOlePackageService wordPackages;
        private readonly FlatOpcPackageService flatOpcPackages;
        private readonly OlePackageExtractor olePackages;

        public SelectedWordObjectService(
            Word.Application application,
            WordOlePackageService wordPackages,
            FlatOpcPackageService flatOpcPackages,
            OlePackageExtractor olePackages)
        {
            this.application = application;
            this.wordPackages = wordPackages;
            this.flatOpcPackages = flatOpcPackages;
            this.olePackages = olePackages;
        }

        public SelectedWordObject CaptureSelection()
        {
            Word.Selection selection = application.Selection;
            if (selection.InlineShapes.Count > 0)
            {
                return CaptureInline(selection.InlineShapes[1]);
            }

            try
            {
                if (selection.ShapeRange.Count > 0)
                {
                    return CaptureFloating(selection.ShapeRange[1]);
                }
            }
            catch (COMException)
            {
            }

            return null;
        }

        public void ReplaceWithPackage(SelectedWordObject target, string zipPath)
        {
            if (target == null)
            {
                throw new ArgumentNullException("target");
            }

            if (target.InlineShape != null)
            {
                ReplaceInline(target, zipPath);
                return;
            }

            if (target.FloatingShape != null)
            {
                ReplaceFloating(target, zipPath);
                return;
            }

            throw new InvalidOperationException("所选对象已不在文档中。");
        }

        private SelectedWordObject CaptureInline(Word.InlineShape inlineShape)
        {
            bool isPicture = inlineShape.Type == Word.WdInlineShapeType.wdInlineShapePicture
                || inlineShape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture;
            bool isOle = inlineShape.Type == Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                || inlineShape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedOLEObject;
            if (!isPicture && !isOle)
            {
                return null;
            }

            SelectedWordObject result = new SelectedWordObject
            {
                InlineShape = inlineShape,
                IsExistingPackage = isOle,
                Width = inlineShape.Width,
                Height = inlineShape.Height,
                ShapeState = WordShapeStateAdapter.Capture(inlineShape)
            };

            try
            {
                string linkedImage = isPicture ? TryGetLinkedImageSource(inlineShape) : null;
                string openXml = null;
                if (string.IsNullOrEmpty(linkedImage) || isOle)
                {
                    openXml = ReadOpenXml(inlineShape.Range);
                }

                if (!string.IsNullOrEmpty(linkedImage))
                {
                    result.DisplayImagePath = linkedImage;
                }
                else
                {
                    result.DisplayImagePath = flatOpcPackages.ExtractDisplayImage(openXml, isOle ? "selected_ole_display" : "selected_image");
                    result.AddTemporaryFile(result.DisplayImagePath);
                }

                if (isOle)
                {
                    result.ExistingZipPath = olePackages.ExtractZip(flatOpcPackages.ExtractOleBytes(openXml), "selected_ole");
                    result.AddTemporaryFile(result.ExistingZipPath);
                }

                return result;
            }
            catch
            {
                result.Dispose();
                throw;
            }
        }

        private SelectedWordObject CaptureFloating(Word.Shape floatingShape)
        {
            bool isPicture = floatingShape.Type == Office.MsoShapeType.msoPicture
                || floatingShape.Type == Office.MsoShapeType.msoLinkedPicture;
            bool isOle = floatingShape.Type == Office.MsoShapeType.msoEmbeddedOLEObject
                || floatingShape.Type == Office.MsoShapeType.msoLinkedOLEObject;
            if (!isPicture && !isOle)
            {
                return null;
            }

            FloatingShapeLayout layout = FloatingShapeLayout.Capture(floatingShape);
            SelectedWordObject result = new SelectedWordObject
            {
                FloatingShape = floatingShape,
                FloatingLayout = layout,
                IsExistingPackage = isOle,
                Width = floatingShape.Width,
                Height = floatingShape.Height,
                ShapeState = layout.ShapeState
            };

            string linkedImage = isPicture ? TryGetLinkedImageSource(floatingShape) : null;
            if (!string.IsNullOrEmpty(linkedImage) && !isOle && !layout.RequiresRenderedDisplay)
            {
                result.DisplayImagePath = linkedImage;
                return result;
            }

            Word.Shape extractionShape = null;
            Word.InlineShape extractionInline = null;
            try
            {
                extractionShape = floatingShape.Duplicate();
                extractionInline = extractionShape.ConvertToInlineShape();
                extractionShape = null;
                string openXml = null;

                if (layout.RequiresRenderedDisplay)
                {
                    result.DisplayImagePath = RenderedImageService.SaveCroppedMetafilePng(
                        (byte[])extractionInline.Range.EnhMetaFileBits,
                        "selected_word_rendered");
                    layout.PrepareForRenderedDisplay();
                }
                else
                {
                    openXml = ReadOpenXml(extractionInline.Range);
                    result.DisplayImagePath = flatOpcPackages.ExtractDisplayImage(
                        openXml,
                        isOle ? "selected_ole_display" : "selected_image");
                }

                result.AddTemporaryFile(result.DisplayImagePath);

                if (isOle)
                {
                    if (string.IsNullOrEmpty(openXml))
                    {
                        openXml = ReadOpenXml(extractionInline.Range);
                    }

                    result.ExistingZipPath = olePackages.ExtractZip(flatOpcPackages.ExtractOleBytes(openXml), "selected_ole");
                    result.AddTemporaryFile(result.ExistingZipPath);
                }

                extractionInline.Delete();
                extractionInline = null;
                return result;
            }
            catch
            {
                TryDeleteShape(extractionShape, extractionInline);
                result.Dispose();
                throw;
            }
        }

        private void ReplaceInline(SelectedWordObject target, string zipPath)
        {
            Word.InlineShape oldShape = target.InlineShape;
            Word.Range insertionRange = oldShape.Range.Duplicate;
            insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertionRange.Select();

            Word.InlineShape newShape = wordPackages.InsertInlinePackage(
                target.DisplayImagePath,
                zipPath,
                target.Width,
                target.Height);

            try
            {
                WordShapeStateAdapter.Apply(newShape, target.ShapeState);
                oldShape.Delete();
                target.InlineShape = newShape;
            }
            catch
            {
                try
                {
                    newShape.Delete();
                }
                catch (COMException)
                {
                }

                throw;
            }
        }

        private void ReplaceFloating(SelectedWordObject target, string zipPath)
        {
            FloatingShapeLayout layout = target.FloatingLayout;
            Word.Shape oldFloating = target.FloatingShape;
            Word.InlineShape newInline = null;
            Word.Shape newFloating = null;

            try
            {
                Word.Range insertionRange = oldFloating.Anchor.Duplicate;
                insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                insertionRange.Select();

                newInline = wordPackages.InsertInlinePackage(
                    target.DisplayImagePath,
                    zipPath,
                    target.Width,
                    target.Height);
                newFloating = newInline.ConvertToShape();
                newInline = null;
                wordPackages.ApplyDisplayImage(newFloating, target.DisplayImagePath);
                layout.Apply(newFloating, false);
                layout.PlaceImmediatelyAbove(newFloating, oldFloating);

                oldFloating.Delete();
                oldFloating = null;
                layout.FinishReplacement(newFloating);
                target.FloatingShape = newFloating;
            }
            catch
            {
                TryDeleteShape(newFloating, newInline);
                throw;
            }
        }

        private static string ReadOpenXml(Word.Range range)
        {
            COMException lastException = null;
            for (int attempt = 0; attempt < 8; attempt++)
            {
                try
                {
                    string openXml = range.WordOpenXML;
                    if (!string.IsNullOrEmpty(openXml))
                    {
                        return openXml;
                    }
                }
                catch (COMException exception)
                {
                    lastException = exception;
                }

                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }

            throw new InvalidOperationException("Word 无法读取所选对象的 Open XML。", lastException);
        }

        private static string TryGetLinkedImageSource(Word.InlineShape inlineShape)
        {
            try
            {
                Word.LinkFormat linkFormat = inlineShape.LinkFormat;
                if (linkFormat == null)
                {
                    return null;
                }

                string source = linkFormat.SourceFullName;
                return File.Exists(source) ? source : null;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string TryGetLinkedImageSource(Word.Shape shape)
        {
            try
            {
                Word.LinkFormat linkFormat = shape.LinkFormat;
                if (linkFormat == null)
                {
                    return null;
                }

                string source = linkFormat.SourceFullName;
                return File.Exists(source) ? source : null;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static void TryDeleteShape(Word.Shape floatingShape, Word.InlineShape inlineShape)
        {
            try
            {
                if (floatingShape != null)
                {
                    floatingShape.Delete();
                }
                else if (inlineShape != null)
                {
                    inlineShape.Delete();
                }
            }
            catch (COMException)
            {
            }
        }
    }

    internal sealed class SelectedWordObject : IDisposable
    {
        private readonly List<string> temporaryFiles = new List<string>();

        public Word.InlineShape InlineShape { get; set; }

        public Word.Shape FloatingShape { get; set; }

        public FloatingShapeLayout FloatingLayout { get; set; }

        public bool IsExistingPackage { get; set; }

        public string DisplayImagePath { get; set; }

        public string ExistingZipPath { get; set; }

        public float Width { get; set; }

        public float Height { get; set; }

        public OfficeShapeState ShapeState { get; set; }

        public void AddTemporaryFile(string path)
        {
            if (!string.IsNullOrEmpty(path))
            {
                temporaryFiles.Add(path);
            }
        }

        public void Dispose()
        {
            foreach (string path in temporaryFiles)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }

            temporaryFiles.Clear();
        }
    }

    internal sealed class FloatingShapeLayout
    {
        private OfficeShapeState shapeState;
        private Word.WdRelativeHorizontalPosition relativeHorizontalPosition;
        private Word.WdRelativeVerticalPosition relativeVerticalPosition;
        private Word.WdWrapType wrapType;
        private Word.WdWrapSideType wrapSide;
        private float distanceBottom;
        private float distanceLeft;
        private float distanceRight;
        private float distanceTop;
        private int allowOverlap;
        private int layoutInCell;
        private int lockAnchor;

        public OfficeShapeState ShapeState
        {
            get { return shapeState; }
        }

        public bool RequiresRenderedDisplay
        {
            get
            {
                return Math.Abs(shapeState.Rotation) > 0.01f
                    || shapeState.HorizontalFlip == (int)Office.MsoTriState.msoTrue
                    || shapeState.VerticalFlip == (int)Office.MsoTriState.msoTrue;
            }
        }

        public void PrepareForRenderedDisplay()
        {
            RenderedImageService.PrepareStateForRenderedDisplay(shapeState);
        }

        public static FloatingShapeLayout Capture(Word.Shape shape)
        {
            return new FloatingShapeLayout
            {
                shapeState = WordShapeStateAdapter.Capture(shape),
                relativeHorizontalPosition = shape.RelativeHorizontalPosition,
                relativeVerticalPosition = shape.RelativeVerticalPosition,
                wrapType = shape.WrapFormat.Type,
                wrapSide = shape.WrapFormat.Side,
                distanceBottom = shape.WrapFormat.DistanceBottom,
                distanceLeft = shape.WrapFormat.DistanceLeft,
                distanceRight = shape.WrapFormat.DistanceRight,
                distanceTop = shape.WrapFormat.DistanceTop,
                allowOverlap = shape.WrapFormat.AllowOverlap,
                layoutInCell = shape.LayoutInCell,
                lockAnchor = shape.LockAnchor
            };
        }

        public void Apply(Word.Shape shape, bool includeIdentity)
        {
            WordShapeStateAdapter.Apply(shape, shapeState, includeIdentity);
            shape.RelativeHorizontalPosition = relativeHorizontalPosition;
            shape.RelativeVerticalPosition = relativeVerticalPosition;
            shape.WrapFormat.Type = wrapType;
            shape.WrapFormat.Side = wrapSide;
            shape.WrapFormat.DistanceBottom = distanceBottom;
            shape.WrapFormat.DistanceLeft = distanceLeft;
            shape.WrapFormat.DistanceRight = distanceRight;
            shape.WrapFormat.DistanceTop = distanceTop;
            shape.WrapFormat.AllowOverlap = allowOverlap;
            shape.LayoutInCell = layoutInCell;
            shape.LockAnchor = lockAnchor;
            WordShapeStateAdapter.Apply(shape, shapeState, includeIdentity);
        }

        public void PlaceImmediatelyAbove(Word.Shape replacement, Word.Shape original)
        {
            WordShapeStateAdapter.MoveToZOrder(replacement, shapeState.ZOrderPosition + 1);
        }

        public void FinishReplacement(Word.Shape replacement)
        {
            WordShapeStateAdapter.ApplyIdentity(replacement, shapeState);
        }
    }
}
