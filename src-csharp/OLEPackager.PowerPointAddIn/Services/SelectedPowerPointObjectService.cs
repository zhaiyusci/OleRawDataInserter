using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OLEPackager.PowerPointAddIn
{
    internal sealed class SelectedPowerPointObjectService
    {
        private readonly PowerPoint.Application application;
        private readonly PowerPointOlePackageService powerPointPackages;
        private readonly PptxPackageService pptxPackages;

        public SelectedPowerPointObjectService(
            PowerPoint.Application application,
            PowerPointOlePackageService powerPointPackages,
            PptxPackageService pptxPackages)
        {
            this.application = application;
            this.powerPointPackages = powerPointPackages;
            this.pptxPackages = pptxPackages;
        }

        public SelectedPowerPointObject CaptureSelection()
        {
            if (application.Windows.Count == 0)
            {
                return null;
            }

            PowerPoint.Selection selection = application.ActiveWindow.Selection;
            if (selection == null
                || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes
                || selection.ShapeRange.Count != 1)
            {
                return null;
            }

            PowerPoint.Shape shape = selection.ShapeRange[1];
            if (shape.Type == Office.MsoShapeType.msoGroup || IsInsideGroup(shape))
            {
                throw new InvalidOperationException("暂不支持替换组合或组合内部的对象。请先取消组合，再管理附件。");
            }

            if (shape.Type == Office.MsoShapeType.msoPlaceholder)
            {
                throw new InvalidOperationException("暂不替换占位符中的图片。请先把图片另存为普通图片对象。");
            }

            bool isPicture = shape.Type == Office.MsoShapeType.msoPicture
                || shape.Type == Office.MsoShapeType.msoLinkedPicture;
            bool isOle = shape.Type == Office.MsoShapeType.msoEmbeddedOLEObject
                || shape.Type == Office.MsoShapeType.msoLinkedOLEObject;
            if (!isPicture && !isOle)
            {
                return null;
            }

            PowerPoint.Slide slide = GetContainingSlide(shape);
            PowerPointShapeState shapeState = PowerPointShapeStateAdapter.Capture(shape);
            SelectedPowerPointObject result = new SelectedPowerPointObject
            {
                Shape = shape,
                Slide = slide,
                ShapeState = shapeState,
                IsExistingPackage = isOle
            };

            try
            {
                string imagePath;
                if (isPicture && shapeState.RequiresRenderedDisplay)
                {
                    imagePath = ExportRenderedImage(shape);
                    shapeState.PrepareForRenderedDisplay();
                }
                else
                {
                    imagePath = pptxPackages.ExtractDisplayImage(application.ActivePresentation, slide, shape);
                }

                result.DisplayImagePath = imagePath;
                result.AddTemporaryFile(imagePath);

                if (isOle)
                {
                    result.ExistingZipPath = pptxPackages.ExtractZip(application.ActivePresentation, slide, shape);
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

        public void ReplaceWithPackage(SelectedPowerPointObject target, string zipPath)
        {
            if (target == null)
            {
                throw new ArgumentNullException("target");
            }

            PowerPoint.Shape oldShape = target.Shape;
            PowerPoint.Shape newShape = null;
            bool animationsTransferred = false;
            OfficeShapeState common = target.ShapeState.Common;

            try
            {
                newShape = powerPointPackages.InsertPackage(
                    target.Slide,
                    target.DisplayImagePath,
                    zipPath,
                    common.Left,
                    common.Top,
                    common.Width,
                    common.Height);
                PowerPointShapeStateAdapter.Apply(newShape, target.ShapeState, false);
                PowerPointShapeStateAdapter.MoveToZOrder(newShape, oldShape.ZOrderPosition + 1);
                PowerPointShapeStateAdapter.TransferAnimations(target.Slide, oldShape, newShape);
                animationsTransferred = true;

                oldShape.Delete();
                oldShape = null;
                PowerPointShapeStateAdapter.MoveToZOrder(newShape, common.ZOrderPosition);
                PowerPointShapeStateAdapter.ApplyIdentity(newShape, common);
                target.Shape = newShape;
            }
            catch
            {
                if (animationsTransferred && oldShape != null && newShape != null)
                {
                    PowerPointShapeStateAdapter.TransferAnimations(target.Slide, newShape, oldShape);
                }

                if (newShape != null)
                {
                    try { newShape.Delete(); }
                    catch (COMException) { }
                }

                throw;
            }
        }

        private static bool IsInsideGroup(PowerPoint.Shape shape)
        {
            try
            {
                return shape.ParentGroup != null;
            }
            catch (COMException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }

        private static PowerPoint.Slide GetContainingSlide(PowerPoint.Shape shape)
        {
            try
            {
                return (PowerPoint.Slide)shape.Parent;
            }
            catch (InvalidCastException)
            {
                throw new InvalidOperationException("所选对象不在普通幻灯片上。");
            }
        }

        private string ExportRenderedImage(PowerPoint.Shape shape)
        {
            string directory = Path.Combine(Path.GetTempPath(), "OLE Packager");
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "selected_powerpoint_rendered_" + Guid.NewGuid().ToString("N") + ".png");
            int slideWidth = Math.Max(1, (int)Math.Ceiling(application.ActivePresentation.PageSetup.SlideWidth * 2f));
            int slideHeight = Math.Max(1, (int)Math.Ceiling(application.ActivePresentation.PageSetup.SlideHeight * 2f));
            shape.Export(
                path,
                PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                slideWidth,
                slideHeight,
                PowerPoint.PpExportMode.ppRelativeToSlide);
            if (!File.Exists(path))
            {
                throw new InvalidOperationException("PowerPoint 没有导出所选对象的显示图。");
            }

            return path;
        }

    }

    internal sealed class SelectedPowerPointObject : IDisposable
    {
        private readonly List<string> temporaryFiles = new List<string>();

        public PowerPoint.Shape Shape { get; set; }
        public PowerPoint.Slide Slide { get; set; }
        public PowerPointShapeState ShapeState { get; set; }
        public bool IsExistingPackage { get; set; }
        public string DisplayImagePath { get; set; }
        public string ExistingZipPath { get; set; }

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
}
