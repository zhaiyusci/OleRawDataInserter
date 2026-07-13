using System;
using System.Drawing;
using System.IO;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OLEPackager.PowerPointAddIn
{
    internal sealed class PowerPointOlePackageService
    {
        private const float DefaultWidthPoints = 360f;
        private const float DefaultHeightPoints = 240f;
        private const float DefaultDpi = 96f;
        private const float SlideMarginPoints = 36f;

        private readonly PowerPoint.Application application;

        public PowerPointOlePackageService(PowerPoint.Application application)
        {
            this.application = application;
        }

        public PowerPoint.Shape InsertPackage(PowerPoint.Slide slide, string imagePath, string zipPath)
        {
            SizeF size = ReadImagePrintSize(imagePath);
            float slideWidth = application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = application.ActivePresentation.PageSetup.SlideHeight;
            size = FitInside(size, Math.Max(72f, slideWidth - (SlideMarginPoints * 2f)), Math.Max(72f, slideHeight - (SlideMarginPoints * 2f)));
            float left = Math.Max(0f, (slideWidth - size.Width) / 2f);
            float top = Math.Max(0f, (slideHeight - size.Height) / 2f);
            return InsertPackage(slide, imagePath, zipPath, left, top, size.Width, size.Height);
        }

        public PowerPoint.Shape InsertPackage(
            PowerPoint.Slide slide,
            string imagePath,
            string zipPath,
            float left,
            float top,
            float width,
            float height)
        {
            if (slide == null)
            {
                throw new InvalidOperationException("请先打开演示文稿并切换到一张幻灯片。");
            }

            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException("显示图片不存在。", imagePath);
            }

            if (!File.Exists(zipPath))
            {
                throw new FileNotFoundException("待嵌入的 zip 不存在。", zipPath);
            }

            PowerPoint.Shape shape = slide.Shapes.AddOLEObject(
                left,
                top,
                width,
                height,
                string.Empty,
                zipPath,
                Office.MsoTriState.msoTrue,
                TransparentIconService.EnsureCreated(),
                0,
                string.Empty,
                Office.MsoTriState.msoFalse);

            try
            {
                ApplyDisplayImage(shape, imagePath);
                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                shape.Width = width;
                shape.Height = height;
                shape.Left = left;
                shape.Top = top;
                return shape;
            }
            catch
            {
                try { shape.Delete(); }
                catch (System.Runtime.InteropServices.COMException) { }
                throw;
            }
        }

        public void ApplyDisplayImage(PowerPoint.Shape shape, string imagePath)
        {
            try
            {
                shape.Line.Visible = Office.MsoTriState.msoFalse;
                shape.Fill.Visible = Office.MsoTriState.msoTrue;
                shape.Fill.UserPicture(imagePath);
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                throw new InvalidOperationException("PowerPoint 无法把所选图片应用为 OLE 的显示图。", exception);
            }
        }

        private static SizeF ReadImagePrintSize(string imagePath)
        {
            try
            {
                using (Image image = Image.FromFile(imagePath))
                {
                    float dpiX = image.HorizontalResolution > 0f ? image.HorizontalResolution : DefaultDpi;
                    float dpiY = image.VerticalResolution > 0f ? image.VerticalResolution : DefaultDpi;
                    return new SizeF(image.Width / dpiX * 72f, image.Height / dpiY * 72f);
                }
            }
            catch (ArgumentException)
            {
                return new SizeF(DefaultWidthPoints, DefaultHeightPoints);
            }
            catch (OutOfMemoryException)
            {
                return new SizeF(DefaultWidthPoints, DefaultHeightPoints);
            }
        }

        private static SizeF FitInside(SizeF source, float maximumWidth, float maximumHeight)
        {
            if (source.Width <= 0f || source.Height <= 0f)
            {
                return new SizeF(DefaultWidthPoints, DefaultHeightPoints);
            }

            float scale = Math.Min(1f, Math.Min(maximumWidth / source.Width, maximumHeight / source.Height));
            return new SizeF(source.Width * scale, source.Height * scale);
        }
    }
}
