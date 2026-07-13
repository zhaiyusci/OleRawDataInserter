using System;
using System.Drawing;
using System.IO;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OLEPackager.WordAddIn
{
    internal sealed class WordOlePackageService
    {
        private const float DefaultWidthPoints = 360f;
        private const float DefaultHeightPoints = 240f;
        private const float DefaultDpi = 96f;

        private readonly Word.Application application;

        public WordOlePackageService(Word.Application application)
        {
            this.application = application;
        }

        public Word.InlineShape InsertInlinePackage(string imagePath, string zipPath)
        {
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException("显示图片不存在。", imagePath);
            }

            if (!File.Exists(zipPath))
            {
                throw new FileNotFoundException("待嵌入的 zip 不存在。", zipPath);
            }

            SizeF displaySize = FitToCurrentTextArea(ReadImagePrintSize(imagePath));
            return InsertInlinePackage(imagePath, zipPath, displaySize.Width, displaySize.Height);
        }

        public Word.InlineShape InsertInlinePackage(
            string imagePath,
            string zipPath,
            float widthPoints,
            float heightPoints)
        {
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException("显示图片不存在。", imagePath);
            }

            if (!File.Exists(zipPath))
            {
                throw new FileNotFoundException("待嵌入的 zip 不存在。", zipPath);
            }

            string iconPath = TransparentIconService.EnsureCreated();
            Word.InlineShape oleObject = application.Selection.InlineShapes.AddOLEObject(
                Type.Missing,
                zipPath,
                false,
                true,
                iconPath,
                0,
                string.Empty,
                Type.Missing);

            try
            {
                ApplyDisplayImage(oleObject, imagePath);
                oleObject.Width = widthPoints;
                oleObject.Height = heightPoints;
                return oleObject;
            }
            catch
            {
                try
                {
                    oleObject.Delete();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                }

                throw;
            }
        }

        public void ApplyDisplayImage(Word.Shape oleObject, string imagePath)
        {
            try
            {
                oleObject.Line.Visible = Office.MsoTriState.msoFalse;
                oleObject.Fill.Visible = Office.MsoTriState.msoTrue;
                oleObject.Fill.UserPicture(imagePath);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw new InvalidOperationException("Word 无法把所选图片应用为 OLE 的显示图片。");
            }
        }

        private static void ApplyDisplayImage(Word.InlineShape oleObject, string imagePath)
        {
            try
            {
                oleObject.Borders.Enable = 0;
                oleObject.Line.Visible = Office.MsoTriState.msoFalse;
                oleObject.Fill.Visible = Office.MsoTriState.msoTrue;
                oleObject.Fill.UserPicture(imagePath);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw new InvalidOperationException("Word 无法把所选图片应用为 OLE 的显示图片。");
            }
        }

        private SizeF FitToCurrentTextArea(SizeF imageSize)
        {
            try
            {
                Word.PageSetup setup = application.Selection.Sections[1].PageSetup;
                float availableWidth = setup.PageWidth - setup.LeftMargin - setup.RightMargin - setup.Gutter;
                float availableHeight = setup.PageHeight - setup.TopMargin - setup.BottomMargin;

                if (availableWidth <= 0f)
                {
                    availableWidth = setup.PageWidth - setup.LeftMargin - setup.RightMargin;
                }

                if (availableHeight <= 0f)
                {
                    availableHeight = setup.PageHeight - setup.TopMargin - setup.BottomMargin;
                }

                if (availableWidth <= 0f || availableHeight <= 0f || imageSize.Width <= 0f || imageSize.Height <= 0f)
                {
                    return imageSize;
                }

                float scale = Math.Min(1f, Math.Min(availableWidth / imageSize.Width, availableHeight / imageSize.Height));
                return new SizeF(imageSize.Width * scale, imageSize.Height * scale);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return imageSize;
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

    }
}
