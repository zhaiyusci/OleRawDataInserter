using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace OLEPackager.Core
{
    public static class RenderedImageService
    {
        private const int MaximumDimension = 4096;

        public static string SaveCroppedMetafilePng(byte[] metafileBytes, string baseName)
        {
            if (metafileBytes == null || metafileBytes.Length == 0)
            {
                throw new ArgumentException("Rendered metafile data is empty.", "metafileBytes");
            }

            string directory = Path.Combine(Path.GetTempPath(), "OLE Packager");
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, baseName + "_" + Guid.NewGuid().ToString("N") + ".png");

            using (MemoryStream stream = new MemoryStream(metafileBytes, false))
            using (Image metafile = Image.FromStream(stream, true, true))
            {
                float scale = Math.Min(1f, (float)MaximumDimension / Math.Max(metafile.Width, metafile.Height));
                int width = Math.Max(1, (int)Math.Ceiling(metafile.Width * scale));
                int height = Math.Max(1, (int)Math.Ceiling(metafile.Height * scale));

                using (Bitmap canvas = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (Graphics graphics = Graphics.FromImage(canvas))
                    {
                        graphics.Clear(Color.Transparent);
                        graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        graphics.DrawImage(metafile, new Rectangle(0, 0, width, height));
                    }

                    Rectangle contentBounds = FindVisibleBounds(canvas);
                    if (contentBounds.Width <= 0 || contentBounds.Height <= 0)
                    {
                        contentBounds = new Rectangle(0, 0, width, height);
                    }

                    contentBounds.Inflate(2, 2);
                    contentBounds.Intersect(new Rectangle(0, 0, width, height));
                    using (Bitmap cropped = canvas.Clone(contentBounds, PixelFormat.Format32bppArgb))
                    {
                        cropped.Save(path, ImageFormat.Png);
                    }
                }
            }

            return path;
        }

        public static void PrepareStateForRenderedDisplay(OfficeShapeState state)
        {
            if (state == null)
            {
                return;
            }

            double radians = state.Rotation * Math.PI / 180d;
            double cosine = Math.Abs(Math.Cos(radians));
            double sine = Math.Abs(Math.Sin(radians));
            float renderedWidth = (float)((state.Width * cosine) + (state.Height * sine));
            float renderedHeight = (float)((state.Width * sine) + (state.Height * cosine));
            float centerX = state.Left + (state.Width / 2f);
            float centerY = state.Top + (state.Height / 2f);

            state.Left = centerX - (renderedWidth / 2f);
            state.Top = centerY - (renderedHeight / 2f);
            state.Width = renderedWidth;
            state.Height = renderedHeight;
            state.Rotation = 0f;
            state.HorizontalFlip = 0;
            state.VerticalFlip = 0;
            state.Picture = null;
            state.Line = null;
            state.Shadow = null;
        }

        private static Rectangle FindVisibleBounds(Bitmap bitmap)
        {
            Rectangle area = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            BitmapData data = bitmap.LockBits(area, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                int stride = Math.Abs(data.Stride);
                byte[] pixels = new byte[stride * bitmap.Height];
                Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);

                int minimumX = bitmap.Width;
                int minimumY = bitmap.Height;
                int maximumX = -1;
                int maximumY = -1;
                for (int y = 0; y < bitmap.Height; y++)
                {
                    int row = y * stride;
                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        if (pixels[row + (x * 4) + 3] == 0)
                        {
                            continue;
                        }

                        minimumX = Math.Min(minimumX, x);
                        minimumY = Math.Min(minimumY, y);
                        maximumX = Math.Max(maximumX, x);
                        maximumY = Math.Max(maximumY, y);
                    }
                }

                return maximumX < minimumX || maximumY < minimumY
                    ? Rectangle.Empty
                    : Rectangle.FromLTRB(minimumX, minimumY, maximumX + 1, maximumY + 1);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }
    }
}
