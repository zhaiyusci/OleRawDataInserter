using System.IO;

namespace OLEPackager.Core
{
    public static class TransparentIconService
    {
        public static string EnsureCreated()
        {
            string directory = Path.Combine(Path.GetTempPath(), "OLE Packager");
            string iconPath = Path.Combine(directory, "OLEPackagerTransparent.ico");
            Directory.CreateDirectory(directory);

            byte[] iconBytes = BuildIconBytes();
            if (!File.Exists(iconPath) || new FileInfo(iconPath).Length != iconBytes.Length)
            {
                File.WriteAllBytes(iconPath, iconBytes);
            }

            return iconPath;
        }

        private static byte[] BuildIconBytes()
        {
            const int iconWidth = 16;
            const int iconHeight = 16;
            const int bitmapHeaderSize = 40;
            const int pixelBytes = iconWidth * iconHeight * 4;
            const int maskBytes = 64;
            const int imageBytes = bitmapHeaderSize + pixelBytes + maskBytes;

            byte[] bytes = new byte[22 + imageBytes];
            int position = 0;

            AppendUInt16(bytes, ref position, 0);
            AppendUInt16(bytes, ref position, 1);
            AppendUInt16(bytes, ref position, 1);
            AppendByte(bytes, ref position, iconWidth);
            AppendByte(bytes, ref position, iconHeight);
            AppendByte(bytes, ref position, 0);
            AppendByte(bytes, ref position, 0);
            AppendUInt16(bytes, ref position, 1);
            AppendUInt16(bytes, ref position, 32);
            AppendUInt32(bytes, ref position, imageBytes);
            AppendUInt32(bytes, ref position, 22);

            AppendUInt32(bytes, ref position, bitmapHeaderSize);
            AppendUInt32(bytes, ref position, iconWidth);
            AppendUInt32(bytes, ref position, iconHeight * 2);
            AppendUInt16(bytes, ref position, 1);
            AppendUInt16(bytes, ref position, 32);
            AppendUInt32(bytes, ref position, 0);
            AppendUInt32(bytes, ref position, pixelBytes);
            AppendUInt32(bytes, ref position, 0);
            AppendUInt32(bytes, ref position, 0);
            AppendUInt32(bytes, ref position, 0);
            AppendUInt32(bytes, ref position, 0);

            position += pixelBytes;
            for (int index = 0; index < maskBytes; index++)
            {
                AppendByte(bytes, ref position, 0xff);
            }

            return bytes;
        }

        private static void AppendByte(byte[] bytes, ref int position, int value)
        {
            bytes[position++] = (byte)(value & 0xff);
        }

        private static void AppendUInt16(byte[] bytes, ref int position, int value)
        {
            AppendByte(bytes, ref position, value);
            AppendByte(bytes, ref position, value >> 8);
        }

        private static void AppendUInt32(byte[] bytes, ref int position, int value)
        {
            AppendByte(bytes, ref position, value);
            AppendByte(bytes, ref position, value >> 8);
            AppendByte(bytes, ref position, value >> 16);
            AppendByte(bytes, ref position, value >> 24);
        }
    }
}
