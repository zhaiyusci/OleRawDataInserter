using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OLEPackager.WordAddIn
{
    internal sealed class FlatOpcPackageService
    {
        private static readonly XNamespace PackageNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage";

        public string ExtractDisplayImage(string openXml, string baseName)
        {
            XElement imagePart = ParseParts(openXml)
                .FirstOrDefault(
                    part => GetAttribute(part, "contentType").StartsWith("image/", StringComparison.OrdinalIgnoreCase));

            if (imagePart == null)
            {
                throw new InvalidDataException("无法从所选对象中找到显示图片。");
            }

            byte[] bytes = ReadBinaryPart(imagePart, "显示图片");
            string extension = GetImageExtension(
                GetAttribute(imagePart, "name"),
                GetAttribute(imagePart, "contentType"));
            string outputPath = BuildTemporaryPath(baseName, extension);
            File.WriteAllBytes(outputPath, bytes);
            return outputPath;
        }

        public byte[] ExtractOleBytes(string openXml)
        {
            XElement olePart = ParseParts(openXml)
                .FirstOrDefault(
                    part => string.Equals(
                        GetAttribute(part, "contentType"),
                        "application/vnd.openxmlformats-officedocument.oleObject",
                        StringComparison.OrdinalIgnoreCase));

            if (olePart == null)
            {
                throw new InvalidDataException("所选对象不是 OLE Packager 嵌入对象。");
            }

            return ReadBinaryPart(olePart, "OLE 对象");
        }

        private static XElement[] ParseParts(string openXml)
        {
            if (string.IsNullOrWhiteSpace(openXml))
            {
                throw new InvalidDataException("Word 没有返回所选对象的 Open XML。");
            }

            try
            {
                XDocument document = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
                return document.Descendants(PackageNamespace + "part").ToArray();
            }
            catch (Exception exception)
            {
                throw new InvalidDataException("无法读取所选对象的 Open XML。", exception);
            }
        }

        private static byte[] ReadBinaryPart(XElement part, string description)
        {
            XElement binaryData = part.Element(PackageNamespace + "binaryData")
                ?? part.Elements().FirstOrDefault(element => element.Name.LocalName == "binaryData");
            if (binaryData == null || string.IsNullOrWhiteSpace(binaryData.Value))
            {
                throw new InvalidDataException("无法读取" + description + "的二进制内容。");
            }

            try
            {
                return Convert.FromBase64String(binaryData.Value);
            }
            catch (FormatException exception)
            {
                throw new InvalidDataException(description + "的 Base64 内容无效。", exception);
            }
        }

        private static string GetAttribute(XElement element, string localName)
        {
            XAttribute attribute = element.Attribute(PackageNamespace + localName)
                ?? element.Attributes().FirstOrDefault(item => item.Name.LocalName == localName);
            return attribute == null ? string.Empty : attribute.Value;
        }

        private static string GetImageExtension(string partName, string contentType)
        {
            string extension = Path.GetExtension(partName).TrimStart('.');
            if (!string.IsNullOrEmpty(extension))
            {
                return "." + extension;
            }

            switch (contentType.ToLowerInvariant())
            {
                case "image/jpeg":
                case "image/jpg":
                    return ".jpg";
                case "image/tiff":
                    return ".tif";
                case "image/x-emf":
                    return ".emf";
                case "image/x-wmf":
                    return ".wmf";
                case "image/bmp":
                    return ".bmp";
                case "image/gif":
                    return ".gif";
                default:
                    return ".png";
            }
        }

        private static string BuildTemporaryPath(string baseName, string extension)
        {
            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(invalidCharacter, '_');
            }

            string directory = Path.Combine(Path.GetTempPath(), "OLE Packager");
            Directory.CreateDirectory(directory);
            return Path.Combine(
                directory,
                string.Format("{0}_{1:N}{2}", baseName, Guid.NewGuid(), extension));
        }
    }
}
