using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OLEPackager.PowerPointAddIn
{
    internal sealed class PptxPackageService
    {
        private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace OfficeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

        private readonly OlePackageExtractor olePackages;

        public PptxPackageService(OlePackageExtractor olePackages)
        {
            this.olePackages = olePackages;
        }

        public string ExtractZip(PowerPoint.Presentation presentation, PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            if (presentation == null || slide == null || shape == null)
            {
                throw new ArgumentNullException("presentation");
            }

            string temporaryDirectory = Path.Combine(Path.GetTempPath(), "OLE Packager");
            Directory.CreateDirectory(temporaryDirectory);
            string pptxPath = Path.Combine(temporaryDirectory, "OLEPackager_" + Guid.NewGuid().ToString("N") + ".pptx");

            try
            {
                presentation.SaveCopyAs(
                    pptxPath,
                    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    Office.MsoTriState.msoFalse);
                byte[] oleBytes = ExtractOleBytes(pptxPath, slide.SlideIndex, shape.Id);
                return olePackages.ExtractZip(oleBytes, "selected_powerpoint_ole");
            }
            finally
            {
                TryDeleteFile(pptxPath);
            }
        }

        public string ExtractDisplayImage(PowerPoint.Presentation presentation, PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            if (presentation == null || slide == null || shape == null)
            {
                throw new ArgumentNullException("presentation");
            }

            string temporaryDirectory = Path.Combine(Path.GetTempPath(), "OLE Packager");
            Directory.CreateDirectory(temporaryDirectory);
            string pptxPath = Path.Combine(temporaryDirectory, "OLEPackager_" + Guid.NewGuid().ToString("N") + ".pptx");
            string imagePath = null;

            try
            {
                presentation.SaveCopyAs(
                    pptxPath,
                    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    Office.MsoTriState.msoFalse);

                using (FileStream stream = File.OpenRead(pptxPath))
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false))
                {
                    ShapePartInfo shapeInfo = FindShapePart(archive, slide.SlideIndex, shape.Id);
                    XElement blip = shapeInfo.ShapeContainer
                        .Descendants(PresentationNamespace + "spPr")
                        .Descendants(DrawingNamespace + "blip")
                        .LastOrDefault(HasImageRelationship)
                        ?? shapeInfo.ShapeContainer
                            .Descendants(DrawingNamespace + "blip")
                            .LastOrDefault(HasImageRelationship);
                    if (blip == null)
                    {
                        throw new InvalidDataException("所选对象没有可提取的原始显示图片。");
                    }

                    string relationshipId = (string)blip.Attribute(OfficeRelationshipNamespace + "embed")
                        ?? (string)blip.Attribute(OfficeRelationshipNamespace + "link");
                    XElement relationship = FindRelationship(archive, shapeInfo.SlidePart, relationshipId);
                    string target = (string)relationship.Attribute("Target");
                    string targetMode = (string)relationship.Attribute("TargetMode");
                    string extension = Path.GetExtension(target);
                    if (string.IsNullOrEmpty(extension) || extension.Length > 10)
                    {
                        extension = ".png";
                    }

                    imagePath = Path.Combine(
                        temporaryDirectory,
                        "selected_powerpoint_image_" + Guid.NewGuid().ToString("N") + extension);

                    if (string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase))
                    {
                        Uri externalUri;
                        string sourcePath = Uri.TryCreate(target, UriKind.Absolute, out externalUri) && externalUri.IsFile
                            ? externalUri.LocalPath
                            : target;
                        if (!File.Exists(sourcePath))
                        {
                            throw new FileNotFoundException("链接图片的源文件不可用。", sourcePath);
                        }

                        File.Copy(sourcePath, imagePath, true);
                    }
                    else
                    {
                        string imagePart = ResolvePartTarget(shapeInfo.SlidePart, target);
                        ZipArchiveEntry imageEntry = GetEntry(archive, imagePart);
                        using (Stream imageStream = imageEntry.Open())
                        using (FileStream output = File.Create(imagePath))
                        {
                            imageStream.CopyTo(output);
                        }
                    }
                }

                return imagePath;
            }
            catch
            {
                TryDeleteFile(imagePath);
                throw;
            }
            finally
            {
                TryDeleteFile(pptxPath);
            }
        }

        internal static byte[] ExtractOleBytes(string pptxPath, int slideIndex, int shapeId)
        {
            using (FileStream stream = File.OpenRead(pptxPath))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false))
            {
                ShapePartInfo shapeInfo = FindShapePart(archive, slideIndex, shapeId);
                XElement shapeContainer = shapeInfo.NonVisualProperties
                    .Ancestors()
                    .FirstOrDefault(element => element.Descendants(PresentationNamespace + "oleObj").Any());
                XElement oleObject = shapeContainer == null
                    ? null
                    : shapeContainer.Descendants(PresentationNamespace + "oleObj").FirstOrDefault();
                string oleRelationshipId = oleObject == null
                    ? null
                    : (string)oleObject.Attribute(OfficeRelationshipNamespace + "id");
                if (string.IsNullOrEmpty(oleRelationshipId))
                {
                    throw new InvalidDataException("所选 PowerPoint 对象没有嵌入的 OLE 数据。");
                }

                string olePart = ResolveRelationshipTarget(archive, shapeInfo.SlidePart, oleRelationshipId);
                ZipArchiveEntry oleEntry = GetEntry(archive, olePart);
                using (Stream oleStream = oleEntry.Open())
                using (MemoryStream bytes = new MemoryStream())
                {
                    oleStream.CopyTo(bytes);
                    return bytes.ToArray();
                }
            }
        }

        private static ShapePartInfo FindShapePart(ZipArchive archive, int slideIndex, int shapeId)
        {
            const string presentationPart = "ppt/presentation.xml";
            XDocument presentation = LoadXml(archive, presentationPart);
            XElement slideId = presentation
                .Descendants(PresentationNamespace + "sldId")
                .ElementAtOrDefault(slideIndex - 1);
            if (slideId == null)
            {
                throw new InvalidDataException("在 PPTX 中找不到当前幻灯片。");
            }

            string slideRelationshipId = (string)slideId.Attribute(OfficeRelationshipNamespace + "id");
            string slidePart = ResolveRelationshipTarget(archive, presentationPart, slideRelationshipId);
            XDocument slide = LoadXml(archive, slidePart);
            XElement nonVisualProperties = slide
                .Descendants(PresentationNamespace + "cNvPr")
                .FirstOrDefault(element => string.Equals(
                    (string)element.Attribute("id"),
                    shapeId.ToString(CultureInfo.InvariantCulture),
                    StringComparison.Ordinal));
            if (nonVisualProperties == null)
            {
                throw new InvalidDataException("在 PPTX 中找不到所选对象。");
            }

            XElement shapeContainer = nonVisualProperties
                .Ancestors()
                .FirstOrDefault(element => element.Descendants(DrawingNamespace + "blip").Any())
                ?? nonVisualProperties.Parent;
            return new ShapePartInfo
            {
                SlidePart = slidePart,
                NonVisualProperties = nonVisualProperties,
                ShapeContainer = shapeContainer
            };
        }

        private static string ResolveRelationshipTarget(ZipArchive archive, string sourcePart, string relationshipId)
        {
            XElement relationship = FindRelationship(archive, sourcePart, relationshipId);
            string target = (string)relationship.Attribute("Target");
            return ResolvePartTarget(sourcePart, target);
        }

        private static XElement FindRelationship(ZipArchive archive, string sourcePart, string relationshipId)
        {
            string relationshipPart = GetRelationshipPartPath(sourcePart);
            XDocument relationships = LoadXml(archive, relationshipPart);
            XElement relationship = relationships
                .Descendants(PackageRelationshipNamespace + "Relationship")
                .FirstOrDefault(element => string.Equals((string)element.Attribute("Id"), relationshipId, StringComparison.Ordinal));
            if (relationship == null)
            {
                throw new InvalidDataException("PPTX 关系不存在：" + relationshipId);
            }

            if (string.IsNullOrEmpty((string)relationship.Attribute("Target")))
            {
                throw new InvalidDataException("PPTX 关系没有目标：" + relationshipId);
            }

            return relationship;
        }

        private static string ResolvePartTarget(string sourcePart, string target)
        {
            Uri sourceUri = new Uri("http://package/" + sourcePart.Replace('\\', '/'));
            Uri targetUri = new Uri(sourceUri, target);
            return Uri.UnescapeDataString(targetUri.AbsolutePath.TrimStart('/'));
        }

        private static bool HasImageRelationship(XElement blip)
        {
            return blip.Attribute(OfficeRelationshipNamespace + "embed") != null
                || blip.Attribute(OfficeRelationshipNamespace + "link") != null;
        }

        private static string GetRelationshipPartPath(string sourcePart)
        {
            string normalized = sourcePart.Replace('\\', '/');
            int separator = normalized.LastIndexOf('/');
            string directory = separator < 0 ? string.Empty : normalized.Substring(0, separator + 1);
            string fileName = separator < 0 ? normalized : normalized.Substring(separator + 1);
            return directory + "_rels/" + fileName + ".rels";
        }

        private static XDocument LoadXml(ZipArchive archive, string partPath)
        {
            ZipArchiveEntry entry = GetEntry(archive, partPath);
            using (Stream stream = entry.Open())
            {
                return XDocument.Load(stream, LoadOptions.None);
            }
        }

        private static ZipArchiveEntry GetEntry(ZipArchive archive, string partPath)
        {
            string normalized = partPath.Replace('\\', '/').TrimStart('/');
            ZipArchiveEntry entry = archive.GetEntry(normalized)
                ?? archive.Entries.FirstOrDefault(candidate => string.Equals(candidate.FullName, normalized, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                throw new InvalidDataException("PPTX 部件不存在：" + normalized);
            }

            return entry;
        }

        private static void TryDeleteFile(string path)
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

        private sealed class ShapePartInfo
        {
            public string SlidePart { get; set; }
            public XElement NonVisualProperties { get; set; }
            public XElement ShapeContainer { get; set; }
        }
    }
}
