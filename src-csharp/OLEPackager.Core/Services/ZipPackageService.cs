using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OLEPackager.Core
{
    public sealed class ZipPackageService
    {
        private static readonly Encoding Utf8ForLegacyPackages = new UTF8Encoding(false);

        public void CreateFromDirectory(string sourceDirectory, string zipPath, bool excludeGeneratedPlotFiles)
        {
            if (!Directory.Exists(sourceDirectory))
            {
                throw new DirectoryNotFoundException(sourceDirectory);
            }

            string fullSource = Path.GetFullPath(sourceDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string root = fullSource + Path.DirectorySeparatorChar;
            PrepareDestination(zipPath);

            using (FileStream stream = new FileStream(zipPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false))
            {
                foreach (string filePath in Directory.GetFiles(fullSource, "*", SearchOption.AllDirectories))
                {
                    string parent = Path.GetDirectoryName(Path.GetFullPath(filePath)).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    if (excludeGeneratedPlotFiles
                        && string.Equals(parent, fullSource, StringComparison.OrdinalIgnoreCase)
                        && IsGeneratedPlotFile(Path.GetFileName(filePath)))
                    {
                        continue;
                    }

                    string entryName = ToEntryName(Path.GetFullPath(filePath).Substring(root.Length));
                    archive.CreateEntryFromFile(filePath, entryName, CompressionLevel.Optimal);
                }

                foreach (string directoryPath in Directory.GetDirectories(fullSource, "*", SearchOption.AllDirectories))
                {
                    if (Directory.GetFileSystemEntries(directoryPath).Length != 0)
                    {
                        continue;
                    }

                    string entryName = ToEntryName(Path.GetFullPath(directoryPath).Substring(root.Length));
                    archive.CreateEntry(entryName.TrimEnd('/') + "/");
                }
            }
        }

        public void ExtractToDirectory(string zipPath, string targetDirectory)
        {
            if (!File.Exists(zipPath))
            {
                throw new FileNotFoundException("嵌入的 zip 文件不存在。", zipPath);
            }

            if (Directory.Exists(targetDirectory))
            {
                Directory.Delete(targetDirectory, true);
            }

            Directory.CreateDirectory(targetDirectory);
            string root = Path.GetFullPath(targetDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string rootPrefix = root + Path.DirectorySeparatorChar;

            using (FileStream stream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Utf8ForLegacyPackages))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string entryName = entry.FullName.Replace('\\', '/');
                    if (string.IsNullOrWhiteSpace(entryName))
                    {
                        continue;
                    }

                    if (entryName.StartsWith("/", StringComparison.Ordinal) || entryName.Contains(":"))
                    {
                        throw new InvalidDataException("zip 中包含不安全的路径：" + entry.FullName);
                    }

                    string relativePath = entryName.Replace('/', Path.DirectorySeparatorChar);
                    string targetPath = Path.GetFullPath(Path.Combine(root, relativePath));
                    if (!targetPath.Equals(root, StringComparison.OrdinalIgnoreCase)
                        && !targetPath.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException("zip 中包含越界路径：" + entry.FullName);
                    }

                    if (entryName.EndsWith("/", StringComparison.Ordinal))
                    {
                        Directory.CreateDirectory(targetPath);
                        continue;
                    }

                    string parent = Path.GetDirectoryName(targetPath);
                    if (!string.IsNullOrEmpty(parent))
                    {
                        Directory.CreateDirectory(parent);
                    }

                    entry.ExtractToFile(targetPath, true);
                }
            }
        }

        private static void PrepareDestination(string zipPath)
        {
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }

            string directory = Path.GetDirectoryName(Path.GetFullPath(zipPath));
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        private static bool IsGeneratedPlotFile(string fileName)
        {
            return string.Equals(fileName, "plot.png", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fileName, "plot.svg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fileName, "plot.pdf", StringComparison.OrdinalIgnoreCase);
        }

        private static string ToEntryName(string relativePath)
        {
            return relativePath.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
        }
    }
}
