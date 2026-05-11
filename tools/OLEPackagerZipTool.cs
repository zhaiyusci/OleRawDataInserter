using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

internal static class OLEPackagerZipTool
{
    private static readonly Encoding Utf8 = new UTF8Encoding(false);
    private static readonly HashSet<string> GeneratedPlotFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "plot.png",
        "plot.svg",
        "plot.pdf"
    };

    private static int Main(string[] args)
    {
        try
        {
            if (args.Length == 0)
            {
                Usage();
                return 2;
            }

            string command = args[0].ToLowerInvariant();
            if (command == "create" && args.Length == 4)
            {
                CreateZip(args[1], args[2], ParseBoolean(args[3]));
                return 0;
            }

            if (command == "list" && args.Length == 3)
            {
                ListZip(args[1], args[2]);
                return 0;
            }

            if (command == "extract" && args.Length == 3)
            {
                ExtractZip(args[1], args[2]);
                return 0;
            }

            Usage();
            return 2;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }

    private static bool ParseBoolean(string value)
    {
        return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || value == "1"
            || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static void CreateZip(string sourcePath, string zipPath, bool excludeGeneratedPlotFiles)
    {
        if (File.Exists(zipPath))
        {
            File.Delete(zipPath);
        }

        string zipDirectory = Path.GetDirectoryName(Path.GetFullPath(zipPath));
        if (!string.IsNullOrEmpty(zipDirectory))
        {
            Directory.CreateDirectory(zipDirectory);
        }

        using (FileStream stream = new FileStream(zipPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false, Utf8))
        {
            if (Directory.Exists(sourcePath))
            {
                string root = EnsureTrailingSeparator(Path.GetFullPath(sourcePath));
                foreach (string filePath in Directory.GetFiles(sourcePath, "*", SearchOption.AllDirectories))
                {
                    string parent = Path.GetDirectoryName(Path.GetFullPath(filePath)).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    string sourceRoot = Path.GetFullPath(sourcePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    if (excludeGeneratedPlotFiles
                        && string.Equals(parent, sourceRoot, StringComparison.OrdinalIgnoreCase)
                        && GeneratedPlotFiles.Contains(Path.GetFileName(filePath)))
                    {
                        continue;
                    }

                    string entryName = ToZipEntryName(GetRelativePath(root, Path.GetFullPath(filePath)));
                    archive.CreateEntryFromFile(filePath, entryName, CompressionLevel.Optimal);
                }

                foreach (string directoryPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
                {
                    if (Directory.GetFileSystemEntries(directoryPath).Length == 0)
                    {
                        string entryName = ToZipEntryName(GetRelativePath(root, Path.GetFullPath(directoryPath)));
                        if (!entryName.EndsWith("/", StringComparison.Ordinal))
                        {
                            entryName += "/";
                        }
                        archive.CreateEntry(entryName);
                    }
                }
            }
            else if (File.Exists(sourcePath))
            {
                archive.CreateEntryFromFile(sourcePath, Path.GetFileName(sourcePath), CompressionLevel.Optimal);
            }
            else
            {
                throw new FileNotFoundException("Source file or folder does not exist.", sourcePath);
            }
        }
    }

    private static void ListZip(string zipPath, string listPath)
    {
        List<string> entries = new List<string>();
        using (FileStream stream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Utf8))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                entries.Add(entry.FullName);
            }
        }

        string listDirectory = Path.GetDirectoryName(Path.GetFullPath(listPath));
        if (!string.IsNullOrEmpty(listDirectory))
        {
            Directory.CreateDirectory(listDirectory);
        }
        File.WriteAllLines(listPath, entries.ToArray(), Utf8);
    }

    private static void ExtractZip(string zipPath, string targetFolder)
    {
        Directory.CreateDirectory(targetFolder);
        string root = Path.GetFullPath(targetFolder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string rootPrefix = root + Path.DirectorySeparatorChar;

        using (FileStream stream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Utf8))
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
                    throw new InvalidDataException("Unsafe zip entry name: " + entry.FullName);
                }

                string relativePath = entryName.Replace('/', Path.DirectorySeparatorChar);
                string targetPath = Path.GetFullPath(Path.Combine(root, relativePath));
                if (!targetPath.Equals(root, StringComparison.OrdinalIgnoreCase)
                    && !targetPath.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("Unsafe zip entry path: " + entry.FullName);
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

    private static string EnsureTrailingSeparator(string path)
    {
        if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            || path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal))
        {
            return path;
        }
        return path + Path.DirectorySeparatorChar;
    }

    private static string GetRelativePath(string root, string fullPath)
    {
        if (!fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Path is outside root: " + fullPath);
        }
        return fullPath.Substring(root.Length);
    }

    private static string ToZipEntryName(string relativePath)
    {
        return relativePath.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
    }

    private static void Usage()
    {
        Console.Error.WriteLine("Usage:");
        Console.Error.WriteLine("  OLEPackagerZipTool create <sourcePath> <zipPath> <excludeGeneratedPlotFiles>");
        Console.Error.WriteLine("  OLEPackagerZipTool list <zipPath> <listPath>");
        Console.Error.WriteLine("  OLEPackagerZipTool extract <zipPath> <targetFolder>");
    }
}
