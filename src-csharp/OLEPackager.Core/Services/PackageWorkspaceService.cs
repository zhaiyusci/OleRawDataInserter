using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace OLEPackager.Core
{
    public sealed class PackageWorkspaceService
    {
        private readonly string rootDirectory;
        private readonly string rootPrefix;

        public PackageWorkspaceService(string rootDirectory)
        {
            if (string.IsNullOrWhiteSpace(rootDirectory))
            {
                throw new ArgumentException("工作文件夹不能为空。", "rootDirectory");
            }

            this.rootDirectory = Path.GetFullPath(rootDirectory)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            rootPrefix = this.rootDirectory + Path.DirectorySeparatorChar;
            Directory.CreateDirectory(this.rootDirectory);
        }

        public string RootDirectory
        {
            get { return rootDirectory; }
        }

        public IList<PackageWorkspaceEntry> GetEntries()
        {
            Directory.CreateDirectory(rootDirectory);
            List<PackageWorkspaceEntry> entries = new List<PackageWorkspaceEntry>();
            Stack<string> pendingDirectories = new Stack<string>();
            pendingDirectories.Push(rootDirectory);

            while (pendingDirectories.Count > 0)
            {
                string currentDirectory = pendingDirectories.Pop();
                foreach (string directoryPath in Directory.GetDirectories(currentDirectory))
                {
                    DirectoryInfo directory = new DirectoryInfo(directoryPath);
                    if ((directory.Attributes & FileAttributes.ReparsePoint) != 0)
                    {
                        continue;
                    }

                    entries.Add(new PackageWorkspaceEntry(
                        directoryPath,
                        GetRelativePath(directoryPath),
                        true,
                        0,
                        directory.LastWriteTime));
                    pendingDirectories.Push(directoryPath);
                }

                foreach (string filePath in Directory.GetFiles(currentDirectory))
                {
                    FileInfo file = new FileInfo(filePath);
                    entries.Add(new PackageWorkspaceEntry(
                        filePath,
                        GetRelativePath(filePath),
                        false,
                        file.Length,
                        file.LastWriteTime));
                }
            }

            return entries
                .OrderBy(entry => entry.RelativePath, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public void AddPaths(IEnumerable<string> sourcePaths)
        {
            if (sourcePaths == null)
            {
                return;
            }

            foreach (string sourcePath in sourcePaths.Where(path => !string.IsNullOrWhiteSpace(path)))
            {
                string fullSourcePath = Path.GetFullPath(sourcePath);
                EnsureSourceIsOutsideWorkspace(fullSourcePath);

                if (File.Exists(fullSourcePath))
                {
                    string destination = GetUniqueDestination(Path.GetFileName(fullSourcePath), false);
                    File.Copy(fullSourcePath, destination, false);
                }
                else if (Directory.Exists(fullSourcePath))
                {
                    string folderName = new DirectoryInfo(fullSourcePath).Name;
                    if (string.IsNullOrEmpty(folderName))
                    {
                        folderName = "folder";
                    }

                    string destination = GetUniqueDestination(folderName, true);
                    CopyDirectory(fullSourcePath, destination);
                }
                else
                {
                    throw new FileNotFoundException("要添加的文件或文件夹不存在。", fullSourcePath);
                }
            }
        }

        public void RemoveEntries(IEnumerable<PackageWorkspaceEntry> entries)
        {
            if (entries == null)
            {
                return;
            }

            foreach (PackageWorkspaceEntry entry in entries
                .Where(item => item != null)
                .OrderByDescending(item => item.FullPath.Length))
            {
                string fullPath = EnsurePathIsInsideWorkspace(entry.FullPath);
                if (File.Exists(fullPath))
                {
                    SetNormalAttributes(fullPath);
                    File.Delete(fullPath);
                }
                else if (Directory.Exists(fullPath))
                {
                    SetNormalAttributesRecursively(fullPath);
                    Directory.Delete(fullPath, true);
                }
            }
        }

        public void Clear()
        {
            Directory.CreateDirectory(rootDirectory);
            foreach (string filePath in Directory.GetFiles(rootDirectory))
            {
                SetNormalAttributes(filePath);
                File.Delete(filePath);
            }

            foreach (string directoryPath in Directory.GetDirectories(rootDirectory))
            {
                SetNormalAttributesRecursively(directoryPath);
                Directory.Delete(directoryPath, true);
            }
        }

        public void OpenRootDirectory()
        {
            Directory.CreateDirectory(rootDirectory);
            Process.Start(new ProcessStartInfo("explorer.exe", "\"" + rootDirectory + "\"")
            {
                UseShellExecute = true
            });
        }

        public void OpenEntry(PackageWorkspaceEntry entry)
        {
            if (entry == null)
            {
                return;
            }

            string fullPath = EnsurePathIsInsideWorkspace(entry.FullPath);
            if (!File.Exists(fullPath) && !Directory.Exists(fullPath))
            {
                throw new FileNotFoundException("所选项目已不存在。", fullPath);
            }

            Process.Start(new ProcessStartInfo(fullPath)
            {
                UseShellExecute = true
            });
        }

        private void CopyDirectory(string sourceDirectory, string destinationDirectory)
        {
            Directory.CreateDirectory(destinationDirectory);
            foreach (string filePath in Directory.GetFiles(sourceDirectory))
            {
                FileInfo sourceFile = new FileInfo(filePath);
                if ((sourceFile.Attributes & FileAttributes.ReparsePoint) != 0)
                {
                    continue;
                }

                File.Copy(filePath, Path.Combine(destinationDirectory, sourceFile.Name), false);
            }

            foreach (string childDirectory in Directory.GetDirectories(sourceDirectory))
            {
                DirectoryInfo sourceChild = new DirectoryInfo(childDirectory);
                if ((sourceChild.Attributes & FileAttributes.ReparsePoint) != 0)
                {
                    continue;
                }

                CopyDirectory(childDirectory, Path.Combine(destinationDirectory, sourceChild.Name));
            }
        }

        private string GetUniqueDestination(string name, bool directory)
        {
            string candidate = Path.Combine(rootDirectory, name);
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
            {
                return candidate;
            }

            string baseName = directory ? name : Path.GetFileNameWithoutExtension(name);
            string extension = directory ? string.Empty : Path.GetExtension(name);
            for (int index = 2; index < int.MaxValue; index++)
            {
                candidate = Path.Combine(rootDirectory, string.Format("{0} ({1}){2}", baseName, index, extension));
                if (!File.Exists(candidate) && !Directory.Exists(candidate))
                {
                    return candidate;
                }
            }

            throw new IOException("无法为添加的项目生成唯一名称。");
        }

        private string GetRelativePath(string fullPath)
        {
            string normalizedPath = EnsurePathIsInsideWorkspace(fullPath);
            return normalizedPath.Substring(rootPrefix.Length);
        }

        private string EnsurePathIsInsideWorkspace(string path)
        {
            string fullPath = Path.GetFullPath(path);
            if (!fullPath.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("路径不在包内容工作文件夹中：" + fullPath);
            }

            return fullPath;
        }

        private void EnsureSourceIsOutsideWorkspace(string sourcePath)
        {
            string sourcePrefix = sourcePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            if (sourcePath.Equals(rootDirectory, StringComparison.OrdinalIgnoreCase)
                || sourcePath.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase)
                || rootDirectory.StartsWith(sourcePrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("不能把工作文件夹本身或其中的项目再次添加到包中。");
            }
        }

        private static void SetNormalAttributesRecursively(string directoryPath)
        {
            foreach (string filePath in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
            {
                SetNormalAttributes(filePath);
            }

            foreach (string childDirectory in Directory.GetDirectories(directoryPath, "*", SearchOption.AllDirectories)
                .OrderByDescending(path => path.Length))
            {
                SetNormalAttributes(childDirectory);
            }

            SetNormalAttributes(directoryPath);
        }

        private static void SetNormalAttributes(string path)
        {
            try
            {
                File.SetAttributes(path, FileAttributes.Normal);
            }
            catch (FileNotFoundException)
            {
            }
            catch (DirectoryNotFoundException)
            {
            }
        }
    }

    public sealed class PackageWorkspaceEntry
    {
        public PackageWorkspaceEntry(
            string fullPath,
            string relativePath,
            bool isDirectory,
            long size,
            DateTime modified)
        {
            FullPath = fullPath;
            RelativePath = relativePath;
            IsDirectory = isDirectory;
            Size = size;
            Modified = modified;
        }

        public string FullPath { get; private set; }

        public string RelativePath { get; private set; }

        public bool IsDirectory { get; private set; }

        public long Size { get; private set; }

        public DateTime Modified { get; private set; }
    }
}
