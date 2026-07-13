using System;
using System.IO;
using System.Windows.Forms;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OLEPackager.WordAddIn
{
    internal sealed class OlePackagerService
    {
        private const string ProductName = "OLE Packager";

        private readonly Word.Application application;
        private readonly ZipPackageService zipPackages;
        private readonly WordOlePackageService wordPackages;
        private readonly SelectedWordObjectService selectedObjects;

        public OlePackagerService(
            Word.Application application,
            ZipPackageService zipPackages,
            WordOlePackageService wordPackages,
            SelectedWordObjectService selectedObjects)
        {
            this.application = application;
            this.zipPackages = zipPackages;
            this.wordPackages = wordPackages;
            this.selectedObjects = selectedObjects;
        }

        public void InsertPlotFolderAsOle()
        {
            string folderPath = PickPlotFolder();
            if (string.IsNullOrEmpty(folderPath))
            {
                return;
            }

            InsertPlotFolder(folderPath);
        }

        internal void InsertPlotFolder(string folderPath)
        {
            ExecuteWordCommand(
                delegate
                {
                    EnsureDocumentIsOpen();
                    string normalizedFolder = Path.GetFullPath(folderPath.Trim());
                    if (!Directory.Exists(normalizedFolder))
                    {
                        throw new DirectoryNotFoundException("文件夹不存在：" + normalizedFolder);
                    }

                    string plotPath = Path.Combine(normalizedFolder, "plot.png");
                    if (!File.Exists(plotPath))
                    {
                        throw new FileNotFoundException("在所选文件夹中找不到 plot.png。", plotPath);
                    }

                    string zipPath = BuildTemporaryZipPath(normalizedFolder);
                    try
                    {
                        zipPackages.CreateFromDirectory(normalizedFolder, zipPath, true);
                        wordPackages.InsertInlinePackage(plotPath, zipPath);
                    }
                    finally
                    {
                        TryDeleteFile(zipPath);
                    }
                });
        }

        public void InsertImageSupportFilesAsOle()
        {
            if (!HasOpenDocument())
            {
                ShowError("请先新建或打开一个 Word 文档。");
                return;
            }

            string workingFolder = BuildWorkingFolderPath("image_support");
            Directory.CreateDirectory(workingFolder);

            try
            {
                using (PackageEditorDialog dialog = new PackageEditorDialog(
                    PackageEditorMode.InsertNew,
                    workingFolder,
                    null,
                    InsertImageAndWorkingFolder))
                {
                    dialog.ShowDialog(GetWordWindow());
                }
            }
            catch (Exception exception)
            {
                ShowError(exception.Message);
            }
            finally
            {
                TryDeleteDirectory(workingFolder);
            }
        }

        public void AttachSupportFilesToSelectedImageAsOle()
        {
            if (!HasOpenDocument())
            {
                ShowError("请先新建或打开一个 Word 文档。");
                return;
            }

            SelectedWordObject target = null;
            string workingFolder = null;
            try
            {
                target = selectedObjects.CaptureSelection();
                if (target == null)
                {
                    MessageBox.Show(
                        "请先在 Word 文档中选中一张图片或一个 OLE Packager 图包。",
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                workingFolder = BuildWorkingFolderPath(target.IsExistingPackage ? "selected_ole_support" : "selected_image_support");
                if (target.IsExistingPackage)
                {
                    zipPackages.ExtractToDirectory(target.ExistingZipPath, workingFolder);
                }
                else
                {
                    Directory.CreateDirectory(workingFolder);
                }

                SelectedWordObject capturedTarget = target;
                using (PackageEditorDialog dialog = new PackageEditorDialog(
                    target.IsExistingPackage ? PackageEditorMode.ManagePackage : PackageEditorMode.AttachToImage,
                    workingFolder,
                    target.DisplayImagePath,
                    delegate(string imagePath, string folder) { ReplaceSelectedObject(capturedTarget, folder); }))
                {
                    dialog.ShowDialog(GetWordWindow());
                }
            }
            catch (Exception exception)
            {
                ShowError(exception.Message);
            }
            finally
            {
                if (target != null)
                {
                    target.Dispose();
                }

                if (!string.IsNullOrEmpty(workingFolder))
                {
                    TryDeleteDirectory(workingFolder);
                }
            }
        }

        public void ShowUsageHelp()
        {
            using (UsageHelpDialog dialog = new UsageHelpDialog())
            {
                dialog.ShowDialog(GetWordWindow());
            }
        }

        private string PickPlotFolder()
        {
            Office.FileDialog dialog = null;
            try
            {
                dialog = application.FileDialog[Office.MsoFileDialogType.msoFileDialogFolderPicker];
                dialog.Title = "选择包含 plot.png 的文件夹";
                dialog.AllowMultiSelect = false;

                if (dialog.Show() != -1)
                {
                    return null;
                }

                return dialog.SelectedItems.Item(1);
            }
            finally
            {
                if (dialog != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(dialog);
                }
            }
        }

        private void InsertImageAndWorkingFolder(string imagePath, string workingFolder)
        {
            imagePath = Path.GetFullPath(imagePath.Trim());
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException("显示图片不存在。", imagePath);
            }

            if (!IsSupportedDisplayImage(imagePath))
            {
                throw new InvalidOperationException("请选择 PNG、JPEG 或 TIFF 图片。");
            }

            if (!Directory.Exists(workingFolder) || !DirectoryHasFiles(workingFolder))
            {
                throw new InvalidOperationException("包内容文件夹为空，请先添加文件。");
            }

            string zipPath = BuildTemporaryZipPathFromBase(Path.GetFileNameWithoutExtension(imagePath) + "_support");
            try
            {
                zipPackages.CreateFromDirectory(workingFolder, zipPath, false);
                wordPackages.InsertInlinePackage(imagePath, zipPath);
            }
            finally
            {
                TryDeleteFile(zipPath);
            }
        }

        private void ReplaceSelectedObject(SelectedWordObject target, string workingFolder)
        {
            if (!Directory.Exists(workingFolder))
            {
                throw new DirectoryNotFoundException("包内容文件夹不存在：" + workingFolder);
            }

            if (!target.IsExistingPackage && !DirectoryHasFiles(workingFolder))
            {
                throw new InvalidOperationException("包内容文件夹为空，请先添加文件。");
            }

            string zipPath = BuildTemporaryZipPathFromBase(
                target.IsExistingPackage ? "managed_ole_support" : "selected_image_support");
            try
            {
                zipPackages.CreateFromDirectory(workingFolder, zipPath, false);
                selectedObjects.ReplaceWithPackage(target, zipPath);
            }
            finally
            {
                TryDeleteFile(zipPath);
            }
        }

        private static string BuildTemporaryZipPath(string folderPath)
        {
            string baseName = new DirectoryInfo(folderPath).Name;
            return BuildTemporaryZipPathFromBase(string.IsNullOrWhiteSpace(baseName) ? "package" : baseName);
        }

        private static string BuildTemporaryZipPathFromBase(string baseName)
        {
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "package";
            }

            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(invalidCharacter, '_');
            }

            string fileName = string.Format(
                "{0}_{1:yyyyMMdd_HHmmss}_{2:N}.zip",
                baseName,
                DateTime.Now,
                Guid.NewGuid());
            string temporaryDirectory = Path.Combine(Path.GetTempPath(), ProductName);
            Directory.CreateDirectory(temporaryDirectory);
            return Path.Combine(temporaryDirectory, fileName);
        }

        private static string BuildWorkingFolderPath(string baseName)
        {
            string directoryName = string.Format(
                "OLEPackager_Work_{0}_{1:yyyyMMdd_HHmmss}_{2:N}",
                baseName,
                DateTime.Now,
                Guid.NewGuid());
            return Path.Combine(Path.GetTempPath(), directoryName);
        }

        private bool HasOpenDocument()
        {
            return application.Documents.Count > 0;
        }

        private void EnsureDocumentIsOpen()
        {
            if (!HasOpenDocument())
            {
                throw new InvalidOperationException("请先新建或打开一个 Word 文档。");
            }
        }

        private WordWindow GetWordWindow()
        {
            return new WordWindow(application.ActiveWindow.Hwnd);
        }

        private static bool IsSupportedDisplayImage(string imagePath)
        {
            string extension = Path.GetExtension(imagePath);
            return string.Equals(extension, ".png", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".jpg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".jpeg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".tif", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".tiff", StringComparison.OrdinalIgnoreCase);
        }

        private static bool DirectoryHasFiles(string directoryPath)
        {
            return Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories).Length > 0;
        }

        private static void TryDeleteDirectory(string directoryPath)
        {
            try
            {
                if (Directory.Exists(directoryPath))
                {
                    Directory.Delete(directoryPath, true);
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static void TryDeleteFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static void ExecuteWordCommand(Action command)
        {
            try
            {
                command();
            }
            catch (Exception exception)
            {
                ShowError(exception.Message);
            }
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private sealed class WordWindow : IWin32Window
        {
            private readonly IntPtr handle;

            public WordWindow(int handle)
            {
                this.handle = new IntPtr(handle);
            }

            public IntPtr Handle
            {
                get { return handle; }
            }
        }
    }
}
