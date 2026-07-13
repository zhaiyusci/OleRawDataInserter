using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OLEPackager.PowerPointAddIn
{
    internal sealed class OlePackagerService
    {
        private const string ProductName = "OLE Packager";

        private readonly PowerPoint.Application application;
        private readonly ZipPackageService zipPackages;
        private readonly PowerPointOlePackageService powerPointPackages;
        private readonly SelectedPowerPointObjectService selectedObjects;

        public OlePackagerService(
            PowerPoint.Application application,
            ZipPackageService zipPackages,
            PowerPointOlePackageService powerPointPackages,
            SelectedPowerPointObjectService selectedObjects)
        {
            this.application = application;
            this.zipPackages = zipPackages;
            this.powerPointPackages = powerPointPackages;
            this.selectedObjects = selectedObjects;
        }

        public void InsertPlotFolderAsOle()
        {
            string folderPath = PickPlotFolder();
            if (!string.IsNullOrEmpty(folderPath))
            {
                InsertPlotFolder(folderPath);
            }
        }

        internal void InsertPlotFolder(string folderPath)
        {
            ExecutePowerPointCommand(delegate
            {
                PowerPoint.Slide slide = GetActiveSlide();
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

                string zipPath = BuildTemporaryZipPath(new DirectoryInfo(normalizedFolder).Name);
                try
                {
                    zipPackages.CreateFromDirectory(normalizedFolder, zipPath, true);
                    powerPointPackages.InsertPackage(slide, plotPath, zipPath);
                }
                finally
                {
                    TryDeleteFile(zipPath);
                }
            });
        }

        public void InsertImageSupportFilesAsOle()
        {
            if (!HasOpenPresentation())
            {
                ShowError("请先新建或打开一个 PowerPoint 演示文稿。");
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
                    dialog.ShowDialog(GetPowerPointWindow());
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
            if (!HasOpenPresentation())
            {
                ShowError("请先新建或打开一个 PowerPoint 演示文稿。");
                return;
            }

            SelectedPowerPointObject target = null;
            string workingFolder = null;
            try
            {
                target = selectedObjects.CaptureSelection();
                if (target == null)
                {
                    ShowError("请在一张普通幻灯片上只选中一个图片或 OLE Packager 对象。");
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

                SelectedPowerPointObject capturedTarget = target;
                using (PackageEditorDialog dialog = new PackageEditorDialog(
                    target.IsExistingPackage ? PackageEditorMode.ManagePackage : PackageEditorMode.AttachToImage,
                    workingFolder,
                    target.DisplayImagePath,
                    delegate(string imagePath, string folder) { ReplaceSelectedObject(capturedTarget, folder); }))
                {
                    dialog.ShowDialog(GetPowerPointWindow());
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
                dialog.ShowDialog(GetPowerPointWindow());
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
                return dialog.Show() == -1 ? dialog.SelectedItems.Item(1) : null;
            }
            finally
            {
                if (dialog != null)
                {
                    Marshal.ReleaseComObject(dialog);
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

            string zipPath = BuildTemporaryZipPath(Path.GetFileNameWithoutExtension(imagePath) + "_support");
            try
            {
                zipPackages.CreateFromDirectory(workingFolder, zipPath, false);
                powerPointPackages.InsertPackage(GetActiveSlide(), imagePath, zipPath);
            }
            finally
            {
                TryDeleteFile(zipPath);
            }
        }

        private void ReplaceSelectedObject(SelectedPowerPointObject target, string workingFolder)
        {
            if (!Directory.Exists(workingFolder))
            {
                throw new DirectoryNotFoundException("包内容文件夹不存在：" + workingFolder);
            }

            if (!target.IsExistingPackage && !DirectoryHasFiles(workingFolder))
            {
                throw new InvalidOperationException("包内容文件夹为空，请先添加文件。");
            }

            string zipPath = BuildTemporaryZipPath(target.IsExistingPackage ? "managed_ole_support" : "selected_image_support");
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

        private PowerPoint.Slide GetActiveSlide()
        {
            if (!HasOpenPresentation() || application.Windows.Count == 0)
            {
                throw new InvalidOperationException("请先新建或打开演示文稿并切换到一张幻灯片。");
            }

            try
            {
                PowerPoint.Slide slide = application.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide != null)
                {
                    return slide;
                }
            }
            catch (COMException)
            {
            }

            try
            {
                PowerPoint.Selection selection = application.ActiveWindow.Selection;
                if (selection != null && selection.SlideRange.Count > 0)
                {
                    return selection.SlideRange[1];
                }
            }
            catch (COMException)
            {
            }

            throw new InvalidOperationException("请切换到普通视图并选中一张幻灯片。");
        }

        private bool HasOpenPresentation()
        {
            return application.Presentations.Count > 0;
        }

        private PowerPointWindow GetPowerPointWindow()
        {
            return new PowerPointWindow(application.ActiveWindow.HWND);
        }

        private static string BuildTemporaryZipPath(string baseName)
        {
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "package";
            }

            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(invalidCharacter, '_');
            }

            string directory = Path.Combine(Path.GetTempPath(), ProductName);
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, string.Format("{0}_{1:yyyyMMdd_HHmmss}_{2:N}.zip", baseName, DateTime.Now, Guid.NewGuid()));
        }

        private static string BuildWorkingFolderPath(string baseName)
        {
            return Path.Combine(Path.GetTempPath(), string.Format("OLEPackager_Work_{0}_{1:yyyyMMdd_HHmmss}_{2:N}", baseName, DateTime.Now, Guid.NewGuid()));
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
            try { if (Directory.Exists(directoryPath)) { Directory.Delete(directoryPath, true); } }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static void TryDeleteFile(string filePath)
        {
            try { if (File.Exists(filePath)) { File.Delete(filePath); } }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private static void ExecutePowerPointCommand(Action command)
        {
            try { command(); }
            catch (Exception exception) { ShowError(exception.Message); }
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private sealed class PowerPointWindow : IWin32Window
        {
            private readonly IntPtr handle;
            public PowerPointWindow(int handle) { this.handle = new IntPtr(handle); }
            public IntPtr Handle { get { return handle; } }
        }
    }
}
