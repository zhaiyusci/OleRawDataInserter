using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OLEPackager.Core
{
    public enum PackageEditorMode
    {
        InsertNew,
        AttachToImage,
        ManagePackage
    }

    public sealed class PackageEditorDialog : Form
    {
        private readonly PackageEditorMode mode;
        private readonly PackageWorkspaceService workspace;
        private readonly Action<string, string> applyPackage;
        private readonly PictureBox imagePreview;
        private readonly TextBox imagePathTextBox;
        private readonly ListView entriesList;
        private readonly Label statusLabel;
        private readonly Button removeButton;
        private readonly Button applyButton;
        private string displayImagePath;

        public PackageEditorDialog(
            PackageEditorMode mode,
            string workingFolder,
            string displayImagePath,
            Action<string, string> applyPackage)
        {
            this.mode = mode;
            workspace = new PackageWorkspaceService(workingFolder);
            this.displayImagePath = displayImagePath;
            this.applyPackage = applyPackage;

            Text = GetWindowTitle(mode);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            ClientSize = new Size(820, 590);
            MinimumSize = new Size(720, 520);
            AllowDrop = true;

            TableLayoutPanel root = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 6,
                Dock = DockStyle.Fill,
                Padding = new Padding(14),
                Margin = Padding.Empty
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 132F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44F));

            Label introLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Text = GetIntroText(mode),
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 0, 6)
            };

            TableLayoutPanel imagePanel = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8)
            };
            imagePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 166F));
            imagePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            imagePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            imagePreview = new PictureBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                SizeMode = PictureBoxSizeMode.Zoom,
                Margin = new Padding(0, 0, 12, 0)
            };

            TableLayoutPanel imageDetails = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 3,
                Dock = DockStyle.Fill,
                Margin = Padding.Empty
            };
            imageDetails.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            imageDetails.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108F));
            imageDetails.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            imageDetails.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            imageDetails.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label imageLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Text = mode == PackageEditorMode.InsertNew ? "显示图片" : "保持所选对象的显示图",
                TextAlign = ContentAlignment.BottomLeft,
                Margin = Padding.Empty
            };
            imageDetails.SetColumnSpan(imageLabel, 2);

            imagePathTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Margin = new Padding(0, 4, 8, 4),
                Text = mode == PackageEditorMode.InsertNew
                    ? (displayImagePath ?? string.Empty)
                    : "来自 Word 中当前选中的对象"
            };

            Button chooseImageButton = CreateButton("选择图片...");
            chooseImageButton.Dock = DockStyle.Fill;
            chooseImageButton.Margin = new Padding(0, 3, 0, 3);
            chooseImageButton.Visible = mode == PackageEditorMode.InsertNew;
            chooseImageButton.Click += ChooseImageButtonClick;

            Label imageHintLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                ForeColor = SystemColors.GrayText,
                Text = mode == PackageEditorMode.InsertNew
                    ? "支持 PNG、JPEG 和 TIFF。图片只作为 Word 中的显示外观，不会自动加入附件列表。"
                    : "重建后会继续使用当前显示图，并保持原对象的尺寸与布局。",
                TextAlign = ContentAlignment.TopLeft,
                Margin = new Padding(0, 8, 0, 0)
            };
            imageDetails.SetColumnSpan(imageHintLabel, 2);

            imageDetails.Controls.Add(imageLabel, 0, 0);
            imageDetails.Controls.Add(imagePathTextBox, 0, 1);
            imageDetails.Controls.Add(chooseImageButton, 1, 1);
            imageDetails.Controls.Add(imageHintLabel, 0, 2);
            imagePanel.Controls.Add(imagePreview, 0, 0);
            imagePanel.Controls.Add(imageDetails, 1, 0);

            FlowLayoutPanel toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 6),
                Padding = Padding.Empty
            };

            Button addFilesButton = CreateButton("添加文件...");
            addFilesButton.Click += AddFilesButtonClick;
            Button addFolderButton = CreateButton("添加文件夹...");
            addFolderButton.Click += AddFolderButtonClick;
            removeButton = CreateButton("移除");
            removeButton.Enabled = false;
            removeButton.Click += RemoveButtonClick;
            Button clearButton = CreateButton("全部清空");
            clearButton.Click += ClearButtonClick;
            Button openFolderButton = CreateButton("在资源管理器中打开");
            openFolderButton.Click += OpenFolderButtonClick;
            Button refreshButton = CreateButton("刷新");
            refreshButton.Click += delegate { ExecuteUiAction(RefreshEntries); };

            toolbar.Controls.Add(addFilesButton);
            toolbar.Controls.Add(addFolderButton);
            toolbar.Controls.Add(removeButton);
            toolbar.Controls.Add(clearButton);
            toolbar.Controls.Add(openFolderButton);
            toolbar.Controls.Add(refreshButton);

            entriesList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = true,
                HideSelection = false,
                AllowDrop = true,
                Margin = Padding.Empty
            };
            entriesList.Columns.Add("包内路径", 430, HorizontalAlignment.Left);
            entriesList.Columns.Add("类型", 90, HorizontalAlignment.Left);
            entriesList.Columns.Add("大小", 90, HorizontalAlignment.Right);
            entriesList.Columns.Add("修改时间", 135, HorizontalAlignment.Left);
            entriesList.SelectedIndexChanged += delegate { removeButton.Enabled = entriesList.SelectedItems.Count > 0; };
            entriesList.DoubleClick += EntriesListDoubleClick;
            entriesList.KeyDown += EntriesListKeyDown;
            entriesList.Resize += delegate { AdjustColumns(); };
            entriesList.DragEnter += EntriesListDragEnter;
            entriesList.DragDrop += EntriesListDragDrop;

            statusLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
                AutoEllipsis = true,
                Margin = new Padding(0, 4, 0, 0)
            };

            FlowLayoutPanel bottomBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Margin = Padding.Empty,
                Padding = new Padding(0, 7, 0, 0)
            };
            Button cancelButton = CreateButton("取消");
            cancelButton.DialogResult = DialogResult.Cancel;
            applyButton = CreateButton(GetActionText(mode));
            applyButton.Click += ApplyButtonClick;
            bottomBar.Controls.Add(cancelButton);
            bottomBar.Controls.Add(applyButton);

            root.Controls.Add(introLabel, 0, 0);
            root.Controls.Add(imagePanel, 0, 1);
            root.Controls.Add(toolbar, 0, 2);
            root.Controls.Add(entriesList, 0, 3);
            root.Controls.Add(statusLabel, 0, 4);
            root.Controls.Add(bottomBar, 0, 5);
            Controls.Add(root);

            AcceptButton = applyButton;
            CancelButton = cancelButton;
            DragEnter += EntriesListDragEnter;
            DragDrop += EntriesListDragDrop;
            Activated += delegate { ExecuteUiAction(RefreshEntries, false); };

            SetPreviewImage(displayImagePath, false);
            RefreshEntries();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && imagePreview != null && imagePreview.Image != null)
            {
                Image image = imagePreview.Image;
                imagePreview.Image = null;
                image.Dispose();
            }

            base.Dispose(disposing);
        }

        private void ChooseImageButtonClick(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择显示图片";
                dialog.Filter = "支持的图片|*.png;*.jpg;*.jpeg;*.tif;*.tiff|PNG 图片|*.png|JPEG 图片|*.jpg;*.jpeg|TIFF 图片|*.tif;*.tiff|所有文件|*.*";
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    SetPreviewImage(dialog.FileName, true);
                }
            }
        }

        private void AddFilesButtonClick(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "添加到包内容";
                dialog.Filter = "所有文件|*.*";
                dialog.CheckFileExists = true;
                dialog.Multiselect = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    ExecuteUiAction(delegate { workspace.AddPaths(dialog.FileNames); });
                }
            }
        }

        private void AddFolderButtonClick(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择要添加到包内容根目录的文件夹";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    ExecuteUiAction(delegate { workspace.AddPaths(new[] { dialog.SelectedPath }); });
                }
            }
        }

        private void RemoveButtonClick(object sender, EventArgs e)
        {
            IList<PackageWorkspaceEntry> selectedEntries = GetSelectedEntries();
            if (selectedEntries.Count == 0)
            {
                return;
            }

            if (MessageBox.Show(
                    this,
                    "确定要从包内容中移除所选项目吗？",
                    "OLE Packager",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ExecuteUiAction(delegate { workspace.RemoveEntries(selectedEntries); });
            }
        }

        private void ClearButtonClick(object sender, EventArgs e)
        {
            if (entriesList.Items.Count == 0)
            {
                return;
            }

            if (MessageBox.Show(
                    this,
                    "确定要清空全部包内容吗？",
                    "OLE Packager",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ExecuteUiAction(workspace.Clear);
            }
        }

        private void OpenFolderButtonClick(object sender, EventArgs e)
        {
            ExecuteUiAction(workspace.OpenRootDirectory, false);
        }

        private void EntriesListDoubleClick(object sender, EventArgs e)
        {
            if (entriesList.SelectedItems.Count == 1)
            {
                PackageWorkspaceEntry entry = entriesList.SelectedItems[0].Tag as PackageWorkspaceEntry;
                ExecuteUiAction(delegate { workspace.OpenEntry(entry); }, false);
            }
        }

        private void EntriesListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                RemoveButtonClick(sender, EventArgs.Empty);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F5)
            {
                ExecuteUiAction(RefreshEntries);
                e.Handled = true;
            }
        }

        private void EntriesListDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void EntriesListDragDrop(object sender, DragEventArgs e)
        {
            string[] paths = e.Data == null ? null : e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths != null && paths.Length > 0)
            {
                ExecuteUiAction(delegate { workspace.AddPaths(paths); });
            }
        }

        private void ApplyButtonClick(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(displayImagePath) || !File.Exists(displayImagePath))
            {
                ShowError("请选择一个存在的显示图片。");
                return;
            }

            IList<PackageWorkspaceEntry> entries = workspace.GetEntries();
            int fileCount = entries.Count(entry => !entry.IsDirectory);
            if (mode != PackageEditorMode.ManagePackage && fileCount == 0)
            {
                ShowError("包内容为空，请先添加至少一个文件。");
                return;
            }

            try
            {
                UseWaitCursor = true;
                applyButton.Enabled = false;
                applyPackage(displayImagePath, workspace.RootDirectory);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                ShowError(exception.Message);
            }
            finally
            {
                UseWaitCursor = false;
                RefreshEntries();
            }
        }

        private void RefreshEntries()
        {
            IList<PackageWorkspaceEntry> entries = workspace.GetEntries();
            entriesList.BeginUpdate();
            try
            {
                entriesList.Items.Clear();
                foreach (PackageWorkspaceEntry entry in entries)
                {
                    ListViewItem item = new ListViewItem(entry.RelativePath)
                    {
                        Tag = entry
                    };
                    item.SubItems.Add(entry.IsDirectory ? "文件夹" : GetFileType(entry.FullPath));
                    item.SubItems.Add(entry.IsDirectory ? string.Empty : FormatBytes(entry.Size));
                    item.SubItems.Add(entry.Modified.ToString("yyyy-MM-dd HH:mm"));
                    entriesList.Items.Add(item);
                }
            }
            finally
            {
                entriesList.EndUpdate();
            }

            int fileCount = entries.Count(entry => !entry.IsDirectory);
            int folderCount = entries.Count(entry => entry.IsDirectory);
            long totalBytes = entries.Where(entry => !entry.IsDirectory).Sum(entry => entry.Size);
            statusLabel.Text = string.Format(
                "{0} 个文件，{1} 个文件夹，共 {2}    工作目录：{3}",
                fileCount,
                folderCount,
                FormatBytes(totalBytes),
                workspace.RootDirectory);
            removeButton.Enabled = false;
            applyButton.Enabled = !string.IsNullOrEmpty(displayImagePath)
                && File.Exists(displayImagePath)
                && (mode == PackageEditorMode.ManagePackage || fileCount > 0);
            AdjustColumns();
        }

        private IList<PackageWorkspaceEntry> GetSelectedEntries()
        {
            return entriesList.SelectedItems
                .Cast<ListViewItem>()
                .Select(item => item.Tag as PackageWorkspaceEntry)
                .Where(entry => entry != null)
                .ToList();
        }

        private void SetPreviewImage(string imagePath, bool showErrors)
        {
            Image newImage = null;
            try
            {
                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                {
                    using (Image source = Image.FromFile(imagePath))
                    {
                        newImage = new Bitmap(source);
                    }
                }
            }
            catch (Exception exception)
            {
                if (showErrors)
                {
                    ShowError("无法预览所选图片：" + exception.Message);
                }

                return;
            }

            Image oldImage = imagePreview.Image;
            imagePreview.Image = newImage;
            if (oldImage != null)
            {
                oldImage.Dispose();
            }

            displayImagePath = imagePath;
            if (mode == PackageEditorMode.InsertNew)
            {
                imagePathTextBox.Text = imagePath ?? string.Empty;
            }

            if (applyButton != null)
            {
                applyButton.Enabled = !string.IsNullOrEmpty(displayImagePath) && File.Exists(displayImagePath);
            }
        }

        private void ExecuteUiAction(Action action)
        {
            ExecuteUiAction(action, true);
        }

        private void ExecuteUiAction(Action action, bool refreshAfter)
        {
            try
            {
                UseWaitCursor = true;
                action();
                if (refreshAfter)
                {
                    RefreshEntries();
                }
            }
            catch (Exception exception)
            {
                ShowError(exception.Message);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        private void AdjustColumns()
        {
            if (entriesList.Columns.Count < 4)
            {
                return;
            }

            int fixedWidth = entriesList.Columns[1].Width
                + entriesList.Columns[2].Width
                + entriesList.Columns[3].Width
                + SystemInformation.VerticalScrollBarWidth
                + 12;
            entriesList.Columns[0].Width = Math.Max(220, entriesList.ClientSize.Width - fixedWidth);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(this, message, "OLE Packager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static Button CreateButton(string text)
        {
            return new Button
            {
                AutoSize = true,
                MinimumSize = new Size(86, 29),
                Height = 29,
                Text = text,
                Margin = new Padding(0, 0, 8, 0),
                Padding = new Padding(8, 0, 8, 0)
            };
        }

        private static string GetWindowTitle(PackageEditorMode mode)
        {
            switch (mode)
            {
                case PackageEditorMode.AttachToImage:
                    return "为所选图片添加包内容";
                case PackageEditorMode.ManagePackage:
                    return "管理所选 OLE 图包";
                default:
                    return "插入图片 + 文件";
            }
        }

        private static string GetIntroText(PackageEditorMode mode)
        {
            switch (mode)
            {
                case PackageEditorMode.AttachToImage:
                    return "把文件或文件夹添加到列表中，然后把当前 Word 图片转换为可双击打开的 OLE 图包。";
                case PackageEditorMode.ManagePackage:
                    return "当前嵌入内容已展开到列表中。可以添加、移除或拖放项目；重建时会保留显示图、尺寸和布局。";
                default:
                    return "选择 Word 中显示的图片，并在下方直接整理要嵌入的文件。也可以从资源管理器拖放文件或文件夹。";
            }
        }

        private static string GetActionText(PackageEditorMode mode)
        {
            return mode == PackageEditorMode.ManagePackage ? "重建" : "插入";
        }

        private static string GetFileType(string path)
        {
            string extension = Path.GetExtension(path).TrimStart('.');
            return string.IsNullOrEmpty(extension) ? "文件" : extension.ToUpperInvariant() + " 文件";
        }

        private static string FormatBytes(long bytes)
        {
            string[] units = { "B", "KB", "MB", "GB" };
            double value = bytes;
            int unit = 0;
            while (value >= 1024D && unit < units.Length - 1)
            {
                value /= 1024D;
                unit++;
            }

            return unit == 0
                ? string.Format("{0} {1}", bytes, units[unit])
                : string.Format("{0:0.##} {1}", value, units[unit]);
        }
    }
}
