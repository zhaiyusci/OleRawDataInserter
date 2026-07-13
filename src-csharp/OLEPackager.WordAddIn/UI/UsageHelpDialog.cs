using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace OLEPackager.WordAddIn
{
    internal sealed class UsageHelpDialog : Form
    {
        public UsageHelpDialog()
        {
            Text = "OLE Packager 使用说明";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            ClientSize = new Size(720, 500);
            MinimumSize = new Size(620, 420);

            TableLayoutPanel root = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 3,
                Dock = DockStyle.Fill,
                Padding = new Padding(14)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));

            TabControl tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty
            };
            tabs.TabPages.Add(CreatePage(
                "插入 OLE 图包",
                "适合 plot.py、原始数据和 plot.png 放在同一文件夹的绘图工作流。\r\n\r\n" +
                "1. 在同一文件夹中准备绘图脚本、数据和其他源文件。\r\n" +
                "2. 生成最终显示图片 plot.png。\r\n" +
                "3. 点击 OLE Packager > Insert OLE Package。\r\n" +
                "4. 选择包含 plot.png 的文件夹。\r\n\r\n" +
                "插件会把文件夹内容嵌入为 zip OLE 对象，并以 plot.png 作为 Word 中的显示图。顶层 plot.png、plot.svg 和 plot.pdf 不进入 zip；同名文件位于子文件夹时仍会保留。"));
            tabs.TabPages.Add(CreatePage(
                "插入图片 + 文件",
                "适合显示图片与附件来自不同位置的工作流。\r\n\r\n" +
                "1. 点击 OLE Packager > Insert Image + Files。\r\n" +
                "2. 选择 PNG、JPEG 或 TIFF 显示图片。\r\n" +
                "3. 使用“添加文件”“添加文件夹”或直接拖放来整理包内容。\r\n" +
                "4. 在列表中检查包内路径、文件大小和修改时间。\r\n" +
                "5. 点击“插入”。\r\n\r\n" +
                "显示图片不会自动放进附件列表。需要把它也作为附件时，可再次通过“添加文件”加入。"));
            tabs.TabPages.Add(CreatePage(
                "管理图片/OLE",
                "适合给已有图片添加附件，或修改现有 OLE 图包。\r\n\r\n" +
                "1. 在 Word 中选中普通图片或 OLE Packager 图包。\r\n" +
                "2. 点击 OLE Packager > Manage Image/OLE Files。\r\n" +
                "3. 普通图片会从空列表开始；已有图包会显示当前嵌入内容。\r\n" +
                "4. 添加、移除、清空或拖放文件和文件夹。\r\n" +
                "5. 点击“插入”或“重建”。\r\n\r\n" +
                "插件会保持原显示图、尺寸、宽高比、裁剪、边框、阴影和辅助文本。浮动对象还会保留旋转、翻转、层级、环绕方式、锚点及相对位置。已有图包允许重建为空 zip。"));

            Label versionLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                ForeColor = SystemColors.GrayText,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "C# VSTO " + Assembly.GetExecutingAssembly().GetName().Version
            };

            FlowLayoutPanel bottomBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 7, 0, 0),
                Margin = Padding.Empty
            };
            Button closeButton = new Button
            {
                DialogResult = DialogResult.OK,
                MinimumSize = new Size(86, 29),
                Text = "关闭"
            };
            bottomBar.Controls.Add(closeButton);

            root.Controls.Add(tabs, 0, 0);
            root.Controls.Add(versionLabel, 0, 1);
            root.Controls.Add(bottomBar, 0, 2);
            Controls.Add(root);
            AcceptButton = closeButton;
            CancelButton = closeButton;
        }

        private static TabPage CreatePage(string title, string text)
        {
            TabPage page = new TabPage(title)
            {
                Padding = new Padding(0)
            };
            Panel scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = SystemColors.Window,
                Padding = new Padding(18)
            };
            Label content = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                Text = text,
                Location = new Point(18, 18)
            };
            scrollPanel.Resize += delegate
            {
                content.MaximumSize = new Size(Math.Max(200, scrollPanel.ClientSize.Width - 48), 0);
            };
            scrollPanel.Controls.Add(content);
            page.Controls.Add(scrollPanel);
            return page;
        }
    }
}
