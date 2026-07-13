using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace OLEPackager.PowerPointAddIn
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

            TableLayoutPanel root = new TableLayoutPanel { ColumnCount = 1, RowCount = 3, Dock = DockStyle.Fill, Padding = new Padding(14) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));

            TabControl tabs = new TabControl { Dock = DockStyle.Fill, Margin = Padding.Empty };
            tabs.TabPages.Add(CreatePage("插入 OLE 图包",
                "1. 把绘图脚本、数据和 plot.png 放在同一文件夹。\r\n" +
                "2. 点击 OLE Packager > Insert OLE Package。\r\n" +
                "3. 选择该文件夹。\r\n\r\n" +
                "插件会把源文件压缩后嵌入当前幻灯片，并以 plot.png 作为显示图。顶层 plot.png、plot.svg 和 plot.pdf 不进入附件包。"));
            tabs.TabPages.Add(CreatePage("插入图片 + 文件",
                "1. 点击 OLE Packager > Insert Image + Files。\r\n" +
                "2. 选择 PNG、JPEG 或 TIFF 显示图片。\r\n" +
                "3. 添加或拖放附件文件和文件夹。\r\n" +
                "4. 点击“插入”。\r\n\r\n" +
                "新对象会按图片比例放入当前幻灯片中央。"));
            tabs.TabPages.Add(CreatePage("管理图片/OLE",
                "1. 在 PowerPoint 中选中一个普通图片或 OLE Packager 图包。\r\n" +
                "2. 点击 OLE Packager > Manage Image/OLE Files。\r\n" +
                "3. 修改附件后点击“插入”或“重建”。\r\n\r\n" +
                "插件会保留尺寸、位置、旋转、裁剪、层级、边框、阴影、名称、替代文本、标签、动作和动画绑定。组内对象与占位符当前不会被替换。"));

            Label versionLabel = new Label { Dock = DockStyle.Fill, ForeColor = SystemColors.GrayText, TextAlign = ContentAlignment.MiddleLeft, Text = "C# VSTO " + Assembly.GetExecutingAssembly().GetName().Version };
            FlowLayoutPanel bottomBar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Padding = new Padding(0, 7, 0, 0), Margin = Padding.Empty };
            Button closeButton = new Button { DialogResult = DialogResult.OK, MinimumSize = new Size(86, 29), Text = "关闭" };
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
            TabPage page = new TabPage(title);
            Panel scrollPanel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = SystemColors.Window, Padding = new Padding(18) };
            Label content = new Label { AutoSize = true, Font = new Font("Segoe UI", 10F), Text = text, Location = new Point(18, 18) };
            scrollPanel.Resize += delegate { content.MaximumSize = new Size(Math.Max(200, scrollPanel.ClientSize.Width - 48), 0); };
            scrollPanel.Controls.Add(content);
            page.Controls.Add(scrollPanel);
            return page;
        }
    }
}
