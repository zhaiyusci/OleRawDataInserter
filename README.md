# OLE Packager

这是从 `缝合怪.docm` 拆出的 Office 加载项项目。它可以把图和支持材料打包为 zip，并作为 OLE 对象插入当前 Word 文档或 PowerPoint 幻灯片；OLE 对象显示为所选图片本身。

当前 `codex/vsto-migration` 分支使用共享 C# Core 加两个 VSTO 宿主加载项实现三个工作流。原 VBA `.dotm`、zip helper 和旧安装器仍保留在仓库中，供迁移对照。

## 终端用户使用

给用户分发这个安装器：

```text
OLEPackager\release\OLEPackagerVstoSetup.exe
```

安装步骤：

1. 关闭 Microsoft Word 和 PowerPoint。
2. 双击 `OLEPackagerVstoSetup.exe`。
3. 重新打开 Word 或 PowerPoint。

安装器默认会为当前 Windows 用户安装 Word 和 PowerPoint 两个 C# VSTO 加载项，也可以在安装类型中只选一个宿主。压缩、解压和 OLE 解析均在加载项中完成；终端用户不需要单独安装 Python、PowerShell 脚本、zip helper 或 7-Zip。

### 工作流一：plot 文件夹

1. 新建一个图目录，把绘图脚本、原始数据和辅助文件都放进去。
2. 用 `plot.py` 画图，并在同一个目录下输出最终图片 `plot.png`。
3. 在 Word 或 PowerPoint 的 `OLE Packager` 选项卡中点击 `Insert OLE Package`。
4. 选择刚才那个包含 `plot.py` 和 `plot.png` 的图目录。
5. 插件会把目录中的原始数据和脚本打包成 zip，并作为 OLE 对象嵌入当前文档或幻灯片；显示外观就是 `plot.png`。

### 工作流二：图片 + 包内容文件夹

1. 在 Word 或 PowerPoint 的 `OLE Packager` 选项卡中点击 `Insert Image + Files`，打开确认窗口。
2. 在窗口中点击 `Choose image...`，选择要显示在文档中的图片，支持 `png`、`jpg`、`jpeg`、`tif`、`tiff`。
3. 使用窗口中的文件列表直接添加文件、添加文件夹、拖放、移除或清空包内容。
4. 需要在 Windows Explorer 中进行更复杂的整理时，点击“在资源管理器中打开”。
5. 确认列表中的包内路径无误后点击“插入”。
6. 插件会把该文件夹中的内容打包成 zip，并作为 OLE 对象嵌入当前文档或幻灯片；显示的是第一步选择的图片。

### 工作流三：管理已有图片或 OLE 图包的附件

1. 在 Word 文档或 PowerPoint 幻灯片中选中一张普通图片，或选中一个已有的 OLE Packager 对象。
2. 在 `OLE Packager` 选项卡中点击 `Manage Image/OLE Files`，打开确认窗口。
3. 如果选中的是普通图片，插件会创建一个空的包内容文件夹。
4. 如果选中的是已有 OLE 对象，插件会先把当前 zip 解压到包内容文件夹。
5. 在编辑器列表中直接新增、删除或拖放文件和文件夹，也可以打开 Windows Explorer 继续整理。
6. 点击“插入”或“重建”。插件会按列表中的当前内容重新生成 zip OLE 对象；显示出来仍然是原来的图，并尽量保持原对象的布局与格式。

PowerPoint 会保留尺寸、位置、层级、宽高比、名称、替代文字、Title、标签、动作和动画绑定。Word 会保留尺寸、位置、环绕、锚点、间距、宽高比、名称和替代文字。由于 Office 的 OLE 图形本身不接受旋转和翻转，插件会在这些场景把最终视觉高质量烘焙为透明图片，同时继续保留 OLE 双击打开附件包的行为。PowerPoint 组内对象和占位符当前不会被替换。

卸载方式：

```text
Windows Settings -> Installed apps -> OLE Packager -> Uninstall
```

VSTO 版的安装目录和注册位置是：

```text
%APPDATA%\OLE Packager\VSTO\Word
%APPDATA%\OLE Packager\VSTO\PowerPoint
HKCU\Software\Microsoft\Office\Word\Addins\OLEPackager.WordAddIn
HKCU\Software\Microsoft\Office\PowerPoint\Addins\OLEPackager.PowerPointAddIn
```

## 打包规则

`Insert OLE Package` 使用包含 `plot.png` 的文件夹。

压缩包内容规则：

- 不额外套一层父文件夹。
- 不生成 `.` / `./` 根目录项。
- 排除顶层生成图文件：`plot.png`、`plot.svg`、`plot.pdf`。
- 保留核心绘图脚本：`plot.py`。
- 保留其他数据、脚本、图片和子文件夹。

`Insert Image + Files` 使用一个显示图片和一个临时包内容文件夹。

压缩包内容规则：

- 包含包内容文件夹中的所有文件。
- 支持文件类型不限。
- 包内容文件夹本身不会进入 zip；它的内容进入 zip 根目录。
- 子文件夹会保留相对路径。
- zip 文件名使用标准 UTF-8 编码；现代 Windows、7-Zip、WinRAR、macOS、Linux 工具通常可以正确显示中文文件/文件夹名。

`Manage Image/OLE Files` 使用文档中已经选中的图片或 OLE Packager 对象。

压缩包内容规则：

- 选中普通图片时，新 zip 只包含包内容文件夹中的文件/文件夹。
- 选中已有 OLE 对象时，旧 zip 会先解压到包内容文件夹；新 zip 由用户编辑后的当前文件夹内容生成。
- 支持文件类型不限。
- 包内容文件夹本身不会进入 zip；它的内容进入 zip 根目录。
- 子文件夹会保留相对路径。
- 显示外观来自文档中原本选中的图片或 OLE 对象，不需要另选图片文件。
- 新生成或重建的 zip 文件名使用标准 UTF-8 编码；旧 zip 如果本身已经乱码，插件无法保证自动恢复原始文件名。

OLE 外观规则：

- 嵌入对象本体是 zip。
- Word 和 PowerPoint 中显示为所选图片：plot 文件夹模式显示 `plot.png`，图片 + 支持文件模式显示用户选择的图片，管理已有图片或 OLE 图包时显示原本选中的图。
- 始终保持显示图片的高宽比。
- Word 新插入图片时按版心等比适配；PowerPoint 新插入图片时按幻灯片可用区域等比适配并居中。
- 管理已有图片或 OLE 图包时，保持原对象的几何位置；旋转或翻转对象使用视觉烘焙保持最终外观。
- 不显示 zip 文件名。
- 使用透明 icon 避免出现默认文件图标。

## 项目结构

```text
OLEPackager/
  LICENSE
  assets/
    insert-ole-package-icon.png
    insert-image-files-icon.png
    manage-image-ole-files-icon.png
    usage-help-icon.png
  customUI/
    customUI14.xml
  dist/
    OLEPackagerZipTool.exe
    OLEPackager.dotm
  installer/
    inno/
      OLEPackager.iss
      OLEPackagerVsto.iss
  release/
    OLEPackagerSetup.exe
    OLEPackagerVstoSetup.exe
  src-csharp/
    OLEPackager.Core/
      Models/
      Services/
      UI/
    OLEPackager.PowerPointAddIn/
      Ribbon/
      Services/
      UI/
    OLEPackager.WordAddIn/
      Ribbon/
      Services/
      UI/
  src/
    ImageSupportFilesDialog.frm
    ImageSupportFilesDialog.frx
    RawDataOleInserter.bas
    RibbonCallbacks.bas
    UsageHelpDialog.frm
    UsageHelpDialog.frx
  test/
    run-word-smoke-test-with-temp-vbom.ps1
    sample-plot/
  tools/
    build-word-addin.ps1
    diagnose-dotm-save.ps1
    OLEPackagerZipTool.cs
    inject-ribbon.ps1
```

## 开发者文档

构建 C# VSTO 版本前，Visual Studio 需要安装“Microsoft 365 开发”工作负载。构建和生成 Inno 安装器：

```powershell
cd C:\path\to\OLEPackager
powershell -ExecutionPolicy Bypass -File tools\build-vsto-addin.ps1 -Configuration Release
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\inno\OLEPackagerVsto.iss
```

`build-vsto-addin.ps1` 只负责编译和签署 VSTO 清单，不安装加载项。安装、升级和卸载统一由 Inno Setup 处理。旧 VBA 版本仍可通过 `tools\build-word-addin.ps1` 构建。

Inno 安装器会从已签署的 VSTO 清单读取公钥，并为 Word、PowerPoint 的精确安装 URL 写入当前用户的 VSTO Inclusion List；它不会把开发证书加入系统根证书信任区。

请看：

```text
OLEPackager\DEVELOPER_NOTES.md
```

## Credits

- Thanks to Youhao SHANG for contributing to the original figure-packaging workflow and the idea of embedding raw data as an OLE object while displaying the figure itself in Word.

## License

This project is released under the MIT License. See [LICENSE](LICENSE).
