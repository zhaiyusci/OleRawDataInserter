# Figure Package Word Add-in

这是从 `缝合怪.docm` 拆出的 Word 加载项项目。它可以把图和支持材料打包为 zip，并作为 OLE 对象插入当前 Word 文档；OLE 对象在 Word 中显示为所选图片本身。

## 终端用户使用

给用户分发这个安装器：

```text
OleRawDataInserter\release\FigurePackageWordAddinSetup.exe
```

安装步骤：

1. 关闭 Microsoft Word。
2. 双击 `FigurePackageWordAddinSetup.exe`。
3. 重新打开 Word。

安装器会为当前 Windows 用户安装 Word 加载项本体和 UTF-8 zip helper；终端用户不需要单独安装 Python、PowerShell 脚本或 7-Zip。

### 工作流一：plot 文件夹

1. 新建一个图目录，把绘图脚本、原始数据和辅助文件都放进去。
2. 用 `plot.py` 画图，并在同一个目录下输出最终图片 `plot.png`。
3. 在 Word 的 `Figure Package` 选项卡中点击 `Insert Figure Package`。
4. 选择刚才那个包含 `plot.py` 和 `plot.png` 的图目录。
5. 插件会把目录中的原始数据和脚本打包成 zip，并作为 OLE 对象嵌入当前 Word 文档；文档中显示的外观就是 `plot.png`。

### 工作流二：图片 + 包内容文件夹

1. 在 Word 的 `Figure Package` 选项卡中点击 `Insert Image + Files`，打开确认窗口。
2. 在窗口中点击 `Choose image...`，选择要显示在文档中的图片，支持 `png`、`jpg`、`jpeg`、`tif`、`tiff`。
3. 插件会创建一个临时包内容文件夹，并在窗口里提示它的位置。
4. 点击 `Open Folder...`，在 Windows Explorer 中把需要嵌入的文件和文件夹复制、拖放、删除或重命名。
5. 确认包内容文件夹无误后，回到 Word 窗口点击 `Insert`。
6. 插件会把该文件夹中的内容打包成 zip，并作为 OLE 对象嵌入当前 Word 文档；文档中显示的是第一步选择的图片。

### 工作流三：管理已有图片或 OLE 图包的附件

1. 在 Word 文档中选中一张已经插入的图片，或选中一个已有的 Figure Package OLE 对象。
2. 在 `Figure Package` 选项卡中点击 `Manage Image/OLE Files`，打开确认窗口。
3. 如果选中的是普通图片，插件会创建一个空的包内容文件夹。
4. 如果选中的是已有 OLE 对象，插件会先把当前 zip 解压到包内容文件夹。
5. 点击 `Open Folder...`，在 Windows Explorer 中直接新增、删除、重命名或移动文件和文件夹。
6. 回到 Word 窗口点击 `Insert` 或 `Rebuild`。插件会按该文件夹的当前内容重新生成 zip OLE 对象；Word 中显示出来仍然是原来的图，并保持原来的大小和位置。

卸载方式：

```text
Windows Settings -> Installed apps -> Figure Package Word Add-in -> Uninstall
```

也可以手动删除：

```text
%APPDATA%\Microsoft\Word\STARTUP\OleRawDataInserter.dotm
%APPDATA%\Microsoft\Word\STARTUP\FigurePackageZipTool.exe
```

## 打包规则

`Insert Figure Package` 使用包含 `plot.png` 的文件夹。

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

`Manage Image/OLE Files` 使用文档中已经选中的图片或 Figure Package OLE 对象。

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
- Word 中显示为所选图片：plot 文件夹模式显示 `plot.png`，图片 + 支持文件模式显示用户选择的图片，管理已有图片或 OLE 图包时显示原本选中的图。
- 始终保持显示图片的高宽比。
- 新插入图片时，如果显示图片的印刷尺寸小于版心，保持原始印刷尺寸。
- 新插入图片时，如果显示图片的印刷尺寸超过版心，等比缩小到能放进版心。
- 管理已有图片或 OLE 图包时，保持原对象在文档中的大小和位置。
- 不显示 zip 文件名。
- 使用透明 icon 避免出现默认文件图标。

## 项目结构

```text
OleRawDataInserter/
  LICENSE
  assets/
    insert-figure-package-icon.png
    insert-image-files-icon.png
    manage-image-ole-files-icon.png
    usage-help-icon.png
  customUI/
    customUI14.xml
  dist/
    FigurePackageZipTool.exe
    OleRawDataInserter.dotm
  installer/
    inno/
      FigurePackageWordAddin.iss
  release/
    FigurePackageWordAddinSetup.exe
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
    FigurePackageZipTool.cs
    inject-ribbon.ps1
```

## 开发者文档

本机开发时推荐使用自动构建脚本：

```powershell
cd "C:\Users\Yu Zhai\Desktop\缝合怪2.1\OleRawDataInserter"
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:exe /out:dist\FigurePackageZipTool.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tools\FigurePackageZipTool.cs
powershell -ExecutionPolicy Bypass -File tools\build-word-addin.ps1 -Install -CloseWord
```

`build-word-addin.ps1` 会先让 Word 保存一个 plain `.dotm`，再保存导入 VBA 后的 `.dotm`；Word 退出后才复制到 `dist`、注入 Ribbon，并在 `-Install` 模式下安装到 Word STARTUP 目录。这是当前验证过的稳定构建方式。

请看：

```text
OleRawDataInserter\DEVELOPER_NOTES.md
```

## Credits

- Thanks to Youhao SHANG for contributing to the original figure-packaging workflow and the idea of embedding raw data as an OLE object while displaying the figure itself in Word.

## License

This project is released under the MIT License. See [LICENSE](LICENSE).
