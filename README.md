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

### 工作流一：plot 文件夹

1. 新建一个图目录，把绘图脚本、原始数据和辅助文件都放进去。
2. 用 `plot.py` 画图，并在同一个目录下输出最终图片 `plot.png`。
3. 在 Word 的 `Figure Package` 选项卡中点击 `Insert Figure Package`。
4. 选择刚才那个包含 `plot.py` 和 `plot.png` 的图目录。
5. 插件会把目录中的原始数据和脚本打包成 zip，并作为 OLE 对象嵌入当前 Word 文档；文档中显示的外观就是 `plot.png`。

### 工作流二：图片 + 支持文件

1. 在 Word 的 `Figure Package` 选项卡中点击 `Insert Image + Files`，打开确认窗口。
2. 在窗口中点击 `Choose image...`，选择要显示在文档中的图片，支持 `png`、`jpg`、`jpeg`、`tif`、`tiff`。
3. 点击 `Add files...` 多选需要嵌入的支持文件，文件类型不限；窗口中会列出已选择的文件。
4. 如果选错了，可以用 `Remove selected` 或 `Clear` 调整支持文件列表。
5. 确认图片和支持文件列表无误后，点击 `Insert`。
6. 插件会把这些支持文件打包成 zip，并作为 OLE 对象嵌入当前 Word 文档；文档中显示的是第一步选择的图片。
7. 如果多选的支持文件有同名文件，zip 内会自动改名为 `name_2.ext` 这类形式，避免覆盖。

### 工作流三：管理已有图片或 OLE 图包的附件

1. 在 Word 文档中选中一张已经插入的图片，或选中一个已有的 Figure Package OLE 对象。
2. 在 `Figure Package` 选项卡中点击 `Manage Image/OLE Files`，打开确认窗口。
3. 如果选中的是普通图片，窗口会从空列表开始；点击 `Add files...` 多选需要嵌入的支持文件，然后点击 `Insert`，插件会把原图片替换成 zip OLE 对象，并保持原来的大小和位置。
4. 如果选中的是已有 OLE 对象，窗口会列出当前 zip 中已有的文件，显示为 `[embedded] filename`。
5. 要删除已有附件，选中对应 `[embedded]` 条目并点击 `Remove selected`；要新增附件，点击 `Add files...`，新增条目会显示为 `[new] full\path`。
6. 确认列表无误后点击 `Insert`。插件会按当前列表重新生成 zip OLE 对象；Word 中显示出来仍然是原来的图，并保持原来的大小和位置。

卸载方式：

```text
Windows Settings -> Installed apps -> Figure Package Word Add-in -> Uninstall
```

也可以手动删除：

```text
%APPDATA%\Microsoft\Word\STARTUP\OleRawDataInserter.dotm
```

## 打包规则

`Insert Figure Package` 使用包含 `plot.png` 的文件夹。

压缩包内容规则：

- 不额外套一层父文件夹。
- 不生成 `.` / `./` 根目录项。
- 排除顶层生成图文件：`plot.png`、`plot.svg`、`plot.pdf`。
- 保留核心绘图脚本：`plot.py`。
- 保留其他数据、脚本、图片和子文件夹。

`Insert Image + Files` 使用一个显示图片和一组支持文件/文件夹。

压缩包内容规则：

- 包含用户添加的支持文件，以及用户添加文件夹中的所有文件。
- 支持文件类型不限。
- 支持文件进入 zip 根目录，不额外套一层父文件夹。
- 支持文件夹会以该文件夹名进入 zip，并保留文件夹内部的相对路径。
- 同名支持文件会自动改名，避免互相覆盖。

`Manage Image/OLE Files` 使用文档中已经选中的图片或 Figure Package OLE 对象。

压缩包内容规则：

- 选中普通图片时，新 zip 只包含用户添加的支持文件/文件夹内容。
- 选中已有 OLE 对象时，新 zip 包含列表中保留的 `[embedded]` 文件和新增的 `[new]` 文件/文件夹内容。
- 支持文件类型不限。
- 支持文件进入 zip 根目录，不额外套一层父文件夹。
- 支持文件夹会以该文件夹名进入 zip，并保留文件夹内部的相对路径。
- 同名新增文件会自动改名，避免互相覆盖。
- 显示外观来自文档中原本选中的图片或 OLE 对象，不需要另选图片文件。

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
    diagnose-dotm-save.ps1
    inject-ribbon.ps1
```

## 开发者文档

请看：

```text
OleRawDataInserter\DEVELOPER_NOTES.md
```

## Credits

- Thanks to Youhao SHANG for contributing to the original figure-packaging workflow and the idea of embedding raw data as an OLE object while displaying the figure itself in Word.

## License

This project is released under the MIT License. See [LICENSE](LICENSE).
