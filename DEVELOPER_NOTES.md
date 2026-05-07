# Developer Notes

这份文档记录如何维护、构建、验证和发布 `Figure Package Word Add-in`。

## 目标产物

最终给终端用户的是单个安装器：

```text
OleRawDataInserter\release\FigurePackageWordAddinSetup.exe
```

用户不需要打开 VBA，不需要运行 PowerShell 脚本，也不需要修改 Word 安全设置。

安装器会把加载项安装到当前用户的 Word STARTUP 目录：

```text
%APPDATA%\Microsoft\Word\STARTUP\OleRawDataInserter.dotm
```

## 关键源码

核心 VBA：

```text
src\RawDataOleInserter.bas
```

图片 + 支持文件对话框：

```text
src\ImageSupportFilesDialog.frm
src\ImageSupportFilesDialog.frx
```

Usage 滚动说明窗口：

```text
src\UsageHelpDialog.frm
src\UsageHelpDialog.frx
```

注意：UserForm 的 `.frm` 和 `.frx` 必须成对保留。VBE 导入 UserForm 时会读取 `.frm` 中的 `OleObjectBlob` 并加载同名 `.frx`；如果只有手写 `.frm` 而没有 `.frx`，Word 可能把它导入成普通模块，随后在 `.Show` 处编译/运行失败。

Ribbon 回调：

```text
src\RibbonCallbacks.bas
```

Ribbon XML：

```text
customUI\customUI14.xml
```

当前 Ribbon 有四个按钮：

- `Insert Figure Package`：选择包含 `plot.png` 的图目录并插入 OLE 图包。
- `Insert Image + Files`：打开确认窗口；用户可以在窗口中选择/查看显示图片和支持文件列表，最后点击 `Insert` 提交。
- `Manage Image/OLE Files`：选择文档中已有图片或 Figure Package OLE 对象后打开确认窗口；普通图片会被替换成显示外观相同的 zip OLE 对象，已有 OLE 对象会列出现有 zip 条目并允许添加/删除文件；两种模式都会保持原来的显示外观、大小和位置。
- `Usage`：打开带垂直滚动条的终端用户使用说明窗口。

Ribbon 图标：

```text
assets\insert-figure-package-icon.png
assets\insert-image-files-icon.png
assets\manage-image-ole-files-icon.png
assets\usage-help-icon.png
```

图标由脚本生成：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\generate-ribbon-icons.ps1
```

Inno Setup 安装器脚本：

```text
installer\inno\FigurePackageWordAddin.iss
```

## 构建依赖

需要安装：

```text
Microsoft Word desktop
Inno Setup 6
```

如果机器上有 `winget`，可以安装 Inno Setup：

```powershell
winget install --id JRSoftware.InnoSetup -e --accept-package-agreements --accept-source-agreements
```

本项目当前使用的 Inno 编译器路径通常是：

```text
%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe
```

## 推荐构建流程

请在项目根目录运行命令：

```powershell
cd "C:\Users\Yu Zhai\Desktop\缝合怪2.1"
```

### 1. 生成 dotm

运行：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\diagnose-dotm-save.ps1
```

这个脚本会启动可见 Word，创建临时 `.dotm`，并导入：

```text
src\RawDataOleInserter.bas
src\RibbonCallbacks.bas
src\ImageSupportFilesDialog.frm
src\ImageSupportFilesDialog.frx
src\UsageHelpDialog.frm
src\UsageHelpDialog.frx
```

输出文件在：

```text
%LOCALAPPDATA%\Temp\OleRawDataInserterBuild\VbaSaveTest.dotm
```

把它复制为正式加载项：

```powershell
Copy-Item -LiteralPath "$env:LOCALAPPDATA\Temp\OleRawDataInserterBuild\VbaSaveTest.dotm" -Destination "OleRawDataInserter\dist\OleRawDataInserter.dotm" -Force
```

### 2. 注入 Ribbon XML 和图标

运行：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\inject-ribbon.ps1 -DotmPath "OleRawDataInserter\dist\OleRawDataInserter.dotm"
```

这个脚本会把以下内容注入 `.dotm`：

```text
customUI/customUI14.xml
customUI/_rels/customUI14.xml.rels
customUI/images/insert-figure-package-icon.png
customUI/images/insert-image-files-icon.png
customUI/images/manage-image-ole-files-icon.png
customUI/images/usage-help-icon.png
```

### 3. 编译安装器

运行：

```powershell
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" "OleRawDataInserter\installer\inno\FigurePackageWordAddin.iss"
```

输出：

```text
OleRawDataInserter\release\FigurePackageWordAddinSetup.exe
```

### 4. 本机安装验证

关闭 Word 后运行：

```powershell
$p = Start-Process -FilePath (Resolve-Path "OleRawDataInserter\release\FigurePackageWordAddinSetup.exe") -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" -Wait -PassThru
"InstallExitCode=$($p.ExitCode)"
Get-Item (Join-Path $env:APPDATA "Microsoft\Word\STARTUP\OleRawDataInserter.dotm")
```

`InstallExitCode=0` 表示安装成功。

## 一条命令重打包

如果只是已经改完源码、要重新生成并安装本机版本，可以运行：

```powershell
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\diagnose-dotm-save.ps1
Copy-Item -LiteralPath "$env:LOCALAPPDATA\Temp\OleRawDataInserterBuild\VbaSaveTest.dotm" -Destination "OleRawDataInserter\dist\OleRawDataInserter.dotm" -Force
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\inject-ribbon.ps1 -DotmPath "OleRawDataInserter\dist\OleRawDataInserter.dotm"
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" "OleRawDataInserter\installer\inno\FigurePackageWordAddin.iss"
$p = Start-Process -FilePath (Resolve-Path "OleRawDataInserter\release\FigurePackageWordAddinSetup.exe") -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" -Wait -PassThru
"InstallExitCode=$($p.ExitCode)"
```

## 功能验证

烟测：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\test\run-word-smoke-test-with-temp-vbom.ps1
```

这个 wrapper 会用 watchdog 在子进程中运行真正的 Word smoke test。默认超时为 120 秒；如果 Word/VBE 因编译错误或模态错误窗口卡住，wrapper 会杀掉本次测试启动的 Word 进程并失败退出，不会无限等待。

调短超时时间进行诊断：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\test\run-word-smoke-test-with-temp-vbom.ps1 -TimeoutSeconds 30
```

成功时应看到：

```text
VBA SmokeTest OK
InlineShapes : 3
FloatingShapes : 1
```

检查最近生成的 zip 内容：

```powershell
$zip = Get-ChildItem "$env:LOCALAPPDATA\Temp" -Filter "sample-plot_*.zip" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
tar.exe -tf $zip.FullName
```

期望规则：

- 不出现外层 `sample-plot/`。
- 不出现 `.` 或 `./` 根目录项。
- 不包含 `plot.png`、`plot.svg`、`plot.pdf`。
- 包含 `plot.py`。
- 包含其他数据和图片。

## 打包实现说明

VBA 中的压缩入口：

```text
CreateZipFromPath
CreateZipFromSupportFiles
```

优先级：

1. `tar.exe`
2. PowerShell `Compress-Archive`
3. `Shell.Application`

当前主要依赖 Windows 自带 `tar.exe`，因为它在 Word VBA 中调用稳定，并能保留中文文件名。

`Insert Figure Package` 的文件夹模式会排除顶层生成图文件。`tar.exe` 命令形态：

```text
tar.exe -a -cf output.zip -C sourceFolder -- item1 item2 ...
```

这里不会使用 `-C sourceFolder .`，因为它会让部分工具显示一个 `.` / `./` 根目录项。VBA 会枚举源文件夹的顶层文件和子文件夹，只把这些条目传给 `tar.exe`；顶层 `plot.png`、`plot.svg`、`plot.pdf` 在枚举阶段直接跳过。如果排除后没有任何条目，则创建真正的空 zip。

`Insert Image + Files` 的多文件模式会先创建临时 staging 文件夹：

```text
%LOCALAPPDATA%\Temp\OleRawDataInserter_<imageBaseName>_<timestamp>\
```

实现规则：

- 用户多选的支持文件会复制到 staging 文件夹。
- 如果不同目录中的支持文件同名，复制时会改成 `name_2.ext`、`name_3.ext` 等形式，避免覆盖。
- staging 文件夹再通过 `CreateZipFromPath(..., False)` 打包；这个模式不会排除 `plot.png`、`plot.svg`、`plot.pdf`。
- zip 创建完成后会删除 staging 文件夹。

`Manage Image/OLE Files` 对普通图片复用 `CreateZipFromSupportFiles` 多文件打包逻辑。它不会把原图文件放进 zip；原图只作为 OLE 对象的显示外观，zip 里只包含用户选择的支持文件。

对已有 OLE 对象，管理逻辑会从 `Range.WordOpenXML` 中取出 `/word/embeddings/oleObject*.bin`，解析 OLE Compound File 中的 `Ole10Native` stream，抽出原 zip。确认窗口里 `[embedded]` 条目表示旧 zip 中保留的文件，`[new]` 条目表示用户新选的本地文件。提交时会解包旧 zip、删除未保留的条目、复制新增文件、重新打包，并用原显示图替换回新的 zip OLE 对象。

## 图片尺寸实现说明

尺寸入口：

```text
GetImagePrintSize
FitSizeToTextArea
GetCurrentTextAreaSize
AttachSupportFilesToInlineImage
AttachSupportFilesToFloatingImage
ManageFilesInInlineOle
ManageFilesInFloatingOle
ExtractImageFromOpenXml
ExtractZipFromOleOpenXml
```

规则：

- 先用 GDI+ 读取显示图片的像素尺寸和 DPI，换算为 Word points，得到原始印刷尺寸。
- 版心宽度取当前节 `PageWidth - LeftMargin - RightMargin - Gutter`。
- 版心高度取当前节 `PageHeight - TopMargin - BottomMargin`。
- 如果图片原始印刷尺寸能放进版心，则不放大，保持原始印刷尺寸。
- 如果图片超过版心，则按宽度比例和高度比例中较小的那个等比缩小。
- 插件会同时设置宽度和高度，缩放比例相同，因此高宽比保持不变。
- 给已有行内图片或行内 OLE 管理附件时，先提取该对象的显示图像，再删除原对象并在同一位置插入新的 OLE 对象，最后恢复原来的 `Width` 和 `Height`。
- 给已有浮动图片或浮动 OLE 管理附件时，先记录 `Width`、`Height`、`Left`、`Top`、相对定位和环绕方式；随后临时转换为行内对象以提取当前选中对象本身，再插入 OLE、转换回浮动 Shape，并恢复这些位置/布局属性。
- 显示图提取优先级是：链接图片的 `LinkFormat.SourceFullName`、`Range.WordOpenXML` 中的内嵌 `image/*` part、最后复制到临时 Word 文档并导出 filtered HTML 后取导出的图片文件。这个 fallback 用于兼容链接图片和某些粘贴来源图片。

## 发布给用户

发布前确认：

```text
release\FigurePackageWordAddinSetup.exe
```

把这个单文件发给用户即可。

用户说明：

1. 关闭 Word。
2. 双击安装器。
3. 重新打开 Word。
4. 在 `Figure Package` 选项卡点击 `Insert Figure Package`、`Insert Image + Files`，或先选中已有图片/OLE 对象再点击 `Manage Image/OLE Files`。

## 卸载

用户可以在 Windows 设置中卸载：

```text
Installed apps -> Figure Package Word Add-in -> Uninstall
```

也可以手动删除：

```text
%APPDATA%\Microsoft\Word\STARTUP\OleRawDataInserter.dotm
```

## 已知注意事项

- Word 需要允许加载宏模板。如果用户环境禁用了所有宏，插件不会运行。
- 构建 `.dotm` 时会启动 Word；如果 Word 卡住，先结束 `WINWORD.EXE` 再重新构建。
- 不建议把自动构建脚本交给终端用户。终端用户只使用 `FigurePackageWordAddinSetup.exe`。
