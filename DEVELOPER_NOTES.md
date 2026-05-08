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
%APPDATA%\Microsoft\Word\STARTUP\FigurePackageZipTool.exe
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
- `Insert Image + Files`：打开确认窗口；用户选择显示图片后，插件创建临时包内容文件夹，用户点击 `Open Folder...` 后用 Windows Explorer 管理内容，最后点击 `Insert` 提交。
- `Manage Image/OLE Files`：选择文档中已有图片或 Figure Package OLE 对象后打开确认窗口；普通图片会使用空的包内容文件夹，已有 OLE 对象会先把现有 zip 解压到包内容文件夹；用户在 Windows Explorer 中增删改文件后提交，两种模式都会保持原来的显示外观、大小和位置。
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

UTF-8 zip helper：

```text
tools\FigurePackageZipTool.cs
dist\FigurePackageZipTool.exe
```

## 构建依赖

需要安装：

```text
Microsoft Word desktop
Inno Setup 6
.NET Framework 4.x C# compiler (csc.exe, included with many Windows/.NET Framework developer setups)
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

### 1. 编译 UTF-8 zip helper

运行：

```powershell
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:exe /out:OleRawDataInserter\dist\FigurePackageZipTool.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll OleRawDataInserter\tools\FigurePackageZipTool.cs
```

输出：

```text
OleRawDataInserter\dist\FigurePackageZipTool.exe
```

如果 64-bit `csc.exe` 不存在，可以改用：

```text
%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

### 2. 生成 dotm 并注入 Ribbon

运行：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\build-word-addin.ps1
```

这个脚本会启动可见 Word，并使用当前验证过的稳定保存顺序：

1. 先保存一个 plain `.dotm` 到 `%LOCALAPPDATA%\Temp\OleRawDataInserterBuild\PlainSaveTest.dotm`，用于预热 Word 的模板保存路径。
2. 再创建第二个文档，导入 VBA 模块和 UserForm，保存为 `%LOCALAPPDATA%\Temp\OleRawDataInserterBuild\VbaSaveTest.dotm`。
3. 关闭 Word。
4. 把 `VbaSaveTest.dotm` 复制为 `dist\OleRawDataInserter.dotm`。
5. 在 Word 已关闭后注入 Ribbon XML 和图标。

第二步会导入：

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
OleRawDataInserter\dist\OleRawDataInserter.dotm
```

脚本内部会调用 `tools\inject-ribbon.ps1`，把以下内容注入 `.dotm`：

```text
customUI/customUI14.xml
customUI/_rels/customUI14.xml.rels
customUI/images/insert-figure-package-icon.png
customUI/images/insert-image-files-icon.png
customUI/images/manage-image-ole-files-icon.png
customUI/images/usage-help-icon.png
```

如果只想生成 `dist\OleRawDataInserter.dotm`，不要加参数。  
如果要同时安装到本机 Word STARTUP 目录，使用：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\build-word-addin.ps1 -Install -CloseWord
```

参数说明：

- `-Install`：构建后复制 `OleRawDataInserter.dotm` 和 `FigurePackageZipTool.exe` 到 `%APPDATA%\Microsoft\Word\STARTUP`。
- `-CloseWord`：构建前结束当前用户正在运行的 `WINWORD.EXE`，避免 STARTUP 模板被锁。
- `-SkipVbomSetup`：不临时修改 `AccessVBOM` 注册表项；仅在已手动启用 “Trust access to the VBA project object model” 时使用。
- `-OfficeVersion`：显式指定 Office 注册表版本，例如 `16.0`。

`tools\diagnose-dotm-save.ps1` 仍保留为诊断脚本：它只验证 Word 是否能保存 plain/VBA `.dotm`，不复制到 `dist`，也不注入 Ribbon。

### 3. 编译安装器

确认以下两个文件都存在：

```text
OleRawDataInserter\dist\OleRawDataInserter.dotm
OleRawDataInserter\dist\FigurePackageZipTool.exe
```

运行：

```powershell
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" "OleRawDataInserter\installer\inno\FigurePackageWordAddin.iss"
```

Inno 脚本会把当前 `dist` 中的两个文件打进安装包：

```text
dist\OleRawDataInserter.dotm
dist\FigurePackageZipTool.exe
```

输出：

```text
OleRawDataInserter\release\FigurePackageWordAddinSetup.exe
```

如果编译器输出 `ArchitecturesInstallIn64BitMode=x64` 相关警告，说明脚本使用了旧架构标识；当前脚本应使用 `x64compatible`。

### 4. 本机安装验证

关闭 Word 后运行：

```powershell
$p = Start-Process -FilePath (Resolve-Path "OleRawDataInserter\release\FigurePackageWordAddinSetup.exe") -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" -Wait -PassThru
"InstallExitCode=$($p.ExitCode)"
Get-Item (Join-Path $env:APPDATA "Microsoft\Word\STARTUP\OleRawDataInserter.dotm")
Get-Item (Join-Path $env:APPDATA "Microsoft\Word\STARTUP\FigurePackageZipTool.exe")
```

`InstallExitCode=0` 表示安装成功。

## 一条命令重打包

如果只是已经改完源码、要重新生成并安装本机版本，可以运行：

```powershell
& "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:exe /out:OleRawDataInserter\dist\FigurePackageZipTool.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll OleRawDataInserter\tools\FigurePackageZipTool.cs
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\build-word-addin.ps1 -Install -CloseWord
```

如果想给这一步加 watchdog，避免 Word 自动化卡住，可以从项目目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\run-word-script-with-timeout.ps1 -ScriptPath OleRawDataInserter\tools\build-word-addin.ps1 -TimeoutSeconds 30
```

注意：`run-word-script-with-timeout.ps1` 用于诊断卡死；如果需要安装，最直接可靠的方式仍是直接运行：

```powershell
powershell -ExecutionPolicy Bypass -File OleRawDataInserter\tools\build-word-addin.ps1 -Install -CloseWord
```

发布安装器时继续运行：

```powershell
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

当前主要依赖随加载项安装的 `FigurePackageZipTool.exe`。这个 helper 由 `tools\FigurePackageZipTool.cs` 编译而来，使用 .NET `System.IO.Compression.ZipArchive`，并显式使用 `System.Text.Encoding.UTF8` 读写 entry names。这样新生成或重建的 zip 会使用标准 UTF-8 文件名，尽量避免中文文件/文件夹名在其他现代压缩包查看器里乱码。

VBA 会按以下顺序寻找 helper：

1. 环境变量 `OLE_RAW_DATA_ZIP_TOOL` 指定的路径，主要用于测试。
2. Word STARTUP 目录：`Application.StartupPath\FigurePackageZipTool.exe`。
3. `.dotm` 所在目录：`ThisDocument.Path\FigurePackageZipTool.exe`。

发布前需要先把 helper 编译到 `dist\FigurePackageZipTool.exe`；Inno 安装脚本会把 `OleRawDataInserter.dotm` 和 `FigurePackageZipTool.exe` 一起安装到 Word STARTUP 目录。

`Insert Figure Package` 的文件夹模式会排除顶层生成图文件。归档脚本会递归枚举源文件夹中的文件和空文件夹，把相对路径写成 zip entry name：

```text
sourceFolder\数据 文件夹\说明.txt -> 数据 文件夹/说明.txt
```

这里不会把源文件夹本身写进 zip，因此不会生成额外父目录，也不会生成 `.` / `./` 根目录项。顶层 `plot.png`、`plot.svg`、`plot.pdf` 在枚举阶段直接跳过。如果排除后没有任何条目，则创建真正的空 zip。

`Insert Image + Files` 的系统文件管理器模式会先创建临时包内容工作目录：

```text
%LOCALAPPDATA%\Temp\OleRawDataInserter_Work_<baseName>_<timestamp>\
```

实现规则：

- UserForm 先创建或准备工作目录，但不会自动打开 Windows Explorer；用户点击 `Open Folder...` 后才会调用 `explorer.exe <workingFolder>`。
- `Insert Image + Files` 和选中普通图片时使用空工作目录。
- 选中已有 OLE 对象时，先用 `ExtractZipToWorkingFolder` 把旧 zip 按 UTF-8 文件名解压到工作目录。
- 用户点击 `Insert` / `Rebuild` 后，通过 `CreateZipFromWorkingFolder(..., False/True)` 打包工作目录内容。
- 工作目录本身不会进入 zip；工作目录里的文件和子文件夹进入 zip 根目录并保留相对路径。
- 选中已有 OLE 对象时允许工作目录为空，此时生成空 zip；插入新图片包或给普通图片加附件时要求工作目录至少包含一个文件。

旧的 `CreateZipFromSupportFiles` / `CreateZipFromManagedFiles` 路径仍保留，主要用于兼容和低层测试；当前交互式 UI 优先走工作目录模式。

对已有 OLE 对象，管理逻辑会从 `Range.WordOpenXML` 中取出 `/word/embeddings/oleObject*.bin`，解析 OLE Compound File 中的 `Ole10Native` stream，抽出原 zip。随后把它解压到工作目录，用户用 Explorer 直接增删改文件；提交时重新打包工作目录，并用原显示图替换回新的 zip OLE 对象。

## 图片尺寸实现说明

尺寸入口：

```text
GetImagePrintSize
FitSizeToTextArea
GetCurrentTextAreaSize
AttachSupportFilesToInlineImage
AttachSupportFilesToFloatingImage
AttachWorkingFolderToInlineImage
AttachWorkingFolderToFloatingImage
ManageFilesInInlineOle
ManageFilesInFloatingOle
ManageWorkingFolderInInlineOle
ManageWorkingFolderInFloatingOle
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

把这个单文件发给用户即可。安装器内部会同时安装：

```text
OleRawDataInserter.dotm
FigurePackageZipTool.exe
```

不要只发送 `.dotm`；没有 `FigurePackageZipTool.exe` 时，UTF-8 zip 打包/列出/解压功能不可用。

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
%APPDATA%\Microsoft\Word\STARTUP\FigurePackageZipTool.exe
```

## 已知注意事项

- Word 需要允许加载宏模板。如果用户环境禁用了所有宏，插件不会运行。
- 构建 `.dotm` 时会启动 Word；如果 Word 卡住，先结束 `WINWORD.EXE` 再重新构建。
- 不建议把自动构建脚本交给终端用户。终端用户只使用 `FigurePackageWordAddinSetup.exe`。
