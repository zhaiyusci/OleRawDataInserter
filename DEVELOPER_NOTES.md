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

Ribbon 回调：

```text
src\RibbonCallbacks.bas
```

Ribbon XML：

```text
customUI\customUI14.xml
```

Ribbon 图标：

```text
assets\figure-package-icon.png
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
customUI/images/figure-package-icon.png
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

成功时应看到：

```text
VBA SmokeTest OK
InlineShapes : 1
```

检查最近生成的 zip 内容：

```powershell
$zip = Get-ChildItem "$env:LOCALAPPDATA\Temp" -Filter "sample-plot_*.zip" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
tar.exe -tf $zip.FullName
```

期望规则：

- 不出现外层 `sample-plot/`。
- 不包含 `plot.png`、`plot.svg`、`plot.pdf`。
- 包含 `plot.py`。
- 包含其他数据和图片。

注意：`tar.exe -tf` 可能显示一个 `./` 根目录项。这不是多套一层文件夹，只是 zip 根目录记录；Windows 解压时不会生成额外父文件夹。

## 打包实现说明

VBA 中的压缩入口：

```text
CreateZipFromPath
```

优先级：

1. `tar.exe`
2. PowerShell `Compress-Archive`
3. `Shell.Application`

当前主要依赖 Windows 自带 `tar.exe`，因为它在 Word VBA 中调用稳定，并能保留中文文件名。

文件夹模式下的 `tar.exe` 命令形态：

```text
tar.exe -a -cf output.zip --exclude=./plot.png --exclude=./plot.svg --exclude=./plot.pdf -C sourceFolder .
```

这里 `-C sourceFolder .` 的作用是把源文件夹的内容放到 zip 根目录，而不是把源文件夹本身放进去。

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
4. 在 `Figure Package` 选项卡点击 `Insert Figure Package`。

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
