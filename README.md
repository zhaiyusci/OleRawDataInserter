# Figure Package Word Add-in

这是从 `缝合怪.docm` 拆出的 Word 加载项项目。它把一个包含 `plot.png` 的图目录打包为 zip，并作为 OLE 对象插入当前 Word 文档；OLE 对象的显示外观使用 `plot.png`。

## 终端用户使用

给用户分发这个安装器：

```text
OleRawDataInserter\release\FigurePackageWordAddinSetup.exe
```

用户操作：

1. 关闭 Microsoft Word。
2. 双击 `FigurePackageWordAddinSetup.exe`。
3. 重新打开 Word。
4. 在 `Figure Package` 选项卡中点击 `Insert Figure Package`。
5. 选择包含 `plot.png` 的文件夹。

卸载方式：

```text
Windows Settings -> Installed apps -> Figure Package Word Add-in -> Uninstall
```

也可以手动删除：

```text
%APPDATA%\Microsoft\Word\STARTUP\OleRawDataInserter.dotm
```

## 打包规则

用户选择的是包含 `plot.png` 的文件夹。

压缩包内容规则：

- 不额外套一层父文件夹。
- 排除顶层生成图文件：`plot.png`、`plot.svg`、`plot.pdf`。
- 保留核心绘图脚本：`plot.py`。
- 保留其他数据、脚本、图片和子文件夹。

OLE 外观规则：

- 嵌入对象本体是 zip。
- Word 中显示为 `plot.png`。
- 不显示 zip 文件名。
- 使用透明 icon 避免出现默认文件图标。

## 项目结构

```text
OleRawDataInserter/
  assets/
    figure-package-icon.png
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
    RawDataOleInserter.bas
    RibbonCallbacks.bas
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
