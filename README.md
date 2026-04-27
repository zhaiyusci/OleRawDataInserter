# Figure Package Word Add-in

这是从 `缝合怪.docm` 拆出的 Word 加载项项目。它把一个包含 `plot.png` 的图目录打包为 zip，并作为 OLE 对象插入当前 Word 文档；OLE 对象的显示外观使用 `plot.png`。

## 终端用户使用

给用户分发这个安装器：

```text
OleRawDataInserter\release\FigurePackageWordAddinSetup.exe
```

典型工作流：

1. 新建一个图目录，把绘图脚本、原始数据和辅助文件都放进去。
2. 用 `plot.py` 画图，并在同一个目录下输出最终图片 `plot.png`。
3. 在 Word 的 `Figure Package` 选项卡中点击 `Insert Figure Package`。
4. 选择刚才那个包含 `plot.py` 和 `plot.png` 的图目录。
5. 插件会把目录中的原始数据和脚本打包成 zip，并作为 OLE 对象嵌入当前 Word 文档；文档中显示的外观就是 `plot.png`。

安装步骤：

1. 关闭 Microsoft Word。
2. 双击 `FigurePackageWordAddinSetup.exe`。
3. 重新打开 Word。

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
- 始终保持 `plot.png` 的高宽比。
- 如果 `plot.png` 的印刷尺寸小于版心，保持原始印刷尺寸。
- 如果 `plot.png` 的印刷尺寸超过版心，等比缩小到能放进版心。
- 不显示 zip 文件名。
- 使用透明 icon 避免出现默认文件图标。

## 项目结构

```text
OleRawDataInserter/
  LICENSE
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

## Credits

- Thanks to Youhao SHANG for contributing to the original figure-packaging workflow and the idea of embedding raw data as an OLE object while displaying the figure itself in Word.

## License

This project is released under the MIT License. See [LICENSE](LICENSE).
