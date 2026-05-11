Attribute VB_Name = "Localization"
Option Explicit

Public Function T(ByVal key As String) As String
    If IsChineseUiLanguage() Then
        T = LocalizedZh(key)
    Else
        T = LocalizedEn(key)
    End If

    If Len(T) = 0 Then T = key
End Function

Public Function TF(ByVal key As String, ByVal value1 As String, Optional ByVal value2 As String = vbNullString) As String
    TF = Replace(T(key), "%1", value1)
    TF = Replace(TF, "%2", value2)
End Function

Private Function IsChineseUiLanguage() As Boolean
    Dim languageId As Long

    On Error Resume Next
    languageId = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If Err.Number <> 0 Then
        Err.Clear
        languageId = Application.Language
    End If
    On Error GoTo 0

    Select Case languageId
        Case 1028, 2052, 3076, 4100, 5124
            IsChineseUiLanguage = True
        Case Else
            IsChineseUiLanguage = False
    End Select
End Function

Private Function LocalizedEn(ByVal key As String) As String
    Select Case key
        Case "app.name": LocalizedEn = "OLE Packager"
        Case "ribbon.group.insert": LocalizedEn = "Insert"
        Case "ribbon.insertPackage.label": LocalizedEn = "Insert OLE Package"
        Case "ribbon.insertPackage.screentip": LocalizedEn = "Insert plot.png with source files"
        Case "ribbon.insertPackage.supertip": LocalizedEn = "Choose a folder containing plot.png. The add-in embeds the whole folder as a zip OLE object and displays plot.png."
        Case "ribbon.insertImageFiles.label": LocalizedEn = "Insert Image + Files"
        Case "ribbon.insertImageFiles.screentip": LocalizedEn = "Insert an image with selected support files"
        Case "ribbon.insertImageFiles.supertip": LocalizedEn = "Open a folder in Explorer, edit package contents, then embed them as a zip OLE object shown as the selected image."
        Case "ribbon.manage.label": LocalizedEn = "Manage Image/OLE Files"
        Case "ribbon.manage.screentip": LocalizedEn = "Manage files attached to an image or OLE object"
        Case "ribbon.manage.supertip": LocalizedEn = "Select an existing picture or OLE Packager object, then add or remove embedded files while preserving the display, size, and position."
        Case "ribbon.usage.label": LocalizedEn = "Usage"
        Case "ribbon.usage.screentip": LocalizedEn = "Show usage instructions"
        Case "ribbon.usage.supertip": LocalizedEn = "Learn how to prepare plot.py, plot.png, and package contents before inserting an OLE package."

        Case "dialog.insert.caption": LocalizedEn = "Insert Image + Package Folder"
        Case "dialog.insert.intro": LocalizedEn = "1. Choose the display image. 2. Click Open Folder... and edit the package contents in Windows Explorer. 3. Return here and click Insert."
        Case "dialog.displayImage": LocalizedEn = "Display image"
        Case "dialog.chooseImage": LocalizedEn = "Choose image..."
        Case "dialog.packageFolder": LocalizedEn = "Package folder"
        Case "dialog.openFolder": LocalizedEn = "Open Folder..."
        Case "dialog.clearFolder": LocalizedEn = "Clear Folder"
        Case "dialog.insertButton": LocalizedEn = "Insert"
        Case "dialog.rebuildButton": LocalizedEn = "Rebuild"
        Case "dialog.attach.caption": LocalizedEn = "Attach Package to Selected Image"
        Case "dialog.attach.intro": LocalizedEn = "Click Open Folder..., add files/folders in Windows Explorer, then return here and click Insert. The selected image keeps its size and position."
        Case "dialog.selectedImage": LocalizedEn = "Selected image in current document"
        Case "dialog.manage.caption": LocalizedEn = "Manage Selected OLE Package"
        Case "dialog.manage.intro": LocalizedEn = "Click Open Folder..., edit the unpacked package in Windows Explorer, then return here and click Rebuild."
        Case "dialog.selectedOle": LocalizedEn = "Selected OLE object in current document"
        Case "dialog.folderPending": LocalizedEn = "The package folder will be created when this window opens."
        Case "dialog.afterRebuild": LocalizedEn = "After editing in Explorer, return here and click Rebuild."
        Case "dialog.afterInsert": LocalizedEn = "After editing in Explorer, return here and click Insert."
        Case "dialog.afterChooseInsert": LocalizedEn = "After choosing an image and editing in Explorer, return here and click Insert."
        Case "dialog.summary": LocalizedEn = "Package folder:" & vbCrLf & "%1" & vbCrLf & "%2"
        Case "dialog.fileDialog.chooseImage": LocalizedEn = "Choose the image to display"
        Case "dialog.filter.images": LocalizedEn = "Image files"
        Case "dialog.filter.png": LocalizedEn = "PNG files"
        Case "dialog.filter.jpeg": LocalizedEn = "JPEG files"
        Case "dialog.filter.tiff": LocalizedEn = "TIFF files"
        Case "dialog.filter.all": LocalizedEn = "All files"

        Case "msg.clearFolder": LocalizedEn = "Clear all files from the package contents folder?"
        Case "msg.chooseImage": LocalizedEn = "Please choose an image."
        Case "msg.chooseSupportedImage": LocalizedEn = "Please choose a PNG, JPG, JPEG, TIF, or TIFF image."
        Case "msg.packageFolderEmpty": LocalizedEn = "Please put at least one support file or folder into the package contents folder."
        Case "msg.selectedOleGone": LocalizedEn = "The selected OLE object is no longer available."
        Case "msg.selectedImageGone": LocalizedEn = "The selected image is no longer available."
        Case "msg.selectExisting": LocalizedEn = "Please select an existing picture or OLE Packager object in the Word document first."
        Case "msg.folderMissing": LocalizedEn = "Folder does not exist: %1"
        Case "msg.plotMissingInFolder": LocalizedEn = "plot.png is not found: %1"
        Case "msg.plotFileMissing": LocalizedEn = "plot.png does not exist: %1"
        Case "msg.choosePlot": LocalizedEn = "Please choose plot.png."
        Case "msg.imageMissing": LocalizedEn = "Image file does not exist: %1"
        Case "msg.rawDataMissing": LocalizedEn = "Raw data path does not exist: %1"
        Case "msg.selectPictureOnly": LocalizedEn = "Please select a picture, not another kind of floating object."
        Case "msg.noSupportFiles": LocalizedEn = "No support files or folders were selected."
        Case "picker.plotFolder": LocalizedEn = "Select the folder containing plot.png"
        Case "picker.plotImage": LocalizedEn = "Choose plot.png"
        Case "usage.title": LocalizedEn = "OLE Packager Usage"
        Case "usage.text": LocalizedEn = BuildUsageEn()
    End Select
End Function

Private Function LocalizedZh(ByVal key As String) As String
    Select Case key
        Case "app.name": LocalizedZh = "OLE Packager"
        Case "ribbon.group.insert": LocalizedZh = DecodeLoc("\u63D2\u5165")
        Case "ribbon.insertPackage.label": LocalizedZh = DecodeLoc("\u63D2\u5165 OLE \u56FE\u5305")
        Case "ribbon.insertPackage.screentip": LocalizedZh = DecodeLoc("\u63D2\u5165 plot.png \u548C\u6E90\u6587\u4EF6")
        Case "ribbon.insertPackage.supertip": LocalizedZh = DecodeLoc("\u9009\u62E9\u5305\u542B plot.png \u7684\u6587\u4EF6\u5939\uFF1B\u63D2\u4EF6\u4F1A\u628A\u6574\u4E2A\u6587\u4EF6\u5939\u6253\u5305\u4E3A zip OLE \u5BF9\u8C61\uFF0C\u5E76\u5728 Word \u4E2D\u663E\u793A plot.png\u3002")
        Case "ribbon.insertImageFiles.label": LocalizedZh = DecodeLoc("\u63D2\u5165\u56FE\u7247 + \u6587\u4EF6")
        Case "ribbon.insertImageFiles.screentip": LocalizedZh = DecodeLoc("\u63D2\u5165\u56FE\u7247\u5E76\u9644\u52A0\u652F\u6301\u6587\u4EF6")
        Case "ribbon.insertImageFiles.supertip": LocalizedZh = DecodeLoc("\u6253\u5F00\u4E00\u4E2A\u6587\u4EF6\u5939\uFF0C\u5728 Explorer \u4E2D\u7F16\u8F91\u5305\u5185\u5BB9\uFF0C\u7136\u540E\u5C06\u5176\u4F5C\u4E3A zip OLE \u5BF9\u8C61\u5D4C\u5165\uFF0C\u5E76\u663E\u793A\u4E3A\u6240\u9009\u56FE\u7247\u3002")
        Case "ribbon.manage.label": LocalizedZh = DecodeLoc("\u7BA1\u7406\u56FE\u7247/OLE \u6587\u4EF6")
        Case "ribbon.manage.screentip": LocalizedZh = DecodeLoc("\u7BA1\u7406\u9644\u52A0\u5230\u56FE\u7247\u6216 OLE \u5BF9\u8C61\u7684\u6587\u4EF6")
        Case "ribbon.manage.supertip": LocalizedZh = DecodeLoc("\u9009\u4E2D\u5DF2\u6709\u56FE\u7247\u6216 OLE Packager \u5BF9\u8C61\uFF0C\u589E\u5220\u5D4C\u5165\u6587\u4EF6\uFF0C\u540C\u65F6\u4FDD\u6301\u663E\u793A\u5916\u89C2\u3001\u5927\u5C0F\u548C\u4F4D\u7F6E\u3002")
        Case "ribbon.usage.label": LocalizedZh = DecodeLoc("\u7528\u6CD5")
        Case "ribbon.usage.screentip": LocalizedZh = DecodeLoc("\u663E\u793A\u7528\u6CD5\u8BF4\u660E")
        Case "ribbon.usage.supertip": LocalizedZh = DecodeLoc("\u4E86\u89E3\u5982\u4F55\u51C6\u5907 plot.py\u3001plot.png \u548C\u5305\u5185\u5BB9\uFF0C\u7136\u540E\u63D2\u5165 OLE \u56FE\u5305\u3002")

        Case "dialog.insert.caption": LocalizedZh = DecodeLoc("\u63D2\u5165\u56FE\u7247 + \u5305\u5185\u5BB9\u6587\u4EF6\u5939")
        Case "dialog.insert.intro": LocalizedZh = DecodeLoc("1. \u9009\u62E9\u663E\u793A\u56FE\u7247\u30022. \u70B9\u51FB\u6253\u5F00\u6587\u4EF6\u5939\uFF0C\u5728 Explorer \u4E2D\u7F16\u8F91\u5305\u5185\u5BB9\u30023. \u56DE\u5230\u8FD9\u91CC\u70B9\u51FB\u63D2\u5165\u3002")
        Case "dialog.displayImage": LocalizedZh = DecodeLoc("\u663E\u793A\u56FE\u7247")
        Case "dialog.chooseImage": LocalizedZh = DecodeLoc("\u9009\u62E9\u56FE\u7247...")
        Case "dialog.packageFolder": LocalizedZh = DecodeLoc("\u5305\u5185\u5BB9\u6587\u4EF6\u5939")
        Case "dialog.openFolder": LocalizedZh = DecodeLoc("\u6253\u5F00\u6587\u4EF6\u5939...")
        Case "dialog.clearFolder": LocalizedZh = DecodeLoc("\u6E05\u7A7A\u6587\u4EF6\u5939")
        Case "dialog.insertButton": LocalizedZh = DecodeLoc("\u63D2\u5165")
        Case "dialog.rebuildButton": LocalizedZh = DecodeLoc("\u91CD\u5EFA")
        Case "dialog.attach.caption": LocalizedZh = DecodeLoc("\u7ED9\u6240\u9009\u56FE\u7247\u9644\u52A0\u5305")
        Case "dialog.attach.intro": LocalizedZh = DecodeLoc("\u70B9\u51FB\u6253\u5F00\u6587\u4EF6\u5939\uFF0C\u5728 Explorer \u4E2D\u6DFB\u52A0\u6587\u4EF6/\u6587\u4EF6\u5939\uFF0C\u7136\u540E\u56DE\u5230\u8FD9\u91CC\u70B9\u51FB\u63D2\u5165\u3002\u6240\u9009\u56FE\u7247\u4F1A\u4FDD\u6301\u539F\u5927\u5C0F\u548C\u4F4D\u7F6E\u3002")
        Case "dialog.selectedImage": LocalizedZh = DecodeLoc("\u5F53\u524D\u6587\u6863\u4E2D\u7684\u6240\u9009\u56FE\u7247")
        Case "dialog.manage.caption": LocalizedZh = DecodeLoc("\u7BA1\u7406\u6240\u9009 OLE \u5305")
        Case "dialog.manage.intro": LocalizedZh = DecodeLoc("\u70B9\u51FB\u6253\u5F00\u6587\u4EF6\u5939\uFF0C\u5728 Explorer \u4E2D\u7F16\u8F91\u89E3\u538B\u540E\u7684\u5305\u5185\u5BB9\uFF0C\u7136\u540E\u56DE\u5230\u8FD9\u91CC\u70B9\u51FB\u91CD\u5EFA\u3002")
        Case "dialog.selectedOle": LocalizedZh = DecodeLoc("\u5F53\u524D\u6587\u6863\u4E2D\u7684\u6240\u9009 OLE \u5BF9\u8C61")
        Case "dialog.folderPending": LocalizedZh = DecodeLoc("\u7A97\u53E3\u6253\u5F00\u540E\u4F1A\u521B\u5EFA\u5305\u5185\u5BB9\u6587\u4EF6\u5939\u3002")
        Case "dialog.afterRebuild": LocalizedZh = DecodeLoc("\u5728 Explorer \u4E2D\u7F16\u8F91\u540E\uFF0C\u56DE\u5230\u8FD9\u91CC\u70B9\u51FB\u91CD\u5EFA\u3002")
        Case "dialog.afterInsert": LocalizedZh = DecodeLoc("\u5728 Explorer \u4E2D\u7F16\u8F91\u540E\uFF0C\u56DE\u5230\u8FD9\u91CC\u70B9\u51FB\u63D2\u5165\u3002")
        Case "dialog.afterChooseInsert": LocalizedZh = DecodeLoc("\u9009\u62E9\u56FE\u7247\u5E76\u5728 Explorer \u4E2D\u7F16\u8F91\u540E\uFF0C\u56DE\u5230\u8FD9\u91CC\u70B9\u51FB\u63D2\u5165\u3002")
        Case "dialog.summary": LocalizedZh = DecodeLoc("\u5305\u5185\u5BB9\u6587\u4EF6\u5939\uFF1A") & vbCrLf & "%1" & vbCrLf & "%2"
        Case "dialog.fileDialog.chooseImage": LocalizedZh = DecodeLoc("\u9009\u62E9\u8981\u663E\u793A\u7684\u56FE\u7247")
        Case "dialog.filter.images": LocalizedZh = DecodeLoc("\u56FE\u7247\u6587\u4EF6")
        Case "dialog.filter.png": LocalizedZh = "PNG " & DecodeLoc("\u6587\u4EF6")
        Case "dialog.filter.jpeg": LocalizedZh = "JPEG " & DecodeLoc("\u6587\u4EF6")
        Case "dialog.filter.tiff": LocalizedZh = "TIFF " & DecodeLoc("\u6587\u4EF6")
        Case "dialog.filter.all": LocalizedZh = DecodeLoc("\u6240\u6709\u6587\u4EF6")

        Case "msg.clearFolder": LocalizedZh = DecodeLoc("\u8981\u6E05\u7A7A\u5305\u5185\u5BB9\u6587\u4EF6\u5939\u4E2D\u7684\u6240\u6709\u6587\u4EF6\u5417\uFF1F")
        Case "msg.chooseImage": LocalizedZh = DecodeLoc("\u8BF7\u9009\u62E9\u4E00\u5F20\u56FE\u7247\u3002")
        Case "msg.chooseSupportedImage": LocalizedZh = DecodeLoc("\u8BF7\u9009\u62E9 PNG\u3001JPG\u3001JPEG\u3001TIF \u6216 TIFF \u56FE\u7247\u3002")
        Case "msg.packageFolderEmpty": LocalizedZh = DecodeLoc("\u8BF7\u628A\u81F3\u5C11\u4E00\u4E2A\u652F\u6301\u6587\u4EF6\u6216\u6587\u4EF6\u5939\u653E\u5165\u5305\u5185\u5BB9\u6587\u4EF6\u5939\u3002")
        Case "msg.selectedOleGone": LocalizedZh = DecodeLoc("\u6240\u9009 OLE \u5BF9\u8C61\u5DF2\u4E0D\u53EF\u7528\u3002")
        Case "msg.selectedImageGone": LocalizedZh = DecodeLoc("\u6240\u9009\u56FE\u7247\u5DF2\u4E0D\u53EF\u7528\u3002")
        Case "msg.selectExisting": LocalizedZh = DecodeLoc("\u8BF7\u5148\u5728 Word \u6587\u6863\u4E2D\u9009\u4E2D\u4E00\u5F20\u5DF2\u6709\u56FE\u7247\u6216 OLE Packager \u5BF9\u8C61\u3002")
        Case "msg.folderMissing": LocalizedZh = DecodeLoc("\u6587\u4EF6\u5939\u4E0D\u5B58\u5728\uFF1A%1")
        Case "msg.plotMissingInFolder": LocalizedZh = DecodeLoc("\u5728\u6587\u4EF6\u5939\u4E2D\u627E\u4E0D\u5230 plot.png\uFF1A%1")
        Case "msg.plotFileMissing": LocalizedZh = DecodeLoc("plot.png \u4E0D\u5B58\u5728\uFF1A%1")
        Case "msg.choosePlot": LocalizedZh = DecodeLoc("\u8BF7\u9009\u62E9 plot.png\u3002")
        Case "msg.imageMissing": LocalizedZh = DecodeLoc("\u56FE\u7247\u6587\u4EF6\u4E0D\u5B58\u5728\uFF1A%1")
        Case "msg.rawDataMissing": LocalizedZh = DecodeLoc("\u539F\u59CB\u6570\u636E\u8DEF\u5F84\u4E0D\u5B58\u5728\uFF1A%1")
        Case "msg.selectPictureOnly": LocalizedZh = DecodeLoc("\u8BF7\u9009\u62E9\u56FE\u7247\uFF0C\u800C\u4E0D\u662F\u5176\u4ED6\u7C7B\u578B\u7684\u6D6E\u52A8\u5BF9\u8C61\u3002")
        Case "msg.noSupportFiles": LocalizedZh = DecodeLoc("\u672A\u9009\u62E9\u652F\u6301\u6587\u4EF6\u6216\u6587\u4EF6\u5939\u3002")
        Case "picker.plotFolder": LocalizedZh = DecodeLoc("\u9009\u62E9\u5305\u542B plot.png \u7684\u6587\u4EF6\u5939")
        Case "picker.plotImage": LocalizedZh = DecodeLoc("\u9009\u62E9 plot.png")
        Case "usage.title": LocalizedZh = DecodeLoc("OLE Packager \u4F7F\u7528\u8BF4\u660E")
        Case "usage.text": LocalizedZh = BuildUsageZh()
    End Select
End Function

Private Function BuildUsageEn() As String
    BuildUsageEn = _
        "OLE Packager Usage" & vbCrLf & vbCrLf & _
        "Mode 1: Plot folder" & vbCrLf & _
        "1. Prepare plot.py, raw data, helper scripts, and images in one folder." & vbCrLf & _
        "2. Run plot.py to generate plot.png in the same folder." & vbCrLf & _
        "3. In Word, click OLE Packager > Insert OLE Package and choose that folder." & vbCrLf & _
        "4. The add-in embeds the folder contents as a zip OLE object and displays plot.png." & vbCrLf & _
        "5. Top-level plot.png, plot.svg, and plot.pdf are excluded from the zip; plot.py and other files are preserved." & vbCrLf & vbCrLf & _
        "Mode 2: Image + package folder" & vbCrLf & _
        "1. Click OLE Packager > Insert Image + Files." & vbCrLf & _
        "2. Choose the display image." & vbCrLf & _
        "3. Click Open Folder..., edit package contents in Explorer, then return to Word and click Insert." & vbCrLf & _
        "4. The package folder contents become the zip root, and the selected image is used as the display." & vbCrLf & vbCrLf & _
        "Mode 3: Manage an existing picture or OLE package" & vbCrLf & _
        "1. Select an existing picture or OLE Packager object in the document." & vbCrLf & _
        "2. Click OLE Packager > Manage Image/OLE Files." & vbCrLf & _
        "3. Existing OLE packages are unpacked into a temporary package folder; pictures start with an empty package folder." & vbCrLf & _
        "4. Edit files in Explorer, then click Insert or Rebuild. The original display, size, and position are preserved."
End Function

Private Function BuildUsageZh() As String
    BuildUsageZh = DecodeLoc( _
        "OLE Packager \u4F7F\u7528\u8BF4\u660E\n\n" & _
        "\u65B9\u5F0F\u4E00\uFF1APlot \u6587\u4EF6\u5939\n" & _
        "1. \u5728\u4E00\u4E2A\u6587\u4EF6\u5939\u4E2D\u51C6\u5907 plot.py\u3001\u539F\u59CB\u6570\u636E\u3001\u8F85\u52A9\u811A\u672C\u548C\u56FE\u7247\u3002\n" & _
        "2. \u8FD0\u884C plot.py \u751F\u6210 plot.png\uFF0C\u5E76\u4E0E plot.py \u653E\u5728\u540C\u4E00\u6587\u4EF6\u5939\u3002\n" & _
        "3. \u5728 Word \u4E2D\u70B9\u51FB OLE Packager > \u63D2\u5165 OLE \u56FE\u5305\uFF0C\u9009\u62E9\u8BE5\u6587\u4EF6\u5939\u3002\n" & _
        "4. \u63D2\u4EF6\u4F1A\u628A\u6587\u4EF6\u5939\u5185\u5BB9\u6253\u5305\u4E3A zip OLE \u5BF9\u8C61\uFF0C\u5E76\u663E\u793A plot.png\u3002\n" & _
        "5. \u9876\u5C42 plot.png\u3001plot.svg\u3001plot.pdf \u4E0D\u4F1A\u8FDB\u5165 zip\uFF1Bplot.py \u548C\u5176\u4ED6\u6587\u4EF6\u4F1A\u4FDD\u7559\u3002\n\n" & _
        "\u65B9\u5F0F\u4E8C\uFF1A\u56FE\u7247 + \u5305\u5185\u5BB9\u6587\u4EF6\u5939\n" & _
        "1. \u70B9\u51FB OLE Packager > \u63D2\u5165\u56FE\u7247 + \u6587\u4EF6\u3002\n" & _
        "2. \u9009\u62E9\u8981\u663E\u793A\u7684\u56FE\u7247\u3002\n" & _
        "3. \u70B9\u51FB\u6253\u5F00\u6587\u4EF6\u5939\uFF0C\u5728 Explorer \u4E2D\u7F16\u8F91\u5305\u5185\u5BB9\uFF0C\u7136\u540E\u56DE\u5230 Word \u70B9\u51FB\u63D2\u5165\u3002\n" & _
        "4. \u5305\u5185\u5BB9\u6587\u4EF6\u5939\u7684\u9876\u5C42\u5185\u5BB9\u4F1A\u6210\u4E3A zip \u6839\u76EE\u5F55\uFF0C\u6240\u9009\u56FE\u7247\u4F5C\u4E3A\u663E\u793A\u5916\u89C2\u3002\n\n" & _
        "\u65B9\u5F0F\u4E09\uFF1A\u7BA1\u7406\u5DF2\u6709\u56FE\u7247\u6216 OLE \u56FE\u5305\n" & _
        "1. \u5728\u6587\u6863\u4E2D\u9009\u4E2D\u5DF2\u6709\u56FE\u7247\u6216 OLE Packager \u5BF9\u8C61\u3002\n" & _
        "2. \u70B9\u51FB OLE Packager > \u7BA1\u7406\u56FE\u7247/OLE \u6587\u4EF6\u3002\n" & _
        "3. \u5DF2\u6709 OLE \u56FE\u5305\u4F1A\u89E3\u538B\u5230\u4E34\u65F6\u5305\u5185\u5BB9\u6587\u4EF6\u5939\uFF1B\u666E\u901A\u56FE\u7247\u4F1A\u4F7F\u7528\u7A7A\u6587\u4EF6\u5939\u3002\n" & _
        "4. \u5728 Explorer \u4E2D\u7F16\u8F91\u6587\u4EF6\u540E\uFF0C\u70B9\u51FB\u63D2\u5165\u6216\u91CD\u5EFA\u3002\u539F\u663E\u793A\u5916\u89C2\u3001\u5927\u5C0F\u548C\u4F4D\u7F6E\u4F1A\u4FDD\u6301\u4E0D\u53D8\u3002")
End Function

Private Function DecodeLoc(ByVal escaped As String) As String
    Dim result As String
    Dim i As Long
    Dim codeText As String
    Dim code As Long

    i = 1
    Do While i <= Len(escaped)
        If Mid$(escaped, i, 2) = "\n" Then
            result = result & vbCrLf
            i = i + 2
        ElseIf Mid$(escaped, i, 2) = "\u" And i + 5 <= Len(escaped) Then
            codeText = Mid$(escaped, i + 2, 4)
            code = CLng("&H" & codeText)
            result = result & LocUnicodeChar(code)
            i = i + 6
        Else
            result = result & Mid$(escaped, i, 1)
            i = i + 1
        End If
    Loop

    DecodeLoc = result
End Function

Private Function LocUnicodeChar(ByVal code As Long) As String
    If code > 32767 Then code = code - 65536
    LocUnicodeChar = ChrW$(code)
End Function

