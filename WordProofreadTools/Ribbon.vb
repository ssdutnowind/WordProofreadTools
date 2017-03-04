'TODO:   按照以下步骤启用功能区(XML)项: 

'1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon()
'End Function

'2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
'   操作(例如单击按钮)。注意: 如果已经从功能区设计器中导出此功能区，
'   请将代码从事件处理程序移动到回调方法，并
'   修改该代码以使用功能区扩展性(RibbonX)编程模型。

'3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。

'有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。

Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
        Me.editTabVisible = True
        Me.checkingTabVisible = True
        CommonModule.ribbon = Me
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("WordProofreadTools.Ribbon.xml")
    End Function

#Region "可用性回调"

    ''' <summary>
    ''' 控制内置Ribbon Tab是否可见
    ''' </summary>
    Private buildinVisible As Boolean = True
    Public Function GetBuildinVisible(control As IRibbonControl) As Boolean
        Return buildinVisible
    End Function

    ''' <summary>
    ''' 编辑Tab是否可见
    ''' </summary>
    Private editTabVisible As Boolean = True
    Public Function GetEditTabVisible(control As IRibbonControl) As Boolean
        Return editTabVisible
    End Function

    ''' <summary>
    ''' 审校Tab是否课件
    ''' </summary>
    Private checkingTabVisible As Boolean = False
    Public Function GetCheckingTabVisible(control As IRibbonControl) As Boolean
        Return checkingTabVisible
    End Function

    ''' <summary>
    ''' 编辑工具的上传按钮是否可用
    ''' </summary>
    Private editUploadEnable As Boolean = False
    Public Function GetEditUploadEnable(control As IRibbonControl) As Boolean
        Return editUploadEnable
    End Function

    ''' <summary>
    ''' 刷新Ribbon状态
    ''' </summary>
    Private Sub RefreshRibbon()
        'If (Not IsNothing(Me.ribbon)) Then
        Me.ribbon.Invalidate()
        'End If
    End Sub

    ''' <summary>
    ''' 设置普通模式（非网站启动）
    ''' </summary>
    Public Sub SetNormalState()
        CommonModule.Log("设置普通模式（非网站启动）")
        Me.editTabVisible = True
        Me.checkingTabVisible = False
        Me.editUploadEnable = False
        Me.buildinVisible = True
        If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = False
            Globals.ThisAddIn.Application.ActiveDocument.Save()
        End If
        Me.RefreshRibbon()
    End Sub

    ''' <summary>
    ''' 设置编辑模式（从网站启动）
    ''' </summary>
    Public Sub SetEditorState()
        CommonModule.Log("设置编辑模式（从网站启动）")
        Me.editTabVisible = True
        Me.checkingTabVisible = False
        Me.editUploadEnable = True
        Me.buildinVisible = True
        If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = False
            Globals.ThisAddIn.Application.ActiveDocument.Save()
        End If
        Me.RefreshRibbon()
    End Sub

    ''' <summary>
    ''' 设置审核模式
    ''' </summary>
    Public Sub SetAuditState()
        CommonModule.Log("设置审核模式")
        Me.editTabVisible = True
        Me.checkingTabVisible = True
        Me.editUploadEnable = False
        Me.buildinVisible = False
        If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            ' 启用跟踪修订
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = True
            ' 显示修订内容
            'doc.ActiveWindow.View.MarkupMode = WdRevisionsMode.wdBalloonRevisions
            ' 显示修订和批注
            'doc.ActiveWindow.View.ShowRevisionsAndComments = True
            ' 显示插入和删除
            'doc.ActiveWindow.View.ShowInsertionsAndDeletions = True
            ' 修改批注和修订作者
            CommonModule.Log("定义修订作者：" + CommonModule.nickName)
            Globals.ThisAddIn.Application.UserInitials = CommonModule.nickName
        End If
        Me.RefreshRibbon()
    End Sub

    ''' <summary>
    ''' 设置校对模式
    ''' </summary>
    Public Sub SetProofreadState()
        CommonModule.Log("设置校对模式")
        Me.editTabVisible = True
        Me.checkingTabVisible = True
        Me.editUploadEnable = False
        Me.buildinVisible = False
        If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            ' 启用跟踪修订
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = True
            ' 显示修订内容
            'doc.ActiveWindow.View.MarkupMode = WdRevisionsMode.wdBalloonRevisions
            ' 显示修订和批注
            'doc.ActiveWindow.View.ShowRevisionsAndComments = True
            ' 显示插入和删除
            'doc.ActiveWindow.View.ShowInsertionsAndDeletions = True
            '修改批注和修订作者
            CommonModule.Log("定义修订作者：" + CommonModule.nickName)
            Globals.ThisAddIn.Application.UserInitials = CommonModule.nickName
        End If
        Me.RefreshRibbon()
    End Sub

#End Region

#Region "图标"

    ''' <summary>
    ''' 获取图标
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function GetIcons(ByVal control As Office.IRibbonControl) As Bitmap

        Select Case control.Id
            Case "BtnUpload1"
                Return New Bitmap(My.Resources.btn_upload)
            Case "BtnUpload2"
                Return New Bitmap(My.Resources.btn_upload)
            Case "BtnExportContents"
                Return New Bitmap(My.Resources.btn_contents)
            Case "BtnExportIndex"
                Return New Bitmap(My.Resources.btn_index)
            Case "BtnExportAbstract"
                Return New Bitmap(My.Resources.btn_abstract)
            Case "BtnRemoveComment"
                Return New Bitmap(My.Resources.btn_remove_comment)
            Case "BtnExportDoc"
                Return New Bitmap(My.Resources.btn_word)
            Case "BtnExportXML"
                Return New Bitmap(My.Resources.btn_xml)
            Case "BtnExportPDF"
                Return New Bitmap(My.Resources.btn_pdf)
            Case "BtnFormatThousands"
                Return New Bitmap(My.Resources.btn_format_thousands)
            Case "BtnFormatNumber"
                Return New Bitmap(My.Resources.btn_format_number1)
            Case "BtnFormatNumber1"
                Return New Bitmap(My.Resources.btn_format_number1)
            Case "BtnFormatNumber2"
                Return New Bitmap(My.Resources.btn_format_number2)
            Case "BtnFormatNumber3"
                Return New Bitmap(My.Resources.btn_format_number3)
            Case "BtnFormatPercent"
                Return New Bitmap(My.Resources.btn_format_percent)
            Case "BtnInsertUnit"
                Return New Bitmap(My.Resources.btn_format_units)
            Case "BtnFormatMoney"
                Return New Bitmap(My.Resources.btn_money_cny)
            Case "BtnFormatMoneyCny"
                Return New Bitmap(My.Resources.btn_money_cny)
            Case "BtnFormatMoneyUsd"
                Return New Bitmap(My.Resources.btn_money_usd)
            Case "BtnFormatMoneyEur"
                Return New Bitmap(My.Resources.btn_money_eur)
            Case "BtnFormatMoneyGbp"
                Return New Bitmap(My.Resources.btn_money_gbp)
            Case "BtnAbout"
                Return New Bitmap(My.Resources.btn_about)
            Case "BtnFileManager"
                Return New Bitmap(My.Resources.btn_file_manager)
            Case Else
                Return Nothing
        End Select

    End Function

    ''' <summary>
    ''' 获取快速批注图标
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function GetCheckingIcon(ByVal control As Office.IRibbonControl) As Bitmap
        Dim id As String = control.Id
        Dim icon As String = id.Replace("BtnChecking", "btn")
        Dim res = My.Resources.ResourceManager.GetObject(icon)
        If IsNothing(res) Then
            Return New Bitmap(My.Resources.btn_1_1)
        Else
            Return CType(res, Bitmap)
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 获取Label内容
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function GetLabel(ByVal control As Office.IRibbonControl) As String
        Select Case CommonModule.taskType
            Case "1"
                Return "初审"
            Case "2"
                Return "复审"
            Case "3"
                Return "终审"
            Case "5"
                Return "排版"
            Case "6"
                Return "一校"
            Case "7"
                Return "二校"
            Case "8"
                Return "三校"
            Case Else
                Return "任务管理"
        End Select
    End Function
#End Region

#Region "功能区回调"
    '在此创建回调方法。有关添加回调方法的详细信息，请访问 http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    ''' <summary>
    ''' 关于
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnAbout_Click(ByVal control As Office.IRibbonControl)
        CommonModule.Log("显示""关于""对话框")
        Dim dialog = New FormAbout()
        dialog.ShowDialog()
    End Sub

    ''' <summary>
    ''' 快速批注
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnChecking_Click(ByVal control As Office.IRibbonControl)
        Dim app = Globals.ThisAddIn.Application
        Dim sel = app.Selection()
        Dim label = control.Tag
        Dim range = sel.Range()
        Dim comment = app.ActiveDocument.Comments.Add(range, label + "： ")
        CommonModule.Log("快速批注：" + label)

    End Sub

    ''' <summary>
    ''' 上传文件
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonUpload_Click(ByVal control As Office.IRibbonControl)
        If (Not String.IsNullOrEmpty(CommonModule.taskId)) Then
            CommonModule.Log("准备上传任务文件……")

            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            ' 保存文档
            doc.Save()
            CommonModule.Log("1.自动保存")

            CommonModule.localFile = doc.Path + "\" + doc.Name
            CommonModule.Log("2.目标文件路径：" + CommonModule.localFile)
            ' 获取文件大小
            Dim info = My.Computer.FileSystem.GetFileInfo(CommonModule.localFile)
            ' 大于100M
            If (info.Length > 100 * 100 * 1024) Then
                CommonModule.ShowAlert("上传文件不能大于100M！", "Warning")
                Return
            End If
            CommonModule.Log("3.文件大小校验通过")
            CommonModule.Log("4.打开""文件上传""对话框")
            ' 开始上传任务
            Dim upload = New FormUpload()
            upload.ShowDialog()

        End If
    End Sub

    ''' <summary>
    ''' 导出目录
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonExportContents_Click(ByVal control As Office.IRibbonControl)
        Dim contentList = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents
        If contentList.Count > 0 Then
            Dim dialog = New SaveFileDialog()
            dialog.Filter = "文本文件|*.txt"
            dialog.FileName = "导出文档目录.txt"
            dialog.Title = "请选择导出的目录文件"
            '确定后开始导出
            If dialog.ShowDialog() = DialogResult.OK Then
                '只获取第一个目录
                Dim contents = contentList(1)
                Dim fileName = dialog.FileName
                Dim content = contents.Range.Text
                ' 拆分之后trim掉末尾的[CR]，最后再用[CRLF]拼接
                ' 将每行中最后一个Tab替换为“......”最后一个Tab为目录内容与页码之间的空白
                Dim contentArray = content.Split(vbCr)
                For i = 0 To contentArray.Length - 1
                    contentArray(i) = contentArray(i).Trim()
                    If (contentArray(i).Length > 2) Then
                        contentArray(i) = contentArray(i).Substring(0, contentArray(i).LastIndexOf(vbTab)) + "......" + contentArray(i).Substring(contentArray(i).LastIndexOf(vbTab) + 1)
                    End If
                Next
                content = String.Join(vbCrLf, contentArray)
                Try
                    My.Computer.FileSystem.WriteAllText(fileName, content, False)
                    CommonModule.Log("导出目录……")
                    CommonModule.ShowAlert("导出目录成功！")
                Catch ex As Exception
                    CommonModule.ShowAlert(ex.Message, "Error")
                End Try
            End If
        Else
            CommonModule.ShowAlert("当前文档内没有目录，请先插入目录后再导出！", "Warning")
        End If
    End Sub

    ''' <summary>
    ''' 导出索引
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonExportIndex_Click(ByVal control As Office.IRibbonControl)
        Dim indexs = Globals.ThisAddIn.Application.ActiveDocument.Indexes
        If indexs.Count > 0 Then
            Dim dialog = New SaveFileDialog()
            dialog.Filter = "文本文件|*.txt"
            dialog.FileName = "导出文档索引.txt"
            dialog.Title = "请选择导出的索引文件"
            '确定后开始导出
            If dialog.ShowDialog() = DialogResult.OK Then
                Dim index = indexs(1)
                Dim fileName = dialog.FileName
                Dim content = index.Range.Text
                Try
                    My.Computer.FileSystem.WriteAllText(fileName, content, False)
                    CommonModule.Log("导出索引……")
                    CommonModule.ShowAlert("导出索引成功！")
                Catch ex As Exception
                    CommonModule.ShowAlert(ex.Message, "Error")
                End Try
            End If
        Else
            CommonModule.ShowAlert("当前文档内没有索引，请先插入索引后再导出！", "Warning")
        End If
    End Sub

    ''' <summary>
    ''' 导出摘要
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnExportAbstract_Click(ByVal control As Office.IRibbonControl)

    End Sub

    ''' <summary>
    ''' 导出Doc
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnExportDoc_Click(ByVal control As Office.IRibbonControl)
        Dim fileName = Globals.ThisAddIn.Application.ActiveDocument.Name
        If (fileName.IndexOf(".") > 0) Then
            fileName = fileName.Substring(0, fileName.LastIndexOf(".")) + ".docx"
        Else
            fileName += ".docx"
        End If

        Dim dialog = New SaveFileDialog()
        dialog.Filter = "Word文件|*.docx"
        dialog.FileName = fileName
        dialog.Title = "请选择导出的Word文件"
        '确定后开始导出
        If dialog.ShowDialog() = DialogResult.OK Then
            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(dialog.FileName, WdSaveFormat.wdFormatDocumentDefault)
            CommonModule.Log("另存为Word……")
        End If
    End Sub

    ''' <summary>
    ''' 导出XML
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnExportXML_Click(ByVal control As Office.IRibbonControl)
        Dim fileName = Globals.ThisAddIn.Application.ActiveDocument.Name
        If (fileName.IndexOf(".") > 0) Then
            fileName = fileName.Substring(0, fileName.LastIndexOf(".")) + ".xml"
        Else
            fileName += ".xml"
        End If

        Dim dialog = New SaveFileDialog()
        dialog.Filter = "XML文件|*.xml"
        dialog.FileName = fileName
        dialog.Title = "请选择导出的XML文件"
        '确定后开始导出
        If dialog.ShowDialog() = DialogResult.OK Then
            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(dialog.FileName, WdSaveFormat.wdFormatXML)
            CommonModule.Log("另存为XML……")
        End If
    End Sub

    ''' <summary>
    ''' 导出PDF
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnExportPDF_Click(ByVal control As Office.IRibbonControl)
        Dim fileName = Globals.ThisAddIn.Application.ActiveDocument.Name
        If (fileName.IndexOf(".") > 0) Then
            fileName = fileName.Substring(0, fileName.LastIndexOf(".")) + ".pdf"
        Else
            fileName += ".pdf"
        End If

        Dim dialog = New SaveFileDialog()
        dialog.Filter = "PDF文件|*.pdf"
        dialog.FileName = fileName
        dialog.Title = "请选择导出的PDF文件"
        '确定后开始导出
        If dialog.ShowDialog() = DialogResult.OK Then
            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(dialog.FileName, WdSaveFormat.wdFormatPDF)
            CommonModule.Log("另存为PDF……")
        End If
    End Sub

    ''' <summary>
    ''' 清空批注
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonRemoveComment_Click(ByVal control As Office.IRibbonControl)
        If CommonModule.ShowAlert("确认要删除所有批注么？", "Question") = vbYes Then
            Dim comments = Globals.ThisAddIn.Application.ActiveDocument.Comments
            For Each comment As Microsoft.Office.Interop.Word.Comment In comments
                comment.DeleteRecursively()
            Next
            CommonModule.Log("清空批注")
            CommonModule.ShowAlert("所有批注已删除完毕！")
        End If
    End Sub

    ''' <summary>
    ''' 打开文件管理窗口
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnFileManager_Click(ByVal control As Office.IRibbonControl)
        Dim form = New FormLocalFiles()
        form.ShowDialog()
    End Sub

#End Region

#Region "格式化回调"

    ''' <summary>
    ''' 千分位格式化
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonFormatThousands_Click(ByVal control As Office.IRibbonControl)
        Dim app = Globals.ThisAddIn.Application
        Dim range As Word.Range = app.Selection.Range
        Dim findObject As Word.Find = app.Selection.Find

        ' 没选中任何文本
        If (String.IsNullOrEmpty(range.Text)) Then
            Return
        End If

        Dim FindChar = "([0-9])([0-9]{3})"
        Dim RepChar = "\1,\2"

        'app.ScreenUpdating = False '关闭屏幕更新

        With findObject
            .ClearFormatting()
            .MatchWildcards = True
            .Text = FindChar
            .Replacement.ClearFormatting()
            .Replacement.Text = RepChar
            '.Execute(Replace:=Word.WdReplace.wdReplaceOne, Wrap:=Word.WdFindWrap.wdFindStop, Forward:=False)
            Do While (.Execute(Replace:=Word.WdReplace.wdReplaceOne, Forward:=False, Wrap:=Word.WdFindWrap.wdFindStop))

            Loop
            'Do While (.Execute(FindText:=FindChar, Replace:=Word.WdReplace.wdReplaceOne, ReplaceWith:=RepChar, Wrap:=Word.WdFindWrap.wdFindStop, Forward:=False))
            'Loop
        End With
        'range.Select()
        'app.ScreenUpdating = False '开启屏幕更新
    End Sub

    ''' <summary>
    ''' 百分比格式化
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonFormatPercent_Click(ByVal control As Office.IRibbonControl)

    End Sub

    ''' <summary>
    ''' 数字转换
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnFormatNumber_Click(ByVal control As Office.IRibbonControl)
        ' 没找到正则等快捷方法，只能笨招了
        Dim src As Array = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
            "１", "２", "３", "４", "５", "６", "７", "８", "９", "０",
        "一", "二", "三", "四", "五", "六", "七", "八", "九", "〇",
        "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖", "零"}
        Dim target As Array
        If (control.Id = "BtnFormatNumber2") Then
            ' 转为大写数字
            target = {"一", "二", "三", "四", "五", "六", "七", "八", "九", "〇"}
        ElseIf (control.Id = "BtnFormatNumber3") Then
            ' 转为汉字数字
            target = {"壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖", "零"}
        Else
            ' 转为阿拉伯数字
            target = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0"}
        End If

        Dim application = Globals.ThisAddIn.Application
        Dim range = application.Selection.Range
        Dim findObject As Word.Find = application.Selection.Find

        ' 没选中任何文本
        If (String.IsNullOrEmpty(range.Text)) Then
            Return
        End If

        CommonModule.Log("数字转换：" + control.Id)

        For i As Integer = 0 To src.Length - 1
            With findObject
                .ClearFormatting()
                .Text = src(i)
                .Replacement.ClearFormatting()
                .Replacement.Text = target(i Mod 10)
                .Execute(Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindStop)
            End With
        Next

    End Sub
#End Region

#Region "快速插入"

    ''' <summary>
    ''' 插入货币符号
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnInsertMoney_Click(ByVal control As Office.IRibbonControl)
        Dim text As String
        Select Case control.Id
            Case "BtnFormatMoneyUsd"
                text = "$"
                Exit Select
            Case "BtnFormatMoneyEur"
                text = "€"
                Exit Select
            Case "BtnFormatMoneyGbp"
                text = "£"
                Exit Select
            Case Else
                text = "￥"
        End Select
        CommonModule.Log("插入货币符号：" + control.Id)
        Dim app = Globals.ThisAddIn.Application
        Dim selection = app.Selection
        Dim range = selection.Range
        range.InsertAfter(text)
        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        range.Select()
    End Sub

    ''' <summary>
    ''' 插入计量单位
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnInsertUnion_Click(ByVal control As Office.IRibbonControl)
        CommonModule.Log("插入计量单位：" + control.Tag)
        Dim text As String = control.Tag
        Dim app = Globals.ThisAddIn.Application
        Dim selection = app.Selection
        Dim range = selection.Range
        range.InsertAfter(text)
        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        range.Select()
    End Sub

#End Region

#Region "事件"

    'Public Sub OnUserLogin(ByVal sender As Object) Handles loginForm.UserLogin
    '    CommonModule.Log("Ribbon：用户登录")
    '    ' 隐藏内置功能区
    '    Me.SetBuildinHidden()
    '    ' 登录
    '    Me.SetLoginEnabled(False)

    'End Sub

#End Region


#Region "帮助器"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
