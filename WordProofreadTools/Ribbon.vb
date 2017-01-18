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
    Private editTabVisible As Boolean = False
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
    ''' 设置普通模式（非网站启动）
    ''' </summary>
    Public Sub SetNormalState()
        Me.editTabVisible = True
        Me.checkingTabVisible = False
        Me.editUploadEnable = False
        Me.buildinVisible = True
        If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = False
        End If
        Me.ribbon.Invalidate()
    End Sub

    ''' <summary>
    ''' 设置编辑模式（从网站启动）
    ''' </summary>
    Public Sub SetEditorState()
        Me.editTabVisible = True
        Me.checkingTabVisible = False
        Me.editUploadEnable = True
        Me.buildinVisible = True
        If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = False
        End If
        Me.ribbon.Invalidate()
    End Sub

    ''' <summary>
    ''' 设置审核模式
    ''' </summary>
    Public Sub SetAuditState()
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
            Globals.ThisAddIn.Application.UserInitials = CommonModule.nickName
        End If
        Me.ribbon.Invalidate()
    End Sub

    ''' <summary>
    ''' 设置校对模式
    ''' </summary>
    Public Sub SetProofreadState()
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
            Globals.ThisAddIn.Application.UserInitials = CommonModule.nickName
        End If
        Me.ribbon.Invalidate()
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
    ''' 快速批注
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnChecking_Click(ByVal control As Office.IRibbonControl)
        Dim app = Globals.ThisAddIn.Application
        Dim sel = app.Selection()
        Dim label = control.Tag
        Dim range = sel.Range()
        Dim comment = app.ActiveDocument.Comments.Add(range, label + "： ")
        'comment.Author = CommonModule.nickName

    End Sub

    ''' <summary>
    ''' 上传文件
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonUpload_Click(ByVal control As Office.IRibbonControl)
        If (Not String.IsNullOrEmpty(CommonModule.taskId)) Then
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            ' 保存文档
            doc.Save()

            CommonModule.localFile = doc.Path + "\" + doc.Name
            ' 获取文件大小
            Dim info = My.Computer.FileSystem.GetFileInfo(CommonModule.localFile)
            ' 大于100M
            If (info.Length > 100 * 100 * 1024) Then
                CommonModule.ShowAlert("上传文件不能大于100M！", "Warning")
                Return
            End If

            ' 开始上传任务
            Dim upload = New FormUpload()
            'upload.StartUpload(CommonModule.localFile)
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
                Try
                    My.Computer.FileSystem.WriteAllText(fileName, content, False)
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
        Dim dialog = New SaveFileDialog()
        dialog.Filter = "Word文件|*.docx"
        dialog.FileName = "导出Word文件.docx"
        dialog.Title = "请选择导出的Word文件"
        '确定后开始导出
        If dialog.ShowDialog() = DialogResult.OK Then
            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(dialog.FileName, WdSaveFormat.wdFormatDocument)
        End If
    End Sub

    ''' <summary>
    ''' 导出XML
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnExportXML_Click(ByVal control As Office.IRibbonControl)
        Dim dialog = New SaveFileDialog()
        dialog.Filter = "XML文件|*.xml"
        dialog.FileName = "导出XML文件.xml"
        dialog.Title = "请选择导出的XML文件"
        '确定后开始导出
        If dialog.ShowDialog() = DialogResult.OK Then
            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(dialog.FileName, WdSaveFormat.wdFormatXML)
        End If
    End Sub

    ''' <summary>
    ''' 导出PDF
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnExportPDF_Click(ByVal control As Office.IRibbonControl)
        Dim dialog = New SaveFileDialog()
        dialog.Filter = "PDF文件|*.pdf"
        dialog.FileName = "导出PDF文件.pdf"
        dialog.Title = "请选择导出的PDF文件"
        '确定后开始导出
        If dialog.ShowDialog() = DialogResult.OK Then
            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(dialog.FileName, WdSaveFormat.wdFormatPDF)
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
        End If
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
