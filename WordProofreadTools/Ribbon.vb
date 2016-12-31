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

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI
    Private WithEvents loginForm As FormLogin = CommonModule.loginForm

    Public Sub New()
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
    Private editDownloadEnable As Boolean = False
    Public Function GetEditUploadEnable(control As IRibbonControl) As Boolean
        Return editDownloadEnable
    End Function

    ''' <summary>
    ''' 设置普通模式（非网站启动）
    ''' </summary>
    Public Sub SetNormalState()

    End Sub

    ''' <summary>
    ''' 设置编辑模式（从网站启动）
    ''' </summary>
    Public Sub SetEditorState()

    End Sub

    ''' <summary>
    ''' 设置审核模式
    ''' </summary>
    Public Sub SetAuditState()

    End Sub

    Public Sub SetProofreadState()

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
        End Select

        Return Nothing
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
        Dim author = control.Id.Replace("BtnChecking_", "") + "主编"
        comment.Author = author

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
