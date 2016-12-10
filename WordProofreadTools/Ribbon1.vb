Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Private WithEvents loginForm As FormLogin = CommonModule.loginForm


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    ''' <summary>
    ''' 弹出登录窗口
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnRibbonLogin_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonLogin.Click
        CommonModule.loginForm.ShowDialog()
    End Sub


#Region "同步窗口显示"

    ''' <summary>
    ''' 切换云端窗格显示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CbxRobbinCloudFiles_Click(sender As Object, e As RibbonControlEventArgs) Handles CbxRobbinCloudFiles.Click
        If Me.CbxRobbinCloudFiles.Checked Then
            CommonModule.ShowCloudFilesPane()
        Else
            CommonModule.HideCloudFilesPane()
        End If
    End Sub

#End Region


#Region "导出功能"

    ''' <summary>
    ''' 导出目录
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnRibbonExportContents_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonExportContents.Click
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
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnRibbonExportIndex_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonExportIndex.Click
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
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnRibbonExprotAbstract_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonExprotAbstract.Click

        'Dim ctrls = Globals.ThisAddIn.Application.ActiveDocument.ContentControls
        'For Each ctrl In ctrls
        '    ctrl.
        'Next
        'Return


        'Dim properties = DirectCast(Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties, Microsoft.Office.Core.DocumentProperties)

        'Common.Log(properties.Item("Subject").Value)
        'Dim dialog = New SaveFileDialog()
        'dialog.Filter = "文本文件|*.txt"
        '    dialog.FileName = "导出文档索引.txt"
        '    dialog.Title = "请选择导出的索引文件"


    End Sub

    Private Sub BtnRibbonProofread_Click(sender As Object, e As RibbonControlEventArgs)
        Dim selection = Globals.ThisAddIn.Application.Selection
        If (selection.Text.Count <= 0) Then
            CommonModule.ShowAlert("请选择内容！", "Warning")
            Return
        End If

        Dim comment = Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, "AAAAA")

        comment.Author = "张主编"


        '' 图片形式拷贝到内存
        'selection.CopyAsPicture()
        '' 判断一下是否有图片
        'If Clipboard.ContainsImage() Then
        '    Dim image = Clipboard.GetData(DataFormats.Bitmap)
        'Else
        '    Common.ShowAlert("生成文档片段快照失败，请稍候重试！", "Error")
        'End If

        'Dim dialog = New FormProofread()
        'dialog.ShowDialog()
    End Sub

    Private Sub BtnRibbonTemplateManage_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonTemplateManage.Click
        Dim dialog = New FormTemplate()
        dialog.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim dialog = New FormTaskProperty()
        dialog.ShowDialog()
    End Sub

#End Region


#Region "用户事件"
    Public Sub OnUserLogin(ByVal sender As Object) Handles loginForm.UserLogin
        CommonModule.Log("Ribbon1：用户登录")
        'Me.StartFromScratch
        'Globals.ThisAddIn.Application.COMAddIns
    End Sub
#End Region

End Class
