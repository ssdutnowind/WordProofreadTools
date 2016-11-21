Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    ''' <summary>
    ''' 弹出登录窗口
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnRibbonLogin_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonLogin.Click
        Dim login = New LoginForm()
        login.ShowDialog()
    End Sub

    ''' <summary>
    ''' 导出目录
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnRibbonExportContents_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonExportContents.Click
        Dim contentList = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents
        If contentList.Count > 0 Then
            Dim contents = contentList(1)
            Common.Log(contents.Range.Text)
        Else
            Common.ShowAlert("当前文档内没有目录！", "Warning")
        End If
    End Sub

    ''' <summary>
    ''' 切换云端窗格显示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CbxRobbinCloudFiles_Click(sender As Object, e As RibbonControlEventArgs) Handles CbxRobbinCloudFiles.Click
        If Me.CbxRobbinCloudFiles.Checked Then
            Common.ShowCloudFilesPane()
        Else
            Common.HideCloudFilesPane()
        End If
    End Sub



End Class
