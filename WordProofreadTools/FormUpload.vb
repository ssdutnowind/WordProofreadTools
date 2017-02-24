Imports System.ComponentModel
Imports System.Net

Public Class FormUpload

    Private WithEvents WC As New WebClient

    ''' <summary>
    ''' 关闭按钮，默认禁用
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' 上传文件到服务器
    ''' </summary>
    ''' <param name="file">文件名</param>
    Public Sub StartUpload(file As String)
        Try
            Dim url = CommonModule.serverPath + ""
            WC.QueryString.Add("taskId", CommonModule.taskId)

            WC.UploadFileAsync(New Uri(url), file)
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub

    ''' <summary>
    ''' 开始上传
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnUpload_Click(sender As Object, e As EventArgs) Handles BtnUpload.Click
        ' 检查备注
        Dim comment = Me.TxtComment.Text
        If (String.IsNullOrEmpty(comment)) Then
            CommonModule.ShowAlert("请输入备注！", "Warning")
            Return
        End If

        Try
            Me.Panel1.Hide()
            Me.BtnUpload.Hide()

            ' 当前Word进程无法直接再打开文件，所以拷贝一份
            Dim dir = My.Computer.FileSystem.GetParentPath(CommonModule.localFile)
            Dim tempFile = dir + "\" + Date.Now.ToString("yyyyMMddHHmmss") + ".docx"

            My.Computer.FileSystem.CopyFile(CommonModule.localFile, tempFile)
            CommonModule.Log("5.生成文件副本：" + vbCrLf + tempFile)
            CommonModule.Log("6.请求参数：")
            Dim url = CommonModule.serverPath + My.Settings.Item("UploadTaskUrl")
            ' 任务ID
            WC.QueryString.Add("taskId", CommonModule.taskId)
            CommonModule.Log("  taskId:" + CommonModule.taskId)
            ' 用户ID
            WC.QueryString.Add("userId", CommonModule.userId)
            CommonModule.Log("  userId:" + CommonModule.userId)
            ' 备注
            WC.QueryString.Add("comment", comment)
            CommonModule.Log("  comment:" + comment)
            ' 批注数量
            WC.QueryString.Add("commentNum", Globals.ThisAddIn.Application.ActiveDocument.Comments.Count)
            CommonModule.Log("  commentNum:" + Globals.ThisAddIn.Application.ActiveDocument.Comments.Count.ToString())
            ' 开始上传
            CommonModule.Log("开始上传……")
            WC.UploadFileAsync(New Uri(url), tempFile)
        Catch ex As Exception
            CommonModule.Log("上传出错：" + ex.ToString())
            CommonModule.ShowAlert(ex.ToString(), "Error")
        End Try
    End Sub

    ''' <summary>
    ''' 监听上传进度
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WC_UploadProgressChanged(ByVal sender As Object, ByVal e As UploadProgressChangedEventArgs) Handles WC.UploadProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    ''' <summary>
    ''' 监听下载完成
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WC_UploadFileComplate(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs) Handles WC.UploadFileCompleted
        If e.Cancelled Then
            ' 任务被取消
            CommonModule.Log("上传任务被取消")
            Me.Close()
        ElseIf e.Error IsNot Nothing Then
            Me.Hide()
            CommonModule.Log("上传任务文件出错！" + vbCrLf + e.Error.Message)
            CommonModule.ShowAlert("上传任务文件出错！" + vbCrLf + e.Error.Message, "Error")
            Me.Close()
        Else
            Me.Hide()
            CommonModule.Log("上传任务文件成功！")
            CommonModule.ShowAlert("上传任务文件成功！")
            Me.Close()
        End If
    End Sub
End Class