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
    ''' 监听下载任务
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WC_UploadProgressChanged(ByVal sender As Object, ByVal e As UploadProgressChangedEventArgs) Handles WC.UploadProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
        ' 上传完毕后关闭当前窗口
        If (e.ProgressPercentage = 100) Then
            Me.Close()
            CommonModule.ShowAlert("任务文件上传成功！")
        End If
    End Sub
End Class