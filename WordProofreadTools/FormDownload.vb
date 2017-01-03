Imports System.Net

Public Class FormDownload

    ''' <summary>
    ''' 关闭按钮，默认禁用
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub FormDownload_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private WithEvents WC As New WebClient

    ''' <summary>
    ''' 下载文件到本地
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="path"></param>
    Public Sub StartDownload(url As String, path As String)
        Try
            My.Computer.FileSystem.CreateDirectory(System.Environment.CurrentDirectory + "\temp")
            WC.DownloadFileAsync(New Uri(url), System.Environment.CurrentDirectory + "\temp\task.docx")
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub

    Private Sub WC_DownloadProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs) Handles WC.DownloadProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

End Class