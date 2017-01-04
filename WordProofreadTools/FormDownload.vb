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

    ''' <summary>
    ''' 初始化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FormDownload_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private WithEvents WC As New WebClient

    ''' <summary>
    ''' 下载文件到本地
    ''' </summary>
    ''' <param name="url">完整下载路径</param>
    ''' <param name="file">文件名</param>
    Public Sub StartDownload(url As String, file As String)
        Dim assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly()
        Dim path = My.Computer.FileSystem.GetParentPath(assemblyInfo.Location)
        MsgBox(path + "\temp")
        Try
            My.Computer.FileSystem.CreateDirectory(path + "\temp")
            CommonModule.localFile = path + "\temp\" + file
            WC.DownloadFileAsync(New Uri(url), CommonModule.localFile)
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub

    ''' <summary>
    ''' 监听下载任务
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WC_DownloadProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs) Handles WC.DownloadProgressChanged
        ProgressBar1.Value = e.ProgressPercentage

        If (e.ProgressPercentage = 100) Then
            Globals.ThisAddIn.Application.Documents.Open(CommonModule.localFile)
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' 下载完毕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WC_DownloadFinished(ByVal sender As Object, ByVal e As DownloadDataCompletedEventArgs) Handles WC.DownloadFileCompleted
        Globals.ThisAddIn.Application.Documents.Open(CommonModule.localFile)
        Me.Close()
    End Sub

End Class