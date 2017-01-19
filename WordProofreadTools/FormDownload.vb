﻿Imports System.ComponentModel
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
        ' 读取程序安装目录
        Dim assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly()
        Dim path = My.Computer.FileSystem.GetParentPath(assemblyInfo.Location)

        Try
            ' 创建目标临时目录
            My.Computer.FileSystem.CreateDirectory(path + "\temp")
            ' 开始异步下载
            CommonModule.localFile = path + "\temp\" + file
            WC.DownloadFileAsync(New Uri(url), CommonModule.localFile)
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub

    ''' <summary>
    ''' 监听下载进度
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WC_DownloadProgressChanged(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs) Handles WC.DownloadProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    ''' <summary>
    ''' 监听下载完成
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub WC_DownloadFileComplate(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs) Handles WC.DownloadFileCompleted
        If e.Cancelled Then
            ' 任务被取消
            Me.Close()
        ElseIf e.Error IsNot Nothing Then
            Me.Hide()
            CommonModule.ShowAlert("下载任务出错！" + vbCrLf + e.Error.Message, "Error")
            Me.Close()
        Else
            ' 打开文件并关闭当前对话框
            Globals.ThisAddIn.Application.Documents.Open(CommonModule.localFile)
            Me.Close()
            'Globals.ThisAddIn.Application.Activate()
        End If
    End Sub

End Class