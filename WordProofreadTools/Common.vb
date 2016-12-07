﻿Imports System.Windows.Forms

Module Common
    ''' <summary>
    ''' 服务器地址
    ''' </summary>
    Public ServerPath As String = ""


    ''' <summary>
    ''' 云端文件窗格
    ''' </summary>
    Public PaneCloudFiles As Microsoft.Office.Tools.CustomTaskPane = Nothing

    ''' <summary>
    ''' 显示云端文件窗格
    ''' </summary>
    Public Sub ShowCloudFilesPane()
        Common.PaneCloudFiles.Visible = True
    End Sub

    ''' <summary>
    ''' 隐藏云端文件窗格
    ''' </summary>
    Public Sub HideCloudFilesPane()
        Common.PaneCloudFiles.Visible = False
    End Sub

    ''' <summary>
    ''' 打印日志
    ''' </summary>
    ''' <param name="info"></param>
    Public Sub Log(info)
        System.Diagnostics.Debug.WriteLine("[" + Date.Now.ToLocalTime + "] " + info)
    End Sub

    ''' <summary>
    ''' 弹出提示信息框
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Public Function ShowAlert(msg As String, Optional type As String = "Info")
        Select Case type
            Case "Warning"
                Return MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Select
            Case "Error"
                Return MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Select
            Case "Question"
                Return MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Question)
                Exit Select
            Case Else
                Return MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Select
    End Function

End Module