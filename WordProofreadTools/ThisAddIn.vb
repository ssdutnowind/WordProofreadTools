﻿Public Class ThisAddIn

    Private WithEvents taskPaneCloudFiles As Microsoft.Office.Tools.CustomTaskPane

    Public ribbon As Microsoft.Office.Core.IRibbonExtensibility

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        CommonModule.Log("系统初始化……")
        '初始化配置
        CommonModule.ServerPath = My.Settings.Item("Server")
        CommonModule.Log("[读取配置]服务器路径：" + CommonModule.ServerPath)

        '初始化云端文件窗格
        Dim cloudFiles = New FormCloudFiles()
        taskPaneCloudFiles = Me.CustomTaskPanes.Add(cloudFiles, "云端文件")
        taskPaneCloudFiles.Width = 250
        CommonModule.PaneCloudFiles = taskPaneCloudFiles

        For Each arg As String In Environment.GetCommandLineArgs
            ' 如果有token参数
            If (arg.IndexOf("token") >= 1) Then
                Dim token As String = ""
                Dim task As String = ""
                arg = arg.Substring(arg.LastIndexOf("/") + 1)
                Dim params = arg.Split(",")
                For Each param As String In params
                    If (param.IndexOf("token") >= 0) Then
                        token = param.Substring(6)
                    ElseIf (param.IndexOf("task") >= 0) Then
                        task = param.Substring(5)
                    End If
                Next

                MsgBox("token: " + token + vbCrLf + "task: " + task)
            End If
        Next

    End Sub

    ''' <summary>
    ''' 监听Pane显示隐藏
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PaneCloudFiles_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles taskPaneCloudFiles.VisibleChanged
        CType(Me.ribbon, Ribbon).SetTaskWindowChecked(taskPaneCloudFiles.Visible)
        'Globals.Ribbons.Ribbon1.CbxRobbinCloudFiles.Checked = taskPaneCloudFiles.Visible
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Me.ribbon = New Ribbon()
        Return Me.ribbon
    End Function
End Class
