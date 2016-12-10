Public Class ThisAddIn

    Private WithEvents taskPaneCloudFiles As Microsoft.Office.Tools.CustomTaskPane

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



    End Sub

    ''' <summary>
    ''' 监听Pane显示隐藏
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PaneCloudFiles_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles taskPaneCloudFiles.VisibleChanged
        Globals.Ribbons.Ribbon1.CbxRobbinCloudFiles.Checked = taskPaneCloudFiles.Visible
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
End Class
