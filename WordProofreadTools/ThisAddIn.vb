Public Class ThisAddIn


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Common.Log("系统初始化……")
        '初始化配置
        Common.ServerPath = My.Settings.Item("Server")
        Common.Log("[读取配置]服务器路径：" + Common.ServerPath)



    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
