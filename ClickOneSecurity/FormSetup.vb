Public Class FormSetup
    Private Sub BtnSetup_Click(sender As Object, e As EventArgs) Handles BtnSetup.Click
        Me.ProgressBar1.Value = 20
        AddLog("开始「协同编纂工具」的安装准备！")
        ' 优化安全策略
        Me.ProgressBar1.Value = 40
        AddLog("正在优化系统安全策略……")
        CheckSystemSecurity()
        ' 注册表注册
        Me.ProgressBar1.Value = 60
        AddLog("正在注册自定义URL协议……")
        RegisterCustomPortol()
        Me.ProgressBar1.Value = 80
        AddLog("完成安装前的准备！")
        StartSetup()
        Me.ProgressBar1.Value = 100
        AddLog("尝试启动Microsoft ClickOne！")
        AddLog("请在完成安装后关闭当前窗口！")

    End Sub

    ''' <summary>
    ''' 记录日志
    ''' </summary>
    ''' <param name="msg"></param>
    Private Sub AddLog(msg)
        Me.TextBox1.Text += msg + vbCrLf + vbCrLf
        Me.Label2.Text = msg
    End Sub

    ''' <summary>
    ''' 调整优化系统安全策略
    ''' </summary>
    Private Sub CheckSystemSecurity()
        Try
            ' 32位
            Dim key As Microsoft.Win32.RegistryKey
            key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey("SOFTWARE\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel")
            key.SetValue("MyComputer", "Enabled")
            key.SetValue("LocalIntranet", "Enabled")
            key.SetValue("Internet", "Enabled")
            key.SetValue("TrustedSites", "Enabled")
            key.SetValue("UntrustedSites", "Enabled")
            key.Close()
        Catch ex As Exception
            Me.TextBox1.Text += ex.Message
        End Try

        Try
            ' 64位
            Dim key As Microsoft.Win32.RegistryKey
            key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey("SOFTWARE\Wow6432Node\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel")
            key.SetValue("MyComputer", "Enabled")
            key.SetValue("LocalIntranet", "Enabled")
            key.SetValue("Internet", "Enabled")
            key.SetValue("TrustedSites", "Enabled")
            key.SetValue("UntrustedSites", "Enabled")
            key.Close()
        Catch ex As Exception
            Me.TextBox1.Text += ex.Message
        End Try
    End Sub

    ''' <summary>
    ''' 注册自定义协议
    ''' </summary>
    Private Sub RegisterCustomPortol()
        '' 当前程序路径
        'Dim assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly()
        'Dim path = My.Computer.FileSystem.GetParentPath(assemblyInfo.Location)
        'path += "\winword.reg"

        'Dim regPath = """" + path + """"

        'Try
        '    Process.Start("regedit", String.Format(" /s {0}", regPath))
        'Catch ex As Exception
        '    Me.TextBox1.Text += ex.Message
        'End Try

        ' [HKEY_CLASSES_ROOT\MSWORD]
        ' @="URL:MSWORD Protocol Handler"
        ' "URL Protocol"=""
        Try
            Dim key As Microsoft.Win32.RegistryKey
            key = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey("MSWORD")
            key.SetValue("", "URL:MSWORD Protocol Handler")
            key.SetValue("URL Protocol", "")
            ' [HKEY_CLASSES_ROOT\MSWORD\shell]
            key = key.CreateSubKey("shell")
            ' [HKEY_CLASSES_ROOT\MSWORD\shell\open]
            key = key.CreateSubKey("open")
            ' [HKEY_CLASSES_ROOT\MSWORD\shell\open\command]
            ' @="WINWORD.exe /n /\"%1\""
            key = key.CreateSubKey("command")
            key.SetValue("", "WINWORD.exe /n /""%1""")
        Catch ex As Exception
            Me.TextBox1.Text += ex.Message
        End Try

    End Sub

    ''' <summary>
    ''' 开始安装
    ''' </summary>
    Private Sub StartSetup()
        ' 当前程序路径
        Dim assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly()
        Dim path = My.Computer.FileSystem.GetParentPath(assemblyInfo.Location)
        path += "\setup.exe"

        Try
            Process.Start(path)
        Catch ex As Exception
            Me.TextBox1.Text += ex.Message
        End Try
    End Sub

    ''' <summary>
    ''' 显示详细日志
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnShowLog_Click(sender As Object, e As EventArgs) Handles BtnShowLog.Click
        Me.Height = 300
    End Sub
End Class