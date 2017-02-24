Public Class Form1
    Private Sub BtnSecurity_Click(sender As Object, e As EventArgs) Handles BtnSecurity.Click

        Dim err As String = ""

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
            err += ex.Message + vbCrLf
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
            err += ex.Message + vbCrLf
        End Try

        Dim msg = "安全策略优化完毕！"

        If (Not String.IsNullOrEmpty(err)) Then
            msg += vbCrLf + "但在优化过程中发生了一些错误，但这些错误可能并不影响继续安装插件，您可尝试重新安装插件确认问题是否解决。"
            msg += vbCrLf + "具体错误信息："
            msg += vbCrLf + err
        End If

        MessageBox.Show(msg)

    End Sub

End Class
