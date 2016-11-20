Module Common
    '服务器地址
    Public ServerPath As String = ""



    Public Sub Log(info)
        System.Diagnostics.Debug.WriteLine("[" + Date.Now.ToLocalTime + "] " + info)
    End Sub

End Module
