''' <summary>
''' 公共应答部分
''' </summary>
Public Class ResponsePublic
    Public code As String

    Public msg As String
End Class

''' <summary>
''' 登录
''' </summary>
Public Class ResponseDoLogin
    Public Class DoLogin
        Public userId As String
        Public userName As String
        Public nickName As String
        Public token As String
    End Class

    Public data As DoLogin
End Class

Public Class DoLogoutResponse

End Class
