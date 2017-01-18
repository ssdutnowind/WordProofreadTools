''' <summary>
''' 公共应答部分
''' </summary>
Public Class ResponsePublic
    Public code As String

    Public msg As String
End Class

''' <summary>
''' 查询任务
''' </summary>
Public Class ResponseQueryTask
    Public Class QueryTask
        Public nickName As String
        Public taskId As String
        Public taskType As String
        Public taskFile As String
        Public fileName As String
    End Class

    Public jsonResult As QueryTask
End Class


