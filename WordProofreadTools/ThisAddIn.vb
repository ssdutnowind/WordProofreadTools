Imports System.Collections
Imports System.Net
Imports Newtonsoft.Json

Public Class ThisAddIn

    'Public WithEvents WC As New WebClient

    Public ribbon As Microsoft.Office.Core.IRibbonExtensibility

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        CommonModule.Log("系统初始化……")
        ' 初始化配置
        CommonModule.serverPath = My.Settings.Item("Server")
        CommonModule.Log("[读取配置]服务器路径：" + CommonModule.serverPath)

        ' 参数判断
        For Each arg As String In Environment.GetCommandLineArgs
            ' 如果有token参数
            If (arg.IndexOf("token") >= 1) Then
                Dim token As String = ""
                Dim task As String = ""
                arg = arg.Substring(arg.LastIndexOf("/") + 1)
                Dim params = arg.Split(",")
                For Each param As String In params
                    If (param.IndexOf("token") >= 0) Then
                        CommonModule.token = param.Substring(6)
                    ElseIf (param.IndexOf("task") >= 0) Then
                        CommonModule.taskId = param.Substring(5)
                    End If
                Next

                MsgBox("token: " + token + vbCrLf + "task: " + task)
            End If
        Next

        If (String.IsNullOrEmpty(CommonModule.taskId) Or String.IsNullOrEmpty(CommonModule.token)) Then
            ' token或者taskId为空则作为离线模式
            CommonModule.ribbon.SetNormalState()
        Else
            ' 否则查询任务属性
            DoQueryTask()
        End If

    End Sub

    Private Sub DoQueryTask()
        ' 查询任务信息
        ' 请求参数
        Dim params = New Hashtable()
        params.Add("taskId", CommonModule.taskId)
        params.Add("token", CommonModule.token)
        ' 下发请求
        Dim request = New HttpRequest()
        Dim response = request.DoSendRequest("queryTask", params)
        If (String.IsNullOrEmpty(response)) Then
            CommonModule.ShowAlert("查询任务信息失败！", "Error")
            Return
        Else
            ' 解析应答
            Dim jsonPublic As ResponsePublic
            Dim json As ResponseQueryTask

            Try
                jsonPublic = JsonConvert.DeserializeObject(Of ResponsePublic)(response)
            Catch ex As Exception
                ' 解析异常也提示失败
                CommonModule.ShowAlert("查询任务信息失败！", "Error")
                Return
            End Try

            If (jsonPublic.code = "00") Then
                ' 成功
                Try
                    json = JsonConvert.DeserializeObject(Of ResponseQueryTask)(response)
                Catch ex As Exception
                    ' 解析异常也提示失败
                    CommonModule.ShowAlert("查询任务信息失败！", "Error")
                    Return
                End Try
                ' 昵称
                CommonModule.nickName = json.data.nickName
                ' 任务类型
                CommonModule.taskType = json.data.taskType
                ' 任务标签
                CommonModule.taskLabel = json.data.taskLabel
                ' 任务文件
                CommonModule.taskFile = json.data.taskFile

                ' 准备就绪开始下载任务文件
                StartDownloadTask()
            Else
                ' 服务器返回异常消息
                CommonModule.ShowAlert(jsonPublic.msg, "Error")
                Return
            End If
        End If

    End Sub

    Private Sub StartDownloadTask()
        Dim file As String = CommonModule.taskFile
        file = file.Substring(file.LastIndexOf("/"))

        CommonModule.Log(file)
        'Dim download = New FormDownload()
        'download.StartDownload("http://localhost:3000/files/test.docx", "task.docx")
        'download.ShowDialog()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Me.ribbon = New Ribbon()
        Return Me.ribbon
    End Function
End Class
