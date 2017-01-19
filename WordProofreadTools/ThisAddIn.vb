Imports System.Collections
Imports System.Net
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json

Public Class ThisAddIn

    WithEvents applicationEvents As ApplicationEvents4_Event

    Public ribbon As Microsoft.Office.Core.IRibbonExtensibility

    ''' <summary>
    ''' 插件初始化
    ''' </summary>
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        CommonModule.Log("系统初始化……")
        ' 初始化配置
        CommonModule.serverPath = My.Settings.Item("Server")
        CommonModule.Log("[读取配置]服务器路径：" + CommonModule.serverPath)

        ' 设置默认状态
        CommonModule.ribbon.SetNormalState()

        ' 参数判断
        For Each arg As String In Environment.GetCommandLineArgs
            ' 如果有taskId参数
            If (arg.IndexOf("taskId") >= 1) Then
                Dim token As String = ""
                Dim task As String = ""
                arg = arg.Substring(arg.LastIndexOf("?") + 1)
                Dim params = arg.Split(",")
                For Each param As String In params
                    If (param.IndexOf("token") >= 0) Then
                        CommonModule.token = param.Substring(6)
                    ElseIf (param.IndexOf("taskId") >= 0) Then
                        CommonModule.taskId = param.Substring(7)
                    ElseIf (param.IndexOf("host") >= 0) Then
                        CommonModule.serverPath = param.Substring(5) + "/"
                    ElseIf (param.IndexOf("userId") >= 0) Then
                        CommonModule.userId = param.Substring(7)
                    End If
                Next
            End If
        Next

        ' 监听事件
        applicationEvents = Globals.ThisAddIn.Application

        If (String.IsNullOrEmpty(CommonModule.taskId) Or String.IsNullOrEmpty(CommonModule.token)) Then
            ' token或者taskId为空则作为离线模式
            CommonModule.ribbon.SetNormalState()
            'CommonModule.ribbon.SetProofreadState()

            ' 插件启动时已经有文档打开
            If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
                applicationEvents4_DocumentOpen(Nothing)
            End If

        Else
            ' 否则查询任务属性
            DoQueryTask()
        End If

    End Sub

    ''' <summary>
    ''' 查询任务信息
    ''' </summary>
    Private Sub DoQueryTask()
        ' 查询任务信息
        ' 请求参数
        Dim params = New Hashtable()
        params.Add("taskId", CommonModule.taskId)
        params.Add("token", CommonModule.token)
        ' 下发请求
        Dim request = New HttpRequest()
        Dim response = request.DoSendRequest(My.Settings.Item("queryTaskUrl"), params)
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
                If (IsNothing(json) Or IsNothing(json.jsonResult)) Then
                    ' 解析异常也提示失败
                    CommonModule.ShowAlert("查询任务信息失败！", "Error")
                End If

                ' 昵称
                CommonModule.nickName = json.jsonResult.nickName
                ' 任务类型
                CommonModule.taskType = json.jsonResult.taskType
                ' 任务文件
                CommonModule.taskFile = json.jsonResult.taskFile
                ' 文件名
                CommonModule.fileName = json.jsonResult.fileName

                ' 数据读取完毕，开始初始化状态
                AfterTaskDataReaded(True)
            Else
                ' 服务器返回异常消息
                CommonModule.ShowAlert(jsonPublic.msg, "Error")
                Return
            End If
        End If

    End Sub

    ''' <summary>
    ''' 解析任务属性的应答
    ''' </summary>
    Private Sub AfterTaskDataReaded(ByVal fromNetwork As Boolean)
        If (CommonModule.taskType = "1" Or CommonModule.taskType = "2" Or CommonModule.taskType = "3") Then
            ' 审核任务
            CommonModule.ribbon.SetAuditState()
        ElseIf (CommonModule.taskType = "6" Or CommonModule.taskType = "7" Or CommonModule.taskType = "8") Then
            ' 校对任务
            CommonModule.ribbon.SetProofreadState()
        Else
            ' 默认编纂任务
            CommonModule.ribbon.SetEditorState()
        End If

        If (fromNetwork = True) Then
            ' 准备就绪开始下载任务文件
            StartDownloadTask()
        End If
    End Sub

    ''' <summary>
    ''' 开始下载任务文件
    ''' </summary>
    Private Sub StartDownloadTask()
        Dim file As String = CommonModule.taskFile
        ' 如果没有返回fileName则从URL解析
        If (String.IsNullOrEmpty(CommonModule.fileName)) Then
            file = file.Substring(file.LastIndexOf("/"))
        Else
            file = CommonModule.fileName
        End If

        Dim download = New FormDownload()
        download.StartDownload(CommonModule.taskFile, file)
        download.ShowDialog()
    End Sub

    ''' <summary>
    ''' 文档打开事件
    ''' </summary>
    ''' <param name="Doc"></param>
    Private Sub applicationEvents4_DocumentOpen(doc As Document) Handles applicationEvents.DocumentOpen
        CommonModule.Log("文档打开……")

        Dim document = Globals.ThisAddIn.Application.ActiveDocument
        If (String.IsNullOrEmpty(CommonModule.taskId)) Then
            CommonModule.Log("打开本地文件……")
            ' 没有taskId说明是本地打开文件，尝试读取response
            Dim taskId As String = CommonModule.ReadDocumentProperty("taskId")

            If (taskId <> Nothing) Then
                CommonModule.Log("读取到自定义属性……")
                CommonModule.taskId = CommonModule.ReadDocumentProperty("taskId")
                CommonModule.taskType = CommonModule.ReadDocumentProperty("taskType")
                CommonModule.nickName = CommonModule.ReadDocumentProperty("nickName")
                CommonModule.userId = CommonModule.ReadDocumentProperty("userId")
                ' 数据读取完毕，开始初始化状态
                AfterTaskDataReaded(False)
            End If
        Else
            ' 有taskId说明是网络下载后打开
            ' 将应答写入自定义属性
            CommonModule.Log("即将写入自定义属性……")
            CommonModule.WriteDocumentProperty("taskId", CommonModule.taskId)
            CommonModule.WriteDocumentProperty("taskType", CommonModule.taskType)
            CommonModule.WriteDocumentProperty("nickName", CommonModule.nickName)
            CommonModule.WriteDocumentProperty("userId", CommonModule.userId)
            doc.Save()
            ' 为了正确设置修订状态，再调用一次
            AfterTaskDataReaded(False)
        End If

    End Sub

    ''' <summary>
    ''' 文档关闭事件
    ''' </summary>
    ''' <param name="doc"></param>
    ''' <param name="cancel"></param>
    Private Sub applicationEvents4_DocumentClose(doc As Document, ByRef cancel As Boolean) Handles applicationEvents.DocumentBeforeClose
        CommonModule.Log("文件关闭……")
        If (Not String.IsNullOrEmpty(CommonModule.token)) Then
            ' 设置默认状态
            CommonModule.ribbon.SetNormalState()
            ' 清空文件相关全局数据
            CommonModule.token = ""
            CommonModule.taskId = ""
            CommonModule.nickName = ""
            CommonModule.taskType = ""
            CommonModule.taskFile = ""
            CommonModule.localFile = ""
            CommonModule.userId = ""
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Me.ribbon = New Ribbon()
        Return Me.ribbon
    End Function
End Class
