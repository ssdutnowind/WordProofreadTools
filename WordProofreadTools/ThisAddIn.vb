﻿Imports System.Collections
Imports System.Net
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json

Public Class ThisAddIn

    WithEvents applicationEvents As ApplicationEvents4_Event

    Public ribbon As Microsoft.Office.Core.IRibbonExtensibility

    ''' <summary>
    ''' 全局异常处理函数
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="args"></param>
    Private Sub GlobalExceptionHandler(sender As Object, args As UnhandledExceptionEventArgs)
        Dim e As Exception = DirectCast(args.ExceptionObject, Exception)
        Dim dialog = New FormErrorDialog()
        dialog.SetErrorMessage(e)
        dialog.ShowDialog()
    End Sub

    ''' <summary>
    ''' 插件初始化
    ''' </summary>
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        CommonModule.Log("[初始化] 系统初始化……")
        ' 处理全局异常
        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        AddHandler currentDomain.UnhandledException, AddressOf GlobalExceptionHandler

        Try
            ' 初始化配置
            CommonModule.serverPath = My.Settings.Item("Server")
            CommonModule.Log("[初始化] 默认服务器路径：" + CommonModule.serverPath)

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
        Catch ex As Exception
            Dim dialog = New FormErrorDialog()
            dialog.SetErrorMessage(ex)
            dialog.ShowDialog()
        End Try
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
            CommonModule.Log("[任务查询] 查询任务信息失败！应答报文为空。")
            CommonModule.ShowAlert("查询任务信息失败！" + vbCrLf + "解析服务器返回的应答报文出错。", "Error")
            Return
        Else
            ' 解析应答
            Dim jsonPublic As ResponsePublic
            Dim json As ResponseQueryTask

            Try
                jsonPublic = JsonConvert.DeserializeObject(Of ResponsePublic)(response)
            Catch ex As Exception
                ' 解析异常也提示失败
                CommonModule.Log("[任务查询] 查询任务信息失败！" + vbCrLf + ex.Message)
                CommonModule.ShowAlert("查询任务信息失败！" + vbCrLf + "解析服务器返回的应答报文出错。", "Error")
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
                    CommonModule.Log("[任务查询] 解析应答报文失败：消息体解析结果为空。")
                End If

                ' 应答报文检查
                If (String.IsNullOrEmpty(json.jsonResult.nickName) Or String.IsNullOrEmpty(json.jsonResult.taskType) Or
                    String.IsNullOrEmpty(json.jsonResult.taskFile) Or String.IsNullOrEmpty(json.jsonResult.fileName)) Then
                    CommonModule.ShowAlert("查询任务信息失败，服务器应答报文格式错误！", "Error")
                    CommonModule.Log("[任务查询] 查询任务信息失败，服务器应答报文格式错误，必要参数为空。")
                    Return
                End If

                ' 昵称
                CommonModule.nickName = json.jsonResult.nickName
                    ' Word限制昵称长度52个字符、缩写长度8个字符
                    If (CommonModule.nickName.Length > 8) Then
                        CommonModule.nickName = CommonModule.nickName.Substring(0, 8)
                        CommonModule.Log("[任务查询] 昵称过长，截取为：" + CommonModule.nickName)
                    End If
                    ' 任务类型
                    CommonModule.taskType = json.jsonResult.taskType
                    ' 任务文件
                    CommonModule.taskFile = json.jsonResult.taskFile
                    ' 文件名
                    CommonModule.fileName = json.jsonResult.fileName
                    ' 替换半角冒号，其余字符不管，到时会报错
                    CommonModule.fileName = CommonModule.fileName.Replace(":", "：")

                    ' 数据读取完毕，开始初始化状态
                    AfterTaskDataReaded(True)
                Else
                    ' 服务器返回异常消息
                    CommonModule.Log("[任务查询] 服务器返回异常：" + jsonPublic.msg)
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
        CommonModule.Log("[初始化] 开始下载任务文件……")
        Dim file As String = CommonModule.taskFile
        ' 如果没有返回fileName则从URL解析
        If (String.IsNullOrEmpty(CommonModule.fileName)) Then
            file = file.Substring(file.LastIndexOf("/"))
        Else
            file = CommonModule.fileName
        End If

        ' 检查非法文件名
        If (file Like "*[\/:*?""<|>]*") Then
            CommonModule.Log("[初始化] 文件名非法：" + file)
            CommonModule.ShowAlert("任务文件名（" + file + "）包含非法字符，请联系管理员！", "Error")
            Return
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
        CommonModule.Log("[初始化] 文档打开……")

        Try
            ' 不知为何触发该事件时也出现过无打开文档的异常
            If (Globals.ThisAddIn.Application.Documents.Count = 0) Then
                Return
            End If

            Dim document = Globals.ThisAddIn.Application.ActiveDocument
            If (String.IsNullOrEmpty(CommonModule.taskId)) Then
                CommonModule.Log("[初始化] 打开本地文件……")
                ' 没有taskId说明是本地打开文件，尝试读取response
                Dim taskId As String = CommonModule.ReadDocumentProperty("taskId")

                If (taskId <> Nothing) Then
                    CommonModule.Log("[初始化] 读取到自定义属性：")
                    CommonModule.taskId = CommonModule.ReadDocumentProperty("taskId")
                    CommonModule.taskType = CommonModule.ReadDocumentProperty("taskType")
                    CommonModule.nickName = CommonModule.ReadDocumentProperty("nickName")
                    CommonModule.userId = CommonModule.ReadDocumentProperty("userId")
                    CommonModule.serverPath = CommonModule.ReadDocumentProperty("serverPath")
                    CommonModule.Log("  taskId：" + CommonModule.taskId)
                    CommonModule.Log("  taskType：" + CommonModule.taskType)
                    CommonModule.Log("  nickName：" + CommonModule.nickName)
                    CommonModule.Log("  userId：" + CommonModule.userId)
                    CommonModule.Log("  serverPath：" + CommonModule.serverPath)
                    ' 数据读取完毕，开始初始化状态
                    AfterTaskDataReaded(False)
                End If
            Else
                ' 有taskId说明是网络下载后打开
                ' 将应答写入自定义属性
                CommonModule.Log("[初始化] 即将写入自定义属性：")
                CommonModule.WriteDocumentProperty("taskId", CommonModule.taskId)
                CommonModule.WriteDocumentProperty("taskType", CommonModule.taskType)
                CommonModule.WriteDocumentProperty("nickName", CommonModule.nickName)
                CommonModule.WriteDocumentProperty("userId", CommonModule.userId)
                CommonModule.WriteDocumentProperty("serverPath", CommonModule.serverPath)
                CommonModule.Log("  taskId：" + CommonModule.taskId)
                CommonModule.Log("  taskType：" + CommonModule.taskType)
                CommonModule.Log("  nickName：" + CommonModule.nickName)
                CommonModule.Log("  userId：" + CommonModule.userId)
                CommonModule.Log("  serverPath：" + CommonModule.serverPath)
                If (Not doc.ReadOnly) Then
                    doc.Save()
                End If

                ' 为了正确设置修订状态，再调用一次
                AfterTaskDataReaded(False)
                End If
        Catch ex As Exception
            Dim dialog = New FormErrorDialog()
            dialog.SetErrorMessage(ex)
            dialog.ShowDialog()
        End Try
    End Sub

    ''' <summary>
    ''' 文档关闭事件
    ''' </summary>
    ''' <param name="doc"></param>
    ''' <param name="cancel"></param>
    Private Sub applicationEvents4_DocumentClose(doc As Document, ByRef cancel As Boolean) Handles applicationEvents.DocumentBeforeClose
        CommonModule.Log("[初始化] 文件关闭……")
        'If (Not String.IsNullOrEmpty(CommonModule.token)) Then
        '    ' 设置默认状态
        '    CommonModule.ribbon.SetNormalState()
        '    ' 清空文件相关全局数据
        '    CommonModule.token = ""
        '    CommonModule.taskId = ""
        '    CommonModule.nickName = ""
        '    CommonModule.taskType = ""
        '    CommonModule.taskFile = ""
        '    CommonModule.localFile = ""
        '    CommonModule.userId = ""
        'End If

    End Sub

    ''' <summary>
    ''' 插件退出
    ''' </summary>
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Me.ribbon = New Ribbon()
        Return Me.ribbon
    End Function
End Class
