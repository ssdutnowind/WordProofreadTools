Imports System.Collections
Imports Newtonsoft.Json

Public Class FormLogin
    Public Event UserLogin(ByVal sender As Object)

    ''' <summary>
    ''' 取消关闭登录
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' 登录
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        Dim userName As String = Me.TxtUsename.Text
        Dim password As String = Me.TxtPassword.Text
        ' 检查用户名和密码
        If (String.IsNullOrEmpty(userName)) Then
            Me.LblMessage.Text = "请输入用户名！"
            Return
        End If
        If (String.IsNullOrEmpty(password)) Then
            Me.LblMessage.Text = "请输入密码！"
            Return
        End If
        Me.LblMessage.Text = ""
        ' 请求参数
        Dim params = New Hashtable()
        params.Add("userName", userName)
        params.Add("password", password)
        ' 下发请求
        Dim request = New HttpRequest()
        Dim response = request.DoSendRequest("doLogin", params)
        If (String.IsNullOrEmpty(response)) Then
            LblMessage.Text = "登录请求失败！"
            Return
        Else
            ' 解析应答
            Dim jsonPublic As ResponsePublic
            Dim json As ResponseDoLogin

            Try
                jsonPublic = JsonConvert.DeserializeObject(Of ResponsePublic)(response)
            Catch ex As Exception
                LblMessage.Text = "登录请求失败！"
                Return
            End Try
            If (jsonPublic.code = "00") Then
                ' 成功
                Try
                    json = JsonConvert.DeserializeObject(Of ResponseDoLogin)(response)
                Catch ex As Exception
                    LblMessage.Text = "登录请求失败！"
                    Return
                End Try

                UserInfo.nickName = json.data.nickName
                UserInfo.userId = json.data.userId
                UserInfo.userName = json.data.userName
                UserInfo.token = json.data.token

                RaiseEvent UserLogin(Me)
            Else
                ' 失败
                LblMessage.Text = jsonPublic.msg
            End If
        End If

    End Sub
End Class