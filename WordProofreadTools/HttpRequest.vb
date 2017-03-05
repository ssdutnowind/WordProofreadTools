Imports System.Collections
Imports System.IO
Imports System.Net

Public Class HttpRequest
    Public Function DoSendRequest(ByVal cmd As String, Optional ByVal params As Hashtable = Nothing, Optional ByVal method As String = "GET")
        ' 完整请求地址
        Dim path As String = CommonModule.serverPath + cmd ' + ".json"

        ' 请求参数
        Dim postData As String = ""
        If (Not params Is Nothing) Then
            For Each dic As DictionaryEntry In params
                postData += dic.Key + "=" + WebUtility.UrlEncode(dic.Value) + "&"
            Next
        End If
        If (postData.Length > 0) Then
            postData = postData.Substring(0, postData.Length - 1)
        End If
        ' 请求对象
        Dim request As WebRequest
        CommonModule.Log("[HttpRequest] 下发请求……")
        If (method = "GET") Then
            If (postData.Length) Then
                path += "?" + postData
            End If
            request = WebRequest.Create(path)
            request.Method = "GET"
            CommonModule.Log("[HttpRequest] 请求地址：" + path)
        Else
            request = WebRequest.Create(path)
            ' 请求方法
            request.Method = "POST"
            CommonModule.Log("[HttpRequest] 请求地址：" + path)
            CommonModule.Log("[HttpRequest] 请求数据：" + postData)
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
            ' ContentType
            request.ContentType = "application/x-www-form-urlencoded"
            ' ContentLength
            request.ContentLength = byteArray.Length
            ' 发送请求
            Dim requestStream As Stream = request.GetRequestStream()
            requestStream.Write(byteArray, 0, byteArray.Length)
            requestStream.Close()
        End If

        ' 接收应答
        Dim response As WebResponse
        Try
            response = request.GetResponse()
        Catch ex As Exception
            Return ""
        End Try

        ' Display the status.
        CommonModule.Log("[HttpRequest] 应答状态：" + CType(response, HttpWebResponse).StatusDescription)
        Dim dataStream = response.GetResponseStream()
        Dim reader As New StreamReader(dataStream)
        Dim responseFromServer As String = reader.ReadToEnd()
        ' Display the content.
        CommonModule.Log("[HttpRequest] 应答内容：" + responseFromServer)
        ' Clean up the streams.
        reader.Close()
        dataStream.Close()
        response.Close()

        Return responseFromServer
    End Function
End Class