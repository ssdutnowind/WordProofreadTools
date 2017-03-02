Public Class FormErrorDialog

    ''' <summary>
    ''' 关闭窗口
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' 设置错误消息内容
    ''' </summary>
    ''' <param name="err"></param>
    Public Sub SetErrorMessage(ByVal err As Exception)
        ' 错误内容
        Dim errMsg = err.Message + vbCrLf + vbCrLf
        errMsg += "详细信息：" + vbCrLf
        ' 详细代码及调用栈
        errMsg += err.Source + vbCrLf + vbCrLf
        errMsg += err.StackTrace

        Me.txtErrorMessage.Text = errMsg
        ' 记录异常信息
        CommonModule.Log(errMsg)
    End Sub
End Class