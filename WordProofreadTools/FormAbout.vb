Imports System.Windows.Forms

Public Class FormAbout

    Private count As Integer = 0

    ''' <summary>
    ''' 点击后显示文本框
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Me.count += 1
        ' 三次双击显示日志
        If (count >= 3) Then
            Me.TextBox1.Text = CommonModule.logs
            Me.Height = 350
            Me.Width = 500
        End If
    End Sub

    ''' <summary>
    ''' 保存日志
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnSaveLog_Click(sender As Object, e As EventArgs) Handles BtnSaveLog.Click
        Dim dialog = New SaveFileDialog()
        dialog.Filter = "文本文件|*.txt"
        dialog.FileName = "系统日志" + Date.Now.ToFileTime.ToString() + ".txt"
        dialog.Title = "请选择导出的日志文件"
        '确定后开始导出
        If dialog.ShowDialog() = DialogResult.OK Then
            Try
                Dim file = dialog.FileName
                My.Computer.FileSystem.WriteAllText(file, CommonModule.logs, False)
                CommonModule.ShowAlert("日志已导出到：" + vbCrLf + file)
            Catch ex As Exception
                CommonModule.ShowAlert(ex.Message, "Error")
            End Try
        End If
    End Sub
End Class