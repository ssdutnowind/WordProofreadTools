Imports System.Windows.Forms

Module CommonModule
    ' 服务器地址
    Public serverPath As String = ""
    ' 任务ID
    Public taskId As String = ""
    ' 临时令牌
    Public token As String = ""
    ' 昵称
    Public nickName As String = ""
    ' 任务类型
    Public taskType As String = ""
    ' 任务文件路径
    Public taskFile As String = ""
    ' 本地临时文件
    Public localFile As String = ""

    Public ribbon As Ribbon = Nothing

    ''' <summary>
    ''' 云端文件窗格
    ''' </summary>
    Public PaneCloudFiles As Microsoft.Office.Tools.CustomTaskPane = Nothing

    ''' <summary>
    ''' 显示云端文件窗格
    ''' </summary>
    Public Sub ShowCloudFilesPane()
        CommonModule.PaneCloudFiles.Visible = True
    End Sub

    ''' <summary>
    ''' 隐藏云端文件窗格
    ''' </summary>
    Public Sub HideCloudFilesPane()
        CommonModule.PaneCloudFiles.Visible = False
    End Sub

    ''' <summary>
    ''' 打印日志
    ''' </summary>
    ''' <param name="info"></param>
    Public Sub Log(info)
        System.Diagnostics.Debug.WriteLine("[" + Date.Now.ToLocalTime + "] " + info)
    End Sub

    ''' <summary>
    ''' 弹出提示信息框
    ''' </summary>
    ''' <param name="msg"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Public Function ShowAlert(msg As String, Optional type As String = "Info")
        Select Case type
            Case "Warning"
                Return MessageBox.Show(msg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Select
            Case "Error"
                Return MessageBox.Show(msg, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Select
            Case "Question"
                Return MessageBox.Show(msg, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Exit Select
            Case Else
                Return MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Select
    End Function

    ''' <summary>
    ''' 读取自定义属性
    ''' </summary>
    ''' <param name="propertyName"></param>
    ''' <returns></returns>
    Public Function ReadDocumentProperty(ByVal propertyName As String) As String
        Dim doc = Globals.ThisAddIn.Application.ActiveDocument
        propertyName = "vsto_" + propertyName
        Dim properties As Office.DocumentProperties = CType(doc.CustomDocumentProperties, Office.DocumentProperties)
        Dim prop As Office.DocumentProperty

        For Each prop In properties
            If prop.Name = propertyName Then
                Return prop.Value.ToString()
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' 写入自定义属性
    ''' </summary>
    ''' <param name="propertyName"></param>
    ''' <param name="value"></param>
    Public Sub WriteDocumentProperty(ByVal propertyName As String, ByVal value As String)
        Dim doc = Globals.ThisAddIn.Application.ActiveDocument
        Dim properties As Microsoft.Office.Core.DocumentProperties
        properties = CType(doc.CustomDocumentProperties, Office.DocumentProperties)

        If ReadDocumentProperty(propertyName) <> Nothing Then
            properties("vsto_" + propertyName).Delete()
        End If

        properties.Add("vsto_" + propertyName, False, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, value)
    End Sub

End Module
