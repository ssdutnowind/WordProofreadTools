Imports System.Windows.Forms

Public Class FormLocalFiles

    Private localPath = String.Empty

    ''' <summary>
    ''' 关闭
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClose_Click(sender As Object, e As EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' 全选按钮
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles BtnCheckAll.Click
        Dim value = False
        If (Me.BtnCheckAll.Text = "全部选择") Then
            value = True
        End If
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            row.Cells.Item("ColumnCheck").Value = value
        Next
        AfterSelectChanged()
    End Sub

    ''' <summary>
    ''' 删除
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnRemove_Click(sender As Object, e As EventArgs) Handles BtnRemove.Click
        Dim result = CommonModule.ShowAlert("确定要删除全部选择的文件么？", "Question")
        If (result = DialogResult.Yes) Then
            Try
                For Each row As DataGridViewRow In Me.DataGridView1.Rows
                    If (row.Cells.Item("ColumnCheck").Value) Then
                        My.Computer.FileSystem.DeleteFile(row.Cells.Item("ColumnPath").Value)
                        row.Visible = False
                    End If
                Next
            Catch ex As Exception
                CommonModule.Log("删除文件出错：" + vbCrLf + ex.Message)
                CommonModule.ShowAlert("删除文件出错：" + vbCrLf + ex.Message, "error")
            End Try
        End If
    End Sub

    ''' <summary>
    ''' 选择修改后
    ''' </summary>
    Private Sub AfterSelectChanged()
        Dim checkAll As Boolean = True
        Dim hasChecked As Boolean = False
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            ' 任何一项未选中
            If (row.Cells.Item("ColumnCheck").Value) Then
                hasChecked = True
            Else
                checkAll = False
            End If
        Next
        If (Me.DataGridView1.Rows.Count = 0) Then
            checkAll = False
        End If

        ' 设置选择按钮文本
        If (checkAll) Then
            Me.BtnCheckAll.Text = "全部取消"
        Else
            Me.BtnCheckAll.Text = "全部选择"
        End If
        ' 设置删除按钮状态
        If (hasChecked) Then
            Me.BtnRemove.Enabled = True
        Else
            Me.BtnRemove.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' 初始化加载文件列表
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FormLocalFiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly()
        localPath = My.Computer.FileSystem.GetParentPath(assemblyInfo.Location) + "\temp"
        Dim fileName = ""
        If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
            fileName = Globals.ThisAddIn.Application.ActiveDocument.FullName
        End If

        ' 判断文件夹是否存在
        If My.Computer.FileSystem.DirectoryExists(localPath) Then
            Dim files = My.Computer.FileSystem.GetFiles(localPath, FileIO.SearchOption.SearchTopLevelOnly, "*.doc")
            Dim fileInfo As IO.FileInfo
            Dim row As Integer = 0
            ' 遍历检索到的所有文件
            For Each file In files
                fileInfo = My.Computer.FileSystem.GetFileInfo(file)
                ' 非当前文件、隐藏文件、系统文件才添加到列表
                If ((fileName.Equals(fileInfo.FullName) Or (fileInfo.Attributes And FileAttribute.Hidden) Or (fileInfo.Attributes And FileAttribute.System))) Then
                    Continue For
                End If
                row = Me.DataGridView1.Rows.Add(False, fileInfo.Name, fileInfo.LastWriteTime.ToLocalTime, file, "打开")
            Next
        End If
    End Sub

    ''' <summary>
    ''' 表格值变化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        ' 仅监控第一列
        If (e.ColumnIndex = 0) Then
            AfterSelectChanged()
        End If

    End Sub

    ''' <summary>
    ''' 实时触发Checkbox修改
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        Dim grid As DataGridView = sender
        If (Not IsNothing(grid)) Then
            grid.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ''' <summary>
    ''' 打开文件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.ColumnIndex = 4) Then
            Dim name = Me.DataGridView1.Rows.Item(e.RowIndex).Cells.Item("ColumnFileName").Value
            Dim result = CommonModule.ShowAlert("确定要打开文件：" + name + "？", "Question")
            If (result = DialogResult.Yes) Then
                Dim path = Me.DataGridView1.Rows.Item(e.RowIndex).Cells.Item("ColumnPath").Value
                Try
                    ' 打开文档
                    If (Globals.ThisAddIn.Application.Documents.Count > 0) Then
                        Globals.ThisAddIn.Application.ActiveDocument.Close()
                    End If

                    Globals.ThisAddIn.Application.Documents.Open(FileName:=path, ConfirmConversions:=True)
                    Me.Close()
                Catch ex As Exception
                    CommonModule.Log("加载Word文件出错：" + vbCrLf + ex.Message)
                    CommonModule.ShowAlert("加载Word文件出错：" + vbCrLf + ex.Message, "error")
                End Try
            End If
        End If
    End Sub
End Class