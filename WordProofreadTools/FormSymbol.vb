Public Class FormSymbol

    ''' <summary>
    ''' 初始化生成列表
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FormSymbol_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 基本单位
        Dim array1 = {
            New String() {"长度", "米", "m", "插入符号"},
            New String() {"质量", "千克（公斤）", "kg", "插入符号"},
            New String() {"时间", "秒"， "s", "插入符号"},
            New String() {"电流", "安[培]", "A", "插入符号"},
            New String() {"热力学温度", "开[尔文]", "K", "插入符号"},
            New String() {"物质的量", "摩[尔]", "mol", "插入符号"},
            New String() {"发光的强度", "坎[德拉]", "cd", "插入符号"}
        }
        For i As Integer = 0 To array1.Length - 1
            Me.DataGridView1.Rows.Add(array1(i))
        Next

        ' 导出单位
        Dim array2 = {
            New String() {"[平面]角", "弧度", "rad", "插入符号"},
            New String() {"立体角", "球面度", "sr", "插入符号"},
            New String() {"频率", "赫[兹]"， "Hz", "插入符号"},
            New String() {"力", "牛[顿]", "N", "插入符号"},
            New String() {"压力，压强，应力", "帕[斯卡]", "Pa", "插入符号"},
            New String() {"能[量]，功，热量", "焦[耳]", "J", "插入符号"},
            New String() {"功率，辐[射能]通量", "瓦[特]", "W", "插入符号"},
            New String() {"电荷[量]", "库[伦]", "C", "插入符号"},
            New String() {"电压，电动势，电位，（电势）", "伏[特]", "V", "插入符号"},
            New String() {"电阻", "欧[姆]", "Ω", "插入符号"},
            New String() {"电导", "西[门子]", "S", "插入符号"},
            New String() {"磁通[量]", "韦[伯]", "Wb", "插入符号"},
            New String() {"磁通[量]密度，磁感应强度", "特[斯拉]", "T", "插入符号"},
            New String() {"电感", "亨[利]", "H", "插入符号"},
            New String() {"光通量", "流[明]", "lm", "插入符号"},
            New String() {"[光]照度", "勒[克斯]", "lx", "插入符号"},
            New String() {"[放射性]活度", "贝可[勒尔]", "Bq", "插入符号"},
            New String() {"吸收剂量", "戈[瑞]", "Gy", "插入符号"},
            New String() {"剂量当量", "希[沃特]", "Sv", "插入符号"}
        }
        For i As Integer = 0 To array1.Length - 1
            Me.DataGridView2.Rows.Add(array1(i))
        Next

    End Sub

    ''' <summary>
    ''' 插入符号
    ''' </summary>
    ''' <param name="symbol"></param>
    Private Sub InsertSymbol(symbol)
        Dim app = Globals.ThisAddIn.Application
        Dim selection = app.Selection
        Dim range = selection.Range
        range.InsertAfter(symbol)
        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        range.Select()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (e.ColumnIndex = 3) Then
            Dim symbol = Me.DataGridView1.Rows.Item(e.RowIndex).Cells(2).Value
            InsertSymbol(symbol)
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If (e.ColumnIndex = 3) Then
            Dim symbol = Me.DataGridView2.Rows.Item(e.RowIndex).Cells(2).Value
            InsertSymbol(symbol)
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        If (e.ColumnIndex = 3) Then
            Dim symbol = Me.DataGridView3.Rows.Item(e.RowIndex).Cells(2).Value
            InsertSymbol(symbol)
        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick
        If (e.ColumnIndex = 3) Then
            Dim symbol = Me.DataGridView4.Rows.Item(e.RowIndex).Cells(2).Value
            InsertSymbol(symbol)
        End If
    End Sub
End Class