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
        For i As Integer = 0 To array2.Length - 1
            Me.DataGridView2.Rows.Add(array2(i))
        Next

        ' 十进制词头
        Dim array3 = {
            New String() {"10的24次方", "尧[它]", "Y", "插入符号"},
            New String() {"10的21次方", "泽[它]", "Z", "插入符号"},
            New String() {"10的18次方", "艾[可萨]"， "E", "插入符号"},
            New String() {"10的15次方", "拍[它]", "P", "插入符号"},
            New String() {"10的12次方", "太[拉]", "T", "插入符号"},
            New String() {"10的9次方", "吉[咖]", "G", "插入符号"},
            New String() {"10的6次方", "兆", "M", "插入符号"},
            New String() {"10的3次方", "千", "k", "插入符号"},
            New String() {"10的2次方", "百", "h", "插入符号"},
            New String() {"10的1次方", "十", "da", "插入符号"},
            New String() {"10的-1次方", "分", "d", "插入符号"},
            New String() {"10的-2次方", "厘", "c", "插入符号"},
            New String() {"10的-3次方", "毫", "m", "插入符号"},
            New String() {"10的-6次方", "微", "μ", "插入符号"},
            New String() {"10的-9次方", "纳[诺]", "n", "插入符号"},
            New String() {"10的-12次方", "皮[可]", "p", "插入符号"},
            New String() {"10的-15次方", "飞[母托]", "f", "插入符号"},
            New String() {"10的-18次方", "阿[托]", "a", "插入符号"},
            New String() {"10的-21次方", "仄[普托]", "z", "插入符号"},
            New String() {"10的-24次方", "幺[科托]", "y", "插入符号"}
        }
        For i As Integer = 0 To array3.Length - 1
            Me.DataGridView3.Rows.Add(array3(i))
        Next

        ' 计量单位
        Dim array4 = {
            New String() {"时间/分", "分", "min", "插入符号"},
            New String() {"时间/小时", "[小]时", "h", "插入符号"},
            New String() {"时间/天", "日，（天）"， "d", "插入符号"},
            New String() {"[平面]角/度", "度", "°", "插入符号"},
            New String() {"[平面]角/分", "[角]分", "′", "插入符号"},
            New String() {"[平面]角/秒", "]角]秒", "″", "插入符号"},
            New String() {"体积", "升", "L", "插入符号"},
            New String() {"质量/吨", "吨", "t", "插入符号"},
            New String() {"质量/原子质量单位", "原子质量单位", "u", "插入符号"},
            New String() {"旋转速度", "转每分", "r/min", "插入符号"},
            New String() {"长度", "海里", "n mile", "插入符号"},
            New String() {"速度", "节", "kn", "插入符号"},
            New String() {"能", "电子伏", "eV", "插入符号"},
            New String() {"级差", "分贝", "dB", "插入符号"},
            New String() {"线密度", "特[克斯]", "tex", "插入符号"},
            New String() {"面积", "公顷", "hm2", "插入符号"}
        }
        For i As Integer = 0 To array4.Length - 1
            Me.DataGridView4.Rows.Add(array4(i))
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

    Private Sub DataGridView2_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs)
        If (e.ColumnIndex = 3) Then
            Dim symbol = Me.DataGridView2.Rows.Item(e.RowIndex).Cells(2).Value
            InsertSymbol(symbol)
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs)
        If (e.ColumnIndex = 3) Then
            Dim symbol = Me.DataGridView3.Rows.Item(e.RowIndex).Cells(2).Value
            InsertSymbol(symbol)
        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs)
        If (e.ColumnIndex = 3) Then
            Dim symbol = Me.DataGridView4.Rows.Item(e.RowIndex).Cells(2).Value
            InsertSymbol(symbol)
        End If
    End Sub
End Class