<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormLocalFiles
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.BtnClose = New System.Windows.Forms.Button()
        Me.BtnRemove = New System.Windows.Forms.Button()
        Me.BtnCheckAll = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ColumnCheck = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.ColumnFileName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColumnTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColumnPath = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColumnOptions = New System.Windows.Forms.DataGridViewButtonColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeRows = False
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColumnCheck, Me.ColumnFileName, Me.ColumnTime, Me.ColumnPath, Me.ColumnOptions})
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView1.Location = New System.Drawing.Point(12, 56)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowTemplate.Height = 23
        Me.DataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(560, 264)
        Me.DataGridView1.TabIndex = 0
        '
        'BtnClose
        '
        Me.BtnClose.Location = New System.Drawing.Point(497, 326)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 1
        Me.BtnClose.Text = "关闭"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BtnRemove
        '
        Me.BtnRemove.Enabled = False
        Me.BtnRemove.Location = New System.Drawing.Point(416, 326)
        Me.BtnRemove.Name = "BtnRemove"
        Me.BtnRemove.Size = New System.Drawing.Size(75, 23)
        Me.BtnRemove.TabIndex = 2
        Me.BtnRemove.Text = "删除"
        Me.BtnRemove.UseVisualStyleBackColor = True
        '
        'BtnCheckAll
        '
        Me.BtnCheckAll.Location = New System.Drawing.Point(335, 326)
        Me.BtnCheckAll.Name = "BtnCheckAll"
        Me.BtnCheckAll.Size = New System.Drawing.Size(75, 23)
        Me.BtnCheckAll.TabIndex = 3
        Me.BtnCheckAll.Text = "全部选择"
        Me.BtnCheckAll.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(209, 12)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "当前窗口列出了本地缓存的工作文件。"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(437, 12)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "您可以删除已经不再使用的文件以释放存储空间，也可以打开一个文件继续编辑。"
        '
        'ColumnCheck
        '
        Me.ColumnCheck.FillWeight = 60.0!
        Me.ColumnCheck.Frozen = True
        Me.ColumnCheck.HeaderText = "选择"
        Me.ColumnCheck.Name = "ColumnCheck"
        Me.ColumnCheck.Width = 40
        '
        'ColumnFileName
        '
        Me.ColumnFileName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ColumnFileName.HeaderText = "文件名"
        Me.ColumnFileName.MinimumWidth = 100
        Me.ColumnFileName.Name = "ColumnFileName"
        Me.ColumnFileName.ReadOnly = True
        '
        'ColumnTime
        '
        Me.ColumnTime.HeaderText = "修改时间"
        Me.ColumnTime.MinimumWidth = 120
        Me.ColumnTime.Name = "ColumnTime"
        Me.ColumnTime.ReadOnly = True
        Me.ColumnTime.Width = 120
        '
        'ColumnPath
        '
        Me.ColumnPath.HeaderText = "路径"
        Me.ColumnPath.Name = "ColumnPath"
        Me.ColumnPath.ReadOnly = True
        Me.ColumnPath.Visible = False
        '
        'ColumnOptions
        '
        Me.ColumnOptions.HeaderText = "操作"
        Me.ColumnOptions.Name = "ColumnOptions"
        Me.ColumnOptions.ReadOnly = True
        Me.ColumnOptions.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColumnOptions.Text = "打开"
        '
        'FormLocalFiles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 361)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnCheckAll)
        Me.Controls.Add(Me.BtnRemove)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormLocalFiles"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "协同编纂工具 - 本地文件管理"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents BtnClose As Windows.Forms.Button
    Friend WithEvents BtnRemove As Windows.Forms.Button
    Friend WithEvents BtnCheckAll As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents ColumnCheck As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColumnFileName As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColumnTime As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColumnPath As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColumnOptions As Windows.Forms.DataGridViewButtonColumn
End Class
