Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.TabProofread = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.BtnRibbonLogin = Me.Factory.CreateRibbonButton
        Me.BtnRibbonLogout = Me.Factory.CreateRibbonButton
        Me.BtnRibbonSetting = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.BtnRibbonExportContents = Me.Factory.CreateRibbonButton
        Me.BtnRibbonExportIndex = Me.Factory.CreateRibbonButton
        Me.BtnRibbonExprotAbstract = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.GayRibbonEditChange = Me.Factory.CreateRibbonGallery
        Me.GayRibbonEditPositoin = Me.Factory.CreateRibbonGallery
        Me.GayRibbonEditSpace = Me.Factory.CreateRibbonGallery
        Me.GayRibbonEditOther = Me.Factory.CreateRibbonGallery
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.CbxRobbinCloudFiles = Me.Factory.CreateRibbonCheckBox
        Me.Tab1.SuspendLayout()
        Me.TabProofread.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'TabProofread
        '
        Me.TabProofread.Groups.Add(Me.Group2)
        Me.TabProofread.Groups.Add(Me.Group3)
        Me.TabProofread.Groups.Add(Me.Group4)
        Me.TabProofread.Groups.Add(Me.Group5)
        Me.TabProofread.Label = "编校工具"
        Me.TabProofread.Name = "TabProofread"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.BtnRibbonLogin)
        Me.Group2.Items.Add(Me.BtnRibbonLogout)
        Me.Group2.Items.Add(Me.BtnRibbonSetting)
        Me.Group2.Label = "用户账户"
        Me.Group2.Name = "Group2"
        '
        'BtnRibbonLogin
        '
        Me.BtnRibbonLogin.Image = CType(resources.GetObject("BtnRibbonLogin.Image"), System.Drawing.Image)
        Me.BtnRibbonLogin.Label = "账户登录"
        Me.BtnRibbonLogin.Name = "BtnRibbonLogin"
        Me.BtnRibbonLogin.ShowImage = True
        '
        'BtnRibbonLogout
        '
        Me.BtnRibbonLogout.Enabled = False
        Me.BtnRibbonLogout.Image = CType(resources.GetObject("BtnRibbonLogout.Image"), System.Drawing.Image)
        Me.BtnRibbonLogout.Label = "账户退出"
        Me.BtnRibbonLogout.Name = "BtnRibbonLogout"
        Me.BtnRibbonLogout.ShowImage = True
        '
        'BtnRibbonSetting
        '
        Me.BtnRibbonSetting.Enabled = False
        Me.BtnRibbonSetting.Image = CType(resources.GetObject("BtnRibbonSetting.Image"), System.Drawing.Image)
        Me.BtnRibbonSetting.Label = "设置"
        Me.BtnRibbonSetting.Name = "BtnRibbonSetting"
        Me.BtnRibbonSetting.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.BtnRibbonExportContents)
        Me.Group3.Items.Add(Me.BtnRibbonExportIndex)
        Me.Group3.Items.Add(Me.BtnRibbonExprotAbstract)
        Me.Group3.Label = "导出"
        Me.Group3.Name = "Group3"
        '
        'BtnRibbonExportContents
        '
        Me.BtnRibbonExportContents.Image = CType(resources.GetObject("BtnRibbonExportContents.Image"), System.Drawing.Image)
        Me.BtnRibbonExportContents.Label = "导出目录"
        Me.BtnRibbonExportContents.Name = "BtnRibbonExportContents"
        Me.BtnRibbonExportContents.ShowImage = True
        '
        'BtnRibbonExportIndex
        '
        Me.BtnRibbonExportIndex.Image = CType(resources.GetObject("BtnRibbonExportIndex.Image"), System.Drawing.Image)
        Me.BtnRibbonExportIndex.Label = "导出索引"
        Me.BtnRibbonExportIndex.Name = "BtnRibbonExportIndex"
        Me.BtnRibbonExportIndex.ShowImage = True
        '
        'BtnRibbonExprotAbstract
        '
        Me.BtnRibbonExprotAbstract.Image = CType(resources.GetObject("BtnRibbonExprotAbstract.Image"), System.Drawing.Image)
        Me.BtnRibbonExprotAbstract.Label = "导出摘要"
        Me.BtnRibbonExprotAbstract.Name = "BtnRibbonExprotAbstract"
        Me.BtnRibbonExprotAbstract.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.GayRibbonEditChange)
        Me.Group4.Items.Add(Me.GayRibbonEditPositoin)
        Me.Group4.Items.Add(Me.GayRibbonEditSpace)
        Me.Group4.Items.Add(Me.GayRibbonEditOther)
        Me.Group4.Label = "校对"
        Me.Group4.Name = "Group4"
        '
        'GayRibbonEditChange
        '
        Me.GayRibbonEditChange.Image = CType(resources.GetObject("GayRibbonEditChange.Image"), System.Drawing.Image)
        RibbonDropDownItemImpl1.Image = CType(resources.GetObject("RibbonDropDownItemImpl1.Image"), System.Drawing.Image)
        RibbonDropDownItemImpl1.Label = "改正"
        RibbonDropDownItemImpl2.Image = CType(resources.GetObject("RibbonDropDownItemImpl2.Image"), System.Drawing.Image)
        RibbonDropDownItemImpl2.Label = "删除"
        RibbonDropDownItemImpl3.Image = CType(resources.GetObject("RibbonDropDownItemImpl3.Image"), System.Drawing.Image)
        RibbonDropDownItemImpl3.Label = "增补"
        RibbonDropDownItemImpl4.Image = CType(resources.GetObject("RibbonDropDownItemImpl4.Image"), System.Drawing.Image)
        RibbonDropDownItemImpl4.Label = "改正上下角"
        Me.GayRibbonEditChange.Items.Add(RibbonDropDownItemImpl1)
        Me.GayRibbonEditChange.Items.Add(RibbonDropDownItemImpl2)
        Me.GayRibbonEditChange.Items.Add(RibbonDropDownItemImpl3)
        Me.GayRibbonEditChange.Items.Add(RibbonDropDownItemImpl4)
        Me.GayRibbonEditChange.Label = "字符改动"
        Me.GayRibbonEditChange.Name = "GayRibbonEditChange"
        Me.GayRibbonEditChange.ShowImage = True
        '
        'GayRibbonEditPositoin
        '
        Me.GayRibbonEditPositoin.Image = CType(resources.GetObject("GayRibbonEditPositoin.Image"), System.Drawing.Image)
        Me.GayRibbonEditPositoin.Label = "方向位置"
        Me.GayRibbonEditPositoin.Name = "GayRibbonEditPositoin"
        Me.GayRibbonEditPositoin.ShowImage = True
        '
        'GayRibbonEditSpace
        '
        Me.GayRibbonEditSpace.Image = CType(resources.GetObject("GayRibbonEditSpace.Image"), System.Drawing.Image)
        Me.GayRibbonEditSpace.Label = "字符间距"
        Me.GayRibbonEditSpace.Name = "GayRibbonEditSpace"
        Me.GayRibbonEditSpace.ShowImage = True
        '
        'GayRibbonEditOther
        '
        Me.GayRibbonEditOther.Image = CType(resources.GetObject("GayRibbonEditOther.Image"), System.Drawing.Image)
        Me.GayRibbonEditOther.Label = "其它"
        Me.GayRibbonEditOther.Name = "GayRibbonEditOther"
        Me.GayRibbonEditOther.ShowImage = True
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.CbxRobbinCloudFiles)
        Me.Group5.Label = "窗口"
        Me.Group5.Name = "Group5"
        '
        'CbxRobbinCloudFiles
        '
        Me.CbxRobbinCloudFiles.Label = "云端文件"
        Me.CbxRobbinCloudFiles.Name = "CbxRobbinCloudFiles"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.TabProofread)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.TabProofread.ResumeLayout(False)
        Me.TabProofread.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents TabProofread As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnRibbonLogin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnRibbonLogout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnRibbonSetting As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnRibbonExportContents As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnRibbonExportIndex As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BtnRibbonExprotAbstract As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GayRibbonEditChange As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents GayRibbonEditPositoin As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents GayRibbonEditSpace As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents GayRibbonEditOther As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents CbxRobbinCloudFiles As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
