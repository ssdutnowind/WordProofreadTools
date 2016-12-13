'TODO:   按照以下步骤启用功能区(XML)项: 

'1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon()
'End Function

'2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
'   操作(例如单击按钮)。注意: 如果已经从功能区设计器中导出此功能区，
'   请将代码从事件处理程序移动到回调方法，并
'   修改该代码以使用功能区扩展性(RibbonX)编程模型。

'3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。

'有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。

Imports Microsoft.Office.Core

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI
    Private WithEvents loginForm As FormLogin = CommonModule.loginForm

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("WordProofreadTools.Ribbon.xml")
    End Function

#Region "可用性回调"



    ''' <summary>
    ''' 控制内置Ribbon Tab是否可见
    ''' </summary>
    Private buildinVisible As Boolean = True
    Public Function GetBuildinVisible(control As IRibbonControl) As Boolean
        Return buildinVisible
    End Function

    ''' <summary>
    ''' 设置内置Tab显示
    ''' </summary>
    Public Sub SetBuildinVisible()
        Me.buildinVisible = True
        Me.ribbon.Invalidate()
    End Sub

    ''' <summary>
    ''' 设置内置Tab隐藏
    ''' </summary>
    Public Sub SetBuildinHidden()
        Me.buildinVisible = False
        Me.ribbon.Invalidate()
    End Sub

    ''' <summary>
    ''' 任务窗口
    ''' </summary>
    Private taskWindowChecked As Boolean = False
    Public Function GetTaskWindowChecked(control As IRibbonControl) As Boolean
        Return taskWindowChecked
    End Function

    ''' <summary>
    ''' 设置内置任务窗口是否显示
    ''' </summary>
    ''' <param name="checked"></param>
    Public Sub SetTaskWindowChecked(ByVal checked As Boolean)
        taskWindowChecked = checked
        Me.ribbon.Invalidate()
    End Sub

    ''' <summary>
    ''' 登录窗口
    ''' </summary>
    Private loginEnable As Boolean = True
    Public Function GetLoginEnabled(control As IRibbonControl) As Boolean
        Return loginEnable
    End Function

    Public Function GetLogoutEnabled(control As IRibbonControl) As Boolean
        Return Not loginEnable

    End Function

    ''' <summary>
    ''' 设置登录是否可用
    ''' </summary>
    ''' <param name="enable"></param>
    Public Sub SetLoginEnabled(enable As Boolean)
        loginEnable = enable
        Me.ribbon.Invalidate()
    End Sub

#End Region


#Region "功能区回调"
    '在此创建回调方法。有关添加回调方法的详细信息，请访问 http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    ''' <summary>
    ''' 登录
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonLogin_Click(ByVal control As Office.IRibbonControl)
        CommonModule.loginForm.ShowDialog()
    End Sub

    ''' <summary>
    ''' 注销
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub BtnRibbonLogout_Click(ByVal control As Office.IRibbonControl)
        'TODO: 一些检查
        If (CommonModule.ShowAlert("确定要退出登录？", "Question") = vbYes) Then
            'TODO: 数据清理
            Me.SetLoginEnabled(True)
            ' 显示内置功能区
            Me.SetBuildinVisible()
        End If
    End Sub


    Public Sub CbxRobbinTaskWindow_Click(ByVal control As Office.IRibbonControl, ByVal isChecked As Boolean)
        If isChecked Then
            CommonModule.ShowCloudFilesPane()
        Else
            CommonModule.HideCloudFilesPane()
        End If
    End Sub

#End Region


#Region "事件"

    Public Sub OnUserLogin(ByVal sender As Object) Handles loginForm.UserLogin
        CommonModule.Log("Ribbon：用户登录")
        ' 隐藏内置功能区
        Me.SetBuildinHidden()
        ' 登录
        Me.SetLoginEnabled(False)

    End Sub

#End Region


#Region "帮助器"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
