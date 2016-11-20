Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnRibbonLogin_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnRibbonLogin.Click
        Dim login = New LoginForm()
        login.ShowDialog()
    End Sub
End Class
