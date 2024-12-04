Public Class LogoutAdmin
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Application.Exit()

    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        AdminMainDashboard.Show()
        Me.Hide()

    End Sub

    Private Sub LogoutAdmin_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        AdminMainDashboard.Close()
        Me.Hide()
        Home.Show()
    End Sub
End Class