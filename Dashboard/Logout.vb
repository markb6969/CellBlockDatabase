Public Class Logout

    Private Sub Logout_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Application.Exit()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MainDashboard.Hide()
        Me.Hide()
        Home.Show()
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        MainDashboard.Show()
        Me.Hide()
    End Sub
End Class