Public Class AdminSettingsControl
    Private Sub btnViewLogs_Click(sender As Object, e As EventArgs) Handles btnViewLogs.Click
        AdminMainDashboard.Show()
        Me.Hide()

    End Sub

    Private Sub btnInmateInfo_Click(sender As Object, e As EventArgs) Handles btnInmateInfo.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToInmateHomeControl()
    End Sub

    Private Sub btnUpdateInfo_Click(sender As Object, e As EventArgs) Handles btnUpdateInfo.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminInfoControl()
    End Sub

    Private Sub btnLogOut_Click(sender As Object, e As EventArgs) Handles btnLogOut.Click
        LogoutAdmin.Show()
        Me.Hide()
    End Sub
End Class
