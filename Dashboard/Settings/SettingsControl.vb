Public Class SettingsControl
    Private Sub btnViewLogs_Click(sender As Object, e As EventArgs) Handles btnViewLogs.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToMainDashboadControl()
    End Sub

    Private Sub btnMangeAccess_Click(sender As Object, e As EventArgs) Handles btnMangeAccess.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToManageHomeControl()
    End Sub

    Private Sub btnInmateInfo_Click(sender As Object, e As EventArgs) Handles btnInmateInfo.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToInmateHomeControl()
    End Sub

    Private Sub btnUpdateInfo_Click(sender As Object, e As EventArgs) Handles btnUpdateInfo.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToSuperAdminControl()

    End Sub

    Private Sub btnLogOut_Click(sender As Object, e As EventArgs) Handles btnLogOut.Click
        Logout.Show()
        Me.Hide()

    End Sub
End Class
