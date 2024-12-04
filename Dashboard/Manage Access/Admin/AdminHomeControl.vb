Public Class AdminHomeControl

    Private Sub AdminHomeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim query As String = "SELECT * FROM accounts WHERE user_level = 'admin'"
        LoadToDGVForDisplay(query, dgvAdmin)
    End Sub

    Private Sub btnSuperAdmin_Click(sender As Object, e As EventArgs) Handles btnSuperAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToSuperAdminControl()
    End Sub

    Private Sub btnClient_Click(sender As Object, e As EventArgs) Handles btnClient.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToClientHomeControl()
    End Sub

    Private Sub btnManageAdmin_Click(sender As Object, e As EventArgs) Handles btnManageAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToManageAdminControl()
    End Sub

    Private Sub dgvAdmin_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAdmin.CellContentClick

    End Sub
End Class
