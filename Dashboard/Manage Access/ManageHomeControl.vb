Public Class ManageHomeControl
    Private Sub ManageHomeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Perform any initialization or data loading tasks here
        Dim sql = "SELECT * FROM accounts WHERE user_level = 'Admin'"
        LoadToDGV(sql, dgvAdmin)
        Dim sqlLogs As String = "SELECT * FROM logs WHERE DATE(dt) = CURDATE()"
        LoadToDGV(sqlLogs, dgvLogs)

    End Sub

    Private Sub btnSuperAdmin_Click(sender As Object, e As EventArgs) Handles btnSuperAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToSuperAdminControl()
    End Sub

    Private Sub btnAdmin_Click(sender As Object, e As EventArgs) Handles btnAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToAdminHomeControl()
    End Sub

    Private Sub btnClient_Click(sender As Object, e As EventArgs) Handles btnClient.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToClientHomeControl()
    End Sub

End Class
