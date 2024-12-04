Public Class AdminVisitorStatus
    Private Sub btnPending_Click(sender As Object, e As EventArgs) Handles btnPending.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminVisitorControl()
    End Sub
End Class
