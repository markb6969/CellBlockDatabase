Public Class AdminVisitorApprove
    Private Sub btnApprove_Click(sender As Object, e As EventArgs) Handles btnApprove.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminVisitorStatusControl()
    End Sub

    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminVisitorStatusControl()
    End Sub
End Class
