Public Class AdminVisitorControl
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminVisitorApproveControl()
    End Sub
End Class
