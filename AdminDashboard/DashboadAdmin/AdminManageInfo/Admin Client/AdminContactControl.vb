Public Class AdminContactControl
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminContactReplyControl()
    End Sub
End Class
