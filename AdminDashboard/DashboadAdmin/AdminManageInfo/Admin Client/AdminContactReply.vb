Public Class AdminContactReply
    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminClientHomeControl()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminClientHomeControl()
    End Sub
End Class
