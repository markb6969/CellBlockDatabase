Public Class ContactReply
    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToClientHomeControl()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToClientHomeControl()
    End Sub
End Class
