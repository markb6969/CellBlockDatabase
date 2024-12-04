Public Class ContactedControl
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToContactRecieveControl()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToClientHomeControl()
    End Sub
End Class
