Public Class InmateReportControls
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToInmatePageControl()
    End Sub
End Class
