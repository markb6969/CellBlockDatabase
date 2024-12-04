Public Class AdminClientHome
    Private Sub btnVisitor_Click(sender As Object, e As EventArgs) Handles btnVisitor.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminVisitorStatusControl()
    End Sub

    Private Sub btnReportConcerns_Click(sender As Object, e As EventArgs) Handles btnReportConcerns.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminConcernControl()
    End Sub

    Private Sub btnContact_Click(sender As Object, e As EventArgs) Handles btnContact.Click
        Dim mainForm As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToAdminContactControl()
    End Sub
End Class
