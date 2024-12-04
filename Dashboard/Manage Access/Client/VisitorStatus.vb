Public Class VisitorStatus

    Private Sub btnPending_Click(sender As Object, e As EventArgs) Handles btnPending.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToVisitationControl()
    End Sub

    Private Sub VisitorStatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Logs("opened the visitor status view")
        ' SQL Queries
        Dim sqlReject As String = "SELECT * FROM visitors WHERE status = 'rejected'"
        Dim sqlPendingVisitors As String = "SELECT * FROM visitors WHERE status = 'pending'"
        Dim sqlApprovedVisitors As String = "SELECT * FROM visitors WHERE status = 'approved'"

        LoadToDGV(sqlReject, dgvRejected)
        LoadToDGV(sqlPendingVisitors, dgvpending)
        LoadToDGV(sqlApprovedVisitors, dgvapproved)
    End Sub


End Class
