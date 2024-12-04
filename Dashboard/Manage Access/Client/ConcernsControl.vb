Public Class ConcernsControl
    Dim sql = "SELECT * FROM concerns"
    Private Sub ConcernsControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Logs("opened concerns information view")
        LoadToDGVForDisplay(sql, dgvReportedConcerns)
    End Sub

End Class
