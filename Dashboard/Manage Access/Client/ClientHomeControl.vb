Public Class ClientHomeControl
    Private Sub btnVisitor_Click(sender As Object, e As EventArgs) Handles btnVisitor.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        mainForm.SwitchToVisitationStatusControl()

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        mainForm.SwitchToVisitationStatusControl()

    End Sub

    Private Sub btnReportConcerns_Click(sender As Object, e As EventArgs) Handles btnReportConcerns.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToConcernsControl()
    End Sub

    Private Sub btnContact_Click(sender As Object, e As EventArgs)
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToContactControl()
    End Sub

    Private Sub ClientHomeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Logs("opened client controll")
        Dim sqlVisits As String = "SELECT * FROM visitors"
        Dim sqlContacts As String = "SELECT * FROM contacts"
        Dim sqlConcerns As String = "SELECT * FROM concerns"
        LoadToDGVForDisplay(sqlVisits, dgvVisitor)
        LoadToDGVForDisplay(sqlContacts, dgvContacts)
        LoadToDGVForDisplay(sqlConcerns, dgvConcerns)
    End Sub

    Private Sub btnContact_Click_1(sender As Object, e As EventArgs) Handles btnContact.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToContactControl()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToMakeAnnouncementControl()

        Dim adminDashboard As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToMakeAnnouncementControl()
            Return
        End If
    End Sub
End Class
