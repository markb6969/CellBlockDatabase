Imports Mysqlx.Crud

Public Class RemoveOptionControl
    ' Form Load event handler to populate txtName
    Private Sub RemoveOptionControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Check if currentPDL is not null
        If currentPDL IsNot Nothing Then
            ' Retrieve first, last, and middle names from currentPDL DataRow
            Dim firstName As String = Convert.ToString(currentPDL("first_name"))
            Dim lastName As String = Convert.ToString(currentPDL("last_name"))
            Dim middleName As String = Convert.ToString(currentPDL("middle_name"))

            ' Handle middle name being possibly empty
            If String.IsNullOrEmpty(middleName) Then
                txtName.Text = $"{firstName} {lastName}"
            Else
                txtName.Text = $"{firstName} {middleName} {lastName}"
            End If
        End If

        txtID.Text = Convert.ToString(currentPDL("pdl_id"))
    End Sub

    ' Other button click event handlers
    Private Sub btnUpdateInmate_Click(sender As Object, e As EventArgs) Handles btnUpdateInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToUpdateInmateControl()
        End If

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)

        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToUpdateInmateControl()
        End If

    End Sub

    Private Sub btnRemoveInmate_Click(sender As Object, e As EventArgs) Handles btnRemoveInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToRemoveInmateControl()
        End If

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)

        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToRemoveInmateControl()
        End If
    End Sub

    Private Sub btnAddInmate_Click(sender As Object, e As EventArgs) Handles btnAddInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToAddInmateControl()
        End If

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)

        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToAddInmateControl()
        End If
    End Sub



    Private Sub btnRelease_Click(sender As Object, e As EventArgs) Handles btnRelease.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToReleaseInmateControl()
        End If

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)

        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToReleaseInmateControl()
        End If
    End Sub
End Class
