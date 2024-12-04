Public Class SuperAdminControl
    ' Called when the form is loaded
    Private Sub SuperAdminControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate fields with currentUser data
        If currentUser IsNot Nothing Then
            txtFirstName.Text = currentUser("first_name").ToString()
            txtLastName.Text = currentUser("last_name").ToString()
            txtUsername.Text = currentUser("username").ToString()
            txtEmail.Text = currentUser("email").ToString()
            txtPhoneNum.Text = currentUser("phone_num").ToString()
            lblShowFullName.Text = currentUser("last_name").ToString() & ", " & currentUser("first_name").ToString()

            ' Combine address fields to create full address
            Dim fullAddress As String = String.Join(", ", New String() {
                currentUser("municipality").ToString(),
                currentUser("region").ToString(),
                currentUser("city").ToString(),
                currentUser("zip").ToString(),
                currentUser("street").ToString(),
                currentUser("country").ToString()
            })

            lblShowAdress.Text = fullAddress
        End If
    End Sub

    Private Sub btnAdmin_Click(sender As Object, e As EventArgs) Handles btnAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToAdminHomeControl()
    End Sub

    Private Sub btnClient_Click(sender As Object, e As EventArgs) Handles btnClient.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToClientHomeControl()
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToManageHomeControl()
    End Sub

    Private Sub btnChangePass_Click(sender As Object, e As EventArgs) Handles btnChangePass.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToChangePassControl()
    End Sub

    Private Sub btnUpdateInfo_Click(sender As Object, e As EventArgs) Handles btnUpdateInfo.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToUpdateSuperAdminControl()
    End Sub
End Class
