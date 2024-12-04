Public Class ChangePasswordControl
    Private Sub btnSuperAdmin_Click(sender As Object, e As EventArgs) Handles btnSuperAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToSuperAdminControl()
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

    Private Sub btnAccountDetails_Click(sender As Object, e As EventArgs) Handles btnAccountDetails.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToSuperAdminControl()
    End Sub

    Private Sub btnUpdateInfo_Click(sender As Object, e As EventArgs) Handles btnUpdateInfo.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToUpdateSuperAdminControl()
    End Sub

    ' Method to update the password in the database
    Private Sub UpdatePassword(userId As String, newPasswordHash As String)
        ' Prepare the data dictionary with the new password hash
        Dim data As New Dictionary(Of String, Object) From {
            {"password_hash", newPasswordHash}
        }

        ' Prepare the conditions dictionary to match the user ID
        Dim conditions As New Dictionary(Of String, Object) From {
            {"id", userId}
        }

        ' Call the UpdateRecord method to update the password hash
        UpdateRecord("accounts", data, conditions)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Check if the old password matches the current password hash
        If txtOldPass.Text <> currentUser("password_hash").ToString() Then
            MessageBox.Show("Old password does not match", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Check if the new password and confirm password match
        If txtNewPass.Text <> txtConfirmPass.Text Then
            MessageBox.Show("New password and confirm password do not match", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Optionally, update the password hash in the database here
        ' Assuming you have a method to update the password in the database (e.g., UpdatePassword)
        Dim newPasswordHash As String = txtNewPass.Text ' Ideally hash the new password
        UpdatePassword(currentUser("id").ToString(), newPasswordHash)
        Logs("Super admin password updated: ID" + currentUser("id"))

        MessageBox.Show("Password updated successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class
