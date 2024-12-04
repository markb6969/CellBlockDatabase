Public Class CreatePasswordControl

    ' Set initial placeholder values when the form loads
    Private Sub CreatePasswordControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set placeholder text for the input fields
        txtEmail.Text = "Email Address"
        txtUsername.Text = "Username"
        txtPassword.Text = "Password"
        txtConfirmPassword.Text = "Confirm Password"
    End Sub

    ' Clear placeholder text when the user clicks into the text box
    Private Sub txtEmail_Enter(sender As Object, e As EventArgs) Handles txtEmail.Enter
        If txtEmail.Text = "Email Address" Then
            txtEmail.Clear() ' Clear the placeholder text
        End If
    End Sub

    Private Sub txtUsername_Enter(sender As Object, e As EventArgs) Handles txtUsername.Enter
        If txtUsername.Text = "Username" Then
            txtUsername.Clear()
        End If
    End Sub

    Private Sub txtPassword_Enter(sender As Object, e As EventArgs) Handles txtPassword.Enter
        If txtPassword.Text = "Password" Then
            txtPassword.Clear()
            txtPassword.PasswordChar = "*" ' Show the password as masked
        End If
    End Sub

    Private Sub txtConfirmPassword_Enter(sender As Object, e As EventArgs) Handles txtConfirmPassword.Enter
        If txtConfirmPassword.Text = "Confirm Password" Then
            txtConfirmPassword.Clear()
            txtConfirmPassword.PasswordChar = "*" ' Show the password as masked
        End If
    End Sub

    ' Restore placeholder text if the user leaves the text box empty
    Private Sub txtEmail_Leave(sender As Object, e As EventArgs) Handles txtEmail.Leave
        If String.IsNullOrWhiteSpace(txtEmail.Text) Then
            txtEmail.Text = "Email Address" ' Restore placeholder text
        End If
    End Sub

    Private Sub txtUsername_Leave(sender As Object, e As EventArgs) Handles txtUsername.Leave
        If String.IsNullOrWhiteSpace(txtUsername.Text) Then
            txtUsername.Text = "Username"
        End If
    End Sub

    Private Sub txtPassword_Leave(sender As Object, e As EventArgs) Handles txtPassword.Leave
        If String.IsNullOrWhiteSpace(txtPassword.Text) Then
            txtPassword.Text = "Password"
            txtPassword.PasswordChar = "" ' Hide the password text when placeholder is shown
        End If
    End Sub

    Private Sub txtConfirmPassword_Leave(sender As Object, e As EventArgs) Handles txtConfirmPassword.Leave
        If String.IsNullOrWhiteSpace(txtConfirmPassword.Text) Then
            txtConfirmPassword.Text = "Confirm Password"
            txtConfirmPassword.PasswordChar = "" ' Hide the password text when placeholder is shown
        End If
    End Sub

    ' Check if all required fields are filled, ignoring placeholder text
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim existingAccs = GetTableData("accounts")
        ' Check if any field is empty or holds placeholder text
        If String.IsNullOrWhiteSpace(txtEmail.Text) OrElse
           txtEmail.Text = "Email Address" OrElse
           String.IsNullOrWhiteSpace(txtUsername.Text) OrElse
           txtUsername.Text = "Username" OrElse
           String.IsNullOrWhiteSpace(txtPassword.Text) OrElse
           txtPassword.Text = "Password" OrElse
           String.IsNullOrWhiteSpace(txtConfirmPassword.Text) OrElse
           txtConfirmPassword.Text = "Confirm Password" OrElse
           String.IsNullOrWhiteSpace(cmbControl.Text) Then

            ' Show a message if any required field is empty or still holds placeholder values
            MessageBox.Show("Please fill in all the required fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Check if passwords match
        If txtPassword.Text <> txtConfirmPassword.Text Then
            MessageBox.Show("Passwords do not match. Please try again.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Check if the username already exists in the table
        Dim foundRows() As DataRow = existingAccs.Select("username = '" & txtUsername.Text.Trim & "'")

        If foundRows.Length > 0 Then
            ' If a row is found, that means the username already exists
            MessageBox.Show("Username already exists. Please choose a different username.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Assign values to accData, replacing the placeholders with Null or empty string if necessary
        accData.Email = If(txtEmail.Text = "Email Address", Nothing, txtEmail.Text)
        accData.Username = If(txtUsername.Text = "Username", Nothing, txtUsername.Text)
        accData.Password_Hash = If(txtPassword.Text = "Password", Nothing, txtPassword.Text)
        accData.user_level = cmbControl.Text

        ' Switch to the next control
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToAccountImageControl()
    End Sub
End Class
