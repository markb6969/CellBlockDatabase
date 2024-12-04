Public Class LogInSuperAdmin

    Private Sub LogInSuperAdmin_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Application.Exit()

    End Sub

    Private Sub btnLogin_Click_1(sender As Object, e As EventArgs) Handles btnLogin.Click
        ' Get data from the "accounts" table
        Dim accounttable = GetTableData("accounts")

        ' Check if the username exists in the table
        Dim foundRows() As DataRow = accounttable.Select("username = '" & txtUsername.Text & "'")

        If foundRows.Length = 0 Then
            ' If no rows are found, display a message box
            MessageBox.Show("Username not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' If the username is found, get the DataRow for the user
        Dim userRow As DataRow = foundRows(0)

        ' Check if the password_hash matches the input password
        If userRow("password_hash").ToString() = txtPassword.Text Then
            currentUser = userRow

            CurrentLoggedUser = New LoggedUser With {
    .id = currentUser("id"),
    .name = currentUser("last_name").ToString() & ", " & currentUser("first_name").ToString() & " " & currentUser("middle_name").ToString(),
    .position = currentUser("user_level").ToString(),
    .username = currentUser("username").ToString(),
    .password = currentUser("password_hash").ToString()
}

            ' If the password matches, show the main dashboard
            Logs("User has logged in")
            MainDashboard.Show()
            Me.Hide()
        Else
            ' If the password doesn't match, show an error message
            MessageBox.Show("Incorrect password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

End Class