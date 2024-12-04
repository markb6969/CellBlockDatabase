Public Class ReportConcernControl

    Dim concernType As String
    Dim email As String
    Dim phoneNumber As String
    Dim subject As String
    Dim concernDetails As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        concernType = cmbConcern.Text
        email = txtEmail.Text
        phoneNumber = txtPhone.Text
        subject = txtSubject.Text
        concernDetails = rtxtMessage.Text

        Dim validationError As String = ValidateFields(concernType, email, phoneNumber, subject, concernDetails)
        If String.IsNullOrEmpty(validationError) Then
            ' If validation passed, prepare the data dictionary
            Dim data As New Dictionary(Of String, Object) From {
                {"concern_type", concernType},
                {"email", email},
                {"phone_number", phoneNumber},
                {"subject", subject},
                {"concern_details", concernDetails},
                {"created_at", DateTime.Now}
            }
            CreateRecord("concerns", data)
            MsgBox("Concern has been filed.")

            ' Reset the fields to their initial state
            ResetFields()
        Else
            ' Show validation error message
            MessageBox.Show(validationError, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    ' Validation function to check the fields
    Private Function ValidateFields(concernType As String, email As String, phoneNumber As String, subject As String, concernDetails As String) As String
        ' Check if any field is null, empty, or contains only white spaces
        If String.IsNullOrWhiteSpace(concernType) Then
            Return "Please select the type of concern."
        ElseIf String.IsNullOrWhiteSpace(email) Then
            Return "Please enter your email address."
        ElseIf String.IsNullOrWhiteSpace(phoneNumber) Then
            Return "Please enter your phone number."
        ElseIf String.IsNullOrWhiteSpace(subject) Then
            Return "Please enter the subject of your concern."
        ElseIf String.IsNullOrWhiteSpace(concernDetails) Then
            Return "Please enter the details of your concern."
        End If

        ' All fields are valid
        Return String.Empty
    End Function


    ' Reset the fields to their initial state
    Private Sub ResetFields()
        cmbConcern.SelectedIndex = -1 ' Reset the ComboBox selection
        txtEmail.Text = "Email address"
        txtEmail.ForeColor = Color.Gray ' Reset placeholder color
        txtPhone.Text = "+63-1234567890"
        txtPhone.ForeColor = Color.Gray ' Reset placeholder color
        txtSubject.Text = "Subject of your Concern"
        txtSubject.ForeColor = Color.Gray ' Reset placeholder color
        rtxtMessage.Text = "Your concerns"
        rtxtMessage.ForeColor = Color.Gray ' Reset placeholder color
    End Sub

    ' Placeholder functionality for Email TextBox
    Private Sub txtEmail_Enter(sender As Object, e As EventArgs) Handles txtEmail.Enter
        If txtEmail.Text = "Email address" Then
            txtEmail.Text = ""
            txtEmail.ForeColor = Color.Black ' Change text color to black when typing
        End If
    End Sub

    Private Sub txtEmail_Leave(sender As Object, e As EventArgs) Handles txtEmail.Leave
        If String.IsNullOrEmpty(txtEmail.Text) Then
            txtEmail.Text = "Email address"
            txtEmail.ForeColor = Color.Gray ' Change text color back to gray for placeholder
        End If
    End Sub

    ' Placeholder functionality for Phone Number TextBox
    Private Sub txtPhone_Enter(sender As Object, e As EventArgs) Handles txtPhone.Enter
        If txtPhone.Text = "+63-1234567890" Then
            txtPhone.Text = ""
            txtPhone.ForeColor = Color.Black ' Change text color to black when typing
        End If
    End Sub

    Private Sub txtPhone_Leave(sender As Object, e As EventArgs) Handles txtPhone.Leave
        If String.IsNullOrEmpty(txtPhone.Text) Then
            txtPhone.Text = "+63-1234567890"
            txtPhone.ForeColor = Color.Gray ' Change text color back to gray for placeholder
        End If
    End Sub

    ' Placeholder functionality for Subject TextBox
    Private Sub txtSubject_Enter(sender As Object, e As EventArgs) Handles txtSubject.Enter
        If txtSubject.Text = "Subject of your Concern" Then
            txtSubject.Text = ""
            txtSubject.ForeColor = Color.Black ' Change text color to black when typing
        End If
    End Sub

    Private Sub txtSubject_Leave(sender As Object, e As EventArgs) Handles txtSubject.Leave
        If String.IsNullOrEmpty(txtSubject.Text) Then
            txtSubject.Text = "Subject of your Concern"
            txtSubject.ForeColor = Color.Gray ' Change text color back to gray for placeholder
        End If
    End Sub

    ' Placeholder functionality for Concerns TextBox (RichTextBox in this case)
    Private Sub rtxtMessage_Enter(sender As Object, e As EventArgs) Handles rtxtMessage.Enter
        If rtxtMessage.Text = "Your concerns" Then
            rtxtMessage.Text = ""
            rtxtMessage.ForeColor = Color.Black ' Change text color to black when typing
        End If
    End Sub

    Private Sub rtxtMessage_Leave(sender As Object, e As EventArgs) Handles rtxtMessage.Leave
        If String.IsNullOrEmpty(rtxtMessage.Text) Then
            rtxtMessage.Text = "Your concerns"
            rtxtMessage.ForeColor = Color.Gray ' Change text color back to gray for placeholder
        End If
    End Sub
    Private Sub txtPhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPhone.KeyPress
        ' Check if the pressed key is a digit or a control key (e.g., backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' If it's not a digit or a control key, suppress the key press
            e.Handled = True
        End If
    End Sub

End Class
