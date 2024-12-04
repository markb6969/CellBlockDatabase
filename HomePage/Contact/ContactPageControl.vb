Imports System.Security.AccessControl

Public Class ContactPageControl

    Dim firstName As String
    Dim email As String
    Dim phoneNumber As String
    Dim preffered As String
    Dim subject As String
    Dim message As String

    Private Sub ContactPageControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the placeholder text for the Name field
        If String.IsNullOrEmpty(txtName.Text) Then
            txtName.Text = "Last name, First name, Middle initial"
            txtName.ForeColor = Color.Gray ' Set placeholder color
        End If
    End Sub


    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        ' Get values from the fields
        firstName = txtName.Text
        email = txtEmail.Text
        phoneNumber = txtPhone.Text
        If cmbContact.SelectedItem Is Nothing OrElse cmbContact.SelectedIndex = -1 Then
            MessageBox.Show("Please select a contact preference.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        preffered = cmbContact.SelectedItem.ToString
        subject = txtSubject.Text
        message = rtxtMessage.Text

        ' Validate fields before submitting
        Dim validationError As String = ValidateFields(firstName, email, phoneNumber, subject, message, preffered)
        If String.IsNullOrEmpty(validationError) Then
            ' If validation passed, prepare the data dictionary
            Dim data As New Dictionary(Of String, Object) From {
                {"first_name", firstName},
                {"email", email},
                {"phone_number", phoneNumber},
                {"preffered", preffered},
                {"subject", subject},
                {"message", message},
                {"created_at", DateTime.Now}
            }

            ' Insert data into the database
            CreateRecord("contacts", data)
            MsgBox("Your contact information has been submitted.")

            ' Reset the fields to their initial state
            ResetFields()
        Else
            ' Show validation error message
            MessageBox.Show(validationError, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    ' Validation function to check the fields
    Private Function ValidateFields(firstName As String, email As String, phoneNumber As String, subject As String, message As String, prefs As String) As String
        ' Check if any field is empty
        If String.IsNullOrEmpty(firstName) Then
            Return "Please enter your first name."
        ElseIf String.IsNullOrEmpty(email) Then
            Return "Please enter your email address."
        ElseIf String.IsNullOrEmpty(phoneNumber) Then
            Return "Please enter your phone number."
        ElseIf String.IsNullOrEmpty(subject) Then
            Return "Please enter the subject."
        ElseIf String.IsNullOrEmpty(message) Then
            Return "Please enter your message."
        ElseIf String.IsNullOrEmpty(prefs) Then
            Return "Please enter your message."
        End If

        ' All fields are valid
        Return String.Empty
    End Function

    ' Reset the fields to their initial state
    Private Sub ResetFields()
        txtName.ForeColor = Color.Gray ' Reset placeholder color
        txtEmail.Text = "Email address"
        txtEmail.ForeColor = Color.Gray ' Reset placeholder color
        txtPhone.Text = "+63-1234567890"
        txtPhone.ForeColor = Color.Gray ' Reset placeholder color
        txtSubject.Text = "Subject"
        txtSubject.ForeColor = Color.Gray ' Reset placeholder color
        rtxtMessage.Text = "Your message"
        rtxtMessage.ForeColor = Color.Gray ' Reset placeholder color
        cmbContact.SelectedIndex = -1 ' Reset the ComboBox selection
    End Sub

    ' Placeholder functionality for First Name TextBox
    Private Sub txtName_Enter(sender As Object, e As EventArgs) Handles txtName.Enter
        If txtName.Text = "Last name, First name, Middle initial" Then
            txtName.Text = ""
            txtName.ForeColor = Color.Black ' Change text color to black when typing
        End If
    End Sub

    Private Sub txtName_Leave(sender As Object, e As EventArgs) Handles txtName.Leave
        If String.IsNullOrEmpty(txtName.Text) Then
            txtName.Text = "Last name, First name, Middle initial"
            txtName.ForeColor = Color.Gray ' Change text color back to gray for placeholder
        End If
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
        If txtSubject.Text = "Subject" Then
            txtSubject.Text = ""
            txtSubject.ForeColor = Color.Black ' Change text color to black when typing
        End If
    End Sub

    Private Sub txtSubject_Leave(sender As Object, e As EventArgs) Handles txtSubject.Leave
        If String.IsNullOrEmpty(txtSubject.Text) Then
            txtSubject.Text = "Subject"
            txtSubject.ForeColor = Color.Gray ' Change text color back to gray for placeholder
        End If
    End Sub

    ' Placeholder functionality for Message RichTextBox
    Private Sub rtxtMessage_Enter(sender As Object, e As EventArgs) Handles rtxtMessage.Enter
        If rtxtMessage.Text = "Your message" Then
            rtxtMessage.Text = ""
            rtxtMessage.ForeColor = Color.Black ' Change text color to black when typing
        End If
    End Sub

    Private Sub rtxtMessage_Leave(sender As Object, e As EventArgs) Handles rtxtMessage.Leave
        If String.IsNullOrEmpty(rtxtMessage.Text) Then
            rtxtMessage.Text = "Your message"
            rtxtMessage.ForeColor = Color.Gray ' Change text color back to gray for placeholder
        End If
    End Sub
    Private Sub txtPhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPhone.KeyPress
        ' Allow control keys like Backspace
        If Char.IsControl(e.KeyChar) Then
            Return
        End If

        ' Allow only digits (0-9)
        If Not Char.IsDigit(e.KeyChar) Then
            ' Cancel the key press if it's not a valid digit
            e.Handled = True
        End If
    End Sub

End Class
