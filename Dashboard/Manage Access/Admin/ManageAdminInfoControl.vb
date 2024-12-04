Public Class ManageAdminInfoControl
    Private Sub ManageAdminInfoControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate the form fields using the DataRow values
        If currentAdmin IsNot Nothing Then
            txtFirstName.Text = If(currentAdmin("first_name") IsNot DBNull.Value, currentAdmin("first_name").ToString(), "")
            txtMiddleName.Text = If(currentAdmin("middle_name") IsNot DBNull.Value, currentAdmin("middle_name").ToString(), "")
            txtLastName.Text = If(currentAdmin("last_name") IsNot DBNull.Value, currentAdmin("last_name").ToString(), "")
            txtStreet.Text = If(currentAdmin("street") IsNot DBNull.Value, currentAdmin("street").ToString(), "")
            txtMunicipality.Text = If(currentAdmin("municipality") IsNot DBNull.Value, currentAdmin("municipality").ToString(), "")
            txtCity.Text = If(currentAdmin("city") IsNot DBNull.Value, currentAdmin("city").ToString(), "")
            txtRegion.Text = If(currentAdmin("region") IsNot DBNull.Value, currentAdmin("region").ToString(), "")
            txtZip.Text = If(currentAdmin("zip") IsNot DBNull.Value, currentAdmin("zip").ToString(), "")
            txtCountry.Text = If(currentAdmin("country") IsNot DBNull.Value, currentAdmin("country").ToString(), "")
            txtPhone.Text = If(currentAdmin("phone_num") IsNot DBNull.Value, currentAdmin("phone_num").ToString(), "")

            txtSuffix.Text = If(currentAdmin("suffix") IsNot DBNull.Value, currentAdmin("suffix").ToString(), "")
            cmbStatus.Text = currentAdmin("civil_status").ToString

            ' Gender based on radio button selection
            If currentAdmin("sex") IsNot DBNull.Value Then
                If currentAdmin("sex").ToString() = "Male" Then
                    RadioButton1.Checked = True
                ElseIf currentAdmin("sex").ToString() = "Female" Then
                    RadioButton2.Checked = True
                End If
            End If

            txtPhone.Text = "09106891723"

            ' If profile picture is available, set it to the PictureBox
            If currentAdmin("image") IsNot DBNull.Value Then
                ' Assuming ByteToImage is your function to convert the byte array to an Image
                pbProfilePic.Image = ByteArrayToImage(currentAdmin("image"))
            End If

            ' If birthdate is available
            If currentAdmin("date_of_birth") IsNot DBNull.Value Then
                dtBirthdate.Value = Convert.ToDateTime(currentAdmin("date_of_birth"))
            End If
        End If
    End Sub

    ' Save the data if needed
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Create a new instance of the UserData class
        Dim user As New UserData()

        ' Populate the UserData instance with the form values
        user.Id = currentAdmin("id") ' Assuming you're getting the current admin ID from somewhere (e.g., DataRow)
        user.First_Name = txtFirstName.Text
        user.Middle_Name = txtMiddleName.Text
        user.Last_Name = txtLastName.Text
        user.Date_Of_Birth = dtBirthdate.Value
        user.Suffix = txtSuffix.Text
        user.Sex = If(RadioButton1.Checked, "Male", "Female") ' Assuming you have Male/Female radio buttons
        user.Civil_Status = cmbStatus.SelectedItem.ToString() ' Assuming cmbStatus is a ComboBox with civil status options
        user.Street = txtStreet.Text
        user.Municipality = txtMunicipality.Text
        user.zip = txtZip.Text
        user.region = txtRegion.Text
        user.city = txtCity.Text
        user.country = txtCountry.Text
        user.user_level = "admin" ' Or whatever the user level should be

        If txtOldPass.Text.ToString.Equals(currentAdmin("password_hash").ToString) Then
            If txtNewPass.Text.Equals(txtConfirmPass.Text) Then
                user.Password_Hash = txtNewPass.Text
            Else
                MsgBox("passwords don't match")
                Return
            End If
        Else
            MsgBox("old password is incorrect")
            Return
        End If


        ' Convert Profile Picture to Byte Array
        If pbProfilePic.Image IsNot Nothing Then
            user.Image = ImageToByteArray(pbProfilePic.Image) ' Assuming you have an ImageToByte function to convert the Image to byte array
        End If

        Dim condition As New Dictionary(Of String, Object) From {
            {"id", currentAdmin("id")}
        }

        ' For example, if you're using a DatabaseModule:
        UpdateRecord(accounts, user.GetDictionary, condition)
        Logs("UpdateDefaultButton admin account: AdminID" + currentAdmin("id").ToString)
        MsgBox("Update successfull")

        ' After saving the data, switch to the Admin Home Control
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToAdminHomeControl()
    End Sub



End Class
