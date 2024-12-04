Imports System.IO

Public Class UpdateInfoControl
    ' Method to load the current user data into the fields
    Private Sub LoadUserData()
        ' Populate text fields with currentUser data from DataRow
        txtFirstName.Text = currentUser("first_name").ToString()
        txtMiddleName.Text = currentUser("middle_name").ToString()
        txtLastName.Text = currentUser("last_name").ToString()
        txtSuffix.Text = currentUser("suffix").ToString()
        txtPhone.Text = currentUser("phone_num").ToString()
        txtZip.Text = currentUser("zip").ToString()
        txtCountry.Text = currentUser("country").ToString()
        txtCity.Text = currentUser("city").ToString()
        txtRegion.Text = currentUser("region").ToString()
        txtMunicipality.Text = currentUser("municipality").ToString()
        txtStreet.Text = currentUser("street").ToString()

        ' Set DateTime field (Date of Birth) with the value from the DataRow
        If currentUser("date_of_birth") IsNot DBNull.Value Then
            dtBirthdate.Value = Convert.ToDateTime(currentUser("date_of_birth"))
        End If

        ' Set sex (assuming sex is stored as "M" or "F")
        If currentUser("sex").ToString() = "M" Then
            RadioButton1.Checked = True ' Male
        Else
            RadioButton2.Checked = True ' Female
        End If

        ' Set civil status (assuming civil_status is a string)
        cmbStatus.SelectedItem = currentUser("civil_status").ToString()

        ' Set profile picture (assuming it's a valid byte array)
        If currentUser("Image") IsNot DBNull.Value Then
            pbProfilePic.Image = ByteArrayToImage(DirectCast(currentUser("Image"), Byte()))
        End If
    End Sub

    ' Convert byte array to image (assuming your Image field is a byte array)
    Private Function ByteArrayToImage(byteArray As Byte()) As Image
        Using ms As New MemoryStream(byteArray)
            Return Image.FromStream(ms)
        End Using
    End Function

    ' Update user information based on input from the form
    Private Sub UpdateUserInfo()
        ' Create a new dictionary to hold the updated user data
        Dim userData As New Dictionary(Of String, Object) From {
            {"first_name", txtFirstName.Text},
            {"middle_name", txtMiddleName.Text},
            {"last_name", txtLastName.Text},
            {"suffix", txtSuffix.Text},
            {"phone_num", txtPhone.Text},
            {"zip", txtZip.Text},
            {"country", txtCountry.Text},
            {"city", txtCity.Text},
            {"region", txtRegion.Text},
            {"municipality", txtMunicipality.Text},
            {"street", txtStreet.Text},
            {"date_of_birth", dtBirthdate.Value},
            {"sex", If(RadioButton1.Checked, "M", "F")},
            {"civil_status", cmbStatus.SelectedItem.ToString()},
            {"Image", ImageToByteArray(pbProfilePic.Image)}
        }

        ' Create a dictionary to hold the conditions for the update (e.g., user id)
        Dim conditions As New Dictionary(Of String, Object) From {
            {"id", currentUser("id")}
        }

        ' Call the UpdateRecord function from your module
        UpdateRecord("accounts", userData, conditions)
    End Sub

    ' Convert image to byte array (for saving profile picture)
    Private Function ImageToByteArray(image As Image) As Byte()
        Using ms As New MemoryStream()
            image.Save(ms, image.RawFormat)
            Return ms.ToArray()
        End Using
    End Function

    ' Call LoadUserData when the form loads (or at an appropriate point in the lifecycle)
    Private Sub UpdateInfoControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadUserData()
    End Sub

    ' Handle the save button click event
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Log the user update action (if needed)
        Logs("Super Admin info updated: ID " + currentUser("id").ToString())
        MsgBox("Update success")
        UpdateUserInfo()
    End Sub
End Class
