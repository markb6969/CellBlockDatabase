Public Class CreateAccountControl
    ' Initial placeholders in the text boxes
    Private Sub CreateAccountControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set placeholder text
        txtFirstName.Text = "First Name"
        txtLastName.Text = "Last Name"
        txtMiddleName.Text = "Middle Name"
        txtStreet.Text = "Street Address"
        txtMunicipality.Text = "Municipality"
        txtCity.Text = "City"
        txtRegion.Text = "Region"
        txtZip.Text = "Zip Code"
        txtCountry.Text = "Country"
        txtPhone.Text = "Phone Number"
        txtSuffix.Text = ""
    End Sub


    Private Sub txtPhone_Enter(sender As Object, e As EventArgs) Handles txtPhone.Enter
        If txtPhone.Text = "Phone Number" Then txtPhone.Clear()
    End Sub

    Private Sub txtPhone_Leave(sender As Object, e As EventArgs) Handles txtPhone.Leave
        If String.IsNullOrWhiteSpace(txtPhone.Text) Then txtPhone.Text = "Phone Number"
    End Sub

    ' Clear placeholder text when the user clicks into the text box
    Private Sub txtFirstName_Enter(sender As Object, e As EventArgs) Handles txtFirstName.Enter
        If txtFirstName.Text = "First Name" Then txtFirstName.Clear()
    End Sub

    Private Sub txtLastName_Enter(sender As Object, e As EventArgs) Handles txtLastName.Enter
        If txtLastName.Text = "Last Name" Then txtLastName.Clear()
    End Sub

    Private Sub txtMiddleName_Enter(sender As Object, e As EventArgs) Handles txtMiddleName.Enter
        If txtMiddleName.Text = "Middle Name" Then txtMiddleName.Clear()
    End Sub

    Private Sub txtStreet_Enter(sender As Object, e As EventArgs) Handles txtStreet.Enter
        If txtStreet.Text = "Street Address" Then txtStreet.Clear()
    End Sub

    Private Sub txtMunicipality_Enter(sender As Object, e As EventArgs) Handles txtMunicipality.Enter
        If txtMunicipality.Text = "Municipality" Then txtMunicipality.Clear()
    End Sub

    Private Sub txtCity_Enter(sender As Object, e As EventArgs) Handles txtCity.Enter
        If txtCity.Text = "City" Then txtCity.Clear()
    End Sub

    Private Sub txtRegion_Enter(sender As Object, e As EventArgs) Handles txtRegion.Enter
        If txtRegion.Text = "Region" Then txtRegion.Clear()
    End Sub

    Private Sub txtZip_Enter(sender As Object, e As EventArgs) Handles txtZip.Enter
        If txtZip.Text = "Zip Code" Then txtZip.Clear()
    End Sub

    Private Sub txtCountry_Enter(sender As Object, e As EventArgs) Handles txtCountry.Enter
        If txtCountry.Text = "Country" Then txtCountry.Clear()
    End Sub

    ' Restore placeholder text if the user leaves the text box empty
    Private Sub txtFirstName_Leave(sender As Object, e As EventArgs) Handles txtFirstName.Leave
        If String.IsNullOrWhiteSpace(txtFirstName.Text) Then txtFirstName.Text = "First Name"
    End Sub

    Private Sub txtLastName_Leave(sender As Object, e As EventArgs) Handles txtLastName.Leave
        If String.IsNullOrWhiteSpace(txtLastName.Text) Then txtLastName.Text = "Last Name"
    End Sub

    Private Sub txtMiddleName_Leave(sender As Object, e As EventArgs) Handles txtMiddleName.Leave
        If String.IsNullOrWhiteSpace(txtMiddleName.Text) Then txtMiddleName.Text = "Middle Name"
    End Sub

    Private Sub txtStreet_Leave(sender As Object, e As EventArgs) Handles txtStreet.Leave
        If String.IsNullOrWhiteSpace(txtStreet.Text) Then txtStreet.Text = "Street Address"
    End Sub

    Private Sub txtMunicipality_Leave(sender As Object, e As EventArgs) Handles txtMunicipality.Leave
        If String.IsNullOrWhiteSpace(txtMunicipality.Text) Then txtMunicipality.Text = "Municipality"
    End Sub

    Private Sub txtCity_Leave(sender As Object, e As EventArgs) Handles txtCity.Leave
        If String.IsNullOrWhiteSpace(txtCity.Text) Then txtCity.Text = "City"
    End Sub

    Private Sub txtRegion_Leave(sender As Object, e As EventArgs) Handles txtRegion.Leave
        If String.IsNullOrWhiteSpace(txtRegion.Text) Then txtRegion.Text = "Region"
    End Sub

    Private Sub txtZip_Leave(sender As Object, e As EventArgs) Handles txtZip.Leave
        If String.IsNullOrWhiteSpace(txtZip.Text) Then txtZip.Text = "Zip Code"
    End Sub

    Private Sub txtCountry_Leave(sender As Object, e As EventArgs) Handles txtCountry.Leave
        If String.IsNullOrWhiteSpace(txtCountry.Text) Then txtCountry.Text = "Country"
    End Sub

    ' Check if all fields are filled (except for Suffix and Middle Name)
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Check if required fields are empty (excluding Suffix and Middle Name)
        If String.IsNullOrWhiteSpace(txtFirstName.Text) OrElse txtFirstName.Text = "First Name" OrElse
           String.IsNullOrWhiteSpace(txtLastName.Text) OrElse txtLastName.Text = "Last Name" OrElse
           String.IsNullOrWhiteSpace(cmbStatus.Text) OrElse
           String.IsNullOrWhiteSpace(dtBirthdate.Text) OrElse
           String.IsNullOrWhiteSpace(cmbSex.Text) OrElse
           String.IsNullOrWhiteSpace(txtStreet.Text) OrElse txtStreet.Text = "Street Address" OrElse
           String.IsNullOrWhiteSpace(txtMunicipality.Text) OrElse txtMunicipality.Text = "Municipality" OrElse
           String.IsNullOrWhiteSpace(txtCity.Text) OrElse txtCity.Text = "City" OrElse
           String.IsNullOrWhiteSpace(txtRegion.Text) OrElse txtRegion.Text = "Region" OrElse
           String.IsNullOrWhiteSpace(txtZip.Text) OrElse txtZip.Text = "Zip Code" OrElse
           String.IsNullOrWhiteSpace(txtCountry.Text) OrElse txtCountry.Text = "Country" OrElse
           String.IsNullOrWhiteSpace(txtPhone.Text) OrElse txtPhone.Text = "Phone Number" Then

            ' Show a message to the user if any field is empty
            MessageBox.Show("Please fill in all the required fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' If all fields are valid, assign values to accData
        accData.First_Name = If(txtFirstName.Text = "First Name", Nothing, txtFirstName.Text)
        accData.Last_Name = If(txtLastName.Text = "Last Name", Nothing, txtLastName.Text)
        accData.Middle_Name = If(txtMiddleName.Text = "Middle Name", Nothing, txtMiddleName.Text)
        accData.Civil_Status = cmbStatus.Text
        accData.Date_Of_Birth = dtBirthdate.Text
        accData.Suffix = If(txtSuffix.Text = "", Nothing, txtSuffix.Text) ' Suffix can be left empty, treated as null
        accData.Sex = cmbSex.Text
        accData.Street = If(txtStreet.Text = "Street Address", Nothing, txtStreet.Text)
        accData.Municipality = If(txtMunicipality.Text = "Municipality", Nothing, txtMunicipality.Text)
        accData.city = If(txtCity.Text = "City", Nothing, txtCity.Text)
        accData.region = If(txtRegion.Text = "Region", Nothing, txtRegion.Text)
        accData.zip = If(txtZip.Text = "Zip Code", Nothing, txtZip.Text)
        accData.country = If(txtCountry.Text = "Country", Nothing, txtCountry.Text)
        accData.phone_num = If(txtPhone.Text = "Phone Number", Nothing, txtPhone.Text)

        ' Switch to the next control
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToAccountPasswordControl()
    End Sub

    ' Switch back to CreateAccountControl on cancel
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToCreateAccountControl()
    End Sub
End Class
