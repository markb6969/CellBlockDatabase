Public Class UpdateInmateControl
    Private Sub UpdateInmateControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate fields with existing data from currentPDL, currentFamily, and currentContacts
        If currentPDL IsNot Nothing Then
            SetPlaceholder(txtFirstName, "First Name", currentPDL("first_name").ToString())
            SetPlaceholder(txtLastName, "Last Name", currentPDL("last_name").ToString())
            SetPlaceholder(txtMiddleName, "Middle Name", currentPDL("middle_name").ToString())
            If Date.TryParse(currentPDL("date_of_birth").ToString(), dtBirthdate.Value) Then
                dtBirthdate.Value = Date.Parse(currentPDL("date_of_birth").ToString())
            End If
            cmbStatus.Text = currentPDL("civil_status").ToString()
            SetPlaceholder(txtSuffix, "Suffix", currentPDL("suffix").ToString())
            If currentPDL("sex").ToString() = "Male" Then
                RadioButton1.Checked = True
            Else
                RadioButton2.Checked = True
            End If
            SetPlaceholder(txtStreet, "Street", currentPDL("street").ToString())
            SetPlaceholder(txtMunicipality, "Municipality", currentPDL("municipality").ToString())
            SetPlaceholder(txtCity, "City/District", currentPDL("district").ToString())
            SetPlaceholder(txtRegion, "Region", currentPDL("region").ToString())
            SetPlaceholder(txtZip, "Zip Code", currentPDL("zip_code").ToString())
            SetPlaceholder(txtCountry, "Country", currentPDL("country").ToString())
            SetPlaceholder(txtNumber, "Phone Number", currentPDL("phone_num").ToString())
            SetPlaceholder(txtHeight, "Height (cm)", currentPDL("height").ToString())
            SetPlaceholder(txtWeight, "Weight (kg)", currentPDL("weight").ToString())
            cmbEyeColor.Text = currentPDL("eye_color").ToString()
            SetPlaceholder(txtIdentifyingMarks, "Identifying Marks", currentPDL("location_of_identifying_marks").ToString())
            SetPlaceholder(txtPhysicalDeformities, "Physical Deformities", currentPDL("deformaties").ToString())
        End If

        If currentFamily IsNot Nothing Then
            SetPlaceholder(txtFatherName, "Father's Full Name", currentFamily("father_full_name").ToString())
            SetPlaceholder(txtFatherAddress, "Father's Address", currentFamily("father_full_address").ToString())
            SetPlaceholder(txtMotherName, "Mother's Full Name", currentFamily("mother_full_name").ToString())
            SetPlaceholder(txtMotherAddress, "Mother's Address", currentFamily("mother_full_address").ToString())
        End If

        If currentContacts IsNot Nothing Then
            SetPlaceholder(txtEmergencyName, "Emergency Contact Name", currentContacts("contact_full_name").ToString())
            SetPlaceholder(txtEmergencyAddress, "Emergency Contact Address", currentContacts("contact_full_address").ToString())
            SetPlaceholder(txtRelationship, "Relationship", currentContacts("relationship_with_pdl").ToString())
        End If

        ' Add focus event handlers for all textboxes
        AddPlaceholders()
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Ask the user if they want to continue with the update
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to continue with the update?", "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.No Then
            ' If the user selects No, exit the method and do nothing
            Return
        End If

        ' Proceed with the update if the user selects Yes
        ' Optional fields can be empty but if not, they are validated
        inmateData.first_name = txtFirstName.Text
        inmateData.last_name = txtLastName.Text
        inmateData.middle_name = txtMiddleName.Text
        inmateData.date_of_birth = dtBirthdate.Text
        inmateData.civil_status = cmbStatus.Text
        inmateData.suffix = txtSuffix.Text

        If RadioButton1.Checked Then
            inmateData.sex = "Male"
        Else
            inmateData.sex = "Female"
        End If

        inmateData.street = txtStreet.Text
        inmateData.municipality = txtMunicipality.Text
        inmateData.district = txtCity.Text
        inmateData.region = txtRegion.Text
        inmateData.zip_code = txtZip.Text
        inmateData.country = txtCountry.Text
        inmateData.phone_num = txtNumber.Text
        inmateData.height = txtHeight.Text
        inmateData.weight = txtWeight.Text
        inmateData.eye_color = cmbEyeColor.Text

        ' Handling optional identifying marks
        inmateData.identifying_mark = ""
        If cbBirthmark.Checked Then
            inmateData.identifying_mark = inmateData.identifying_mark + "Birthmark,"
        End If
        If cbMole.Checked Then
            inmateData.identifying_mark = inmateData.identifying_mark + "Mole"
        End If
        If cbScar.Checked Then
            inmateData.identifying_mark = inmateData.identifying_mark + "Scar"
        End If
        If cbTatoo.Checked Then
            inmateData.identifying_mark = inmateData.identifying_mark + "Tatoo"
        End If

        inmateData.location_of_identifying_marks = txtIdentifyingMarks.Text
        inmateData.deformaties = txtPhysicalDeformities.Text

        ' Collect family background details
        family.father_full_name = txtFatherName.Text
        family.father_full_address = txtFatherAddress.Text
        family.mother_full_address = txtMotherAddress.Text
        family.mother_full_name = txtMotherName.Text
        family.pdl_id = selectedItem

        ' Collect emergency contact information
        contacts.contact_full_name = txtEmergencyName.Text
        contacts.contact_full_address = txtEmergencyAddress.Text
        contacts.relationship_with_pdl = txtRelationship.Text
        contacts.number = txtNumber.Text
        contacts.pdl_id = selectedItem

        ' Switch to the next control
        Dim condition As New Dictionary(Of String, Object) From {
        {"pdl_id", selectedItem}
    }

        Dim pdlData = inmateData.GetDictionary
        pdlData.Add("pdl_id", selectedItem)

        UpdateRecord(pdl, pdlData, condition)
        Logs("PDL record updated: PDLID" + crimeCase.pdl_id)
        UpdateRecord(family_background, family.getDictionary, condition)
        Logs("PDL family_background record updated: PDLID" + crimeCase.pdl_id)
        UpdateRecord(emergency_contact, contacts.getDictionary, condition)
        Logs("PDL emergency record updated: PDLID" + crimeCase.pdl_id)
        MsgBox("update Successful")


        ' Switch to the update control page
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainForm IsNot Nothing Then
            selectedItem = 0
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub


    Private Sub SetPlaceholder(txtBox As TextBox, placeholder As String, value As String)
        If String.IsNullOrEmpty(value) Then
            txtBox.Text = placeholder
            txtBox.ForeColor = Color.Gray
        Else
            txtBox.Text = value
            txtBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub AddPlaceholders()
        ' Add focus event handlers for each field with a placeholder
        AddHandler txtFirstName.GotFocus, AddressOf RemovePlaceholder
        AddHandler txtFirstName.LostFocus, AddressOf ApplyPlaceholder
        ' Repeat for all other TextBoxes...
    End Sub

    Private Sub RemovePlaceholder(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = CType(sender, TextBox)
        If txtBox.ForeColor = Color.Gray Then
            txtBox.Text = ""
            txtBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ApplyPlaceholder(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = CType(sender, TextBox)
        If String.IsNullOrWhiteSpace(txtBox.Text) Then
            txtBox.Text = "Placeholder Text" ' Replace with specific placeholder
            txtBox.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub
End Class
