Public Class AddInmateControl
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Check for empty required fields
        If String.IsNullOrWhiteSpace(txtFirstName.Text) Then
            MsgBox("First Name is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtLastName.Text) Then
            MsgBox("Last Name is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(dtBirthdate.Text) Then
            MsgBox("Date of Birth is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(cmbStatus.Text) Then
            MsgBox("Civil Status is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtStreet.Text) Then
            MsgBox("Street is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtMunicipality.Text) Then
            MsgBox("Municipality is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtCity.Text) Then
            MsgBox("District/City is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtRegion.Text) Then
            MsgBox("Region is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtZip.Text) Then
            MsgBox("Zip Code is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtCountry.Text) Then
            MsgBox("Country is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtNumber.Text) Then
            MsgBox("Phone Number is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtHeight.Text) Then
            MsgBox("Height is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtWeight.Text) Then
            MsgBox("Weight is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(cmbEyeColor.Text) Then
            MsgBox("Eye Color is required.", MsgBoxStyle.Critical)
            Return
        End If

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

        CreateRecord(pdl, inmateData.GetDictionary)

        ' Collect family background details
        family.father_full_name = txtFatherName.Text
        family.father_full_address = txtFatherAddress.Text
        family.mother_full_address = txtMotherAddress.Text
        family.mother_full_name = txtMotherName.Text

        ' Collect emergency contact information
        contacts.contact_full_name = txtEmergencyName.Text
        contacts.contact_full_address = txtEmergencyAddress.Text
        contacts.relationship_with_pdl = txtRelationship.Text
        contacts.number = txtNumber.Text

        Dim data As DataTable = GetTableData(pdl)
        family.pdl_id = Convert.ToInt32(Data.Rows(Data.Rows.Count - 1)("pdl_id"))
        contacts.pdl_id = Convert.ToInt32(Data.Rows(Data.Rows.Count - 1)("pdl_id"))

        CreateRecord(family_background, family.getDictionary)
        CreateRecord(emergency_contact, contacts.getDictionary)

        ' Switch to next form if mainForm is valid
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateCriminalCaseControl()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub
End Class
