Public Class addMedicalinfoControl
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Check if any required field is empty
        If String.IsNullOrWhiteSpace(txtMedicalAddress.Text) Then
            MsgBox("Medical Address is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtDoctorName.Text) Then
            MsgBox("Doctor Name is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(dtDateofMedical.Text) Then
            MsgBox("Date of Medical Examination is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtChronicIllness.Text) Then
            MsgBox("Chronic Illness information is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtAlergies.Text) Then
            MsgBox("Allergies information is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(cmbMental.Text) Then
            MsgBox("Mental Health Status is required.", MsgBoxStyle.Critical)
            Return
        End If

        ' Proceed with saving data if all required fields are filled
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        Dim data As DataTable = GetTableData(pdl)

        If data.Rows.Count > 0 Then
            ' Get the last record by checking the last row in the DataTable
            medData.pdl_id = Convert.ToInt32(data.Rows(data.Rows.Count - 1)("pdl_id"))
        Else
            MsgBox("No records found in the table.", MsgBoxStyle.Information)
        End If

        ' Assign the form values to medData object
        medData.done_at = txtMedicalAddress.Text
        medData.doctor = txtDoctorName.Text
        medData.date_of_medical_examination = dtDateofMedical.Text
        medData.chronic_illnesses = txtChronicIllness.Text
        medData.allergies = txtAlergies.Text

        If cbYesSelfharm.Checked Then
            medData.risk_of_self_harm = True
        Else
            medData.risk_of_self_harm = False
        End If

        If cbYesTreatment.Checked Then
            medData.pychiatric_treatment_required = True
        Else
            medData.pychiatric_treatment_required = False
        End If

        medData.mental_health_status = cmbMental.Text

        ' Save the record
        CreateRecord(medical, medData.GetDictionary)
        Logs("Created medical record: " + +inmateData.last_name + ", " + inmateData.first_name + " " + inmateData.middle_name)

        ' Switch to the next form if mainForm is valid
        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub
End Class
