Public Class UpdateMedicalinfoControl
    ' Assuming currentMedical is a global variable holding the current medical data as a DataRow
    Private Sub LoadCurrentMedicalInfo()
        ' Check if currentMedical contains data
        If currentMedical IsNot Nothing Then
            ' Populate the form fields with the current medical data
            txtMedicalAddress.Text = currentMedical("done_at").ToString()
            txtDoctorName.Text = currentMedical("doctor").ToString()
            txtChronicIllness.Text = currentMedical("chronic_illnesses").ToString()
            txtAlergies.Text = currentMedical("allergies").ToString()
            cmbMental.Text = currentMedical("mental_health_status").ToString()

            If currentMedical("psychiatric_treatment_required") = 0 Then
                cbNoTreatment.Checked = True
                cbYesTreatment.Checked = False
            ElseIf currentMedical("psychiatric_treatment_required") = 1 Then
                cbYesSelfharm.Checked = True
                cbNoSelfharm.Checked = False
            End If

            If currentMedical("risk_of_self_harm") = 0 Then
                cbNoSelfharm.Checked = True
                cbYesSelfharm.Checked = False
            ElseIf currentMedical("risk_of_self_harm") = 1 Then
                cbYesSelfharm.Checked = True
                cbNoSelfharm.Checked = False
            End If

        Else
                MsgBox("No medical data found to load.", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub UpdateMedicalinfoControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Call the method to load the current medical info when the form loads
        LoadCurrentMedicalInfo()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Show confirmation message before proceeding
        Dim result As MsgBoxResult = MsgBox("Are you sure you want to save the changes?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Confirm Save")

        If result = MsgBoxResult.Yes Then

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

            medData.pdl_id = selectedItem

            Dim condition As New Dictionary(Of String, Object) From {
                {"pdl_id", selectedItem}
            }

            ' Update the record
            UpdateRecord(medical, medData.GetDictionary, condition)
            Logs("Medcial record updated: medicalID" + crimeCase.pdl_id)

            ' Display success message
            MsgBox("Medical information updated successfully.", MsgBoxStyle.Information)

            ' Switch to the home control after saving the data
            Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
            If mainForm IsNot Nothing Then
                selectedItem = 0
                mainForm.SwitchToInmateHomeControl()
            End If
        Else
            ' If the user chose "No", do nothing (just return)
            Return
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub
End Class
