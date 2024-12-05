Public Class ViewMedicalInfo
    Private Sub ViewMedicalInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Logs("viewing Medical")
        PopulateFields()
    End Sub

    Private Sub PopulateFields()
        ' Populate textName, textSex, textCivil, textHeight, textWeight, textPhone, textBirth, textID
        If currentPDL IsNot Nothing Then
            textName.Text = $"{Convert.ToString(currentPDL("first_name"))} " &
                        $"{Convert.ToString(currentPDL("middle_name"))} " &
                        $"{Convert.ToString(currentPDL("last_name"))}"
            textSex.Text = Convert.ToString(currentPDL("sex"))
            textCivil.Text = Convert.ToString(currentPDL("civil_status"))
            textHeight.Text = Convert.ToString(currentPDL("height"))
            textWeight.Text = Convert.ToString(currentPDL("weight"))
            textPhone.Text = Convert.ToString(currentPDL("phone_num"))
            textBirth.Text = currentPDL("date_of_birth").ToString
            textID.Text = Convert.ToString(currentPDL("pdl_id"))
            If currentCase IsNot Nothing AndAlso currentCase.Table.Columns.Contains("mugshot") AndAlso
       currentCase("mugshot") IsNot DBNull.Value AndAlso currentCase("mugshot") IsNot Nothing Then
                pbMugshot.Image = ByteArrayToImage(CType(currentCase("mugshot"), Byte()))
            Else
                pbMugshot.Image = Nothing ' Or set a default image
            End If

        End If

        ' Populate medical fields: txtMedID, txtRisk, txtDateOfMedical, txtDoctor, txtHospital, txtMental, txtIllnesses, txtAllergies
        If currentMedical IsNot Nothing Then
            txtMedID.Text = Convert.ToString(currentMedical("pdl_id"))
            txtRisk.Text = Convert.ToString(currentMedical("risk_of_self_harm"))
            txtDateOfMedical.Text = currentMedical("date_of_medical_examination").ToString
            txtDoctor.Text = Convert.ToString(currentMedical("doctor"))
            txtHospital.Text = Convert.ToString(currentMedical("done_at"))
            txtMental.Text = Convert.ToString(currentMedical("mental_health_status"))
            txtIllnesses.Text = Convert.ToString(currentMedical("chronic_illnesses"))
            txtAllergies.Text = Convert.ToString(currentMedical("allergies"))
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Logs("Exited medical view")
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        selectedItem = 0
        If mainform IsNot Nothing Then
            mainform.SwitchToInmateHomeControl()
            Return
        End If


        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        selectedItem = 0
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToInmateHomeControl()
            Return
        End If

    End Sub
End Class
