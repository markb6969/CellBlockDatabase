Public Class ReleaseInmateControl
    Private inmateDetails As New Dictionary(Of String, Object)

    Private Sub ReleaseInmateControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load currentPDL data into fields and populate the dictionary
        If currentPDL IsNot Nothing Then
            txtPdlId.Text = Convert.ToString(currentPDL("pdl_id"))
            txtFirstName.Text = Convert.ToString(currentPDL("first_name"))
            txtMiddleName.Text = Convert.ToString(currentPDL("middle_name"))
            txtInmateLastName.Text = Convert.ToString(currentPDL("last_name"))

            Dim sex As String = Convert.ToString(currentPDL("sex")).ToLower()
            If sex = "male" Then
                RadioButton1.Checked = True
            ElseIf sex = "female" Then
                RadioButton2.Checked = True
            End If

            ' Add to dictionary
            inmateDetails("pdl_id") = txtPdlId.Text
            inmateDetails("first_name") = txtFirstName.Text
            inmateDetails("middle_name") = txtMiddleName.Text
            inmateDetails("last_name") = txtInmateLastName.Text
            inmateDetails("sex") = sex
        End If
    End Sub

    Private Sub btnRelease_Click(sender As Object, e As EventArgs) Handles btnRelease.Click

        ' Validate required fields
        If String.IsNullOrWhiteSpace(txtPdlId.Text) OrElse
       String.IsNullOrWhiteSpace(txtFirstName.Text) OrElse
       String.IsNullOrWhiteSpace(txtMiddleName.Text) OrElse
       String.IsNullOrWhiteSpace(txtInmateLastName.Text) OrElse
       cmbTypeOfRelease.SelectedItem Is Nothing OrElse
       String.IsNullOrWhiteSpace(rtxtReasonRelease.Text) OrElse
       String.IsNullOrWhiteSpace(txtOfficerName.Text) OrElse
       cmbPosition.SelectedItem Is Nothing OrElse
       pbOfficerSignature.Image Is Nothing OrElse
       pbInmateSignature.Image Is Nothing Then

            MessageBox.Show("Please fill in all required fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Ask for confirmation
        Dim confirmResult As DialogResult = MessageBox.Show(
        "Are you sure you want to release this inmate?",
        "Confirm Release",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question)

        If confirmResult = DialogResult.No Then
            Return
        End If

        ' Populate the remaining fields into the dictionary
        inmateDetails("release_date") = dtReleaseDate.Value
        inmateDetails("type_of_release") = cmbTypeOfRelease.SelectedItem
        inmateDetails("reason_release") = rtxtReasonRelease.Text
        inmateDetails("officer_name") = txtOfficerName.Text
        inmateDetails("officer_position") = cmbPosition.SelectedItem
        inmateDetails("officer_signature") = ImageToByteArray(pbOfficerSignature.Image)
        inmateDetails("officer_date_signed") = dtOfficerDate.Value
        inmateDetails("inmate_signature") = ImageToByteArray(pbInmateSignature.Image)
        inmateDetails("date_signed_by_pdl") = dtpDateSignByPDL.Value

        ' Update and create records
        Dim data As New Dictionary(Of String, Object) From {
        {"status", "released"}}

        Dim condition As New Dictionary(Of String, Object) From {
        {"pdl_id", selectedItem}
    }

        UpdateRecord(pdl, data, condition)
        CreateRecord(inmatereleasedetails, inmateDetails)

        ' Display success message
        MessageBox.Show("Inmate release process completed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Navigate back to the inmate home control
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainForm IsNot Nothing Then
            Logs("inmate has been released: " + inmateDetails("last_name") + ", " + inmateDetails("first_name"))
            selectedItem = 0
            mainForm.SwitchToInmateHomeControl()
        End If

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            Logs("inmate has been released: " + inmateDetails("last_name") + ", " + inmateDetails("first_name"))
            selectedItem = 0
            adminDashboard.SwitchToInmateHomeControl()
        End If

    End Sub

    Private Sub btnOfficerSignature_Click(sender As Object, e As EventArgs) Handles btnOfficerSignature.Click
        Dim openFileDialog As New OpenFileDialog With {
            .Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp"
        }

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedImage = Image.FromFile(openFileDialog.FileName)
            pbOfficerSignature.Image = selectedImage
        End If
    End Sub

    Private Sub btnInmateSignature_Click(sender As Object, e As EventArgs) Handles btnInmateSignature.Click
        Dim openFileDialog As New OpenFileDialog With {
            .Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp"
        }

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedImage = Image.FromFile(openFileDialog.FileName)
            pbInmateSignature.Image = selectedImage
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainform.SwitchToInmateHomeControl()
    End Sub


End Class
