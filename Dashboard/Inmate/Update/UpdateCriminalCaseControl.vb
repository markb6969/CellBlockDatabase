Public Class UpdateCriminalCaseControl

    Private Sub UpdateCriminalCaseControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load the current case details when the form is loaded
        LoadCurrentCase()
    End Sub

    ' Load current case details into the form
    Public Sub LoadCurrentCase()
        If currentCase IsNot Nothing Then
            txtOffense.Text = currentCase("offence_charge").ToString()
            txtModus.Text = currentCase("modus_operandi").ToString()
            txtPlace.Text = currentCase("place_of_arrest").ToString()

            txtOfficerName.Text = currentCase("arresting_officer").ToString()
            txtLawyerName.Text = currentCase("lawyer").ToString()
            txtSentence.Text = currentCase("sentence").ToString

            ' If the mugshot exists, convert it to an image and display it
            Dim mugshotBytes As Byte() = TryCast(currentCase("mugshot"), Byte())
            If mugshotBytes IsNot Nothing Then
                pbMugshot.Image = ByteArrayToImage(mugshotBytes)
            Else
                pbMugshot.Image = Nothing ' Or set to a default placeholder image
            End If

            cmbCells.Text = currentCase("cellblock").ToString()
            cbYes.Checked = Convert.ToBoolean(currentCase("multipleCase"))
        Else
            MsgBox("No case data is available.", MsgBoxStyle.Exclamation)
        End If
    End Sub

    ' Handle switching to the next control
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Show confirmation message before proceeding
        Dim result As MsgBoxResult = MsgBox("Are you sure you want to save the changes?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Confirm Save")

        If result = MsgBoxResult.Yes Then
            ' First, update the case data before moving to the next control
            UpdateCaseData()

            Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
            If mainForm IsNot Nothing Then
                mainForm.SwitchToInmateHomeControl()
            End If
            selectedItem = 0
            ' Display success message after updating
            MsgBox("Criminal case updated successfully.", MsgBoxStyle.Information)
        Else
            ' If the user chose "No", do nothing (just return)
            Return
        End If
    End Sub

    ' Function to update the DataRow with the form values
    Public Sub UpdateCaseData()
        If currentCase IsNot Nothing Then
            crimeCase.offence_charge = txtOffense.Text
            crimeCase.modus_operandi = txtModus.Text
            crimeCase.place_of_arrest = txtPlace.Text
            crimeCase.arresting_officer = txtOfficerName.Text
            crimeCase.lawyer = txtLawyerName.Text
            crimeCase.cellblock = cmbCells.Text
            crimeCase.multipleCase = cbYes.Checked
            crimeCase.pdl_id = selectedItem
            crimeCase.sentence = txtSentence.Text

            ' Save the mugshot image back to the case data
            If pbMugshot.Image IsNot Nothing Then
                crimeCase.mugshot = ImageToByteArray(pbMugshot.Image)
            End If
        End If
        Dim condition As New Dictionary(Of String, Object) From {
          {"pdl_id", selectedItem}
        }
        UpdateRecord(criminal_case, crimeCase.getDictionary, condition)
        Logs("Crime record updated: CaseID" + crimeCase.pdl_id)
    End Sub

    ' Enhanced function for choosing an image
    Private Sub btnChooseImg_Click(sender As Object, e As EventArgs) Handles btnChooseImg.Click
        Dim dialogResult As DialogResult = MessageBox.Show("Would you like to use the camera to capture the mugshot?",
                                                            "Choose Image Source", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

        If dialogResult = DialogResult.Yes Then
            ' Use camera to capture image
            Dim cameraForm As New CameraCaptureForm()
            If cameraForm.ShowDialog() = DialogResult.OK Then
                pbMugshot.Image = capturedImage ' Use the globally captured image
            End If
        ElseIf dialogResult = DialogResult.No Then
            ' Use file selection dialog
            Dim openFileDialog As New OpenFileDialog With {
                .Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp",
                .Title = "Select Mugshot Image"
            }
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                pbMugshot.Image = Image.FromFile(openFileDialog.FileName)
                currentCase("mugshot") = ImageToByteArray(pbMugshot.Image) ' Update DataRow mugshot
            End If
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub
End Class
