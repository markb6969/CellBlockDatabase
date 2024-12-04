Imports AForge.Video
Imports AForge.Video.DirectShow

Public Class AddCriminalCaseControl
    Dim videoSource As VideoCaptureDevice
    Dim capturedBitmap As Bitmap
    Public Event FormClosing(sender As Object, e As FormClosingEventArgs)

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)


        ' Check for empty required fields
        If String.IsNullOrWhiteSpace(txtOffense.Text) Then
            MsgBox("Offense is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtModus.Text) Then
            MsgBox("Modus Operandi is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtPlace.Text) Then
            MsgBox("Place of Arrest is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(dtDateofArrest.Text) Then
            MsgBox("Date of Arrest is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtOfficerName.Text) Then
            MsgBox("Arresting Officer is required.", MsgBoxStyle.Critical)
            Return
        End If

        If String.IsNullOrWhiteSpace(txtLawyerName.Text) Then
            MsgBox("Lawyer's Name is required.", MsgBoxStyle.Critical)
            Return
        End If
        If String.IsNullOrWhiteSpace(txtSentence.Text) Then
            MsgBox("sentence is required.", MsgBoxStyle.Critical)
            Return
        End If
        ' Setting values to the crimeCase object
        crimeCase.offence_charge = txtOffense.Text
        crimeCase.modus_operandi = txtModus.Text
        crimeCase.place_of_arrest = txtPlace.Text
        crimeCase.date_of_arrest = dtDateofArrest.Text
        crimeCase.arresting_officer = txtOfficerName.Text
        crimeCase.lawyer = txtLawyerName.Text
        crimeCase.pdl_id = selectedItem
        crimeCase.sentence = txtSentence.Text

        ' Conditional checkbox for multiple cases
        If cbYes.Checked Then
            crimeCase.multipleCase = True
        Else
            crimeCase.multipleCase = False
        End If

        ' Get the table data
        Dim data As DataTable = GetTableData(pdl)

        If data.Rows.Count > 0 Then
            ' Get the last record by checking the last row in the DataTable
            crimeCase.pdl_id = Convert.ToInt32(data.Rows(data.Rows.Count - 1)("pdl_id"))
            crime.pdl_id = Convert.ToInt32(data.Rows(data.Rows.Count - 1)("pdl_id"))
        Else
            MsgBox("No records found in the table.", MsgBoxStyle.Information)
            Return
        End If

        crime.crime_commited = txtOffense.Text
        CreateRecord(crimes, crime.getDictionary)

        ' Handling the mugshot
        If pbMugshot.Image Is Nothing Then
            MsgBox("Mugshot is required.", MsgBoxStyle.Critical)
            Return
        End If

        crimeCase.mugshot = ImageToByteArray(pbMugshot.Image)
        crimeCase.cellblock = cmbCells.Text
        inmateData.cell = cmbCells.Text

        Dim conditions As New Dictionary(Of String, Object) From {
        {"pdl_id", crimeCase.pdl_id}
    }
        UpdateRecord(pdl, inmateData.GetDictionary, conditions)

        CreateRecord(criminal_case, crimeCase.getDictionary)

        Logs("Added PDL record: " + inmateData.last_name + ", " + inmateData.first_name + " " + inmateData.middle_name)
        Logs("Added Criminal case: " + inmateData.last_name + ", " + inmateData.first_name + " " + inmateData.middle_name)
        Logs("Added Crime: " + inmateData.last_name + ", " + inmateData.first_name + " " + inmateData.middle_name)

        ' Switch to the next form
        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateMedicalInfoControl()
        End If
    End Sub


    ' Handle the camera capture and display the image
    Private Sub btnChooseImg_Click(sender As Object, e As EventArgs) Handles btnChooseImg.Click
        ' Open CameraCaptureForm as a modal dialog
        Dim cameraForm As New CameraCaptureForm()

        ' Show the form and wait for user interaction
        If cameraForm.ShowDialog() = DialogResult.OK Then
            ' Set the captured image from the global variable to pbMugshot
            pbMugshot.Image = capturedImage
        End If
    End Sub

    ' Capture the frame from the camera and display it
    Private Sub videoSource_NewFrame(sender As Object, eventArgs As NewFrameEventArgs)
        ' Capture the frame from the camera
        If capturedBitmap IsNot Nothing Then
            capturedBitmap.Dispose() ' Release the previous frame
        End If

        ' Create a new Bitmap from the captured frame
        capturedBitmap = CType(eventArgs.Frame.Clone(), Bitmap)

        ' Set the captured frame to the PictureBox (pbMugshot)
        pbMugshot.Image = capturedBitmap
    End Sub

    ' Stop video capture when form is closed
    Private Sub AddCriminalCaseControl_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Stop the video capture when the form is closed
        If videoSource IsNot Nothing AndAlso videoSource.IsRunning Then
            videoSource.SignalToStop()
            videoSource.WaitForStop()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If
    End Sub
End Class
