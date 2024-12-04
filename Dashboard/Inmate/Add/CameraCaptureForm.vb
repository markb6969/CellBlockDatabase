Imports AForge.Video
Imports AForge.Video.DirectShow

Public Class CameraCaptureForm
    Dim videoSource As VideoCaptureDevice
    Dim capturedBitmap As Bitmap

    ' Set up the camera capture when form loads
    Private Sub CameraCaptureForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize video capture device
        Dim videoDevices As New FilterInfoCollection(FilterCategory.VideoInputDevice)

        If videoDevices.Count = 0 Then
            MsgBox("No camera devices found.", MsgBoxStyle.Critical)
            Return
        End If

        videoSource = New VideoCaptureDevice(videoDevices(0).MonikerString)

        ' Handle new frame
        AddHandler videoSource.NewFrame, AddressOf videoSource_NewFrame

        ' Start the video capture
        videoSource.Start()
    End Sub

    ' Handle new frames from the camera
    Private Sub videoSource_NewFrame(sender As Object, eventArgs As NewFrameEventArgs)
        If capturedBitmap IsNot Nothing Then
            capturedBitmap.Dispose() ' Release the previous frame
        End If

        capturedBitmap = CType(eventArgs.Frame.Clone(), Bitmap)
        pbCamera.Image = capturedBitmap ' Display the captured frame in the PictureBox
    End Sub

    ' Capture the current image from the camera and assign it to the global image variable
    Private Sub btnCapture_Click(sender As Object, e As EventArgs) Handles btnCapture.Click
        If capturedBitmap IsNot Nothing Then
            Try
                ' Assign the captured image to the global variable in the module
                capturedImage = New Bitmap(capturedBitmap) ' Use the global image variable from the module
                MsgBox("Image captured successfully!")
                Logs("Mugshot captured: " + inmateData.last_name + ", " + inmateData.first_name + " " + inmateData.middle_name)
            Catch ex As ArgumentException
                MsgBox("Error capturing image: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        Else
            MsgBox("No valid image captured.", MsgBoxStyle.Information)
        End If
    End Sub

    ' Save the captured image and close the camera form
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If capturedImage IsNot Nothing Then
            ' Return the captured image to the parent form
            Me.DialogResult = DialogResult.OK
            Logs("Mugshot saved: " + inmateData.last_name + ", " + inmateData.first_name + " " + inmateData.middle_name)
            Me.Close() ' Close the camera form

        Else
            MsgBox("No image captured yet.", MsgBoxStyle.Information)
        End If
    End Sub

    ' Stop the camera when closing the form
    Private Sub CameraCaptureForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If videoSource IsNot Nothing AndAlso videoSource.IsRunning Then
            videoSource.SignalToStop()
            videoSource.WaitForStop()
        End If
    End Sub
End Class
