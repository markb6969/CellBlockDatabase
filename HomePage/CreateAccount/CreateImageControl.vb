Imports System.Net.Mime.MediaTypeNames
Imports System.IO
Imports System.Drawing


Public Class CreateImageControl
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        ' Check if pbImage or pbID does not have an image
        If pbImage.Image Is Nothing Then
            MessageBox.Show("Please upload a profile image before submitting.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If pbID.Image Is Nothing Then
            MessageBox.Show("Please upload an ID image before submitting.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Convert images to byte arrays and assign to accData
        accData.Image = ImageToByteArray(pbImage.Image)
        accData.IdImage = ImageToByteArray(pbID.Image)
        Dim Data = accData.GetDictionary()

        CreateRecord("accounts", Data)
        ' Switch to the next control
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToCreateAccountControl()
    End Sub


    Private Sub btnChooseImg_Click(sender As Object, e As EventArgs) Handles btnChooseImg.Click
        ' Open File Dialog to choose an image for pbImage
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*"
        openFileDialog.Title = "Select an Image"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFilePath As String = openFileDialog.FileName

            ' Load the selected image into pbImage (e.g., PictureBox)
            If IO.File.Exists(selectedFilePath) Then
                pbImage.Image = Drawing.Image.FromFile(selectedFilePath)
                pbImage.Tag = selectedFilePath ' Store the file path for further processing if needed
            Else
                MessageBox.Show("The selected file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub btnChooseId_Click(sender As Object, e As EventArgs) Handles btnChooseId.Click
        ' Open File Dialog to choose an image for pbImage (same as the first button)
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*"
        openFileDialog.Title = "Select an ID Image"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFilePath As String = openFileDialog.FileName

            ' Load the selected image into pbImage (e.g., PictureBox)
            If IO.File.Exists(selectedFilePath) Then
                pbID.Image = Drawing.Image.FromFile(selectedFilePath)
                pbID.Tag = selectedFilePath ' Store the file path for further processing if needed
            Else
                MessageBox.Show("The selected file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub


End Class
