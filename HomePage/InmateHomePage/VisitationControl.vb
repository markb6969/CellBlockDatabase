Imports System.Data.SqlClient
Imports System.IO

Public Class VisitationControl

    Private pdl_id As String
    Private first_name As String
    Private last_name As String
    Private middle_name As String

    Private Sub VisitationControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set initial placeholder text
        SetPlaceholders()

        ' Center align text in specified RichTextBoxes
        cmbIdentification.Text = "type of ID"
        CenterAlignRichTextBoxes(RichTextBox2, RichTextBox4, RichTextBox5, RichTextBox6, RichTextBox10, RichTextBox14, RichTextBox15, RichTextBox16)

        ' Load the DataGridView with initial data
        Dim sql As String = "SELECT * FROM pdl"
        LoadToDGVForDisplay(sql, dgvPDL)
    End Sub

    ' Function to set the placeholders for all text fields
    Private Sub SetPlaceholders()
        txtVisitorFirstName.Text = "[Enter First Name]"
        txtVisitorLastName.Text = "[Enter Last Name]"
        txtVisitorMiddleName.Text = "[Enter Middle Name]"
        txtContactNumber.Text = "[Enter Contact Number]"
        txtEmail.Text = "[Enter Email Address]"
        txtStreet.Text = "[Enter Street Address]"
        txtMunicipality.Text = "[Enter Municipality]"
        txtCity.Text = "[Enter City]"
        txtRegion.Text = "[Enter Region]"
        txtZip.Text = "[Enter Zip Code]"
        txtCountry.Text = "[Enter Country]"
        txtPDLName.Text = "[Enter PDL Name]"
        txtInmateSearch.Text = "[Enter Inmate Search Information]"

        ' Set placeholder text color to gray
        txtVisitorFirstName.ForeColor = Color.Gray
        txtVisitorLastName.ForeColor = Color.Gray
        txtVisitorMiddleName.ForeColor = Color.Gray
        txtContactNumber.ForeColor = Color.Gray
        txtEmail.ForeColor = Color.Gray
        txtStreet.ForeColor = Color.Gray
        txtMunicipality.ForeColor = Color.Gray
        txtCity.ForeColor = Color.Gray
        txtRegion.ForeColor = Color.Gray
        txtZip.ForeColor = Color.Gray
        txtCountry.ForeColor = Color.Gray
        txtPDLName.ForeColor = Color.Gray
        txtInmateSearch.ForeColor = Color.Gray
    End Sub

    ' Event to remove the placeholder text when focus is on the field
    Private Sub RemovePlaceholder(sender As Object, e As EventArgs) Handles txtVisitorFirstName.Enter, txtVisitorLastName.Enter, txtVisitorMiddleName.Enter, txtContactNumber.Enter, txtEmail.Enter, txtStreet.Enter, txtMunicipality.Enter, txtCity.Enter, txtRegion.Enter, txtZip.Enter, txtCountry.Enter, txtPDLName.Enter, txtInmateSearch.Enter
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If textBox.Text = "[Enter First Name]" OrElse textBox.Text = "[Enter Last Name]" OrElse textBox.Text = "[Enter Middle Name]" OrElse textBox.Text = "[Enter Contact Number]" OrElse textBox.Text = "[Enter Email Address]" OrElse textBox.Text = "[Enter Street Address]" OrElse textBox.Text = "[Enter Municipality]" OrElse textBox.Text = "[Enter City]" OrElse textBox.Text = "[Enter Region]" OrElse textBox.Text = "[Enter Zip Code]" OrElse textBox.Text = "[Enter Country]" OrElse textBox.Text = "[Enter PDL Name]" OrElse textBox.Text = "[Enter Inmate Search Information]" Then
            textBox.Text = ""
            textBox.ForeColor = Color.Black ' Change text color when focused
        End If
    End Sub

    ' Event to restore the placeholder text if the field is left blank
    Private Sub RestorePlaceholder(sender As Object, e As EventArgs) Handles txtVisitorFirstName.Leave, txtVisitorLastName.Leave, txtVisitorMiddleName.Leave, txtContactNumber.Leave, txtEmail.Leave, txtStreet.Leave, txtMunicipality.Leave, txtCity.Leave, txtRegion.Leave, txtZip.Leave, txtCountry.Leave, txtPDLName.Leave, txtInmateSearch.Leave
        Dim textBox As TextBox = DirectCast(sender, TextBox)
        If String.IsNullOrWhiteSpace(textBox.Text) Then
            ' Restore the placeholder text
            Select Case textBox.Name
                Case "txtVisitorFirstName"
                    textBox.Text = "[Enter First Name]"
                Case "txtVisitorLastName"
                    textBox.Text = "[Enter Last Name]"
                Case "txtVisitorMiddleName"
                    textBox.Text = "[Enter Middle Name]"
                Case "txtContactNumber"
                    textBox.Text = "[Enter Contact Number]"
                Case "txtEmail"
                    textBox.Text = "[Enter Email Address]"
                Case "txtStreet"
                    textBox.Text = "[Enter Street Address]"
                Case "txtMunicipality"
                    textBox.Text = "[Enter Municipality]"
                Case "txtCity"
                    textBox.Text = "[Enter City]"
                Case "txtRegion"
                    textBox.Text = "[Enter Region]"
                Case "txtZip"
                    textBox.Text = "[Enter Zip Code]"
                Case "txtCountry"
                    textBox.Text = "[Enter Country]"
                Case "txtPDLName"
                    textBox.Text = "[Enter PDL Name]"
                Case "txtInmateSearch"
                    textBox.Text = "[Enter Inmate Search Information]"
            End Select
            textBox.ForeColor = Color.Gray ' Change text color when placeholder is restored
        End If
    End Sub
    ' Function to validate if all required fields are filled
    Private Function ValidateFields() As Boolean
        ' Check if any required fields are empty, not selected, contain placeholder text,
        ' or if email and contact number match their placeholders
        If String.IsNullOrWhiteSpace(txtPDLName.Text) OrElse
       txtPDLName.Text = "[Enter Name]" OrElse
       String.IsNullOrWhiteSpace(cmbRelationship.Text) OrElse
       cmbRelationship.Text = "[Select Relationship]" OrElse
       String.IsNullOrWhiteSpace(txtVisitorLastName.Text) OrElse
       txtVisitorLastName.Text = "[Enter Last Name]" OrElse
       String.IsNullOrWhiteSpace(txtVisitorMiddleName.Text) OrElse
       txtVisitorMiddleName.Text = "[Enter Middle Name]" OrElse
       String.IsNullOrWhiteSpace(txtVisitorFirstName.Text) OrElse
       txtVisitorFirstName.Text = "[Enter First Name]" OrElse
       String.IsNullOrWhiteSpace(txtContactNumber.Text) OrElse
       txtContactNumber.Text = "[Enter Contact Number]" OrElse ' Check for placeholder
       String.IsNullOrWhiteSpace(txtEmail.Text) OrElse
       txtEmail.Text = "[Enter Email]" OrElse ' Check for placeholder
       String.IsNullOrWhiteSpace(txtStreet.Text) OrElse
       txtStreet.Text = "[Enter Street]" OrElse
       String.IsNullOrWhiteSpace(txtMunicipality.Text) OrElse
       txtMunicipality.Text = "[Enter Municipality]" OrElse
       String.IsNullOrWhiteSpace(txtCity.Text) OrElse
       txtCity.Text = "[Enter City]" OrElse
       String.IsNullOrWhiteSpace(txtRegion.Text) OrElse
       txtRegion.Text = "[Enter Region]" OrElse
       String.IsNullOrWhiteSpace(txtZip.Text) OrElse
       txtZip.Text = "[Enter Zip]" OrElse
       String.IsNullOrWhiteSpace(txtCountry.Text) OrElse
       txtCountry.Text = "[Enter Country]" OrElse
       cmbPurpose.SelectedItem Is Nothing OrElse
       cmbIdentification.SelectedItem Is Nothing OrElse
       dtpDate.Value = Nothing OrElse
       (Not rbYes.Checked AndAlso Not rbNo.Checked) OrElse
       (Not rbFemale.Checked AndAlso Not rbMale.Checked) Then
            Return False
        End If
        Return True
    End Function



    ' Submit button handler
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        ' Validate if all required fields are filled
        If ValidateFields() Then
            ' Prepare the data to be inserted into the "visitors" table
            Dim data As New Dictionary(Of String, Object) From {
            {"pdl_id", pdl_id},
            {"visit_date", dtpDate.Value},
            {"purpose", cmbPurpose.SelectedItem.ToString()},
            {"pdl_name", txtPDLName.Text},
            {"relationship", cmbRelationship.Text},
            {"is_victim", If(rbYes.Checked, "Yes", "No")}, ' Victim status
            {"gender", If(rbFemale.Checked, "Female", "Male")}, ' Gender selection
            {"visitor_last_name", txtVisitorLastName.Text},
            {"visitor_middle_name", txtVisitorMiddleName.Text},
            {"visitor_first_name", txtVisitorFirstName.Text},
            {"valid_id", ImageToByteArray(pbValidID.Image)},
            {"id_type", cmbIdentification.SelectedItem.ToString()},
            {"contact_number", txtContactNumber.Text},
            {"email", txtEmail.Text},
            {"street", txtStreet.Text},
            {"municipality", txtMunicipality.Text},
            {"city", txtCity.Text},
            {"region", txtRegion.Text},
            {"zip", txtZip.Text},
            {"country", txtCountry.Text}
        }

            ' Call the CreateRecord method to insert the data into the "visitors" table
            CreateRecord("visitors", data)

            ' Show success message
            MessageBox.Show("Visitor data saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ResetFields()
            SetPlaceholders()
        Else
            ' If any fields are empty, show a message
            MessageBox.Show("Please fill in all required fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub


    Private Sub ResetFields()
        ' Reset textboxes
        txtPDLName.Clear()
        txtVisitorLastName.Clear()
        txtVisitorMiddleName.Clear()
        txtVisitorFirstName.Clear()
        txtContactNumber.Clear()
        txtEmail.Clear()
        txtStreet.Clear()
        txtMunicipality.Clear()
        txtCity.Clear()
        txtRegion.Clear()
        txtZip.Clear()
        txtCountry.Clear()

        ' Reset combo boxes
        cmbPurpose.SelectedIndex = -1
        cmbRelationship.SelectedIndex = -1
        cmbIdentification.SelectedIndex = -1

        ' Reset radio buttons
        rbYes.Checked = False
        rbNo.Checked = False
        rbFemale.Checked = False
        rbMale.Checked = False

        ' Reset date picker
        dtpDate.Value = DateTime.Now

        ' Reset the image
        pbValidID.Image = Nothing
    End Sub

    ' KeyPress event handler for numeric fields
    Private Sub txtContactNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtContactNumber.KeyPress, txtZip.KeyPress
        ' Allow only numbers to be entered in these fields
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub CenterAlignRichTextBoxes(ParamArray textBoxes As RichTextBox())
        For Each rtb As RichTextBox In textBoxes
            rtb.SelectionAlignment = HorizontalAlignment.Center
        Next
    End Sub

    ' Function to handle the search for inmates in the database
    Private Sub txtInmateSearch_TextChanged(sender As Object, e As EventArgs) Handles txtInmateSearch.TextChanged
        ' This function generates a search query for first name, middle name, or last name
        Dim searchText As String = txtInmateSearch.Text.Trim()
        Dim searchQuery As String = "
        SELECT * FROM pdl 
        WHERE 
            first_name LIKE '%" & searchText & "%' OR 
            last_name LIKE '%" & searchText & "%' OR 
            middle_name LIKE '%" & searchText & "%' OR 
            CONCAT(first_name, ' ', last_name) LIKE '%" & searchText & "%' OR
            CONCAT(last_name, ' ', first_name) LIKE '%" & searchText & "%' OR
            CONCAT(first_name, ' ', middle_name, ' ', last_name) LIKE '%" & searchText & "%' OR
            CONCAT(last_name, ' ', middle_name, ' ', first_name) LIKE '%" & searchText & "%'"
        LoadToDGVForDisplay(searchQuery, dgvPDL) ' Update the DataGridView with search results
    End Sub


    ' Function to handle the action when a row in the DataGridView is selected
    Private Sub dgvPDL_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPDL.CellClick
        ' Assuming the first column of the DataGridView contains the PDL's ID (or adjust accordingly)
        If e.RowIndex >= 0 Then
            Dim selectedRow As DataGridViewRow = dgvPDL.Rows(e.RowIndex)
            pdl_id = selectedRow.Cells(0).Value.ToString()

            ' Get first name, middle name, and last name from columns 2, 3, and 4
            Dim firstName As String = selectedRow.Cells(1).Value.ToString()
            Dim middleName As String = selectedRow.Cells(2).Value.ToString()
            Dim lastName As String = selectedRow.Cells(3).Value.ToString()

            ' Format as "Last Name, First Name Middle Name"
            txtPDLName.Text = lastName & ", " & firstName & " " & middleName
            MsgBox("pdl selected, please continue to form")
        End If
    End Sub
    Private Sub pbValidID_Click(sender As Object, e As EventArgs) Handles pbValidID.Click
        ' Create and configure the OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Select a Valid ID"
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif" ' Restrict to image file types

        ' Show the dialog and check if the user selected a file
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Set the selected image to the PictureBox
            pbValidID.Image = Image.FromFile(openFileDialog.FileName)
        End If
    End Sub


End Class
