Public Class AddStaffControl

    ' Declare variables for text fields
    Dim firstName As String
    Dim middleName As String
    Dim lastName As String
    Dim suffix As String
    Dim street As String
    Dim municipality As String
    Dim city As String
    Dim regionVar As String
    Dim zip As String
    Dim phone As String

    ' Declare variables for radio buttons
    Dim gender As String ' Can be "Male" or "Female"

    ' Declare variable for date picker
    Dim birthdate As DateTime

    ' Declare variables for combo boxes
    Dim status As String ' For cmbStatus
    Dim authority
    Dim position

    Dim imgID
    Dim imgSelf

    Private Sub AddStaffControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set placeholders for all the fields
        Dim sql = "SELECT * FROM staffdetails"
        LoadToDGVForDisplay(sql, dgvAllStaff)
        SetPlaceholders()
    End Sub

    Private Sub SetPlaceholders()
        ' Textboxes placeholders
        SetPlaceholder(txtFirstName, "First Name")
        SetPlaceholder(txtMiddleName, "Middle Name")
        SetPlaceholder(txtLastName, "Last Name")
        SetPlaceholder(txtSuffix, "Suffix")
        SetPlaceholder(txtStreet, "Street Address")
        SetPlaceholder(txtMunicipality, "Municipality")
        SetPlaceholder(txtCity, "City")
        SetPlaceholder(txtRegion, "Region")
        SetPlaceholder(txtZip, "ZIP Code")
        SetPlaceholder(txtPhone, "Phone Number")

        ' ComboBox placeholders
        SetComboBoxPlaceholder(cmbStatus, "Select Status")
        SetComboBoxPlaceholder(cmbPosition, "Select Position")
        SetComboBoxPlaceholder(cmbAuthority, "Select Authority")
    End Sub

    Private Sub SetPlaceholder(txtBox As TextBox, placeholder As String)
        txtBox.Text = placeholder
        txtBox.ForeColor = Color.Gray ' Set the placeholder color
        AddHandler txtBox.Enter, AddressOf TextBox_Enter
        AddHandler txtBox.Leave, AddressOf TextBox_Leave
        txtBox.Tag = placeholder ' Store the placeholder in Tag for reference
    End Sub

    Private Sub SetComboBoxPlaceholder(cmbBox As ComboBox, placeholder As String)
        cmbBox.Text = placeholder
        cmbBox.ForeColor = Color.Gray ' Set the placeholder color
        AddHandler cmbBox.Enter, AddressOf ComboBox_Enter
        AddHandler cmbBox.Leave, AddressOf ComboBox_Leave
        cmbBox.Tag = placeholder ' Store the placeholder in Tag for reference
    End Sub

    Private Sub TextBox_Enter(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = CType(sender, TextBox)
        If txtBox.Text = txtBox.Tag.ToString() Then
            txtBox.Text = ""
            txtBox.ForeColor = Color.Black ' Change text color when typing
        End If
    End Sub

    Private Sub TextBox_Leave(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = CType(sender, TextBox)
        If String.IsNullOrEmpty(txtBox.Text) Then
            txtBox.Text = txtBox.Tag.ToString()
            txtBox.ForeColor = Color.Gray ' Restore placeholder text color
        End If
    End Sub

    Private Sub ComboBox_Enter(sender As Object, e As EventArgs)
        Dim cmbBox As ComboBox = CType(sender, ComboBox)
        If cmbBox.Text = cmbBox.Tag.ToString() Then
            cmbBox.Text = ""
            cmbBox.ForeColor = Color.Black ' Change text color when selecting
        End If
    End Sub

    Private Sub ComboBox_Leave(sender As Object, e As EventArgs)
        Dim cmbBox As ComboBox = CType(sender, ComboBox)
        If String.IsNullOrEmpty(cmbBox.Text) Then
            cmbBox.Text = cmbBox.Tag.ToString()
            cmbBox.ForeColor = Color.Gray ' Restore placeholder text color
        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Retrieve text field values
        firstName = txtFirstName.Text.Trim()
        middleName = txtMiddleName.Text.Trim()
        lastName = txtLastName.Text.Trim()
        suffix = txtSuffix.Text.Trim()
        street = txtStreet.Text.Trim()
        municipality = txtMunicipality.Text.Trim()
        city = txtCity.Text.Trim()
        regionVar = txtRegion.Text.Trim()
        zip = txtZip.Text.Trim()
        phone = txtPhone.Text.Trim()

        ' Validate that none of the fields are empty or contain only white spaces
        If String.IsNullOrWhiteSpace(firstName) OrElse firstName = "First Name" Then
            MessageBox.Show("Please enter a valid first name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(middleName) OrElse middleName = "Middle Name" Then
            MessageBox.Show("Please enter a valid middle name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(lastName) OrElse lastName = "Last Name" Then
            MessageBox.Show("Please enter a valid last name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(suffix) OrElse suffix = "Suffix" Then
            MessageBox.Show("Please enter a valid suffix.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(street) OrElse street = "Street Address" Then
            MessageBox.Show("Please enter a valid street address.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(municipality) OrElse municipality = "Municipality" Then
            MessageBox.Show("Please enter a valid municipality.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(city) OrElse city = "City" Then
            MessageBox.Show("Please enter a valid city.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(regionVar) OrElse regionVar = "Region" Then
            MessageBox.Show("Please enter a valid region.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(zip) OrElse zip = "ZIP Code" Then
            MessageBox.Show("Please enter a valid zip code.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        ElseIf String.IsNullOrWhiteSpace(phone) OrElse phone = "Phone Number" Then
            MessageBox.Show("Please enter a valid phone number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Retrieve gender based on selected radio button
        If RadioButton1.Checked Then
            gender = "Male"
        ElseIf RadioButton2.Checked Then
            gender = "Female"
        Else
            gender = String.Empty ' No selection made
        End If

        ' Retrieve date picker value
        birthdate = dtBirthdate.Value

        ' Validate combo box selections
        If cmbStatus.SelectedItem Is Nothing OrElse cmbStatus.Text = "Select Status" Then
            MessageBox.Show("Please select a valid status.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If cmbPosition.SelectedItem Is Nothing OrElse cmbPosition.Text = "Select Position" Then
            MessageBox.Show("Please select a valid position.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If cmbAuthority.SelectedItem Is Nothing OrElse cmbAuthority.Text = "Select Authority" Then
            MessageBox.Show("Please select a valid authority.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        position = cmbPosition.Text
        status = cmbStatus.Text
        authority = cmbAuthority.Text

        ' Retrieve file paths for ID and image
        imgID = ImageToByteArray(pbID.Image)
        imgSelf = ImageToByteArray(pbImage.Image)

        ' Create a Dictionary of staff details
        Dim staffDetails As New Dictionary(Of String, Object) From {
            {"FirstName", firstName},
            {"MiddleName", middleName},
            {"LastName", lastName},
            {"Suffix", suffix},
            {"Street", street},
            {"Municipality", municipality},
            {"City", city},
            {"Region", regionVar},
            {"ZIP", zip},
            {"Phone", phone},
            {"Gender", gender},
            {"Birthdate", birthdate},
            {"Status", status},
            {"Position", position},
            {"Authority", authority},
            {"IDFile", imgID},
            {"SelfImage", imgSelf}
        }

        ' If all validations pass, create the record
        CreateRecord("staffdetails", staffDetails)

        MsgBox("staff added success")
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        MainDashboard.SwitchToAddStaffControlHome()

    End Sub


    Private Sub btnAddCellblock_Click(sender As Object, e As EventArgs) Handles btnAddCellblock.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.switchToAddEntity()
        End If

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)

        If adminDashboard IsNot Nothing Then
            adminDashboard.switchToAddEntity()
        End If
    End Sub



    Private Sub btnChooseId_Click(sender As Object, e As EventArgs) Handles btnChooseId.Click
        ' Create an OpenFileDialog to allow the user to choose an image file for the ID
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*" ' Filter for image files

        ' Show the dialog and check if the user selected a file
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Assign the chosen file to the PictureBox for ID
            pbID.Image = Image.FromFile(openFileDialog.FileName)
        End If
    End Sub

    Private Sub btnChooseImg_Click(sender As Object, e As EventArgs) Handles btnChooseImg.Click
        ' Create an OpenFileDialog to allow the user to choose an image file for the profile image
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*" ' Filter for image files

        ' Show the dialog and check if the user selected a file
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Assign the chosen file to the PictureBox for profile image
            pbImage.Image = Image.FromFile(openFileDialog.FileName)
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        MainDashboard.SwitchToAddStaffControlHome()
    End Sub

End Class
