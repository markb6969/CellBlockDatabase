Public Class EditStaffControl

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

    Private Sub EditStaffControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadToDGVForDisplay("SELECT * FROM staffdetails", dgvAllStaff)
    End Sub

    Private Sub btnAddStaff_Click(sender As Object, e As EventArgs) Handles btnAddStaff.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToAddStaffControlHome()
        End If
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

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        MainDashboard.SwitchToAddStaffControlHome()
    End Sub

    Private Sub dgvAllStaff_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAllStaff.CellClick
        ' Check if the clicked row is valid
        If e.RowIndex >= 0 Then
            ' Get the selected row
            currentStaff = CType(dgvAllStaff.Rows(e.RowIndex).DataBoundItem, DataRowView).Row

            ' Load data into text fields
            txtFirstName.Text = currentStaff("FirstName").ToString()
            txtMiddleName.Text = currentStaff("MiddleName").ToString()
            txtLastName.Text = currentStaff("LastName").ToString()
            txtSuffix.Text = currentStaff("Suffix").ToString()
            txtStreet.Text = currentStaff("Street").ToString()
            txtMunicipality.Text = currentStaff("Municipality").ToString()
            txtCity.Text = currentStaff("City").ToString()
            txtRegion.Text = currentStaff("Region").ToString()
            txtZip.Text = currentStaff("ZIP").ToString()
            txtPhone.Text = currentStaff("Phone").ToString()
            txtStaffID.Text = currentStaff("StaffID")

            ' Load data into radio buttons (Gender)
            If currentStaff("Gender").ToString() = "Male" Then
                RadioButton1.Checked = True
            ElseIf currentStaff("Gender").ToString() = "Female" Then
                RadioButton2.Checked = True
            End If

            ' Load data into date picker
            dtBirthdate.Value = Convert.ToDateTime(currentStaff("Birthdate"))

            ' Load data into combo boxes
            cmbStatus.SelectedItem = currentStaff("Status").ToString()
            cmbPosition.SelectedItem = currentStaff("Position").ToString()
            cmbAuthority.SelectedItem = currentStaff("Authority").ToString()

            ' Load data into picture boxes (images)
            If Not IsDBNull(currentStaff("IDFile")) Then
                imgID = CType(currentStaff("IDFile"), Byte())
                pbID.Image = ByteArrayToImage(imgID)
            End If

            If Not IsDBNull(currentStaff("SelfImage")) Then
                imgSelf = CType(currentStaff("SelfImage"), Byte())
                pbImage.Image = ByteArrayToImage(imgSelf)
            End If

            MsgBox("Staff selected")
        End If
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrWhiteSpace(txtStaffID.Text.ToString) OrElse txtStaffID.Text.Equals("Enter Staff ID") Then
            ' Handle case where StaffID is empty or null
            MessageBox.Show("please select a staff first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        ' Create a dictionary to store the staff data
        Dim staffData As New Dictionary(Of String, Object)

        ' Add data from text fields
        staffData.Add("FirstName", txtFirstName.Text)
        staffData.Add("MiddleName", txtMiddleName.Text)
        staffData.Add("LastName", txtLastName.Text)
        staffData.Add("Suffix", txtSuffix.Text)
        staffData.Add("Street", txtStreet.Text)
        staffData.Add("Municipality", txtMunicipality.Text)
        staffData.Add("City", txtCity.Text)
        staffData.Add("Region", txtRegion.Text)
        staffData.Add("ZIP", txtZip.Text)
        staffData.Add("Phone", txtPhone.Text)

        ' Add data from radio buttons (Gender)
        If RadioButton1.Checked Then
            staffData.Add("Gender", "Male")
        ElseIf RadioButton2.Checked Then
            staffData.Add("Gender", "Female")
        End If

        ' Add data from date picker (Birthdate)
        staffData.Add("Birthdate", dtBirthdate.Value)

        ' Add data from combo boxes
        staffData.Add("Status", cmbStatus.SelectedItem)
        staffData.Add("Position", cmbPosition.SelectedItem)
        staffData.Add("Authority", cmbAuthority.SelectedItem)

        ' Add data from picture boxes (ID and Self Image)
        If pbID.Image IsNot Nothing Then
            staffData.Add("IDFile", ImageToByteArray(pbID.Image)) ' Convert image to byte array using your ImgToByteArray function
        End If

        If pbImage.Image IsNot Nothing Then
            staffData.Add("SelfImage", ImageToByteArray(pbImage.Image)) ' Convert image to byte array using your ImgToByteArray function
        End If


        Dim result As DialogResult = MessageBox.Show("Are you sure you want to save the changes?", "Confirm Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        ' If the user clicks Yes, proceed with the update
        If result = DialogResult.Yes Then
            Dim condition As New Dictionary(Of String, Object) From {
            {"StaffID", currentStaff("StaffID")}
            }
            ' Proceed with the update
            UpdateRecord("staffdetails", staffData, condition)
            MsgBox("Staff update success")
            txtStaffID.Text = ""
            Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

            If mainForm IsNot Nothing Then
                mainForm.SwitchToAddStaffControlHome()
            End If
        End If
    End Sub


    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        ' Get the text from the search box
        Dim searchText As String = txtSearch.Text.Trim()

        If String.IsNullOrWhiteSpace(searchText) Or searchText.Equals("Search Staff") Then
            ' If search text is empty, reload the original data
            LoadToDGVForDisplay("SELECT * FROM staffdetails", dgvAllStaff)
        Else

            ' Get the original staff data (DataTable)
            Dim staffData As DataTable = GetTableData("staffdetails")

            ' Create a DataView to filter the data
            Dim staffView As New DataView(staffData)

            ' Apply the filter to search FirstName, LastName, and MiddleName in all combinations
            Dim filter As String = $"FirstName LIKE '%{searchText}%' OR " &
                                    $"LastName LIKE '%{searchText}%' OR " &
                                    $"MiddleName LIKE '%{searchText}%'"

            ' Apply the filter to the DataView
            staffView.RowFilter = filter

            ' Set the filtered data to the DataGridView (dgvAllStaff)
            dgvAllStaff.DataSource = staffView

        End If
    End Sub
    Private Sub txtSearch_Enter(sender As Object, e As EventArgs) Handles txtSearch.Enter
        ' If the TextBox is not empty, clear the placeholder
        If txtSearch.Text = "Search Staff" Then
            txtSearch.Text = ""
            txtSearch.ForeColor = Color.Black ' Change text color to black when user types
        End If
    End Sub

    Private Sub txtSearch_Leave(sender As Object, e As EventArgs) Handles txtSearch.Leave
        ' If the TextBox is empty, show the placeholder text again
        If String.IsNullOrWhiteSpace(txtSearch.Text) Then
            txtSearch.Text = "Search Staff"
            txtSearch.ForeColor = Color.Gray ' Change text color to gray for placeholder
        End If
    End Sub

    Private Sub txtZip_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtZip.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPhone.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

End Class
