Public Class InmateHomeControl
    ' Method to load the unordered table data
    Public Sub LoadInmateData()
        Dim dt As DataTable = GetTableData(pdl)
        DataGridView1.DataSource = dt
    End Sub

    Private Sub btnAddInmate_Click(sender As Object, e As EventArgs) Handles btnAddInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToAddInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToAddInmateControl()
            Return
        End If
    End Sub

    Private Sub btnUpdateInmate_Click(sender As Object, e As EventArgs) Handles btnUpdateInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToUpdateInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToUpdateInmateControl()
            Return
        End If
    End Sub

    Private Sub btnRemoveInmate_Click(sender As Object, e As EventArgs) Handles btnRemoveInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToRemoveInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToRemoveInmateControl()
            Return
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFilter.SelectedIndexChanged
        ' Get the selected filter from the combo box
        Dim selectedFilter As String = cmbFilter.SelectedItem.ToString()
        Dim orderBy As String = ""

        ' Determine the ORDER BY clause based on the selected filter
        Select Case selectedFilter
            Case "A-Z"
                orderBy = "ORDER BY first_name ASC"
            Case "Date (oldest)"
                orderBy = "ORDER BY date_of_birth ASC"
            Case "Date (newest)"
                orderBy = "ORDER BY date_of_birth DESC"
            Case "ID"
                orderBy = "ORDER BY pdl_id ASC"
        End Select

        ' Fetch the sorted data using the new function
        Dim dt As DataTable = GetTableOrder(pdl, orderBy)

        ' Bind the data to the DataGridView
        DataGridView1.DataSource = dt
    End Sub

    ' Call the LoadInmateData method when the control loads
    Private Sub InmateHomeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadInmateData()
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        ' Get the search input
        Dim searchText As String = RichTextBox1.Text.Trim()

        ' If the search box is empty, load the unordered data
        If String.IsNullOrWhiteSpace(searchText) Then
            LoadInmateData() ' Reload the full dataset
            Return
        End If

        ' Build the filter condition for the SQL query
        Dim filter As String = $"LOWER(first_name) LIKE '%{searchText.ToLower()}%' OR " &
                               $"LOWER(middle_name) LIKE '%{searchText.ToLower()}%' OR " &
                               $"LOWER(last_name) LIKE '%{searchText.ToLower()}%' OR " &
                               $"LOWER(CONCAT(first_name, ' ', last_name)) LIKE '%{searchText.ToLower()}%' OR " &
                               $"LOWER(CONCAT(last_name, ' ', first_name)) LIKE '%{searchText.ToLower()}%' OR " &
                               $"LOWER(CONCAT(first_name, ' ', middle_name, ' ', last_name)) LIKE '%{searchText.ToLower()}%'"

        ' Fetch the filtered data using the GetTableData method
        Dim dt As DataTable = GetTableData(pdl, filter)

        ' Bind the data to the DataGridView
        DataGridView1.DataSource = dt
    End Sub

    ' Event to handle row selection in DataGridView
    ' Declare a variable to store the selected pdl_id
    Private selectedItem As String

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Ensure that a valid row is clicked (not the header)
        If e.RowIndex >= 0 Then
            ' Get the DataRowView for the clicked row
            Dim selectedRowView As DataRowView = TryCast(DataGridView1.Rows(e.RowIndex).DataBoundItem, DataRowView)

            ' Ensure the selected row can be cast to a DataRowView
            If selectedRowView IsNot Nothing Then
                ' Set the currentPDL to the DataRow
                currentPDL = selectedRowView.Row
                ' Save the pdl_id to selectedItem
                selectedItem = CInt(currentPDL("pdl_id"))
                LoadCriminalCaseInformation(selectedItem)
                LoadMedicalInformation(selectedItem)
                txtID.Text = selectedItem
            Else
                currentPDL = Nothing
                selectedItem = Nothing
            End If
        End If
    End Sub

    Private Sub btnViewPDL_Click(sender As Object, e As EventArgs) Handles btnViewPDL.Click
        If currentPDL Is Nothing Then
            MessageBox.Show("Please select a row first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainform IsNot Nothing Then
            mainform.SwitchToViewPDL()
            Return
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToViewPDL()
            Return
        End If

    End Sub

    Private Sub btnViewMed_Click(sender As Object, e As EventArgs) Handles btnViewMed.Click
        ' Check if a row is selected
        If currentPDL Is Nothing Then
            MessageBox.Show("Please select a row first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainform IsNot Nothing Then
            mainform.SwitchtoViewMedical()
            Return
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchtoViewMedical()
            Return
        End If
    End Sub

    Private Sub btnViewCase_Click(sender As Object, e As EventArgs) Handles btnViewCase.Click
        ' Check if a row is selected
        If currentPDL Is Nothing Then
            MessageBox.Show("Please select a row first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainform IsNot Nothing Then
            mainform.switchtoviewcase()
            Return
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.switchtoviewcase()
            Return
        End If

    End Sub
End Class
