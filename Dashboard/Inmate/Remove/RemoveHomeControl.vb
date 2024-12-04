Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient

Public Class RemoveHomeControl
    ' Declare a DataTable to store the inmate data
    Private inmateData As DataTable
    Private matchedPDLId As Integer?

    ' Load data into the DataGridView
    Private Sub LoadData()
        Dim query As String = "SELECT * FROM `pdl` WHERE `status` = 'In custody';" ' Replace 'pdl' with your table name

        ' Use your existing LoadToDGV function to load data
        LoadToDGV(query, dgvRows)

        ' Copy the data from the DataGridView to a DataTable for filtering
        inmateData = New DataTable()

        For Each column As DataGridViewColumn In dgvRows.Columns
            inmateData.Columns.Add(column.Name, column.ValueType)
        Next

        For Each row As DataGridViewRow In dgvRows.Rows
            If Not row.IsNewRow Then
                Dim dataRow As DataRow = inmateData.NewRow()
                For Each column As DataGridViewColumn In dgvRows.Columns
                    dataRow(column.Name) = row.Cells(column.Name).Value
                Next
                inmateData.Rows.Add(dataRow)
            End If
        Next
    End Sub

    ' Apply filters to the DataTable
    Private Sub ApplyFilters()
        If inmateData Is Nothing Then Return

        Dim filterConditions As New List(Of String)

        ' Primary filter: Filter by PDL ID (ensure type compatibility)
        If Not String.IsNullOrWhiteSpace(txtPDLId.Text) Then
            filterConditions.Add($"Convert(pdl_id, 'System.String') LIKE '%{txtPDLId.Text}%'")
        End If

        ' Secondary filter: Filter by Name (combining first, last, and middle names)
        If Not String.IsNullOrEmpty(txtPDLName.Text) Then
            Dim nameFilter As String = $"(first_name LIKE '%{txtPDLName.Text}%' OR " &
                                   $"last_name LIKE '%{txtPDLName.Text}%' OR " &
                                   $"middle_name LIKE '%{txtPDLName.Text}%')"
            ' If PDL ID is already a filter, add nameFilter as an AND condition
            If filterConditions.Count > 0 Then
                filterConditions.Add($"({String.Join(" OR ", filterConditions)}) AND {nameFilter}")
            Else
                filterConditions.Add(nameFilter)
            End If
        End If

        ' Combine all conditions
        Dim filter As String = String.Join(" OR ", filterConditions)

        ' If no filters are applied, clear the filter and show all rows
        If String.IsNullOrWhiteSpace(filter) Then
            dgvRows.DataSource = inmateData
            matchedPDLId = Nothing
        Else
            ' Apply the filter to the DataTable's DefaultView
            Dim filteredData As DataTable = inmateData.Clone()
            Dim rows = inmateData.Select(filter)
            For Each row As DataRow In rows
                filteredData.ImportRow(row)
            Next
            dgvRows.DataSource = filteredData

            ' Assign the first matching PDL ID to the variable
            If rows.Length > 0 AndAlso Not IsDBNull(rows(0)("pdl_id")) Then
                matchedPDLId = Convert.ToInt32(rows(0)("pdl_id"))
            Else
                matchedPDLId = Nothing
            End If
        End If
    End Sub

    ' Handle PDL ID TextBox TextChanged
    Private Sub txtPDLId_TextChanged(sender As Object, e As EventArgs) Handles txtPDLId.TextChanged
        ApplyFilters()
    End Sub

    ' Handle PDL Name TextBox TextChanged
    Private Sub txtPDLName_TextChanged(sender As Object, e As EventArgs) Handles txtPDLName.TextChanged
        ApplyFilters()
    End Sub

    ' Handle btnNext click to switch control
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Ensure a valid item is selected before proceeding
        If selectedItem = 0 Then
            MessageBox.Show("Please select a valid item before proceeding.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Proceed to switch to the next control if the check passes
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainForm IsNot Nothing Then
            mainForm.SwitchToRemoveOptionControl()
        End If


        ' Proceed to switch to the next control if the check passes
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToRemoveOptionControl()
        End If
    End Sub


    ' Handle btnRemoveInmate click to switch to RemoveInmateControl
    Private Sub btnRemoveInmate_Click(sender As Object, e As EventArgs) Handles btnRemoveInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToRemoveInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToInmateHomeControl()
            Return
        End If
    End Sub

    ' Handle btnUpdateInmate click to switch to UpdateInmateControl
    Private Sub btnUpdateInmate_Click(sender As Object, e As EventArgs) Handles btnUpdateInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToUpdateInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToUpdateInmateControl()
            Return
        End If
    End Sub

    ' Handle btnAddInmate click to switch to AddInmateControl
    Private Sub btnAddInmate_Click(sender As Object, e As EventArgs) Handles btnAddInmate.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToAddInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToUpdateInmateControl()
            Return
        End If

    End Sub

    ' Handle control Load event
    Private Sub RemoveHomeControl_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Load data into the DataGridView when the control loads
        LoadData()
    End Sub

    ' Handle DataGridView RowHeaderMouseClick to select an item (optional)

    Private Sub dgvRows_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRows.CellClick
        ' Ensure the click is on a valid row (not the header)
        If e.RowIndex >= 0 Then
            ' Get the DataRow corresponding to the clicked row index
            Dim dataTable As DataTable = CType(dgvRows.DataSource, DataTable)
            Dim selectedRow As DataRow = dataTable.Rows(e.RowIndex)

            ' Retrieve the PDL ID (assuming 'pdl_id' is the column name for PDL ID)
            selectedItem = Convert.ToInt32(selectedRow("pdl_id"))
            txtID.Text = selectedItem

            ' Store the entire row in the currentPDL
            currentPDL = selectedRow
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToInmateHomeControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToInmateHomeControl()
            Return
        End If
    End Sub
End Class
