Imports System.Data.SqlClient
Imports CellBlockIM.My
Imports MySql.Data.MySqlClient

Public Class UpdateHomeControl
    ' Declare a DataTable to store the inmate data
    Private inmateData As DataTable

    ' Load data into the DataGridView
    Private Sub LoadData()
        Dim query As String = "SELECT * FROM `pdl` WHERE `status` = 'In custody';"

        ' Use your existing LoadToDGV function to load data
        LoadToDGV(query, dgvInmates)

        ' Copy the data from the DataGridView to a DataTable for filtering
        inmateData = New DataTable()

        For Each column As DataGridViewColumn In dgvInmates.Columns
            inmateData.Columns.Add(column.Name, column.ValueType)
        Next

        For Each row As DataGridViewRow In dgvInmates.Rows
            If Not row.IsNewRow Then
                Dim dataRow As DataRow = inmateData.NewRow()
                For Each column As DataGridViewColumn In dgvInmates.Columns
                    dataRow(column.Name) = row.Cells(column.Name).Value
                Next
                inmateData.Rows.Add(dataRow)
            End If
        Next
    End Sub


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
            dgvInmates.DataSource = inmateData
            matchedPDLId = Nothing
        Else
            ' Apply the filter to the DataTable's DefaultView
            Dim filteredData As DataTable = inmateData.Clone()
            Dim rows = inmateData.Select(filter)
            For Each row As DataRow In rows
                filteredData.ImportRow(row)
            Next
            dgvInmates.DataSource = filteredData

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

    ' Handle DataGridView RowHeaderMouseClick to select an item
    Private Sub dgvInmates_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInmates.CellClick
        ' Ensure the click is on a valid row (not the header)
        If e.RowIndex >= 0 Then
            ' Get the selected row
            Dim selectedRow As DataGridViewRow = dgvInmates.Rows(e.RowIndex)

            ' Retrieve the PDL ID or other data you want from the selected row
            selectedItem = Convert.ToInt32(selectedRow.Cells("pdl_id").Value)
            txtID.Text = selectedItem

        End If
    End Sub

    ' Handle control Load event
    Private Sub UpdateHomeControl_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Load data into the DataGridView when the control loads
        LoadData()

        ' Populate ComboBox with options
        cmbSelect.Items.AddRange(New String() {"pdl information", "criminal case", "medical information", "add offense"})
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Check if a valid item is selected
        If selectedItem < 0 Then
            MessageBox.Show("Please select a valid PDL from the list.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Check if a ComboBox option is selected
        If cmbSelect.SelectedItem Is Nothing Then
            MessageBox.Show("Please select an option from the dropdown.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)

        ' Perform actions based on the selected ComboBox option
        Select Case cmbSelect.SelectedItem.ToString()
            Case "pdl information"

                LoadPDLInformation(selectedItem)

                If mainForm IsNot Nothing Then
                    mainForm.SwitchToUpdateAddInmateControl()
                    Return
                End If


                LoadPDLInformation(selectedItem)
                If adminDashboard IsNot Nothing Then
                    adminDashboard.SwitchToUpdateAddInmateControl()
                    Return
                End If



            Case "criminal case"

                LoadCriminalCaseInformation(selectedItem)

                If mainForm IsNot Nothing Then
                    mainForm.SwitchToUpdateCrminalCaseControl()
                    Return
                End If


                LoadCriminalCaseInformation(selectedItem)

                If adminDashboard IsNot Nothing Then
                    adminDashboard.SwitchToUpdateCrminalCaseControl()
                    Return
                End If



            Case "medical information"

                LoadMedicalInformation(selectedItem)

                If mainForm IsNot Nothing Then
                    mainForm.SwitchToUpdateMedicalControl()
                    Return
                End If


                LoadMedicalInformation(selectedItem)

                If adminDashboard IsNot Nothing Then
                    adminDashboard.SwitchToUpdateMedicalControl()
                    Return
                End If


            Case "add offense"

                AddOffenseForm.Show()
            Case Else
                MessageBox.Show("Invalid Selection")
        End Select
    End Sub


    ' Load PDL Information from the database and store it in the currentPDL variable

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

    Private Sub btnRemoveInmate_Click(sender As Object, e As EventArgs) Handles btnRemoveInmate.Click
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainform IsNot Nothing Then
            mainform.SwitchToRemoveInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToRemoveInmateControl()
            Return
        End If
    End Sub

    Private Sub btnAddInmate_Click(sender As Object, e As EventArgs) Handles btnAddInmate.Click
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        If mainForm IsNot Nothing Then
            mainform.SwitchToAddInmateControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToAddInmateControl()
            Return
        End If
    End Sub
End Class
