Public Class VisitorControl
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' Check if a row is selected
        If currentPending Is Nothing Then
            ' Show a message if no row is selected
            MessageBox.Show("Please select a visitor from the list before proceeding.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return ' Exit the method if no row is selected
        End If

        ' Proceed with switching the control if a row is selected
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToVisitationApproveControl()
    End Sub

    ' Load all visitors with "pending" status when the form loads
    Private Sub VisitorControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' SQL Queries to fetch visitor data with "pending" status
        Dim sqlAllVisitors As String = "SELECT * FROM visitors WHERE status = 'pending'"
        LoadToDGV(sqlAllVisitors, dgvPending)
    End Sub

    ' Event triggered when the text in txtVisitorName changes (for filtering)
    Private Sub txtVisitorName_TextChanged(sender As Object, e As EventArgs) Handles txtVisitorName.TextChanged
        Dim filterText As String = txtVisitorName.Text.Trim()
        Dim query As String

        If String.IsNullOrWhiteSpace(filterText) Then
            ' If the filter text is empty, load all visitors with 'pending' status
            query = "SELECT * FROM visitors WHERE status = 'pending'"
        Else
            ' Filter by first_name, middle_name, or last_name
            query = $"SELECT * FROM visitors WHERE status = 'pending' AND " &
                    $"(visitor_first_name LIKE '%{filterText}%' OR visitor_middle_name LIKE '%{filterText}%' OR visitor_last_name LIKE '%{filterText}%')"
        End If

        LoadToDGVForDisplay(query, dgvPending)
    End Sub

    ' Event triggered when a cell in dgvPending is clicked
    Private Sub dgvPending_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPending.CellClick
        ' Check if the row index is valid (skip header row)
        If e.RowIndex >= 0 Then
            ' Get the selected row (from the clicked cell's row index)
            Dim selectedRow As DataGridViewRow = dgvPending.Rows(e.RowIndex)

            ' Retrieve the DataRow from the DataGridViewRow (assuming the DataGridView is bound to a DataTable)
            currentPending = CType(selectedRow.DataBoundItem, DataRowView).Row

            ' Now you can access the data from the selected DataRow
            Dim visitorFirstName As String = currentPending("visitor_first_name").ToString()
            Dim visitorMiddleName As String = currentPending("visitor_middle_name").ToString()
            Dim visitorLastName As String = currentPending("visitor_last_name").ToString()

            ' Example: Show the visitor name or perform other actions
            MessageBox.Show($"Selected Visitor: {visitorFirstName} {visitorMiddleName} {visitorLastName}")
        End If
    End Sub
End Class
