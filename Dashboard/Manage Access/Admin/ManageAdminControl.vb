Public Class ManageAdminControl

    Private Sub ManageAdminControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim query As String = "SELECT * FROM accounts WHERE user_level = 'admin'"
        LoadToDGVForDisplay(query, DataGridView1)
    End Sub
    Private Sub btnSuperAdmin_Click(sender As Object, e As EventArgs) Handles btnSuperAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToSuperAdminControl()
    End Sub

    Private Sub btnAdmin_Click(sender As Object, e As EventArgs) Handles btnAdmin.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToAdminHomeControl()
    End Sub

    Private Sub btnClient_Click(sender As Object, e As EventArgs) Handles btnClient.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToClientHomeControl()
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToManageAdminInfoControl()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToAdminHomeControl()
    End Sub


    Private Sub txtAdminName_TextChanged(sender As Object, e As EventArgs) Handles txtAdminName.TextChanged
        Dim filterText As String = txtAdminName.Text.Trim()
        Dim query As String

        If String.IsNullOrWhiteSpace(filterText) Then
            ' If the filter text is empty, load all data
            query = "SELECT * FROM accounts WHERE user_level = 'admin'"
        Else
            ' Filter by first_name, last_name, or middle_name
            query = $"SELECT * FROM accounts WHERE user_level = 'admin' AND " &
                    $"(first_name LIKE '%{filterText}%' OR last_name LIKE '%{filterText}%' OR middle_name LIKE '%{filterText}%')"
        End If

        LoadToDGVForDisplay(query, DataGridView1)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            ' Get the selected row
            Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Retrieve the DataRow from the DataGridView's DataSource (assuming it is bound to a DataTable)
            currentAdmin = CType(selectedRow.DataBoundItem, DataRowView).Row

            ' Retrieve the "id" value from the DataRow
            Dim id As Object = currentAdmin("id")

            ' Display or use the ID as needed
            If id IsNot Nothing Then
                MessageBox.Show($"Admin selected: {id}", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("ID is null or empty.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub
End Class
