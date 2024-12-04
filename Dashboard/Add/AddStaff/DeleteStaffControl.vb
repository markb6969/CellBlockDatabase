Public Class DeleteStaffControl


    Private Sub DeleteStaffControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadToDGVForDisplay("SELECT * FROM staffdetails", dgvAllStaff)
    End Sub
    Private Sub btnAddStaff_Click(sender As Object, e As EventArgs) Handles btnAddStaff.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToAddStaffControlHome()
        End If
    End Sub


    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        MainDashboard.SwitchToAddStaffControlHome()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If String.IsNullOrWhiteSpace(txtStaffID.Text.ToString) Then
            MessageBox.Show("Please select a staff first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim confirmation As DialogResult = MessageBox.Show("Are you sure you want to delete this staff?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirmation = DialogResult.Yes Then
            Dim condition As New Dictionary(Of String, Object) From {
            {"StaffID", currentStaff("StaffID")}
        }
            DeleteRecord("staffdetails", condition)
            MsgBox("Staff information deleted")

            Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
            MainDashboard.SwitchToAddStaffControlHome()
        End If
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
            txtStaffID.Text = currentStaff("StaffID").ToString
            MsgBox("Staff selected")
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
End Class
