Public Class AddStaffHomeControl
    Private Sub AddStaffHomeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Attach event handlers for AddEditDeleteControl events
        AddHandler AddEditDeleteControl1.AddButtonClick, AddressOf HandleAddButtonClick
        AddHandler AddEditDeleteControl1.EditButtonClick, AddressOf HandleEditButtonClick
        AddHandler AddEditDeleteControl1.DeleteButtonClick, AddressOf HandleDeleteButtonClick

        Dim sql = "SELECT * FROM staffdetails"
        LoadToDGVForDisplay(sql, dgvStaffInfo)
    End Sub

    Private Sub HandleAddButtonClick()
        ' Get reference to MainDashboard
        Dim mainDashboard As MainDashboard = TryCast(Me.FindForm(), MainDashboard)
        If mainDashboard IsNot Nothing Then
            mainDashboard.SwitchToAddStaffControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToAddStaffControl()
            Return
        End If
    End Sub

    Private Sub HandleEditButtonClick()
        ' Get reference to MainDashboard
        Dim mainDashboard As MainDashboard = TryCast(Me.FindForm(), MainDashboard)
        If mainDashboard IsNot Nothing Then
            mainDashboard.SwitchToEditStaffControl()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToEditStaffControl()
            Return
        End If
    End Sub

    Private Sub HandleDeleteButtonClick()
        ' Get reference to MainDashboard
        Dim mainDashboard As MainDashboard = TryCast(Me.FindForm(), MainDashboard)
        If mainDashboard IsNot Nothing Then
            mainDashboard.SwitchToDeleteStaffControl()
        End If


        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.FindForm(), AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToDeleteStaffControl()
            Return
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
