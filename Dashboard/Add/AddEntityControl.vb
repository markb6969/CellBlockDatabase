Public Class AddEntityControl

    ' Load the initial data when the form is loaded
    Private Sub AddEntityControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sql = "SELECT * FROM pdl"
        LoadToDGV(sql, dgvCellblockInfo)
    End Sub

    ' Handle ComboBox selection change to apply filter
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim selectedOption As String = ComboBox1.SelectedItem.ToString()
        Dim sql As String = "SELECT * FROM pdl WHERE cell = @cell"

        ' Apply different filter based on selected ComboBox item
        Select Case selectedOption
            Case "minimum security"
                sql = "SELECT * FROM pdl WHERE cell = 'minimum security'"
            Case "medium security"
                sql = "SELECT * FROM pdl WHERE cell = 'medium security'"
            Case "maximum security"
                sql = "SELECT * FROM pdl WHERE cell = 'maximum security'"
            Case "juvenile correction"
                sql = "SELECT * FROM pdl WHERE cell = 'juvenile correction'"
            Case "women's facility"
                sql = "SELECT * FROM pdl WHERE cell = 'women''s facility'" ' Note the double single quote for apostrophe
            Case "rehabilitational facility"
                sql = "SELECT * FROM pdl WHERE cell = 'rehabilitational facility'"
            Case Else
                sql = "SELECT * FROM pdl" ' Default to show all data
        End Select

        ' Load the filtered data into the DataGridView
        LoadToDGV(sql, dgvCellblockInfo)
    End Sub

    ' Button click to switch to Add Staff Control Home
    Private Sub btnAddStaff_Click(sender As Object, e As EventArgs) Handles btnAddStaff.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.SwitchToAddStaffControlHome()
        End If

        'ADMIN DASHBOARD
        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToAddStaffControlHome()
            Return
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainForm IsNot Nothing Then
            mainForm.switchToAddEntity()
        End If


        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)

        If adminDashboard IsNot Nothing Then
            adminDashboard.switchToAddEntity()
        End If
    End Sub

    ' Button click to switch to Move PDL Control Home

End Class
