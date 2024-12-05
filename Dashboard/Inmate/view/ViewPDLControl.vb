Public Class ViewPDLControl
    ' Method to load the details into the User Control fields
    Public Sub LoadDetails()
        Logs("Viewing PDL information")
        ' Use currentPDL directly from the module (assuming currentPDL is a DataRow)
        textName.Text = currentPDL("last_name").ToString() + ", " + currentPDL("first_name").ToString() + " " + currentPDL("middle_name").ToString()
        textHeight.Text = currentPDL("height").ToString()
        textWeight.Text = currentPDL("weight").ToString()
        textPhone.Text = currentPDL("phone_num").ToString()
        textSex.Text = currentPDL("sex").ToString()
        textID.Text = currentPDL("pdl_id").ToString()
        textCivil.Text = currentPDL("civil_status").ToString()
        textAddress.Text = "Address: " &
                        currentPDL("Municipality").ToString() & ", " &
                        currentPDL("Street").ToString() & ", " &
                        currentPDL("Region").ToString() & ", " &
                        currentPDL("District").ToString() & ", " &
                        currentPDL("Zip_code").ToString() & ", " &
                        currentPDL("Country").ToString()
        textBirth.Text = currentPDL("date_of_birth")
        textHair.Text = currentPDL("hair_color").ToString()
        textEyes.Text = currentPDL("eye_color").ToString()
        textStatus.Text = currentPDL("status").ToString()
        If currentCase IsNot Nothing AndAlso currentCase.Table.Columns.Contains("mugshot") AndAlso
   currentCase("mugshot") IsNot DBNull.Value AndAlso currentCase("mugshot") IsNot Nothing Then
            pbMugshot.Image = ByteArrayToImage(CType(currentCase("mugshot"), Byte()))
        Else
            pbMugshot.Image = Nothing ' Optionally set a default image if "mugshot" is missing or invalid
        End If

    End Sub


    ' Event handler for form load (optional, if you want to load data when the form loads)
    Private Sub viewPDL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load the data into the user control fields
        LoadDetails()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Logs("exited PDL view information")
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        selectedItem = 0

        If mainform IsNot Nothing Then
            mainform.SwitchToInmateHomeControl()
            Return
        End If

        Dim adminDashboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        selectedItem = 0

        If adminDashboard IsNot Nothing Then
            adminDashboard.SwitchToInmateHomeControl()
            Return
        End If

    End Sub

End Class
