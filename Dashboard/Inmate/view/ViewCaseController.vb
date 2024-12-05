Imports System.Globalization

Public Class ViewCaseController

    ' Form Load Event to Load Case Details
    Private Sub ViewCaseController_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Call LoadCaseDetails to populate the fields when the form is loaded
        Logs("viewing case")
        LoadCaseDetails()
    End Sub

    ' Method to load the data into the form
    Public Sub LoadCaseDetails()
        ' Ensure the data rows are not null
        If currentPDL Is Nothing OrElse currentCase Is Nothing Then
            MessageBox.Show("No data to load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Personal Information - Populate TextBoxes
        textName.Text = $"{currentPDL("first_name")} {currentPDL("middle_name")} {currentPDL("last_name")}"
        textSex.Text = currentPDL("sex").ToString()
        textCivil.Text = currentPDL("civil_status").ToString()
        textHeight.Text = currentPDL("height").ToString()
        textWeight.Text = currentPDL("weight").ToString()
        textPhone.Text = currentPDL("phone_num").ToString()
        textBirth.Text = Date.Parse(currentPDL("date_of_birth").ToString()).ToString("D")
        textID.Text = currentPDL("pdl_id").ToString()

        ' Criminal Case Information - Populate TextBoxes
        txtCrime.Text = currentCase("offence_charge").ToString()
        txtModus.Text = currentCase("modus_operandi").ToString()
        txtPlace.Text = currentCase("place_of_arrest").ToString()

        Dim arrestDateString As String = currentCase("date_of_arrest").ToString()
        Dim arrestDate As DateTime

        ' Try parsing the date
        If DateTime.TryParseExact(arrestDateString, "dd MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, arrestDate) Then
            txtArrest.Text = arrestDate.ToString("D")
        Else
            txtArrest.Text = "Invalid Date Format"
        End If
        txtArrest.Text = currentCase("date_of_arrest").ToString
        txtLawyer.Text = currentCase("lawyer").ToString()
        txtCell.Text = currentCase("cellblock").ToString()
        txtSentence.Text = currentCase("sentence").ToString()

        ' Calculate Remaining Sentence - Assuming "sentence" is in days

        ' Attempt to parse the arrest date
        If Date.TryParse(currentCase("date_of_arrest").ToString(), arrestDate) Then
            ' Attempt to parse the sentence (assuming it is in days)
            Dim sentenceDays As Integer
            If Integer.TryParse(currentCase("sentence").ToString(), sentenceDays) Then
                ' Calculate the total days from arrest to today
                Dim daysSinceArrest As Integer = (Date.Now - arrestDate).Days
                ' Subtract the sentence length from the days since arrest to get remaining days
                Dim remainingDays As Integer = Math.Max(sentenceDays - daysSinceArrest, 0)
                txtRemaining.Text = $"{remainingDays} days remaining"
            Else
                txtRemaining.Text = "Invalid sentence value"
            End If
        Else
            txtRemaining.Text = "Invalid Arrest Date"
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Logs("Exited view case")
        Dim mainform As MainDashboard = TryCast(Me.ParentForm, MainDashboard)

        If mainform IsNot Nothing Then
            mainform.SwitchToInmateHomeControl()
            Return
        End If


        Dim adminDasboard As AdminMainDashboard = TryCast(Me.ParentForm, AdminMainDashboard)
        If adminDasboard IsNot Nothing Then
            adminDasboard.SwitchToInmateHomeControl()
            Return
        End If

    End Sub
End Class
