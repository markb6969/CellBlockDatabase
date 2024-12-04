Imports System.IO

Public Class VisitorApproveControl

    ' Load the current pending info and fill the form
    Private Sub FillAdminFields()
        Logs("Opened Visitation approval form")

        ' TextBoxes for visitor's first, middle, and last name
        txtVisitorFirstName.Text = currentPending("visitor_first_name").ToString()
        txtVisitorMiddleName.Text = currentPending("visitor_middle_name").ToString()
        txtVisitorLastName.Text = currentPending("visitor_last_name").ToString()

        ' TextBox for PDL Name
        txtPDLName.Text = currentPending("pdl_name").ToString()

        ' ComboBox for purpose and ID type
        cmbPurpose.SelectedItem = currentPending("purpose").ToString()
        cmbIdentification.SelectedItem = currentPending("id_type").ToString()

        ' ComboBox for relationship (if applicable)
        cmbRelationship.SelectedItem = currentPending("relationship").ToString()

        ' RadioButtons for victim status and gender
        If currentPending("is_victim").ToString() = "Yes" Then
            rbYes.Checked = True
        Else
            rbNo.Checked = True
        End If

        If currentPending("gender").ToString() = "Female" Then
            rbFemale.Checked = True
        Else
            rbMale.Checked = True
        End If

        ' DateTimePickers for birthdate and visit date
        dtBirthdate.Text = currentPending("visit_date").ToString
        DateTimePicker1.Value = Convert.ToDateTime(currentPending("visit_date"))

        ' TextBoxes for address details
        txtStreet.Text = currentPending("street").ToString()
        txtMunicipality.Text = currentPending("municipality").ToString()
        txtCity.Text = currentPending("city").ToString()
        txtRegion.Text = currentPending("region").ToString()
        txtZip.Text = currentPending("zip").ToString()
        txtCountry.Text = currentPending("country").ToString()

        ' TextBoxes for contact details
        txtContactNumber.Text = currentPending("contact_number").ToString()
        txtEmail.Text = currentPending("email").ToString()

        ' If the ID image exists, convert and assign it (you can write a method to convert byte array to image)
        If currentPending("valid_id") IsNot DBNull.Value Then
            pbID.Image = ByteArrayToImage(CType(currentPending("valid_id"), Byte()))
        End If
    End Sub

    ' Method to convert ByteArray to Image (if needed for displaying ID image)
    Private Function ByteArrayToImage(byteArray As Byte()) As Image
        Using ms As New MemoryStream(byteArray)
            Return Image.FromStream(ms)
        End Using
    End Function

    ' When the form loads, populate the fields with current pending data
    Private Sub VisitorApproveControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillAdminFields()
    End Sub

    ' Event handler when the approve button is clicked
    Private Sub btnApprove_Click(sender As Object, e As EventArgs) Handles btnApprove.Click

        ' Dictionary for the update data
        Dim data As New Dictionary(Of String, Object) From {
            {"status", "approved"}
        }

        ' Condition for finding the specific pending visitor
        Dim condition As New Dictionary(Of String, Object) From {
            {"visitor_id", currentPending("visitor_id")}
        }

        ' Update the record in the "visitors" table
        UpdateRecord("visitors", data, condition)
        Logs("Visitation approved, visitorID" + currentPending("visitor_id"))

        ' Switch to the Visitation Status control on the main form
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToVisitationStatusControl()

        ' Clear currentPending after approval
        currentPending = Nothing
    End Sub
    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        ' Dictionary for the update data
        Dim data As New Dictionary(Of String, Object) From {
            {"status", "rejected"}
        }

        ' Condition for finding the specific pending visitor
        Dim condition As New Dictionary(Of String, Object) From {
            {"visitor_id", currentPending("visitor_id")}
        }

        ' Update the record in the "visitors" table
        UpdateRecord("visitors", data, condition)
        Logs("Visitation rejected, visitorID" + currentPending("visitor_id"))

        ' Switch to the Visitation Status control on the main form
        Dim mainForm As MainDashboard = TryCast(Me.ParentForm, MainDashboard)
        mainForm.SwitchToVisitationStatusControl()

        ' Clear currentPending after rejection
        currentPending = Nothing
    End Sub

End Class

