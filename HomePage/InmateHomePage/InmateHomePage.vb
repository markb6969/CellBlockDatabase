Public Class InmateHomePage
    Private Sub InmateHomePage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.SelectAll()
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox2.SelectAll()
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox3.SelectAll()
        RichTextBox3.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox4.SelectAll()
        RichTextBox4.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox5.SelectAll()
        RichTextBox5.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox6.SelectAll()
        RichTextBox6.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox7.SelectAll()
        RichTextBox7.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox8.SelectAll()
        RichTextBox8.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox9.SelectAll()
        RichTextBox9.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox10.SelectAll()
        RichTextBox10.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox12.SelectAll()
        RichTextBox12.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox14.SelectAll()
        RichTextBox14.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox15.SelectAll()
        RichTextBox15.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox16.SelectAll()
        RichTextBox16.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub btnGoGenerate_Click(sender As Object, e As EventArgs)
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToPDLInfoPageControl()
    End Sub



    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToPDLInfoPageControl()
    End Sub

    Private Sub btnVisitation_Click_1(sender As Object, e As EventArgs) Handles btnVisitation.Click
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToVisitationControl()
    End Sub

    Private Sub btnReportConcern_Click_1(sender As Object, e As EventArgs) Handles btnReportConcern.Click
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToReportConcernPageControl()
    End Sub
End Class
