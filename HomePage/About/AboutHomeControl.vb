Public Class AboutHomeControl
    Private Sub AboutHomeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.SelectAll()
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub btnLearnMore_Click(sender As Object, e As EventArgs) Handles btnLearnMore.Click
        Dim mainForm As Home = TryCast(Me.ParentForm, Home)
        mainForm.SwitchToAboutHistoryControl()
    End Sub
End Class
