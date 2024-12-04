Public Class AddEditDeleteControl
    Public Event AddButtonClick()
    Public Event EditButtonClick()
    Public Event DeleteButtonClick()

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        RaiseEvent AddButtonClick()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        RaiseEvent EditButtonClick()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        RaiseEvent DeleteButtonClick()
    End Sub

    Private Sub AddEditDeleteControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
