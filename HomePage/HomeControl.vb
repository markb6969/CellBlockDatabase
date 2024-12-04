Public Class HomeControl
    Dim images(2) As Bitmap
    Dim descriptions(2) As String
    Dim pos As Integer = 0
    Private Sub HomeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rtxtHeadline.SelectAll()
        rtxtHeadline.SelectionAlignment = HorizontalAlignment.Center

        images(0) = My.Resources.Report1
        images(1) = My.Resources.Report2
        images(2) = My.Resources.Report3

        descriptions(0) = "MANILA, Philippines — Lawmakers believe that dismissed Bamban, Tarlac mayor Alice Guo’s agitation during the latest House of Representatives’ quad committee hearing was a sign that something about her true personality was revealed."
        descriptions(1) = "MANILA, Philippines — The Philippine National Police (PNP) has released footage Of the surrender And subsequent arrest Of Apollo Quiboloy On Sunday evening, ending the weeks-Long search For the Kingdom Of Jesus Christ leader. Quiboloy peacefully surrendered after being cornered by police, And following numerous negotiations With authorities, according To the PNP."
        descriptions(2) = "MANILA, Philippines – Authorities made a breakthrough on Thursday, October 10, when they caught Chinese national Lyu Dong, believed to be one of the big bosses of shady Philippine Offshore Gaming Operators (POGOs), particularly the scam hubs in Bamban, Tarlac and in Porac, Pampanga in Central Luzon."

        UpdateImagePositions()
    End Sub

    Private Sub UpdateImagePositions()
        pbMainImage.Image = images(pos)

        Dim prevPos As Integer = If(pos - 1 < 0, images.Length - 1, pos - 1)
        pbPrevImage.Image = images(prevPos)

        Dim nextPos As Integer = If(pos + 1 > images.Length - 1, 0, pos + 1)
        pbNextImage.Image = images(nextPos)

        rtxtHeadline.Text = descriptions(pos)
    End Sub

    Private Sub btnPreviousImg_Click(sender As Object, e As EventArgs) Handles btnPreviousImg.Click
        pos -= 1
        If pos < 0 Then
            pos = images.Length - 1
        End If

        UpdateImagePositions()
    End Sub

    Private Sub btnNextImg_Click(sender As Object, e As EventArgs) Handles btnNextImg.Click
        pos += 1
        If pos > images.Length - 1 Then
            pos = 0
        End If

        UpdateImagePositions()
    End Sub
End Class
