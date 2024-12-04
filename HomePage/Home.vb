Public Class Home
    Dim images(2) As Bitmap
    Dim descriptions(2) As String
    Dim pos As Integer = 0
    Private Sub Home_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rtxtHeadline.SelectAll()
        rtxtHeadline.SelectionAlignment = HorizontalAlignment.Center
        openConn(db_name)
        isConnectedToLocalServer()


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


    ' SWITCH PANELS

    Public Sub switchPanel(childPanel As Panel, newControl As UserControl)
        childPanel.Controls.Clear()
        newControl.Dock = DockStyle.Fill
        childPanel.Controls.Add(newControl)
        newControl.Show()
    End Sub


    Public Sub SwitchToMainHomeControl()
        switchPanel(pnlMain, New HomeControl())
    End Sub
    Public Sub SwitchToAboutHomeControl()
        switchPanel(pnlMain, New AboutHomeControl())
    End Sub
    Public Sub SwitchToAboutHistoryControl()
        switchPanel(pnlMain, New AboutHistoryControl())
    End Sub
    Public Sub SwitchToInmatePageControl()
        switchPanel(pnlMain, New InmateHomePage())
    End Sub
    Public Sub SwitchToMapPageControl()
        switchPanel(pnlMain, New MapPageControl())
    End Sub
    Public Sub SwitchToContactPageControl()
        switchPanel(pnlMain, New ContactPageControl())
    End Sub
    Public Sub SwitchToPDLInfoPageControl()
        switchPanel(pnlMain, New InmateReportControls())
    End Sub
    Public Sub SwitchToReportConcernPageControl()
        switchPanel(pnlMain, New ReportConcernControl())
    End Sub
    Public Sub SwitchToVisitationControl()
        switchPanel(pnlMain, New VisitationControl())
    End Sub
    Public Sub SwitchToCreateAccountControl()
        switchPanel(pnlMain, New AccountHomeControl())
    End Sub

    Public Sub SwitchToAccountCreateControl()
        switchPanel(pnlMain, New CreateAccountControl())
    End Sub
    Public Sub SwitchToAccountPasswordControl()
        switchPanel(pnlMain, New CreatePasswordControl())
    End Sub
    Public Sub SwitchToAccountImageControl()
        switchPanel(pnlMain, New CreateImageControl())
    End Sub

    'NAV
    Private Sub btnAbout1_Click(sender As Object, e As EventArgs) Handles btnAbout1.Click
        SwitchToAboutHomeControl()
    End Sub

    Private Sub btnHome_Click(sender As Object, e As EventArgs) Handles btnHome.Click
        SwitchToMainHomeControl()
    End Sub

    Private Sub btnInmate1_Click(sender As Object, e As EventArgs) Handles btnInmate1.Click
        SwitchToInmatePageControl()
    End Sub

    Private Sub btnLocation1_Click(sender As Object, e As EventArgs)
        SwitchToMapPageControl()
    End Sub

    Private Sub btnContact1_Click(sender As Object, e As EventArgs) Handles btnContact1.Click
        SwitchToContactPageControl()
    End Sub

    Private Sub btnAccount_Click(sender As Object, e As EventArgs) Handles btnAccount.Click
        SwitchToCreateAccountControl()
    End Sub
End Class