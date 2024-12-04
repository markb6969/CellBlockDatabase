<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddMultipleOffenseControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.Panel25 = New System.Windows.Forms.Panel()
        Me.dtDateofArrest2 = New System.Windows.Forms.DateTimePicker()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.txtLastName2 = New System.Windows.Forms.TextBox()
        Me.lblModus = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.txtModus2 = New System.Windows.Forms.TextBox()
        Me.lblOffense = New System.Windows.Forms.Label()
        Me.pnlTxtUsername = New System.Windows.Forms.Panel()
        Me.txtOffense2 = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.txtOfficerUnit4 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.txtOfficerUnit3 = New System.Windows.Forms.TextBox()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.txtOfficerRank4 = New System.Windows.Forms.TextBox()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.txtOfficerName4 = New System.Windows.Forms.TextBox()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.txtOfficerRank3 = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtOfficerName3 = New System.Windows.Forms.TextBox()
        Me.Panel25.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.pnlTxtUsername.SuspendLayout()
        Me.Panel11.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.Panel10.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblDate
        '
        Me.lblDate.AutoSize = True
        Me.lblDate.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate.Location = New System.Drawing.Point(29, 259)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(88, 21)
        Me.lblDate.TabIndex = 49
        Me.lblDate.Text = "Date of Arrest"
        '
        'Panel25
        '
        Me.Panel25.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel25.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel25.Controls.Add(Me.dtDateofArrest2)
        Me.Panel25.Location = New System.Drawing.Point(33, 283)
        Me.Panel25.Name = "Panel25"
        Me.Panel25.Size = New System.Drawing.Size(335, 44)
        Me.Panel25.TabIndex = 48
        '
        'dtDateofArrest2
        '
        Me.dtDateofArrest2.Font = New System.Drawing.Font("Poppins", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtDateofArrest2.Location = New System.Drawing.Point(51, 7)
        Me.dtDateofArrest2.Name = "dtDateofArrest2"
        Me.dtDateofArrest2.Size = New System.Drawing.Size(230, 27)
        Me.dtDateofArrest2.TabIndex = 24
        Me.dtDateofArrest2.Value = New Date(2024, 10, 14, 11, 6, 41, 0)
        '
        'lblLastName
        '
        Me.lblLastName.AutoSize = True
        Me.lblLastName.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastName.Location = New System.Drawing.Point(29, 179)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(99, 21)
        Me.lblLastName.TabIndex = 47
        Me.lblLastName.Text = "Where Arrested"
        '
        'Panel5
        '
        Me.Panel5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel5.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.txtLastName2)
        Me.Panel5.Location = New System.Drawing.Point(33, 203)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(335, 44)
        Me.Panel5.TabIndex = 46
        '
        'txtLastName2
        '
        Me.txtLastName2.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtLastName2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtLastName2.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastName2.ForeColor = System.Drawing.Color.DarkGray
        Me.txtLastName2.Location = New System.Drawing.Point(13, 9)
        Me.txtLastName2.Name = "txtLastName2"
        Me.txtLastName2.Size = New System.Drawing.Size(311, 24)
        Me.txtLastName2.TabIndex = 0
        Me.txtLastName2.Text = "Address"
        '
        'lblModus
        '
        Me.lblModus.AutoSize = True
        Me.lblModus.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModus.Location = New System.Drawing.Point(31, 99)
        Me.lblModus.Name = "lblModus"
        Me.lblModus.Size = New System.Drawing.Size(113, 21)
        Me.lblModus.TabIndex = 45
        Me.lblModus.Text = "Modus Operanda"
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Controls.Add(Me.txtModus2)
        Me.Panel6.Location = New System.Drawing.Point(34, 123)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(334, 44)
        Me.Panel6.TabIndex = 44
        '
        'txtModus2
        '
        Me.txtModus2.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtModus2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtModus2.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtModus2.ForeColor = System.Drawing.Color.DarkGray
        Me.txtModus2.Location = New System.Drawing.Point(12, 9)
        Me.txtModus2.Name = "txtModus2"
        Me.txtModus2.Size = New System.Drawing.Size(311, 24)
        Me.txtModus2.TabIndex = 0
        Me.txtModus2.Text = "Method used"
        '
        'lblOffense
        '
        Me.lblOffense.AutoSize = True
        Me.lblOffense.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOffense.Location = New System.Drawing.Point(30, 18)
        Me.lblOffense.Name = "lblOffense"
        Me.lblOffense.Size = New System.Drawing.Size(102, 21)
        Me.lblOffense.TabIndex = 43
        Me.lblOffense.Text = "Offense Charge"
        '
        'pnlTxtUsername
        '
        Me.pnlTxtUsername.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTxtUsername.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.pnlTxtUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTxtUsername.Controls.Add(Me.txtOffense2)
        Me.pnlTxtUsername.Location = New System.Drawing.Point(34, 42)
        Me.pnlTxtUsername.Name = "pnlTxtUsername"
        Me.pnlTxtUsername.Size = New System.Drawing.Size(334, 44)
        Me.pnlTxtUsername.TabIndex = 42
        '
        'txtOffense2
        '
        Me.txtOffense2.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOffense2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOffense2.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOffense2.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOffense2.Location = New System.Drawing.Point(12, 9)
        Me.txtOffense2.Name = "txtOffense2"
        Me.txtOffense2.Size = New System.Drawing.Size(311, 24)
        Me.txtOffense2.TabIndex = 0
        Me.txtOffense2.Text = "Nature of offense"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(493, 440)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(33, 21)
        Me.Label8.TabIndex = 56
        Me.Label8.Text = "Unit"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(491, 359)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(33, 21)
        Me.Label6.TabIndex = 57
        Me.Label6.Text = "Unit"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(376, 440)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(39, 21)
        Me.Label7.TabIndex = 58
        Me.Label7.Text = "Rank"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(374, 359)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 21)
        Me.Label5.TabIndex = 59
        Me.Label5.Text = "Rank"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(33, 440)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(103, 21)
        Me.Label4.TabIndex = 60
        Me.Label4.Text = "Arresting Officer"
        '
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel11.Controls.Add(Me.txtOfficerUnit4)
        Me.Panel11.Location = New System.Drawing.Point(497, 464)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(110, 44)
        Me.Panel11.TabIndex = 50
        '
        'txtOfficerUnit4
        '
        Me.txtOfficerUnit4.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerUnit4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerUnit4.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerUnit4.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerUnit4.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerUnit4.Name = "txtOfficerUnit4"
        Me.txtOfficerUnit4.Size = New System.Drawing.Size(91, 24)
        Me.txtOfficerUnit4.TabIndex = 0
        Me.txtOfficerUnit4.Text = "Rank"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(31, 359)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 21)
        Me.Label3.TabIndex = 61
        Me.Label3.Text = "Arresting Officer"
        '
        'Panel9
        '
        Me.Panel9.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel9.Controls.Add(Me.txtOfficerUnit3)
        Me.Panel9.Location = New System.Drawing.Point(495, 383)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(110, 44)
        Me.Panel9.TabIndex = 51
        '
        'txtOfficerUnit3
        '
        Me.txtOfficerUnit3.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerUnit3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerUnit3.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerUnit3.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerUnit3.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerUnit3.Name = "txtOfficerUnit3"
        Me.txtOfficerUnit3.Size = New System.Drawing.Size(91, 24)
        Me.txtOfficerUnit3.TabIndex = 0
        Me.txtOfficerUnit3.Text = "Rank"
        '
        'Panel10
        '
        Me.Panel10.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel10.Controls.Add(Me.txtOfficerRank4)
        Me.Panel10.Location = New System.Drawing.Point(380, 464)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(110, 44)
        Me.Panel10.TabIndex = 52
        '
        'txtOfficerRank4
        '
        Me.txtOfficerRank4.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerRank4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerRank4.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerRank4.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerRank4.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerRank4.Name = "txtOfficerRank4"
        Me.txtOfficerRank4.Size = New System.Drawing.Size(91, 24)
        Me.txtOfficerRank4.TabIndex = 0
        Me.txtOfficerRank4.Text = "Rank"
        '
        'Panel7
        '
        Me.Panel7.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel7.Controls.Add(Me.txtOfficerName4)
        Me.Panel7.Location = New System.Drawing.Point(37, 464)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(335, 44)
        Me.Panel7.TabIndex = 53
        '
        'txtOfficerName4
        '
        Me.txtOfficerName4.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerName4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerName4.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerName4.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerName4.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerName4.Name = "txtOfficerName4"
        Me.txtOfficerName4.Size = New System.Drawing.Size(311, 24)
        Me.txtOfficerName4.TabIndex = 0
        Me.txtOfficerName4.Text = "Last Name/ First Name/ Middle Name"
        '
        'Panel8
        '
        Me.Panel8.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel8.Controls.Add(Me.txtOfficerRank3)
        Me.Panel8.Location = New System.Drawing.Point(378, 383)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(110, 44)
        Me.Panel8.TabIndex = 54
        '
        'txtOfficerRank3
        '
        Me.txtOfficerRank3.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerRank3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerRank3.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerRank3.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerRank3.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerRank3.Name = "txtOfficerRank3"
        Me.txtOfficerRank3.Size = New System.Drawing.Size(91, 24)
        Me.txtOfficerRank3.TabIndex = 0
        Me.txtOfficerRank3.Text = "Rank"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.txtOfficerName3)
        Me.Panel3.Location = New System.Drawing.Point(35, 383)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(335, 44)
        Me.Panel3.TabIndex = 55
        '
        'txtOfficerName3
        '
        Me.txtOfficerName3.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerName3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerName3.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerName3.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerName3.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerName3.Name = "txtOfficerName3"
        Me.txtOfficerName3.Size = New System.Drawing.Size(311, 24)
        Me.txtOfficerName3.TabIndex = 0
        Me.txtOfficerName3.Text = "Last Name/ First Name/ Middle Name"
        '
        'AddMultipleOffenseControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Panel11)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel9)
        Me.Controls.Add(Me.Panel10)
        Me.Controls.Add(Me.Panel7)
        Me.Controls.Add(Me.Panel8)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.Panel25)
        Me.Controls.Add(Me.lblLastName)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.lblModus)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.lblOffense)
        Me.Controls.Add(Me.pnlTxtUsername)
        Me.Name = "AddMultipleOffenseControl"
        Me.Size = New System.Drawing.Size(914, 576)
        Me.Panel25.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.pnlTxtUsername.ResumeLayout(False)
        Me.pnlTxtUsername.PerformLayout()
        Me.Panel11.ResumeLayout(False)
        Me.Panel11.PerformLayout()
        Me.Panel9.ResumeLayout(False)
        Me.Panel9.PerformLayout()
        Me.Panel10.ResumeLayout(False)
        Me.Panel10.PerformLayout()
        Me.Panel7.ResumeLayout(False)
        Me.Panel7.PerformLayout()
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblDate As Label
    Friend WithEvents Panel25 As Panel
    Friend WithEvents dtDateofArrest2 As DateTimePicker
    Friend WithEvents lblLastName As Label
    Friend WithEvents Panel5 As Panel
    Friend WithEvents txtLastName2 As TextBox
    Friend WithEvents lblModus As Label
    Friend WithEvents Panel6 As Panel
    Friend WithEvents txtModus2 As TextBox
    Friend WithEvents lblOffense As Label
    Friend WithEvents pnlTxtUsername As Panel
    Friend WithEvents txtOffense2 As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Panel11 As Panel
    Friend WithEvents txtOfficerUnit4 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Panel9 As Panel
    Friend WithEvents txtOfficerUnit3 As TextBox
    Friend WithEvents Panel10 As Panel
    Friend WithEvents txtOfficerRank4 As TextBox
    Friend WithEvents Panel7 As Panel
    Friend WithEvents txtOfficerName4 As TextBox
    Friend WithEvents Panel8 As Panel
    Friend WithEvents txtOfficerRank3 As TextBox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents txtOfficerName3 As TextBox
End Class
