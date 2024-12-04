<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddArrestingOfficerControl
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
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.txtOfficerUnit = New System.Windows.Forms.TextBox()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.txtOfficerRank = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtOfficerName = New System.Windows.Forms.TextBox()
        Me.Panel9.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(469, 41)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(33, 21)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "Unit"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(352, 41)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 21)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "Rank"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(9, 41)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(103, 21)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "Arresting Officer"
        '
        'Panel9
        '
        Me.Panel9.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel9.Controls.Add(Me.txtOfficerUnit)
        Me.Panel9.Location = New System.Drawing.Point(473, 65)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(110, 44)
        Me.Panel9.TabIndex = 38
        '
        'txtOfficerUnit
        '
        Me.txtOfficerUnit.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerUnit.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerUnit.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerUnit.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerUnit.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerUnit.Name = "txtOfficerUnit"
        Me.txtOfficerUnit.Size = New System.Drawing.Size(91, 24)
        Me.txtOfficerUnit.TabIndex = 0
        Me.txtOfficerUnit.Text = "Rank"
        '
        'Panel8
        '
        Me.Panel8.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel8.Controls.Add(Me.txtOfficerRank)
        Me.Panel8.Location = New System.Drawing.Point(356, 65)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(110, 44)
        Me.Panel8.TabIndex = 39
        '
        'txtOfficerRank
        '
        Me.txtOfficerRank.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerRank.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerRank.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerRank.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerRank.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerRank.Name = "txtOfficerRank"
        Me.txtOfficerRank.Size = New System.Drawing.Size(91, 24)
        Me.txtOfficerRank.TabIndex = 0
        Me.txtOfficerRank.Text = "Rank"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.txtOfficerName)
        Me.Panel3.Location = New System.Drawing.Point(13, 65)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(335, 44)
        Me.Panel3.TabIndex = 40
        '
        'txtOfficerName
        '
        Me.txtOfficerName.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtOfficerName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOfficerName.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOfficerName.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOfficerName.Location = New System.Drawing.Point(13, 9)
        Me.txtOfficerName.Name = "txtOfficerName"
        Me.txtOfficerName.Size = New System.Drawing.Size(311, 24)
        Me.txtOfficerName.TabIndex = 0
        Me.txtOfficerName.Text = "Last Name/ First Name/ Middle Name"
        '
        'AddArrestingOfficerControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel9)
        Me.Controls.Add(Me.Panel8)
        Me.Controls.Add(Me.Panel3)
        Me.Name = "AddArrestingOfficerControl"
        Me.Size = New System.Drawing.Size(609, 150)
        Me.Panel9.ResumeLayout(False)
        Me.Panel9.PerformLayout()
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Panel9 As Panel
    Friend WithEvents txtOfficerUnit As TextBox
    Friend WithEvents Panel8 As Panel
    Friend WithEvents txtOfficerRank As TextBox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents txtOfficerName As TextBox
End Class
