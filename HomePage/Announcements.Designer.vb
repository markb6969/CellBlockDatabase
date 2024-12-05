<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Announcements
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.rtxtDescripton = New System.Windows.Forms.RichTextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Poppins SemiBold", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(16, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(103, 62)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Title"
        '
        'rtxtDescripton
        '
        Me.rtxtDescripton.BackColor = System.Drawing.Color.White
        Me.rtxtDescripton.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.rtxtDescripton.Cursor = System.Windows.Forms.Cursors.Default
        Me.rtxtDescripton.Enabled = False
        Me.rtxtDescripton.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtxtDescripton.Location = New System.Drawing.Point(64, 50)
        Me.rtxtDescripton.Name = "rtxtDescripton"
        Me.rtxtDescripton.ReadOnly = True
        Me.rtxtDescripton.Size = New System.Drawing.Size(624, 91)
        Me.rtxtDescripton.TabIndex = 1
        Me.rtxtDescripton.Text = "Description"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.rtxtDescripton)
        Me.Panel1.Controls.Add(Me.lblTitle)
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(754, 144)
        Me.Panel1.TabIndex = 0
        '
        'Announcements
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Controls.Add(Me.Panel1)
        Me.Name = "Announcements"
        Me.Size = New System.Drawing.Size(760, 153)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblTitle As Label
    Friend WithEvents rtxtDescripton As RichTextBox
    Friend WithEvents Panel1 As Panel
End Class
