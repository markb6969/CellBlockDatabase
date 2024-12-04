<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportsControl
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rtxtDescripton = New System.Windows.Forms.RichTextBox()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.btnGenerate)
        Me.Panel1.Controls.Add(Me.rtxtDescripton)
        Me.Panel1.Controls.Add(Me.lblTitle)
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(569, 144)
        Me.Panel1.TabIndex = 1
        '
        'rtxtDescripton
        '
        Me.rtxtDescripton.BackColor = System.Drawing.Color.White
        Me.rtxtDescripton.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.rtxtDescripton.Cursor = System.Windows.Forms.Cursors.Default
        Me.rtxtDescripton.Enabled = False
        Me.rtxtDescripton.Font = New System.Drawing.Font("Poppins", 12.0!)
        Me.rtxtDescripton.Location = New System.Drawing.Point(50, 60)
        Me.rtxtDescripton.Name = "rtxtDescripton"
        Me.rtxtDescripton.ReadOnly = True
        Me.rtxtDescripton.Size = New System.Drawing.Size(459, 68)
        Me.rtxtDescripton.TabIndex = 1
        Me.rtxtDescripton.Text = "Description"
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Poppins SemiBold", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(20, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(103, 62)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Title"
        '
        'btnGenerate
        '
        Me.btnGenerate.BackColor = System.Drawing.Color.Transparent
        Me.btnGenerate.FlatAppearance.BorderSize = 0
        Me.btnGenerate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGenerate.Image = Global.CellBlockIM.My.Resources.Resources.pdf_file__2_
        Me.btnGenerate.Location = New System.Drawing.Point(504, 9)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(62, 56)
        Me.btnGenerate.TabIndex = 2
        Me.btnGenerate.UseVisualStyleBackColor = False
        '
        'ReportsControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ReportsControl"
        Me.Size = New System.Drawing.Size(575, 150)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents rtxtDescripton As RichTextBox
    Friend WithEvents lblTitle As Label
    Friend WithEvents btnGenerate As Button
End Class
