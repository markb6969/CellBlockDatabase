<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CameraCaptureForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.pbCamera = New System.Windows.Forms.PictureBox()
        Me.btnCapture = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.pbCamera, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pbCamera
        '
        Me.pbCamera.BackColor = System.Drawing.SystemColors.Control
        Me.pbCamera.Location = New System.Drawing.Point(24, 17)
        Me.pbCamera.Name = "pbCamera"
        Me.pbCamera.Size = New System.Drawing.Size(809, 627)
        Me.pbCamera.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbCamera.TabIndex = 0
        Me.pbCamera.TabStop = False
        '
        'btnCapture
        '
        Me.btnCapture.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCapture.Location = New System.Drawing.Point(921, 333)
        Me.btnCapture.Name = "btnCapture"
        Me.btnCapture.Size = New System.Drawing.Size(189, 81)
        Me.btnCapture.TabIndex = 1
        Me.btnCapture.Text = "Capture"
        Me.btnCapture.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(921, 438)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(189, 73)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Panel1.Controls.Add(Me.pbCamera)
        Me.Panel1.Location = New System.Drawing.Point(35, 22)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(853, 655)
        Me.Panel1.TabIndex = 3
        '
        'CameraCaptureForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1180, 705)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCapture)
        Me.Name = "CameraCaptureForm"
        Me.Text = "CameraCaptureForm"
        CType(Me.pbCamera, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pbCamera As PictureBox
    Friend WithEvents btnCapture As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents Panel1 As Panel
End Class
