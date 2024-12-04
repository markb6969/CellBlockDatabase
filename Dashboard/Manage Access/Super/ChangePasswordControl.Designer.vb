<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChangePasswordControl
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.txtConfirmPass = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtNewPass = New System.Windows.Forms.TextBox()
        Me.pnlTxtUsername = New System.Windows.Forms.Panel()
        Me.txtOldPass = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.btnChangePass = New System.Windows.Forms.Button()
        Me.btnUpdateInfo = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnAccountDetails = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnSuperAdmin = New System.Windows.Forms.Button()
        Me.btnAdmin = New System.Windows.Forms.Button()
        Me.btnClient = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.pnlTxtUsername.SuspendLayout()
        Me.Panel7.SuspendLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.57143!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 71.42857!))
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel7, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(-1, 117)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1297, 554)
        Me.TableLayoutPanel1.TabIndex = 47
        '
        'Panel2
        '
        Me.Panel2.AutoScroll = True
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.btnSave)
        Me.Panel2.Controls.Add(Me.Panel5)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.pnlTxtUsername)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(374, 4)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(919, 546)
        Me.Panel2.TabIndex = 5
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.Green
        Me.btnSave.FlatAppearance.BorderSize = 0
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSave.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.Location = New System.Drawing.Point(713, 596)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(295, 57)
        Me.btnSave.TabIndex = 24
        Me.btnSave.Text = "Save changes"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.White
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.txtConfirmPass)
        Me.Panel5.Location = New System.Drawing.Point(317, 321)
        Me.Panel5.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(521, 54)
        Me.Panel5.TabIndex = 21
        '
        'txtConfirmPass
        '
        Me.txtConfirmPass.BackColor = System.Drawing.Color.White
        Me.txtConfirmPass.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtConfirmPass.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtConfirmPass.ForeColor = System.Drawing.Color.DarkGray
        Me.txtConfirmPass.Location = New System.Drawing.Point(16, 11)
        Me.txtConfirmPass.Margin = New System.Windows.Forms.Padding(4)
        Me.txtConfirmPass.Name = "txtConfirmPass"
        Me.txtConfirmPass.Size = New System.Drawing.Size(495, 30)
        Me.txtConfirmPass.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.txtNewPass)
        Me.Panel3.Location = New System.Drawing.Point(317, 220)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(521, 54)
        Me.Panel3.TabIndex = 21
        '
        'txtNewPass
        '
        Me.txtNewPass.BackColor = System.Drawing.Color.White
        Me.txtNewPass.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtNewPass.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewPass.ForeColor = System.Drawing.Color.DarkGray
        Me.txtNewPass.Location = New System.Drawing.Point(16, 11)
        Me.txtNewPass.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNewPass.Name = "txtNewPass"
        Me.txtNewPass.Size = New System.Drawing.Size(495, 30)
        Me.txtNewPass.TabIndex = 0
        '
        'pnlTxtUsername
        '
        Me.pnlTxtUsername.BackColor = System.Drawing.Color.White
        Me.pnlTxtUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTxtUsername.Controls.Add(Me.txtOldPass)
        Me.pnlTxtUsername.Location = New System.Drawing.Point(317, 126)
        Me.pnlTxtUsername.Margin = New System.Windows.Forms.Padding(4)
        Me.pnlTxtUsername.Name = "pnlTxtUsername"
        Me.pnlTxtUsername.Size = New System.Drawing.Size(521, 54)
        Me.pnlTxtUsername.TabIndex = 21
        '
        'txtOldPass
        '
        Me.txtOldPass.BackColor = System.Drawing.Color.White
        Me.txtOldPass.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtOldPass.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOldPass.ForeColor = System.Drawing.Color.DarkGray
        Me.txtOldPass.Location = New System.Drawing.Point(16, 11)
        Me.txtOldPass.Margin = New System.Windows.Forms.Padding(4)
        Me.txtOldPass.Name = "txtOldPass"
        Me.txtOldPass.Size = New System.Drawing.Size(495, 30)
        Me.txtOldPass.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(109, 334)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(152, 26)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Confirm password"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(109, 240)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(122, 26)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "New password"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(109, 144)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(116, 26)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Old password"
        '
        'Panel7
        '
        Me.Panel7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel7.AutoScroll = True
        Me.Panel7.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel7.Controls.Add(Me.PictureBox8)
        Me.Panel7.Controls.Add(Me.PictureBox4)
        Me.Panel7.Controls.Add(Me.PictureBox3)
        Me.Panel7.Controls.Add(Me.PictureBox2)
        Me.Panel7.Controls.Add(Me.btnChangePass)
        Me.Panel7.Controls.Add(Me.btnUpdateInfo)
        Me.Panel7.Controls.Add(Me.btnBack)
        Me.Panel7.Controls.Add(Me.btnAccountDetails)
        Me.Panel7.Location = New System.Drawing.Point(4, 4)
        Me.Panel7.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(362, 546)
        Me.Panel7.TabIndex = 4
        '
        'PictureBox8
        '
        Me.PictureBox8.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(252, Byte), Integer))
        Me.PictureBox8.Image = Global.CellBlockIM.My.Resources.Resources.key
        Me.PictureBox8.Location = New System.Drawing.Point(16, 191)
        Me.PictureBox8.Margin = New System.Windows.Forms.Padding(4)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(33, 31)
        Me.PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox8.TabIndex = 24
        Me.PictureBox8.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.PictureBox4.Image = Global.CellBlockIM.My.Resources.Resources.edit
        Me.PictureBox4.Location = New System.Drawing.Point(15, 251)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(4)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(36, 33)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 24
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.PictureBox3.Image = Global.CellBlockIM.My.Resources.Resources.admin_with_cogwheels__1_
        Me.PictureBox3.Location = New System.Drawing.Point(16, 133)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(33, 31)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 24
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Image = Global.CellBlockIM.My.Resources.Resources.left
        Me.PictureBox2.Location = New System.Drawing.Point(16, 76)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(33, 31)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 24
        Me.PictureBox2.TabStop = False
        '
        'btnChangePass
        '
        Me.btnChangePass.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnChangePass.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(252, Byte), Integer))
        Me.btnChangePass.FlatAppearance.BorderSize = 0
        Me.btnChangePass.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChangePass.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnChangePass.Location = New System.Drawing.Point(-1, 178)
        Me.btnChangePass.Margin = New System.Windows.Forms.Padding(4)
        Me.btnChangePass.Name = "btnChangePass"
        Me.btnChangePass.Size = New System.Drawing.Size(433, 60)
        Me.btnChangePass.TabIndex = 0
        Me.btnChangePass.Text = "Change Password"
        Me.btnChangePass.UseVisualStyleBackColor = False
        '
        'btnUpdateInfo
        '
        Me.btnUpdateInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateInfo.AutoSize = True
        Me.btnUpdateInfo.FlatAppearance.BorderSize = 0
        Me.btnUpdateInfo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUpdateInfo.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdateInfo.Location = New System.Drawing.Point(-1, 238)
        Me.btnUpdateInfo.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateInfo.Name = "btnUpdateInfo"
        Me.btnUpdateInfo.Size = New System.Drawing.Size(437, 60)
        Me.btnUpdateInfo.TabIndex = 0
        Me.btnUpdateInfo.Text = "Update Personal Info"
        Me.btnUpdateInfo.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBack.FlatAppearance.BorderSize = 0
        Me.btnBack.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBack.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.Location = New System.Drawing.Point(-1, 62)
        Me.btnBack.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(433, 60)
        Me.btnBack.TabIndex = 0
        Me.btnBack.Text = "Back to Home"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnAccountDetails
        '
        Me.btnAccountDetails.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAccountDetails.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.btnAccountDetails.FlatAppearance.BorderSize = 0
        Me.btnAccountDetails.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAccountDetails.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAccountDetails.Location = New System.Drawing.Point(-1, 119)
        Me.btnAccountDetails.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAccountDetails.Name = "btnAccountDetails"
        Me.btnAccountDetails.Size = New System.Drawing.Size(433, 60)
        Me.btnAccountDetails.TabIndex = 0
        Me.btnAccountDetails.Text = "View Account Details"
        Me.btnAccountDetails.UseVisualStyleBackColor = False
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.btnSuperAdmin)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnAdmin)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnClient)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel2.Margin = New System.Windows.Forms.Padding(4)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(1296, 123)
        Me.FlowLayoutPanel2.TabIndex = 48
        '
        'btnSuperAdmin
        '
        Me.btnSuperAdmin.FlatAppearance.BorderSize = 0
        Me.btnSuperAdmin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSuperAdmin.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSuperAdmin.ForeColor = System.Drawing.Color.Teal
        Me.btnSuperAdmin.Location = New System.Drawing.Point(4, 4)
        Me.btnSuperAdmin.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSuperAdmin.Name = "btnSuperAdmin"
        Me.btnSuperAdmin.Size = New System.Drawing.Size(271, 95)
        Me.btnSuperAdmin.TabIndex = 24
        Me.btnSuperAdmin.Text = "Me"
        Me.btnSuperAdmin.UseVisualStyleBackColor = True
        '
        'btnAdmin
        '
        Me.btnAdmin.FlatAppearance.BorderSize = 0
        Me.btnAdmin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAdmin.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdmin.Location = New System.Drawing.Point(283, 4)
        Me.btnAdmin.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAdmin.Name = "btnAdmin"
        Me.btnAdmin.Size = New System.Drawing.Size(276, 95)
        Me.btnAdmin.TabIndex = 24
        Me.btnAdmin.Text = "Admin"
        Me.btnAdmin.UseVisualStyleBackColor = True
        '
        'btnClient
        '
        Me.btnClient.FlatAppearance.BorderSize = 0
        Me.btnClient.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClient.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClient.Location = New System.Drawing.Point(567, 4)
        Me.btnClient.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClient.Name = "btnClient"
        Me.btnClient.Size = New System.Drawing.Size(313, 95)
        Me.btnClient.TabIndex = 24
        Me.btnClient.Text = "Client"
        Me.btnClient.UseVisualStyleBackColor = True
        '
        'ChangePasswordControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.FlowLayoutPanel2)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "ChangePasswordControl"
        Me.Size = New System.Drawing.Size(1296, 671)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.pnlTxtUsername.ResumeLayout(False)
        Me.pnlTxtUsername.PerformLayout()
        Me.Panel7.ResumeLayout(False)
        Me.Panel7.PerformLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnSave As Button
    Friend WithEvents Panel5 As Panel
    Friend WithEvents txtConfirmPass As TextBox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents txtNewPass As TextBox
    Friend WithEvents pnlTxtUsername As Panel
    Friend WithEvents txtOldPass As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Panel7 As Panel
    Friend WithEvents PictureBox8 As PictureBox
    Friend WithEvents PictureBox4 As PictureBox
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents btnChangePass As Button
    Friend WithEvents btnUpdateInfo As Button
    Friend WithEvents btnBack As Button
    Friend WithEvents btnAccountDetails As Button
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents btnSuperAdmin As Button
    Friend WithEvents btnAdmin As Button
    Friend WithEvents btnClient As Button
End Class
