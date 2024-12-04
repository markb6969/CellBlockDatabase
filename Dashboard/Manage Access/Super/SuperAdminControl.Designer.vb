<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SuperAdminControl
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
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.txtPhoneNum = New System.Windows.Forms.TextBox()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.pnlTxtUsername = New System.Windows.Forms.Panel()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblShowFullName = New System.Windows.Forms.Label()
        Me.lblShowAdress = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.pbProfilePic = New System.Windows.Forms.PictureBox()
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
        Me.Panel6.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.pnlTxtUsername.SuspendLayout()
        CType(Me.pbProfilePic, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 117)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1296, 554)
        Me.TableLayoutPanel1.TabIndex = 40
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.AutoScroll = True
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Panel6)
        Me.Panel2.Controls.Add(Me.Panel5)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.Panel1)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.pnlTxtUsername)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.lblShowFullName)
        Me.Panel2.Controls.Add(Me.lblShowAdress)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.lblFirstName)
        Me.Panel2.Controls.Add(Me.pbProfilePic)
        Me.Panel2.Location = New System.Drawing.Point(374, 4)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(918, 546)
        Me.Panel2.TabIndex = 2
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.White
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Controls.Add(Me.txtPhoneNum)
        Me.Panel6.Location = New System.Drawing.Point(317, 617)
        Me.Panel6.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(690, 54)
        Me.Panel6.TabIndex = 21
        '
        'txtPhoneNum
        '
        Me.txtPhoneNum.BackColor = System.Drawing.Color.White
        Me.txtPhoneNum.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtPhoneNum.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhoneNum.ForeColor = System.Drawing.Color.DarkGray
        Me.txtPhoneNum.Location = New System.Drawing.Point(16, 11)
        Me.txtPhoneNum.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtPhoneNum.Name = "txtPhoneNum"
        Me.txtPhoneNum.Size = New System.Drawing.Size(668, 30)
        Me.txtPhoneNum.TabIndex = 0
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.White
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.txtEmail)
        Me.Panel5.Location = New System.Drawing.Point(317, 524)
        Me.Panel5.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(690, 54)
        Me.Panel5.TabIndex = 21
        '
        'txtEmail
        '
        Me.txtEmail.BackColor = System.Drawing.Color.White
        Me.txtEmail.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtEmail.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmail.ForeColor = System.Drawing.Color.DarkGray
        Me.txtEmail.Location = New System.Drawing.Point(16, 11)
        Me.txtEmail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(668, 30)
        Me.txtEmail.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.txtUsername)
        Me.Panel3.Location = New System.Drawing.Point(317, 423)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(690, 54)
        Me.Panel3.TabIndex = 21
        '
        'txtUsername
        '
        Me.txtUsername.BackColor = System.Drawing.Color.White
        Me.txtUsername.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtUsername.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsername.ForeColor = System.Drawing.Color.DarkGray
        Me.txtUsername.Location = New System.Drawing.Point(16, 11)
        Me.txtUsername.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(668, 30)
        Me.txtUsername.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.txtLastName)
        Me.Panel1.Location = New System.Drawing.Point(673, 329)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(334, 54)
        Me.Panel1.TabIndex = 21
        '
        'txtLastName
        '
        Me.txtLastName.BackColor = System.Drawing.Color.White
        Me.txtLastName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtLastName.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastName.ForeColor = System.Drawing.Color.DarkGray
        Me.txtLastName.Location = New System.Drawing.Point(16, 11)
        Me.txtLastName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(312, 30)
        Me.txtLastName.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(109, 635)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(124, 26)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Phone number"
        '
        'pnlTxtUsername
        '
        Me.pnlTxtUsername.BackColor = System.Drawing.Color.White
        Me.pnlTxtUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTxtUsername.Controls.Add(Me.txtFirstName)
        Me.pnlTxtUsername.Location = New System.Drawing.Point(317, 329)
        Me.pnlTxtUsername.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pnlTxtUsername.Name = "pnlTxtUsername"
        Me.pnlTxtUsername.Size = New System.Drawing.Size(334, 54)
        Me.pnlTxtUsername.TabIndex = 21
        '
        'txtFirstName
        '
        Me.txtFirstName.BackColor = System.Drawing.Color.White
        Me.txtFirstName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtFirstName.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstName.ForeColor = System.Drawing.Color.DarkGray
        Me.txtFirstName.Location = New System.Drawing.Point(16, 11)
        Me.txtFirstName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(312, 30)
        Me.txtFirstName.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(109, 537)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 26)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Email"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(109, 443)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 26)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Username"
        '
        'lblShowFullName
        '
        Me.lblShowFullName.AutoSize = True
        Me.lblShowFullName.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShowFullName.Location = New System.Drawing.Point(345, 162)
        Me.lblShowFullName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblShowFullName.Name = "lblShowFullName"
        Me.lblShowFullName.Size = New System.Drawing.Size(192, 39)
        Me.lblShowFullName.TabIndex = 20
        Me.lblShowFullName.Text = "Full Name "
        '
        'lblShowAdress
        '
        Me.lblShowAdress.AutoSize = True
        Me.lblShowAdress.Font = New System.Drawing.Font("Poppins", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShowAdress.Location = New System.Drawing.Point(349, 214)
        Me.lblShowAdress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblShowAdress.Name = "lblShowAdress"
        Me.lblShowAdress.Size = New System.Drawing.Size(100, 42)
        Me.lblShowAdress.TabIndex = 20
        Me.lblShowAdress.Text = "Adress"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(109, 347)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(61, 26)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Name "
        '
        'lblFirstName
        '
        Me.lblFirstName.AutoSize = True
        Me.lblFirstName.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFirstName.Location = New System.Drawing.Point(109, 89)
        Me.lblFirstName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(115, 26)
        Me.lblFirstName.TabIndex = 20
        Me.lblFirstName.Text = "Profile Picture"
        '
        'pbProfilePic
        '
        Me.pbProfilePic.Location = New System.Drawing.Point(115, 118)
        Me.pbProfilePic.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.pbProfilePic.Name = "pbProfilePic"
        Me.pbProfilePic.Size = New System.Drawing.Size(203, 187)
        Me.pbProfilePic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbProfilePic.TabIndex = 0
        Me.pbProfilePic.TabStop = False
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
        Me.Panel7.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(362, 546)
        Me.Panel7.TabIndex = 4
        '
        'PictureBox8
        '
        Me.PictureBox8.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.PictureBox8.Image = Global.CellBlockIM.My.Resources.Resources.key
        Me.PictureBox8.Location = New System.Drawing.Point(16, 191)
        Me.PictureBox8.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(36, 33)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 24
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(252, Byte), Integer))
        Me.PictureBox3.Image = Global.CellBlockIM.My.Resources.Resources.admin_with_cogwheels__1_
        Me.PictureBox3.Location = New System.Drawing.Point(16, 133)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.btnChangePass.FlatAppearance.BorderSize = 0
        Me.btnChangePass.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChangePass.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnChangePass.Location = New System.Drawing.Point(-1, 178)
        Me.btnChangePass.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnChangePass.Name = "btnChangePass"
        Me.btnChangePass.Size = New System.Drawing.Size(433, 60)
        Me.btnChangePass.TabIndex = 0
        Me.btnChangePass.Text = "Change Password"
        Me.btnChangePass.UseVisualStyleBackColor = True
        '
        'btnUpdateInfo
        '
        Me.btnUpdateInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateInfo.FlatAppearance.BorderSize = 0
        Me.btnUpdateInfo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUpdateInfo.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdateInfo.Location = New System.Drawing.Point(-1, 238)
        Me.btnUpdateInfo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnUpdateInfo.Name = "btnUpdateInfo"
        Me.btnUpdateInfo.Size = New System.Drawing.Size(433, 60)
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
        Me.btnBack.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.btnAccountDetails.BackColor = System.Drawing.Color.FromArgb(CType(CType(236, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(252, Byte), Integer))
        Me.btnAccountDetails.FlatAppearance.BorderSize = 0
        Me.btnAccountDetails.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAccountDetails.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAccountDetails.Location = New System.Drawing.Point(-1, 119)
        Me.btnAccountDetails.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.FlowLayoutPanel2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(1296, 123)
        Me.FlowLayoutPanel2.TabIndex = 41
        '
        'btnSuperAdmin
        '
        Me.btnSuperAdmin.FlatAppearance.BorderSize = 0
        Me.btnSuperAdmin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSuperAdmin.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSuperAdmin.ForeColor = System.Drawing.Color.Teal
        Me.btnSuperAdmin.Location = New System.Drawing.Point(4, 4)
        Me.btnSuperAdmin.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.btnAdmin.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.btnClient.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClient.Name = "btnClient"
        Me.btnClient.Size = New System.Drawing.Size(313, 95)
        Me.btnClient.TabIndex = 24
        Me.btnClient.Text = "Client"
        Me.btnClient.UseVisualStyleBackColor = True
        '
        'SuperAdminControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.FlowLayoutPanel2)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "SuperAdminControl"
        Me.Size = New System.Drawing.Size(1296, 671)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pnlTxtUsername.ResumeLayout(False)
        Me.pnlTxtUsername.PerformLayout()
        CType(Me.pbProfilePic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel7.ResumeLayout(False)
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel6 As Panel
    Friend WithEvents txtPhoneNum As TextBox
    Friend WithEvents Panel5 As Panel
    Friend WithEvents txtEmail As TextBox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents txtUsername As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents txtLastName As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents pnlTxtUsername As Panel
    Friend WithEvents txtFirstName As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents lblShowFullName As Label
    Friend WithEvents lblShowAdress As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents lblFirstName As Label
    Friend WithEvents pbProfilePic As PictureBox
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
