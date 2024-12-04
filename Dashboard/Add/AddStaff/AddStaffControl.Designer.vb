<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AddStaffControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.dgvAllStaff = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnChooseId = New System.Windows.Forms.Button()
        Me.btnChooseImg = New System.Windows.Forms.Button()
        Me.lblID = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblUploadImg = New System.Windows.Forms.Label()
        Me.Panel21 = New System.Windows.Forms.Panel()
        Me.pbID = New System.Windows.Forms.PictureBox()
        Me.Panel22 = New System.Windows.Forms.Panel()
        Me.pbImage = New System.Windows.Forms.PictureBox()
        Me.Panel20 = New System.Windows.Forms.Panel()
        Me.cmbAuthority = New System.Windows.Forms.ComboBox()
        Me.Panel18 = New System.Windows.Forms.Panel()
        Me.cmbPosition = New System.Windows.Forms.ComboBox()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.txtZip = New System.Windows.Forms.TextBox()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.txtRegion = New System.Windows.Forms.TextBox()
        Me.Panel14 = New System.Windows.Forms.Panel()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.Panel15 = New System.Windows.Forms.Panel()
        Me.txtMunicipality = New System.Windows.Forms.TextBox()
        Me.Panel16 = New System.Windows.Forms.Panel()
        Me.txtStreet = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.lblPhone = New System.Windows.Forms.Label()
        Me.lblAdress = New System.Windows.Forms.Label()
        Me.Panel17 = New System.Windows.Forms.Panel()
        Me.dtBirthdate = New System.Windows.Forms.DateTimePicker()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Panel13 = New System.Windows.Forms.Panel()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.Panel12 = New System.Windows.Forms.Panel()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.txtSuffix = New System.Windows.Forms.TextBox()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.lblMiddle = New System.Windows.Forms.Label()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.txtMiddleName = New System.Windows.Forms.TextBox()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.pnlTxtUsername = New System.Windows.Forms.Panel()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.lblCellblockID = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnAddCellblock = New System.Windows.Forms.Button()
        Me.btnAddStaff = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvAllStaff, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel21.SuspendLayout()
        CType(Me.pbID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel22.SuspendLayout()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel20.SuspendLayout()
        Me.Panel18.SuspendLayout()
        Me.Panel11.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel14.SuspendLayout()
        Me.Panel15.SuspendLayout()
        Me.Panel16.SuspendLayout()
        Me.Panel17.SuspendLayout()
        Me.Panel13.SuspendLayout()
        Me.Panel12.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.Panel10.SuspendLayout()
        Me.pnlTxtUsername.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel7, 0, 1)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 100)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 75.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 465.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(969, 445)
        Me.TableLayoutPanel1.TabIndex = 47
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Panel1, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel2, 1, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel2.Margin = New System.Windows.Forms.Padding(3, 3, 100, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(866, 364)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.dgvAllStaff)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(3, 3, 7, 50)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(423, 311)
        Me.Panel1.TabIndex = 0
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Image = Global.CellBlockIM.My.Resources.Resources.plus__1_
        Me.PictureBox2.Location = New System.Drawing.Point(65, 48)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(39, 38)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 19
        Me.PictureBox2.TabStop = False
        '
        'dgvAllStaff
        '
        Me.dgvAllStaff.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAllStaff.BackgroundColor = System.Drawing.Color.White
        Me.dgvAllStaff.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAllStaff.Location = New System.Drawing.Point(14, 159)
        Me.dgvAllStaff.Margin = New System.Windows.Forms.Padding(20, 20, 20, 20)
        Me.dgvAllStaff.Name = "dgvAllStaff"
        Me.dgvAllStaff.RowHeadersWidth = 51
        Me.dgvAllStaff.Size = New System.Drawing.Size(396, 140)
        Me.dgvAllStaff.TabIndex = 22
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Teal
        Me.Label3.Location = New System.Drawing.Point(110, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(109, 25)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Add Staff"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 25)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "All Staff"
        '
        'Panel2
        '
        Me.Panel2.AutoScroll = True
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.Add(Me.btnChooseId)
        Me.Panel2.Controls.Add(Me.btnChooseImg)
        Me.Panel2.Controls.Add(Me.lblID)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.lblUploadImg)
        Me.Panel2.Controls.Add(Me.Panel21)
        Me.Panel2.Controls.Add(Me.Panel22)
        Me.Panel2.Controls.Add(Me.Panel20)
        Me.Panel2.Controls.Add(Me.Panel18)
        Me.Panel2.Controls.Add(Me.Panel11)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.Panel5)
        Me.Panel2.Controls.Add(Me.Panel6)
        Me.Panel2.Controls.Add(Me.Panel14)
        Me.Panel2.Controls.Add(Me.Panel15)
        Me.Panel2.Controls.Add(Me.Panel16)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.lblDate)
        Me.Panel2.Controls.Add(Me.lblPhone)
        Me.Panel2.Controls.Add(Me.lblAdress)
        Me.Panel2.Controls.Add(Me.Panel17)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.Panel13)
        Me.Panel2.Controls.Add(Me.Panel12)
        Me.Panel2.Controls.Add(Me.Panel8)
        Me.Panel2.Controls.Add(Me.lblLastName)
        Me.Panel2.Controls.Add(Me.Panel9)
        Me.Panel2.Controls.Add(Me.lblMiddle)
        Me.Panel2.Controls.Add(Me.Panel10)
        Me.Panel2.Controls.Add(Me.lblFirstName)
        Me.Panel2.Controls.Add(Me.pnlTxtUsername)
        Me.Panel2.Controls.Add(Me.lblCellblockID)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(436, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(427, 358)
        Me.Panel2.TabIndex = 1
        '
        'btnChooseId
        '
        Me.btnChooseId.BackColor = System.Drawing.Color.Gainsboro
        Me.btnChooseId.FlatAppearance.BorderSize = 0
        Me.btnChooseId.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChooseId.Font = New System.Drawing.Font("Poppins", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnChooseId.ForeColor = System.Drawing.Color.Black
        Me.btnChooseId.Location = New System.Drawing.Point(368, 1402)
        Me.btnChooseId.Name = "btnChooseId"
        Me.btnChooseId.Size = New System.Drawing.Size(139, 42)
        Me.btnChooseId.TabIndex = 53
        Me.btnChooseId.Text = "Browse"
        Me.btnChooseId.UseVisualStyleBackColor = False
        '
        'btnChooseImg
        '
        Me.btnChooseImg.BackColor = System.Drawing.Color.Gainsboro
        Me.btnChooseImg.FlatAppearance.BorderSize = 0
        Me.btnChooseImg.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChooseImg.Font = New System.Drawing.Font("Poppins", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnChooseImg.ForeColor = System.Drawing.Color.Black
        Me.btnChooseImg.Location = New System.Drawing.Point(78, 1402)
        Me.btnChooseImg.Name = "btnChooseImg"
        Me.btnChooseImg.Size = New System.Drawing.Size(139, 42)
        Me.btnChooseImg.TabIndex = 54
        Me.btnChooseImg.Text = "Browse"
        Me.btnChooseImg.UseVisualStyleBackColor = False
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblID.Location = New System.Drawing.Point(317, 1168)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(110, 21)
        Me.lblID.TabIndex = 52
        Me.lblID.Text = "Upload a valid ID"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(35, 1118)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(154, 25)
        Me.Label6.TabIndex = 51
        Me.Label6.Text = "Upload Photo"
        '
        'lblUploadImg
        '
        Me.lblUploadImg.AutoSize = True
        Me.lblUploadImg.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUploadImg.Location = New System.Drawing.Point(35, 1168)
        Me.lblUploadImg.Name = "lblUploadImg"
        Me.lblUploadImg.Size = New System.Drawing.Size(94, 21)
        Me.lblUploadImg.TabIndex = 51
        Me.lblUploadImg.Text = "Upload Image"
        '
        'Panel21
        '
        Me.Panel21.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel21.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel21.Controls.Add(Me.pbID)
        Me.Panel21.Location = New System.Drawing.Point(322, 1192)
        Me.Panel21.Name = "Panel21"
        Me.Panel21.Size = New System.Drawing.Size(222, 194)
        Me.Panel21.TabIndex = 49
        '
        'pbID
        '
        Me.pbID.Location = New System.Drawing.Point(8, 7)
        Me.pbID.Name = "pbID"
        Me.pbID.Size = New System.Drawing.Size(204, 177)
        Me.pbID.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbID.TabIndex = 0
        Me.pbID.TabStop = False
        '
        'Panel22
        '
        Me.Panel22.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel22.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel22.Controls.Add(Me.pbImage)
        Me.Panel22.Location = New System.Drawing.Point(39, 1192)
        Me.Panel22.Name = "Panel22"
        Me.Panel22.Size = New System.Drawing.Size(222, 194)
        Me.Panel22.TabIndex = 50
        '
        'pbImage
        '
        Me.pbImage.Location = New System.Drawing.Point(8, 7)
        Me.pbImage.Name = "pbImage"
        Me.pbImage.Size = New System.Drawing.Size(204, 177)
        Me.pbImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbImage.TabIndex = 0
        Me.pbImage.TabStop = False
        '
        'Panel20
        '
        Me.Panel20.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel20.Controls.Add(Me.cmbAuthority)
        Me.Panel20.Location = New System.Drawing.Point(44, 998)
        Me.Panel20.Name = "Panel20"
        Me.Panel20.Size = New System.Drawing.Size(352, 44)
        Me.Panel20.TabIndex = 47
        '
        'cmbAuthority
        '
        Me.cmbAuthority.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.cmbAuthority.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbAuthority.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbAuthority.ForeColor = System.Drawing.Color.DarkGray
        Me.cmbAuthority.FormattingEnabled = True
        Me.cmbAuthority.Items.AddRange(New Object() {"minimum security", "medium security", "maximum security", "juvenile correction", "women's facility", "rehabilitational facility"})
        Me.cmbAuthority.Location = New System.Drawing.Point(4, 3)
        Me.cmbAuthority.Name = "cmbAuthority"
        Me.cmbAuthority.Size = New System.Drawing.Size(344, 34)
        Me.cmbAuthority.TabIndex = 0
        '
        'Panel18
        '
        Me.Panel18.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel18.Controls.Add(Me.cmbPosition)
        Me.Panel18.Location = New System.Drawing.Point(42, 917)
        Me.Panel18.Name = "Panel18"
        Me.Panel18.Size = New System.Drawing.Size(352, 44)
        Me.Panel18.TabIndex = 47
        '
        'cmbPosition
        '
        Me.cmbPosition.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.cmbPosition.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbPosition.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPosition.ForeColor = System.Drawing.Color.DarkGray
        Me.cmbPosition.FormattingEnabled = True
        Me.cmbPosition.Items.AddRange(New Object() {"Correctional Officer", "Disciplinary Officer" & Global.Microsoft.VisualBasic.ChrW(8232), "Medical Officer", "Maintenance Staff", "Cafeteria staff"})
        Me.cmbPosition.Location = New System.Drawing.Point(3, 3)
        Me.cmbPosition.Name = "cmbPosition"
        Me.cmbPosition.Size = New System.Drawing.Size(344, 34)
        Me.cmbPosition.TabIndex = 0
        '
        'Panel11
        '
        Me.Panel11.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel11.Controls.Add(Me.txtPhone)
        Me.Panel11.Location = New System.Drawing.Point(44, 829)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(352, 44)
        Me.Panel11.TabIndex = 47
        '
        'txtPhone
        '
        Me.txtPhone.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtPhone.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtPhone.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhone.ForeColor = System.Drawing.Color.DarkGray
        Me.txtPhone.Location = New System.Drawing.Point(13, 9)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(257, 24)
        Me.txtPhone.TabIndex = 0
        Me.txtPhone.Text = "+63 123-4567-890"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.ComboBox1)
        Me.Panel3.Location = New System.Drawing.Point(232, 746)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(164, 44)
        Me.Panel3.TabIndex = 41
        '
        'ComboBox1
        '
        Me.ComboBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.ComboBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboBox1.Font = New System.Drawing.Font("Poppins", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.ForeColor = System.Drawing.Color.DarkGray
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua & Deps", "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina", "Burundi", "Cambodia", "Cameroon", "Canada", "Cape Verde", "Central African Rep", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo", "Congo {Democratic Rep}", "Costa Rica", "Croatia", "Cuba", "Cyprus", "Czech Republic", "Denmark", "Djibouti", "Dominica", "Dominican Republic", "East Timor", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Honduras", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland {Republic}", "Israel", "Italy", "Ivory Coast", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Korea North", "Korea South", "Kosovo", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Macedonia", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar, {Burma}", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "Norway", "Oman", "Pakistan", "Palau", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania", "Russian Federation", "Rwanda", "St Kitts & Nevis", "St Lucia", "Saint Vincent & the Grenadines", "Samoa", "San Marino", "Sao Tome & Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname", "Swaziland", "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand", "Togo", "Tonga", "Trinidad & Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States", "Uruguay", "Uzbekistan", "Vanuatu", "Vatican City", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe"})
        Me.ComboBox1.Location = New System.Drawing.Point(5, 5)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(154, 34)
        Me.ComboBox1.TabIndex = 0
        Me.ComboBox1.Text = "Country"
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.txtZip)
        Me.Panel5.Location = New System.Drawing.Point(44, 746)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(182, 44)
        Me.Panel5.TabIndex = 42
        '
        'txtZip
        '
        Me.txtZip.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtZip.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtZip.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtZip.ForeColor = System.Drawing.Color.DarkGray
        Me.txtZip.Location = New System.Drawing.Point(15, 9)
        Me.txtZip.Name = "txtZip"
        Me.txtZip.Size = New System.Drawing.Size(119, 24)
        Me.txtZip.TabIndex = 0
        Me.txtZip.Text = "ZIP Code"
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Controls.Add(Me.txtRegion)
        Me.Panel6.Location = New System.Drawing.Point(232, 696)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(164, 44)
        Me.Panel6.TabIndex = 43
        '
        'txtRegion
        '
        Me.txtRegion.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtRegion.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtRegion.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegion.ForeColor = System.Drawing.Color.DarkGray
        Me.txtRegion.Location = New System.Drawing.Point(15, 9)
        Me.txtRegion.Name = "txtRegion"
        Me.txtRegion.Size = New System.Drawing.Size(111, 24)
        Me.txtRegion.TabIndex = 0
        Me.txtRegion.Text = "Region"
        '
        'Panel14
        '
        Me.Panel14.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel14.Controls.Add(Me.txtCity)
        Me.Panel14.Location = New System.Drawing.Point(44, 696)
        Me.Panel14.Name = "Panel14"
        Me.Panel14.Size = New System.Drawing.Size(182, 44)
        Me.Panel14.TabIndex = 44
        '
        'txtCity
        '
        Me.txtCity.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtCity.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtCity.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCity.ForeColor = System.Drawing.Color.DarkGray
        Me.txtCity.Location = New System.Drawing.Point(15, 9)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(119, 24)
        Me.txtCity.TabIndex = 0
        Me.txtCity.Text = "City/District"
        '
        'Panel15
        '
        Me.Panel15.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel15.Controls.Add(Me.txtMunicipality)
        Me.Panel15.Location = New System.Drawing.Point(44, 646)
        Me.Panel15.Name = "Panel15"
        Me.Panel15.Size = New System.Drawing.Size(353, 44)
        Me.Panel15.TabIndex = 45
        '
        'txtMunicipality
        '
        Me.txtMunicipality.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtMunicipality.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtMunicipality.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMunicipality.ForeColor = System.Drawing.Color.DarkGray
        Me.txtMunicipality.Location = New System.Drawing.Point(15, 9)
        Me.txtMunicipality.Name = "txtMunicipality"
        Me.txtMunicipality.Size = New System.Drawing.Size(333, 24)
        Me.txtMunicipality.TabIndex = 0
        Me.txtMunicipality.Text = "Municipality"
        '
        'Panel16
        '
        Me.Panel16.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel16.Controls.Add(Me.txtStreet)
        Me.Panel16.Location = New System.Drawing.Point(44, 596)
        Me.Panel16.Name = "Panel16"
        Me.Panel16.Size = New System.Drawing.Size(353, 44)
        Me.Panel16.TabIndex = 46
        '
        'txtStreet
        '
        Me.txtStreet.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtStreet.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStreet.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStreet.ForeColor = System.Drawing.Color.DarkGray
        Me.txtStreet.Location = New System.Drawing.Point(15, 9)
        Me.txtStreet.Name = "txtStreet"
        Me.txtStreet.Size = New System.Drawing.Size(333, 24)
        Me.txtStreet.TabIndex = 0
        Me.txtStreet.Text = "Street"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(42, 974)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(121, 21)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Authorization Level"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(40, 893)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 21)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Position"
        '
        'lblDate
        '
        Me.lblDate.AutoSize = True
        Me.lblDate.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate.Location = New System.Drawing.Point(40, 487)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(82, 21)
        Me.lblDate.TabIndex = 38
        Me.lblDate.Text = "Date of birth"
        '
        'lblPhone
        '
        Me.lblPhone.AutoSize = True
        Me.lblPhone.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhone.Location = New System.Drawing.Point(42, 805)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(98, 21)
        Me.lblPhone.TabIndex = 39
        Me.lblPhone.Text = "Phone number"
        '
        'lblAdress
        '
        Me.lblAdress.AutoSize = True
        Me.lblAdress.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAdress.Location = New System.Drawing.Point(42, 572)
        Me.lblAdress.Name = "lblAdress"
        Me.lblAdress.Size = New System.Drawing.Size(57, 21)
        Me.lblAdress.TabIndex = 40
        Me.lblAdress.Text = "Address"
        '
        'Panel17
        '
        Me.Panel17.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel17.Controls.Add(Me.dtBirthdate)
        Me.Panel17.Location = New System.Drawing.Point(44, 511)
        Me.Panel17.Name = "Panel17"
        Me.Panel17.Size = New System.Drawing.Size(353, 44)
        Me.Panel17.TabIndex = 37
        '
        'dtBirthdate
        '
        Me.dtBirthdate.Font = New System.Drawing.Font("Poppins", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtBirthdate.Location = New System.Drawing.Point(26, 7)
        Me.dtBirthdate.Name = "dtBirthdate"
        Me.dtBirthdate.Size = New System.Drawing.Size(300, 27)
        Me.dtBirthdate.TabIndex = 24
        Me.dtBirthdate.Value = New Date(2024, 10, 14, 11, 6, 41, 0)
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(40, 379)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(30, 21)
        Me.Label8.TabIndex = 32
        Me.Label8.Text = "Sex"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(245, 379)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(74, 21)
        Me.Label9.TabIndex = 33
        Me.Label9.Text = "Civil status"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Gray
        Me.Label10.Location = New System.Drawing.Point(76, 302)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(70, 21)
        Me.Label10.TabIndex = 34
        Me.Label10.Text = "(optional)"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(40, 302)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(42, 21)
        Me.Label11.TabIndex = 35
        Me.Label11.Text = "Suffix"
        '
        'Panel13
        '
        Me.Panel13.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel13.Controls.Add(Me.cmbStatus)
        Me.Panel13.Location = New System.Drawing.Point(242, 403)
        Me.Panel13.Name = "Panel13"
        Me.Panel13.Size = New System.Drawing.Size(129, 44)
        Me.Panel13.TabIndex = 28
        '
        'cmbStatus
        '
        Me.cmbStatus.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.cmbStatus.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbStatus.Font = New System.Drawing.Font("Poppins", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbStatus.ForeColor = System.Drawing.Color.Gray
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Items.AddRange(New Object() {"Single", "Married ", "Widowed", "Separated", "Divorced"})
        Me.cmbStatus.Location = New System.Drawing.Point(3, 6)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(121, 31)
        Me.cmbStatus.TabIndex = 25
        Me.cmbStatus.Text = "Single"
        '
        'Panel12
        '
        Me.Panel12.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel12.Controls.Add(Me.RadioButton1)
        Me.Panel12.Controls.Add(Me.RadioButton2)
        Me.Panel12.Location = New System.Drawing.Point(45, 403)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(129, 63)
        Me.Panel12.TabIndex = 29
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Font = New System.Drawing.Font("Poppins", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(17, 5)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(59, 27)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Male"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Font = New System.Drawing.Font("Poppins", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(17, 29)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(76, 27)
        Me.RadioButton2.TabIndex = 0
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Female"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Panel8
        '
        Me.Panel8.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel8.Controls.Add(Me.txtSuffix)
        Me.Panel8.Location = New System.Drawing.Point(44, 326)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(129, 44)
        Me.Panel8.TabIndex = 30
        '
        'txtSuffix
        '
        Me.txtSuffix.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtSuffix.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtSuffix.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSuffix.ForeColor = System.Drawing.Color.DarkGray
        Me.txtSuffix.Location = New System.Drawing.Point(13, 9)
        Me.txtSuffix.Name = "txtSuffix"
        Me.txtSuffix.Size = New System.Drawing.Size(102, 24)
        Me.txtSuffix.TabIndex = 0
        Me.txtSuffix.Text = "--"
        '
        'lblLastName
        '
        Me.lblLastName.AutoSize = True
        Me.lblLastName.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastName.Location = New System.Drawing.Point(40, 222)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(71, 21)
        Me.lblLastName.TabIndex = 36
        Me.lblLastName.Text = "Last Name"
        '
        'Panel9
        '
        Me.Panel9.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel9.Controls.Add(Me.txtLastName)
        Me.Panel9.Location = New System.Drawing.Point(44, 246)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(353, 44)
        Me.Panel9.TabIndex = 31
        '
        'txtLastName
        '
        Me.txtLastName.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtLastName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtLastName.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastName.ForeColor = System.Drawing.Color.DarkGray
        Me.txtLastName.Location = New System.Drawing.Point(13, 9)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(335, 24)
        Me.txtLastName.TabIndex = 0
        Me.txtLastName.Text = "Enter Last Name"
        '
        'lblMiddle
        '
        Me.lblMiddle.AutoSize = True
        Me.lblMiddle.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMiddle.Location = New System.Drawing.Point(42, 142)
        Me.lblMiddle.Name = "lblMiddle"
        Me.lblMiddle.Size = New System.Drawing.Size(87, 21)
        Me.lblMiddle.TabIndex = 27
        Me.lblMiddle.Text = "Middle Name"
        '
        'Panel10
        '
        Me.Panel10.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel10.Controls.Add(Me.txtMiddleName)
        Me.Panel10.Location = New System.Drawing.Point(45, 166)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(352, 44)
        Me.Panel10.TabIndex = 26
        '
        'txtMiddleName
        '
        Me.txtMiddleName.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtMiddleName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtMiddleName.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMiddleName.ForeColor = System.Drawing.Color.DarkGray
        Me.txtMiddleName.Location = New System.Drawing.Point(12, 9)
        Me.txtMiddleName.Name = "txtMiddleName"
        Me.txtMiddleName.Size = New System.Drawing.Size(335, 24)
        Me.txtMiddleName.TabIndex = 0
        Me.txtMiddleName.Text = "Enter Middle Name"
        '
        'lblFirstName
        '
        Me.lblFirstName.AutoSize = True
        Me.lblFirstName.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFirstName.Location = New System.Drawing.Point(41, 61)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(71, 21)
        Me.lblFirstName.TabIndex = 25
        Me.lblFirstName.Text = "First Name"
        '
        'pnlTxtUsername
        '
        Me.pnlTxtUsername.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.pnlTxtUsername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlTxtUsername.Controls.Add(Me.txtFirstName)
        Me.pnlTxtUsername.Location = New System.Drawing.Point(45, 85)
        Me.pnlTxtUsername.Name = "pnlTxtUsername"
        Me.pnlTxtUsername.Size = New System.Drawing.Size(352, 44)
        Me.pnlTxtUsername.TabIndex = 24
        '
        'txtFirstName
        '
        Me.txtFirstName.BackColor = System.Drawing.Color.FromArgb(CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.txtFirstName.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtFirstName.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstName.ForeColor = System.Drawing.Color.DarkGray
        Me.txtFirstName.Location = New System.Drawing.Point(12, 9)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(335, 24)
        Me.txtFirstName.TabIndex = 0
        Me.txtFirstName.Text = "Enter First Name"
        '
        'lblCellblockID
        '
        Me.lblCellblockID.AutoSize = True
        Me.lblCellblockID.Font = New System.Drawing.Font("Poppins", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCellblockID.Location = New System.Drawing.Point(235, 60)
        Me.lblCellblockID.Name = "lblCellblockID"
        Me.lblCellblockID.Size = New System.Drawing.Size(162, 21)
        Me.lblCellblockID.TabIndex = 23
        Me.lblCellblockID.Text = "Staff ID: (auto generated) "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(18, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(185, 25)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Staff Information"
        '
        'Panel7
        '
        Me.Panel7.Controls.Add(Me.btnSave)
        Me.Panel7.Controls.Add(Me.btnCancel)
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel7.Location = New System.Drawing.Point(3, 373)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(963, 69)
        Me.Panel7.TabIndex = 1
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.FromArgb(CType(CType(117, Byte), Integer), CType(CType(117, Byte), Integer), CType(CType(117, Byte), Integer))
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSave.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.Color.White
        Me.btnSave.Location = New System.Drawing.Point(467, 17)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(201, 49)
        Me.btnSave.TabIndex = 34
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.White
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancel.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.Color.Black
        Me.btnCancel.Location = New System.Drawing.Point(260, 17)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(201, 49)
        Me.btnCancel.TabIndex = 34
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.btnAddCellblock)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnAddStaff)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(972, 100)
        Me.FlowLayoutPanel2.TabIndex = 48
        '
        'btnAddCellblock
        '
        Me.btnAddCellblock.FlatAppearance.BorderSize = 0
        Me.btnAddCellblock.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddCellblock.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddCellblock.ForeColor = System.Drawing.Color.Black
        Me.btnAddCellblock.Location = New System.Drawing.Point(3, 3)
        Me.btnAddCellblock.Name = "btnAddCellblock"
        Me.btnAddCellblock.Size = New System.Drawing.Size(296, 77)
        Me.btnAddCellblock.TabIndex = 24
        Me.btnAddCellblock.Text = "Containment"
        Me.btnAddCellblock.UseVisualStyleBackColor = True
        '
        'btnAddStaff
        '
        Me.btnAddStaff.FlatAppearance.BorderSize = 0
        Me.btnAddStaff.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddStaff.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddStaff.ForeColor = System.Drawing.Color.Teal
        Me.btnAddStaff.Location = New System.Drawing.Point(305, 3)
        Me.btnAddStaff.Name = "btnAddStaff"
        Me.btnAddStaff.Size = New System.Drawing.Size(235, 77)
        Me.btnAddStaff.TabIndex = 24
        Me.btnAddStaff.Text = "Add Staff"
        Me.btnAddStaff.UseVisualStyleBackColor = True
        '
        'AddStaffControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.FlowLayoutPanel2)
        Me.Name = "AddStaffControl"
        Me.Size = New System.Drawing.Size(972, 545)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvAllStaff, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel21.ResumeLayout(False)
        CType(Me.pbID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel22.ResumeLayout(False)
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel20.ResumeLayout(False)
        Me.Panel18.ResumeLayout(False)
        Me.Panel11.ResumeLayout(False)
        Me.Panel11.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.Panel14.ResumeLayout(False)
        Me.Panel14.PerformLayout()
        Me.Panel15.ResumeLayout(False)
        Me.Panel15.PerformLayout()
        Me.Panel16.ResumeLayout(False)
        Me.Panel16.PerformLayout()
        Me.Panel17.ResumeLayout(False)
        Me.Panel13.ResumeLayout(False)
        Me.Panel12.ResumeLayout(False)
        Me.Panel12.PerformLayout()
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        Me.Panel9.ResumeLayout(False)
        Me.Panel9.PerformLayout()
        Me.Panel10.ResumeLayout(False)
        Me.Panel10.PerformLayout()
        Me.pnlTxtUsername.ResumeLayout(False)
        Me.pnlTxtUsername.PerformLayout()
        Me.Panel7.ResumeLayout(False)
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents dgvAllStaff As DataGridView
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnChooseId As Button
    Friend WithEvents btnChooseImg As Button
    Friend WithEvents lblID As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents lblUploadImg As Label
    Friend WithEvents Panel21 As Panel
    Friend WithEvents pbID As PictureBox
    Friend WithEvents Panel22 As Panel
    Friend WithEvents pbImage As PictureBox
    Friend WithEvents Panel20 As Panel
    Friend WithEvents cmbAuthority As ComboBox
    Friend WithEvents Panel18 As Panel
    Friend WithEvents cmbPosition As ComboBox
    Friend WithEvents Panel11 As Panel
    Friend WithEvents txtPhone As TextBox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Panel5 As Panel
    Friend WithEvents txtZip As TextBox
    Friend WithEvents Panel6 As Panel
    Friend WithEvents txtRegion As TextBox
    Friend WithEvents Panel14 As Panel
    Friend WithEvents txtCity As TextBox
    Friend WithEvents Panel15 As Panel
    Friend WithEvents txtMunicipality As TextBox
    Friend WithEvents Panel16 As Panel
    Friend WithEvents txtStreet As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents lblDate As Label
    Friend WithEvents lblPhone As Label
    Friend WithEvents lblAdress As Label
    Friend WithEvents Panel17 As Panel
    Friend WithEvents dtBirthdate As DateTimePicker
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Panel13 As Panel
    Friend WithEvents cmbStatus As ComboBox
    Friend WithEvents Panel12 As Panel
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents Panel8 As Panel
    Friend WithEvents txtSuffix As TextBox
    Friend WithEvents lblLastName As Label
    Friend WithEvents Panel9 As Panel
    Friend WithEvents txtLastName As TextBox
    Friend WithEvents lblMiddle As Label
    Friend WithEvents Panel10 As Panel
    Friend WithEvents txtMiddleName As TextBox
    Friend WithEvents lblFirstName As Label
    Friend WithEvents pnlTxtUsername As Panel
    Friend WithEvents txtFirstName As TextBox
    Friend WithEvents lblCellblockID As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Panel7 As Panel
    Friend WithEvents btnSave As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents btnAddCellblock As Button
    Friend WithEvents btnAddStaff As Button
End Class
