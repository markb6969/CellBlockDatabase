<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class InmateHomeControl
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
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.fpnCenter = New System.Windows.Forms.FlowLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbFilter = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.btnViewCase = New System.Windows.Forms.Button()
        Me.btnViewPDL = New System.Windows.Forms.Button()
        Me.btnViewMed = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnAddInmate = New System.Windows.Forms.Button()
        Me.btnUpdateInmate = New System.Windows.Forms.Button()
        Me.btnRemoveInmate = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.fpnCenter.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel3, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 95)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 450.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(972, 450)
        Me.TableLayoutPanel1.TabIndex = 47
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 3
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.57469!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 74.53416!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 7.919255!))
        Me.TableLayoutPanel3.Controls.Add(Me.fpnCenter, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Panel2, 0, 0)
        Me.TableLayoutPanel3.Cursor = System.Windows.Forms.Cursors.Default
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(966, 444)
        Me.TableLayoutPanel3.TabIndex = 1
        '
        'fpnCenter
        '
        Me.fpnCenter.Controls.Add(Me.Panel1)
        Me.fpnCenter.Controls.Add(Me.DataGridView1)
        Me.fpnCenter.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.fpnCenter.Location = New System.Drawing.Point(171, 2)
        Me.fpnCenter.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.fpnCenter.Name = "fpnCenter"
        Me.fpnCenter.Size = New System.Drawing.Size(715, 439)
        Me.fpnCenter.TabIndex = 2
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbFilter)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.RichTextBox1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(2, 2)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(712, 59)
        Me.Panel1.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(21, 21)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 28)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "filter"
        '
        'cmbFilter
        '
        Me.cmbFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbFilter.FormattingEnabled = True
        Me.cmbFilter.Items.AddRange(New Object() {"A-Z", "Date(oldest)", "Date(Newest)", "ID"})
        Me.cmbFilter.Location = New System.Drawing.Point(71, 20)
        Me.cmbFilter.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmbFilter.Name = "cmbFilter"
        Me.cmbFilter.Size = New System.Drawing.Size(92, 28)
        Me.cmbFilter.TabIndex = 2
        Me.cmbFilter.Text = "ID"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(333, 18)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 28)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Search"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.RichTextBox1.Location = New System.Drawing.Point(401, 18)
        Me.RichTextBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(264, 27)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        '
        'DataGridView1
        '
        Me.DataGridView1.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Raised
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(2, 65)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(712, 363)
        Me.DataGridView1.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtID)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.PictureBox1)
        Me.Panel2.Controls.Add(Me.btnViewCase)
        Me.Panel2.Controls.Add(Me.btnViewPDL)
        Me.Panel2.Controls.Add(Me.btnViewMed)
        Me.Panel2.Location = New System.Drawing.Point(2, 2)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(165, 439)
        Me.Panel2.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 179)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 28)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Selected ID:"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.CellBlockIM.My.Resources.Resources.admin_with_cogwheels1
        Me.PictureBox1.Location = New System.Drawing.Point(2, 2)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(163, 169)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'txtID
        '
        Me.txtID.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.Location = New System.Drawing.Point(16, 202)
        Me.txtID.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(133, 26)
        Me.txtID.TabIndex = 0
        '
        'btnViewCase
        '
        Me.btnViewCase.FlatAppearance.BorderSize = 0
        Me.btnViewCase.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnViewCase.Font = New System.Drawing.Font("Poppins", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewCase.Location = New System.Drawing.Point(2, 345)
        Me.btnViewCase.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnViewCase.Name = "btnViewCase"
        Me.btnViewCase.Size = New System.Drawing.Size(163, 57)
        Me.btnViewCase.TabIndex = 3
        Me.btnViewCase.Text = "View case"
        Me.btnViewCase.UseVisualStyleBackColor = True
        '
        'btnViewPDL
        '
        Me.btnViewPDL.FlatAppearance.BorderSize = 0
        Me.btnViewPDL.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnViewPDL.Font = New System.Drawing.Font("Poppins", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewPDL.Location = New System.Drawing.Point(2, 231)
        Me.btnViewPDL.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnViewPDL.Name = "btnViewPDL"
        Me.btnViewPDL.Size = New System.Drawing.Size(163, 54)
        Me.btnViewPDL.TabIndex = 1
        Me.btnViewPDL.Text = "View pdl Info"
        Me.btnViewPDL.UseVisualStyleBackColor = True
        '
        'btnViewMed
        '
        Me.btnViewMed.FlatAppearance.BorderSize = 0
        Me.btnViewMed.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnViewMed.Font = New System.Drawing.Font("Poppins", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewMed.Location = New System.Drawing.Point(2, 289)
        Me.btnViewMed.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnViewMed.Name = "btnViewMed"
        Me.btnViewMed.Size = New System.Drawing.Size(163, 51)
        Me.btnViewMed.TabIndex = 2
        Me.btnViewMed.Text = "View medical"
        Me.btnViewMed.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.btnAddInmate)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnUpdateInmate)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnRemoveInmate)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(972, 100)
        Me.FlowLayoutPanel2.TabIndex = 49
        '
        'btnAddInmate
        '
        Me.btnAddInmate.FlatAppearance.BorderSize = 0
        Me.btnAddInmate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddInmate.Font = New System.Drawing.Font("Poppins SemiBold", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddInmate.ForeColor = System.Drawing.Color.Black
        Me.btnAddInmate.Location = New System.Drawing.Point(3, 3)
        Me.btnAddInmate.Name = "btnAddInmate"
        Me.btnAddInmate.Size = New System.Drawing.Size(237, 77)
        Me.btnAddInmate.TabIndex = 24
        Me.btnAddInmate.Text = "Add "
        Me.btnAddInmate.UseVisualStyleBackColor = True
        '
        'btnUpdateInmate
        '
        Me.btnUpdateInmate.FlatAppearance.BorderSize = 0
        Me.btnUpdateInmate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUpdateInmate.Font = New System.Drawing.Font("Poppins SemiBold", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdateInmate.Location = New System.Drawing.Point(246, 3)
        Me.btnUpdateInmate.Name = "btnUpdateInmate"
        Me.btnUpdateInmate.Size = New System.Drawing.Size(214, 77)
        Me.btnUpdateInmate.TabIndex = 24
        Me.btnUpdateInmate.Text = "Update"
        Me.btnUpdateInmate.UseVisualStyleBackColor = True
        '
        'btnRemoveInmate
        '
        Me.btnRemoveInmate.FlatAppearance.BorderSize = 0
        Me.btnRemoveInmate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRemoveInmate.Font = New System.Drawing.Font("Poppins SemiBold", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveInmate.Location = New System.Drawing.Point(466, 3)
        Me.btnRemoveInmate.Name = "btnRemoveInmate"
        Me.btnRemoveInmate.Size = New System.Drawing.Size(235, 77)
        Me.btnRemoveInmate.TabIndex = 24
        Me.btnRemoveInmate.Text = "Release"
        Me.btnRemoveInmate.UseVisualStyleBackColor = True
        '
        'InmateHomeControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.FlowLayoutPanel2)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "InmateHomeControl"
        Me.Size = New System.Drawing.Size(972, 545)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.fpnCenter.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents btnAddInmate As Button
    Friend WithEvents btnUpdateInmate As Button
    Friend WithEvents btnRemoveInmate As Button
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents fpnCenter As FlowLayoutPanel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbFilter As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Panel2 As Panel
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnViewCase As Button
    Friend WithEvents btnViewPDL As Button
    Friend WithEvents btnViewMed As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents txtID As TextBox
End Class
