<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RemoveOptionControl
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
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.btnRelease = New System.Windows.Forms.Button()
        Me.lblPdlId = New System.Windows.Forms.Label()
        Me.lblPdlName = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnAddInmate = New System.Windows.Forms.Button()
        Me.btnUpdateInmate = New System.Windows.Forms.Button()
        Me.btnRemoveInmate = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 119)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1296, 551)
        Me.TableLayoutPanel1.TabIndex = 68
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel3, 0, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(4, 4)
        Me.TableLayoutPanel2.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 543.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 543.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 543.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(1288, 543)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 3
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24.39025!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.21951!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24.39025!))
        Me.TableLayoutPanel3.Controls.Add(Me.Panel1, 1, 0)
        Me.TableLayoutPanel3.Cursor = System.Windows.Forms.Cursors.Default
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(4, 4)
        Me.TableLayoutPanel3.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(1280, 535)
        Me.TableLayoutPanel3.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.txtID)
        Me.Panel1.Controls.Add(Me.txtName)
        Me.Panel1.Controls.Add(Me.btnRelease)
        Me.Panel1.Controls.Add(Me.lblPdlId)
        Me.Panel1.Controls.Add(Me.lblPdlName)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(316, 4)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(647, 527)
        Me.Panel1.TabIndex = 0
        '
        'txtID
        '
        Me.txtID.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.Location = New System.Drawing.Point(38, 276)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(426, 41)
        Me.txtID.TabIndex = 43
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Location = New System.Drawing.Point(38, 114)
        Me.txtName.Name = "txtName"
        Me.txtName.ReadOnly = True
        Me.txtName.Size = New System.Drawing.Size(426, 41)
        Me.txtName.TabIndex = 42
        '
        'btnRelease
        '
        Me.btnRelease.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRelease.BackColor = System.Drawing.Color.FromArgb(CType(CType(117, Byte), Integer), CType(CType(117, Byte), Integer), CType(CType(117, Byte), Integer))
        Me.btnRelease.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRelease.Font = New System.Drawing.Font("Poppins", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRelease.ForeColor = System.Drawing.Color.White
        Me.btnRelease.Location = New System.Drawing.Point(38, 433)
        Me.btnRelease.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRelease.Name = "btnRelease"
        Me.btnRelease.Size = New System.Drawing.Size(136, 52)
        Me.btnRelease.TabIndex = 41
        Me.btnRelease.Text = "Next"
        Me.btnRelease.UseVisualStyleBackColor = False
        '
        'lblPdlId
        '
        Me.lblPdlId.AutoSize = True
        Me.lblPdlId.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPdlId.ForeColor = System.Drawing.Color.Black
        Me.lblPdlId.Location = New System.Drawing.Point(32, 217)
        Me.lblPdlId.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPdlId.Name = "lblPdlId"
        Me.lblPdlId.Size = New System.Drawing.Size(128, 31)
        Me.lblPdlId.TabIndex = 30
        Me.lblPdlId.Text = "(PDL ID)"
        '
        'lblPdlName
        '
        Me.lblPdlName.AutoSize = True
        Me.lblPdlName.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPdlName.ForeColor = System.Drawing.Color.Black
        Me.lblPdlName.Location = New System.Drawing.Point(32, 60)
        Me.lblPdlName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPdlName.Name = "lblPdlName"
        Me.lblPdlName.Size = New System.Drawing.Size(169, 31)
        Me.lblPdlName.TabIndex = 30
        Me.lblPdlName.Text = "(PDL name)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Teal
        Me.Label3.Location = New System.Drawing.Point(32, 15)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(225, 31)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "PDL Information"
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.btnAddInmate)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnUpdateInmate)
        Me.FlowLayoutPanel2.Controls.Add(Me.btnRemoveInmate)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel2.Margin = New System.Windows.Forms.Padding(4)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(1296, 123)
        Me.FlowLayoutPanel2.TabIndex = 69
        '
        'btnAddInmate
        '
        Me.btnAddInmate.FlatAppearance.BorderSize = 0
        Me.btnAddInmate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddInmate.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddInmate.ForeColor = System.Drawing.Color.Black
        Me.btnAddInmate.Location = New System.Drawing.Point(4, 4)
        Me.btnAddInmate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAddInmate.Name = "btnAddInmate"
        Me.btnAddInmate.Size = New System.Drawing.Size(316, 95)
        Me.btnAddInmate.TabIndex = 24
        Me.btnAddInmate.Text = "Add "
        Me.btnAddInmate.UseVisualStyleBackColor = True
        '
        'btnUpdateInmate
        '
        Me.btnUpdateInmate.FlatAppearance.BorderSize = 0
        Me.btnUpdateInmate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUpdateInmate.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdateInmate.ForeColor = System.Drawing.Color.Black
        Me.btnUpdateInmate.Location = New System.Drawing.Point(328, 4)
        Me.btnUpdateInmate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateInmate.Name = "btnUpdateInmate"
        Me.btnUpdateInmate.Size = New System.Drawing.Size(285, 95)
        Me.btnUpdateInmate.TabIndex = 24
        Me.btnUpdateInmate.Text = "Update"
        Me.btnUpdateInmate.UseVisualStyleBackColor = True
        '
        'btnRemoveInmate
        '
        Me.btnRemoveInmate.FlatAppearance.BorderSize = 0
        Me.btnRemoveInmate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRemoveInmate.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveInmate.ForeColor = System.Drawing.Color.Teal
        Me.btnRemoveInmate.Location = New System.Drawing.Point(621, 4)
        Me.btnRemoveInmate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRemoveInmate.Name = "btnRemoveInmate"
        Me.btnRemoveInmate.Size = New System.Drawing.Size(313, 95)
        Me.btnRemoveInmate.TabIndex = 24
        Me.btnRemoveInmate.Text = "Remove"
        Me.btnRemoveInmate.UseVisualStyleBackColor = True
        '
        'RemoveOptionControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.FlowLayoutPanel2)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "RemoveOptionControl"
        Me.Size = New System.Drawing.Size(1296, 671)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnRelease As Button
    Friend WithEvents lblPdlId As Label
    Friend WithEvents lblPdlName As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents btnAddInmate As Button
    Friend WithEvents btnUpdateInmate As Button
    Friend WithEvents btnRemoveInmate As Button
    Friend WithEvents txtID As TextBox
    Friend WithEvents txtName As TextBox
End Class
