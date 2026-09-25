<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExpenses
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
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbCurrencyType = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpExpenseDate = New System.Windows.Forms.DateTimePicker()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbCategory = New System.Windows.Forms.ComboBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtVoucherNumber = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ExpensesReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintSelectToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.New1ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BulkExpenseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileToolStripMenuItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblTotalAmount = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ExportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.GroupBox1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(77, 69)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(270, 25)
        Me.txtDescription.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(361, 188)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 19)
        Me.Label1.TabIndex = 45
        Me.Label1.Text = "فئة المصروف"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbCurrencyType
        '
        Me.cmbCurrencyType.FormattingEnabled = True
        Me.cmbCurrencyType.Items.AddRange(New Object() {"شيكل", "دينار", "دولار"})
        Me.cmbCurrencyType.Location = New System.Drawing.Point(106, 105)
        Me.cmbCurrencyType.Name = "cmbCurrencyType"
        Me.cmbCurrencyType.Size = New System.Drawing.Size(82, 28)
        Me.cmbCurrencyType.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label4.Enabled = False
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(197, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 19)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "نوع العملة"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtpExpenseDate
        '
        Me.dtpExpenseDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpExpenseDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpExpenseDate.Location = New System.Drawing.Point(172, 144)
        Me.dtpExpenseDate.Name = "dtpExpenseDate"
        Me.dtpExpenseDate.Size = New System.Drawing.Size(174, 25)
        Me.dtpExpenseDate.TabIndex = 5
        '
        'Label16
        '
        Me.Label16.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label16.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label16.Enabled = False
        Me.Label16.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(361, 150)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(67, 19)
        Me.Label16.TabIndex = 36
        Me.Label16.Text = "تاريخ الصرف"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label3.Enabled = False
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(361, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 19)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "قيمة المبلغ"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(292, 106)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(57, 25)
        Me.txtAmount.TabIndex = 3
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(77, 223)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(272, 56)
        Me.txtNotes.TabIndex = 7
        Me.txtNotes.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label6.Enabled = False
        Me.Label6.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(361, 238)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 19)
        Me.Label6.TabIndex = 32
        Me.Label6.Text = "ملاحظات"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbCategory
        '
        Me.cmbCategory.FormattingEnabled = True
        Me.cmbCategory.Location = New System.Drawing.Point(264, 185)
        Me.cmbCategory.Name = "cmbCategory"
        Me.cmbCategory.Size = New System.Drawing.Size(82, 28)
        Me.cmbCategory.TabIndex = 6
        '
        'Label17
        '
        Me.Label17.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label17.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label17.Enabled = False
        Me.Label17.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(361, 76)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 19)
        Me.Label17.TabIndex = 12
        Me.Label17.Text = "البيان"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtVoucherNumber
        '
        Me.txtVoucherNumber.Location = New System.Drawing.Point(270, 36)
        Me.txtVoucherNumber.Name = "txtVoucherNumber"
        Me.txtVoucherNumber.Size = New System.Drawing.Size(81, 25)
        Me.txtVoucherNumber.TabIndex = 1
        '
        'Label18
        '
        Me.Label18.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label18.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label18.Enabled = False
        Me.Label18.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(361, 39)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label18.Size = New System.Drawing.Size(67, 19)
        Me.Label18.TabIndex = 10
        Me.Label18.Text = "رقم السند"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.txtDescription)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cmbCurrencyType)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.dtpExpenseDate)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtAmount)
        Me.GroupBox1.Controls.Add(Me.txtNotes)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.cmbCategory)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.txtVoucherNumber)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial Narrow", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(346, 47)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.GroupBox1.Size = New System.Drawing.Size(440, 308)
        Me.GroupBox1.TabIndex = 45
        Me.GroupBox1.TabStop = False
        '
        'ExpensesReportToolStripMenuItem
        '
        Me.ExpensesReportToolStripMenuItem.Name = "ExpensesReportToolStripMenuItem"
        Me.ExpensesReportToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.ExpensesReportToolStripMenuItem.Text = "تقرير المصروفات"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExpensesReportToolStripMenuItem, Me.PrintSelectToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.ReportToolStripMenuItem.Text = "تقرير"
        '
        'PrintSelectToolStripMenuItem
        '
        Me.PrintSelectToolStripMenuItem.Name = "PrintSelectToolStripMenuItem"
        Me.PrintSelectToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.PrintSelectToolStripMenuItem.Text = "طباعة محدد"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.ExitToolStripMenuItem.Text = "خروج"
        '
        'DeleteToolStripMenuItem
        '
        Me.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem"
        Me.DeleteToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.DeleteToolStripMenuItem.Text = "حذف"
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.SaveToolStripMenuItem.Text = "حفظ"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.New1ToolStripMenuItem, Me.BulkExpenseToolStripMenuItem})
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.NewToolStripMenuItem.Text = "جديد"
        '
        'New1ToolStripMenuItem
        '
        Me.New1ToolStripMenuItem.Name = "New1ToolStripMenuItem"
        Me.New1ToolStripMenuItem.Size = New System.Drawing.Size(218, 22)
        Me.New1ToolStripMenuItem.Text = "سند صرف جديد لشخص"
        '
        'BulkExpenseToolStripMenuItem
        '
        Me.BulkExpenseToolStripMenuItem.Name = "BulkExpenseToolStripMenuItem"
        Me.BulkExpenseToolStripMenuItem.Size = New System.Drawing.Size(218, 22)
        Me.BulkExpenseToolStripMenuItem.Text = "سند صرف جديد لعدة اشخاص"
        '
        'FileToolStripMenuItemToolStripMenuItem
        '
        Me.FileToolStripMenuItemToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.SaveToolStripMenuItem, Me.DeleteToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItemToolStripMenuItem.Name = "FileToolStripMenuItemToolStripMenuItem"
        Me.FileToolStripMenuItemToolStripMenuItem.Size = New System.Drawing.Size(42, 20)
        Me.FileToolStripMenuItemToolStripMenuItem.Text = "ملف"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(112, 362)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(100, 23)
        Me.ProgressBar1.TabIndex = 43
        '
        'lblTotalAmount
        '
        Me.lblTotalAmount.AutoSize = True
        Me.lblTotalAmount.Enabled = False
        Me.lblTotalAmount.Location = New System.Drawing.Point(11, 372)
        Me.lblTotalAmount.Name = "lblTotalAmount"
        Me.lblTotalAmount.Size = New System.Drawing.Size(38, 13)
        Me.lblTotalAmount.TabIndex = 42
        Me.lblTotalAmount.Text = "Label1"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItemToolStripMenuItem, Me.ReportToolStripMenuItem, Me.ExportToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.MenuStrip1.Size = New System.Drawing.Size(795, 24)
        Me.MenuStrip1.TabIndex = 44
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ExportToolStripMenuItem
        '
        Me.ExportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExportToExcelToolStripMenuItem})
        Me.ExportToolStripMenuItem.Name = "ExportToolStripMenuItem"
        Me.ExportToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.ExportToolStripMenuItem.Text = "تصدير"
        '
        'ExportToExcelToolStripMenuItem
        '
        Me.ExportToExcelToolStripMenuItem.Name = "ExportToExcelToolStripMenuItem"
        Me.ExportToExcelToolStripMenuItem.Size = New System.Drawing.Size(155, 22)
        Me.ExportToExcelToolStripMenuItem.Text = "تصدير الى اكسل"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DataGridView1.Location = New System.Drawing.Point(0, 393)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DataGridView1.Size = New System.Drawing.Size(795, 150)
        Me.DataGridView1.TabIndex = 41
        '
        'frmExpenses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(795, 543)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.lblTotalAmount)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.DataGridView1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmExpenses"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "شاشة المصروفات"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtDescription As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbCurrencyType As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents dtpExpenseDate As DateTimePicker
    Friend WithEvents Label16 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtAmount As TextBox
    Friend WithEvents txtNotes As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents cmbCategory As ComboBox
    Friend WithEvents Label17 As Label
    Friend WithEvents txtVoucherNumber As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents ExpensesReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintSelectToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DeleteToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItemToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents lblTotalAmount As Label
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ExportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExportToExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents New1ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents BulkExpenseToolStripMenuItem As ToolStripMenuItem
End Class
