<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPayments
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.lblTotalAmount = New System.Windows.Forms.Label()
        Me.PaymentReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtMobileNumber = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbCurrencyType = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpDueDate = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PrintReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GeneralReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmbPlayerName = New System.Windows.Forms.ComboBox()
        Me.dtpPaymentDate = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblSearchResult = New System.Windows.Forms.Label()
        Me.ExporteToExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbPaymentMethod = New System.Windows.Forms.ComboBox()
        Me.ExporteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblSummary = New System.Windows.Forms.Label()
        Me.txtTransferorName = New System.Windows.Forms.TextBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.New1ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BulkPaymentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ImageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddimageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteimageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintimageToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtVoucherNumber = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTotalAmount
        '
        Me.lblTotalAmount.AutoSize = True
        Me.lblTotalAmount.Enabled = False
        Me.lblTotalAmount.Location = New System.Drawing.Point(709, 495)
        Me.lblTotalAmount.Name = "lblTotalAmount"
        Me.lblTotalAmount.Size = New System.Drawing.Size(48, 13)
        Me.lblTotalAmount.TabIndex = 50
        Me.lblTotalAmount.Text = "الإجمالي"
        Me.lblTotalAmount.Visible = False
        '
        'PaymentReportToolStripMenuItem
        '
        Me.PaymentReportToolStripMenuItem.Name = "PaymentReportToolStripMenuItem"
        Me.PaymentReportToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.PaymentReportToolStripMenuItem.Text = "تقرير الايرادات"
        '
        'txtMobileNumber
        '
        Me.txtMobileNumber.Location = New System.Drawing.Point(355, 153)
        Me.txtMobileNumber.Name = "txtMobileNumber"
        Me.txtMobileNumber.Size = New System.Drawing.Size(97, 20)
        Me.txtMobileNumber.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label5.Enabled = False
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(468, 155)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 17)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "رقم الجوال"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbCurrencyType
        '
        Me.cmbCurrencyType.FormattingEnabled = True
        Me.cmbCurrencyType.Items.AddRange(New Object() {"شيكل", "دينار", "دولار"})
        Me.cmbCurrencyType.Location = New System.Drawing.Point(92, 154)
        Me.cmbCurrencyType.Name = "cmbCurrencyType"
        Me.cmbCurrencyType.Size = New System.Drawing.Size(82, 21)
        Me.cmbCurrencyType.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label4.Enabled = False
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(200, 158)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 17)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "نوع العملة"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtpDueDate
        '
        Me.dtpDueDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDueDate.Location = New System.Drawing.Point(281, 270)
        Me.dtpDueDate.Name = "dtpDueDate"
        Me.dtpDueDate.Size = New System.Drawing.Size(174, 20)
        Me.dtpDueDate.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label2.Enabled = False
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(468, 272)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 17)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "تاريخ الاستحقاق"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PrintReportToolStripMenuItem
        '
        Me.PrintReportToolStripMenuItem.Name = "PrintReportToolStripMenuItem"
        Me.PrintReportToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.PrintReportToolStripMenuItem.Text = "طباعة التقرير"
        '
        'GeneralReportToolStripMenuItem
        '
        Me.GeneralReportToolStripMenuItem.Name = "GeneralReportToolStripMenuItem"
        Me.GeneralReportToolStripMenuItem.Size = New System.Drawing.Size(143, 22)
        Me.GeneralReportToolStripMenuItem.Text = "تقرير عام"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GeneralReportToolStripMenuItem, Me.PaymentReportToolStripMenuItem, Me.PrintReportToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.ReportToolStripMenuItem.Text = "تقرير"
        '
        'cmbPlayerName
        '
        Me.cmbPlayerName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbPlayerName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbPlayerName.FormattingEnabled = True
        Me.cmbPlayerName.Location = New System.Drawing.Point(176, 73)
        Me.cmbPlayerName.Name = "cmbPlayerName"
        Me.cmbPlayerName.Size = New System.Drawing.Size(272, 21)
        Me.cmbPlayerName.TabIndex = 2
        '
        'dtpPaymentDate
        '
        Me.dtpPaymentDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpPaymentDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPaymentDate.Location = New System.Drawing.Point(278, 233)
        Me.dtpPaymentDate.Name = "dtpPaymentDate"
        Me.dtpPaymentDate.Size = New System.Drawing.Size(174, 20)
        Me.dtpPaymentDate.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label3.Enabled = False
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(468, 194)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 17)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "قيمة المبلغ"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSearchResult
        '
        Me.lblSearchResult.AutoSize = True
        Me.lblSearchResult.Enabled = False
        Me.lblSearchResult.Location = New System.Drawing.Point(709, 458)
        Me.lblSearchResult.Name = "lblSearchResult"
        Me.lblSearchResult.Size = New System.Drawing.Size(56, 13)
        Me.lblSearchResult.TabIndex = 49
        Me.lblSearchResult.Text = "عدد النتائج"
        Me.lblSearchResult.Visible = False
        '
        'ExporteToExcelToolStripMenuItem
        '
        Me.ExporteToExcelToolStripMenuItem.Name = "ExporteToExcelToolStripMenuItem"
        Me.ExporteToExcelToolStripMenuItem.Size = New System.Drawing.Size(155, 22)
        Me.ExporteToExcelToolStripMenuItem.Text = "تصدير الى اكسل"
        '
        'Label16
        '
        Me.Label16.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label16.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label16.Enabled = False
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(468, 233)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(74, 17)
        Me.Label16.TabIndex = 36
        Me.Label16.Text = "تاريخ الدفع"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.ExitToolStripMenuItem.Text = "خروج"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(391, 192)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(57, 20)
        Me.txtAmount.TabIndex = 4
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(186, 312)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(272, 56)
        Me.txtNotes.TabIndex = 9
        Me.txtNotes.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label6.Enabled = False
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(468, 311)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(74, 17)
        Me.Label6.TabIndex = 32
        Me.Label6.Text = "ملاحظات"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbPaymentMethod
        '
        Me.cmbPaymentMethod.FormattingEnabled = True
        Me.cmbPaymentMethod.Items.AddRange(New Object() {"نقدا", "تحويل بنكي", "شيك"})
        Me.cmbPaymentMethod.Location = New System.Drawing.Point(93, 189)
        Me.cmbPaymentMethod.Name = "cmbPaymentMethod"
        Me.cmbPaymentMethod.Size = New System.Drawing.Size(82, 21)
        Me.cmbPaymentMethod.TabIndex = 6
        '
        'ExporteToolStripMenuItem
        '
        Me.ExporteToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExporteToExcelToolStripMenuItem})
        Me.ExporteToolStripMenuItem.Name = "ExporteToolStripMenuItem"
        Me.ExporteToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.ExporteToolStripMenuItem.Text = "تصدير"
        '
        'lblSummary
        '
        Me.lblSummary.AutoSize = True
        Me.lblSummary.Enabled = False
        Me.lblSummary.Location = New System.Drawing.Point(407, 464)
        Me.lblSummary.Name = "lblSummary"
        Me.lblSummary.Size = New System.Drawing.Size(38, 13)
        Me.lblSummary.TabIndex = 51
        Me.lblSummary.Text = "Label7"
        Me.lblSummary.Visible = False
        '
        'txtTransferorName
        '
        Me.txtTransferorName.Location = New System.Drawing.Point(176, 113)
        Me.txtTransferorName.Name = "txtTransferorName"
        Me.txtTransferorName.Size = New System.Drawing.Size(272, 20)
        Me.txtTransferorName.TabIndex = 3
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(9, 458)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(452, 27)
        Me.ProgressBar1.TabIndex = 53
        '
        'Label14
        '
        Me.Label14.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label14.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label14.Enabled = False
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(468, 116)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label14.Size = New System.Drawing.Size(74, 17)
        Me.Label14.TabIndex = 16
        Me.Label14.Text = "اسم المحول"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.SaveToolStripMenuItem.Text = "حفظ"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.New1ToolStripMenuItem, Me.BulkPaymentToolStripMenuItem})
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.NewToolStripMenuItem.Text = "جديد"
        '
        'New1ToolStripMenuItem
        '
        Me.New1ToolStripMenuItem.Name = "New1ToolStripMenuItem"
        Me.New1ToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.New1ToolStripMenuItem.Text = "وصل جديد للاعب واحد"
        '
        'BulkPaymentToolStripMenuItem
        '
        Me.BulkPaymentToolStripMenuItem.Name = "BulkPaymentToolStripMenuItem"
        Me.BulkPaymentToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.BulkPaymentToolStripMenuItem.Text = "وصل جديد لعدة لاعبين"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.SaveToolStripMenuItem, Me.DeleteToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(42, 20)
        Me.FileToolStripMenuItem.Text = "ملف"
        '
        'DeleteToolStripMenuItem
        '
        Me.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem"
        Me.DeleteToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.DeleteToolStripMenuItem.Text = "حذف"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ExporteToolStripMenuItem, Me.ReportToolStripMenuItem, Me.ImageToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.MenuStrip1.Size = New System.Drawing.Size(1266, 24)
        Me.MenuStrip1.TabIndex = 48
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ImageToolStripMenuItem
        '
        Me.ImageToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddimageToolStripMenuItem, Me.DeleteimageToolStripMenuItem, Me.PrintimageToolStripMenuItem})
        Me.ImageToolStripMenuItem.Name = "ImageToolStripMenuItem"
        Me.ImageToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.ImageToolStripMenuItem.Text = "صورة"
        '
        'AddimageToolStripMenuItem
        '
        Me.AddimageToolStripMenuItem.Name = "AddimageToolStripMenuItem"
        Me.AddimageToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.AddimageToolStripMenuItem.Text = "اضافة صورة"
        '
        'DeleteimageToolStripMenuItem
        '
        Me.DeleteimageToolStripMenuItem.Name = "DeleteimageToolStripMenuItem"
        Me.DeleteimageToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.DeleteimageToolStripMenuItem.Text = "حذف صورة"
        '
        'PrintimageToolStripMenuItem
        '
        Me.PrintimageToolStripMenuItem.Name = "PrintimageToolStripMenuItem"
        Me.PrintimageToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.PrintimageToolStripMenuItem.Text = "طباعة صورة"
        '
        'Label17
        '
        Me.Label17.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label17.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label17.Enabled = False
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(468, 77)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(74, 17)
        Me.Label17.TabIndex = 12
        Me.Label17.Text = "الاسم رباعي"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtVoucherNumber
        '
        Me.txtVoucherNumber.Location = New System.Drawing.Point(367, 37)
        Me.txtVoucherNumber.Name = "txtVoucherNumber"
        Me.txtVoucherNumber.Size = New System.Drawing.Size(81, 20)
        Me.txtVoucherNumber.TabIndex = 1
        '
        'Label18
        '
        Me.Label18.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label18.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label18.Enabled = False
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(468, 38)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(74, 17)
        Me.Label18.TabIndex = 10
        Me.Label18.Text = "رقم الوصل"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(9, 34)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(436, 421)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 52
        Me.PictureBox1.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.txtMobileNumber)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.cmbCurrencyType)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.dtpDueDate)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cmbPlayerName)
        Me.GroupBox1.Controls.Add(Me.dtpPaymentDate)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtAmount)
        Me.GroupBox1.Controls.Add(Me.txtNotes)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.cmbPaymentMethod)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtTransferorName)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.txtVoucherNumber)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(712, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.GroupBox1.Size = New System.Drawing.Size(548, 419)
        Me.GroupBox1.TabIndex = 47
        Me.GroupBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(199, 190)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 17)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "طريقة الدفع"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DataGridView1.Location = New System.Drawing.Point(0, 465)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1266, 165)
        Me.DataGridView1.TabIndex = 46
        '
        'frmPayments
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1266, 630)
        Me.Controls.Add(Me.lblTotalAmount)
        Me.Controls.Add(Me.lblSearchResult)
        Me.Controls.Add(Me.lblSummary)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.DataGridView1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPayments"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "شاشة الايرادات"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTotalAmount As Label
    Friend WithEvents PaymentReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents txtMobileNumber As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents cmbCurrencyType As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents dtpDueDate As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents PrintReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GeneralReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents cmbPlayerName As ComboBox
    Friend WithEvents dtpPaymentDate As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents lblSearchResult As Label
    Friend WithEvents ExporteToExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Label16 As Label
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents txtAmount As TextBox
    Friend WithEvents txtNotes As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents cmbPaymentMethod As ComboBox
    Friend WithEvents ExporteToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents lblSummary As Label
    Friend WithEvents txtTransferorName As TextBox
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents Label14 As Label
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DeleteToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents Label17 As Label
    Friend WithEvents txtVoucherNumber As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ImageToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AddimageToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DeleteimageToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintimageToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Label1 As Label
    Friend WithEvents New1ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents BulkPaymentToolStripMenuItem As ToolStripMenuItem
End Class
