Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — عارض سجل التدقيق
' فلاتر: نطاق تاريخ + مستخدم + نوع حدث + بحث حر، مع تحديد أقصى صفوف
' وتصدير CSV وإعادة تدوير السجلات القديمة (المدير فقط).
' واجهة مبنية بالكود بالكامل — نفس نمط frmLogin
' ═══════════════════════════════════════════════════════════════════
Public Class frmAuditLog
    Inherits System.Windows.Forms.Form

    Private Const MaxRows As Integer = 2000

    Private lblFrom As New Label()
    Private dtpFrom As New DateTimePicker()
    Private lblTo As New Label()
    Private dtpTo As New DateTimePicker()
    Private lblUser As New Label()
    Private cmbUser As New ComboBox()
    Private lblAction As New Label()
    Private cmbAction As New ComboBox()
    Private lblSearch As New Label()
    Private txtSearch As New TextBox()
    Private btnSearch As New Button()
    Private btnReset As New Button()
    Private btnExport As New Button()
    Private btnPurge As New Button()
    Private btnVerify As New Button()
    Private lblSummary As New Label()
    Private dgv As New DataGridView()

    Private ReadOnly _service As New AuditService()
    Private _isLoading As Boolean = False

    Public Sub New()
        MyBase.New()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Try
            Me.Text = "سجل التدقيق — من قام بماذا ومتى"
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Size = New Size(1050, 680)
            Me.MinimumSize = New Size(900, 600)
            Me.RightToLeft = RightToLeft.Yes
            Me.BackColor = Color.FromArgb(243, 244, 246)
            Me.Font = New Font("Arial", 9.0!)

            ' ── شريط الفلاتر (أعلى الشاشة) ──
            lblFrom.Text = "من:"
            lblFrom.AutoSize = True
            lblFrom.Location = New Point(990, 22)

            dtpFrom.Format = DateTimePickerFormat.Short
            dtpFrom.Width = 110
            dtpFrom.Location = New Point(870, 18)
            dtpFrom.Value = DateTime.Now.AddDays(-30)

            lblTo.Text = "إلى:"
            lblTo.AutoSize = True
            lblTo.Location = New Point(820, 22)

            dtpTo.Format = DateTimePickerFormat.Short
            dtpTo.Width = 110
            dtpTo.Location = New Point(700, 18)
            dtpTo.Value = DateTime.Now

            lblUser.Text = "المستخدم:"
            lblUser.AutoSize = True
            lblUser.Location = New Point(620, 22)

            cmbUser.DropDownStyle = ComboBoxStyle.DropDownList
            cmbUser.Width = 140
            cmbUser.Location = New Point(470, 18)

            lblAction.Text = "الحدث:"
            lblAction.AutoSize = True
            lblAction.Location = New Point(408, 22)

            cmbAction.DropDownStyle = ComboBoxStyle.DropDownList
            cmbAction.Width = 150
            cmbAction.Location = New Point(250, 18)

            lblSearch.Text = "بحث:"
            lblSearch.AutoSize = True
            lblSearch.Location = New Point(208, 22)

            txtSearch.Width = 80
            txtSearch.Location = New Point(125, 18)

            btnSearch.Text = "عرض"
            btnSearch.Size = New Size(50, 25)
            btnSearch.Location = New Point(70, 17)
            btnSearch.BackColor = Color.FromArgb(100, 210, 130)
            btnSearch.FlatStyle = FlatStyle.Flat
            AddHandler btnSearch.Click, AddressOf DoSearch

            btnReset.Text = "مسح"
            btnReset.Size = New Size(45, 25)
            btnReset.Location = New Point(20, 17)
            btnReset.FlatStyle = FlatStyle.Flat
            AddHandler btnReset.Click, AddressOf DoReset

            ' ── الشبكة ──
            dgv.Location = New Point(15, 55)
            dgv.Size = New Size(1010, 540)
            dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            dgv.ReadOnly = True
            dgv.AllowUserToAddRows = False
            dgv.AllowUserToDeleteRows = False
            dgv.AllowUserToResizeRows = False
            dgv.RowHeadersVisible = False
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.MultiSelect = False
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.BackgroundColor = Color.White
            dgv.BorderStyle = BorderStyle.FixedSingle
            dgv.RightToLeft = RightToLeft.Yes
            AddHandler dgv.CellFormatting, AddressOf Dgv_CellFormatting

            ' ── صف السفلى: ملخص + أزرار ──
            lblSummary.AutoSize = True
            lblSummary.Location = New Point(640, 610)
            lblSummary.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            lblSummary.Font = New Font("Arial", 9.0!, FontStyle.Bold)
            lblSummary.ForeColor = Color.FromArgb(40, 48, 68)

            btnExport.Text = "تصدير CSV"
            btnExport.Size = New Size(110, 32)
            btnExport.Location = New Point(500, 605)
            btnExport.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            btnExport.FlatStyle = FlatStyle.Flat
            AddHandler btnExport.Click, AddressOf DoExportCsv

            btnPurge.Text = "حذف السجلات الأقدم من سنة (المدير)"
            btnPurge.Size = New Size(250, 32)
            btnPurge.Location = New Point(230, 605)
            btnPurge.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            btnPurge.FlatStyle = FlatStyle.Flat
            btnPurge.BackColor = Color.FromArgb(240, 190, 190)
            AddHandler btnPurge.Click, AddressOf DoPurge

            ' 🌟 L-12: زر فحص سلامة سلسلة الهاش — للقراءة فقط ولا يعدل شيئاً
            btnVerify.Text = "فحص سلامة السجل"
            btnVerify.Size = New Size(150, 32)
            btnVerify.Location = New Point(60, 605)
            btnVerify.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            btnVerify.FlatStyle = FlatStyle.Flat
            btnVerify.BackColor = Color.FromArgb(205, 225, 250)
            AddHandler btnVerify.Click, AddressOf DoVerifyChain

            Me.Controls.AddRange({lblFrom, dtpFrom, lblTo, dtpTo, lblUser, cmbUser,
                                  lblAction, cmbAction, lblSearch, txtSearch,
                                  btnSearch, btnReset, dgv, lblSummary, btnExport, btnPurge, btnVerify})
        Catch ex As Exception
            DatabaseModule.LogError("frmAuditLog.BuildUI", ex)
        End Try
    End Sub

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        LoadFiltersAsync()
        DoSearch(Nothing, EventArgs.Empty)
    End Sub

    ''' <summary>تعبئة قائمتي المستخدم ونوع الحدث من السجل نفسه</summary>
    Private Async Sub LoadFiltersAsync()
        Try
            Dim users As List(Of String) = Await _service.GetDistinctUsernamesAsync()
            cmbUser.Items.Clear()
            cmbUser.Items.Add("(الكل)")
            For Each u As String In users
                cmbUser.Items.Add(u)
            Next
            If cmbUser.Items.Count > 0 Then cmbUser.SelectedIndex = 0

            Dim actions As List(Of String) = Await _service.GetDistinctActionsAsync()
            cmbAction.Items.Clear()
            cmbAction.Items.Add("(الكل)")
            For Each a As String In actions
                cmbAction.Items.Add(a)
            Next
            If cmbAction.Items.Count > 0 Then cmbAction.SelectedIndex = 0
        Catch ex As Exception
            DatabaseModule.LogError("frmAuditLog.LoadFiltersAsync", ex)
        End Try
    End Sub

    ''' <summary>تنفيذ الاستعلام حسب الفلاتر وعرض النتائج</summary>
    Private Async Sub DoSearch(sender As Object, e As EventArgs)
        If _isLoading Then Return
        _isLoading = True
        Try
            Me.Cursor = Cursors.WaitCursor

            ' 🌟 إصلاح: حدود الفلتر بثقافة ثابتة — كانت بثقافة النظام فتكسر المقارنة النصية مع المخزن
            Dim fromDate As String = dtpFrom.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
            Dim toDate As String = dtpTo.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
            Dim user As String = If(cmbUser.SelectedIndex > 0, Convert.ToString(cmbUser.SelectedItem), "")
            Dim action As String = If(cmbAction.SelectedIndex > 0, Convert.ToString(cmbAction.SelectedItem), "")
            Dim search As String = txtSearch.Text.Trim()

            Dim dt As DataTable = Await _service.QueryAsync(fromDate, toDate, user, action, search, MaxRows)

            If dt Is Nothing Then
                MessageBox.Show("تعذر قراءة السجل — راجع ملف اللوج", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' تسميات عربية للأعمدة
            dt.Columns("Id").ColumnName = "م"
            dt.Columns("Timestamp").ColumnName = "التاريخ والوقت"
            dt.Columns("Username").ColumnName = "المستخدم"
            dt.Columns("Action").ColumnName = "الحدث"
            dt.Columns("EntityType").ColumnName = "النوع"
            dt.Columns("EntityId").ColumnName = "المعرف"
            dt.Columns("Details").ColumnName = "التفاصيل"

            dgv.DataSource = dt

            ' تخفيت عمود التفاصيل حتى لا يبتلع الشاشة
            If dgv.Columns.Contains("التفاصيل") Then
                dgv.Columns("التفاصيل").FillWeight = 220.0!
            End If
            If dgv.Columns.Contains("م") Then
                dgv.Columns("م").FillWeight = 35.0!
            End If

            Dim total As Long = Await _service.GetTotalCountAsync()
            ' 🌟 إصلاح: الفشل يعيد -1 وكان يُعرض للمستخدم "من أصل -1 سجلاً"
            lblSummary.Text = If(total < 0,
                $"معروض: {dt.Rows.Count} سجلاً (حد العرض {MaxRows})",
                $"معروض: {dt.Rows.Count} من أصل {total} سجلاً (حد العرض {MaxRows})")
        Catch ex As Exception
            DatabaseModule.LogError("frmAuditLog.DoSearch", ex)
            MessageBox.Show("خطأ في العرض: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            _isLoading = False
        End Try
    End Sub

    ''' <summary>إعادة الفلاتر لوضعها الافتراضي</summary>
    Private Sub DoReset(sender As Object, e As EventArgs)
        dtpFrom.Value = DateTime.Now.AddDays(-30)
        dtpTo.Value = DateTime.Now
        If cmbUser.Items.Count > 0 Then cmbUser.SelectedIndex = 0
        If cmbAction.Items.Count > 0 Then cmbAction.SelectedIndex = 0
        txtSearch.Clear()
        DoSearch(Nothing, EventArgs.Empty)
    End Sub

    ''' <summary>تظليل خلايا نوع الحدث الحرجة بألوان دالة</summary>
    Private Sub Dgv_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        Try
            If e Is Nothing OrElse e.CellStyle Is Nothing Then Return
            If dgv.Columns(e.ColumnIndex).Name <> "الحدث" Then Return
            If e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return

            Dim actionText As String = Convert.ToString(e.Value)
            If actionText.Contains("حذف") OrElse actionText.Contains("فاشلة") Then
                e.CellStyle.BackColor = Color.FromArgb(255, 205, 205)
            ElseIf actionText.Contains("إضافة") Then
                e.CellStyle.BackColor = Color.FromArgb(205, 240, 215)
            ElseIf actionText.Contains("تعديل") OrElse actionText.Contains("استعادة") Then
                e.CellStyle.BackColor = Color.FromArgb(255, 235, 190)
            End If
        Catch
            ' التنسيق الشكلي لا يُسقط العرض أبداً
        End Try
    End Sub

    ''' <summary>تصدير النتائج المعروضة إلى CSV (يستخدم الكاتب الموحد WriteGridCsv)</summary>
    Private Sub DoExportCsv(sender As Object, e As EventArgs)
        Try
            If dgv.Rows.Count = 0 Then
                MessageBox.Show("لا توجد بيانات معروضة للتصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Title = "تصدير سجل التدقيق"
                sfd.Filter = "ملف CSV|*.csv"
                sfd.DefaultExt = "csv"
                sfd.FileName = $"AuditLog_{DateTime.Now:yyyy-MM-dd_HH-mm}.csv"
                If sfd.ShowDialog() <> DialogResult.OK Then Return

                ' 🌟 UTF-8 مع BOM حتى يفتح الملف بالعربية في Excel مباشرة
                Using writer As New StreamWriter(sfd.FileName, False, New UTF8Encoding(True))
                    UtilityModule.WriteGridCsv(dgv, writer, New String() {})
                End Using

                MessageBox.Show($"تم تصدير {dgv.Rows.Count} سجلاً بنجاح" & vbCrLf & sfd.FileName,
                                "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("frmAuditLog.DoExportCsv", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>حذف السجلات الأقدم من فترة الاحتفاظ — للمدير فقط</summary>
    Private Async Sub DoPurge(sender As Object, e As EventArgs)
        Try
            ' صلاحية المدير فقط (نفس نمط شاشة إدارة المستخدمين)
            If Not UserSession.CurrentUsername.ToLower() = "admin" Then
                MessageBox.Show("هذه العملية متاحة لمدير النظام (admin) فقط", "تنبيه",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' مدة الاحتفاظ من الإعدادات (افتراضي 365 يوماً) — قابلة للضبط من مكان واحد
            Dim keepDays As Integer = AppSettingsStore.GetIntSetting(AppSettingsStore.Key_AuditRetentionDays, 365, 30, 3650)

            Dim confirm As DialogResult = MessageBox.Show(
                $"سيتم حذف نهائياً كل السجلات الأقدم من {keepDays} يوماً." & vbCrLf &
                "(لتغيير المدة عدّل الإعداد Audit_RetentionDays)" & vbCrLf & vbCrLf &
                "هل أنت متأكد؟", "تأكيد التنظيف",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If confirm <> DialogResult.Yes Then Return

            Me.Cursor = Cursors.WaitCursor
            Dim deleted As Integer = Await _service.PurgeOlderThanAsync(keepDays)

            If deleted >= 0 Then
                Dim audit As New AuditService()
                audit.Log("تنظيف سجل التدقيق", "سجل", "", $"حُذف {deleted} سجلاً أقدم من {keepDays} يوماً")
                MessageBox.Show($"تم حذف {deleted} سجلاً قديماً.", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DoSearch(Nothing, EventArgs.Empty)
            Else
                MessageBox.Show("فشل التنظيف — راجع ملف اللوج", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmAuditLog.DoPurge", ex)
            MessageBox.Show("خطأ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>فحص سلامة سلسلة هاش سجل التدقيق (L-12) — للقراءة فقط ولا يعدل أي بيانات</summary>
    Private Async Sub DoVerifyChain(sender As Object, e As EventArgs)
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim result As Tuple(Of Boolean, String) = Await _service.VerifyAuditChainAsync()
            If result.Item1 Then
                MessageBox.Show("السجل سليم — لا يوجد تعديل خارجي مكتشف." & vbCrLf & result.Item2,
                                "فحص سلامة السجل", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("نتيجة الفحص: يوجد خلل محتمل." & vbCrLf & result.Item2,
                                "فحص سلامة السجل", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmAuditLog.DoVerifyChain", ex)
            MessageBox.Show("خطأ في الفحص: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

End Class
