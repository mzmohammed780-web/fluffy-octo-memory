Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Threading.Tasks
Imports System.Globalization

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — شاشة تقرير المتأخرات وغرامات التأخير
' إعدادات الاشتراك الشهري والغرامة والعملة تُحفظ في AppSettings وتظهر
' مسبقة التعبئة في كل مرة. التقرير قراءة فقط — لا يعدل أي بيانات.
' واجهة مبنية بالكود بالكامل — نفس نمط frmLogin
' ═══════════════════════════════════════════════════════════════════
Public Class frmDelinquency
    Inherits System.Windows.Forms.Form

    Private lblYear As New Label()
    Private numYear As New NumericUpDown()
    Private lblCurrency As New Label()
    Private cmbCurrency As New ComboBox()
    Private lblDue As New Label()
    Private txtMonthlyDue As New TextBox()
    Private lblFine As New Label()
    Private txtFinePerMonth As New TextBox()
    Private chkOnlyDelinquent As New CheckBox()
    Private btnGenerate As New Button()
    Private btnExportHtml As New Button()
    Private btnExportCsv As New Button()
    Private lblSummary As New Label()
    Private dgv As New DataGridView()

    Private ReadOnly _service As New DelinquencyService()
    Private _lastResult As DelinquencyService.ReportResult = Nothing
    Private _isBusy As Boolean = False

    Public Sub New()
        MyBase.New()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Try
            Me.Text = "تقرير المتأخرات وغرامات التأخير"
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Size = New Size(1100, 660)
            Me.MinimumSize = New Size(950, 580)
            Me.RightToLeft = RightToLeft.Yes
            Me.BackColor = Color.FromArgb(243, 244, 246)
            Me.Font = New Font("Arial", 9.0!)

            ' ── صف الإعدادات ──
            lblYear.Text = "السنة:"
            lblYear.AutoSize = True
            lblYear.Location = New Point(1020, 24)

            numYear.Minimum = 2000
            numYear.Maximum = 2100
            numYear.Value = DateTime.Now.Year
            numYear.Width = 75
            numYear.Location = New Point(935, 20)

            lblCurrency.Text = "العملة:"
            lblCurrency.AutoSize = True
            lblCurrency.Location = New Point(860, 24)

            cmbCurrency.DropDownStyle = ComboBoxStyle.DropDownList
            cmbCurrency.Width = 90
            cmbCurrency.Location = New Point(760, 20)
            cmbCurrency.Items.AddRange({AppConstants.Currency_ILS, AppConstants.Currency_USD, AppConstants.Currency_JOD})

            lblDue.Text = "الاشتراك الشهري:"
            lblDue.AutoSize = True
            lblDue.Location = New Point(655, 24)

            txtMonthlyDue.Width = 90
            txtMonthlyDue.Location = New Point(555, 20)
            txtMonthlyDue.TextAlign = HorizontalAlignment.Center
            AddHandler txtMonthlyDue.KeyPress, AddressOf UtilityModule.AmountKeyPressFilter

            lblFine.Text = "غرامة الشهر:"
            lblFine.AutoSize = True
            lblFine.Location = New Point(465, 24)

            txtFinePerMonth.Width = 90
            txtFinePerMonth.Location = New Point(365, 20)
            txtFinePerMonth.TextAlign = HorizontalAlignment.Center
            AddHandler txtFinePerMonth.KeyPress, AddressOf UtilityModule.AmountKeyPressFilter

            chkOnlyDelinquent.Text = "المتأخرون فقط"
            chkOnlyDelinquent.AutoSize = True
            chkOnlyDelinquent.Location = New Point(240, 22)
            AddHandler chkOnlyDelinquent.CheckedChanged, AddressOf chkOnlyDelinquent_CheckedChanged

            btnGenerate.Text = "توليد التقرير"
            btnGenerate.Size = New Size(110, 28)
            btnGenerate.Location = New Point(120, 17)
            btnGenerate.BackColor = Color.FromArgb(100, 210, 130)
            btnGenerate.FlatStyle = FlatStyle.Flat
            btnGenerate.Font = New Font("Arial", 9.0!, FontStyle.Bold)
            AddHandler btnGenerate.Click, AddressOf DoGenerate

            btnExportHtml.Text = "طباعة / PDF"
            btnExportHtml.Size = New Size(100, 28)
            btnExportHtml.Location = New Point(15, 17)
            btnExportHtml.FlatStyle = FlatStyle.Flat
            btnExportHtml.Enabled = False
            AddHandler btnExportHtml.Click, AddressOf DoExportHtml

            ' ── الشبكة ──
            dgv.Location = New Point(15, 58)
            dgv.Size = New Size(1065, 540)
            dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            dgv.ReadOnly = True
            dgv.AllowUserToAddRows = False
            dgv.AllowUserToDeleteRows = False
            dgv.RowHeadersVisible = False
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.BackgroundColor = Color.White
            dgv.BorderStyle = BorderStyle.FixedSingle
            dgv.RightToLeft = RightToLeft.Yes

            ' ── الملخص ──
            lblSummary.AutoSize = True
            lblSummary.Location = New Point(700, 608)
            lblSummary.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
            lblSummary.Font = New Font("Arial", 9.5!, FontStyle.Bold)
            lblSummary.ForeColor = Color.FromArgb(40, 48, 68)

            btnExportCsv.Text = "تصدير CSV"
            btnExportCsv.Size = New Size(100, 30)
            btnExportCsv.Location = New Point(560, 603)
            btnExportCsv.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            btnExportCsv.FlatStyle = FlatStyle.Flat
            btnExportCsv.Enabled = False
            AddHandler btnExportCsv.Click, AddressOf DoExportCsv

            ' ── ملاحظة توضيحية أسفل الشاشة ──
            Dim lblNote As New Label()
            lblNote.Text =
                "ملاحظة: الغرامة تقديرية للعرض فقط ولا تُسجل في القاعدة. المدفوعات بعملات أخرى تُعرض دون دمج (لا يوجد سعر صرف)."
            lblNote.AutoSize = True
            lblNote.Location = New Point(15, 610)
            lblNote.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            lblNote.ForeColor = Color.FromArgb(120, 70, 0)

            Me.Controls.AddRange({lblYear, numYear, lblCurrency, cmbCurrency, lblDue, txtMonthlyDue,
                                  lblFine, txtFinePerMonth, chkOnlyDelinquent, btnGenerate, btnExportHtml,
                                  dgv, lblSummary, btnExportCsv, lblNote})

            ' ── تعبئة مسبقة من الإعدادات المحفوظة ──
            Dim savedDue As Decimal = AppSettingsStore.GetDecimalSetting(AppSettingsStore.Key_Delinq_MonthlyDue, 0D)
            Dim savedFine As Decimal = AppSettingsStore.GetDecimalSetting(AppSettingsStore.Key_Delinq_FinePerMonth, 0D)
            Dim savedCurrency As String = AppSettingsStore.GetSetting(AppSettingsStore.Key_Delinq_Currency, AppConstants.Currency_ILS)

            If savedDue > 0D Then txtMonthlyDue.Text = savedDue.ToString("0.##", CultureInfo.InvariantCulture)
            If savedFine > 0D Then txtFinePerMonth.Text = savedFine.ToString("0.##", CultureInfo.InvariantCulture)
            Dim idx As Integer = cmbCurrency.Items.IndexOf(savedCurrency)
            If idx >= 0 Then cmbCurrency.SelectedIndex = idx
            If cmbCurrency.SelectedIndex < 0 Then cmbCurrency.SelectedIndex = 0
        Catch ex As Exception
            DatabaseModule.LogError("frmDelinquency.BuildUI", ex)
        End Try
    End Sub

    ''' <summary>قراءة المدخلات والتحقق منها — يعيد False مع رسالة عند فشل التحقق</summary>
    Private Function TryReadInputs(ByRef year As Integer, ByRef currency As String,
                                   ByRef monthlyDue As Decimal, ByRef finePerMonth As Decimal) As Boolean
        year = CInt(numYear.Value)

        If cmbCurrency.SelectedIndex < 0 Then
            MessageBox.Show("اختر العملة المرجعية", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        currency = Convert.ToString(cmbCurrency.SelectedItem)

        ' 🌟 تحليل المبالغ بثقافة ثابتة — نفس سياسة المشروع الموحدة
        If Not Decimal.TryParse(txtMonthlyDue.Text.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, monthlyDue) OrElse monthlyDue <= 0D Then
            MessageBox.Show("أدخل اشتراكاً شهرياً صحيحاً أكبر من صفر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        If Not Decimal.TryParse(txtFinePerMonth.Text.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, finePerMonth) OrElse finePerMonth < 0D Then
            MessageBox.Show("أدخل غرامة شهر صحيحة (0 = بدون غرامة)", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Return True
    End Function

    ''' <summary>توليد التقرير وعرضه</summary>
    Private Async Sub DoGenerate(sender As Object, e As EventArgs)
        If _isBusy Then Return

        Dim year As Integer = 0
        Dim currency As String = ""
        Dim monthlyDue As Decimal = 0D
        Dim finePerMonth As Decimal = 0D

        If Not TryReadInputs(year, currency, monthlyDue, finePerMonth) Then Return

        _isBusy = True
        Try
            Me.Cursor = Cursors.WaitCursor
            btnGenerate.Enabled = False
            btnGenerate.Text = "جاري الحساب..."

            ' 🌟 حفظ الإعدادات لهذا النسق حتى لا يعيد المستخدم إدخالها كل مرة
            AppSettingsStore.SetDecimalSetting(AppSettingsStore.Key_Delinq_MonthlyDue, monthlyDue)
            AppSettingsStore.SetDecimalSetting(AppSettingsStore.Key_Delinq_FinePerMonth, finePerMonth)
            AppSettingsStore.SetSetting(AppSettingsStore.Key_Delinq_Currency, currency)

            Dim result As DelinquencyService.ReportResult =
                Await _service.GenerateReportAsync(year, currency, monthlyDue, finePerMonth)

            If Not result.Success Then
                MessageBox.Show(result.ErrorMessage, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            _lastResult = result
            ShowResult(result)

            btnExportHtml.Enabled = True
            btnExportCsv.Enabled = True
        Catch ex As Exception
            DatabaseModule.LogError("frmDelinquency.DoGenerate", ex)
            MessageBox.Show("خطأ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isBusy = False
            Me.Cursor = Cursors.Default
            btnGenerate.Enabled = True
            btnGenerate.Text = "توليد التقرير"
        End Try
    End Sub

    ''' <summary>تعبئة الشبكة من نتيجة التقرير مع مراعاة فلتر "المتأخرون فقط"</summary>
    Private Sub ShowResult(result As DelinquencyService.ReportResult)
        Dim dt As New DataTable()
        dt.Columns.Add("م", GetType(Integer))
        dt.Columns.Add("معرف اللاعب", GetType(String))
        dt.Columns.Add("الاسم", GetType(String))
        dt.Columns.Add("الهاتف", GetType(String))
        dt.Columns.Add("تاريخ التحاقه", GetType(String))
        dt.Columns.Add("أشهر مستحقة", GetType(Integer))
        dt.Columns.Add($"مدفوع ({result.Currency})", GetType(String))
        dt.Columns.Add("مدفوع بعملات أخرى", GetType(String))
        dt.Columns.Add("المستحق", GetType(String))
        dt.Columns.Add("العجز", GetType(String))
        dt.Columns.Add("أشهر متأخرة", GetType(Integer))
        dt.Columns.Add("غرامة تقديرية", GetType(String))

        Dim seq As Integer = 0
        Dim inv As CultureInfo = CultureInfo.InvariantCulture

        For Each row As DelinquencyService.DelinquencyRow In result.Rows
            If chkOnlyDelinquent.Checked AndAlso row.MonthsBehind <= 0 Then Continue For
            seq += 1
            dt.Rows.Add(seq, row.PlayerId, row.PlayerName, row.Phone, row.JoinDate,
                        row.MonthsDue,
                        row.PaidSelected.ToString("0.##", inv),
                        row.PaidOtherCurrencies,
                        row.ExpectedTotal.ToString("0.##", inv),
                        row.Shortfall.ToString("0.##", inv),
                        row.MonthsBehind,
                        row.EstimatedFine.ToString("0.##", inv))
        Next

        dgv.DataSource = dt

        ' تظليل صفوف المتأخرين
        dgv.ClearSelection()
        Dim delinquentCount As Integer = 0
        For Each r As DataGridViewRow In dgv.Rows
            Dim behind As Integer = If(r.Cells("أشهر متأخرة").Value Is Nothing, 0, Convert.ToInt32(r.Cells("أشهر متأخرة").Value))
            If behind > 0 Then
                r.DefaultCellStyle.BackColor = Color.FromArgb(255, 220, 220)
                delinquentCount += 1
            End If
        Next

        lblSummary.Text =
            $"متأخرون: {result.DelinquentCount} من {result.Rows.Count} لاعباً" &
            $" — إجمالي العجز: {result.TotalShortfall.ToString("0.##", inv)} {result.Currency}" &
            $" — إجمالي الغرامات التقديرية: {result.TotalFine.ToString("0.##", inv)} {result.Currency}"
    End Sub

    Private Sub chkOnlyDelinquent_CheckedChanged(sender As Object, e As EventArgs)
        If _lastResult IsNot Nothing Then ShowResult(_lastResult)
    End Sub

    ''' <summary>تصدير HTML للطباعة/PDF</summary>
    Private Sub DoExportHtml(sender As Object, e As EventArgs)
        Try
            If dgv.Rows.Count = 0 Then
                MessageBox.Show("لا توجد بيانات للتصدير — ولّد التقرير أولاً", "تنبيه",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 🌟 إصلاح: العنوان كان من قيم الواجهة الحالية — تغيير السنة بعد التوليد يصدّر
            ' بيانات سنة قديمة بعنوان سنة جديدة. الآن من نتيجة التقرير المحفوظة
            If _lastResult Is Nothing Then
                MessageBox.Show("ولّد التقرير أولاً ثم صدّر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Dim subtitle As String =
                $"سنة {_lastResult.Year} — اشتراك شهري {_lastResult.MonthlyDue.ToString("0.##", Globalization.CultureInfo.InvariantCulture)} {_lastResult.Currency}"
            Dim footer As String =
                "الغرامة تقديرية وتُحسب: (العجز ÷ الاشتراك الشهري مقرباً للأعلى) × غرامة الشهر. " &
                "المدفوعات بعملات أخرى تُعرض دون دمج."

            Dim path As String = ReportExporter.ExportGridToHtml(
                dgv, "تقرير المتأخرات وغرامات التأخير", subtitle, footer, landscape:=True,
                excludedColumnNames:=New String() {"معرف اللاعب", "تاريخ التحاقه"})

            Dim audit As New AuditService()
            audit.Log(AuditService.Act_ReportExport, "تقرير", "", "تقرير المتأخرات HTML")

            MessageBox.Show("فُتح التقرير في المتصفح للطباعة/الحفظ كـ PDF." & vbCrLf & path,
                            "تم", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            DatabaseModule.LogError("frmDelinquency.DoExportHtml", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>تصدير CSV (الكاتب الموحد)</summary>
    Private Sub DoExportCsv(sender As Object, e As EventArgs)
        Try
            If dgv.Rows.Count = 0 Then
                MessageBox.Show("لا توجد بيانات للتصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Title = "تصدير تقرير المتأخرات"
                sfd.Filter = "ملف CSV|*.csv"
                sfd.DefaultExt = "csv"
                ' 🌟 إصلاح: اسم الملف من نتيجة التقرير المحفوظة لا من قيم الواجهة الحالية
                If _lastResult Is Nothing Then
                    MessageBox.Show("ولّد التقرير أولاً ثم صدّر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                sfd.FileName = $"Delinquency_{_lastResult.Year}_{DateTime.Now:yyyy-MM-dd_HH-mm}.csv"
                If sfd.ShowDialog() <> DialogResult.OK Then Return

                Using writer As New System.IO.StreamWriter(sfd.FileName, False, New System.Text.UTF8Encoding(True))
                    UtilityModule.WriteGridCsv(dgv, writer, New String() {})
                End Using

                Dim audit As New AuditService()
                audit.Log(AuditService.Act_ReportExport, "تقرير", "", "تقرير المتأخرات CSV")

                MessageBox.Show($"تم التصدير بنجاح" & vbCrLf & sfd.FileName, "نجاح",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("frmDelinquency.DoExportCsv", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
