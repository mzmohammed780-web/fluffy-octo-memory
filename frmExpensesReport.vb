Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System.Globalization ' 🌟 لضبط تحليل المبالغ مستقل عن إعدادات الويندوز
Imports System.Drawing.Printing
Imports System.Linq
Imports System.Threading.Tasks

Public Class frmExpensesReport

#Region "Variables"

    Private WithEvents printDoc As New Printing.PrintDocument()
    Private printPreview As New PrintPreviewDialog()
    Private printColumns As New List(Of String)
    Private _printRows As List(Of DataGridViewRow)

    Private _totalILS As Decimal = 0
    Private _totalUSD As Decimal = 0
    Private _totalJOD As Decimal = 0
    Private _totalCount As Integer = 0
    Private isLoadingDefault As Boolean = True   ' 🌟 هل التحميل الحالي هو العرض الافتراضي (بدون فلترة)؟
    ' 🌟 إصلاح زر الخروج: كان FormClosing يلغي كل إغلاق برمجي (CloseReason.None) —
    ' وMe.Close() البرمجي يعطي بالضبط هذه القيمة، فكانت قائمة الخروج لا تعمل أبداً.
    ' الآن: العلم يُرفع قبل الإغلاق المتعمّد فقط
    Private _programmaticClose As Boolean = False

#End Region

#Region "Constructor & Form Events"

    Public Sub New()
        InitializeComponent()
        Me.DialogResult = Windows.Forms.DialogResult.None
    End Sub

    Private Sub frmExpensesReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DialogResult = Windows.Forms.DialogResult.None
        Try
            SetupForm()
            SetDefaultDates()
        Catch ex As Exception
            DatabaseModule.LogError("frmExpensesReport_Load", ex)
            MessageBox.Show($"خطأ في تحميل التقرير: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub frmExpensesReport_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            isLoadingDefault = True          ' 🌟 عرض كل المصاريف عند الفتح
            Await LoadExpensesReportAsync()
        Catch ex As Exception
            DatabaseModule.LogError("frmExpensesReport_Shown", ex)
        End Try
    End Sub
    Private Sub frmExpensesReport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = CloseReason.None AndAlso Not _programmaticClose Then
            e.Cancel = True
            Return
        End If

        Try
            printDoc?.Dispose()
            printPreview?.Dispose()
        Catch
        End Try
    End Sub

#End Region

#Region "Form Setup"

    Private Sub SetupForm()
        Me.Text = "تقرير المصروفات"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.RightToLeft = RightToLeft.Yes

        If dgvReport IsNot Nothing Then
            dgvReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvReport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dgvReport.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvReport.RowHeadersVisible = False
            dgvReport.BackgroundColor = Color.White
            dgvReport.EnableHeadersVisualStyles = False
            dgvReport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(64, 64, 64)      ' 🌟 توحيد مع GridHelper
            dgvReport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            dgvReport.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold) ' 🌟 توحيد مع باقي الشاشات
            dgvReport.ColumnHeadersHeight = 35                                                   ' 🌟 توحيد
            dgvReport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255) ' 🌟 توحيد AliceBlue
            dgvReport.DefaultCellStyle.Font = New Font("Arial", 9)
            dgvReport.DefaultCellStyle.Padding = New Padding(2)
            dgvReport.DefaultCellStyle.SelectionBackColor = Color.FromArgb(135, 206, 250)
            dgvReport.DefaultCellStyle.SelectionForeColor = Color.Black
            AddHandler dgvReport.DataBindingComplete, AddressOf OnDataBindingComplete
        End If

        If lblTotal IsNot Nothing Then
            lblTotal.Font = New Font("Arial", 10, FontStyle.Bold)
            lblTotal.ForeColor = Color.Black
            lblTotal.BackColor = Color.FromArgb(255, 255, 200)
            lblTotal.BorderStyle = BorderStyle.FixedSingle
            lblTotal.Padding = New Padding(10, 0, 10, 0)
            lblTotal.TextAlign = ContentAlignment.MiddleLeft
        End If
    End Sub

    Private Sub OnDataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        Try
            If dgvReport.Columns.Contains("Description") Then dgvReport.Columns("Description").MinimumWidth = 200
            If dgvReport.Columns.Contains("Notes") Then dgvReport.Columns("Notes").MinimumWidth = 150
            If dgvReport.Columns.Contains("Category") Then dgvReport.Columns("Category").MinimumWidth = 100
            ' 🌟 توسيع عمود التاريخ ليصبح 120 ليتسع للتاريخ بشكل مريح
            If dgvReport.Columns.Contains("ExpenseDate") Then dgvReport.Columns("ExpenseDate").Width = 120
        Catch ex As Exception
            DatabaseModule.LogError("OnDataBindingComplete", ex)
        End Try
    End Sub

    Private Sub SetDefaultDates()
        ' 🌟 الفترة الافتراضية = من بداية السنة الحالية إلى نهايتها (تعرض بيانات كافية عند الفتح)
        ' يمكن للمستخدم تعديل "من" لاحقاً لمدى أوسع إن أراد
        If dtpFrom IsNot Nothing Then dtpFrom.Value = New Date(Date.Today.Year, 1, 1)
        If dtpTo IsNot Nothing Then dtpTo.Value = Date.Today
    End Sub
#End Region

#Region "Menu Click Handlers"

    Private Async Sub ShowReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowReportToolStripMenuItem.Click
        Await LoadExpensesReportAsync()
    End Sub

    Private Sub PrintReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintReportToolStripMenuItem.Click
        If dgvReport.Rows.Count = 0 Then
            MessageBox.Show("لا توجد بيانات للطباعة. الرجاء عرض التقرير أولاً.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Using frmSelect As New frmSelectColumns()
                frmSelect.OperationType = "Print"
                frmSelect.SourceDataGridView = dgvReport
                frmSelect.StartPosition = FormStartPosition.CenterParent

                If frmSelect.ShowDialog() <> DialogResult.OK Then Return

                printColumns = frmSelect.SelectedColumns
                If printColumns Is Nothing OrElse printColumns.Count = 0 Then
                    MessageBox.Show("لم يتم تحديد أي عمود للطباعة", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            End Using

            _printRows = dgvReport.Rows.Cast(Of DataGridViewRow)().
                         Where(Function(r) Not r.IsNewRow).
                         OrderBy(Function(r) UtilityModule.SafeDate(r.Cells("ExpenseDate").Value)).
                         ToList()

            Using helper As New PrintHelper() With {
                .SourceGrid = dgvReport,
                .RowsToPrint = _printRows,
                .ColumnsToPrint = printColumns,
                .ReportTitle = "تقرير المصروفات",
                .FooterSummary = lblTotal.Text
            }
                helper.ShowPreview()
            End Using

        Catch ex As Exception
            DatabaseModule.LogError("PrintReportToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ أثناء الطباعة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        ExportToCSV()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        _programmaticClose = True ' 🌟 إصلاح: رفع العلم قبل الإغلاق حتى لا يلغيه FormClosing
        Me.Close()
    End Sub

#End Region

#Region "Report Loading - Memory Date Filtering"

    Private Async Function LoadExpensesReportAsync() As Task
        Try
            Dim fromDate As Date = dtpFrom.Value.Date
            Dim toDate As Date = dtpTo.Value.Date

            ' 🌟 التقاط حالة التحميل الافتراضي قبل تصفيرها — لعرض تذييل صادق (كل الفترات)
            Dim wasDefaultLoad As Boolean = isLoadingDefault

            If fromDate > toDate Then
                MessageBox.Show("تاريخ البداية لا يمكن أن يكون أكبر من تاريخ النهاية", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' 🌟 الفلترة داخل SQL مباشرة — العرض الافتراضي الأول يجلب الكل، وبعدها الفلترة بالفترة (على فهرس ExpenseDate)
            Dim sql As String = "SELECT VoucherNumber, Description, Amount, CurrencyType, ExpenseDate, Category, Notes FROM Expenses"

            Dim params As New List(Of SQLiteParameter)
            If Not isLoadingDefault Then
                sql &= " WHERE ExpenseDate >= @from AND ExpenseDate <= @to"
                ' 🌟 إصلاح: حدود الفلتر بثقافة ثابتة لتطابق التخزين ISO على أي جهاز
                params.Add(New SQLiteParameter("@from", fromDate.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)))
                params.Add(New SQLiteParameter("@to", toDate.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)))
            End If

            Dim resultDt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, params.ToArray())
            isLoadingDefault = False   ' 🌟 الاستعلامات القادمة (زر عرض التقرير) تُفلتر بالفترة طبيعياً

            If resultDt IsNot Nothing Then
                ' 🌟 تحويل التواريخ المخزنة ISO إلى صيغة العرض dd/MM/yyyy
                UtilityModule.FormatDateColumnsForDisplay(resultDt, "ExpenseDate")

                dgvReport.DataSource = resultDt
                AddSequenceColumn()
                SetArabicHeaders()
                CalculateTotals(resultDt, wasDefaultLoad)

                Dim totalCount As Integer = resultDt.Rows.Count
                Me.Text = $"تقرير المصروفات - {totalCount} مصروف"

                If totalCount = 0 Then
                    lblTotal.Text = "لا توجد مصروفات مطابقة"
                    lblTotal.BackColor = Color.LightCoral
                Else
                    lblTotal.BackColor = Color.FromArgb(255, 255, 200)
                End If
            End If

        Catch ex As Exception
            DatabaseModule.LogError("LoadExpensesReportAsync", ex)
            MessageBox.Show("خطأ في تحميل التقرير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Private Sub AddSequenceColumn()
        Try
            If Not dgvReport.Columns.Contains("SeqColumn") Then
                Dim seqColumn As New DataGridViewTextBoxColumn() With {
                    .Name = "SeqColumn",
                    .HeaderText = "ت",
                    .Width = 45,
                    .ReadOnly = True
                }
                seqColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                seqColumn.DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
                dgvReport.Columns.Insert(0, seqColumn)
            End If

            For i As Integer = 0 To dgvReport.Rows.Count - 1
                dgvReport.Rows(i).Cells("SeqColumn").Value = i + 1
            Next
        Catch ex As Exception
            DatabaseModule.LogError("AddSequenceColumn", ex)
        End Try
    End Sub

    Private Sub SetArabicHeaders()
        For Each col As DataGridViewColumn In dgvReport.Columns
            Select Case col.Name
                Case "ExpenseDate" : col.HeaderText = "التاريخ"
                Case "VoucherNumber" : col.HeaderText = "رقم السند"
                Case "Description" : col.HeaderText = "البيان"
                Case "Amount" : col.HeaderText = "المبلغ"
                Case "CurrencyType" : col.HeaderText = "العملة"
                Case "Category" : col.HeaderText = "الفئة"
                Case "Notes" : col.HeaderText = "ملاحظات"
                Case "SeqColumn" : col.HeaderText = "ت"
            End Select

            Select Case col.Name
                Case "Description", "Notes", "Category"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    col.DefaultCellStyle.WrapMode = DataGridViewTriState.True
                Case "ExpenseDate"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    col.DefaultCellStyle.Format = "dd/MM/yyyy"
                Case "VoucherNumber"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                Case "Amount"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    col.DefaultCellStyle.Format = "#,##0.##"
                Case "CurrencyType", "SeqColumn"
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End Select
        Next
    End Sub

    ''' <summary>حساب مجاميع العملات وعرض التذييل</summary>
    ''' <param name="wasDefaultLoad">🌟 إذا كان العرض الافتراضي (كل الفترات) لا يُعرض نطاق تواريخ كأنه مطبّق</param>
    Private Sub CalculateTotals(dt As DataTable, wasDefaultLoad As Boolean)
        Try
            ' 🌟 توحيد حساب المجاميع: UtilityModule.SumCurrencyTotalsFromTable (كان مكرراً في 4 نماذج)
            Dim t As CurrencyTotals = UtilityModule.SumCurrencyTotalsFromTable(dt, "Amount", "CurrencyType")
            _totalILS = t.ILS
            _totalUSD = t.USD
            _totalJOD = t.JOD
            _totalCount = t.Count

            ' 🌟 إصلاح تذييل مضلل: كان يعرض نطاق التاريخ حتى لو كانت الأرقام لكل الفترات
            Dim periodText As String = If(wasDefaultLoad, "كل الفترات", $"{dtpFrom.Value:dd/MM/yyyy} - {dtpTo.Value:dd/MM/yyyy}")
            Dim result As String = $"{periodText}   |   {_totalCount} مصروف"
            If _totalILS > 0 Then result &= $"   |   {UtilityModule.FormatCurrencyPart(_totalILS, AppConstants.Currency_ILS)}"
            If _totalUSD > 0 Then result &= $"   |   {UtilityModule.FormatCurrencyPart(_totalUSD, AppConstants.Currency_USD)}"
            If _totalJOD > 0 Then result &= $"   |   {UtilityModule.FormatCurrencyPart(_totalJOD, AppConstants.Currency_JOD)}"

            lblTotal.Text = result
        Catch ex As Exception
            DatabaseModule.LogError("CalculateTotals", ex)
        End Try
    End Sub
#End Region

#Region "Export to CSV"

    Private Sub ExportToCSV()
        If dgvReport.Rows.Count = 0 Then
            MessageBox.Show("لا توجد بيانات للتصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Using sfd As New SaveFileDialog()
                sfd.Filter = "CSV Files|*.csv"
                sfd.DefaultExt = "csv"
                sfd.FileName = $"تقرير المصروفات_{dtpFrom.Value:yyyy-MM-dd}_الي_{dtpTo.Value:yyyy-MM-dd}"

                If sfd.ShowDialog() = DialogResult.OK Then
                    Cursor = Cursors.WaitCursor
                    Dim utf8WithBom As New UTF8Encoding(True)
                    Using sw As New StreamWriter(sfd.FileName, False, utf8WithBom)
                        sw.WriteLine("تقرير المصروفات")
                        sw.WriteLine($"الفترة: {dtpFrom.Value:dd/MM/yyyy} - {dtpTo.Value:dd/MM/yyyy}")
                        sw.WriteLine($"تاريخ التصدير: {DateTime.Now:dd/MM/yyyy HH:mm}")
                        sw.WriteLine()

                        ' 🌟 توحيد CSV: المنطق المشترك في UtilityModule.WriteGridCsv (كان منسوخاً في 5 نماذج)
                        ' + تهيئة التاريخ والمبلغ بثقافة ثابتة داخل الدالة المشتركة
                        UtilityModule.WriteGridCsv(dgvReport, sw,
                            New String() {"SeqColumn"},
                            Nothing,
                            New String() {"ExpenseDate"},
                            New String() {"Amount"})

                        sw.WriteLine()
                        sw.WriteLine("الإجماليات:")
                        ' 🌟 إصلاح: سطر الإجماليات كان يُكتب خاماً — فواصل الآلاف (#,##0.##)
                        ' تفتح أعمدة جديدة في Excel وينزاح السطر. الآن تهريب سليم
                        sw.WriteLine("الإجماليات: " & UtilityModule.EscapeCsvValue(lblTotal.Text))
                    End Using

                    Cursor = Cursors.Default
                    MessageBox.Show("تم تصدير التقرير بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    If MessageBox.Show("هل تريد فتح الملف الآن؟", "فتح الملف", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        Try
                            Process.Start(New ProcessStartInfo(sfd.FileName) With {.UseShellExecute = True})
                        Catch
                            MessageBox.Show("تعذّر فتح الملف تلقائياً. يمكنك فتحه يدوياً من: " & sfd.FileName)
                        End Try
                    End If
                End If
            End Using
        Catch ex As Exception
            Cursor = Cursors.Default
            DatabaseModule.LogError("ExportToCSV", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

End Class
