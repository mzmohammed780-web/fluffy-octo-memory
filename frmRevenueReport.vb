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
Imports System.Windows.Forms

' 🌟 ملاحظة توثيق (أولوية 3): الكلاس "frmRevenueReport" (تقرير الإيرادات والمقبوضات) اسمه سليم،
'    لكن اسم الملف "Form5.vb" غير معبّر — أعد تسمية الملف إلى frmRevenueReport.vb
'    من Visual Studio (كليك يمين على الملف ← Rename) دون أي تأثير على الكود.
' 🌟 هذه النسخة خالية تماماً من الأقواس المعقوفة وعلامة $ حتى لا تنكسر عند نسخها
'    من الشات مع النص العربي (تغيّر اتجاه النص قد يقلب مواضع الأقواس بصرياً).
Public Class frmRevenueReport

#Region "Variables"

    Private WithEvents printDoc As New Printing.PrintDocument()
    Private printPreview As New PrintPreviewDialog()
    Private printColumns As New List(Of String)
    Private _printRows As List(Of DataGridViewRow)

#End Region

#Region "Constructor & Form Events"

    Public Sub New()
        InitializeComponent()
        Me.DialogResult = DialogResult.None
    End Sub

    Private Sub frmRevenueReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DialogResult = DialogResult.None
        Try
            SetupForm()
            SetDefaultDates()
        Catch ex As Exception
            DatabaseModule.LogError("frmRevenueReport_Load", ex)
            MessageBox.Show("خطأ في تحميل النموذج: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub frmRevenueReport_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            Await LoadRevenueReportAsync()
        Catch ex As Exception
            DatabaseModule.LogError("frmRevenueReport_Shown", ex)
        End Try
    End Sub

    Private Sub frmRevenueReport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            ' 🌟 استبدال عامل الوصول الشرطي بفحص صريح (أسلوب متوافق مع كل إصدارات المترجم)
            If printDoc IsNot Nothing Then printDoc.Dispose()
            If printPreview IsNot Nothing Then printPreview.Dispose()
        Catch
        End Try
    End Sub

#End Region

#Region "Form Setup"

    Private Sub SetupForm()
        Me.Text = "تقرير الإيرادات والمقبوضات"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.RightToLeft = RightToLeft.Yes

        If dgvReport IsNot Nothing Then
            dgvReport.RightToLeft = RightToLeft.Yes
            dgvReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgvReport.RowHeadersVisible = False
            dgvReport.BackgroundColor = Color.White
            dgvReport.EnableHeadersVisualStyles = False
            dgvReport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 58, 64)
            dgvReport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            dgvReport.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9.5!, FontStyle.Bold)
            dgvReport.ColumnHeadersHeight = 38
            dgvReport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 250)
            dgvReport.DefaultCellStyle.Font = New Font("Arial", 9)
            dgvReport.DefaultCellStyle.SelectionBackColor = Color.FromArgb(135, 206, 250)
            dgvReport.DefaultCellStyle.SelectionForeColor = Color.Black
            AddHandler dgvReport.DataBindingComplete, AddressOf OnDataBindingComplete
        End If

        If lblTotal IsNot Nothing Then
            lblTotal.Font = New Font("Arial", 10, FontStyle.Bold)
            lblTotal.ForeColor = Color.Black
            lblTotal.BackColor = Color.FromArgb(255, 255, 200)
            lblTotal.BorderStyle = BorderStyle.FixedSingle
            lblTotal.TextAlign = ContentAlignment.MiddleLeft
        End If
    End Sub

    Private Sub OnDataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        Try
            If dgvReport.Columns.Contains(AppConstants.Col_Pay_PlayerName) Then dgvReport.Columns(AppConstants.Col_Pay_PlayerName).Width = 180
            If dgvReport.Columns.Contains(AppConstants.Col_Pay_TransferorName) Then dgvReport.Columns(AppConstants.Col_Pay_TransferorName).Width = 160
            If dgvReport.Columns.Contains(AppConstants.Col_Pay_Notes) Then dgvReport.Columns(AppConstants.Col_Pay_Notes).Width = 180
            If dgvReport.Columns.Contains(AppConstants.Col_Pay_Amount) Then dgvReport.Columns(AppConstants.Col_Pay_Amount).Width = 90
            If dgvReport.Columns.Contains(AppConstants.Col_Pay_VoucherNumber) Then dgvReport.Columns(AppConstants.Col_Pay_VoucherNumber).Width = 90
            If dgvReport.Columns.Contains(AppConstants.Col_Pay_PaymentDate) Then dgvReport.Columns(AppConstants.Col_Pay_PaymentDate).Width = 95

            ' 🌟 تم زيادة عرض عمود تاريخ الاستحقاق ليصبح 120 بدلاً من 95
            If dgvReport.Columns.Contains(AppConstants.Col_Pay_DueDate) Then dgvReport.Columns(AppConstants.Col_Pay_DueDate).Width = 120
        Catch ex As Exception
            DatabaseModule.LogError("OnDataBindingComplete", ex)
        End Try
    End Sub

    Private Sub SetDefaultDates()
        If dtpFrom IsNot Nothing Then dtpFrom.Value = New Date(Date.Today.Year, Date.Today.Month, 1)
        If dtpTo IsNot Nothing Then dtpTo.Value = Date.Today
    End Sub

#End Region

#Region "Menu Click Handlers"

    ' 🌟 تم إزالة التكرار في الـ Handles
    Private Async Sub ShowReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowReportToolStripMenuItem.Click
        Await LoadRevenueReportAsync()
    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        ExportToExcel()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

#End Region

#Region "Report Loading - SQL Date Filtering"


    Private Async Function LoadRevenueReportAsync() As Task
        Try
            Dim fromDate As Date = dtpFrom.Value.Date
            Dim toDate As Date = dtpTo.Value.Date

            If fromDate > toDate Then
                MessageBox.Show("تاريخ البداية لا يمكن أن يكون أكبر من تاريخ النهاية!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' 🌟 الفلترة داخل SQL مباشرة: التواريخ مخزنة ISO (yyyy-MM-dd) فالمقارنة النصية صحيحة،
            ' وتستخدم فهرس idx_payment_date بدل تحميل الجدول كاملاً في الذاكرة
            Dim sql As String = "SELECT " & AppConstants.Col_Pay_PlayerName & ", " & AppConstants.Col_Pay_TransferorName & ", " &
                                AppConstants.Col_Pay_VoucherNumber & ", " & AppConstants.Col_Pay_Amount & ", " & AppConstants.Col_Pay_CurrencyType & ", " &
                                AppConstants.Col_Pay_PaymentMethod & ", " & AppConstants.Col_Pay_PaymentDate & ", " & AppConstants.Col_Pay_DueDate & ", " & AppConstants.Col_Pay_Notes & " " &
                                "FROM " & AppConstants.Table_Payments & " " &
                                "WHERE " & AppConstants.Col_Pay_PaymentDate & " >= @from AND " & AppConstants.Col_Pay_PaymentDate & " <= @to"

            ' 🌟 إصلاح: حدود الفلتر بثقافة ثابتة لتطابق التخزين ISO على أي جهاز
            ' 🌟 تمرير البارامترات مباشرة (ParamArray) دون مصفوفة مهيّأة بالأقواس المعقوفة
            Dim resultDt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql,
                New SQLiteParameter("@from", fromDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
                New SQLiteParameter("@to", toDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))

            ' 🌟 تحويل التواريخ المخزنة ISO إلى صيغة العرض dd/MM/yyyy
            UtilityModule.FormatDateColumnsForDisplay(resultDt, AppConstants.Col_Pay_PaymentDate, AppConstants.Col_Pay_DueDate)

            dgvReport.SuspendLayout()
            dgvReport.DataSource = resultDt
            AddSequenceColumn()
            SetArabicHeaders()
            CalculateTotals(resultDt)
            dgvReport.ResumeLayout()

            Dim totalCount As Integer = resultDt.Rows.Count
            Me.Text = "تقرير الإيرادات - " & totalCount.ToString() & " دفعة"

            If totalCount = 0 Then
                lblTotal.Text = "  لا توجد دفعات مسجلة في هذه الفترة المحددة"
                lblTotal.BackColor = Color.FromArgb(255, 220, 220)
            Else
                lblTotal.BackColor = Color.FromArgb(220, 255, 220)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("LoadRevenueReportAsync", ex)
            MessageBox.Show("خطأ في تحميل التقرير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Private Sub AddSequenceColumn()
        Try
            GridHelper.FillSequenceNumbers(dgvReport)
        Catch ex As Exception
            DatabaseModule.LogError("frmRevenueReport.AddSequenceColumn", ex)
        End Try
    End Sub


    Private Sub SetArabicHeaders()
        For Each col As DataGridViewColumn In dgvReport.Columns
            Select Case col.Name
                Case AppConstants.Col_Pay_PaymentDate : col.HeaderText = "تاريخ الدفع"
                Case AppConstants.Col_Pay_PlayerName : col.HeaderText = "اسم اللاعب"
                Case AppConstants.Col_Pay_TransferorName : col.HeaderText = "اسم المحول"
                Case AppConstants.Col_Pay_VoucherNumber : col.HeaderText = "رقم السند"
                Case AppConstants.Col_Pay_Amount : col.HeaderText = "المبلغ"
                Case AppConstants.Col_Pay_CurrencyType : col.HeaderText = "العملة"
                Case AppConstants.Col_Pay_PaymentMethod : col.HeaderText = "طريقة الدفع"
                Case AppConstants.Col_Pay_DueDate : col.HeaderText = "تاريخ الاستحقاق"
                Case AppConstants.Col_Pay_Notes : col.HeaderText = "ملاحظات"
                Case "SeqColumn" : col.HeaderText = "ت"
            End Select

            Select Case col.Name
                Case AppConstants.Col_Pay_PlayerName, AppConstants.Col_Pay_TransferorName, AppConstants.Col_Pay_Notes
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                Case AppConstants.Col_Pay_PaymentDate, AppConstants.Col_Pay_DueDate
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                Case AppConstants.Col_Pay_Amount
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter   ' 🌟 توسيط بدل اليمين
                    col.DefaultCellStyle.Format = "#,##0.##"
                Case Else
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End Select
        Next
    End Sub

    Private Sub CalculateTotals(dt As DataTable)
        Try
            ' 🌟 توحيد حساب المجاميع: UtilityModule.SumCurrencyTotalsFromTable (كان مكرراً في 4 نماذج)
            Dim t As CurrencyTotals = UtilityModule.SumCurrencyTotalsFromTable(
                dt, AppConstants.Col_Pay_Amount, AppConstants.Col_Pay_CurrencyType)

            ' 🌟 استبدال السلاسل المُقحمة (Interpolation) بدمج نصي عادي — آمن تماماً عند النسخ من الشات
            Dim result As String = "  الفترة: " & dtpFrom.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) &
                                   " إلى " & dtpTo.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) &
                                   "  |  عدد الدفعات: " & t.Count.ToString() &
                                   "  |  " & UtilityModule.FormatCurrencyPart(t.ILS, AppConstants.Currency_ILS)
            If t.USD > 0 Then result &= "  |  " & UtilityModule.FormatCurrencyPart(t.USD, AppConstants.Currency_USD)
            If t.JOD > 0 Then result &= "  |  " & UtilityModule.FormatCurrencyPart(t.JOD, AppConstants.Currency_JOD)

            lblTotal.Text = result
        Catch ex As Exception
            DatabaseModule.LogError("CalculateTotals", ex)
        End Try
    End Sub

#End Region

#Region "Printing & Export"

    Private Sub ExportToExcel()
        If dgvReport.Rows.Count = 0 Then
            MessageBox.Show("لا توجد بيانات للتصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Using sfd As New SaveFileDialog()
                sfd.Filter = "CSV Files|*.csv"
                sfd.DefaultExt = "csv"
                sfd.FileName = "تقرير_الإيرادات_" & dtpFrom.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) &
                               "_إلى_" & dtpTo.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)

                If sfd.ShowDialog() = DialogResult.OK Then
                    Cursor = Cursors.WaitCursor
                    Using sw As New StreamWriter(sfd.FileName, False, New UTF8Encoding(True))
                        sw.WriteLine("تقرير الإيرادات والمقبوضات")
                        sw.WriteLine("الفترة: " & dtpFrom.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) &
                                     " إلى " & dtpTo.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture))
                        sw.WriteLine("تاريخ التصدير: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture))
                        sw.WriteLine()

                        ' 🌟 توحيد CSV: المنطق المشترك في UtilityModule.WriteGridCsv (كان منسوخاً في 5 نماذج)
                        ' + تهيئة التاريخ والمبلغ بثقافة ثابتة داخل الدالة المشتركة
                        ' 🌟 مصفوفات معرّفة بلا أقواس معقوفة لتكون آمنة عند النسخ
                        Dim excludedColumn(0) As String
                        excludedColumn(0) = "SeqColumn"
                        Dim dateColumns(1) As String
                        dateColumns(0) = AppConstants.Col_Pay_PaymentDate
                        dateColumns(1) = AppConstants.Col_Pay_DueDate
                        Dim amountColumns(0) As String
                        amountColumns(0) = AppConstants.Col_Pay_Amount
                        UtilityModule.WriteGridCsv(dgvReport, sw, excludedColumn, Nothing, dateColumns, amountColumns)

                        sw.WriteLine()
                        sw.WriteLine("الإجماليات:")
                        ' 🌟 إصلاح: سطر الإجماليات كان يُكتب خاماً — فواصل الآلاف تنزاح أعمدة Excel
                        sw.WriteLine("الإجماليات: " & UtilityModule.EscapeCsvValue(lblTotal.Text))
                    End Using

                    Cursor = Cursors.Default
                    MessageBox.Show("تم تصدير التقرير بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    If MessageBox.Show("هل تريد فتح الملف الآن؟", "فتح الملف", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        ' 🌟 إصلاح: فشل الفتح بعد تصدير ناجح كان يعرض "خطأ في التصدير" مضللاً
                        Try
                            Dim psi As New ProcessStartInfo(sfd.FileName)
                            psi.UseShellExecute = True
                            Process.Start(psi)
                        Catch
                            MessageBox.Show("تعذّر فتح الملف تلقائياً. يمكنك فتحه يدوياً من: " & sfd.FileName)
                        End Try
                    End If
                End If
            End Using
        Catch ex As Exception
            Cursor = Cursors.Default
            DatabaseModule.LogError("ExportToExcel", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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
                         OrderBy(Function(r) UtilityModule.SafeDate(r.Cells(AppConstants.Col_Pay_PaymentDate).Value)).
                         ToList()

            ' 🌟 استبدال تهيئة الكائن بالأقواس المعقوفة بتعيين صريح — آمن تماماً عند النسخ
            Dim helper As New PrintHelper()
            helper.SourceGrid = dgvReport
            helper.RowsToPrint = _printRows
            helper.ColumnsToPrint = printColumns
            helper.ReportTitle = "تقرير الإيرادات والمقبوضات"
            helper.FooterSummary = lblTotal.Text
            Try
                helper.ShowPreview()
            Finally
                helper.Dispose()
            End Try

        Catch ex As Exception
            DatabaseModule.LogError("PrintReportToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ أثناء الطباعة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

End Class
