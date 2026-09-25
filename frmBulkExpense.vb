Option Explicit On
Option Strict On

Imports System.Data
Imports System.Data.SQLite
Imports System.Drawing
Imports System.Globalization ' 🌟 لضبط تحليل المبالغ مستقل عن إعدادات الويندوز
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

''' <summary>
''' سند صرف شامل: سند واحد + عدة أسطر، كل سطر بيان + مبلغ + عملة خاصة به
''' السند الواحد قد يحتوي أكثر من عملة (مثلاً: 100 شيكل + 50 دولار)
''' الإجماليات والمتبقي تُحسب لكل عملة على حدة (لا خلط عملات)
''' الأسطر بـ DataGridView ملوّن: ت | سند | بيان | تاريخ | مبلغ | عملة
''' حفظ ذري — الفئة ذكية والبيان AutoComplete
''' </summary>
Public Class frmBulkExpense

    ' سطر = (البيان، المبلغ، العملة)
    Private _lines As New List(Of Tuple(Of String, Decimal, String))
    ' 🌟 حماية من الحفظ المزدوج (الضغط مرتين أثناء الحفظ كان يحفظ نفس السند مرتين)
    Private _isSaving As Boolean = False

    Private ReadOnly lblVoucher As New Label()
    Private ReadOnly txtVoucher As New TextBox()
    Private ReadOnly lblDate As New Label()
    Private ReadOnly dtpDate As New DateTimePicker()

    Private ReadOnly lblTotalILS As New Label()
    Private ReadOnly txtTotalILS As New TextBox()
    Private ReadOnly lblTotalUSD As New Label()
    Private ReadOnly txtTotalUSD As New TextBox()
    Private ReadOnly lblTotalJOD As New Label()
    Private ReadOnly txtTotalJOD As New TextBox()

    Private ReadOnly lblDesc As New Label()
    Private ReadOnly txtDesc As New TextBox()
    Private ReadOnly lblAmount As New Label()
    Private ReadOnly txtAmount As New TextBox()
    Private ReadOnly lblLineCurrency As New Label()
    Private ReadOnly cmbLineCurrency As New ComboBox()
    Private ReadOnly lblCategory As New Label()
    Private ReadOnly cmbCategory As New ComboBox()

    Private ReadOnly dgvLines As New DataGridView()

    Private ReadOnly btnRemove As New Button()
    Private ReadOnly btnAdd As New Button()

    Private ReadOnly lblRunningILS As New Label()
    Private ReadOnly lblRunningILSValue As New Label()
    Private ReadOnly lblRunningUSD As New Label()
    Private ReadOnly lblRunningUSDValue As New Label()
    Private ReadOnly lblRunningJOD As New Label()
    Private ReadOnly lblRunningJODValue As New Label()

    Private ReadOnly lblRemainingILS As New Label()
    Private ReadOnly lblRemainingILSValue As New Label()
    Private ReadOnly lblRemainingUSD As New Label()
    Private ReadOnly lblRemainingUSDValue As New Label()
    Private ReadOnly lblRemainingJOD As New Label()
    Private ReadOnly lblRemainingJODValue As New Label()

    Private ReadOnly btnSave As New Button()
    Private ReadOnly btnCancel As New Button()

    ' 🌟 لوحة المجاميع المشتركة (منطق موحّد مع frmBulkPayment في UtilityModule)
    Private ReadOnly _totals As New BulkTotalsPanel()

    Public Sub New()
        SetupUI()
    End Sub

    Private Sub SetupUI()
        Me.Text = "سند صرف شامل لعدة مستفيدين"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(680, 600)   ' 🌟 كانت 780: أطول من شاشات اللابتوب (768) وفي ~200px فراغ تحت زر الحفظ
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.RightToLeft = RightToLeft.Yes
        Me.Font = New Font("Arial", 9)

        Dim y As Integer = 15

        ' ── صف 1: رقم السند الشامل | التاريخ ──
        lblVoucher.Text = "رقم السند الشامل:" : lblVoucher.AutoSize = True : lblVoucher.Location = New Point(510, y + 3)
        txtVoucher.Location = New Point(375, y) : txtVoucher.Width = 130
        Me.Controls.AddRange({lblVoucher, txtVoucher})

        lblDate.Text = "التاريخ:" : lblDate.AutoSize = True : lblDate.Location = New Point(255, y + 3)
        dtpDate.Location = New Point(140, y) : dtpDate.Width = 110 : dtpDate.Format = DateTimePickerFormat.Short
        Me.Controls.AddRange({lblDate, dtpDate})

        y += 36

        ' ── صف 2: 🌟 الإجماليات لكل عملة ──
        lblTotalILS.Text = "إجمالي شيكل:" : lblTotalILS.AutoSize = True : lblTotalILS.Location = New Point(560, y + 3)
        txtTotalILS.Location = New Point(455, y) : txtTotalILS.Width = 100
        Me.Controls.AddRange({lblTotalILS, txtTotalILS})

        lblTotalUSD.Text = "دولار:" : lblTotalUSD.AutoSize = True : lblTotalUSD.Location = New Point(370, y + 3)
        txtTotalUSD.Location = New Point(280, y) : txtTotalUSD.Width = 85
        Me.Controls.AddRange({lblTotalUSD, txtTotalUSD})

        lblTotalJOD.Text = "دينار:" : lblTotalJOD.AutoSize = True : lblTotalJOD.Location = New Point(195, y + 3)
        txtTotalJOD.Location = New Point(105, y) : txtTotalJOD.Width = 85
        Me.Controls.AddRange({lblTotalJOD, txtTotalJOD})

        y += 36

        ' ── صف 3: البيان | المبلغ | عملة السطر ──
        lblDesc.Text = "البيان:" : lblDesc.AutoSize = True : lblDesc.Location = New Point(590, y + 3)
        txtDesc.Location = New Point(400, y) : txtDesc.Width = 185
        Me.Controls.AddRange({lblDesc, txtDesc})

        lblAmount.Text = "المبلغ:" : lblAmount.AutoSize = True : lblAmount.Location = New Point(330, y + 3)
        txtAmount.Location = New Point(240, y) : txtAmount.Width = 85
        Me.Controls.AddRange({lblAmount, txtAmount})

        lblLineCurrency.Text = "العملة:" : lblLineCurrency.AutoSize = True : lblLineCurrency.Location = New Point(165, y + 3)
        cmbLineCurrency.Location = New Point(70, y) : cmbLineCurrency.Width = 90 : cmbLineCurrency.DropDownStyle = ComboBoxStyle.DropDownList
        cmbLineCurrency.Items.AddRange(New Object() {AppConstants.Currency_ILS, AppConstants.Currency_USD, AppConstants.Currency_JOD})
        cmbLineCurrency.SelectedIndex = 0
        Me.Controls.AddRange({lblLineCurrency, cmbLineCurrency})

        ' الفئة — سطر مستقل تحت العملة
        lblCategory.Text = "الفئة:" : lblCategory.AutoSize = True : lblCategory.Location = New Point(610, y + 35)
        cmbCategory.Location = New Point(500, y + 32) : cmbCategory.Width = 95 : cmbCategory.DropDownStyle = ComboBoxStyle.DropDown
        Me.Controls.AddRange({lblCategory, cmbCategory})

        ' 🌟 Enter للتنقل والإضافة
        AddHandler txtDesc.KeyDown, Sub(s, e)
                                        If e.KeyCode = Keys.Enter Then
                                            e.SuppressKeyPress = True
                                            txtAmount.Focus()
                                        End If
                                    End Sub
        AddHandler txtAmount.KeyDown, Sub(s, e)
                                          If e.KeyCode = Keys.Enter Then
                                              e.SuppressKeyPress = True
                                              cmbLineCurrency.Focus()
                                          End If
                                      End Sub
        AddHandler cmbLineCurrency.KeyDown, Sub(s, e)
                                                If e.KeyCode = Keys.Enter Then
                                                    e.SuppressKeyPress = True
                                                    AddLine()
                                                End If
                                            End Sub

        y += 78

        ' ── 🌟 DataGridView الأسطر ──
        SetupLinesGrid()
        dgvLines.Location = New Point(15, y)
        dgvLines.Size = New Size(635, 230)
        dgvLines.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right  ' 🌟 بدون Bottom حتى لا تنكمش الشبكة مع الارتفاع الجديد
        Me.Controls.Add(dgvLines)

        y += 238

        ' ── صف الأزرار + مجموعات المدخل ──
        btnRemove.Text = "🗑️ حذف المحدد" : btnRemove.Size = New Size(130, 36)
        btnRemove.Location = New Point(15, y)
        btnRemove.BackColor = Color.FromArgb(250, 180, 180) : btnRemove.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnRemove)

        btnAdd.Text = "➕ إضافة"
        btnAdd.Size = New Size(100, 36) : btnAdd.Location = New Point(155, y)
        btnAdd.BackColor = Color.FromArgb(173, 216, 230) : btnAdd.FlatStyle = FlatStyle.Flat
        btnAdd.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(btnAdd)

        ' المجموع المُدخل لكل عملة — يمين
        lblRunningILS.Text = "شيكل مُدخل:" : lblRunningILS.AutoSize = True
        lblRunningILS.Location = New Point(560, y)
        lblRunningILS.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRunningILS)

        lblRunningILSValue.Text = "0" : lblRunningILSValue.AutoSize = True
        lblRunningILSValue.Location = New Point(455, y)
        lblRunningILSValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRunningILSValue)

        lblRunningUSD.Text = "دولار مُدخل:" : lblRunningUSD.AutoSize = True
        lblRunningUSD.Location = New Point(560, y + 22)
        lblRunningUSD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRunningUSD)

        lblRunningUSDValue.Text = "0" : lblRunningUSDValue.AutoSize = True
        lblRunningUSDValue.Location = New Point(455, y + 22)
        lblRunningUSDValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRunningUSDValue)

        lblRunningJOD.Text = "دينار مُدخل:" : lblRunningJOD.AutoSize = True
        lblRunningJOD.Location = New Point(560, y + 44)
        lblRunningJOD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRunningJOD)

        lblRunningJODValue.Text = "0" : lblRunningJODValue.AutoSize = True
        lblRunningJODValue.Location = New Point(455, y + 44)
        lblRunningJODValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRunningJODValue)

        ' ── المتبقي لكل عملة — وسط ──
        lblRemainingILS.Text = "شيكل متبقي:" : lblRemainingILS.AutoSize = True
        lblRemainingILS.Location = New Point(390, y)
        lblRemainingILS.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRemainingILS)

        lblRemainingILSValue.Text = "—" : lblRemainingILSValue.AutoSize = True
        lblRemainingILSValue.Location = New Point(285, y)
        lblRemainingILSValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRemainingILSValue)

        lblRemainingUSD.Text = "دولار متبقي:" : lblRemainingUSD.AutoSize = True
        lblRemainingUSD.Location = New Point(390, y + 22)
        lblRemainingUSD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRemainingUSD)

        lblRemainingUSDValue.Text = "—" : lblRemainingUSDValue.AutoSize = True
        lblRemainingUSDValue.Location = New Point(285, y + 22)
        lblRemainingUSDValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRemainingUSDValue)

        lblRemainingJOD.Text = "دينار متبقي:" : lblRemainingJOD.AutoSize = True
        lblRemainingJOD.Location = New Point(390, y + 44)
        lblRemainingJOD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRemainingJOD)

        lblRemainingJODValue.Text = "—" : lblRemainingJODValue.AutoSize = True
        lblRemainingJODValue.Location = New Point(285, y + 44)
        lblRemainingJODValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRemainingJODValue)

        y += 88

        ' ── 💾 الحفظ ──
        btnSave.Text = "💾 حفظ السند الشامل"
        btnSave.Size = New Size(280, 44) : btnSave.Location = New Point(15, y)
        btnSave.BackColor = Color.FromArgb(100, 210, 130) : btnSave.FlatStyle = FlatStyle.Flat
        btnSave.Font = New Font("Arial", 10, FontStyle.Bold)
        Me.Controls.Add(btnSave)

        btnCancel.Text = "إلغاء"
        btnCancel.Size = New Size(120, 44) : btnCancel.Location = New Point(310, y)
        btnCancel.BackColor = Color.FromArgb(200, 200, 200) : btnCancel.FlatStyle = FlatStyle.Flat
        btnCancel.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnCancel)
        Me.CancelButton = btnCancel

        ' ─────────────────────────────────────────
        ' 🌟 ربط الأحداث
        ' ─────────────────────────────────────────
        AddHandler btnRemove.Click, Sub(s, e) RemoveSelected()
        AddHandler btnAdd.Click, Sub(s, e) AddLine()
        AddHandler btnSave.Click, AddressOf SaveBulk

        AddHandler txtAmount.KeyPress, AddressOf AmountKeyPress
        AddHandler txtTotalILS.KeyPress, AddressOf AmountKeyPress
        AddHandler txtTotalUSD.KeyPress, AddressOf AmountKeyPress
        AddHandler txtTotalJOD.KeyPress, AddressOf AmountKeyPress

        AddHandler txtTotalILS.TextChanged, Sub(s, e) UpdateTotals()
        AddHandler txtTotalUSD.TextChanged, Sub(s, e) UpdateTotals()
        AddHandler txtTotalJOD.TextChanged, Sub(s, e) UpdateTotals()

        ' 🌟 تسجيل عناصر المجاميع باللوحة المشتركة
        _totals.Register(AppConstants.Currency_ILS, txtTotalILS, lblRunningILSValue, lblRemainingILSValue)
        _totals.Register(AppConstants.Currency_USD, txtTotalUSD, lblRunningUSDValue, lblRemainingUSDValue)
        _totals.Register(AppConstants.Currency_JOD, txtTotalJOD, lblRunningJODValue, lblRemainingJODValue)

        AddHandler txtVoucher.TextChanged, Sub(s, e) RefreshLinesGrid()
        AddHandler dtpDate.ValueChanged, Sub(s, e) RefreshLinesGrid()

        AddHandler txtVoucher.Leave, AddressOf TxtVoucher_Leave
    End Sub

    ' ─────────────────────────────────────────────
    ' 🌟 DataGridView (ت | سند | بيان | تاريخ | مبلغ | عملة)
    ' ─────────────────────────────────────────────
    Private Sub SetupLinesGrid()
        dgvLines.ReadOnly = True
        dgvLines.AllowUserToAddRows = False
        dgvLines.AllowUserToDeleteRows = False
        dgvLines.AllowUserToResizeRows = False
        dgvLines.RowHeadersVisible = False
        dgvLines.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvLines.MultiSelect = False
        dgvLines.BackgroundColor = Color.White
        dgvLines.EnableHeadersVisualStyles = False
        dgvLines.GridColor = Color.Gainsboro

        dgvLines.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(64, 64, 64)
        dgvLines.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvLines.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        dgvLines.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvLines.ColumnHeadersHeight = 32

        dgvLines.DefaultCellStyle.Font = New Font("Arial", 9)
        dgvLines.DefaultCellStyle.SelectionBackColor = Color.FromArgb(135, 206, 250)
        dgvLines.DefaultCellStyle.SelectionForeColor = Color.Black
        dgvLines.RowsDefaultCellStyle.BackColor = Color.White
        dgvLines.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255)

        ' 🌟 L-06: فحص القيمة الراجعة قبل الاستخدام — يمنع استثناء نظري في بيئات محدودة
        Dim dblBufProp As Reflection.PropertyInfo = dgvLines.GetType().GetProperty("DoubleBuffered",
            Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        If dblBufProp IsNot Nothing Then
            dblBufProp.SetValue(dgvLines, True, Nothing)
        End If

        Dim colSeq As New DataGridViewTextBoxColumn() With {
            .Name = "SeqColumn", .HeaderText = "ت", .Width = 40, .ReadOnly = True
        }
        colSeq.DefaultCellStyle = New DataGridViewCellStyle() With {
            .Alignment = DataGridViewContentAlignment.MiddleCenter,
            .BackColor = Color.FromArgb(232, 240, 254),
            .SelectionBackColor = Color.FromArgb(232, 240, 254),
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(40, 48, 68)
        }

        Dim colVoucher As New DataGridViewTextBoxColumn() With {
            .Name = "VoucherNumber", .HeaderText = "رقم السند",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, .ReadOnly = True
        }
        colVoucher.DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleCenter}

        Dim colDesc As New DataGridViewTextBoxColumn() With {
            .Name = "Description", .HeaderText = "البيان",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, .ReadOnly = True
        }
        colDesc.DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleLeft}

        Dim colDate As New DataGridViewTextBoxColumn() With {
            .Name = "ExpenseDate", .HeaderText = "التاريخ",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, .ReadOnly = True
        }
        colDate.DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleCenter}

        Dim colAmount As New DataGridViewTextBoxColumn() With {
            .Name = "Amount", .HeaderText = "المبلغ",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, .ReadOnly = True
        }
        colAmount.DefaultCellStyle = New DataGridViewCellStyle() With {
            .Format = "#,##0.##",
            .Alignment = DataGridViewContentAlignment.MiddleCenter,
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(0, 100, 60)
        }

        Dim colCurrency As New DataGridViewTextBoxColumn() With {
            .Name = "Currency", .HeaderText = "العملة",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, .ReadOnly = True
        }
        colCurrency.DefaultCellStyle = New DataGridViewCellStyle() With {
            .Alignment = DataGridViewContentAlignment.MiddleCenter,
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .ForeColor = Color.FromArgb(140, 60, 0)
        }

        dgvLines.Columns.AddRange(New DataGridViewColumn() {colSeq, colVoucher, colDesc, colDate, colAmount, colCurrency})

        ' 🌟 إصلاح: منع الفرز — الفرز يخلي CurrentRow.Index لا يطابق فهرس _lines فيُحذف سطر خطأ
        For Each c As DataGridViewColumn In dgvLines.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    ''' <summary>إعادة رسم الجريد — تسلسل + سند وتاريخ مشتركين + عملة كل سطر</summary>
    Private Sub RefreshLinesGrid()
        Dim voucher As String = txtVoucher.Text.Trim()
        Dim dateStr As String = dtpDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)

        dgvLines.Rows.Clear()
        For i As Integer = 0 To _lines.Count - 1
            dgvLines.Rows.Add((i + 1).ToString(), voucher, _lines(i).Item1, dateStr, _lines(i).Item2, _lines(i).Item3)
        Next
    End Sub

    ' ─────────────────────────────────────────────
    ' الحسابات لكل عملة على حدة
    ' ─────────────────────────────────────────────
    Private Function GetEnteredSum(currency As String) As Decimal
        Return _lines.Where(Function(l) l.Item3 = currency).Sum(Function(l) l.Item2)
    End Function

    Private Function GetDeclaredTotal(currency As String, ByRef parsed As Decimal) As Boolean
        ' 🌟 المنطق المشترك في BulkTotalsPanel (كان منسوخاً حرفياً بين النموذجين)
        Return _totals.ParseDeclared(currency, parsed)
    End Function

    ''' <summary>تحديث المدخل والمتبقي لكل عملة — بألوان حسب الحالة</summary>
    Private Sub UpdateTotals()
        ' 🌟 المنطق المشترك في BulkTotalsPanel (كان منسوخاً حرفياً بين النموذجين)
        _totals.Refresh(AddressOf GetEnteredSum)
    End Sub

    ' ─────────────────────────────────────────────
    ' تحميل النموذج
    ' ─────────────────────────────────────────────
    Private Async Sub frmBulkExpense_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim srv As New DropdownService()
            Dim saved As List(Of String) = Await srv.GetItemsAsync("ExpenseCategory")
            For Each v In saved
                If Not cmbCategory.Items.Contains(v) Then cmbCategory.Items.Add(v)
            Next

            Dim realDt As DataTable = Await DatabaseModule.ExecuteQueryAsync(
                "SELECT DISTINCT TRIM(Category) AS C FROM Expenses WHERE Category IS NOT NULL AND TRIM(Category) <> '' ORDER BY C")
            If realDt IsNot Nothing Then
                For Each row As DataRow In realDt.Rows
                    Dim v As String = UtilityModule.SafeString(row("C"))
                    If v <> "" AndAlso Not cmbCategory.Items.Contains(v) Then cmbCategory.Items.Add(v)
                Next
            End If
            If cmbCategory.Items.Count > 0 Then cmbCategory.SelectedIndex = 0

            Dim descDt As DataTable = Await DatabaseModule.ExecuteQueryAsync(
                "SELECT DISTINCT TRIM(Description) AS D FROM Expenses WHERE Description IS NOT NULL AND TRIM(Description) <> '' ORDER BY D")
            If descDt IsNot Nothing Then
                Dim src As New AutoCompleteStringCollection()
                For Each row As DataRow In descDt.Rows
                    Dim v As String = UtilityModule.SafeString(row("D"))
                    If v <> "" Then src.Add(v)
                Next
                txtDesc.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                txtDesc.AutoCompleteSource = AutoCompleteSource.CustomSource
                txtDesc.AutoCompleteCustomSource = src
            End If

            dtpDate.Value = Date.Today
            RefreshLinesGrid()
            UpdateTotals()
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkExpense_Load", ex)
        End Try
    End Sub

    ' ─────────────────────────────────────────────
    ' فحص رقم السند عند المغادرة
    ' ─────────────────────────────────────────────
    Private Async Sub TxtVoucher_Leave(sender As Object, e As EventArgs)
        Try
            If String.IsNullOrWhiteSpace(txtVoucher.Text) Then Return
            Dim v As String = txtVoucher.Text.Trim()

            Dim service As New ExpenseService()
            If Await service.IsVoucherNumberExistsAsync(v) Then
                MessageBox.Show($"رقم السند {v} مستخدم مسبقاً!" & vbCrLf &
                                "السند الشامل يجب أن يكون برقم فريد — الرجاء استخدام رقم مختلف.",
                                "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtVoucher.Focus()
                txtVoucher.SelectAll()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkExpense.TxtVoucher_Leave", ex)
        End Try
    End Sub

    ' ─────────────────────────────────────────────
    ' الإدخال
    ' ─────────────────────────────────────────────
    Private Sub AmountKeyPress(sender As Object, e As KeyPressEventArgs)
        ' 🌟 الفلتر المشترك في UtilityModule (كان منسوخاً حرفياً بين النموذجين)
        UtilityModule.AmountKeyPressFilter(sender, e)
    End Sub

    Private Sub AddLine()
        Try
            Dim desc As String = txtDesc.Text.Trim()
            If String.IsNullOrEmpty(desc) Then
                MessageBox.Show("أدخل البيان (اسم المستفيد / وصف المصروف)", "تنبيه",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtDesc.Focus()
                Return
            End If

            Dim amount As Decimal
            ' 🌟 تحليل المبلغ بثقافة ثابتة (مطابق لفلتر KeyPress الذي يفرض النقطة)
            If Not Decimal.TryParse(txtAmount.Text, NumberStyles.Number, CultureInfo.InvariantCulture, amount) OrElse amount <= 0 Then
                MessageBox.Show("أدخل مبلغاً صحيحاً أكبر من صفر", "تنبيه",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Return
            End If

            Dim currency As String = If(cmbLineCurrency.SelectedItem?.ToString(), "")
            If String.IsNullOrEmpty(currency) Then
                MessageBox.Show("اختر عملة السطر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbLineCurrency.Focus()
                Return
            End If

            ' 🌟 حاجز التجاوز — لعملة هذا السطر فقط
            Dim declared As Decimal
            If GetDeclaredTotal(currency, declared) AndAlso declared > 0 Then
                Dim newSum As Decimal = GetEnteredSum(currency) + amount
                If newSum > declared Then
                    MessageBox.Show(
                        $"لا يمكن إضافة هذا المبلغ!" & vbCrLf &
                        $"إجمالي {currency}: {declared:#,##0.##}" & vbCrLf &
                        $"المُدخل من {currency}: {GetEnteredSum(currency):#,##0.##}" & vbCrLf &
                        $"المطلوب إضافته: {amount:#,##0.##}" & vbCrLf &
                        $"الزائد: {(newSum - declared):#,##0.##}" & vbCrLf & vbCrLf &
                        "عدّل المبلغ ليطابق المتبقي المتاح لهذه العملة.",
                        $"تجاوز إجمالي {currency}", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtAmount.Focus()
                    txtAmount.SelectAll()
                    Return
                End If
            End If

            _lines.Add(Tuple.Create(desc, amount, currency))
            RefreshLinesGrid()

            txtDesc.Clear() : txtAmount.Clear()
            txtDesc.Focus()
            UpdateTotals()
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkExpense.AddLine", ex)
        End Try
    End Sub

    Private Sub RemoveSelected()
        Try
            If dgvLines.CurrentRow Is Nothing OrElse dgvLines.CurrentRow.Index < 0 Then
                MessageBox.Show("حدد سطراً من الجدول أولاً", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim idx As Integer = dgvLines.CurrentRow.Index
            If idx >= 0 AndAlso idx < _lines.Count Then
                _lines.RemoveAt(idx)
                RefreshLinesGrid()
                UpdateTotals()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkExpense.RemoveSelected", ex)
        End Try
    End Sub

    ' ─────────────────────────────────────────────
    ' الحفظ (ذرية — كل الأسطر بعملاتها)
    ' ─────────────────────────────────────────────
    Private Async Sub SaveBulk(sender As Object, e As EventArgs)
        ' 🌟 إصلاح حفظ مزدوج: حماية فورية أعلى الدالة — كان زر الحفظ يُعطّل فقط بعد
        ' الاستعلام والتأكيد، فالضغطة الثانية أثناء الانتظار كانت تحفظ نفس السند مرتين
        If _isSaving Then Return
        ' 🌟 إصلاح صلاحيات: الحفظ كان يتجاوز بوابة CanEdit (ExpenseService.InsertExpenseAsync) مباشرة
        If Not UserSession.CanEdit Then
            MessageBox.Show("ليست لديك صلاحية الإدخال/التعديل", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        _isSaving = True
        Try
            Dim voucher As String = txtVoucher.Text.Trim()

            If String.IsNullOrEmpty(voucher) Then
                MessageBox.Show("أدخل رقم السند الشامل", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtVoucher.Focus() : Return
            End If

            Dim svcCheck As New ExpenseService()
            If Await svcCheck.IsVoucherNumberExistsAsync(voucher) Then
                MessageBox.Show($"رقم السند {voucher} مستخدم مسبقاً — الرجاء استخدام رقم مختلف.",
                                "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtVoucher.Focus() : Return
            End If

            If _lines.Count = 0 Then
                MessageBox.Show("أضف سطراً واحداً على الأقل", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' 🌟 التحقق لكل عملة مستخدمة: إجمالي معلن + لا تجاوز
            Dim currenciesUsed As List(Of String) = _lines.Select(Function(l) l.Item3).Distinct().ToList()

            For Each cur As String In currenciesUsed
                Dim entered As Decimal = GetEnteredSum(cur)
                Dim declared As Decimal

                If Not GetDeclaredTotal(cur, declared) OrElse declared <= 0 Then
                    MessageBox.Show($"يوجد أسطر بعملة {cur} لكن لم تدخل الإجمالي للسند!" & vbCrLf &
                                    $"أدخل [إجمالي {cur}] في الأعلى.",
                                    "إجمالي مفقود", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If entered > declared Then
                    MessageBox.Show($"مجموع أسطر {cur} ({entered:#,##0.##}) أكبر من الإجمالي المعلن ({declared:#,##0.##})!",
                                    "تجاوز", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            Next

            ' إجمالي معلن بدون أسطر → تنبيه
            For Each cur As String In New String() {AppConstants.Currency_ILS, AppConstants.Currency_USD, AppConstants.Currency_JOD}
                Dim declared As Decimal
                If GetDeclaredTotal(cur, declared) AndAlso declared > 0 AndAlso GetEnteredSum(cur) = 0 Then
                    Dim res As DialogResult = MessageBox.Show(
                        $"أدخلت إجمالي {cur} ({declared:#,##0.##}) لكن لا يوجد أسطر بهذه العملة!" & vbCrLf &
                        "هل تريد المتابعة؟",
                        "إجمالي بدون أسطر", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button2)
                    If res <> DialogResult.Yes Then Return
                End If
            Next

            Dim category As String = cmbCategory.Text.Trim()
            Dim dateStr As String = dtpDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            Dim modified As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            ' 🌟 ملخص الحفظ لكل عملة
            Dim summaryBuilder As New StringBuilder()
            For Each cur As String In currenciesUsed
                Dim lineCount As Integer = _lines.Where(Function(l) l.Item3 = cur).Count()
                summaryBuilder.AppendLine($"• {cur}: {lineCount} سطر — إجمالي {GetEnteredSum(cur):#,##0.##}")
            Next

            If MessageBox.Show($"حفظ {_lines.Count} سطر مصروف برقم السند {voucher}؟" & vbCrLf & vbCrLf &
                               summaryBuilder.ToString(),
                               "تأكيد الحفظ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return

            btnSave.Enabled = False

            ' 🌟 كل سطر بعملته — بمعاملة واحدة
            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))
            For Each line In _lines
                Dim sql As String =
                    "INSERT INTO Expenses (VoucherNumber, Description, Amount, CurrencyType, ExpenseDate, Category, LastModifiedBy, LastModifiedDate) " &
                    "VALUES (@vno, @desc, @amt, @cur, @edate, @cat, @lmb, @lmd)"
                Dim prms As SQLiteParameter() = {
                    New SQLiteParameter("@vno", voucher),
                    New SQLiteParameter("@desc", line.Item1),
                    New SQLiteParameter("@amt", line.Item2),
                    New SQLiteParameter("@cur", line.Item3),
                    New SQLiteParameter("@edate", dateStr),
                    New SQLiteParameter("@cat", category),
                    New SQLiteParameter("@lmb", UserSession.CurrentUsername),
                    New SQLiteParameter("@lmd", modified)
                }
                queries.Add(Tuple.Create(sql, prms))
            Next

            Dim ok As Boolean = Await Task.Run(Function() DatabaseModule.ExecuteTransaction(queries))

            If ok Then
                ' 🌟 إصلاح توثيق: سند الصرف الشامل كان بلا أي أثر في سجل التدقيق
                Try
                    Dim audit As New AuditService()
                    audit.Log(AuditService.Act_ExpenseAdd, "مصروف", voucher,
                              $"سند صرف شامل: {_lines.Count} سطر — " & summaryBuilder.ToString().Replace(vbCr, "").Replace(vbLf, " "))
                Catch
                End Try
                If category <> "" Then
                    Await (New DropdownService()).AddItemAsync("ExpenseCategory", category)
                End If

                MessageBox.Show($"تم حفظ {_lines.Count} سطر مصروف برقم السند {voucher} بنجاح!",
                                "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("فشل الحفظ — تم التراجع بالكامل.", "خطأ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                btnSave.Enabled = True
            End If

        Catch ex As Exception
            DatabaseModule.LogError("frmBulkExpense.SaveBulk", ex)
            MessageBox.Show("خطأ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnSave.Enabled = True
        Finally
            _isSaving = False
        End Try
    End Sub

End Class
