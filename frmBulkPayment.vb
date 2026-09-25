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
''' سند استلام شامل لعدة لاعبين: سند واحد + عدة أسطر
''' كل سطر: لاعب + مبلغ + عملة خاصة به (نفس الشخص مرتين بعملتين = مسموح)
''' الإجماليات والمتبقي لكل عملة على حدة
''' 🌟 صورة السند تُحفظ مع كل الأسطر (أي سطر يُفتح لاحقاً يظهر السند)
''' حفظ ذري بمعاملة واحدة
''' </summary>
Public Class frmBulkPayment

    Private _service As New PaymentService()

    ' سطر = (اسم اللاعب، PlayerId، المبلغ، العملة)
    Private _lines As New List(Of Tuple(Of String, Long, Decimal, String))
    ' 🌟 حماية من الحفظ المزدوج (الضغط مرتين أثناء الحفظ كان يحفظ نفس السند مرتين)
    Private _isSaving As Boolean = False

    Private _playerNames As New List(Of String)
    Private _playerIds As New Dictionary(Of String, Long)
    ' 🌟 خريطة الاسم المعروض → الاسم الحقيقي: الأسماء المكررة تُعرض بلاحقة (هوية) لتمييزها،
    ' والخريطة تُستخدم لاستعادة الاسم الأصلي عند الحفظ حتى لا يتلوّث عمود الاسم بالقاعدة
    Private _playerRealNames As New Dictionary(Of String, String)
    Private _voucherImage As Image = Nothing   ' 🌟 صورة السند المشتركة

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

    Private ReadOnly lblPlayer As New Label()
    Private ReadOnly cmbPlayerName As New ComboBox()
    Private ReadOnly lblAmount As New Label()
    Private ReadOnly txtAmount As New TextBox()
    Private ReadOnly lblLineCurrency As New Label()
    Private ReadOnly cmbLineCurrency As New ComboBox()

    ' 🌟 صورة السند
    Private ReadOnly grpImage As New GroupBox()
    Private ReadOnly pbVoucherImage As New PictureBox()
    Private ReadOnly btnPickImage As New Button()
    Private ReadOnly btnClearImage As New Button()

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

    ' 🌟 لوحة المجاميع المشتركة (منطق موحّد مع frmBulkExpense في UtilityModule)
    Private ReadOnly _totals As New BulkTotalsPanel()

    Public Sub New()
        SetupUI()
    End Sub

    Private Sub SetupUI()
        Me.Text = "سند استلام شامل لعدة لاعبين"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(740, 880)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.RightToLeft = RightToLeft.Yes
        Me.Font = New Font("Arial", 9)

        Dim y As Integer = 15

        ' ── صف 1: رقم السند الشامل | التاريخ ──
        lblVoucher.Text = "رقم السند الشامل:" : lblVoucher.AutoSize = True : lblVoucher.Location = New Point(570, y + 3)
        txtVoucher.Location = New Point(435, y) : txtVoucher.Width = 130
        Me.Controls.AddRange({lblVoucher, txtVoucher})

        lblDate.Text = "التاريخ:" : lblDate.AutoSize = True : lblDate.Location = New Point(315, y + 3)
        dtpDate.Location = New Point(200, y) : dtpDate.Width = 110 : dtpDate.Format = DateTimePickerFormat.Short
        Me.Controls.AddRange({lblDate, dtpDate})

        y += 36

        ' ── صف 2: 🌟 الإجماليات لكل عملة ──
        lblTotalILS.Text = "إجمالي شيكل:" : lblTotalILS.AutoSize = True : lblTotalILS.Location = New Point(620, y + 3)
        txtTotalILS.Location = New Point(515, y) : txtTotalILS.Width = 100
        Me.Controls.AddRange({lblTotalILS, txtTotalILS})

        lblTotalUSD.Text = "دولار:" : lblTotalUSD.AutoSize = True : lblTotalUSD.Location = New Point(430, y + 3)
        txtTotalUSD.Location = New Point(340, y) : txtTotalUSD.Width = 85
        Me.Controls.AddRange({lblTotalUSD, txtTotalUSD})

        lblTotalJOD.Text = "دينار:" : lblTotalJOD.AutoSize = True : lblTotalJOD.Location = New Point(255, y + 3)
        txtTotalJOD.Location = New Point(165, y) : txtTotalJOD.Width = 85
        Me.Controls.AddRange({lblTotalJOD, txtTotalJOD})

        y += 36

        ' ── صف 3: اللاعب | المبلغ | عملة السطر ──
        lblPlayer.Text = "اللاعب:" : lblPlayer.AutoSize = True : lblPlayer.Location = New Point(640, y + 3)
        cmbPlayerName.Location = New Point(420, y) : cmbPlayerName.Width = 215
        cmbPlayerName.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        cmbPlayerName.AutoCompleteSource = AutoCompleteSource.ListItems
        Me.Controls.AddRange({lblPlayer, cmbPlayerName})

        lblAmount.Text = "المبلغ:" : lblAmount.AutoSize = True : lblAmount.Location = New Point(345, y + 3)
        txtAmount.Location = New Point(250, y) : txtAmount.Width = 90
        Me.Controls.AddRange({lblAmount, txtAmount})

        lblLineCurrency.Text = "العملة:" : lblLineCurrency.AutoSize = True : lblLineCurrency.Location = New Point(180, y + 3)
        cmbLineCurrency.Location = New Point(85, y) : cmbLineCurrency.Width = 90 : cmbLineCurrency.DropDownStyle = ComboBoxStyle.DropDownList
        cmbLineCurrency.Items.AddRange(New Object() {AppConstants.Currency_ILS, AppConstants.Currency_USD, AppConstants.Currency_JOD})
        cmbLineCurrency.SelectedIndex = 0
        Me.Controls.AddRange({lblLineCurrency, cmbLineCurrency})

        y += 36

        ' ── 🌟 صورة السند (GroupBox) ──
        grpImage.Text = "صورة السند (تُحفظ مع كل الأسطر)"
        grpImage.Location = New Point(15, y)
        grpImage.Size = New Size(690, 150)
        Me.Controls.Add(grpImage)

        pbVoucherImage.Location = New Point(430, 22)
        pbVoucherImage.Size = New Size(245, 115)
        pbVoucherImage.SizeMode = PictureBoxSizeMode.Zoom
        pbVoucherImage.BorderStyle = BorderStyle.FixedSingle
        pbVoucherImage.BackColor = Color.WhiteSmoke
        grpImage.Controls.Add(pbVoucherImage)

        btnPickImage.Text = "📎 اختيار صورة السند"
        btnPickImage.Size = New Size(180, 34)
        btnPickImage.Location = New Point(225, 30)
        btnPickImage.BackColor = Color.FromArgb(173, 216, 230)
        btnPickImage.FlatStyle = FlatStyle.Flat
        grpImage.Controls.Add(btnPickImage)

        btnClearImage.Text = "🗑️ حذف الصورة"
        btnClearImage.Size = New Size(180, 34)
        btnClearImage.Location = New Point(225, 75)
        btnClearImage.BackColor = Color.FromArgb(250, 200, 180)
        btnClearImage.FlatStyle = FlatStyle.Flat
        btnClearImage.Enabled = False
        grpImage.Controls.Add(btnClearImage)

        Dim lblImgHint As New Label() With {
            .Text = "يُفضّل صورة واضحة للسند الورقي" & vbCrLf & "ستُضغط تلقائياً (500px)",
            .AutoSize = True,
            .Location = New Point(20, 45),
            .ForeColor = Color.Gray
        }
        grpImage.Controls.Add(lblImgHint)

        y += 158

        ' ── 🌟 DataGridView الأسطر ──
        SetupLinesGrid()
        dgvLines.Location = New Point(15, y)
        dgvLines.Size = New Size(695, 220)
        dgvLines.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(dgvLines)

        y += 228

        ' ── صف الأزرار + المجاميع ──
        btnRemove.Text = "🗑️ حذف المحدد" : btnRemove.Size = New Size(130, 36)
        btnRemove.Location = New Point(15, y)
        btnRemove.BackColor = Color.FromArgb(250, 180, 180) : btnRemove.FlatStyle = FlatStyle.Flat
        Me.Controls.Add(btnRemove)

        btnAdd.Text = "➕ إضافة"
        btnAdd.Size = New Size(100, 36) : btnAdd.Location = New Point(155, y)
        btnAdd.BackColor = Color.FromArgb(173, 216, 230) : btnAdd.FlatStyle = FlatStyle.Flat
        btnAdd.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(btnAdd)

        ' المجموع المُدخل لكل عملة
        lblRunningILS.Text = "شيكل مُدخل:" : lblRunningILS.AutoSize = True
        lblRunningILS.Location = New Point(610, y)
        lblRunningILS.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRunningILS)

        lblRunningILSValue.Text = "0" : lblRunningILSValue.AutoSize = True
        lblRunningILSValue.Location = New Point(520, y)
        lblRunningILSValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRunningILSValue)

        lblRunningUSD.Text = "دولار مُدخل:" : lblRunningUSD.AutoSize = True
        lblRunningUSD.Location = New Point(610, y + 22)
        lblRunningUSD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRunningUSD)

        lblRunningUSDValue.Text = "0" : lblRunningUSDValue.AutoSize = True
        lblRunningUSDValue.Location = New Point(520, y + 22)
        lblRunningUSDValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRunningUSDValue)

        lblRunningJOD.Text = "دينار مُدخل:" : lblRunningJOD.AutoSize = True
        lblRunningJOD.Location = New Point(610, y + 44)
        lblRunningJOD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRunningJOD)

        lblRunningJODValue.Text = "0" : lblRunningJODValue.AutoSize = True
        lblRunningJODValue.Location = New Point(520, y + 44)
        lblRunningJODValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRunningJODValue)

        ' المتبقي لكل عملة
        lblRemainingILS.Text = "شيكل متبقي:" : lblRemainingILS.AutoSize = True
        lblRemainingILS.Location = New Point(430, y)
        lblRemainingILS.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRemainingILS)

        lblRemainingILSValue.Text = "—" : lblRemainingILSValue.AutoSize = True
        lblRemainingILSValue.Location = New Point(330, y)
        lblRemainingILSValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRemainingILSValue)

        lblRemainingUSD.Text = "دولار متبقي:" : lblRemainingUSD.AutoSize = True
        lblRemainingUSD.Location = New Point(430, y + 22)
        lblRemainingUSD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRemainingUSD)

        lblRemainingUSDValue.Text = "—" : lblRemainingUSDValue.AutoSize = True
        lblRemainingUSDValue.Location = New Point(330, y + 22)
        lblRemainingUSDValue.Font = New Font("Arial", 9.5!, FontStyle.Bold)
        Me.Controls.Add(lblRemainingUSDValue)

        lblRemainingJOD.Text = "دينار متبقي:" : lblRemainingJOD.AutoSize = True
        lblRemainingJOD.Location = New Point(430, y + 44)
        lblRemainingJOD.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblRemainingJOD)

        lblRemainingJODValue.Text = "—" : lblRemainingJODValue.AutoSize = True
        lblRemainingJODValue.Location = New Point(330, y + 44)
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
        AddHandler btnPickImage.Click, AddressOf PickImage
        AddHandler btnClearImage.Click, AddressOf ClearImage

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
    ' 🌟 DataGridView (ت | سند | لاعب | تاريخ | مبلغ | عملة)
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

        Dim colPlayer As New DataGridViewTextBoxColumn() With {
            .Name = "PlayerName", .HeaderText = "اللاعب",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, .ReadOnly = True
        }
        colPlayer.DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleLeft}

        Dim colDate As New DataGridViewTextBoxColumn() With {
            .Name = "PaymentDate", .HeaderText = "التاريخ",
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

        dgvLines.Columns.AddRange(New DataGridViewColumn() {colSeq, colVoucher, colPlayer, colDate, colAmount, colCurrency})

        ' 🌟 إصلاح: منع الفرز بالنقر على رؤوس الأعمدة — الفرز يخلي CurrentRow.Index (ترتيب العرض)
        ' لا يطابق فهرس _lines (ترتيب الإدخال) فيُحذف سطر خطأ بصمت
        For Each c As DataGridViewColumn In dgvLines.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    ''' <summary>إعادة رسم الجريد</summary>
    Private Sub RefreshLinesGrid()
        Dim voucher As String = txtVoucher.Text.Trim()
        Dim dateStr As String = dtpDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)

        dgvLines.Rows.Clear()
        For i As Integer = 0 To _lines.Count - 1
            ' 🌟 إصلاح: كان عمود المبلغ يستقبل Item2 (رقم الهوية!) بدل Item3 (المبلغ)
            dgvLines.Rows.Add((i + 1).ToString(), voucher, _lines(i).Item1, dateStr, _lines(i).Item3, _lines(i).Item4)
        Next
    End Sub

    ' ─────────────────────────────────────────────
    ' الحسابات لكل عملة
    ' ─────────────────────────────────────────────
    Private Function GetEnteredSum(currency As String) As Decimal
        Return _lines.Where(Function(l) l.Item4 = currency).Sum(Function(l) l.Item3)
    End Function

    Private Function GetDeclaredTotal(currency As String, ByRef parsed As Decimal) As Boolean
        ' 🌟 المنطق المشترك في BulkTotalsPanel (كان منسوخاً حرفياً بين النموذجين)
        Return _totals.ParseDeclared(currency, parsed)
    End Function

    Private Sub UpdateTotals()
        ' 🌟 المنطق المشترك في BulkTotalsPanel (كان منسوخاً حرفياً بين النموذجين)
        _totals.Refresh(AddressOf GetEnteredSum)
    End Sub

    ' ─────────────────────────────────────────────
    ' 🌟 صورة السند
    ' ─────────────────────────────────────────────
    Private Sub PickImage(sender As Object, e As EventArgs)
        Try
            Using ofd As New OpenFileDialog()
                ofd.Title = "اختر صورة السند"
                ofd.Filter = "ملفات الصور|*.jpg;*.jpeg;*.png;*.bmp|كل الملفات|*.*"
                If ofd.ShowDialog() = DialogResult.OK Then
                    Using fs As New IO.FileStream(ofd.FileName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                        Using img As Image = Image.FromStream(fs)
                            SafeDisposeVoucherImage()
                            _voucherImage = New Bitmap(img)
                        End Using
                    End Using
                    pbVoucherImage.Image = _voucherImage
                    btnClearImage.Enabled = True
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkPayment.PickImage", ex)
            MessageBox.Show("خطأ في تحميل الصورة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ClearImage(sender As Object, e As EventArgs)
        SafeDisposeVoucherImage()
        btnClearImage.Enabled = False
    End Sub

    Private Sub SafeDisposeVoucherImage()
        If _voucherImage IsNot Nothing Then
            _voucherImage.Dispose()
            _voucherImage = Nothing
        End If
        pbVoucherImage.Image = Nothing
    End Sub

    ' ─────────────────────────────────────────────
    ' تحميل النموذج (اللاعبون النشطون)
    ' ─────────────────────────────────────────────
    Private Async Sub frmBulkPayment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim result As Tuple(Of DataTable, DataTable) = Await _service.LoadPlayersComboDataAsync()
            If result.Item1 IsNot Nothing Then
                For Each row As DataRow In result.Item1.Rows
                    Dim name As String = UtilityModule.SafeString(row(AppConstants.Col_PlayerName))
                    Dim pid As Long = 0
                    If name <> "" AndAlso Long.TryParse(UtilityModule.SafeString(row(AppConstants.Col_PlayerId)), pid) AndAlso pid > 0 Then
                        If Not _playerIds.ContainsKey(name) Then
                            _playerNames.Add(name)
                            _playerIds(name) = pid
                            _playerRealNames(name) = name
                        Else
                            ' 🌟 إصلاح دقة البيانات: كان الاسم المكرر الثاني يُتجاهل بصمت،
                            ' فالدفعة تُحسب للاعب خاطئ. الآن تُعرض نسخة بلاحقة (رقم الهوية)
                            Dim displayName As String = $"{name} ({pid})"
                            If Not _playerIds.ContainsKey(displayName) Then
                                _playerNames.Add(displayName)
                                _playerIds(displayName) = pid
                                _playerRealNames(displayName) = name
                            End If
                        End If
                    End If
                Next
            End If
            _playerNames.Sort()
            cmbPlayerName.Items.AddRange(_playerNames.ToArray())

            dtpDate.Value = Date.Today
            RefreshLinesGrid()
            UpdateTotals()
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkPayment_Load", ex)
        End Try
    End Sub

    ' ─────────────────────────────────────────────
    ' فحص رقم السند عند المغادرة (رقم شامل = فريد نهائياً)
    ' ─────────────────────────────────────────────
    Private Async Sub TxtVoucher_Leave(sender As Object, e As EventArgs)
        Try
            If String.IsNullOrWhiteSpace(txtVoucher.Text) Then Return
            Dim v As String = txtVoucher.Text.Trim()

            If Await _service.IsVoucherNumberUsedAsync(v) Then
                MessageBox.Show($"رقم السند {v} مستخدم مسبقاً!" & vbCrLf &
                                "السند الشامل يجب أن يكون برقم فريد — الرجاء استخدام رقم مختلف.",
                                "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtVoucher.Focus()
                txtVoucher.SelectAll()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkPayment.TxtVoucher_Leave", ex)
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
            Dim name As String = cmbPlayerName.Text.Trim()
            If String.IsNullOrEmpty(name) OrElse Not _playerIds.ContainsKey(name) Then
                MessageBox.Show("اختر لاعباً من القائمة", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbPlayerName.Focus()
                Return
            End If

            ' 🌟 منع تكرار (نفس اللاعب + نفس العملة) بالنافذة
            Dim currency As String = If(cmbLineCurrency.SelectedItem?.ToString(), "")
            ' 🌟 استعادة الاسم الحقيقي (قد يكون المعروض بلاحقة هوية لتمييز الأسماء المكررة)
            Dim realName As String = If(_playerRealNames.ContainsKey(name), _playerRealNames(name), name)
            ' 🌟 إصلاح: المقارنة كانت بالاسم المعروض — لاعبان متطابقا الاسم (وهذا سبب البلاحقة أصلاً)
            ' كان أحدهما يُمنع ظلماً. الآن المقارنة برقم الهوية
            If _lines.Any(Function(l) l.Item2 = _playerIds(name) AndAlso l.Item4 = currency) Then
                MessageBox.Show($"اللاعب مضاف مسبقاً بعملة {currency}!" & vbCrLf &
                                "نفس اللاعب يُضاف مرة بكل عملة (مثلاً مرة شيكل ومرة دولار).",
                                "تكرار", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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

            ' 🌟 حاجز التجاوز لعملة السطر
            Dim declared As Decimal
            If Not GetDeclaredTotal(currency, declared) Then
                ' 🌟 إصلاح: إجمالي معلن غير رقمي كان يتخطى الحاجز صمتاً
                MessageBox.Show($"قيمة [إجمالي {currency}] غير صالحة — أدخل رقماً.", "إجمالي غير صالح",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If declared > 0 Then
                Dim newSum As Decimal = GetEnteredSum(currency) + amount
                If newSum > declared Then
                    MessageBox.Show(
                        $"لا يمكن إضافة هذا المبلغ!" & vbCrLf &
                        $"إجمالي {currency}: {declared:#,##0.##}" & vbCrLf &
                        $"المُدخل من {currency}: {GetEnteredSum(currency):#,##0.##}" & vbCrLf &
                        $"المطلوب إضافته: {amount:#,##0.##}" & vbCrLf &
                        $"الزائد: {(newSum - declared):#,##0.##}",
                        $"تجاوز إجمالي {currency}", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtAmount.Focus()
                    txtAmount.SelectAll()
                    Return
                End If
            End If

            _lines.Add(Tuple.Create(realName, _playerIds(name), amount, currency))
            RefreshLinesGrid()

            txtAmount.Clear()
            cmbPlayerName.Focus()
            UpdateTotals()
        Catch ex As Exception
            DatabaseModule.LogError("frmBulkPayment.AddLine", ex)
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
            DatabaseModule.LogError("frmBulkPayment.RemoveSelected", ex)
        End Try
    End Sub

    ' ─────────────────────────────────────────────
    ' الحفظ (ذرية — الأسطر بعملاتها + الصورة مع كل سطر)
    ' ─────────────────────────────────────────────
    Private Async Sub SaveBulk(sender As Object, e As EventArgs)
        ' 🌟 إصلاح حفظ مزدوج: حماية فورية أعلى الدالة — كان زر الحفظ يُعطّل فقط بعد
        ' الاستعلام والتأكيد، فالضغطة الثانية أثناء الانتظار كانت تحفظ نفس السند مرتين
        If _isSaving Then Return
        ' 🌟 إصلاح صلاحيات: الحفظ كان يتجاوز بوابة CanEdit مباشرة عبر ExecuteTransaction
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

            ' 🌟 رقم السند الشامل فريد نهائياً
            If Await _service.IsVoucherNumberUsedAsync(voucher) Then
                MessageBox.Show($"رقم السند {voucher} مستخدم مسبقاً — الرجاء استخدام رقم مختلف.",
                                "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtVoucher.Focus() : Return
            End If

            If _lines.Count = 0 Then
                MessageBox.Show("أضف لاعباً واحداً على الأقل", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' 🌟 التحقق لكل عملة
            Dim currenciesUsed As List(Of String) = _lines.Select(Function(l) l.Item4).Distinct().ToList()

            For Each cur As String In currenciesUsed
                Dim entered As Decimal = GetEnteredSum(cur)
                Dim declared As Decimal

                If Not GetDeclaredTotal(cur, declared) OrElse declared <= 0 Then
                    MessageBox.Show($"يوجد أسطر بعملة {cur} لكن لم تدخل الإجمالي!" & vbCrLf &
                                    $"أدخل [إجمالي {cur}] في الأعلى.",
                                    "إجمالي مفقود", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If entered > declared Then
                    MessageBox.Show($"مجموع أسطر {cur} ({entered:#,##0.##}) أكبر من الإجمالي ({declared:#,##0.##})!",
                                    "تجاوز", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            Next

            Dim dateStr As String = dtpDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            Dim modified As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            ' 🔴 H-04: صورة السند تُخزن مرة واحدة بالجدول المستقل VoucherImages —
            ' كانت الـ BLOB نفسها تُنسخ مع كل سطر (سند 20 لاعباً = الصورة 20 مرة = 0.6-1.6MB لكل سند!)
            Dim imageParam As Object = UtilityModule.ImageToDBValue(_voucherImage)
            Dim hasImage As Boolean = imageParam IsNot Nothing AndAlso Not IsDBNull(imageParam)

            ' ملخص
            Dim summaryBuilder As New StringBuilder()
            For Each cur As String In currenciesUsed
                Dim lineCount As Integer = _lines.Where(Function(l) l.Item4 = cur).Count()
                summaryBuilder.AppendLine($"• {cur}: {lineCount} لاعب — إجمالي {GetEnteredSum(cur):#,##0.##}")
            Next
            If hasImage Then
                summaryBuilder.AppendLine("• مع صورة السند ✅ (تُخزن مرة واحدة — جدول مستقل)")
            End If

            If MessageBox.Show($"حفظ {_lines.Count} دفعة برقم السند {voucher}؟" & vbCrLf & vbCrLf &
                               summaryBuilder.ToString(),
                               "تأكيد الحفظ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return

            btnSave.Enabled = False

            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))
            For Each line In _lines
                ' 🔴 H-04: الإدراج بلا عمود VoucherImage — الصورة بمكانها الوحيد أدناه
                Dim sql As String =
                    "INSERT INTO " & AppConstants.Table_Payments & " (PlayerId, Playername, PaymentDate, VoucherNumber, Amount, CurrencyType, CreatedDate, LastModifiedBy, LastModifiedDate) " &
                    "VALUES (@pid, @pname, @pdate, @vno, @amt, @cur, @created, @lmb, @lmd)"
                Dim prms As SQLiteParameter() = {
                    New SQLiteParameter("@pid", CObj(line.Item2)),
                    New SQLiteParameter("@pname", line.Item1),
                    New SQLiteParameter("@pdate", dateStr),
                    New SQLiteParameter("@vno", voucher),
                    New SQLiteParameter("@amt", line.Item3),
                    New SQLiteParameter("@cur", line.Item4),
                    New SQLiteParameter("@created", modified),
                    New SQLiteParameter("@lmb", UserSession.CurrentUsername),
                    New SQLiteParameter("@lmd", modified)
                }
                queries.Add(Tuple.Create(sql, prms))
            Next

            ' 🔴 H-04: صورة السند تُخزن مرة واحدة — بنفس المعاملة (ذرية مع الأسطر كلياً)
            If hasImage Then
                queries.Add(Tuple.Create(
                    "INSERT OR REPLACE INTO VoucherImages (VoucherNumber, ImageData, CreatedDate, LastModifiedDate) " &
                    "VALUES (@vno, @img, @cdate, @cdate2)",
                    New SQLiteParameter() {
                        New SQLiteParameter("@vno", voucher),
                        New SQLiteParameter("@img", imageParam),
                        New SQLiteParameter("@cdate", modified),
                        New SQLiteParameter("@cdate2", modified)}))
            End If

            Dim ok As Boolean = Await Task.Run(Function() DatabaseModule.ExecuteTransaction(queries))

            If ok Then
                ' 🌟 إصلاح توثيق: السند الشامل كله كان بلا أي أثر في سجل التدقيق
                Try
                    Dim audit As New AuditService()
                    audit.Log(AuditService.Act_PaymentAdd, "دفعة", voucher,
                              $"سند شامل: {_lines.Count} سطر — " & summaryBuilder.ToString().Replace(vbCr, "").Replace(vbLf, " "))
                Catch
                End Try
                MessageBox.Show($"تم حفظ {_lines.Count} دفعة برقم السند {voucher} بنجاح!",
                                "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("فشل الحفظ — تم التراجع بالكامل.", "خطأ",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                btnSave.Enabled = True
            End If

        Catch ex As Exception
            DatabaseModule.LogError("frmBulkPayment.SaveBulk", ex)
            MessageBox.Show("خطأ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnSave.Enabled = True
        Finally
            _isSaving = False
        End Try
    End Sub

    ' ─────────────────────────────────────────────
    ' تنظيف عند الإغلاق
    ' ─────────────────────────────────────────────
    Private Sub frmBulkPayment_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SafeDisposeVoucherImage()
    End Sub

End Class
