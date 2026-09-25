Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Linq
Imports System.Globalization ' 🌟 لضبط تحليل المبالغ بشكل ثابت مستقل عن إعدادات الويندوز

' 🌟 ملاحظة توثيق (أولوية 3): الكلاس "frmPayments" (شاشة إدخال الإيرادات/الدفعات) اسمه سليم،
'    لكن اسم الملف "Form3.vb" غير معبّر — أعد تسمية الملف إلى frmPayments.vb
'    من Visual Studio (كليك يمين على الملف ← Rename) دون أي تأثير على الكود.
Public Class frmPayments

#Region "Variables"

    Private isEditMode As Boolean = False
    Private currentVoucherNumber As String = ""
    Private currentPaymentId As Long = -1   ' 🌟 المعرف الفريد للدفعة الحالية

    Private suppressSearch As Boolean = False
    Private _isSaving As Boolean = False
    Private playerNames As New List(Of String)
    Private playerPhones As New Dictionary(Of String, String)
    Private playerIds As New Dictionary(Of String, Long)   ' 🌟 اسم اللاعب → معرفه الفريد
    ' 🔴 H-03: كل معرفات اللاعبين لكل اسم (مع الهاتف) — التكرار يُحل بحوار صريح لا اختيار صامت لأصغر هوية
    Private playerChoices As New Dictionary(Of String, List(Of Tuple(Of Long, String)))
    Private archivedPlayers As New HashSet(Of String)
    Private TempData As DataTable
    Private printColumns As New List(Of String)
    Private _printRows As List(Of DataGridViewRow)
    Private isReportMode As Boolean = False
    Private totalPaymentsCount As Integer = 0
    Private _cancellationTokenSource As CancellationTokenSource
    Private WithEvents searchTimer As New System.Windows.Forms.Timer()
    Private isClosing As Boolean = False
    Private _imageChanged As Boolean = False
    Private tempImagePath As String = ""
    Private WithEvents imgPrintDoc As New Printing.PrintDocument()
    ' 🌟 عدّاد حماية من سباق تحميل صورة السند (نتيجة قديمة لا تُعرض فوق الجديدة)
    Private _voucherImageLoadId As Integer = 0
    ' 🌟 L-09: حالة الترقيم التزايدي لقائمة الدفعات — لا تُحمَّل كل السطور دفعة واحدة
    Private Const LoadMorePageSize As Integer = 500
    Private _filterPlayerName As String = ""
    Private _filterVoucherNumber As String = ""
    Private _filterTransferorName As String = ""
    Private _filterNotes As String = ""
    Private _filterPaymentMethod As String = ""
    Private _filterCurrencyType As String = ""
    Private _currentTotals As CurrencyTotals
    Private _totalRowsInFilter As Integer = 0
    Private _hasMoreRows As Boolean = False
    Private _isLoadingMore As Boolean = False
    Private miLoadMore As ToolStripMenuItem = Nothing

#End Region

#Region "Double Buffering"

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property

#End Region

#Region "Form Events"

    Private Sub frmPayments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.SetStyle(ControlStyles.DoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)
            Me.UpdateStyles()

            DatabaseModule.EnsurePaymentsTableExists()
            SetupSearchTimer()
            SetupDataGridView()
            SetDefaultComboBoxValues()
            SetAddMode()
            DisablePaymentControls()
            Me.Text = "شاشة ادخال الايرادات"

            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee

            SaveToolStripMenuItem.Enabled = UserSession.CanEdit
            DeleteToolStripMenuItem.Enabled = UserSession.CanDelete

            ' 🌟 [أولوية 4] بند تصدير HTML (طباعة/PDF) داخل قائمة التصدير الموجودة
            AddHtmlExportItem()
            ' 🌟 L-09: بند "تحميل المزيد" بشريط القوائم (ترقيم تزايدي لقائمة الدفعات)
            AddLoadMoreMenuItem()

        Catch ex As Exception
            DatabaseModule.LogError("frmPayments_Load", ex)
            MessageBox.Show("خطأ في تحميل النموذج: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub frmPayments_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            Await Task.WhenAll(LoadAllPlayersComboAsync(), LoadPaymentsAsync())
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments_Shown", ex)
        End Try
    End Sub

    Private Async Function LoadAllPlayersComboAsync(Optional updateCombo As Boolean = True) As Task
        Try
            Dim service As New PaymentService()
            Dim result As Tuple(Of DataTable, DataTable) = Await service.LoadPlayersComboDataAsync()

            Dim localPlayerNames As New List(Of String)
            Dim localPlayerPhones As New Dictionary(Of String, String)
            Dim localPlayerIds As New Dictionary(Of String, Long)
            Dim localPlayerChoices As New Dictionary(Of String, List(Of Tuple(Of Long, String)))
            Dim localArchivedPlayers As New HashSet(Of String)
            Dim autoComplete As New AutoCompleteStringCollection()

            If result.Item1 IsNot Nothing Then
                For Each row As DataRow In result.Item1.Rows
                    Dim name = UtilityModule.SafeString(row(AppConstants.Col_PlayerName))
                    Dim phone = UtilityModule.SafeString(row(AppConstants.Col_Fone))
                    If Not String.IsNullOrEmpty(name) Then
                        If Not localPlayerPhones.ContainsKey(name) Then
                            localPlayerNames.Add(name)
                            localPlayerPhones(name) = phone
                            autoComplete.Add(name)
                        End If
                        ' 🌟 حفظ معرف اللاعب — أساس الربط الجديد بالـ ID بدل الاسم
                        Dim pid As Long = 0
                        If Long.TryParse(UtilityModule.SafeString(row(AppConstants.Col_PlayerId)), pid) AndAlso pid > 0 Then
                            localPlayerIds(name) = pid
                            ' 🔴 H-03: تجميع كل المعرفات لكل اسم — التكرار يُحل بحوار صريح لا اختيار صامت
                            If Not localPlayerChoices.ContainsKey(name) Then
                                localPlayerChoices(name) = New List(Of Tuple(Of Long, String))
                            End If
                            localPlayerChoices(name).Add(Tuple.Create(pid, UtilityModule.SafeString(row(AppConstants.Col_Fone))))
                        End If
                    End If
                Next
            End If

            If result.Item2 IsNot Nothing Then
                For Each row As DataRow In result.Item2.Rows
                    Dim name = UtilityModule.SafeString(row(AppConstants.Col_PlayerName))
                    If Not String.IsNullOrEmpty(name) Then
                        localArchivedPlayers.Add(name)
                        ' 🔴 H-03: المؤرشفون مرشحون أيضاً — تكرار (نشط + مؤرشف) يُحل بالحوار لا صامتاً
                        Dim aid As Long = 0
                        If Long.TryParse(UtilityModule.SafeString(row(AppConstants.Col_PlayerId)), aid) AndAlso aid > 0 Then
                            If Not localPlayerChoices.ContainsKey(name) Then
                                localPlayerChoices(name) = New List(Of Tuple(Of Long, String))
                            End If
                            Dim aphone As String = UtilityModule.SafeString(row(AppConstants.Col_Fone))
                            localPlayerChoices(name).Add(Tuple.Create(aid, If(String.IsNullOrWhiteSpace(aphone), "مؤرشف", aphone & " — مؤرشف")))
                        End If
                    End If
                Next
            End If

            localPlayerNames.Sort()

            If updateCombo Then
                If Me.InvokeRequired Then
                    Me.Invoke(Sub() UpdateComboBoxUI(localPlayerNames, localPlayerPhones, localArchivedPlayers, autoComplete, localPlayerIds, localPlayerChoices))
                Else
                    UpdateComboBoxUI(localPlayerNames, localPlayerPhones, localArchivedPlayers, autoComplete, localPlayerIds, localPlayerChoices)
                End If
            Else
                ' 🔴 H-03: تحديث الكاش فقط — دون مساس بالكمبو حتى لا يُمحى ما كتبه المستخدم
                If Me.InvokeRequired Then
                    Me.Invoke(Sub() UpdatePlayerCachesOnly(localPlayerNames, localPlayerPhones, localArchivedPlayers, localPlayerIds, localPlayerChoices))
                Else
                    UpdatePlayerCachesOnly(localPlayerNames, localPlayerPhones, localArchivedPlayers, localPlayerIds, localPlayerChoices)
                End If
            End If

        Catch ex As Exception
            DatabaseModule.LogError("LoadAllPlayersComboAsync", ex)
        End Try
    End Function
    ''' <summary>🔴 H-03: تحديث كاش اللاعبين (أسماء/هواتف/معرفات/خيارات) دون تغيير عناصر الكمبو —
    ''' يُستدعى قبل كل حفظ ليكون حل التكرار على بيانات حية لا على كاش لحظة فتح الشاشة</summary>
    Private Sub UpdatePlayerCachesOnly(names As List(Of String), phones As Dictionary(Of String, String),
                                       archived As HashSet(Of String),
                                       ids As Dictionary(Of String, Long),
                                       choices As Dictionary(Of String, List(Of Tuple(Of Long, String))))
        playerPhones.Clear()
        playerNames.Clear()
        archivedPlayers.Clear()
        playerIds.Clear()
        playerChoices.Clear()

        For Each kv In phones
            playerPhones(kv.Key) = kv.Value
        Next
        For Each kv In ids
            playerIds(kv.Key) = kv.Value
        Next
        For Each kv In choices
            playerChoices(kv.Key) = kv.Value
        Next
        playerNames.AddRange(names)
        For Each a In archived
            archivedPlayers.Add(a)
        Next
    End Sub

    Private Sub UpdateComboBoxUI(names As List(Of String), phones As Dictionary(Of String, String),
                                 archived As HashSet(Of String), autoColl As AutoCompleteStringCollection,
                                 ids As Dictionary(Of String, Long),
                                 choices As Dictionary(Of String, List(Of Tuple(Of Long, String))))
        playerPhones.Clear()
        playerNames.Clear()
        archivedPlayers.Clear()
        playerIds.Clear()
        playerChoices.Clear()

        For Each kv In phones
            playerPhones(kv.Key) = kv.Value
        Next
        For Each kv In ids
            playerIds(kv.Key) = kv.Value
        Next
        For Each kv In choices
            playerChoices(kv.Key) = kv.Value
        Next
        playerNames.AddRange(names)
        For Each a In archived
            archivedPlayers.Add(a)
        Next

        suppressSearch = True
        cmbPlayerName.BeginUpdate()
        Try
            cmbPlayerName.Items.Clear()
            cmbPlayerName.Items.Add("-- اختر اللاعب --")
            cmbPlayerName.Items.AddRange(names.ToArray())
            cmbPlayerName.SelectedIndex = 0

            cmbPlayerName.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cmbPlayerName.AutoCompleteSource = AutoCompleteSource.CustomSource
            cmbPlayerName.AutoCompleteCustomSource = autoColl
        Finally
            cmbPlayerName.EndUpdate()
            suppressSearch = False
        End Try
    End Sub
    Private Sub frmPayments_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        isClosing = True
        searchTimer?.Stop()
        searchTimer?.Dispose()
        CancelSearch()
    End Sub
#End Region

#Region "Initialization"

    Private Sub SetupSearchTimer()
        searchTimer.Interval = 800
        searchTimer.Enabled = False
    End Sub

    Private Sub SetupDataGridView()
        ' 🌟 الستايل الموحد — اختلافات هذه الشاشة: تحديد فردي + قراءة فقط + خط 9
        GridHelper.ApplyCommonGridStyle(DataGridView1, multiSelect:=False, readOnlyGrid:=True, cellFontSize:=9.0F)

        AddCustomColumns()

        AddHandler DataGridView1.RowPrePaint, AddressOf DataGridView1_RowPrePaint
        AddHandler DataGridView1.CellDoubleClick, AddressOf DataGridView1_CellDoubleClick
        AddHandler DataGridView1.DataBindingComplete, AddressOf DataGridView1_DataBindingComplete
        AddHandler DataGridView1.Enter, AddressOf DataGridView1_Enter
        AddHandler DataGridView1.Leave, AddressOf DataGridView1_Leave
        AddHandler DataGridView1.CurrentCellDirtyStateChanged, AddressOf OnCurrentCellDirtyStateChanged
        AddHandler DataGridView1.CellValueChanged, AddressOf OnCellValueChanged
        AddHandler DataGridView1.ColumnHeaderMouseClick, AddressOf DataGridView1_ColumnHeaderMouseClick
        AddHandler DataGridView1.CellClick, AddressOf DataGridView1_CellClick
    End Sub
    Private Sub AddCustomColumns()
        Try
            If Not DataGridView1.Columns.Contains(AppConstants.Grid_SelectColumn) Then
                Dim checkColumn As New DataGridViewCheckBoxColumn() With {
                    .Name = AppConstants.Grid_SelectColumn,
                    .HeaderText = "□",
                    .Width = 40,
                    .ReadOnly = False,
                    .DisplayIndex = 0
                }
                DataGridView1.Columns.Insert(0, checkColumn)
            End If

            If Not DataGridView1.Columns.Contains(AppConstants.Grid_SeqColumn) Then
                Dim seqCol As New DataGridViewTextBoxColumn() With {
                    .Name = AppConstants.Grid_SeqColumn,
                    .HeaderText = "ت",
                    .Width = 40,
                    .ReadOnly = True,
                    .DisplayIndex = 1
                }
                seqCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns.Insert(1, seqCol)
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.AddCustomColumns", ex)
        End Try
    End Sub

    Private Async Function LoadPaymentsAsync() As Task
        Try
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            Cursor = Cursors.WaitCursor
            DataGridView1.SuspendLayout()

            ' 🌟 L-09: عرض كامل = فلاتر فارغة — الصفحة الأولى فقط ثم التحميل التزايدي عند التمرير
            ClearPaymentFilters()
            Dim dt As DataTable = Await FetchPaymentsPageAsync(0, LoadMorePageSize, CancellationToken.None)

            If dt IsNot Nothing Then
                DataGridView1.DataSource = dt
                SetArabicHeaders()
                Await RefreshTotalsAsync()
                totalPaymentsCount = _totalRowsInFilter
                Me.Text = "شاشة ادخال الايرادات"
                UpdateLoadMoreState()
            End If

        Catch ex As Exception
            DatabaseModule.LogError("LoadPaymentsAsync", ex)
            MessageBox.Show("خطأ في تحميل البيانات: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            DataGridView1.ResumeLayout()
            ProgressBar1.Visible = False
            Cursor = Cursors.Default
        End Try
    End Function

    Private Sub SetDefaultComboBoxValues()
        suppressSearch = True
        Try
            cmbPaymentMethod.SelectedIndex = -1
            cmbCurrencyType.SelectedIndex = -1
            cmbPlayerName.Text = ""
            txtMobileNumber.Text = ""
            dtpDueDate.Enabled = False
        Finally
            suppressSearch = False
        End Try
    End Sub

    Private Sub SetArabicHeaders()
        For Each col As DataGridViewColumn In DataGridView1.Columns
            Select Case col.Name
                Case AppConstants.Grid_SelectColumn : col.HeaderText = "□"
                Case AppConstants.Col_Pay_PlayerName : col.HeaderText = "اسم اللاعب"
                Case AppConstants.Col_Pay_TransferorName : col.HeaderText = "اسم المحول"
                Case AppConstants.Col_Pay_VoucherNumber : col.HeaderText = "رقم السند"
                Case AppConstants.Col_Pay_Amount : col.HeaderText = "المبلغ"
                Case AppConstants.Col_Pay_CurrencyType : col.HeaderText = "العملة"
                Case AppConstants.Col_Pay_PaymentMethod : col.HeaderText = "طريقة الدفع"
                Case AppConstants.Col_Pay_PaymentDate : col.HeaderText = "تاريخ الدفع"
                Case AppConstants.Col_Pay_DueDate : col.HeaderText = "تاريخ الاستحقاق"
                Case "MobileNumber" : col.HeaderText = "رقم الجوال"
                Case AppConstants.Col_Pay_Notes : col.HeaderText = "ملاحظات"
                Case AppConstants.Grid_SeqColumn : col.HeaderText = "ت"
                Case AppConstants.Col_Pay_LastModifiedBy : col.HeaderText = "آخر تعديل بواسطة"
                Case "LastModifiedDate" : col.HeaderText = "تاريخ آخر تعديل"
            End Select
        Next
    End Sub

#End Region

#Region "DataGridView Events"

    Private Sub DataGridView1_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs)
    End Sub

    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        Try
            DataGridView1.SuspendLayout()
            AddSequenceColumn()
            ColorizeRows()
            UpdateColumnHeaderState()
        Catch ex As Exception
            DatabaseModule.LogError("DataGridView1_DataBindingComplete", ex)
        Finally
            DataGridView1.ResumeLayout()
        End Try
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If isReportMode OrElse e.RowIndex < 0 Then Return
        If e.ColumnIndex >= 0 AndAlso DataGridView1.Columns(e.ColumnIndex).Name = AppConstants.Grid_SelectColumn Then Return

        Try
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            Dim name As String = UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_PlayerName).Value)

            If archivedPlayers.Contains(name) Then
                MessageBox.Show("هذا اللاعب موجود في الأرشيف، لا يمكن تعديل دفعاته.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            suppressSearch = True
            Try
                LoadVoucherData(e.RowIndex)
            Finally
                suppressSearch = False
            End Try

        Catch ex As Exception
            DatabaseModule.LogError("DataGridView1_CellDoubleClick", ex)
            MessageBox.Show("خطأ في عرض البيانات: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_Enter(sender As Object, e As EventArgs)
        suppressSearch = True
        searchTimer.Stop()
    End Sub

    Private Sub DataGridView1_Leave(sender As Object, e As EventArgs)
        suppressSearch = False
    End Sub

    Private Sub OnCurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
    ' 🌟 النقر على خلية التحديد يقلب العلامة — الجريد readonly فرح يديرها برمجياً
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
            If DataGridView1.Columns(e.ColumnIndex).Name = AppConstants.Grid_SelectColumn Then
                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
                If Not row.IsNewRow Then
                    Dim cur As Object = row.Cells(AppConstants.Grid_SelectColumn).Value
                    Dim curVal As Boolean = (cur IsNot Nothing AndAlso Not IsDBNull(cur) AndAlso Convert.ToBoolean(cur))
                    row.Cells(AppConstants.Grid_SelectColumn).Value = Not curVal
                    UpdateColumnHeaderState()
                End If
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.DataGridView1_CellClick", ex)
        End Try
    End Sub
    Private Sub OnCellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If DataGridView1.Columns.Contains(AppConstants.Grid_SelectColumn) AndAlso
           e.ColumnIndex = DataGridView1.Columns(AppConstants.Grid_SelectColumn).Index AndAlso
           e.RowIndex >= 0 Then
            UpdateColumnHeaderState()
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs)
        Try
            If Not DataGridView1.Columns.Contains(AppConstants.Grid_SelectColumn) Then Return
            If e.ColumnIndex <> DataGridView1.Columns(AppConstants.Grid_SelectColumn).Index Then Return

            DataGridView1.SuspendLayout()
            GridHelper.ToggleAllRows(DataGridView1)
            DataGridView1.ResumeLayout()
        Catch ex As Exception
            DatabaseModule.LogError("DataGridView1_ColumnHeaderMouseClick", ex)
        End Try
    End Sub
#End Region

#Region "Grid Helpers"

    Private Sub AddSequenceColumn()
        Try
            GridHelper.FillSequenceNumbers(DataGridView1)
            ReorderColumns()
            SetColumnsAlignment()
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.AddSequenceColumn", ex)
        End Try
    End Sub
    Private Sub SetColumnsAlignment()
        Dim centerCols As String() = {
            AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn, AppConstants.Col_Pay_Amount, AppConstants.Col_Pay_CurrencyType,
            AppConstants.Col_Pay_PaymentMethod, AppConstants.Col_Pay_PaymentDate,
            AppConstants.Col_Pay_DueDate, "MobileNumber", AppConstants.Col_Pay_VoucherNumber
        }
        For Each colName In centerCols
            If DataGridView1.Columns.Contains(colName) Then
                DataGridView1.Columns(colName).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
        Next

        For Each leftCol As String In {AppConstants.Col_Pay_PlayerName, AppConstants.Col_Pay_TransferorName, AppConstants.Col_Pay_Notes}
            If DataGridView1.Columns.Contains(leftCol) Then
                DataGridView1.Columns(leftCol).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If
        Next
    End Sub

    Private Sub ReorderColumns()
        Try
            Dim existing As New HashSet(Of String)
            For Each col As DataGridViewColumn In DataGridView1.Columns
                existing.Add(col.Name)
            Next

            Dim order As String() = {
                AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn, AppConstants.Col_Pay_PlayerName, AppConstants.Col_Pay_TransferorName,
                AppConstants.Col_Pay_VoucherNumber, AppConstants.Col_Pay_Amount, AppConstants.Col_Pay_CurrencyType,
                AppConstants.Col_Pay_PaymentMethod, AppConstants.Col_Pay_PaymentDate,
                AppConstants.Col_Pay_DueDate, "MobileNumber", AppConstants.Col_Pay_Notes
            }

            Dim validOrder = order.Where(Function(n) existing.Contains(n)).ToArray()
            For i As Integer = 0 To validOrder.Length - 1
                DataGridView1.Columns(validOrder(i)).DisplayIndex = i
            Next

            Dim widths As New Dictionary(Of String, Integer) From {
                {AppConstants.Grid_SelectColumn, 40}, {AppConstants.Grid_SeqColumn, 40},
                {AppConstants.Col_Pay_PlayerName, 200}, {AppConstants.Col_Pay_TransferorName, 200}, {AppConstants.Col_Pay_VoucherNumber, 90},
                {AppConstants.Col_Pay_Amount, 90}, {AppConstants.Col_Pay_CurrencyType, 70}, {AppConstants.Col_Pay_PaymentMethod, 80},
                {AppConstants.Col_Pay_PaymentDate, 90}, {AppConstants.Col_Pay_DueDate, 100},
                {"MobileNumber", 100}, {AppConstants.Col_Pay_Notes, 200}
            }
            For Each kv In widths
                If DataGridView1.Columns.Contains(kv.Key) Then
                    DataGridView1.Columns(kv.Key).Width = kv.Value
                End If
            Next

            If DataGridView1.Columns.Contains(AppConstants.Col_Pay_Amount) Then
                DataGridView1.Columns(AppConstants.Col_Pay_Amount).DefaultCellStyle.Format = "#,##0.##"
            End If
            If DataGridView1.Columns.Contains(AppConstants.Col_Pay_Notes) Then
                DataGridView1.Columns(AppConstants.Col_Pay_Notes).DefaultCellStyle.WrapMode = DataGridViewTriState.True
                DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            End If


            ' 🌟 إخفاء أعمدة المعرفات التقنية — موجودة بالبيانات لكن غير مرئية
            If DataGridView1.Columns.Contains(AppConstants.Col_Pay_ID) Then
                DataGridView1.Columns(AppConstants.Col_Pay_ID).Visible = False
            End If
            If DataGridView1.Columns.Contains(AppConstants.Col_Pay_PlayerId) Then
                DataGridView1.Columns(AppConstants.Col_Pay_PlayerId).Visible = False
            End If
        Catch ex As Exception
            DatabaseModule.LogError("ReorderColumns", ex)
        End Try
    End Sub

    Private Sub ColorizeRows()
    End Sub

    Private Function GetSelectedRowsFromCheckBox() As List(Of DataGridViewRow)
        Return GridHelper.GetCheckedRows(DataGridView1)
    End Function
    Private Sub UpdateColumnHeaderState()
        GridHelper.UpdateSelectHeaderState(DataGridView1)
    End Sub
#End Region

#Region "Control Enable/Disable"

    Private Sub DisablePaymentControls()
        cmbPlayerName.Enabled = False
        txtMobileNumber.Enabled = False
        dtpPaymentDate.Enabled = False
        dtpDueDate.Enabled = False
        txtAmount.Enabled = False
        cmbPaymentMethod.Enabled = False
        txtTransferorName.Enabled = False
        txtNotes.Enabled = False
        cmbCurrencyType.Enabled = False
        txtVoucherNumber.Enabled = False

        AddimageToolStripMenuItem.Enabled = False
        DeleteimageToolStripMenuItem.Enabled = False
        PrintimageToolStripMenuItem.Enabled = False

        SaveToolStripMenuItem.Enabled = False
        DeleteToolStripMenuItem.Enabled = False
        GeneralReportToolStripMenuItem.Enabled = True
        PaymentReportToolStripMenuItem.Enabled = True
        PrintReportToolStripMenuItem.Enabled = True
        ExporteToExcelToolStripMenuItem.Enabled = True
    End Sub

    Private Sub EnablePaymentControls()
        cmbPlayerName.Enabled = True
        txtMobileNumber.Enabled = False
        dtpPaymentDate.Enabled = True
        dtpDueDate.Enabled = False
        txtAmount.Enabled = True
        cmbPaymentMethod.Enabled = True
        txtTransferorName.Enabled = True
        txtNotes.Enabled = True
        cmbCurrencyType.Enabled = True
        txtVoucherNumber.Enabled = True

        AddimageToolStripMenuItem.Enabled = True
        PrintimageToolStripMenuItem.Enabled = True

        SaveToolStripMenuItem.Enabled = UserSession.CanEdit
        DeleteToolStripMenuItem.Enabled = UserSession.CanDelete
        NewToolStripMenuItem.Enabled = True

        GeneralReportToolStripMenuItem.Enabled = True
        PaymentReportToolStripMenuItem.Enabled = True
        PrintReportToolStripMenuItem.Enabled = True
        ExporteToExcelToolStripMenuItem.Enabled = True

        If isEditMode Then
            txtVoucherNumber.Enabled = False
            DeleteimageToolStripMenuItem.Enabled = UserSession.CanDelete
        End If
    End Sub

    Private Sub SetSavingState(saving As Boolean)
        ' 🌟 إصلاح ثغرة صلاحيات: عند انتهاء الحفظ تُستعاد الأزرار حسب صلاحيات المستخدم الفعلية
        SaveToolStripMenuItem.Enabled = Not saving AndAlso UserSession.CanEdit
        DeleteToolStripMenuItem.Enabled = Not saving AndAlso UserSession.CanDelete
        NewToolStripMenuItem.Enabled = Not saving
    End Sub

#End Region

#Region "Menu Click Handlers"

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click

    End Sub

    Private Async Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Try
            ' 🌟 L-02: الحالة بعلم منطقي صريح بدل نص زر الحفظ — التغييرات الشكلية لا تكسر المنطق
            If isReportMode Then
                Await ReturnFromReportModeAsync()
            ElseIf isEditMode Then
                Await UpdatePaymentAsync()
            Else
                Await InsertPaymentAsync()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("SaveToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ أثناء الحفظ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Try
            Await DeletePaymentAsync()
        Catch ex As Exception
            DatabaseModule.LogError("DeleteToolStripMenuItem_Click", ex)
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Async Function ReturnFromReportModeAsync() As Task
        If TempData IsNot Nothing Then
            DataGridView1.SuspendLayout()
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = TempData
            SetArabicHeaders()
            UpdateTotalAmount()
            DataGridView1.ResumeLayout()
            TempData = Nothing
        Else
            Await LoadPaymentsAsync()
        End If

        If lblSummary IsNot Nothing Then lblSummary.Text = ""
        SaveToolStripMenuItem.Text = "حفظ"
        isReportMode = False
        UpdateLoadMoreState()   ' 🌟 L-09: إعادة إظهار "تحميل المزيد" إن بقيت سطور غير محمّلة
        DisablePaymentControls()
        NewToolStripMenuItem.Enabled = True
        Me.Text = "شاشة ادخال الايرادات"
    End Function
#End Region

#Region "Data Operations"

    ''' <summary>
    ''' 🔴 H-03 — تحديد معرف اللاعب عند الحفظ:
    '''   * الاسم من القائمة بمعرّف واحد → يُستخدم مباشرة من الذاكرة (لا بحث صامت بالاسم)
    '''   * الاسم مكرر → حوار اختيار صريح بين أصحاب الهويات — والإلغاء يلغي الحفظ كلياً
    '''   * اسم مكتوب يدوياً خارج القائمة → الملاذ الأخير: البحث بالاسم (يشمل الأرشيف)
    ''' </summary>
    Private Async Function ResolvePlayerIdForSaveAsync(playerName As String) As Task(Of Tuple(Of Object, Boolean))
        ' النتيجة: Item1 = معرف اللاعب (أو Nothing) — Item2 = True إذا ألغى المستخدم الاختيار
        If String.IsNullOrEmpty(playerName) Then Return Tuple.Create(Of Object, Boolean)(Nothing, False)

        If playerChoices.ContainsKey(playerName) Then
            Dim options As List(Of Tuple(Of Long, String)) = playerChoices(playerName)
            If options.Count = 1 Then
                Return Tuple.Create(Of Object, Boolean)(options(0).Item1, False)
            ElseIf options.Count > 1 Then
                Using picker As New frmPickPlayer(playerName, options)
                    If picker.ShowDialog(Me) = DialogResult.OK AndAlso picker.SelectedId > 0 Then
                        Return Tuple.Create(Of Object, Boolean)(picker.SelectedId, False)
                    End If
                End Using
                ' ألغى المستخدم حوار اختيار اللاعب — إشارة إلغاء صريحة، لا حفظ صامت
                Return Tuple.Create(Of Object, Boolean)(Nothing, True)
            End If
        End If

        ' الاسم خارج القائمة (مكتوب يدوياً) — الملاذ الأخير فقط، ويشمل الأرشيف
        ' 🌟 إصلاح v2: الاسم خارج القائمة — نحاول أولاً جلب كل المطابقين لعرض حوار اختيار صريح
        Dim fallbackService As New PaymentService()
        Dim allMatches As List(Of Tuple(Of Long, String)) = Await fallbackService.GetAllPlayerMatchesByNameAsync(playerName)

        If allMatches IsNot Nothing AndAlso allMatches.Count > 1 Then
            ' أسماء مكررة → حوار اختيار صريح (نفس نمط الأسماء داخل القائمة)
            Using picker As New frmPickPlayer(playerName, allMatches)
                If picker.ShowDialog(Me) = DialogResult.OK AndAlso picker.SelectedId > 0 Then
                    Return Tuple.Create(Of Object, Boolean)(picker.SelectedId, False)
                End If
            End Using
            Return Tuple.Create(Of Object, Boolean)(Nothing, True)   ' ألغى المستخدم
        ElseIf allMatches IsNot Nothing AndAlso allMatches.Count = 1 Then
            Return Tuple.Create(Of Object, Boolean)(allMatches(0).Item1, False)
        End If

        Return Tuple.Create(Of Object, Boolean)(Nothing, False)
    End Function

    Private Async Function InsertPaymentAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True
            SetSavingState(True)
            ' 🔴 H-03: تحديث كاش اللاعبين قبل الحفظ — اللاعب المضاف حديثاً يدخل حل التكرار فوراً
            Await LoadAllPlayersComboAsync(updateCombo:=False)
            If Not ValidateDataBasic() Then Return

            Dim cleanVoucher As String = txtVoucherNumber.Text.Trim()
            Dim service As New PaymentService()

            Dim playerName As String = ""
            If Not TryNormalizePlayerName(playerName) Then Return

            ' 🔴 H-03: الربط بالمعرف من القائمة المحمّلة — الأسماء المكررة تُحل بحوار صريح لا اختيار صامت
            Dim playerIdObj As Object = Nothing
            If Not String.IsNullOrEmpty(playerName) Then
                Dim resolveResult As Tuple(Of Object, Boolean) = Await ResolvePlayerIdForSaveAsync(playerName)
                If resolveResult.Item2 Then Return   ' ألغى المستخدم اختيار اللاعب — لا حفظ صامت
                playerIdObj = resolveResult.Item1
            End If
            If Not String.IsNullOrWhiteSpace(cleanVoucher) Then
                ' 🌟 إصلاح: الفحص كان بلا عملة/لاعب فيصبح حارساً ميتاً (شرط العملة '' لا يطابق أي صف طبيعي)
                Dim duplicateFound As Boolean
                If playerIdObj IsNot Nothing Then
                    duplicateFound = Await service.IsVoucherNumberExistsAsync(cleanVoucher,
                        playerId:=playerIdObj, currency:=cmbCurrencyType.Text)
                Else
                    ' سند "محول فقط" بلا لاعب — مسار فحص مخصص له
                    duplicateFound = Await service.IsVoucherNumberExistsForTransferorAsync(cleanVoucher)
                End If
                If duplicateFound Then
                    MessageBox.Show("رقم السند موجود مسبقاً! الرجاء استخدام رقم سند مختلف.", "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtVoucherNumber.Focus()
                    Return
                End If
            End If

            Dim amountValue As Decimal
            ' 🌟 تحليل المبلغ بثقافة ثابتة (InvariantCulture) حتى لا تُفسد المبالغ على أجهزة تستخدم الفاصلة بدل النقطة
            If Not Decimal.TryParse(txtAmount.Text, NumberStyles.Number, CultureInfo.InvariantCulture, amountValue) OrElse amountValue <= 0 Then
                MessageBox.Show("الرجاء إدخال مبلغ صحيح أكبر من صفر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Return
            End If
            Dim modifiedDateStr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            Dim photoParam As Object = UtilityModule.ImageToDBValue(PictureBox1.Image)

            ' 🌟             ' 🌟 حفظ التواريخ بصيغة ISO (yyyy-MM-dd) — العرض dd/MM/yyyy في الواجهة فقط
            Dim affected As Integer = Await service.InsertPaymentAsync(
                playerName, playerIdObj, dtpPaymentDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), txtVoucherNumber.Text,
                amountValue, cmbPaymentMethod.Text, txtTransferorName.Text,
                txtNotes.Text, dtpDueDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), cmbCurrencyType.Text,
                UserSession.CurrentUsername, modifiedDateStr, photoParam)
            If affected > 0 Then
                MessageBox.Show("تم إضافة الدفعة بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Await LoadPaymentsAsync()
                SetAddMode()
                DisablePaymentControls()
            Else
                MessageBox.Show("لم يتم إدراج البيانات — تأكد من صحة المعلومات.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("InsertPaymentAsync", ex)
            MessageBox.Show("خطأ في الإضافة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function
    Private Async Function UpdatePaymentAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True
            SetSavingState(True)
            ' 🔴 H-03: تحديث كاش اللاعبين قبل الحفظ — اللاعب المضاف حديثاً يدخل حل التكرار فوراً
            Await LoadAllPlayersComboAsync(updateCombo:=False)
            If currentPaymentId <= 0 Then
                MessageBox.Show("الرجاء اختيار دفعة للتحديث", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If Not ValidateDataBasic() Then Return

            Dim cleanVoucher As String = txtVoucherNumber.Text.Trim()
            Dim service As New PaymentService()

            Dim playerName As String = ""
            If Not TryNormalizePlayerName(playerName) Then Return

            ' 🔴 H-03: تحديث الربط بالمعرف من القائمة المحمّلة — التكرار بحوار صريح لا اختيار صامت
            Dim playerIdObj As Object = Nothing
            If Not String.IsNullOrEmpty(playerName) Then
                Dim resolveResult As Tuple(Of Object, Boolean) = Await ResolvePlayerIdForSaveAsync(playerName)
                If resolveResult.Item2 Then Return   ' ألغى المستخدم اختيار اللاعب — لا حفظ صامت
                playerIdObj = resolveResult.Item1
            End If
            If Not String.IsNullOrWhiteSpace(cleanVoucher) AndAlso cleanVoucher <> currentVoucherNumber Then
                ' 🌟 إصلاح: تمرير اللاعب والعملة حتى يعمل الفحص فعلياً (مع استثناء الدفعة الحالية)
                Dim duplicateFound As Boolean
                If playerIdObj IsNot Nothing Then
                    duplicateFound = Await service.IsVoucherNumberExistsAsync(cleanVoucher, currentPaymentId,
                        playerId:=playerIdObj, currency:=cmbCurrencyType.Text)
                Else
                    duplicateFound = Await service.IsVoucherNumberExistsForTransferorAsync(cleanVoucher)
                End If
                If duplicateFound Then
                    MessageBox.Show("رقم السند موجود مسبقاً! الرجاء استخدام رقم سند مختلف.", "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtVoucherNumber.Focus()
                    Return
                End If
            End If

            Dim amountValue As Decimal
            ' 🌟 تحليل المبلغ بثقافة ثابتة (InvariantCulture) حتى لا تُفسد المبالغ على أجهزة تستخدم الفاصلة بدل النقطة
            If Not Decimal.TryParse(txtAmount.Text, NumberStyles.Number, CultureInfo.InvariantCulture, amountValue) OrElse amountValue <= 0 Then
                MessageBox.Show("الرجاء إدخال مبلغ صحيح أكبر من صفر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Return
            End If
            Dim modifiedDateStr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            Dim photoParam As Object = Nothing
            If _imageChanged Then
                photoParam = UtilityModule.ImageToDBValue(PictureBox1.Image)
            End If

            ' 🌟             ' 🌟 حفظ التواريخ بصيغة ISO (yyyy-MM-dd) — العرض dd/MM/yyyy في الواجهة فقط
            Dim affected As Integer = Await service.UpdatePaymentAsync(
    currentPaymentId,
    playerName, playerIdObj, dtpPaymentDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), txtVoucherNumber.Text,
    amountValue, cmbPaymentMethod.Text, txtTransferorName.Text,
        txtNotes.Text, dtpDueDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), cmbCurrencyType.Text,
    UserSession.CurrentUsername, modifiedDateStr, photoParam, _imageChanged)
            If affected > 0 Then
                MessageBox.Show("تم تحديث الدفعة بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                _imageChanged = False
                Await LoadPaymentsAsync()
                SetAddMode()
                DisablePaymentControls()
            Else
                MessageBox.Show("لم يتم تحديث أي بيانات — تأكد من رقم السند.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("UpdatePaymentAsync", ex)
            MessageBox.Show("خطأ في التحديث: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function
    Private Async Function DeletePaymentAsync() As Task
        If _isSaving Then Return
        Try
            Dim service As New PaymentService()
            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            Dim idsToDelete As New List(Of Long)

            ' 🌟 1. إذا حدد المستخدم سجلات من الجدول، نحذفها بالمعرف الفريد
            If selectedRows.Count > 0 Then
                For Each row In selectedRows
                    If DataGridView1.Columns.Contains(AppConstants.Col_Pay_ID) AndAlso
                       row.Cells(AppConstants.Col_Pay_ID).Value IsNot Nothing AndAlso
                       Not IsDBNull(row.Cells(AppConstants.Col_Pay_ID).Value) Then
                        idsToDelete.Add(Convert.ToInt64(row.Cells(AppConstants.Col_Pay_ID).Value))
                    End If
                Next

                If idsToDelete.Count = 0 Then
                    MessageBox.Show("السجلات المحددة لا تحتوي على معرف صالح.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If MessageBox.Show($"هل أنت متأكد من حذف {idsToDelete.Count} دفعة نهائياً؟", "تأكيد الحذف",
                                 MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return

                _isSaving = True
                SetSavingState(True)

                ' 🌟 إصلاح (M-08): الحذف الجماعي الآن بمعاملة واحدة ذرية — كان سجلاً بسجل بلا معاملة،
                ' فأي انقطاع بالمنتصف يترك نصف الحذف (نفس إصلاح frmExpenses سابقاً)
                Dim deleteQueries As New List(Of Tuple(Of String, SQLiteParameter()))
                For Each pid As Long In idsToDelete
                    deleteQueries.Add(Tuple.Create(
                        $"DELETE FROM {AppConstants.Table_Payments} WHERE ID = @id",
                        New SQLiteParameter() {New SQLiteParameter("@id", pid)}))
                Next

                ' 🌟 بوابات الصلاحية: الحلقة القديمة كانت تمر عبر PaymentService.DeletePaymentAsync
                ' التي تفرض CanDelete — المعاملة الجماعية تفحصها صراحة
                If Not UserSession.CanDelete Then
                    MessageBox.Show("ليست لديك صلاحية الحذف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim txOk As Boolean = Await Task.Run(Function() DatabaseModule.ExecuteTransaction(deleteQueries))

                ' 🌟 فشل المعاملة = لم يُحذف أي سجل (كل شيء تُراجع)
                If Not txOk Then
                    MessageBox.Show("فشل حذف السجلات المحددة — لم يُحذف أي سجل.", "خطأ",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' 🌟 توثيق جماعي في سجل التدقيق (كان الحذف الجماعي بلا أثر فيه)
                Try
                    Dim audit As New AuditService()
                    audit.Log(AuditService.Act_PaymentDelete, "دفعة", String.Join(",", idsToDelete),
                              $"حذف جماعي: {idsToDelete.Count} سجل")
                Catch
                End Try

                MessageBox.Show($"تم حذف {idsToDelete.Count} دفعة بنجاح.", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Await LoadPaymentsAsync()
                SetAddMode()
                DisablePaymentControls()
                Return
            End If

            ' 🌟 2. حذف الدفعة المفتوحة حالياً بالمعرف الفريد
            If currentPaymentId <= 0 Then
                MessageBox.Show("الرجاء اختيار دفعة من الجدول أو تحديد سجلات بواسطة ✔ للحذف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If MessageBox.Show("هل أنت متأكد من حذف هذه الدفعة نهائياً؟", "تأكيد الحذف",
                              MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                              MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return

            _isSaving = True
            SetSavingState(True)

            Dim affected As Integer = Await service.DeletePaymentAsync(currentPaymentId)

            If affected > 0 Then
                MessageBox.Show("تم حذف الدفعة بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Await LoadPaymentsAsync()
                SetAddMode()
                DisablePaymentControls()
            Else
                MessageBox.Show("فشل حذف الدفعة — ربما حُذفت مسبقاً.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("DeletePaymentAsync", ex)
            MessageBox.Show("خطأ في الحذف: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function
    Private Function TryNormalizePlayerName(ByRef playerName As String) As Boolean
        Dim raw As String = cmbPlayerName.Text.Trim()

        If String.IsNullOrWhiteSpace(raw) OrElse raw = "-- اختر اللاعب --" Then
            playerName = ""
            Return True
        End If

        If Not playerNames.Contains(raw) Then
            ' 🌟 إذا لا يوجد اسم محول أيضاً، لا معنى للسؤال — سند بلا لاعب وبلا محول = سند يتيم
            If String.IsNullOrWhiteSpace(txtTransferorName.Text) Then
                MessageBox.Show(
                    "الاسم غير موجود في قائمة اللاعبين، وحقل المحول فارغ." & vbCrLf &
                    "الرجاء إدخال اسم لاعب صحيح أو اسم محول (واحد على الأقل).",
                    "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbPlayerName.Focus()
                playerName = ""
                Return False
            End If

            Dim ask As DialogResult = MessageBox.Show(
                "الاسم غير موجود في قائمة اللاعبين." & vbCrLf &
                "هل تريد المتابعة باستخدام اسم المحول فقط؟" & vbCrLf &
                "ملاحظة: لن يتم ربط الدفعة بأي لاعب موجود.",
                "تأكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)

            If ask = DialogResult.No Then
                cmbPlayerName.Focus()
                playerName = ""
                Return False
            End If
            playerName = ""
            Return True
        End If

        playerName = raw
        Return True
    End Function
#End Region

#Region "Voucher Image Operations"

    Private Sub AddimageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddimageToolStripMenuItem.Click
        Try
            Using ofd As New OpenFileDialog()
                ofd.Title = "اختر صورة السند"
                ofd.Filter = "ملفات الصور|*.jpg;*.jpeg;*.png;*.bmp|كل الملفات|*.*"
                If ofd.ShowDialog() = DialogResult.OK Then
                    Using fs As New FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using img As Image = Image.FromStream(fs)
                            If PictureBox1.Image IsNot Nothing Then
                                PictureBox1.Image.Dispose()
                            End If
                            PictureBox1.Image = New Bitmap(img)
                            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                        End Using
                    End Using
                    tempImagePath = ofd.FileName
                    _imageChanged = True
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("AddimageToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ في تحميل الصورة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DeleteimageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteimageToolStripMenuItem.Click
        Try
            If PictureBox1.Image Is Nothing Then
                MessageBox.Show("لا توجد صورة للحذف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If MessageBox.Show("هل أنت متأكد من حذف صورة السند؟", "تأكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                If PictureBox1.Image IsNot Nothing Then
                    PictureBox1.Image.Dispose()
                End If
                PictureBox1.Image = Nothing
                tempImagePath = ""
                _imageChanged = True
                If isEditMode Then
                    MessageBox.Show("تم حذف الصورة، الرجاء الضغط على تحديث لحفظ التغييرات", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            DatabaseModule.LogError("DeleteimageToolStripMenuItem_Click", ex)
        End Try
    End Sub

    Private Sub PrintimageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintimageToolStripMenuItem.Click
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("لا توجد صورة لطباعتها", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim prnPreview As New PrintPreviewDialog() With {
                .Document = imgPrintDoc,
                .WindowState = FormWindowState.Maximized,
                .Text = "معاينة صورة السند"
            }
            prnPreview.ShowDialog()
        Catch ex As Exception
            DatabaseModule.LogError("PrintimageToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ في طباعة الصورة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub imgPrintDoc_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles imgPrintDoc.PrintPage
        If PictureBox1.Image Is Nothing Then
            e.HasMorePages = False
            Return
        End If

        Dim img As Image = PictureBox1.Image
        Dim ratio As Double = Math.Min(CDbl(e.MarginBounds.Width) / img.Width, CDbl(e.MarginBounds.Height) / img.Height)
        Dim newW As Integer = CInt(img.Width * ratio)
        Dim newH As Integer = CInt(img.Height * ratio)
        Dim x As Integer = e.MarginBounds.Left + (e.MarginBounds.Width - newW) \ 2
        Dim y As Integer = e.MarginBounds.Top + (e.MarginBounds.Height - newH) \ 2

        e.Graphics.DrawImage(img, x, y, newW, newH)
        e.HasMorePages = False
    End Sub

    Private Async Function LoadVoucherImageAsync(paymentId As Long) As Task
        _voucherImageLoadId += 1
        Dim thisLoadId As Integer = _voucherImageLoadId
        Try
            If PictureBox1.Image IsNot Nothing Then
                PictureBox1.Image.Dispose()
                PictureBox1.Image = Nothing
            End If

            If paymentId <= 0 Then Return

            Dim service As New PaymentService()
            Dim photoBytes As Byte() = Await service.GetVoucherImageAsync(paymentId)

            ' 🌟 تجاهل النتيجة إذا بدأ تحميل صورة أخرى أثناء الانتظار (حماية من سباق التحميل)
            If thisLoadId <> _voucherImageLoadId Then Return

            If photoBytes IsNot Nothing AndAlso photoBytes.Length > 0 Then
                Dim bmp As Image = UtilityModule.ByteArrayToImage(photoBytes)
                If bmp IsNot Nothing Then
                    PictureBox1.Image = bmp
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                End If
            End If
        Catch ex As Exception
            DatabaseModule.LogError("LoadVoucherImageAsync", ex)
        End Try
    End Function
#End Region

#Region "Voucher Validation - UI"

    Private Async Sub txtVoucherNumber_Leave(sender As Object, e As EventArgs) Handles txtVoucherNumber.Leave
        If isClosing OrElse isEditMode Then Return
        Try
            If String.IsNullOrWhiteSpace(txtVoucherNumber.Text) Then Return
            Dim cleanVoucher As String = txtVoucherNumber.Text.Trim()

            Dim service As New PaymentService()
            Dim owners As List(Of String) = Await service.GetVoucherOwnersAsync(cleanVoucher)

            If owners.Count > 0 Then
                Dim namesList As String = String.Join("، ", owners.Take(5))
                If owners.Count > 5 Then namesList &= " وغيرهم..."

                MessageBox.Show(
                    "رقم السند " & cleanVoucher & " مستخدم مسبقاً من: " & namesList & vbCrLf & vbCrLf &
                    "✔ إذا هذا سند صرف شامل لنفس المجموعة → تابع الإدخال عادي (تغيير اسم اللاعب فقط)" & vbCrLf &
                    "✖ إذا هذا تكرار بالغلط → غيّر رقم السند",
                    "رقم سند مستخدم", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            DatabaseModule.LogError("txtVoucherNumber_Leave", ex)
        End Try
    End Sub
    Private Sub HighlightVoucherInGrid(voucherNumber As String)
        Try
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow AndAlso DataGridView1.Columns.Contains(AppConstants.Col_Pay_VoucherNumber) Then
                    If UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_VoucherNumber).Value) = voucherNumber Then
                        DataGridView1.ClearSelection()
                        row.Selected = True
                        DataGridView1.CurrentCell = row.Cells(1)
                        LoadVoucherData(row.Index)
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            DatabaseModule.LogError("HighlightVoucherInGrid", ex)
        End Try
    End Sub

    Private Async Sub LoadVoucherData(rowIndex As Integer)
        Try
            Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)
            ' 🌟 حفظ المعرف الفريد للدفعة (يُستخدم للتحديث/الحذف/الصورة)
            currentPaymentId = -1
            If DataGridView1.Columns.Contains(AppConstants.Col_Pay_ID) AndAlso
               row.Cells(AppConstants.Col_Pay_ID).Value IsNot Nothing AndAlso
               Not IsDBNull(row.Cells(AppConstants.Col_Pay_ID).Value) Then
                currentPaymentId = Convert.ToInt64(row.Cells(AppConstants.Col_Pay_ID).Value)
            End If
            suppressSearch = True
            Try
                cmbPlayerName.Text = UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_PlayerName).Value)
                txtMobileNumber.Text = UtilityModule.SafeString(row.Cells("MobileNumber").Value)
                Dim payDate = UtilityModule.SafeDate(row.Cells(AppConstants.Col_Pay_PaymentDate).Value)
                dtpPaymentDate.Value = If(payDate = Date.MinValue, Date.Today, payDate)
                txtVoucherNumber.Text = UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_VoucherNumber).Value)
                currentVoucherNumber = txtVoucherNumber.Text
                ' 🌟 إصلاح: عرض المبلغ بثقافة ثابتة — كان بثقافة النظام فيفشل التحليل عند الحفظ على أجهزة الفاصلة ","
                txtAmount.Text = UtilityModule.ToDecimalSafe(row.Cells(AppConstants.Col_Pay_Amount).Value).ToString("0.##", CultureInfo.InvariantCulture)
                cmbPaymentMethod.Text = UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_PaymentMethod).Value)
                txtTransferorName.Text = UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_TransferorName).Value)
                txtNotes.Text = UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_Notes).Value)
                cmbCurrencyType.Text = UtilityModule.SafeString(row.Cells(AppConstants.Col_Pay_CurrencyType).Value)


                ' 🌟 تاريخ الاستحقاق: نعرض قيمة السجل إن وجدت، وإلا يبقى المحسوب تلقائياً (دفع + شهر)
                Dim dueDateVal = row.Cells(AppConstants.Col_Pay_DueDate).Value
                If dueDateVal IsNot Nothing AndAlso Not IsDBNull(dueDateVal) AndAlso UtilityModule.SafeDate(dueDateVal) <> Date.MinValue Then
                    dtpDueDate.Value = UtilityModule.SafeDate(dueDateVal)
                End If
                If String.IsNullOrWhiteSpace(txtMobileNumber.Text) AndAlso
                   playerPhones.ContainsKey(cmbPlayerName.Text) Then
                    txtMobileNumber.Text = playerPhones(cmbPlayerName.Text)
                End If
            Finally
                suppressSearch = False
            End Try

            _imageChanged = False
            Await LoadVoucherImageAsync(currentPaymentId)

            isEditMode = True
            EnablePaymentControls()
            SaveToolStripMenuItem.Text = "تحديث"

        Catch ex As Exception
            DatabaseModule.LogError("LoadVoucherData", ex)
            MessageBox.Show("خطأ في تحميل بيانات السند: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Smart Search"

    Private Async Sub OnSearchTimerTick(sender As Object, e As EventArgs) Handles searchTimer.Tick
        searchTimer.Stop()
        If suppressSearch Then Return

        CancelSearch()
        _cancellationTokenSource = New CancellationTokenSource()
        Try
            Await PerformSmartSearchAsync(_cancellationTokenSource.Token)
        Catch ex As OperationCanceledException
        Catch ex As Exception
            DatabaseModule.LogError("OnSearchTimerTick", ex)
        End Try
    End Sub

    Private Async Function PerformSmartSearchAsync(ct As System.Threading.CancellationToken) As Task
        Try
            If Not HasSearchCriteria() Then
                Await LoadPaymentsAsync()
                Return
            End If

            ' 🌟 L-09: حفظ الفلاتر الحالية — التحميل التزايدي لاحقاً يستعلم بنفس الشروط
            SetPaymentFiltersFromControls()
            ' 🌟 تم إزالة تمرير نص حالة الدفع للبحث
            Dim dt As DataTable = Await FetchPaymentsPageAsync(0, LoadMorePageSize, ct)

            ' 🌟 إصلاح: نتيجة بحث أقدم انطلقت أولاً كانت قد تعود آخراً وتطمس الأحدث —
            ' فحص الإلغاء بعد Await يضمن عدم تطبيق نتيجة بحث سبق إلغاؤه ببحث أحدث
            If dt IsNot Nothing AndAlso Not ct.IsCancellationRequested Then
                DataGridView1.SuspendLayout()
                DataGridView1.DataSource = dt
                SetArabicHeaders()
                Await RefreshTotalsAsync(ct)
                DataGridView1.ResumeLayout()
                Me.Text = "شاشة ادخال الايرادات"
                UpdateLoadMoreState()
            End If

        Catch ex As Exception
            If Not ct.IsCancellationRequested Then
                DatabaseModule.LogError("PerformSmartSearchAsync", ex)
                MessageBox.Show("خطأ في البحث: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Function

    Private Function HasSearchCriteria() As Boolean
        Return Not ((String.IsNullOrWhiteSpace(cmbPlayerName.Text) OrElse
                     cmbPlayerName.Text = "-- اختر اللاعب --") AndAlso
                    String.IsNullOrWhiteSpace(txtVoucherNumber.Text) AndAlso
                    String.IsNullOrWhiteSpace(txtTransferorName.Text) AndAlso
                    String.IsNullOrWhiteSpace(txtNotes.Text) AndAlso
                    (cmbPaymentMethod.SelectedIndex < 0 OrElse String.IsNullOrWhiteSpace(cmbPaymentMethod.Text)) AndAlso
                    (cmbCurrencyType.SelectedIndex < 0 OrElse String.IsNullOrWhiteSpace(cmbCurrencyType.Text)))
    End Function

    Private Sub CancelSearch()
        If _cancellationTokenSource IsNot Nothing Then
            Try
                _cancellationTokenSource.Cancel()
                _cancellationTokenSource.Dispose()
            Catch
            End Try
            _cancellationTokenSource = Nothing
        End If
    End Sub

#End Region

#Region "L-09 — الترقيم التزايدي لقائمة الدفعات"

    ' 🌟 L-09: لا تُحمَّل كل الدفعات دفعة واحدة — الصفحة الأولى (LoadMorePageSize) ثم دفعات
    ' لاحقة عند التمرير لآخر القائمة أو من بند "تحميل المزيد" بشريط القوائم.
    ' المجاميع والعدد الكلي يُجلبان دائماً من SQL فتبقى أرقام الشريط السفلي صحيحة كاملة.

    Private Sub ClearPaymentFilters()
        _filterPlayerName = ""
        _filterVoucherNumber = ""
        _filterTransferorName = ""
        _filterNotes = ""
        _filterPaymentMethod = ""
        _filterCurrencyType = ""
    End Sub

    Private Sub SetPaymentFiltersFromControls()
        _filterPlayerName = cmbPlayerName.Text
        _filterVoucherNumber = txtVoucherNumber.Text
        _filterTransferorName = txtTransferorName.Text
        _filterNotes = txtNotes.Text
        _filterPaymentMethod = cmbPaymentMethod.Text
        _filterCurrencyType = cmbCurrencyType.Text
    End Sub

    ''' <summary>جلب صفحة سطور وفق الفلاتر الحالية (فلاتر فارغة = كل الدفعات)</summary>
    Private Function FetchPaymentsPageAsync(offset As Integer, limit As Integer, ct As CancellationToken) As Task(Of DataTable)
        Dim service As New PaymentService()
        Return service.SearchPaymentsPagedAsync(
            _filterPlayerName, _filterVoucherNumber, _filterTransferorName,
            _filterNotes, _filterPaymentMethod, _filterCurrencyType, ct, limit, offset)
    End Function

    ''' <summary>جلب المجاميع الحقيقية والعدد الكلي لكل النتائج المطابقة (لا المعروض منها فقط).
    ''' إلغاء بحث أحدث (ct) يُبتلع بصمت — نفس دلالات مسار البحث</summary>
    Private Async Function RefreshTotalsAsync(Optional ct As CancellationToken = Nothing) As Task
        Try
            Dim service As New PaymentService()
            Dim result As Tuple(Of CurrencyTotals, Integer) = Await service.GetPaymentsTotalsAsync(
                _filterPlayerName, _filterVoucherNumber, _filterTransferorName,
                _filterNotes, _filterPaymentMethod, _filterCurrencyType, ct)

            If result IsNot Nothing AndAlso Not ct.IsCancellationRequested Then
                _currentTotals = result.Item1
                _totalRowsInFilter = result.Item2
                _hasMoreRows = CurrentLoadedRowsCount() < _totalRowsInFilter
                UpdateTotalAmount()
            End If
        Catch ex As OperationCanceledException
            ' إلغاء مقصود — لا رسالة ولا لوج
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.RefreshTotalsAsync", ex)
        End Try
    End Function

    ''' <summary>عدد السطور المحمّلة فعلياً بالعرض الحالي</summary>
    Private Function CurrentLoadedRowsCount() As Integer
        If TypeOf DataGridView1.DataSource Is DataTable Then
            Return CType(DataGridView1.DataSource, DataTable).Rows.Count
        End If
        Return 0
    End Function

    ''' <summary>جلب الدفعة التالية من السطور ودمجها بالعرض الحالي.
    ''' علم _isLoadingMore يمنع التداخل، والدمج يحدّث الشبكة حياً (ListChanged)
    ''' مع إعادة ترقيم التسلسل 1..N لكل المعروض</summary>
    Private Async Function LoadMoreRowsAsync() As Task
        If _isLoadingMore OrElse Not _hasMoreRows OrElse isReportMode OrElse isClosing Then Return
        If Not (TypeOf DataGridView1.DataSource Is DataTable) Then Return

        _isLoadingMore = True
        Try
            Dim loaded As Integer = CurrentLoadedRowsCount()
            Dim dt As DataTable = Await FetchPaymentsPageAsync(loaded, LoadMorePageSize, CancellationToken.None)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                DataGridView1.SuspendLayout()
                CType(DataGridView1.DataSource, DataTable).Merge(dt)
                SetArabicHeaders()
                AddSequenceColumn()
                DataGridView1.ResumeLayout()
            End If

            Await RefreshTotalsAsync()
            UpdateLoadMoreState()
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.LoadMoreRowsAsync", ex)
        Finally
            _isLoadingMore = False
        End Try
    End Function

    ''' <summary>تحديث بند "تحميل المزيد" (نصه وظهوره) وفق حالة التحميل</summary>
    Private Sub UpdateLoadMoreState()
        Try
            _hasMoreRows = CurrentLoadedRowsCount() < _totalRowsInFilter
            If miLoadMore IsNot Nothing Then
                miLoadMore.Visible = _hasMoreRows AndAlso Not isReportMode
                miLoadMore.Text = $"تحميل المزيد — معروض {CurrentLoadedRowsCount():N0} من {_totalRowsInFilter:N0}"
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.UpdateLoadMoreState", ex)
        End Try
    End Sub

    ''' <summary>إضافة بند "تحميل المزيد" لشريط القوائم ديناميكياً
    ''' (بدون تعديل Designer — نفس نمط AddHtmlExportItem)</summary>
    Private Sub AddLoadMoreMenuItem()
        Try
            If miLoadMore IsNot Nothing Then Return
            Dim owner As ToolStrip = TryCast(SaveToolStripMenuItem.Owner, ToolStrip)
            If owner Is Nothing Then Return

            For Each existing As ToolStripItem In owner.Items
                If existing.Name = "MiLoadMorePayments" Then
                    miLoadMore = TryCast(existing, ToolStripMenuItem)
                    Return
                End If
            Next

            Dim mi As New ToolStripMenuItem("تحميل المزيد")
            mi.Name = "MiLoadMorePayments"
            AddHandler mi.Click, AddressOf MiLoadMore_Click
            owner.Items.Add(mi)
            miLoadMore = mi
            miLoadMore.Visible = False
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.AddLoadMoreMenuItem", ex)
        End Try
    End Sub

    Private Async Sub MiLoadMore_Click(sender As Object, e As EventArgs)
        Await LoadMoreRowsAsync()
    End Sub

    ''' <summary>تحميل تلقائي للدفعة التالية عند الاقتراب من نهاية القائمة
    ''' (يشمل سحب السكرول بار وعجلة الفأرة والتنقل بالأسهم)</summary>
    Private Async Sub DataGridView1_Scroll(sender As Object, e As ScrollEventArgs) Handles DataGridView1.Scroll
        Try
            If Not _hasMoreRows OrElse _isLoadingMore OrElse isReportMode Then Return
            If e.ScrollOrientation <> ScrollOrientation.VerticalScroll Then Return

            Dim first As Integer = DataGridView1.FirstDisplayedScrollingRowIndex
            If first < 0 Then Return
            Dim shown As Integer = DataGridView1.DisplayedRowCount(True)
            If first + shown >= DataGridView1.Rows.Count - 2 Then
                Await LoadMoreRowsAsync()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmPayments.DataGridView1_Scroll", ex)
        End Try
    End Sub

#End Region

#Region "UI Helpers"

    Private Sub SetAddMode()
        isEditMode = False
        currentPaymentId = -1   ' 🌟 تصفير المعرف
        currentVoucherNumber = ""
        ClearFields()
        SaveToolStripMenuItem.Text = "حفظ"
    End Sub
    Private Sub ClearFields()
        suppressSearch = True
        Try
            cmbPlayerName.Text = ""
            txtMobileNumber.Text = ""
            txtVoucherNumber.Text = ""
            txtTransferorName.Text = ""
            txtNotes.Text = ""
            dtpPaymentDate.Value = Date.Today
            dtpDueDate.Value = Date.Today.AddMonths(1)
            txtAmount.Text = ""
            cmbPaymentMethod.SelectedIndex = -1
            cmbCurrencyType.SelectedIndex = -1
            cmbPlayerName.BackColor = Color.White

            If PictureBox1.Image IsNot Nothing Then
                PictureBox1.Image.Dispose()
            End If
            PictureBox1.Image = Nothing
            _imageChanged = False
        Finally
            suppressSearch = False
        End Try
    End Sub

    Private Sub UpdateTotalAmount()
        Try
            ' 🌟 L-09: المجاميع من الاستعلام الكامل (_currentTotals) — لا من صفوف الشبكة المعروضة،
            ' فمع الترقيم التزايدي قد لا تكون كل السطور محمّلة بعد
            Dim t As CurrencyTotals = _currentTotals

            If lblTotalAmount IsNot Nothing Then
                Dim resultText As String = $"إجمالي الدفعات: {t.Count}   |   {UtilityModule.FormatCurrencyPart(t.ILS, AppConstants.Currency_ILS)}"
                If t.USD > 0 Then resultText &= $"   |   {UtilityModule.FormatCurrencyPart(t.USD, AppConstants.Currency_USD)}"
                If t.JOD > 0 Then resultText &= $"   |   {UtilityModule.FormatCurrencyPart(t.JOD, AppConstants.Currency_JOD)}"
                If _hasMoreRows Then
                    resultText &= $"   |   معروض: {CurrentLoadedRowsCount():N0} من {_totalRowsInFilter:N0}"
                End If
                lblTotalAmount.Text = resultText
            End If
        Catch ex As Exception
            DatabaseModule.LogError("UpdateTotalAmount", ex)
        End Try
    End Sub

    Private Function ValidateDataBasic() As Boolean
        Dim errMsg As String = ""

        If Not ValidationModule.IsPaymentBasicDataValid(
            txtVoucherNumber.Text, txtAmount.Text, cmbPlayerName.Text, txtTransferorName.Text,
            cmbPaymentMethod.Text, cmbCurrencyType.Text, errMsg) Then

            MessageBox.Show(errMsg, "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            If errMsg.Contains("السند") Then
                txtVoucherNumber.Focus()
            ElseIf errMsg.Contains("المبلغ") Then
                txtAmount.Focus()
            ElseIf errMsg.Contains("اللاعب") OrElse errMsg.Contains("المحول") Then
                cmbPlayerName.Focus()
            ElseIf errMsg.Contains("طريقة الدفع") Then
                cmbPaymentMethod.Focus()
            ElseIf errMsg.Contains("العملة") Then
                cmbCurrencyType.Focus()
            End If

            Return False
        End If

        Return True
    End Function
#End Region

#Region "Input Events"

    Private Sub txtAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAmount.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "."c AndAlso
           Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = "."c AndAlso DirectCast(sender, TextBox).Text.Contains(".") Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtVoucherNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtVoucherNumber.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "+"c AndAlso
           Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub cmbPlayerName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPlayerName.SelectedIndexChanged
        Try
            Dim name As String = cmbPlayerName.Text
            If playerPhones.ContainsKey(name) Then
                txtMobileNumber.Text = playerPhones(name)
            End If
            UpdatePlayerNameColor()
        Catch ex As Exception
            DatabaseModule.LogError("cmbPlayerName_SelectedIndexChanged", ex)
        End Try
    End Sub

    Private Sub cmbPlayerName_TextUpdate(sender As Object, e As EventArgs) Handles cmbPlayerName.TextUpdate
        UpdatePlayerNameColor()
    End Sub

    Private Sub UpdatePlayerNameColor()
        If archivedPlayers.Contains(cmbPlayerName.Text) Then
            cmbPlayerName.BackColor = Color.LightCoral
        Else
            cmbPlayerName.BackColor = Color.White
        End If
    End Sub

    Private Sub dtpPaymentDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpPaymentDate.ValueChanged
        dtpDueDate.Value = dtpPaymentDate.Value.AddMonths(1)
    End Sub

#End Region

#Region "Text Changed Handlers (Search Triggers)"

    Private Sub AllTextChangedHandlers(sender As Object, e As EventArgs) Handles _
        cmbPlayerName.TextChanged, txtVoucherNumber.TextChanged,
        txtTransferorName.TextChanged, txtNotes.TextChanged
        TriggerSearch()
    End Sub

    Private Sub cmbPaymentMethod_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPaymentMethod.SelectedIndexChanged
        TriggerSearch()
    End Sub

    Private Sub cmbCurrencyType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCurrencyType.SelectedIndexChanged
        TriggerSearch()
    End Sub

    Private Sub TriggerSearch()
        If searchTimer IsNot Nothing AndAlso Not suppressSearch Then
            searchTimer.Stop()
            searchTimer.Start()
        End If
    End Sub

#End Region

#Region "Export Functions"

    Private Sub ExportGeneralReportToExcel()
        Try
            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            If selectedRows.Count = 0 Then
                MessageBox.Show("الرجاء تحديد سجل واحد على الأقل للتصدير (وضع علامة ✔ بجانب السجل).", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Filter = "CSV Files|*.csv|جميع الملفات|*.*"
                sfd.DefaultExt = "csv"
                sfd.FileName = "تقرير_الايرادات_" & DateTime.Now.ToString("yyyy-MM-dd")

                If sfd.ShowDialog() <> DialogResult.OK Then Return

                Cursor = Cursors.WaitCursor
                Using sw As New StreamWriter(sfd.FileName, False, New UTF8Encoding(True))
                    sw.WriteLine("تقرير الايرادات")
                    sw.WriteLine("تاريخ التصدير: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm"))
                    sw.WriteLine()

                    ' 🌟 توحيد CSV: المنطق المشترك في UtilityModule.WriteGridCsv (كان منسوخاً في 5 نماذج)
                    ' 🌟 إصلاح: تمرير الصفوف المحددة فقط — كان يصدّر كل الصفوف بينما الرسالة تدّعي تصدير المحدد
                    UtilityModule.WriteGridCsv(DataGridView1, sw,
                        New String() {AppConstants.Grid_SeqColumn, AppConstants.Grid_SelectColumn},
                        rowsToExport:=selectedRows)

                    sw.WriteLine()
                    sw.WriteLine("إحصائيات:")
                    If lblTotalAmount IsNot Nothing Then
                        sw.WriteLine(lblTotalAmount.Text.Replace("   |   ", ","))
                    End If
                End Using

                Cursor = Cursors.Default
                MessageBox.Show($"تم تصدير بيانات {selectedRows.Count} سجل بنجاح!", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If MessageBox.Show("هل تريد فتح الملف الآن؟", "فتح الملف", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Try
                        Process.Start(New ProcessStartInfo(sfd.FileName) With {.UseShellExecute = True})
                    Catch
                        MessageBox.Show("تعذّر فتح الملف تلقائياً. يمكنك فتحه يدوياً من: " & sfd.FileName)
                    End Try
                End If
            End Using

        Catch ex As Exception
            DatabaseModule.LogError("ExportGeneralReportToExcel", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ExporteToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExporteToExcelToolStripMenuItem.Click
        ExportGeneralReportToExcel()
    End Sub

#End Region

#Region "Report & Overdue Filter Functions"

    Private Async Sub GeneralReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GeneralReportToolStripMenuItem.Click
        ' 🌟 حماية: منع الدخول لوضع التقرير مرتين — الضغطة الثانية كانت تخزّن التقرير المُجمّع
        ' في TempData وكأنه البيانات الأصلية، فيُرجع زر "رجوع" بيانات خاطئة
        If isReportMode Then
            MessageBox.Show("أنت حالياً في وضع التقرير العام — اضغط (رجوع) أولاً للخروج منه", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim errorOccurred As Boolean = False
        Try
            ClearFields()

            If DataGridView1.DataSource IsNot Nothing AndAlso TypeOf DataGridView1.DataSource Is DataTable Then
                TempData = CType(DataGridView1.DataSource, DataTable).Copy()
            End If

            DisablePaymentControls()
            SaveToolStripMenuItem.Enabled = True
            NewToolStripMenuItem.Enabled = False
            isReportMode = True
            UpdateLoadMoreState()   ' 🌟 L-09: إخفاء "تحميل المزيد" أثناء وضع التقرير

            ' 🌟 استعلام SQL موحد بدون عمود حالة الدفع (Status)
            ' 🌟 استعلام التقرير العام — الإجماليات منفصلة لكل عملة (لا خلط بين العملات)
            ' 🌟 التقرير العام — الربط بالمعرف الفريد، مع احتياط بالاسم للسجلات القديمة غير المربوطة
            ' 🌟 التقرير العام — الربط برقم الهوية (لا يتأثر بتغيير الأسماء) + يشمل اللاعبين المتأرشفين
            Dim sql As String =
                "SELECT P.Playername, P.Fone AS MobileNumber, " &
                "IFNULL(GROUP_CONCAT(Pay.VoucherNumber, ', '), '') AS VoucherNumbers, " &
                "COUNT(Pay.VoucherNumber) AS PaymentsCount, " &
                $"IFNULL(SUM(CASE WHEN IFNULL(Pay.CurrencyType, '') IN ('', '{AppConstants.Currency_ILS}') THEN Pay.Amount ELSE 0 END), 0) AS TotalILS, " &
                $"IFNULL(SUM(CASE WHEN Pay.CurrencyType = '{AppConstants.Currency_USD}' THEN Pay.Amount ELSE 0 END), 0) AS TotalUSD, " &
                $"IFNULL(SUM(CASE WHEN Pay.CurrencyType = '{AppConstants.Currency_JOD}' THEN Pay.Amount ELSE 0 END), 0) AS TotalJOD, " &
                "IFNULL(MAX(Pay.PaymentDate), '') AS LastPaymentDate " &
                "FROM (SELECT Playerid, Playername, Fone FROM Players " &
                "      UNION ALL " &
                "      SELECT Playerid, Playername, Fone FROM archive) P " &
                                $"LEFT JOIN {AppConstants.Table_Payments} Pay ON Pay.PlayerId = P.Playerid " &
                "GROUP BY P.Playerid, P.Playername " &
                "ORDER BY P.Playername ASC"
            Dim reportDt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql)

            If reportDt IsNot Nothing Then
                For Each row As DataRow In reportDt.Rows
                    Dim lastDateStr As String = UtilityModule.SafeString(row("LastPaymentDate"))
                    If Not String.IsNullOrWhiteSpace(lastDateStr) Then
                        row("LastPaymentDate") = UtilityModule.SafeDate(lastDateStr).ToString("dd/MM/yyyy")
                    End If
                Next
            End If

            DataGridView1.SuspendLayout()

            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = reportDt

            AddCustomColumns()
            AddSequenceColumn()

            SetReportHeaders()
            ColorReportRows()
            UpdateReportSummary()
            DataGridView1.ResumeLayout()

            SaveToolStripMenuItem.Text = "رجوع"
            Me.Text = "شاشة ادخال الايرادات - التقرير العام"

        Catch ex As Exception
            errorOccurred = True
            DatabaseModule.LogError("GeneralReportToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ في التقرير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If errorOccurred Then
            Await ReturnFromReportModeAsync()
        End If
    End Sub

    Private Sub SetReportHeaders()
        Dim headerMap As New Dictionary(Of String, String) From {
            {"Playername", "اسم اللاعب"}, {"MobileNumber", "رقم الجوال"},
            {"VoucherNumbers", "أرقام السندات"}, {"PaymentsCount", "عدد الدفعات"},
            {"TotalILS", "إجمالي شيكل"}, {"TotalUSD", "إجمالي دولار"}, {"TotalJOD", "إجمالي دينار"},
            {"LastPaymentDate", "آخر دفعة"}
        }
        For Each kv In headerMap
            If DataGridView1.Columns.Contains(kv.Key) Then
                DataGridView1.Columns(kv.Key).HeaderText = kv.Value
            End If
        Next

        Dim widthMap As New Dictionary(Of String, Integer) From {
            {"Playername", 200}, {"MobileNumber", 100}, {"VoucherNumbers", 200},
            {"PaymentsCount", 80}, {"TotalILS", 110}, {"TotalUSD", 110}, {"TotalJOD", 110},
            {"LastPaymentDate", 90}
        }
        For Each kv In widthMap
            If DataGridView1.Columns.Contains(kv.Key) Then
                DataGridView1.Columns(kv.Key).Width = kv.Value
            End If
        Next

        ' 🌟 تنسيق أعمدة المبالغ الثلاثة
        For Each c In {"TotalILS", "TotalUSD", "TotalJOD"}
            If DataGridView1.Columns.Contains(c) Then
                DataGridView1.Columns(c).DefaultCellStyle.Format = "#,##0.##"
                DataGridView1.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
        Next
        For Each c In {"PaymentsCount", "LastPaymentDate"}
            If DataGridView1.Columns.Contains(c) Then
                DataGridView1.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
        Next
    End Sub
    Private Sub ColorReportRows()
        ' 🌟 تم إزالة التلوين المعتمد على حالة الدفع
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.IsNewRow Then Continue For
            ' يمكن إضافة تلوين بسيط للتبديل إذا رغبت
        Next
    End Sub

    Private Sub UpdateReportSummary()
        Try
            Dim totalILS As Decimal = 0
            Dim totalUSD As Decimal = 0
            Dim totalJOD As Decimal = 0
            Dim totalPayments As Integer = 0
            Dim totalPlayers As Integer = 0

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.IsNewRow Then Continue For
                totalPlayers += 1

                If DataGridView1.Columns.Contains("PaymentsCount") AndAlso
                   row.Cells("PaymentsCount").Value IsNot Nothing AndAlso
                   Not IsDBNull(row.Cells("PaymentsCount").Value) Then
                    totalPayments += Convert.ToInt32(row.Cells("PaymentsCount").Value)
                End If

                ' 🌟 جمع كل عملة على حدة — لا خلط بين العملات
                If DataGridView1.Columns.Contains("TotalILS") AndAlso
                   row.Cells("TotalILS").Value IsNot Nothing AndAlso
                   Not IsDBNull(row.Cells("TotalILS").Value) Then
                    totalILS += Convert.ToDecimal(row.Cells("TotalILS").Value)
                End If

                If DataGridView1.Columns.Contains("TotalUSD") AndAlso
                   row.Cells("TotalUSD").Value IsNot Nothing AndAlso
                   Not IsDBNull(row.Cells("TotalUSD").Value) Then
                    totalUSD += Convert.ToDecimal(row.Cells("TotalUSD").Value)
                End If

                If DataGridView1.Columns.Contains("TotalJOD") AndAlso
                   row.Cells("TotalJOD").Value IsNot Nothing AndAlso
                   Not IsDBNull(row.Cells("TotalJOD").Value) Then
                    totalJOD += Convert.ToDecimal(row.Cells("TotalJOD").Value)
                End If
            Next

            If lblSummary IsNot Nothing Then
                Dim resultText As String = $"إجمالي اللاعبين: {totalPlayers} | عدد الدفعات: {totalPayments} | شيكل: {totalILS:#,##0.##} ₪"
                If totalUSD > 0 Then resultText &= $" | دولار: {totalUSD:#,##0.##} $"
                If totalJOD > 0 Then resultText &= $" | دينار: {totalJOD:#,##0.##} د.أ"
                lblSummary.Text = resultText
            End If
        Catch ex As Exception
            DatabaseModule.LogError("UpdateReportSummary", ex)
        End Try
    End Sub
    Private Sub PrintReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintReportToolStripMenuItem.Click
        PrintReport()
    End Sub

    Private Sub PrintReport()
        Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
        If selectedRows.Count = 0 Then
            MessageBox.Show("الرجاء تحديد سجل واحد على الأقل للطباعة (وضع علامة ✔ بجانب السجل).", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Using frm As New frmSelectColumns()
            frm.OperationType = "Print"
            frm.SourceDataGridView = DataGridView1
            ' 🌟 استثناء الأعمدة التقنية الخاصة بشاشة المدفوعات
            frm.ExcludedColumns = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "ID", "playerphoto", "VoucherImage", AppConstants.Grid_SelectColumn,
        AppConstants.Grid_SeqColumn, AppConstants.Grid_DataSource
    }
            frm.StartPosition = FormStartPosition.CenterParent

            If frm.ShowDialog() <> DialogResult.OK Then Return

            printColumns = frm.SelectedColumns
            If printColumns Is Nothing OrElse printColumns.Count = 0 Then
                MessageBox.Show("لم يتم تحديد أي عمود للطباعة", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End Using
        _printRows = selectedRows

        Using helper As New PrintHelper() With {
            .SourceGrid = DataGridView1,
            .RowsToPrint = _printRows,
            .ColumnsToPrint = printColumns,
            .ReportTitle = If(isReportMode, "التقرير العام للاعبين والمدفوعات", "تقرير المقبوضات والإيرادات"),
            .FooterSummary = If(lblSummary IsNot Nothing AndAlso Not String.IsNullOrEmpty(lblSummary.Text), lblSummary.Text, lblTotalAmount.Text)
        }
            Dim prnDoc As New Printing.PrintDocument()
            AddHandler prnDoc.PrintPage, AddressOf helper.PrintPage

            prnDoc.DefaultPageSettings.Landscape = True
            prnDoc.DefaultPageSettings.Margins = New Printing.Margins(50, 50, 50, 50)

            Dim prnPreview As New PrintPreviewDialog() With {
                .Document = prnDoc,
                .WindowState = FormWindowState.Maximized,
                .Text = "معاينة قبل الطباعة"
            }
            prnPreview.ShowDialog()

            RemoveHandler prnDoc.PrintPage, AddressOf helper.PrintPage
            prnDoc.Dispose()
            prnPreview.Dispose()
        End Using
    End Sub

    Private Sub PaymentReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PaymentReportToolStripMenuItem.Click
        Using frm As New frmRevenueReport()
            frm.ShowDialog()
        End Using
    End Sub





    Private Sub New1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles New1ToolStripMenuItem.Click
        SetAddMode()
        EnablePaymentControls()
        txtVoucherNumber.Enabled = True
        txtVoucherNumber.Focus()
    End Sub



#End Region
    ' 🌟 سند استلام شامل لعدة لاعبين
    Private Async Sub BulkPaymentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BulkPaymentToolStripMenuItem.Click
        Using frm As New frmBulkPayment()
            If frm.ShowDialog(Me) = DialogResult.OK Then
                Await LoadPaymentsAsync()
                SetAddMode()
                DisablePaymentControls()
            End If
        End Using
    End Sub

#Region "Priority 4 — HTML Export (🌟 أولوية 4)"

    ''' <summary>إضافة بند تصدير HTML ديناميكياً داخل قائمة التصدير الموجودة (بدون تعديل Designer)</summary>
    Private Sub AddHtmlExportItem()
        Try
            If ExporteToExcelToolStripMenuItem Is Nothing OrElse ExporteToExcelToolStripMenuItem.Owner Is Nothing Then
                DatabaseModule.LogInfo("Form3.AddHtmlExportItem", "لم تُعثر قائمة التصدير")
                Return
            End If
            Dim owner As ToolStripDropDown = CType(ExporteToExcelToolStripMenuItem.Owner, ToolStripDropDown)

            For Each existing As ToolStripItem In owner.Items
                If existing.Name = "MiExportHtmlPayments" Then Return
            Next

            Dim mi As New ToolStripMenuItem("تصدير HTML (طباعة / حفظ كـ PDF)")
            mi.Name = "MiExportHtmlPayments"
            AddHandler mi.Click, AddressOf MiExportHtml_Click
            owner.Items.Add(mi)
        Catch ex As Exception
            DatabaseModule.LogError("Form3.AddHtmlExportItem", ex)
        End Try
    End Sub

    Private Sub MiExportHtml_Click(sender As Object, e As EventArgs)
        Try
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("لا توجد بيانات معروضة للتصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim path As String = ReportExporter.ExportGridToHtml(
                DataGridView1, "تقرير الإيرادات (الدفعات)", "كما تظهر في الشاشة الحالية",
                landscape:=True,
                excludedColumnNames:=New String() {AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn,
                                                   AppConstants.Grid_DataSource, "VoucherImage"})

            Dim audit As New AuditService()
            audit.Log(AuditService.Act_ReportExport, "تقرير", "", "تقرير الإيرادات HTML")

            MessageBox.Show("فُتح التقرير في المتصفح — اضغط زر الطباعة لحفظه PDF." & vbCrLf & path,
                            "تم", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            DatabaseModule.LogError("Form3.MiExportHtml_Click", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

End Class
