Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Globalization ' 🌟 لضبط تحليل المبالغ مستقل عن إعدادات الويندوز
Imports System.IO
Imports System.Text
Imports System.Linq
Imports System.Threading


Public Class frmExpenses

#Region "Variables"

    Private isEditMode As Boolean = False
    Private currentVoucherNumber As String = ""
    Private WithEvents searchTimer As New System.Windows.Forms.Timer()
    Private suppressSearch As Boolean = False
    Private _isSaving As Boolean = False
    Private _cancellationTokenSource As CancellationTokenSource
    Private isClosing As Boolean = False
    ' متغيرات الإجمالي للطباعة
    Private _totalILS As Decimal = 0
    Private _totalUSD As Decimal = 0
    Private _totalJOD As Decimal = 0
    Private _totalCount As Integer = 0
    Private currentExpenseID As Integer = -1

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

    Private Async Sub frmExpenses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.SetStyle(ControlStyles.DoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)
            Me.UpdateStyles()

            DatabaseModule.EnsureExpensesTableExists()
            SetupSearchTimer()
            SetupDataGridView()
            SetAddMode()
            DisableControls()
            Me.Text = "شاشة ادخال المصروفات"

            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            AddExpenseIcon()

            ' --- تطبيق صلاحيات المستخدم --- 
            If Not UserSession.CanEdit Then SaveToolStripMenuItem.Enabled = False
            If Not UserSession.CanDelete Then DeleteToolStripMenuItem.Enabled = False
            ' ------------------------------

            ' 🌟 [أولوية 4] بند تصدير HTML (طباعة/PDF) داخل قائمة التصدير الموجودة
            AddHtmlExportItem()

            ' 🌟 الفئة الذكية + AutoComplete للبيان
            Await LoadCategoriesAsync()
            Await SetupDescriptionAutoCompleteAsync()
        Catch ex As Exception
            DatabaseModule.LogError("frmExpenses_Load", ex)
            MessageBox.Show("خطأ في تحميل النموذج: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub frmExpenses_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Await LoadExpensesAsync()
    End Sub

    Private Sub frmExpenses_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        isClosing = True
        searchTimer?.Stop()
        searchTimer?.Dispose()
        CancelSearch()
    End Sub
#End Region

#Region "Smart Helpers (Category & Description AutoComplete)"

    ''' <summary>تحميل فئات المصروفات: الجدول الذكي DropdownItems + القيم الفعلية بالقاعدة</summary>
    Private Async Function LoadCategoriesAsync() As Task
        Try
            Dim currentText As String = cmbCategory.Text

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

            cmbCategory.Text = currentText
        Catch ex As Exception
            DatabaseModule.LogError("LoadCategoriesAsync", ex)
        End Try
    End Function

    ''' <summary>AutoComplete لحقل البيان من القيم الفعلية المخزنة بجدول المصروفات</summary>
    Private Async Function SetupDescriptionAutoCompleteAsync() As Task
        Try
            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(
                "SELECT DISTINCT TRIM(Description) AS D FROM Expenses WHERE Description IS NOT NULL AND TRIM(Description) <> '' ORDER BY D")

            Dim src As New AutoCompleteStringCollection()
            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    Dim v As String = UtilityModule.SafeString(row("D"))
                    If v <> "" Then src.Add(v)
                Next
            End If

            txtDescription.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            txtDescription.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtDescription.AutoCompleteCustomSource = src
        Catch ex As Exception
            DatabaseModule.LogError("SetupDescriptionAutoCompleteAsync", ex)
        End Try
    End Function

    Private Async Function SaveCategoryToSmartListAsync() As Task
        Try
            If cmbCategory.Text.Trim() <> "" Then
                Await (New DropdownService()).AddItemAsync("ExpenseCategory", cmbCategory.Text.Trim())
            End If
        Catch ex As Exception
            DatabaseModule.LogError("SaveCategoryToSmartListAsync", ex)
        End Try
    End Function
#End Region

#Region "Initialization"

    Private Sub SetupSearchTimer()
        searchTimer.Interval = 800
        searchTimer.Enabled = False
    End Sub

    Private Sub SetupDataGridView()
        ' 🌟 الستايل الموحد — اختلافات هذه الشاشة: قراءة فقط (علامات التحديد تدار برمجياً في CellClick)
        GridHelper.ApplyCommonGridStyle(DataGridView1, multiSelect:=True, readOnlyGrid:=True, cellFontSize:=8.0F)

        AddCustomColumns()

        AddHandler DataGridView1.DataBindingComplete, AddressOf OnDataBindingComplete
        AddHandler DataGridView1.CurrentCellDirtyStateChanged, AddressOf OnCurrentCellDirtyStateChanged
        AddHandler DataGridView1.CellValueChanged, AddressOf OnCellValueChanged
        AddHandler DataGridView1.CellDoubleClick, AddressOf OnCellDoubleClick
        AddHandler DataGridView1.Enter, AddressOf OnDataGridViewEnter
        AddHandler DataGridView1.Leave, AddressOf OnDataGridViewLeave
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
                Dim seqColumn As New DataGridViewTextBoxColumn() With {
                    .Name = AppConstants.Grid_SeqColumn,
                    .HeaderText = "ت",
                    .Width = 35,
                    .ReadOnly = True,
                    .DisplayIndex = 1
                }
                seqColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns.Insert(1, seqColumn)
            End If
        Catch ex As Exception
            DatabaseModule.LogError("AddCustomColumns", ex)
        End Try
    End Sub

#End Region

#Region "DataGridView Event Handlers"

    Private Sub OnDataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        Try
            DataGridView1.SuspendLayout()
            ' 🌟 إخفاء عمود الـ ID
            If DataGridView1.Columns.Contains("ID") Then
                DataGridView1.Columns("ID").Visible = False
            End If
            ' 🌟 تغيير الترويسة للأسماء العربية وتعديل الأبعاد بعد وجود الأعمدة الفعلية
            AddRowNumbers()
            SetArabicHeaders()
            CustomizeColumnWidths()
            UpdateColumnHeaderState()
        Catch ex As Exception
            DatabaseModule.LogError("OnDataBindingComplete", ex)
        Finally
            DataGridView1.ResumeLayout()
        End Try
    End Sub
    Private Sub OnCurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub OnCellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If DataGridView1.Columns.Contains(AppConstants.Grid_SelectColumn) AndAlso
           e.ColumnIndex = DataGridView1.Columns(AppConstants.Grid_SelectColumn).Index AndAlso
           e.RowIndex >= 0 Then
            UpdateColumnHeaderState()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
            If DataGridView1.Columns(e.ColumnIndex).Name = AppConstants.Grid_SelectColumn Then
                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
                If Not row.IsNewRow Then
                    ' 🌟 إصلاح: حماية من قيمة فارغة (Nothing/DBNull) قبل قلب العلامة — نفس نمط frmPayments
                    Dim cur As Object = row.Cells(AppConstants.Grid_SelectColumn).Value
                    Dim curVal As Boolean = (cur IsNot Nothing AndAlso Not IsDBNull(cur) AndAlso Convert.ToBoolean(cur))
                    row.Cells(AppConstants.Grid_SelectColumn).Value = Not curVal
                    UpdateColumnHeaderState()
                End If
            End If
        Catch ex As Exception
            DatabaseModule.LogError("DataGridView1_CellClick", ex)
        End Try
    End Sub

    Private Sub OnCellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e.RowIndex < 0 Then Exit Sub
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            If row.IsNewRow Then Exit Sub

            suppressSearch = True
            LoadExpenseData(e.RowIndex)
            suppressSearch = False

        Catch ex As Exception
            DatabaseModule.LogError("OnCellDoubleClick", ex)
            MessageBox.Show("خطأ في عرض البيانات: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            suppressSearch = False
        End Try
    End Sub

    Private Sub OnDataGridViewEnter(sender As Object, e As EventArgs)
        suppressSearch = True
        searchTimer.Stop()
    End Sub

    Private Sub OnDataGridViewLeave(sender As Object, e As EventArgs)
        suppressSearch = False
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

    Private Sub AddRowNumbers()
        GridHelper.FillSequenceNumbers(DataGridView1)
    End Sub
    Private Sub SetArabicHeaders()
        Try
            If DataGridView1.Columns.Count = 0 Then Return
            DataGridView1.SuspendLayout()
            DataGridView1.RightToLeft = RightToLeft.Yes
            DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            For Each col As DataGridViewColumn In DataGridView1.Columns
                Select Case col.Name
                    Case AppConstants.Grid_SelectColumn : col.HeaderText = "□"
                    Case AppConstants.Grid_SeqColumn : col.HeaderText = "ت"
                    Case "VoucherNumber" : col.HeaderText = "رقم السند"
                    Case "Description" : col.HeaderText = "البيان"
                    Case "Amount" : col.HeaderText = "المبلغ"
                    Case "CurrencyType" : col.HeaderText = "العملة"
                    Case "ExpenseDate" : col.HeaderText = "التاريخ"
                    Case "Category" : col.HeaderText = "الفئة"
                    Case "Notes" : col.HeaderText = "ملاحظات"
                    Case "LastModifiedBy" : col.HeaderText = "آخر تعديل بواسطة"
                    Case "LastModifiedDate" : col.HeaderText = "تاريخ آخر تعديل"
                End Select

                Select Case col.Name
                    Case AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn, "VoucherNumber", "Amount", "CurrencyType", "ExpenseDate",
                         "Category", "LastModifiedBy"
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    Case Else
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                End Select
            Next
        Catch ex As Exception
            DatabaseModule.LogError("SetArabicHeaders", ex)
        Finally
            DataGridView1.ResumeLayout()
        End Try
    End Sub

    Private Sub CustomizeColumnWidths()
        Try
            If DataGridView1.Columns.Count = 0 Then Return
            DataGridView1.SuspendLayout()
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Dim widths As New Dictionary(Of String, Integer) From {
                {AppConstants.Grid_SelectColumn, 40}, {AppConstants.Grid_SeqColumn, 35},
                {"VoucherNumber", 90}, {"Description", 200},
                {"Amount", 90}, {"CurrencyType", 80},
                {"ExpenseDate", 90}, {"Category", 120},
                {"Notes", 200}
            }
            For Each col As DataGridViewColumn In DataGridView1.Columns
                If widths.ContainsKey(col.Name) Then col.Width = widths(col.Name)
            Next

            ' 🌟 إصلاح الفواصل العشرية وتوسيط المبلغ
            If DataGridView1.Columns.Contains("Amount") Then
                Dim amountCol As DataGridViewColumn = DataGridView1.Columns("Amount")
                amountCol.DefaultCellStyle.Format = "#,##0.##"
                amountCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
            If DataGridView1.Columns.Contains("Notes") Then
                DataGridView1.Columns("Notes").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            End If
        Catch ex As Exception
            DatabaseModule.LogError("CustomizeColumnWidths", ex)
        Finally
            DataGridView1.ResumeLayout()
        End Try
    End Sub
    Private Function GetSelectedRowsFromCheckBox() As List(Of DataGridViewRow)
        Return GridHelper.GetCheckedRows(DataGridView1)
    End Function
    Private Sub UpdateColumnHeaderState()
        GridHelper.UpdateSelectHeaderState(DataGridView1)
    End Sub
#End Region

#Region "Data Loading"
    Private Async Function LoadExpensesAsync() As Task
        Try
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            Cursor = Cursors.WaitCursor
            DataGridView1.SuspendLayout()

            Dim service As New ExpenseService()
            Dim dt As DataTable = Await service.GetAllExpensesAsync()

            If dt IsNot Nothing Then
                DataGridView1.DataSource = dt
                Me.Text = $"شاشة ادخال المصروفات - {dt.Rows.Count} سجل"
                UpdateTotalAmount()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("LoadExpensesAsync", ex)
            MessageBox.Show("خطأ في تحميل البيانات: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            DataGridView1.ResumeLayout()
            ProgressBar1.Visible = False
            Cursor = Cursors.Default
        End Try
    End Function
#End Region

#Region "Control Enable/Disable"

    Private Sub DisableControls()
        txtVoucherNumber.Enabled = False
        txtDescription.Enabled = False
        txtAmount.Enabled = False
        cmbCurrencyType.Enabled = False
        dtpExpenseDate.Enabled = False
        cmbCategory.Enabled = False
        txtNotes.Enabled = False

        SaveToolStripMenuItem.Enabled = False
        DeleteToolStripMenuItem.Enabled = False
        PrintSelectToolStripMenuItem.Enabled = True
    End Sub

    Private Sub EnableControls()
        txtVoucherNumber.Enabled = True
        txtDescription.Enabled = True
        txtAmount.Enabled = True
        cmbCurrencyType.Enabled = True
        dtpExpenseDate.Enabled = True
        cmbCategory.Enabled = True
        txtNotes.Enabled = True

        ' --- تطبيق صلاحيات المستخدم --- 
        SaveToolStripMenuItem.Enabled = UserSession.CanEdit
        DeleteToolStripMenuItem.Enabled = UserSession.CanDelete
        ' ------------------------------

        PrintSelectToolStripMenuItem.Enabled = True

        If isEditMode Then
            txtVoucherNumber.Enabled = False
        End If
    End Sub

    Private Sub SetSavingState(saving As Boolean)
        If saving Then
            SaveToolStripMenuItem.Enabled = False
            DeleteToolStripMenuItem.Enabled = False
            NewToolStripMenuItem.Enabled = False
        Else
            NewToolStripMenuItem.Enabled = True
            SaveToolStripMenuItem.Enabled = UserSession.CanEdit
            DeleteToolStripMenuItem.Enabled = UserSession.CanDelete
        End If
    End Sub

#End Region

#Region "Menu Click Handlers"

    Private Async Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If isEditMode Then
            Await UpdateExpenseAsync()
        Else
            Await InsertExpenseAsync()
        End If
    End Sub

    Private Async Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Await DeleteExpenseAsync()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    ' 🌟 سند صرف شامل لعدة مستفيدين
    Private Async Sub BulkExpenseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BulkExpenseToolStripMenuItem.Click
        Using frm As New frmBulkExpense()
            If frm.ShowDialog(Me) = DialogResult.OK Then
                Await LoadExpensesAsync()
                SetAddMode()
                DisableControls()
            End If
        End Using
    End Sub

#End Region

#Region "Data Operations - Async"
    Private Async Function InsertExpenseAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True
            SetSavingState(True)

            If Not ValidateData() Then Return

            Dim amountValue As Decimal
            ' 🌟 تحليل المبلغ بثقافة ثابتة + رسالة واضحة بدل السكوت الصامت
            If Not Decimal.TryParse(txtAmount.Text, NumberStyles.Number, CultureInfo.InvariantCulture, amountValue) OrElse amountValue <= 0 Then
                MessageBox.Show("الرجاء إدخال مبلغ صحيح أكبر من صفر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Return
            End If

            Dim service As New ExpenseService()
            If Await service.IsVoucherNumberExistsAsync(txtVoucherNumber.Text.Trim()) Then
                MessageBox.Show("رقم السند موجود مسبقاً! إذا كان سنداً شاملاً لعدة مستفيدين استخدم (صرف لعدة مستفيدين) من القائمة، وإلا استخدم رقماً مختلفاً.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtVoucherNumber.Focus()
                Return
            End If

            ' 🌟 حفظ التواريخ بصيغة ISO (yyyy-MM-dd) — العرض dd/MM/yyyy في الواجهة فقط
            Dim expenseDateStr As String = dtpExpenseDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            Dim modifiedDateStr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            Dim affectedRows As Integer = Await service.InsertExpenseAsync(
                txtVoucherNumber.Text, txtDescription.Text, amountValue,
                cmbCurrencyType.Text, expenseDateStr, cmbCategory.Text, txtNotes.Text,
                UserSession.CurrentUsername, modifiedDateStr)

            If affectedRows > 0 Then
                MessageBox.Show("تم إضافة المصروف بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' 🌟 حفظ الفئة بالجدول الذكي
                Await SaveCategoryToSmartListAsync()

                Await LoadExpensesAsync()
                SetAddMode()
                DisableControls()
                txtVoucherNumber.Focus()
            End If

        Catch ex As Exception
            DatabaseModule.LogError("InsertExpenseAsync", ex)
            MessageBox.Show("خطأ في الإضافة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function
    Private Async Function UpdateExpenseAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True
            SetSavingState(True)

            If currentExpenseID = -1 Then
                MessageBox.Show("الرجاء اختيار مصروف للتحديث", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If Not ValidateData() Then Return

            Dim amountValue As Decimal
            ' 🌟 تحليل المبلغ بثقافة ثابتة + رسالة واضحة بدل السكوت الصامت
            If Not Decimal.TryParse(txtAmount.Text, NumberStyles.Number, CultureInfo.InvariantCulture, amountValue) OrElse amountValue <= 0 Then
                MessageBox.Show("الرجاء إدخال مبلغ صحيح أكبر من صفر", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAmount.Focus()
                Return
            End If

            Dim service As New ExpenseService()
            Dim newVoucher As String = txtVoucherNumber.Text.Trim()
            If newVoucher <> currentVoucherNumber Then
                If Await service.IsVoucherNumberExistsAsync(newVoucher) Then
                    MessageBox.Show("رقم السند موجود مسبقاً! إذا كان سنداً شاملاً لعدة مستفيدين استخدم (صرف لعدة مستفيدين) من القائمة، وإلا استخدم رقماً مختلفاً.", "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtVoucherNumber.Focus()
                    Return
                End If
            End If

            ' 🌟 حفظ التواريخ بصيغة ISO (yyyy-MM-dd) — العرض dd/MM/yyyy في الواجهة فقط
            Dim expenseDateStr As String = dtpExpenseDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            Dim modifiedDateStr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            Dim affectedRows As Integer = Await service.UpdateExpenseAsync(
                txtVoucherNumber.Text, txtDescription.Text, amountValue,
                cmbCurrencyType.Text, expenseDateStr, cmbCategory.Text, txtNotes.Text, currentExpenseID,
                UserSession.CurrentUsername, modifiedDateStr)

            If affectedRows > 0 Then
                MessageBox.Show("تم تحديث المصروف بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' 🌟 حفظ الفئة بالجدول الذكي
                Await SaveCategoryToSmartListAsync()

                Await LoadExpensesAsync()
                SetAddMode()
                DisableControls()
                txtVoucherNumber.Focus()
            Else
                MessageBox.Show("لم يتم تحديث أي بيانات — تأكد من رقم السند.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("UpdateExpenseAsync", ex)
            MessageBox.Show("خطأ في التحديث: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function
    Private Async Function DeleteExpenseAsync() As Task
        If _isSaving Then Return
        Try
            Dim service As New ExpenseService()
            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            Dim idsToDelete As New List(Of Integer)

            ' 🌟 1. إذا حدد المستخدم سجلات من الجدول، نحذفها بالمعرف الفريد
            If selectedRows.Count > 0 Then
                For Each row In selectedRows
                    If DataGridView1.Columns.Contains("ID") AndAlso
                       row.Cells("ID").Value IsNot Nothing AndAlso
                       Not IsDBNull(row.Cells("ID").Value) Then
                        idsToDelete.Add(Convert.ToInt32(row.Cells("ID").Value))
                    End If
                Next

                If idsToDelete.Count = 0 Then
                    MessageBox.Show("السجلات المحددة لا تحتوي على معرف صالح.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If MessageBox.Show($"هل أنت متأكد من حذف {idsToDelete.Count} مصروف نهائياً؟", "تأكيد الحذف",
                                 MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return

                _isSaving = True
                SetSavingState(True)

                Dim successCount As Integer = 0

                ' 🌟 إصلاح: الحذف الجماعي الآن بمعاملة واحدة ذرية — كان حذف سجل بسجل بدون معاملة،
                ' فأي فصل بالمنتصف يترك الحذف ناقصاً (نصف السجلات محذوفة ونصفها لا)
                Dim deleteQueries As New List(Of Tuple(Of String, SQLiteParameter()))
                For Each eid As Integer In idsToDelete
                    deleteQueries.Add(Tuple.Create(
                        $"DELETE FROM {AppConstants.Table_Expenses} WHERE ID = @id",
                        New SQLiteParameter() {New SQLiteParameter("@id", eid)}))
                Next
                ' 🌟 إصلاح صلاحيات: كان الحذف الجماعي يتجاوز ExpenseService.DeleteExpenseAsync
                ' (التي تفرض CanDelete وتوثق في التدقيق) وينفذ DELETE مباشرة
                If Not UserSession.CanDelete Then
                    MessageBox.Show("ليست لديك صلاحية الحذف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim txOk As Boolean = Await Task.Run(Function() DatabaseModule.ExecuteTransaction(deleteQueries))
                If txOk Then successCount = idsToDelete.Count

                ' 🌟 إصلاح: فشل المعاملة كان يعرض "تم حذف 0 مصروف بنجاح" ثم يستمر كنجاح
                If Not txOk Then
                    MessageBox.Show("فشل حذف السجلات المحددة — لم يُحذف أي سجل.", "خطأ",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' 🌟 إصلاح توثيق: الحذف الجماعي كان بلا أثر في سجل التدقيق
                Try
                    Dim audit As New AuditService()
                    audit.Log(AuditService.Act_ExpenseDelete, "مصروف", String.Join(",", idsToDelete),
                              $"حذف جماعي: {idsToDelete.Count} سجل")
                Catch
                End Try

                MessageBox.Show($"تم حذف {successCount} مصروف بنجاح.", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Await LoadExpensesAsync()
                SetAddMode()
                DisableControls()
                txtVoucherNumber.Focus()
                Return
            End If

            ' 🌟 2. حذف المصروف المفتوح حالياً بالمعرف الفريد
            If currentExpenseID = -1 Then
                MessageBox.Show("الرجاء اختيار مصروف للحذف أو تحديد سجلات من الجدول", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If MessageBox.Show("هل أنت متأكد من حذف هذا المصروف نهائياً؟", "تأكيد الحذف",
                             MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                             MessageBoxDefaultButton.Button2) <> DialogResult.Yes Then Return

            _isSaving = True
            SetSavingState(True)

            Dim affected As Integer = Await service.DeleteExpenseAsync(currentExpenseID)

            If affected > 0 Then
                MessageBox.Show("تم حذف المصروف بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Await LoadExpensesAsync()
                SetAddMode()
                DisableControls()
                txtVoucherNumber.Focus()
            Else
                MessageBox.Show("فشل حذف المصروف — ربما تم حذفه مسبقاً.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("DeleteExpenseAsync", ex)
            MessageBox.Show("خطأ في الحذف: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function
#End Region

#Region "Validation"

    Private Function ValidateData() As Boolean
        Dim errMsg As String = ""

        ' --- استدعاء التحقق الموحد من الموديول ---
        If Not ValidationModule.IsExpenseDataValid(
            txtDescription.Text, txtAmount.Text, txtVoucherNumber.Text, errMsg) Then

            MessageBox.Show(errMsg, "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' تحديد أي مربع نص يجب التركيز عليه حسب رسالة الخطأ
            If errMsg.Contains("بيان") Then
                txtDescription.Focus()
            ElseIf errMsg.Contains("المبلغ") Then
                txtAmount.Focus()
            ElseIf errMsg.Contains("السند") Then
                txtVoucherNumber.Focus()
            End If

            Return False
        End If

        Return True
    End Function
#End Region

#Region "Voucher Validation - UI"

    Private Async Sub txtVoucherNumber_Leave(sender As Object, e As EventArgs) Handles txtVoucherNumber.Leave
        If isClosing Then Return
        Try
            If String.IsNullOrWhiteSpace(txtVoucherNumber.Text) Then Return
            Dim cleanVoucher As String = txtVoucherNumber.Text.Trim()

            If Not isEditMode Then
                Dim service As New ExpenseService()
                If Await service.IsVoucherNumberExistsAsync(cleanVoucher) Then
                    MessageBox.Show($"رقم السند {cleanVoucher} مستخدم مسبقاً." & vbCrLf &
                                    "إذا كان سنداً شاملاً لعدة مستفيدين → استخدم (صرف لعدة مستفيدين) من القائمة." & vbCrLf &
                                    "إلا فاستخدم رقماً مختلفاً.", "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        Catch ex As Exception
            DatabaseModule.LogError("txtVoucherNumber_Leave", ex)
        End Try
    End Sub

    Private Sub LoadExpenseData(rowIndex As Integer)
        Try
            Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)

            If Not DataGridView1.Columns.Contains("ID") Then
                MessageBox.Show("العمود ID غير موجود في الجدول", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 🌟 حفظ الـ ID الحقيقي للسند في الذاكرة
            ' 🌟 إصلاح: كان بلا فحص DBNull/Nothing — Nothing يعيد 0 بصمت فيدخل وضع تعديل وهمي
            Dim idObj As Object = row.Cells("ID").Value
            If idObj Is Nothing OrElse IsDBNull(idObj) Then Return
            currentExpenseID = Convert.ToInt32(idObj)
            If currentExpenseID <= 0 Then Return

            currentVoucherNumber = UtilityModule.SafeString(row.Cells("VoucherNumber").Value)
            txtVoucherNumber.Text = currentVoucherNumber
            txtDescription.Text = UtilityModule.SafeString(row.Cells("Description").Value)
            ' 🌟 إصلاح: عرض المبلغ بثقافة ثابتة — عمود REAL يصل Double ويُعرض بثقافة النظام
            txtAmount.Text = UtilityModule.ToDecimalSafe(row.Cells("Amount").Value).ToString("0.##", CultureInfo.InvariantCulture)
            cmbCurrencyType.Text = UtilityModule.SafeString(row.Cells("CurrencyType").Value)

            Dim parsedDate As Date = UtilityModule.SafeDate(row.Cells("ExpenseDate").Value)
            If parsedDate <> Date.MinValue Then
                dtpExpenseDate.Checked = True
                dtpExpenseDate.Value = parsedDate
            Else
                dtpExpenseDate.Checked = False
            End If

            cmbCategory.Text = UtilityModule.SafeString(row.Cells("Category").Value)
            txtNotes.Text = UtilityModule.SafeString(row.Cells("Notes").Value)

            isEditMode = True
            EnableControls()
            SaveToolStripMenuItem.Text = "تحديث"

        Catch ex As Exception
            DatabaseModule.LogError("LoadExpenseData", ex)
            MessageBox.Show("خطأ في تحميل بيانات المصروف: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        Catch ex As Exception
            If _cancellationTokenSource Is Nothing OrElse Not _cancellationTokenSource.IsCancellationRequested Then
                DatabaseModule.LogError("OnSearchTimerTick", ex)
            End If
        End Try
    End Sub

    Private Async Function PerformSmartSearchAsync(ct As Threading.CancellationToken) As Task
        Try
            If Not HasSearchCriteria() Then
                Await LoadExpensesAsync()
                Return
            End If

            Dim service As New ExpenseService()
            Dim dt As DataTable = Await service.SearchExpensesAsync(
                txtVoucherNumber.Text, txtDescription.Text,
                cmbCurrencyType.Text, cmbCategory.Text, txtNotes.Text, ct)

            If dt IsNot Nothing Then
                DataGridView1.SuspendLayout()
                DataGridView1.DataSource = dt
                Me.Text = $"شاشة ادخال المصروفات - {dt.Rows.Count} نتيجة"
                UpdateTotalAmount()
                DataGridView1.ResumeLayout()
            End If
        Catch ex As Exception
            If Not ct.IsCancellationRequested Then
                DatabaseModule.LogError("PerformSmartSearchAsync", ex)
                MessageBox.Show("خطأ في البحث: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Function

    Private Function HasSearchCriteria() As Boolean
        Return Not (String.IsNullOrWhiteSpace(txtVoucherNumber.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtDescription.Text) AndAlso
                   String.IsNullOrWhiteSpace(cmbCurrencyType.Text) AndAlso
                   String.IsNullOrWhiteSpace(cmbCategory.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtNotes.Text))
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

#Region "UI Helpers"

    Private Sub SetAddMode()
        isEditMode = False
        currentVoucherNumber = ""
        currentExpenseID = -1 ' 🌟 تصفير الـ ID
        ClearFields()
        SaveToolStripMenuItem.Text = "حفظ"
    End Sub
    Private Sub ClearFields()
        ' 🌟 كتم البحث أثناء التفريغ — التفريغ لا يعتبر معيار بحث
        suppressSearch = True
        Try
            txtVoucherNumber.Text = ""
            txtDescription.Text = ""
            txtAmount.Text = ""
            cmbCurrencyType.SelectedIndex = -1
            dtpExpenseDate.Value = Date.Today
            cmbCategory.SelectedIndex = -1
            txtNotes.Text = ""
        Finally
            suppressSearch = False
        End Try
    End Sub
    Private Sub UpdateTotalAmount()
        Try
            Dim totalILS As Decimal = 0
            Dim totalUSD As Decimal = 0
            Dim totalJOD As Decimal = 0
            Dim count As Integer = 0

            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    Dim amountObj = row.Cells("Amount").Value
                    If amountObj IsNot Nothing AndAlso Not IsDBNull(amountObj) Then
                        Dim amount As Decimal = Convert.ToDecimal(amountObj)
                        Dim currency As String = UtilityModule.SafeString(row.Cells("CurrencyType").Value)

                        Select Case currency
                            Case AppConstants.Currency_USD : totalUSD += amount
                            Case AppConstants.Currency_JOD : totalJOD += amount
                            Case Else : totalILS += amount
                        End Select
                        count += 1
                    End If
                End If
            Next

            If lblTotalAmount IsNot Nothing Then
                lblTotalAmount.Text = $"إجمالي المصروفات: {count}   |   " &
                                      $"شيكل: {totalILS:#,##0.##} ₪   |   " &
                                      $"دولار: {totalUSD:#,##0.##} $   |   " &
                                      $"دينار: {totalJOD:#,##0.##} د.أ"
            End If
        Catch ex As Exception
            DatabaseModule.LogError("UpdateTotalAmount", ex)
        End Try
    End Sub
#End Region

#Region "Text Changed Handlers"

    Private Sub AllTextChangedHandlers(sender As Object, e As EventArgs) Handles _
        txtVoucherNumber.TextChanged, txtDescription.TextChanged, txtNotes.TextChanged

        If searchTimer IsNot Nothing AndAlso Not suppressSearch Then
            searchTimer.Stop()
            searchTimer.Start()
        End If
    End Sub

    Private Sub cmbCurrencyType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCurrencyType.SelectedIndexChanged
        If searchTimer IsNot Nothing AndAlso Not suppressSearch Then
            searchTimer.Stop()
            searchTimer.Start()
        End If
    End Sub

    Private Sub cmbCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCategory.SelectedIndexChanged
        If searchTimer IsNot Nothing AndAlso Not suppressSearch Then
            searchTimer.Stop()
            searchTimer.Start()
        End If
    End Sub

#End Region

#Region "Printing - Enhanced with PrintHelper"

    Private Sub CalculatePrintTotals(rows As List(Of DataGridViewRow))
        ' 🌟 توحيد حساب المجاميع: UtilityModule.SumCurrencyTotalsFromRows (كان مكرراً في 4 نماذج)
        Dim t As CurrencyTotals = UtilityModule.SumCurrencyTotalsFromRows(rows, "Amount", "CurrencyType")
        _totalILS = t.ILS
        _totalUSD = t.USD
        _totalJOD = t.JOD
        _totalCount = t.Count
    End Sub
    Private Sub PrintSelectedRows()
        Try
            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            If selectedRows.Count = 0 Then
                MessageBox.Show("الرجاء تحديد سجل واحد على الأقل للطباعة (وضع علامة ✔ بجانب السجل).", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using frm As New frmSelectColumns()
                frm.OperationType = "Print"
                frm.SourceDataGridView = DataGridView1
                frm.StartPosition = FormStartPosition.CenterParent
                If frm.ShowDialog() = DialogResult.OK Then
                    Dim printCols = frm.SelectedColumns
                    Dim printRows = selectedRows.OrderBy(Function(r) UtilityModule.SafeDate(r.Cells("ExpenseDate").Value)).ToList()

                    CalculatePrintTotals(printRows)
                    ' 🌟 تهيئة موحدة عبر FormatCurrencyPart
                    Dim summary As String = $"عدد السجلات: {_totalCount} | {UtilityModule.FormatCurrencyPart(_totalILS, AppConstants.Currency_ILS)} | {UtilityModule.FormatCurrencyPart(_totalUSD, AppConstants.Currency_USD)} | {UtilityModule.FormatCurrencyPart(_totalJOD, AppConstants.Currency_JOD)}"

                    Using helper As New PrintHelper() With {
                        .SourceGrid = DataGridView1,
                        .RowsToPrint = printRows,
                        .ColumnsToPrint = printCols,
                        .ReportTitle = "تقرير المصروفات (المحدد)",
                        .FooterSummary = summary
                    }
                        helper.ShowPreview()
                    End Using
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("PrintSelectedRows", ex)
            MessageBox.Show("خطأ في الطباعة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub PrintSelectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintSelectToolStripMenuItem.Click
        PrintSelectedRows()
    End Sub

#End Region

#Region "Export Functions"

    Private Sub ExportToExcel()
        Try
            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            If selectedRows.Count = 0 Then
                MessageBox.Show("الرجاء تحديد سجل واحد على الأقل للتصدير (وضع علامة ✔ بجانب السجل).", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Filter = "CSV Files|*.csv"
                sfd.DefaultExt = "csv"
                sfd.FileName = "تقرير المصروفات " & DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)

                If sfd.ShowDialog() = DialogResult.OK Then
                    Cursor = Cursors.WaitCursor
                    Dim utf8WithBom As New UTF8Encoding(True)

                    Using sw As New StreamWriter(sfd.FileName, False, utf8WithBom)
                        sw.WriteLine("تقرير المصروفات")
                        sw.WriteLine("تاريخ التصدير: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm"))
                        sw.WriteLine()

                        ' 🌟 توحيد CSV: المنطق المشترك في UtilityModule.WriteGridCsv (كان منسوخاً في 5 نماذج)
                        ' 🌟 إصلاح: تمرير الصفوف المحددة فقط — كان يصدّر كل الصفوف بينما الرسالة تدّعي تصدير المحدد
                        UtilityModule.WriteGridCsv(DataGridView1, sw,
                            New String() {AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn},
                            rowsToExport:=selectedRows,
                            dateColumnNames:=New String() {"ExpenseDate"},
                            amountColumnNames:=New String() {"Amount"})

                        CalculatePrintTotals(selectedRows)
                        sw.WriteLine()
                        sw.WriteLine("إحصائيات السجلات المحددة:")

                        ' 🌟 تهيئة موحدة عبر FormatCurrencyPart (تظهر العملات ذات المبالغ الموجبة فقط)
                        Dim statsParts As New List(Of String)
                        If _totalILS > 0 Then statsParts.Add(UtilityModule.FormatCurrencyPart(_totalILS, AppConstants.Currency_ILS))
                        If _totalUSD > 0 Then statsParts.Add(UtilityModule.FormatCurrencyPart(_totalUSD, AppConstants.Currency_USD))
                        If _totalJOD > 0 Then statsParts.Add(UtilityModule.FormatCurrencyPart(_totalJOD, AppConstants.Currency_JOD))
                        Dim stats As String = String.Join(" | ", statsParts)

                        sw.WriteLine(stats)
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
                End If
            End Using
        Catch ex As Exception
            Cursor = Cursors.Default
            DatabaseModule.LogError("ExportToExcel", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        ExportToExcel()
    End Sub

#End Region

#Region "Priority 4 — HTML Export (🌟 أولوية 4)"

    ''' <summary>إضافة بند تصدير HTML ديناميكياً داخل قائمة التصدير الموجودة (بدون تعديل Designer)</summary>
    Private Sub AddHtmlExportItem()
        Try
            If ExportToExcelToolStripMenuItem Is Nothing OrElse ExportToExcelToolStripMenuItem.Owner Is Nothing Then
                DatabaseModule.LogInfo("frmExpenses.AddHtmlExportItem", "لم تُعثر قائمة التصدير")
                Return
            End If
            ' 🌟 إصلاح خطأ ترجمة: Owner يعيد ToolStrip — تحويل صريح إلزامي مع Option Strict On
            Dim owner As ToolStripDropDown = CType(ExportToExcelToolStripMenuItem.Owner, ToolStripDropDown)

            For Each existing As ToolStripItem In owner.Items
                If existing.Name = "MiExportHtmlExpenses" Then Return
            Next

            Dim mi As New ToolStripMenuItem("تصدير HTML (طباعة / حفظ كـ PDF)")
            mi.Name = "MiExportHtmlExpenses"
            AddHandler mi.Click, AddressOf MiExportHtml_Click
            owner.Items.Add(mi)
        Catch ex As Exception
            DatabaseModule.LogError("frmExpenses.AddHtmlExportItem", ex)
        End Try
    End Sub

    Private Sub MiExportHtml_Click(sender As Object, e As EventArgs)
        Try
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("لا توجد بيانات معروضة للتصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim path As String = ReportExporter.ExportGridToHtml(
                DataGridView1, "تقرير المصروفات", "كما تظهر في الشاشة الحالية",
                landscape:=True,
                excludedColumnNames:=New String() {AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn, AppConstants.Grid_DataSource})

            Dim audit As New AuditService()
            audit.Log(AuditService.Act_ReportExport, "تقرير", "", "تقرير المصروفات HTML")

            MessageBox.Show("فُتح التقرير في المتصفح — اضغط زر الطباعة لحفظه PDF." & vbCrLf & path,
                            "تم", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            DatabaseModule.LogError("frmExpenses.MiExportHtml_Click", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Report Forms"
    Private Sub ExpensesReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExpensesReportToolStripMenuItem.Click
        Try
            Using frm As New frmExpensesReport()
                frm.ShowDialog(Me)
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("ExpensesReportToolStripMenuItem_Click", ex)
            MessageBox.Show("تعذر فتح نموذج التقرير بسبب الخطأ التالي:" & vbCrLf & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

    Private Sub txtAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAmount.KeyPress
        Dim txt As TextBox = TryCast(sender, TextBox)
        If txt Is Nothing Then Return

        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "." AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = "." AndAlso txt.Text.Contains(".") Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtVoucherNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtVoucherNumber.KeyPress
        Dim txt As TextBox = TryCast(sender, TextBox)
        If txt Is Nothing Then Return

        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "+" AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub AddExpenseIcon()
        Try
            Dim lblIcon As New Label() With {
                .Font = New Font("Segoe MDL2 Assets", 48, FontStyle.Regular),
                .Text = ChrW(&HE72D),
                .Location = New Point(20, 10),
                .AutoSize = True,
                .ForeColor = Color.DarkRed
            }
            Me.Controls.Add(lblIcon)
        Catch ex As Exception
            DatabaseModule.LogInfo("AddExpenseIcon", "تعذر إضافة الأيقونة: " & ex.Message)
        End Try
    End Sub

    Private Async Sub New1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles New1ToolStripMenuItem.Click
        SetAddMode()
        EnableControls()
        txtVoucherNumber.Enabled = True
        txtVoucherNumber.Focus()
        Await LoadExpensesAsync()
    End Sub
End Class
