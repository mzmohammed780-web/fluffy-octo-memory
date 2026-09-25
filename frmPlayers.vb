Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Linq

' ═════════════════════════════════════════════════════════════════
' 🌟 ملاحظة توثيق (أولوية 3): هذا الملف هو شاشة اللاعبين الرئيسية —
'    الشاشة الأولى بعد تسجيل الدخول (يفتحها frmLogin بـ New frmPlayers()).
'    اسم الكلاس "frmPlayers" موروث من الإنشاء التلقائي وغير معبّر، ولا يُعاد
'    تسميته يدوياً في الملف لأن المُصمم (Designer.vb) مرتبط بالاسم؛
'    لإعادة التسمية بأمان: في Visual Studio كليك يمين على frmPlayers في
'    مستكشف الحلول ← Rename (سيحدّث الكلاس ومراجعه في frmLogin تلقائياً).
' ═════════════════════════════════════════════════════════════════
Public Class frmPlayers

#Region "Variables"

    Private isEditMode As Boolean = False
    Private currentPlayerID As String = ""
    Private tempImagePath As String = ""
    Private suppressSearch As Boolean = False
    Private searchResultCount As Integer = 0
    Private _isSaving As Boolean = False
    Private _imageChanged As Boolean = False
    Private _cancellationTokenSource As CancellationTokenSource
    ' 🌟 عدّاد حماية من سباق تحميل الصور (نتيجة تحميل قديم لا تُعرض فوق الجديدة)
    Private _photoLoadId As Integer = 0
    Private WithEvents playerPhotoPrintDoc As New Printing.PrintDocument()
    Private isClosing As Boolean = False
    ' 🌟 L-10: علم حماية يمنع تكرار النسخ الاحتياطي الخلفي عند إعادة طلب الإغلاق
    Private autoBackupStarted As Boolean = False
    Private printColumns As New List(Of String)

    ' عناصر شريط الإحصائيات (Dashboard) البرمجي
    Private pnlDashboard As FlowLayoutPanel
    Private lblDashPlayers As Label
    Private lblDashRevenue As Label
    Private lblDashExpense As Label
    Private lblDashNet As Label


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

    Private Async Sub frmPlayers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.SetStyle(ControlStyles.DoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)
            Me.UpdateStyles()
            ' 🌟 إصلاح: Initialize كانت تعرض MessageBox من داخل خيط خلفي (سلوك غير مدعوم) —
            ' الآن صامتة والرسالة تُعرض هنا من خيط الواجهة
            If Not Await Task.Run(Function() DatabaseModule.Initialize(False)) Then
                MessageBox.Show("فشل في تهيئة قاعدة البيانات. سيتم إغلاق التطبيق." & vbCrLf &
                                DatabaseModule.LastError, "خطأ فادح",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End If

            ' 🌟 [تعديل 4] نسخة احتياطية تلقائية قبل أي ترحيل —
            ' لو فشل ترحيل الجداول بنص الطريق تبقى قاعدة البيانات الأصلية محفوظة
            Try
                DatabaseModule.PerformAutoBackup(14)
            Catch
                ' فشل النسخ الاحتياطي لا يوقف الإقلاع — ستُؤخذ نسخة أخرى عند الإغلاق
            End Try

            Try
                DatabaseModule.EnsurePlayersTableExists()
                DatabaseModule.EnsureArchiveTableExists()
                DatabaseModule.EnsurePaymentsTableExists()
                DatabaseModule.EnsureExpensesTableExists()
                DatabaseModule.EnsureUsersTableExists()
                DatabaseModule.EnsureDropdownItemsTableExists()   ' 🌟 عناصر القوائم المنسدلة

                ' 🌟 ترحيل لمرة واحدة: Payment → Revenues
                DatabaseModule.RenamePaymentToRevenues()
            Catch ex As Exception
                DatabaseModule.LogError("frmPlayers_Load - EnsureTables", ex)
                MessageBox.Show("فشل في إنشاء جداول قاعدة البيانات: " & ex.Message & vbCrLf & "سيتم إغلاق التطبيق.",
                              "خطأ فادح", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End Try
            ' 🌟 M-01: ترحيل التواريخ يعمل مرة واحدة فقط عبر علم الإتمام في AppSettings —
            ' بدل مسح سبعة أعمدة تاريخ بكل إقلاع، صار ما بيتعمل بكل إقلاع هو SELECT واحد خفيف
            Try
                AppSettingsStore.EnsureAppSettingsTableExists()
                If AppSettingsStore.GetSetting(AppSettingsStore.Key_Migration_DatesIso_Done, "0") <> "1" Then
                    DatabaseModule.MigrateDatesToIsoFormat()
                End If
            Catch ex As Exception
                DatabaseModule.LogError("frmPlayers_Load - MigrateDates", ex)
            End Try

            ' 🌟 [أولوية 4] تهيئة جداول المزايا الأمنية الجديدة + توثيق تشغيل النظام
            Try
                AppSettingsStore.EnsureAppSettingsTableExists()
                DatabaseModule.TightenDataFolderPermissionsOnce()   ' 🔴 H-02: تقييد صلاحيات مجلد البيانات (مرة واحدة — تثبيتات قديمة)
                Dim auditStartup As New AuditService()
                auditStartup.EnsureAuditLogTableExists()
                auditStartup.Log("تشغيل النظام", "جلسة", "",
                    $"مستخدم الجلسة: {If(String.IsNullOrWhiteSpace(UserSession.CurrentUsername), "غير معروف", UserSession.CurrentUsername)}")

                ' تنظيف السجلات الأقدم من فترة الاحتفاظ في الخلفية — لا يبطئ الإقلاع أبداً
                Dim keepDays As Integer = AppSettingsStore.GetIntSetting(AppSettingsStore.Key_AuditRetentionDays, 365, 30, 3650)
                Dim purgeSvc As New AuditService()
                Dim purgeTask As Task = Task.Run(
                    Async Function()
                        Await purgeSvc.PurgeOlderThanAsync(keepDays)
                    End Function)
            Catch ex As Exception
                DatabaseModule.LogError("frmPlayers_Load - Priority4Init", ex)
            End Try

            SetupDataGridView()
            SetupInputRestrictions()
            UpdateSelectedCount()
            Me.Text = "نظام إدارة اللاعبين - الإصدار 2.0"

            TimerSearch.Enabled = False
            UpdateImageButtonsState()
            DisableAllControls()
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee

            ' تطبيق صلاحيات المستخدم
            SaveToolStripMenuItem.Enabled = UserSession.CanEdit
            DeleteToolStripMenuItem.Enabled = UserSession.CanDelete

            ' 🌟 [أولوية 4] قفل الخمول التلقائي + بنود القائمة الأمنية الجديدة
            IdleLockManager.Instance.Startup()
            AddSecurityMenuItems()

        Catch ex As Exception
            DatabaseModule.LogError("frmPlayers_Load", ex)
            MessageBox.Show("خطأ في تحميل النموذج: " & ex.Message, "خطأ",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub frmPlayers_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        SetupDashboardUI()
        Await LoadPlayersOnlyAsync()
        Await LoadCustomDropdownItemsAsync()   ' 🌟 عناصر القوائم المخصصة من القاعدة
        Await RefreshDashboardAsync()
    End Sub
    Private Sub frmPlayers_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        isClosing = True
        CancelSearch()
        TimerSearch.Enabled = False
        SafeDisposeImage()

        ' 🌟 تشغيل النسخ الاحتياطي التلقائي فقط عند إغلاق البرنامج بالكامل، وليس عند تسجيل الخروج
        ' 🌟 L-10: عند إغلاق المستخدم تنطلق النسخة على خيط خلفي حتى لا يمدد زمن الخروج مع كبر القاعدة —
        '    علم الحماية يمنع التكرار عند إعادة طلب الإغلاق، وأي سبب آخر (إطفاء/إنهاء مهمة) يبقى متزامناً كما هو
        If Not UserSession.IsLoggingOut Then
            If autoBackupStarted Then
                ' النسخة انطلقت على الخيط الخلفي قبل لحظات — لا تكررها ولا تحجب الإغلاق
            ElseIf e.CloseReason = CloseReason.UserClosing Then
                autoBackupStarted = True
                e.Cancel = True
                Me.Cursor = Cursors.WaitCursor
                Dim backupThread As New System.Threading.Thread(
                    Sub()
                        Try
                            DatabaseModule.PerformAutoBackup(14)
                        Catch
                        End Try
                        ' إعادة طلب الإغلاق من خيط الواجهة بعد اكتمال النسخ
                        Try
                            Me.BeginInvoke(New Action(Sub() Me.Close()))
                        Catch
                        End Try
                    End Sub)
                backupThread.IsBackground = False
                backupThread.Start()
                ' 🌟 إصلاح v2: مهلة 30 ثانية — لو تعطّل النسخ (قرص ممتلئ/ملف مقفول)
                ' لا يتعلّق الإغلاق للأبد
                If Not backupThread.Join(TimeSpan.FromSeconds(30)) Then
                    DatabaseModule.LogWarn("frmPlayers.FormClosing", "انتهت مهلة النسخ الاحتياطي — إغلاق متابعة على أي حال")
                End If
                Return
            Else
                DatabaseModule.PerformAutoBackup(14)
            End If
        End If

        ' 🌟 [أولوية 4] إيقاف مراقبة الخمول + توثيق نهاية الجلسة في سجل التدقيق
        IdleLockManager.Instance.Shutdown()
        Dim auditClose As New AuditService()
        If UserSession.IsLoggingOut Then
            auditClose.Log(AuditService.Act_Logout, "جلسة", "", "")
        Else
            auditClose.Log(AuditService.Act_AppClosed, "جلسة", "", "")
        End If
    End Sub

#End Region

#Region "Dashboard UI & Logic"

    Private Sub SetupDashboardUI()
        Try
            If pnlDashboard IsNot Nothing Then Return

            ' 🌟 FlowLayoutPanel: العناصر تتسلسل تلقائياً من اليمين لليسار
            ' ولا تتراكب مهما كان عرض النافذة — وعلى النوافذ الضيقة تلتف لسطر ثانٍ بدل أن تختفي
            pnlDashboard = New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .AutoSize = True,
                .BackColor = Color.FromArgb(40, 48, 68),
                .Font = New Font("Arial", 9.5!, FontStyle.Bold),
                .RightToLeft = RightToLeft.Yes,
                .FlowDirection = FlowDirection.LeftToRight,   ' مع RTL يتدفق تلقائياً من اليمين لليسار
                .Padding = New Padding(12, 7, 12, 5)
            }

            lblDashPlayers = New Label() With {.Text = "👥 لاعبون نشطون: ...", .ForeColor = Color.White, .AutoSize = True, .Margin = New Padding(10, 0, 10, 0)}
            lblDashRevenue = New Label() With {.Text = "💰 إيراد الشهر: ...", .ForeColor = Color.LightGreen, .AutoSize = True, .Margin = New Padding(10, 0, 10, 0)}
            lblDashExpense = New Label() With {.Text = "💸 مصروف الشهر: ...", .ForeColor = Color.LightPink, .AutoSize = True, .Margin = New Padding(10, 0, 10, 0)}
            lblDashNet = New Label() With {.Text = "⚖️ الصافي: ...", .ForeColor = Color.LightCyan, .AutoSize = True, .Margin = New Padding(10, 0, 10, 0)}

            pnlDashboard.Controls.AddRange({lblDashPlayers, lblDashRevenue, lblDashExpense, lblDashNet})
            Me.Controls.Add(pnlDashboard)
            pnlDashboard.BringToFront()
        Catch ex As Exception
            DatabaseModule.LogError("SetupDashboardUI", ex)
        End Try
    End Sub
    Private Async Function RefreshDashboardAsync() As Task
        Try
            Dim srv As New DashboardService()
            Dim stats = Await srv.GetStatsAsync()

            If lblDashPlayers IsNot Nothing Then
                lblDashPlayers.Text = $"👥 لاعبون نشطون: {stats.ActivePlayers}"
                lblDashRevenue.Text = $"💰 إيراد الشهر: {stats.MonthRevenuesILS:#,##0.##} ₪"
                lblDashExpense.Text = $"💸 مصروف الشهر: {stats.MonthExpensesILS:#,##0.##} ₪"
                lblDashNet.Text = $"⚖️ الصافي: {stats.MonthNetILS:#,##0.##} ₪"


            End If
        Catch ex As Exception
            DatabaseModule.LogError("RefreshDashboardAsync", ex)
        End Try
    End Function

#End Region

#Region "DataGridView Setup"

    Private Sub SetupDataGridView()
        ' 🌟 الستايل الموحد — الافتراضيات تناسب هذه الشاشة: تحديد متعدد + قابل للتعديل + خط 8
        GridHelper.ApplyCommonGridStyle(DataGridView1)

        AddCustomColumns()

        AddHandler DataGridView1.DataBindingComplete, AddressOf OnDataBindingComplete
        AddHandler DataGridView1.RowPrePaint, AddressOf OnRowPrePaint
        AddHandler DataGridView1.CurrentCellDirtyStateChanged, AddressOf OnCurrentCellDirtyStateChanged
        AddHandler DataGridView1.CellValueChanged, AddressOf OnCellValueChanged
        AddHandler DataGridView1.CellDoubleClick, AddressOf OnCellDoubleClick
        AddHandler DataGridView1.Enter, AddressOf OnDataGridViewEnter
        AddHandler DataGridView1.Leave, AddressOf OnDataGridViewLeave
        AddHandler DataGridView1.ColumnHeaderMouseClick, AddressOf DataGridView1_ColumnHeaderMouseClick
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
            If DataGridView1.Columns.Contains(AppConstants.Grid_DataSource) Then
                DataGridView1.Columns(AppConstants.Grid_DataSource).Visible = False
            End If
            If DataGridView1.Columns.Contains(AppConstants.Col_PlayerPhoto) Then
                DataGridView1.Columns(AppConstants.Col_PlayerPhoto).Visible = False
            End If
            AddRowNumbers()
            SetArabicHeaders()
            CustomizeColumnWidths()

            ' 🌟 منع التعديل الوهمي: خلايا البيانات للقراءة فقط والتعديل الحقيقي عبر النموذج وزر الحفظ،
            ' أما عمود التحديد ✔ فيبقى قابلاً للنقر (كان بإمكان المستخدم تعديل خلية بالجريد ويظن أنها انحفظت
            ' بينما التعديل يضيع عند إعادة تحميل البيانات)
            For Each col As DataGridViewColumn In DataGridView1.Columns
                col.ReadOnly = (col.Name <> AppConstants.Grid_SelectColumn)
            Next
            UpdateColumnHeaderState()
            ColorizeRows()
        Catch ex As Exception
            DatabaseModule.LogError("OnDataBindingComplete", ex)
        Finally
            DataGridView1.ResumeLayout()
        End Try
    End Sub

    Private Sub OnRowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs)
        ' تم إفراغ هذه الدالة لمنع التقطيع أثناء التمرير (Scrolling)
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
            UpdateSelectedCount()
            UpdateColumnHeaderState()
        End If
    End Sub

    Private Async Sub OnCellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Try
            Await LoadSelectedPlayerDataAsync(e.RowIndex)

            If DataGridView1.Columns.Contains(AppConstants.Grid_DataSource) Then
                Dim sourceValue = DataGridView1.Rows(e.RowIndex).Cells(AppConstants.Grid_DataSource).Value
                If sourceValue IsNot Nothing AndAlso sourceValue.ToString() = AppConstants.Source_Archive Then
                    ' إذا كان السند مؤرشف، نغلق كل شيء
                    DisableAllControls()
                    RestorefromarchifeToolStripMenuItem.Enabled = UserSession.CanEdit ' 🌟 إصلاح: كان متاحاً لغير المصرح ويفشل بصمت داخل الخدمة
                Else
                    ' إذا كان نشطاً، نفتح كل شيء (ومربع الهوية تم تفعيله بالفعل داخل SetEditMode)
                    EnableAllControls()
                    ' 🌟 تم حذف السطر الذي كان يعطل txtPlayerID هنا
                End If
            End If
        Catch ex As Exception
            DatabaseModule.LogError("OnCellDoubleClick", ex)
            MessageBox.Show("خطأ في تحميل بيانات اللاعب: " & ex.Message, "خطأ",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub OnDataGridViewEnter(sender As Object, e As EventArgs)
        suppressSearch = True
        TimerSearch.Enabled = False
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
            UpdateSelectedCount()
            DataGridView1.ResumeLayout()
        Catch ex As Exception
            DatabaseModule.LogError("DataGridView1_ColumnHeaderMouseClick", ex)
        End Try
    End Sub
#End Region

#Region "Data Loading"

    Private Async Function LoadPlayersOnlyAsync() As Task
        Try
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            Cursor = Cursors.WaitCursor
            DataGridView1.SuspendLayout()

            Dim service As New PlayerService()
            Dim dt As DataTable = Await service.GetAllPlayersAsync()

            If dt IsNot Nothing Then
                ' 🌟 تم إزالة DataSource = Nothing لتسريع العرض
                DataGridView1.DataSource = dt
                searchResultCount = dt.Rows.Count
                Me.Text = $"نظام إدارة اللاعبين "
                UpdateSearchResultLabel()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("LoadPlayersOnlyAsync", ex)
            MessageBox.Show("خطأ في تحميل البيانات: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            DataGridView1.ResumeLayout()
            ProgressBar1.Visible = False
            Cursor = Cursors.Default
        End Try
    End Function

    Private Async Function LoadArchiveOnlyAsync() As Task
        Try
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            Cursor = Cursors.WaitCursor
            DataGridView1.SuspendLayout()

            Dim service As New PlayerService()
            Dim dt As DataTable = Await service.GetAllArchivedPlayersAsync()

            If dt IsNot Nothing Then
                ' 🌟 تم إزالة DataSource = Nothing
                DataGridView1.DataSource = dt
                searchResultCount = dt.Rows.Count
                Me.Text = "نظام إدارة اللاعبين - الأرشيف"
                UpdateSearchResultLabel()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("LoadArchiveOnlyAsync", ex)
            MessageBox.Show("خطأ في تحميل الأرشيف: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            DataGridView1.ResumeLayout()
            ProgressBar1.Visible = False
            Cursor = Cursors.Default
        End Try
    End Function

    Private Async Function LoadSelectedPlayerDataAsync(rowIndex As Integer) As Task
        Try
            Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)
            Dim sourceValue As Object = Nothing
            If DataGridView1.Columns.Contains(AppConstants.Grid_DataSource) Then
                sourceValue = row.Cells(AppConstants.Grid_DataSource).Value
            End If

            suppressSearch = True
            Try
                currentPlayerID = SafeGetString(row.Cells(AppConstants.Col_PlayerId).Value)
                txtPlayerID.Text = currentPlayerID
                txtPlayerName.Text = SafeGetString(row.Cells(AppConstants.Col_PlayerName).Value)

                Dim bDate = row.Cells(AppConstants.Col_BirthDate).Value
                If bDate IsNot Nothing AndAlso Not IsDBNull(bDate) AndAlso Not String.IsNullOrWhiteSpace(bDate.ToString()) Then
                    Dim bd As Date = UtilityModule.SafeDate(bDate)
                    ' 🌟 إصلاح: تاريخ تالف يُحلَّل لـ Date.MinValue (0001) كان يرمي استثناء MinDate على الـ DateTimePicker
                    If bd <> Date.MinValue Then
                        dtpBirthDate.Checked = True
                        dtpBirthDate.Value = bd
                    Else
                        dtpBirthDate.Checked = False
                    End If
                Else
                    dtpBirthDate.Checked = False
                End If

                txtFatherName.Text = SafeGetString(row.Cells(AppConstants.Col_FatherName).Value)
                txtFatherID.Text = SafeGetString(row.Cells(AppConstants.Col_FatherId).Value)
                txtWifeName.Text = SafeGetString(row.Cells(AppConstants.Col_WifeName).Value)
                txtWifeID.Text = SafeGetString(row.Cells(AppConstants.Col_WifeId).Value)
                txtFone.Text = SafeGetString(row.Cells(AppConstants.Col_Fone).Value)
                txtAltFone.Text = SafeGetString(row.Cells(AppConstants.Col_AltFone).Value)
                txtAddress.Text = SafeGetString(row.Cells(AppConstants.Col_Address).Value)

                Dim dIn = row.Cells(AppConstants.Col_DateIn).Value
                If dIn IsNot Nothing AndAlso Not IsDBNull(dIn) AndAlso Not String.IsNullOrWhiteSpace(dIn.ToString()) Then
                    Dim dInDate As Date = UtilityModule.SafeDate(dIn)
                    ' 🌟 إصلاح: نفس حماية Date.MinValue
                    If dInDate <> Date.MinValue Then
                        dtpDateIn.Checked = True
                        dtpDateIn.Value = dInDate
                    Else
                        dtpDateIn.Checked = False
                    End If
                Else
                    dtpDateIn.Checked = False
                End If

                txtNotes.Text = SafeGetString(row.Cells(AppConstants.Col_Notes).Value)
                cmbGender.Text = SafeGetString(row.Cells(AppConstants.Col_Gender).Value)
                cmbMaritalStatus.Text = SafeGetString(row.Cells(AppConstants.Col_MaritalStatus).Value)
                cmbTeam.Text = SafeGetString(row.Cells(AppConstants.Col_Team).Value)
                cmbRolle.Text = SafeGetString(row.Cells(AppConstants.Col_Rolle).Value)
                cmbJobTitle.Text = SafeGetString(row.Cells(AppConstants.Col_JobTitle).Value)
            Finally
                suppressSearch = False
            End Try

            _imageChanged = False
            Await LoadPlayerPhotoAsync(currentPlayerID, sourceValue)
            UpdateImageButtonsState()

            If sourceValue IsNot Nothing AndAlso sourceValue.ToString() = AppConstants.Source_Archive Then
                SetArchiveViewMode()
            Else
                SetEditMode()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("LoadSelectedPlayerDataAsync", ex)
            Throw
        End Try
    End Function

    Private Async Function LoadPlayerPhotoAsync(playerID As String, Optional source As Object = Nothing) As Task
        _photoLoadId += 1
        Dim thisLoadId As Integer = _photoLoadId
        Try
            SafeDisposeImage()

            If String.IsNullOrWhiteSpace(playerID) Then
                SetDefaultImage()
                Return
            End If

            Cursor = Cursors.WaitCursor

            Dim service As New PlayerService()
            Dim sourceStr As String = If(source IsNot Nothing, source.ToString(), "")
            Dim photoBytes As Byte() = Await service.GetPlayerPhotoAsync(playerID, sourceStr)

            ' 🌟 تجاهل النتيجة إذا بدأ تحميل صورة أخرى أثناء الانتظار (حماية من سباق التحميل)
            If thisLoadId <> _photoLoadId Then Return

            ' 🌟 استخدام الدالة الآمنة لمنع خطأ GDI+
            If photoBytes IsNot Nothing AndAlso photoBytes.Length > 0 Then
                Dim bmp As Image = UtilityModule.ByteArrayToImage(photoBytes)
                If bmp IsNot Nothing Then
                    SafeDisposeImage()
                    PictureBox1.Image = bmp
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                End If
                Return
            End If

            SetDefaultImage()

        Catch ex As Exception
            DatabaseModule.LogError("LoadPlayerPhotoAsync", ex)
            SetDefaultImage()
        Finally
            Cursor = Cursors.Default
        End Try
    End Function

    Private Sub SetDefaultImage()
        SafeDisposeImage()
        PictureBox1.Image = Nothing
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
    End Sub

#End Region

#Region "Search - Players & Archive Combined"

    Private Async Sub TimerSearch_Tick(sender As Object, e As EventArgs) Handles TimerSearch.Tick
        TimerSearch.Enabled = False
        If suppressSearch Then Return

        ' 🌟 إلغاء البحث القديم فوراً والبدء بواحد جديد
        CancelSearch()
        _cancellationTokenSource = New CancellationTokenSource()
        Try
            Await PerformSmartSearchAsync(_cancellationTokenSource.Token)
        Catch ex As OperationCanceledException
            ' تم الإلغاء، تجاهل الخطأ
        Catch ex As Exception
            DatabaseModule.LogError("TimerSearch_Tick", ex)
        End Try
    End Sub

    Private Async Function PerformSmartSearchAsync(cancellationToken As CancellationToken) As Task
        Try
            cancellationToken.ThrowIfCancellationRequested()

            If Not HasSearchCriteria() Then
                ' 🌟 إصلاح: كان يعيد تحميل النشطين دائماً حتى لو كنت في عرض الأرشيف
                ' (نفس أسلوب الفحص الموجود في سطر التحكم بالقوائم: نص الزر = "عرض فقط")
                If SaveToolStripMenuItem.Text = "عرض فقط" Then
                    Await LoadArchiveOnlyAsync()
                Else
                    Await LoadPlayersOnlyAsync()
                End If
                Return
            End If

            Dim birthYear As String = If(dtpBirthDate.Checked, dtpBirthDate.Value.ToString("yyyy", Globalization.CultureInfo.InvariantCulture), "")
            Dim dateInYear As String = If(dtpDateIn.Checked, dtpDateIn.Value.ToString("yyyy", Globalization.CultureInfo.InvariantCulture), "")

            Dim service As New PlayerService()
            Dim dt As DataTable = Await service.SearchPlayersAsync(
                txtPlayerID.Text, txtPlayerName.Text, birthYear,
                txtFatherName.Text, txtFatherID.Text, txtWifeName.Text, txtWifeID.Text,
                txtFone.Text, txtAltFone.Text, txtAddress.Text, dateInYear, txtNotes.Text,
                cmbGender.Text, cmbMaritalStatus.Text, cmbTeam.Text, cmbRolle.Text, cmbJobTitle.Text,
                cancellationToken)

            If cancellationToken.IsCancellationRequested Then Return ' خروج آمن

            If InvokeRequired Then
                Invoke(Sub() DisplaySearchResults(dt))
            Else
                DisplaySearchResults(dt)
            End If

        Catch ex As OperationCanceledException
            ' تجاهل
        Catch ex As Exception
            DatabaseModule.LogError("PerformSmartSearchAsync", ex)
            MessageBox.Show("خطأ في البحث: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub DisplaySearchResults(dt As DataTable)
        Try
            DataGridView1.SuspendLayout()
            ' 🌟 تم إزالة DataSource = Nothing
            DataGridView1.DataSource = dt
            searchResultCount = If(dt IsNot Nothing, dt.Rows.Count, 0)
            UpdateSearchResultLabel()
        Catch ex As Exception
            DatabaseModule.LogError("DisplaySearchResults", ex)
        Finally
            DataGridView1.ResumeLayout()
        End Try
    End Sub

    Private Function HasSearchCriteria() As Boolean
        Return Not (String.IsNullOrWhiteSpace(txtPlayerID.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtPlayerName.Text) AndAlso
                   Not dtpBirthDate.Checked AndAlso
                   String.IsNullOrWhiteSpace(txtFatherName.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtFatherID.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtWifeName.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtWifeID.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtFone.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtAltFone.Text) AndAlso
                   String.IsNullOrWhiteSpace(txtAddress.Text) AndAlso
                   Not dtpDateIn.Checked AndAlso
                   String.IsNullOrWhiteSpace(txtNotes.Text) AndAlso
                   String.IsNullOrWhiteSpace(cmbGender.Text) AndAlso
                   String.IsNullOrWhiteSpace(cmbMaritalStatus.Text) AndAlso
                   String.IsNullOrWhiteSpace(cmbTeam.Text) AndAlso
                   String.IsNullOrWhiteSpace(cmbRolle.Text) AndAlso
                   String.IsNullOrWhiteSpace(cmbJobTitle.Text))
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

#Region "Data Operations"

    Private Async Function InsertPlayerAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True
            SetSavingState(True)

            ' 🌟 إذا حقل الهوية فارغ: توليد رقم مؤقت تلقائي (نطاق 900000000+) والحفظ ممكن فوراً
            If String.IsNullOrWhiteSpace(txtPlayerID.Text) Then
                Dim srvGen As New PlayerService()
                Dim tempId As Long = Await srvGen.GenerateTempPlayerIdAsync()
                If tempId = 0 Then
                    MessageBox.Show("تعذر توليد رقم مؤقت — جرّب إدخال رقم الهوية يدوياً.", "تنبيه",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' 🌟 إصلاح: الإسناد كان يطلق بحثاً غير مقصود أثناء ظهور الرسالة — كتم مثل باقي المسارات
                suppressSearch = True
                Try
                    txtPlayerID.Text = tempId.ToString()
                Finally
                    suppressSearch = False
                End Try
                TimerSearch.Enabled = False
                MessageBox.Show($"لم يتم إدخال رقم هوية — تم توليد رقم مؤقت: {tempId}" & vbCrLf &
                                "يمكنك تحديثه برقم الهوية الحقيقي لاحقاً (تعديل بيانات اللاعب).",
                                "رقم مؤقت", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            If Not ValidateAllData() Then Return

            Dim playerIdNum As Long
            If Not Long.TryParse(txtPlayerID.Text, playerIdNum) Then
                MessageBox.Show("رقم الهوية غير صالح", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPlayerID.Focus()
                Return
            End If
            Dim service As New PlayerService()
            If Await service.PlayerExistsAsync(txtPlayerID.Text) Then
                MessageBox.Show("رقم الهوية موجود مسبقاً", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPlayerID.Focus()
                Return
            End If

            Dim photoParam As Object = UtilityModule.ImageToDBValue(PictureBox1.Image)

            ' 🌟             ' 🌟 حفظ التواريخ بصيغة ISO (yyyy-MM-dd) — العرض dd/MM/yyyy في الواجهة فقط
            Dim birthDateVal As Object = If(dtpBirthDate.Checked, CObj(dtpBirthDate.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)), DBNull.Value)
            Dim dateInVal As Object = If(dtpDateIn.Checked, CObj(dtpDateIn.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)), DBNull.Value)
            Dim result As Integer = Await service.InsertPlayerAsync(
                playerIdNum, txtPlayerName.Text, birthDateVal,
                txtFatherName.Text, UtilityModule.ToDBValue(txtFatherID.Text, True), txtWifeName.Text, UtilityModule.ToDBValue(txtWifeID.Text, True),
                txtFone.Text, txtAltFone.Text, txtAddress.Text, dateInVal, txtNotes.Text,
                cmbGender.Text, cmbMaritalStatus.Text, cmbTeam.Text, cmbRolle.Text, cmbJobTitle.Text, photoParam,
                UserSession.CurrentUsername)

            If result > 0 Then
                MessageBox.Show("تم الحفظ بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' 🌟 حفظ أي قيم جديدة كتبها المستخدم بالقوائم المنسدلة
                CaptureCustomDropdownItems()

                Await LoadPlayersOnlyAsync()
                Await RefreshDashboardAsync()
                ClearFields()
                SetAddMode()
                UpdateImageButtonsState()
                txtPlayerID.Focus()
            Else
                MessageBox.Show("فشل الحفظ — لم يتم إدراج البيانات.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("InsertPlayerAsync", ex)
            MessageBox.Show("خطأ في إضافة اللاعب: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function

    Private Async Function UpdatePlayerAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True
            SetSavingState(True)

            If String.IsNullOrWhiteSpace(txtPlayerID.Text) Then
                MessageBox.Show("الرجاء اختيار لاعب للتحديث", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If Not ValidateAllData() Then Return

            Dim newPlayerIDStr As String = txtPlayerID.Text.Trim()
            Dim newPlayerIDNum As Long
            If Not Long.TryParse(newPlayerIDStr, newPlayerIDNum) Then
                MessageBox.Show("رقم الهوية الجديد غير صالح", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim service As New PlayerService()

            If newPlayerIDStr <> currentPlayerID Then
                If Await service.PlayerExistsAsync(newPlayerIDStr) Then
                    MessageBox.Show("رقم الهوية الجديد موجود مسبقاً لاعب آخر!", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtPlayerID.Focus()
                    Return
                End If
            End If

            Dim photoParam As Object = Nothing
            If _imageChanged Then
                photoParam = UtilityModule.ImageToDBValue(PictureBox1.Image)
            End If

            ' 🌟             ' 🌟 حفظ التواريخ بصيغة ISO (yyyy-MM-dd) — العرض dd/MM/yyyy في الواجهة فقط
            Dim birthDateVal As Object = If(dtpBirthDate.Checked, CObj(dtpBirthDate.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)), DBNull.Value)
            Dim dateInVal As Object = If(dtpDateIn.Checked, CObj(dtpDateIn.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)), DBNull.Value)
            Dim oldPlayerIDNum As Long
            If Not Long.TryParse(currentPlayerID, oldPlayerIDNum) Then
                MessageBox.Show("رقم هوية اللاعب الأصلي غير صالح — أعد اختيار اللاعب من الجدول", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim result As Integer = Await service.UpdatePlayerAsync(
                newPlayerIDNum, oldPlayerIDNum, txtPlayerName.Text, birthDateVal,
                txtFatherName.Text, UtilityModule.ToDBValue(txtFatherID.Text, True), txtWifeName.Text, UtilityModule.ToDBValue(txtWifeID.Text, True),
                txtFone.Text, txtAltFone.Text, txtAddress.Text, dateInVal, txtNotes.Text,
                cmbGender.Text, cmbMaritalStatus.Text, cmbTeam.Text, cmbRolle.Text, cmbJobTitle.Text,
                photoParam, _imageChanged,
                UserSession.CurrentUsername)

            If result > 0 Then
                MessageBox.Show("تم تحديث بيانات اللاعب بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' 🌟 حفظ أي قيم جديدة كتبها المستخدم بالقوائم المنسدلة
                CaptureCustomDropdownItems()

                _imageChanged = False
                Await LoadPlayersOnlyAsync()
                Await RefreshDashboardAsync()
                ClearFields()
                SetAddMode()
                UpdateImageButtonsState()
                txtPlayerID.Focus()
            Else
                MessageBox.Show("لم يتم تحديث أي بيانات — تأكد من صحة رقم الهوية.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("UpdatePlayerAsync", ex)
            MessageBox.Show("خطأ في تحديث اللاعب: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            SetSavingState(False)
        End Try
    End Function
    Private Async Function DeletePlayersAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True

            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("لا يوجد لاعبين للحذف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            Dim playersToDelete As New List(Of DataGridViewRow)

            If selectedRows.Count > 0 Then
                playersToDelete = selectedRows
            ElseIf DataGridView1.CurrentRow IsNot Nothing Then
                playersToDelete.Add(DataGridView1.CurrentRow)
            Else
                MessageBox.Show("الرجاء اختيار لاعب أو تحديد لاعبين للحذف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            For Each row In playersToDelete
                If DataGridView1.Columns.Contains(AppConstants.Grid_DataSource) AndAlso
                   row.Cells(AppConstants.Grid_DataSource).Value IsNot Nothing AndAlso
                   row.Cells(AppConstants.Grid_DataSource).Value.ToString() = AppConstants.Source_Archive Then
                    MessageBox.Show("لا يمكن حذف بيانات مؤرشفة. قم بإلغاء تحديد اللاعبين المؤرشفين.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            Next

            Dim firstName As String = SafeGetString(playersToDelete(0).Cells(AppConstants.Col_PlayerName).Value)
            Dim confirmMsg As String = If(playersToDelete.Count > 1,
                $"هل أنت متأكد من حذف {playersToDelete.Count} لاعب/لاعبين؟" & vbCrLf & "سيتم نقل البيانات إلى الأرشيف",
                $"هل أنت متأكد من حذف اللاعب: {firstName}؟" & vbCrLf & "سيتم نقل البيانات إلى الأرشيف")

            If Not ConfirmAction(confirmMsg, "تأكيد الحذف") Then Return

            Cursor = Cursors.WaitCursor
            Dim successCount As Integer = 0
            Dim failedNames As New List(Of String)
            Dim service As New PlayerService()

            For Each row In playersToDelete
                Try
                    Dim pid As String = SafeGetString(row.Cells(AppConstants.Col_PlayerId).Value)
                    ' 🌟 تنفيذ قاعدة البيانات خارج خيط الواجهة حتى لا تتجمد الشاشة
                    If Await Task.Run(Function() service.MoveToArchive(pid)) Then
                        successCount += 1
                    Else
                        failedNames.Add(SafeGetString(row.Cells(AppConstants.Col_PlayerName).Value))
                    End If
                Catch ex As Exception
                    failedNames.Add(SafeGetString(row.Cells(AppConstants.Col_PlayerName).Value))
                    DatabaseModule.LogError("DeletePlayersAsync - row", ex)
                End Try
            Next

            If failedNames.Count = 0 Then
                MessageBox.Show(If(playersToDelete.Count > 1, $"تم حذف {successCount} لاعب بنجاح", "تم حذف اللاعب بنجاح"), "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show($"تم حذف {successCount} لاعب بنجاح" & vbCrLf & $"فشل حذف {failedNames.Count} لاعب:" & vbCrLf & String.Join(vbCrLf, failedNames), "نتيجة الحذف", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            Await LoadPlayersOnlyAsync()
            Await RefreshDashboardAsync()
            ClearFields()
            SetAddMode()
            UpdateImageButtonsState()

        Catch ex As Exception
            DatabaseModule.LogError("DeletePlayersAsync", ex)
            MessageBox.Show("خطأ في حذف اللاعبين: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            Cursor = Cursors.Default
        End Try
    End Function

    Private Async Function RestoreFromArchiveAsync() As Task
        If _isSaving Then Return
        Try
            _isSaving = True

            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("لا يوجد لاعبين في الأرشيف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            Dim playersToRestore As New List(Of DataGridViewRow)

            If selectedRows.Count > 0 Then
                playersToRestore = selectedRows
            ElseIf DataGridView1.CurrentRow IsNot Nothing AndAlso
                   DataGridView1.Columns.Contains(AppConstants.Grid_DataSource) AndAlso
                   DataGridView1.CurrentRow.Cells(AppConstants.Grid_DataSource).Value?.ToString() = AppConstants.Source_Archive Then
                playersToRestore.Add(DataGridView1.CurrentRow)
            Else
                MessageBox.Show("الرجاء اختيار لاعب من الأرشيف للاستعادة", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            For Each row In playersToRestore
                If row.Cells(AppConstants.Grid_DataSource).Value?.ToString() <> AppConstants.Source_Archive Then
                    MessageBox.Show("بعض اللاعبين المحددين ليسوا من الأرشيف", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            Next

            Dim firstName As String = SafeGetString(playersToRestore(0).Cells(AppConstants.Col_PlayerName).Value)
            Dim confirmMsg As String = If(playersToRestore.Count > 1,
                $"هل أنت متأكد من استعادة {playersToRestore.Count} لاعب؟",
                $"هل أنت متأكد من استعادة اللاعب: {firstName}؟")

            If Not ConfirmAction(confirmMsg, "تأكيد الاستعادة") Then Return

            Cursor = Cursors.WaitCursor
            Dim successCount As Integer = 0
            Dim failedNames As New List(Of String)
            Dim service As New PlayerService()

            For Each row In playersToRestore
                Try
                    Dim pid As String = SafeGetString(row.Cells(AppConstants.Col_PlayerId).Value)
                    ' 🌟 تنفيذ قاعدة البيانات خارج خيط الواجهة حتى لا تتجمد الشاشة
                    If Await Task.Run(Function() service.RestoreFromArchive(pid)) Then
                        successCount += 1
                    Else
                        failedNames.Add(SafeGetString(row.Cells(AppConstants.Col_PlayerName).Value))
                    End If
                Catch ex As Exception
                    failedNames.Add(SafeGetString(row.Cells(AppConstants.Col_PlayerName).Value))
                    DatabaseModule.LogError("RestoreFromArchiveAsync - row", ex)
                End Try
            Next

            If failedNames.Count = 0 Then
                MessageBox.Show($"تم استعادة {successCount} لاعب بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show($"تم استعادة {successCount} — فشل استعادة {failedNames.Count}: " & String.Join("، ", failedNames), "نتيجة الاستعادة", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

            ClearFields()
            Await LoadPlayersOnlyAsync()
            Await RefreshDashboardAsync()
            UpdateImageButtonsState()

        Catch ex As Exception
            DatabaseModule.LogError("RestoreFromArchiveAsync", ex)
            MessageBox.Show("خطأ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isSaving = False
            Cursor = Cursors.Default
        End Try
    End Function

#End Region

#Region "Selection Operations"

    Private Function GetSelectedRowsFromCheckBox() As List(Of DataGridViewRow)
        Return GridHelper.GetCheckedRows(DataGridView1)
    End Function
    Private Sub UpdateSelectedCount()
        Dim count As Integer = GridHelper.CountCheckedRows(DataGridView1)
        If lblSelectedCount IsNot Nothing Then
            lblSelectedCount.Text = "تم تحديد: " & count
            lblSelectedCount.ForeColor = If(count > 0, Color.Green, Color.Gray)
        End If
    End Sub
    Private Sub UpdateColumnHeaderState()
        GridHelper.UpdateSelectHeaderState(DataGridView1)
    End Sub
#End Region

#Region "ID Validation & Immediate Leave Events"

    Private Async Sub txtPlayerID_Leave(sender As Object, e As EventArgs) Handles txtPlayerID.Leave
        If isClosing Then Return
        Try
            If String.IsNullOrWhiteSpace(txtPlayerID.Text) Then Return

            ' 🌟 استخدام ValidationModule الموحد
            Dim cleanID As String = New String(txtPlayerID.Text.Where(AddressOf Char.IsDigit).ToArray())

            ' 🌟 الأرقام المؤقتة المولدة (نطاق 900000000+) معفاة من فحص Luhn — نفس استثناء ValidateAllData
            Dim isTempId As Boolean = (cleanID.Length = 9 AndAlso cleanID.StartsWith("900", StringComparison.Ordinal))
            If Not isTempId Then
                Dim errMsg As String = ""
                If Not ValidationModule.IsIdValid(txtPlayerID.Text, errMsg) Then
                    MessageBox.Show(errMsg, "خطأ في رقم الهوية", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtPlayerID.Focus()
                    Return
                End If
            End If

            ' فحص التكرار في قاعدة البيانات
            If Not isEditMode Then
                Dim playerName As String = Await CheckPlayerNameExistsAsync(cleanID)
                If Not String.IsNullOrEmpty(playerName) Then
                    MessageBox.Show($"رقم الهوية {cleanID} موجود مسبقاً للاعب: {playerName}" & vbCrLf & "سيتم نقلك إلى بيانات هذا اللاعب", "رقم مكرر", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Await HighlightPlayerInGrid(cleanID)
                End If
            End If

        Catch ex As Exception
            DatabaseModule.LogError("txtPlayerID_Leave", ex)
        End Try
    End Sub

    Private Sub txtFatherID_Leave(sender As Object, e As EventArgs) Handles txtFatherID.Leave
        If String.IsNullOrWhiteSpace(txtFatherID.Text) Then Return

        ' 🌟 استخدام ValidationModule الموحد
        ' 🌟 L-01: التنسيق إجباري، وفحص Luhn تحذير تأكيد اختياري — هويات قديمة صحيحة قد لا تجتازه
        Dim errMsg As String = ""
        If Not ValidationModule.IsIdFormatValid(txtFatherID.Text, errMsg) Then
            MessageBox.Show(errMsg, "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtFatherID.Focus()
            Return
        End If
        Dim cleanFatherId As String = New String(txtFatherID.Text.Where(AddressOf Char.IsDigit).ToArray())
        If Not ValidationModule.IsLuhnValid(cleanFatherId) Then
            Dim confirmFather As DialogResult = MessageBox.Show(
                "رقم هوية الأب لا يجتاز فحص Luhn — قد يكون هوية قديمة صحيحة." & vbCrLf &
                "هل تريد المتابعة بهذا الرقم؟", "تأكيد رقم الهوية",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            If confirmFather = DialogResult.No Then txtFatherID.Focus()
        End If
    End Sub

    Private Sub txtWifeID_Leave(sender As Object, e As EventArgs) Handles txtWifeID.Leave
        If String.IsNullOrWhiteSpace(txtWifeID.Text) Then Return

        ' 🌟 استخدام ValidationModule الموحد
        ' 🌟 L-01: التنسيق إجباري، وفحص Luhn تحذير تأكيد اختياري — هويات قديمة صحيحة قد لا تجتازه
        Dim errMsg As String = ""
        If Not ValidationModule.IsIdFormatValid(txtWifeID.Text, errMsg) Then
            MessageBox.Show(errMsg, "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtWifeID.Focus()
            Return
        End If
        Dim cleanWifeId As String = New String(txtWifeID.Text.Where(AddressOf Char.IsDigit).ToArray())
        If Not ValidationModule.IsLuhnValid(cleanWifeId) Then
            Dim confirmWife As DialogResult = MessageBox.Show(
                "رقم هوية الزوجة/الأم لا يجتاز فحص Luhn — قد يكون هوية قديمة صحيحة." & vbCrLf &
                "هل تريد المتابعة بهذا الرقم؟", "تأكيد رقم الهوية",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            If confirmWife = DialogResult.No Then txtWifeID.Focus()
        End If
    End Sub

    Private Async Function CheckPlayerNameExistsAsync(playerID As String) As Task(Of String)
        Dim service As New PlayerService()
        Return Await service.GetPlayerNameByIdAsync(playerID)
    End Function

    Private Async Function HighlightPlayerInGrid(playerID As String) As Task
        Try
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                Dim row As DataGridViewRow = DataGridView1.Rows(i)
                If Not row.IsNewRow AndAlso
                   row.Cells(AppConstants.Col_PlayerId).Value IsNot Nothing AndAlso
                   row.Cells(AppConstants.Col_PlayerId).Value.ToString() = playerID Then
                    DataGridView1.ClearSelection()
                    row.Selected = True
                    DataGridView1.CurrentCell = row.Cells(1)
                    Await LoadSelectedPlayerDataAsync(i)
                    Exit For
                End If
            Next
        Catch ex As Exception
            DatabaseModule.LogError("HighlightPlayerInGrid", ex)
        End Try
    End Function

#End Region

#Region "Image Helpers"

    Private Sub UpdateImageButtonsState()
        Dim hasData As Boolean = Not String.IsNullOrWhiteSpace(txtPlayerID.Text) AndAlso Not String.IsNullOrWhiteSpace(txtPlayerName.Text)
        Dim isArchiveView As Boolean = (SaveToolStripMenuItem.Text = "عرض فقط")

        AddimageToolStripMenuItem.Enabled = hasData AndAlso Not isArchiveView AndAlso UserSession.CanEdit
        DeleteimageToolStripMenuItem.Enabled = hasData AndAlso PictureBox1.Image IsNot Nothing AndAlso Not isArchiveView AndAlso UserSession.CanDelete
        SaveimageToolStripMenuItem.Enabled = hasData AndAlso PictureBox1.Image IsNot Nothing
        PrintColToolStripMenuItem.Enabled = True
    End Sub

    Private Function UpdateImageButtonsStateBasedOnData() As Boolean
        Return Not String.IsNullOrWhiteSpace(txtPlayerID.Text) AndAlso Not String.IsNullOrWhiteSpace(txtPlayerName.Text)
    End Function

#End Region

#Region "Printing and Export"

    Private Sub PrintSelectedRows()
        Try
            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            If selectedRows.Count = 0 Then
                MessageBox.Show("الرجاء تحديد لاعب واحد على الأقل للطباعة (وضع علامة ✔ بجانب اسم اللاعب).", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using frm As New frmSelectColumns()
                frm.OperationType = "Print"
                frm.SourceDataGridView = DataGridView1
                frm.StartPosition = FormStartPosition.CenterParent
                If frm.ShowDialog() = DialogResult.OK Then
                    printColumns = frm.SelectedColumns
                    Dim printRows = selectedRows.OrderBy(Function(r) SafeGetString(r.Cells(AppConstants.Col_PlayerName).Value)).ToList()

                    Using helper As New PrintHelper() With {
                        .SourceGrid = DataGridView1,
                        .RowsToPrint = printRows,
                        .ColumnsToPrint = printColumns,
                        .ReportTitle = "تقرير اللاعبين",
                        .FooterSummary = "عدد السجلات المطبوعة: " & printRows.Count
                    }
                        helper.ShowPreview()
                    End Using
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("خطأ في فتح نموذج اختيار الأعمدة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportToExcel()
        Try
            Dim selectedRows As List(Of DataGridViewRow) = GetSelectedRowsFromCheckBox()
            If selectedRows.Count = 0 Then
                MessageBox.Show("الرجاء تحديد لاعب واحد على الأقل للتصدير (وضع علامة ✔ بجانب اسم اللاعب).", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Filter = "CSV Files|*.csv"
                sfd.DefaultExt = "csv"
                sfd.FileName = "تقرير_اللاعبين_" & DateTime.Now.ToString("yyyy-MM-dd")

                If sfd.ShowDialog() = DialogResult.OK Then
                    Cursor = Cursors.WaitCursor
                    Using sw As New StreamWriter(sfd.FileName, False, New UTF8Encoding(True))
                        ' 🌟 توحيد CSV: المنطق المشترك في UtilityModule.WriteGridCsv (كان منسوخاً في 5 نماذج)
                        UtilityModule.WriteGridCsv(DataGridView1, sw,
                            New String() {AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn, AppConstants.Grid_DataSource},
                            selectedRows,
                            New String() {AppConstants.Col_BirthDate, AppConstants.Col_DateIn, AppConstants.Col_DeletedDate})
                    End Using
                    MessageBox.Show($"تم تصدير بيانات {selectedRows.Count} لاعب بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("ExportToExcel", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

#End Region

#Region "UI Helpers"
    ''' <summary>حفظ أي قيم جديدة كتبها المستخدم بالقوائم المدارة إلى قاعدة البيانات (يُستدعى بعد نجاح الحفظ)</summary>
    Private Sub CaptureCustomDropdownItems()
        Try
            Dim srv As New DropdownService()

            CaptureComboAsync(cmbTeam, "Team", srv)
            CaptureComboAsync(cmbRolle, "Rolle", srv)
            CaptureComboAsync(cmbJobTitle, "JobTitle", srv)


        Catch ex As Exception
            DatabaseModule.LogError("CaptureCustomDropdownItems", ex)
        End Try
    End Sub

    ''' <summary>يضيف النص المكتوب حالياً في الكومبوبوكس (إن وجد) إلى قاعدة البيانات ثم يرسل كل عناصر القائمة</summary>
    Private Async Sub CaptureComboAsync(combo As ComboBox, fieldName As String, srv As DropdownService)
        Try
            ' 🌟 1. النص المكتوب يدوياً (مهم!): لا يدخل إلى Items تلقائياً — نرسله صراحة
            Dim typedText As String = combo.Text.Trim()
            If typedText <> "" Then
                Await srv.AddItemAsync(fieldName, typedText)
            End If

            ' 🌟 2. ثم كل عناصر القائمة الحالية (إن وُجدت إضافات برمجية سابقة)
            For Each it As Object In combo.Items
                Await srv.AddItemAsync(fieldName, it.ToString())
            Next

            ' 🌟 3. أضف النص المكتوب أيضاً إلى Items المحلية فوراً حتى يظهر في القائمة دون إعادة تشغيل
            If typedText <> "" Then
                Dim exists As Boolean = False
                For Each it As Object In combo.Items
                    If String.Equals(it.ToString().Trim(), typedText, StringComparison.OrdinalIgnoreCase) Then
                        exists = True
                        Exit For
                    End If
                Next
                If Not exists Then combo.Items.Add(typedText)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("CaptureComboAsync(" & fieldName & ")", ex)
        End Try
    End Sub
    ''' <summary>تحميل عناصر القوائم المخصصة من قاعدة البيانات ودمجها مع عناصر الـ Designer</summary>
    Private Async Function LoadCustomDropdownItemsAsync() As Task
        Try
            Dim srv As New DropdownService()

            Await MergeComboItemsAsync(cmbTeam, "Team", srv)
            Await MergeComboItemsAsync(cmbRolle, "Rolle", srv)
            Await MergeComboItemsAsync(cmbJobTitle, "JobTitle", srv)


        Catch ex As Exception
            DatabaseModule.LogError("LoadCustomDropdownItemsAsync", ex)
        End Try
    End Function

    ''' <summary>دمج عناصر محفوظة بالقاعدة مع عناصر الكومبوبوكس الحالية بدون تكرار</summary>
    Private Async Function MergeComboItemsAsync(combo As ComboBox, fieldName As String, srv As DropdownService) As Task
        Dim items As List(Of String) = Await srv.GetItemsAsync(fieldName)
        For Each v As String In items
            Dim exists As Boolean = False
            For Each it As Object In combo.Items
                If String.Equals(it.ToString().Trim(), v, StringComparison.OrdinalIgnoreCase) Then
                    exists = True
                    Exit For
                End If
            Next
            If Not exists Then combo.Items.Add(v)
        Next
    End Function
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
                    Case AppConstants.Col_PlayerId : col.HeaderText = "رقم الهوية"
                    Case AppConstants.Col_PlayerName : col.HeaderText = "الاسم الكامل"
                    Case AppConstants.Col_BirthDate : col.HeaderText = "تاريخ الميلاد"
                    Case AppConstants.Col_FatherName : col.HeaderText = "اسم الأب"
                    Case AppConstants.Col_FatherId : col.HeaderText = "رقم الأب"
                    Case AppConstants.Col_WifeName : col.HeaderText = "اسم الزوجة/الأم"
                    Case AppConstants.Col_WifeId : col.HeaderText = "رقم الزوجة/الأم"
                    Case AppConstants.Col_Fone : col.HeaderText = "رقم الهاتف"
                    Case AppConstants.Col_AltFone : col.HeaderText = "هاتف بديل"
                    Case AppConstants.Col_Address : col.HeaderText = "العنوان"
                    Case AppConstants.Col_DateIn : col.HeaderText = "تاريخ الانتساب"
                    Case AppConstants.Col_Notes : col.HeaderText = "ملاحظات"
                    Case AppConstants.Col_Gender : col.HeaderText = "الجنس"
                    Case AppConstants.Col_MaritalStatus : col.HeaderText = "الحالة الاجتماعية"
                    Case AppConstants.Col_Team : col.HeaderText = "الفريق"
                    Case AppConstants.Col_Rolle : col.HeaderText = "الدور"
                    Case AppConstants.Col_JobTitle : col.HeaderText = "المسمى الوظيفي"
                    Case AppConstants.Col_DeletedDate : col.HeaderText = "تاريخ الأرشفة"
                    Case AppConstants.Col_LastModifiedBy : col.HeaderText = "آخر تعديل بواسطة"
                End Select

                Select Case col.Name
                    Case AppConstants.Col_PlayerId, AppConstants.Col_FatherId, AppConstants.Col_WifeId, AppConstants.Col_Fone, AppConstants.Col_AltFone,
                         AppConstants.Col_BirthDate, AppConstants.Col_DateIn, AppConstants.Col_DeletedDate, AppConstants.Grid_SeqColumn, AppConstants.Col_LastModifiedBy
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
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
                {AppConstants.Col_PlayerName, 200}, {AppConstants.Col_FatherName, 200}, {AppConstants.Col_WifeName, 200},
                {AppConstants.Col_Address, 150}, {AppConstants.Col_Notes, 150},
                {AppConstants.Col_PlayerId, 80}, {AppConstants.Col_FatherId, 85}, {AppConstants.Col_WifeId, 95},
                {AppConstants.Col_Fone, 90}, {AppConstants.Col_AltFone, 90},
                {AppConstants.Col_BirthDate, 85}, {AppConstants.Col_DateIn, 95}, {AppConstants.Col_DeletedDate, 95},
                {AppConstants.Col_Gender, 60}, {AppConstants.Col_MaritalStatus, 100},
                {AppConstants.Col_Team, 90}, {AppConstants.Col_Rolle, 90}, {AppConstants.Col_JobTitle, 110},
                {AppConstants.Col_LastModifiedBy, 100}
            }
            For Each col As DataGridViewColumn In DataGridView1.Columns
                If widths.ContainsKey(col.Name) Then col.Width = widths(col.Name)
            Next
        Catch ex As Exception
            DatabaseModule.LogError("CustomizeColumnWidths", ex)
        Finally
            DataGridView1.ResumeLayout()
        End Try
    End Sub

    Private Sub UpdateSearchResultLabel()
        If lblSearchResult Is Nothing Then Return
        If searchResultCount > 0 Then
            lblSearchResult.Text = "عدد النتائج: " & searchResultCount
            lblSearchResult.ForeColor = Color.Green
        Else
            lblSearchResult.Text = "لا توجد نتائج"
            lblSearchResult.ForeColor = Color.Red
        End If
        ' 🌟 إصلاح: كان ينشئ Font جديداً في كل تحديث دون التخلص من القديم — تراكم مقابض GDI
        Dim oldFont = lblSearchResult.Font
        lblSearchResult.Font = New Font("Arial", 10, FontStyle.Bold)
        If oldFont IsNot Nothing AndAlso Not oldFont.IsSystemFont Then oldFont.Dispose()
    End Sub

    Private Sub ClearFields()
        suppressSearch = True
        Try
            txtPlayerID.Text = ""
            txtPlayerName.Text = ""
            dtpBirthDate.Checked = False
            txtFatherName.Text = ""
            txtFatherID.Text = ""
            txtWifeName.Text = ""
            txtWifeID.Text = ""
            txtFone.Text = ""
            txtAltFone.Text = ""
            txtAddress.Text = ""
            dtpDateIn.Checked = False
            txtNotes.Text = ""
            cmbGender.SelectedIndex = -1
            cmbGender.Text = ""
            cmbMaritalStatus.SelectedIndex = -1
            cmbMaritalStatus.Text = ""
            cmbTeam.SelectedIndex = -1
            cmbTeam.Text = ""
            cmbJobTitle.SelectedIndex = -1
            cmbJobTitle.Text = ""
            cmbRolle.SelectedIndex = -1
            cmbRolle.Text = ""
            SafeDisposeImage()
            tempImagePath = ""
            _imageChanged = False
        Finally
            suppressSearch = False
        End Try
    End Sub

    Private Sub SafeDisposeImage()
        If PictureBox1.Image IsNot Nothing Then
            Dim oldImage As Image = PictureBox1.Image
            PictureBox1.Image = Nothing
            oldImage.Dispose()
        End If
    End Sub

    Private Sub SetSavingState(isSaving As Boolean)
        ' 🌟 إصلاح ثغرة صلاحيات: عند انتهاء الحفظ تُستعاد الأزرار حسب صلاحيات المستخدم الفعلية
        SaveToolStripMenuItem.Enabled = Not isSaving AndAlso UserSession.CanEdit
        DeleteToolStripMenuItem.Enabled = Not isSaving AndAlso UserSession.CanDelete
        AddimageToolStripMenuItem.Enabled = Not isSaving AndAlso UserSession.CanEdit
    End Sub

    Private Sub SetAddMode()
        isEditMode = False
        currentPlayerID = ""
        _imageChanged = False
        EnableAllControls()
        txtPlayerID.Enabled = True
        txtPlayerID.Focus()
        SaveToolStripMenuItem.Text = "حفظ"
        SaveToolStripMenuItem.Enabled = UserSession.CanEdit ' 🌟 إصلاح ثغرة: كان يُفعّل الحفظ بدون فحص الصلاحية
        UpdateImageButtonsState()
    End Sub

    Private Sub SetEditMode()
        isEditMode = True
        currentPlayerID = txtPlayerID.Text ' 🌟 حفظ رقم الهوية القديم قبل أي تعديل
        _imageChanged = False
        EnableAllControls()
        txtPlayerID.Enabled = True ' 🌟 السماح بتعديل رقم الهوية
        txtPlayerID.Focus()
        txtPlayerID.SelectAll()
        SaveToolStripMenuItem.Enabled = UserSession.CanEdit ' 🌟 إصلاح ثغرة: كان يُفعّل الحفظ بدون فحص الصلاحية
        SaveToolStripMenuItem.Text = "تحديث"
        UpdateImageButtonsState()
    End Sub
    Private Sub SetArchiveViewMode()
        MessageBox.Show("هذه بيانات مؤرشفة — يمكنك العرض فقط", "معلومات", MessageBoxButtons.OK, MessageBoxIcon.Information)
        DisableAllControls()
        RestorefromarchifeToolStripMenuItem.Enabled = UserSession.CanEdit ' 🌟 إصلاح: كان متاحاً لغير المصرح ويفشل بصمت داخل الخدمة
        SaveToolStripMenuItem.Text = "عرض فقط"
        UpdateImageButtonsState()
    End Sub

    Private Sub DisableAllControls()
        DisableInputControlsRecursive(Me)
        SaveToolStripMenuItem.Enabled = False
        DeleteToolStripMenuItem.Enabled = False
        RestorefromarchifeToolStripMenuItem.Enabled = False
        AddimageToolStripMenuItem.Enabled = False
        DeleteimageToolStripMenuItem.Enabled = False
        SaveimageToolStripMenuItem.Enabled = False
        ExportToExcelToolStripMenuItem.Enabled = False
        PrintColToolStripMenuItem.Enabled = True
        UpdateImageButtonsState()
    End Sub

    Private Sub EnableAllControls()
        EnableInputControlsRecursive(Me)
        SaveToolStripMenuItem.Enabled = UserSession.CanEdit
        AddimageToolStripMenuItem.Enabled = UserSession.CanEdit
        DeleteToolStripMenuItem.Enabled = UserSession.CanDelete
        DeleteimageToolStripMenuItem.Enabled = UserSession.CanDelete

        ' 🌟 يجب أن يكون مغلقًا هنا، ويُفتح فقط في SetArchiveViewMode
        RestorefromarchifeToolStripMenuItem.Enabled = False

        SaveimageToolStripMenuItem.Enabled = True
        PrintColToolStripMenuItem.Enabled = True
        ExportToExcelToolStripMenuItem.Enabled = True
        UpdateImageButtonsState()
    End Sub

    Private Sub DisableInputControlsRecursive(parent As Control)
        For Each ctrl As Control In parent.Controls
            If TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox OrElse TypeOf ctrl Is DateTimePicker Then
                ctrl.Enabled = False
            End If
            If ctrl.HasChildren Then
                DisableInputControlsRecursive(ctrl)
            End If
        Next
    End Sub

    Private Sub EnableInputControlsRecursive(parent As Control)
        For Each ctrl As Control In parent.Controls
            If TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox OrElse TypeOf ctrl Is DateTimePicker Then
                ctrl.Enabled = True
            End If
            If ctrl.HasChildren Then
                EnableInputControlsRecursive(ctrl)
            End If
        Next
    End Sub

    Private Sub ColorizeRows()
        Try
            If DataGridView1.Rows.Count = 0 Then Return
            If Not DataGridView1.Columns.Contains(AppConstants.Grid_DataSource) Then Return

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.IsNewRow Then Continue For

                Dim sourceValue = row.Cells(AppConstants.Grid_DataSource).Value
                If sourceValue IsNot Nothing AndAlso sourceValue.ToString() = AppConstants.Source_Archive Then
                    row.DefaultCellStyle.BackColor = Color.LightCoral
                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(205, 92, 92)
                Else
                    row.DefaultCellStyle.BackColor = Color.White
                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(173, 216, 230)
                End If
            Next
        Catch ex As Exception
            DatabaseModule.LogError("ColorizeRows", ex)
        End Try
    End Sub
#End Region

#Region "Input Validation"

    Private Sub SetupInputRestrictions()
        txtPlayerID.MaxLength = 9
        txtFatherID.MaxLength = 9
        txtWifeID.MaxLength = 9
        txtFone.MaxLength = 10
        txtAltFone.MaxLength = 10

        dtpBirthDate.ShowCheckBox = True
        dtpBirthDate.Checked = False
        dtpBirthDate.Format = DateTimePickerFormat.Short

        dtpDateIn.ShowCheckBox = True
        dtpDateIn.Checked = False
        dtpDateIn.Format = DateTimePickerFormat.Short
    End Sub

    ' 🌟 تم تنظيف هذه الدالة بالكامل واستخدام ValidationModule
    Private Function ValidateAllData() As Boolean
        Dim errMsg As String = ""

        ' 1. التحقق من رقم هوية اللاعب (أرقام النطاق المؤقت 900xxx معفاة من فحص Luhn)
        If Not String.IsNullOrWhiteSpace(txtPlayerID.Text) AndAlso
           txtPlayerID.Text.StartsWith("900", StringComparison.Ordinal) AndAlso
           txtPlayerID.Text.Length = 9 Then
            ' رقم مؤقت مولّد — مقبول كما هو
        ElseIf Not ValidationModule.IsIdValid(txtPlayerID.Text, errMsg) Then
            MessageBox.Show(errMsg, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPlayerID.Focus() : Return False
        End If

        ' 2. التحقق من رقم هوية الأب (إذا تم إدخاله)
        ' 🌟 L-01: التنسيق إجباري، وفحص Luhn تحذير تأكيد اختياري للحقول الثانوية
        If Not String.IsNullOrWhiteSpace(txtFatherID.Text) Then
            If Not ValidationModule.IsIdFormatValid(txtFatherID.Text, errMsg) Then
                MessageBox.Show(errMsg, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFatherID.Focus() : Return False
            End If
            Dim cleanFatherId As String = New String(txtFatherID.Text.Where(AddressOf Char.IsDigit).ToArray())
            If Not ValidationModule.IsLuhnValid(cleanFatherId) Then
                Dim confirmFather As DialogResult = MessageBox.Show(
                    "رقم هوية الأب لا يجتاز فحص Luhn — قد يكون هوية قديمة صحيحة." & vbCrLf &
                    "هل تريد المتابعة بحفظ هذا الرقم؟", "تأكيد رقم الهوية",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
                If confirmFather = DialogResult.No Then txtFatherID.Focus() : Return False
            End If
        End If

        ' 3. التحقق من رقم هوية الزوجة/الأم (إذا تم إدخاله)
        ' 🌟 L-01: التنسيق إجباري، وفحص Luhn تحذير تأكيد اختياري للحقول الثانوية
        If Not String.IsNullOrWhiteSpace(txtWifeID.Text) Then
            If Not ValidationModule.IsIdFormatValid(txtWifeID.Text, errMsg) Then
                MessageBox.Show(errMsg, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtWifeID.Focus() : Return False
            End If
            Dim cleanWifeId As String = New String(txtWifeID.Text.Where(AddressOf Char.IsDigit).ToArray())
            If Not ValidationModule.IsLuhnValid(cleanWifeId) Then
                Dim confirmWife As DialogResult = MessageBox.Show(
                    "رقم هوية الزوجة/الأم لا يجتاز فحص Luhn — قد يكون هوية قديمة صحيحة." & vbCrLf &
                    "هل تريد المتابعة بحفظ هذا الرقم؟", "تأكيد رقم الهوية",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
                If confirmWife = DialogResult.No Then txtWifeID.Focus() : Return False
            End If
        End If

        ' 4. التحقق من أرقام الهواتف
        If Not String.IsNullOrWhiteSpace(txtFone.Text) Then
            If Not ValidationModule.IsPhoneValid(txtFone.Text, errMsg) Then
                MessageBox.Show(errMsg, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtFone.Focus() : Return False
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtAltFone.Text) Then
            If Not ValidationModule.IsPhoneValid(txtAltFone.Text, errMsg) Then
                MessageBox.Show(errMsg, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtAltFone.Focus() : Return False
            End If
        End If

        ' 5. التحقق من الاسم وتاريخ الميلاد (موحد هنا بدلاً من ValidateBeforeSave)
        Dim birthDateVal As Date? = If(dtpBirthDate.Checked, dtpBirthDate.Value, Nothing)
        If Not ValidationModule.IsPlayerBasicDataValid(txtPlayerName.Text, birthDateVal, errMsg) Then
            MessageBox.Show(errMsg, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            If errMsg.Contains("اسم") Then txtPlayerName.Focus() Else dtpBirthDate.Focus()
            Return False
        End If

        Return True
    End Function

#End Region

#Region "Helper Functions"

    Private Function SafeGetString(value As Object) As String
        Return UtilityModule.SafeString(value)
    End Function

    Private Function ConfirmAction(message As String, title As String) As Boolean
        Return MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

#End Region

#Region "Menu Click Handlers"

    Private Async Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        ' 🌟 إصلاح ثغرة: فحص صلاحية التعديل قبل فتح وضع الإضافة
        If Not UserSession.CanEdit Then
            MessageBox.Show("ليس لديك صلاحية لإضافة لاعبين", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Await LoadPlayersOnlyAsync()
            ClearFields()
            EnableAllControls()
            SetAddMode()
            txtPlayerID.Focus()
        Catch ex As Exception
            DatabaseModule.LogError("NewToolStripMenuItem_Click", ex)
        End Try
    End Sub

    Private Async Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Try
            If isEditMode Then
                Await UpdatePlayerAsync()
            Else
                Await InsertPlayerAsync()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("SaveToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ أثناء الحفظ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Try
            Await DeletePlayersAsync()
        Catch ex As Exception
            DatabaseModule.LogError("DeleteToolStripMenuItem_Click", ex)
        End Try
    End Sub

    Private Async Sub RestorefromarchifeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestorefromarchifeToolStripMenuItem.Click
        Try
            Await RestoreFromArchiveAsync()
        Catch ex As Exception
            DatabaseModule.LogError("RestorefromarchifeToolStripMenuItem_Click", ex)
        End Try
    End Sub

    Private Async Sub ShowplayersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowplayersToolStripMenuItem.Click
        Try
            Await LoadPlayersOnlyAsync()
            ClearFields()
            DisableAllControls()
            SaveToolStripMenuItem.Text = "حفظ"
            UpdateImageButtonsState()
        Catch ex As Exception
            DatabaseModule.LogError("ShowplayersToolStripMenuItem_Click", ex)
        End Try
    End Sub

    Private Async Sub ShowdeleteplayersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowdeleteplayersToolStripMenuItem.Click
        Try
            Await LoadArchiveOnlyAsync()
            ClearFields()
            DisableAllControls()
            RestorefromarchifeToolStripMenuItem.Enabled = UserSession.CanEdit ' 🌟 إصلاح: كان متاحاً لغير المصرح ويفشل بصمت داخل الخدمة
            SaveToolStripMenuItem.Text = "عرض فقط"
            UpdateImageButtonsState()
        Catch ex As Exception
            DatabaseModule.LogError("SowdeleteplayersToolStripMenuItem_Click", ex)
        End Try
    End Sub
    Private Sub PrintimageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintimageToolStripMenuItem.Click
        If PictureBox1.Image Is Nothing Then
            MessageBox.Show("لا توجد صورة لطباعتها. الرجاء تحديد لاعب له صورة.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim prnPreview As New PrintPreviewDialog() With {
                .Document = playerPhotoPrintDoc,
                .WindowState = FormWindowState.Maximized,
                .Text = "معاينة صورة اللاعب"
            }
            prnPreview.ShowDialog()
        Catch ex As Exception
            DatabaseModule.LogError("PrintPlayerPhoto", ex)
            MessageBox.Show("خطأ في طباعة الصورة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub playerPhotoPrintDoc_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles playerPhotoPrintDoc.PrintPage
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

    Private Sub AddimageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddimageToolStripMenuItem.Click
        Try
            If Not UpdateImageButtonsStateBasedOnData() Then
                MessageBox.Show("الرجاء إدخال بيانات اللاعب أولاً", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Using ofd As New OpenFileDialog()
                ofd.Title = "اختر صورة اللاعب"
                ofd.Filter = "ملفات الصور|*.jpg;*.jpeg;*.png;*.bmp;*.gif|كل الملفات|*.*"
                If ofd.ShowDialog() = DialogResult.OK Then
                    Using fs As New FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using img As Image = Image.FromStream(fs)
                            SafeDisposeImage()
                            PictureBox1.Image = New Bitmap(img)
                        End Using
                    End Using
                    tempImagePath = ofd.FileName
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                    _imageChanged = True
                    UpdateImageButtonsState()
                    If isEditMode Then
                        MessageBox.Show("تم تغيير الصورة، الرجاء الضغط على تحديث لحفظ التغييرات", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
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
            If ConfirmAction("هل أنت متأكد من حذف الصورة؟", "تأكيد") Then
                SafeDisposeImage()
                tempImagePath = ""
                _imageChanged = True
                UpdateImageButtonsState()
                If isEditMode Then
                    MessageBox.Show("تم حذف الصورة" & vbCrLf & "الرجاء الضغط على تحديث لحفظ التغييرات", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            DatabaseModule.LogError("DeleteimageToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ في حذف الصورة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SaveimageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveimageToolStripMenuItem.Click
        Try
            If PictureBox1.Image Is Nothing Then
                MessageBox.Show("لا توجد صورة لحفظها", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Using sfd As New SaveFileDialog()
                sfd.Title = "حفظ الصورة"
                sfd.Filter = "JPEG|*.jpg|PNG|*.png|BMP|*.bmp"
                sfd.FileName = "player_" & txtPlayerID.Text & ".jpg"
                If sfd.ShowDialog() = DialogResult.OK Then
                    Dim format As Imaging.ImageFormat = Imaging.ImageFormat.Jpeg
                    Select Case Path.GetExtension(sfd.FileName).ToLower()
                        Case ".png" : format = Imaging.ImageFormat.Png
                        Case ".bmp" : format = Imaging.ImageFormat.Bmp
                    End Select
                    Using bmp As New Bitmap(PictureBox1.Image)
                        bmp.Save(sfd.FileName, format)
                    End Using
                    MessageBox.Show("تم حفظ الصورة بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("SaveimageToolStripMenuItem_Click", ex)
            MessageBox.Show("خطأ في حفظ الصورة: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PrintColToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintColToolStripMenuItem.Click
        PrintSelectedRows()
    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        ExportToExcel()
    End Sub

    Private Sub CopyDBToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyDBToolStripMenuItem.Click
        ' 🌟 إصلاح صلاحيات: بندا النسخ/الاستعادة كانا متاحين لكل المستخدمين بلا فحص
        If Not UserSession.CanDelete Then
            MessageBox.Show("النسخ الاحتياطي متاح للمستخدم المخوّل فقط", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        BackupDatabaseUI()
    End Sub

    Private Sub RestoreCopyDBToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreCopyDBToolStripMenuItem.Click
        ' 🌟 إصلاح صلاحيات: الاستعادة تدمّر القاعدة بالكامل — كانت بلا أي بوابة صلاحيات
        If Not UserSession.CanDelete Then
            MessageBox.Show("الاستعادة متاحة للمستخدم المخوّل فقط", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        RestoreDatabaseUI()
    End Sub

    Private Async Sub PaymentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PaymentToolStripMenuItem.Click
        Using frm As New frmPayments()
            frm.ShowDialog()
        End Using
        Await RefreshDashboardAsync()
    End Sub

    Private Async Sub ExpenseToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles ExpenseToolStripMenuItem1.Click
        Using frm As New frmExpenses()
            frm.ShowDialog()
        End Using
        Await RefreshDashboardAsync()
    End Sub

#End Region

#Region "Text Changed Handlers"

    Private Sub AllTextChangedHandlers(sender As Object, e As EventArgs) Handles _
        txtPlayerID.TextChanged, txtPlayerName.TextChanged,
        txtFatherName.TextChanged, txtFatherID.TextChanged, txtWifeName.TextChanged,
        txtWifeID.TextChanged, txtFone.TextChanged, txtAltFone.TextChanged,
        txtAddress.TextChanged, txtNotes.TextChanged
        If TimerSearch IsNot Nothing AndAlso Not suppressSearch Then
            TimerSearch.Stop()
            TimerSearch.Start()
        End If
        UpdateImageButtonsState()
    End Sub

    Private Sub DateTimePicker_ValueChanged(sender As Object, e As EventArgs) Handles dtpBirthDate.ValueChanged, dtpDateIn.ValueChanged
        If TimerSearch IsNot Nothing AndAlso Not suppressSearch Then
            TimerSearch.Stop()
            TimerSearch.Start()
        End If
        UpdateImageButtonsState()
    End Sub

    Private Sub AllComboBoxChangedHandlers(sender As Object, e As EventArgs) Handles _
        cmbGender.SelectedIndexChanged, cmbMaritalStatus.SelectedIndexChanged,
        cmbTeam.SelectedIndexChanged, cmbJobTitle.SelectedIndexChanged,
        cmbRolle.SelectedIndexChanged
        If TimerSearch IsNot Nothing AndAlso Not suppressSearch Then
            TimerSearch.Stop()
            TimerSearch.Start()
        End If
        UpdateImageButtonsState()
    End Sub

    Private Sub AllComboBoxTextChanged(sender As Object, e As EventArgs) Handles _
        cmbGender.TextChanged, cmbMaritalStatus.TextChanged, cmbTeam.TextChanged,
        cmbJobTitle.TextChanged, cmbRolle.TextChanged
        If TimerSearch IsNot Nothing AndAlso Not suppressSearch Then
            TimerSearch.Stop()
            TimerSearch.Start()
        End If
        UpdateImageButtonsState()
    End Sub

#End Region

#Region "Key Press Handlers"

    Private Sub NumericKeyPress(sender As Object, e As KeyPressEventArgs) Handles _
        txtPlayerID.KeyPress, txtFatherID.KeyPress, txtWifeID.KeyPress,
        txtFone.KeyPress, txtAltFone.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPlayerID_TextChanged(sender As Object, e As EventArgs) Handles txtPlayerID.TextChanged
        If TypeOf sender Is TextBox Then
            Dim txt As TextBox = DirectCast(sender, TextBox)
            Dim cleanText As String = System.Text.RegularExpressions.Regex.Replace(txt.Text, "[^0-9]", "")
            If txt.Text <> cleanText Then
                Dim timerWasEnabled As Boolean = TimerSearch IsNot Nothing AndAlso TimerSearch.Enabled
                If timerWasEnabled Then TimerSearch.Stop()
                suppressSearch = True
                txt.Text = cleanText
                txt.SelectionStart = txt.Text.Length
                suppressSearch = False
                If timerWasEnabled Then TimerSearch.Start()
            End If
        End If
    End Sub

#End Region

#Region "Backup and Restore"

    Private Sub BackupDatabaseUI()
        Try
            If Not DatabaseModule.IsInitialized Then
                If Not DatabaseModule.Initialize(False) Then
                    MessageBox.Show("قاعدة البيانات غير مهيأة", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If
            If String.IsNullOrEmpty(DatabaseModule.CurrentDatabasePath) OrElse Not File.Exists(DatabaseModule.CurrentDatabasePath) Then
                MessageBox.Show("لم يتم العثور على ملف قاعدة البيانات", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            Using sfd As New SaveFileDialog()
                sfd.Title = "حفظ نسخة احتياطية"
                sfd.Filter = "ملفات قاعدة البيانات|*.db|جميع الملفات|*.*"
                sfd.DefaultExt = "db"
                sfd.FileName = $"Backup_PlayersDB_{DateTime.Now:yyyy-MM-dd_HH-mm}.db"
                If sfd.ShowDialog() = DialogResult.OK Then
                    Cursor = Cursors.WaitCursor
                    If DatabaseModule.BackupDatabase(sfd.FileName) Then
                        ' 🌟 [أولوية 4] توثيق النسخ الاحتياطي في سجل التدقيق
                        Dim audit As New AuditService()
                        audit.Log(AuditService.Act_Backup, "قاعدة بيانات", "", System.IO.Path.GetFileName(sfd.FileName))

                        MessageBox.Show($"تم إنشاء النسخة الاحتياطية بنجاح!" & vbCrLf & $"الموقع: {sfd.FileName}", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("فشل إنشاء النسخة الاحتياطية", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("BackupDatabaseUI", ex)
            MessageBox.Show($"خطأ في النسخ الاحتياطي: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ''' <summary>🌟 M-06: هل الملف قاعدة SQLite سليمة؟ — فحص التوقيع القياسي (أول 16 بايت)
    ''' "SQLite format 3" + بايت صفري — نفس الفحص الذي يطبقه مسار الاستعادة المشفرة</summary>
    Private Function IsValidSqliteFile(filePath As String) As Boolean
        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                If fs.Length < 100 Then Return False   ' ترويسة SQLite أصلاً 100 بايت
                Dim header(15) As Byte
                fs.Read(header, 0, 16)
                Dim expected As Byte() = {
                    &H53, &H51, &H4C, &H69, &H74, &H65, &H20, &H66,
                    &H6F, &H72, &H6D, &H61, &H74, &H20, &H33, &H0}   ' "SQLite format 3\0"
                For i As Integer = 0 To 15
                    If header(i) <> expected(i) Then Return False
                Next
                Return True
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("IsValidSqliteFile", ex)
            Return False
        End Try
    End Function

    Private Sub RestoreDatabaseUI()
        Try
            Dim userInput As String = InputBox(
                "تحذير: استعادة النسخة الاحتياطية ستكتب فوق قاعدة البيانات الحالية!" & vbCrLf &
                "جميع البيانات غير المحفوظة ستفقد." & vbCrLf & vbCrLf &
                "للتأكيد، اكتب الكلمة التالية: استعادة" & vbCrLf & vbCrLf &
                "أو اضغط إلغاء للتراجع.",
                "تأكيد استعادة النسخة الاحتياطية",
                "")

            If userInput.Trim() <> "استعادة" Then
                MessageBox.Show("تم إلغاء العملية. لم يتم تغيير أي بيانات.", "إلغاء", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using ofd As New OpenFileDialog()
                ofd.Title = "اختر ملف النسخة الاحتياطية"
                ofd.Filter = "ملفات قاعدة البيانات|*.db;*.adbak|نسخة تلقائية محمية|*.adbak|جميع الملفات|*.*"
                If ofd.ShowDialog() = DialogResult.OK Then
                    If Not File.Exists(ofd.FileName) Then
                        MessageBox.Show("الملف المختار غير موجود", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                    ' 🔴 H-02: النسخ التلقائية صارت مشفّرة (.adbak) — نعطّل الحماية إلى ملف مؤقت نصي أولاً
                    Dim restoreSource As String = ofd.FileName
                    Dim plainTemp As String = ""
                    If DatabaseModule.IsProtectedAutoBackup(ofd.FileName) Then
                        plainTemp = Path.Combine(Path.GetTempPath(), $"AutoBakRest_{Guid.NewGuid():N}.db")
                        If Not DatabaseModule.UnprotectAutoBackup(ofd.FileName, plainTemp) Then
                            MessageBox.Show("فشل فك حماية النسخة التلقائية." & vbCrLf &
                                            "الملف تالف أو منقول من جهاز آخر — مفتاح الحماية مربوط بجهاز الإنشاء.",
                                            "فك الحماية فشل", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Return
                        End If
                        restoreSource = plainTemp
                    End If

                    ' 🌟 M-06: فحص أن الملف قاعدة SQLite سليمة قبل أي عملية —
                    ' سابقاً أي ملف كان بيتقبّل وبيتكبى فوق القاعدة
                    If Not IsValidSqliteFile(restoreSource) Then
                        If plainTemp <> "" Then
                            Try
                                File.Delete(plainTemp)
                            Catch
                            End Try
                        End If
                        MessageBox.Show("الملف المختار ليس قاعدة بيانات SQLite سليمة." & vbCrLf &
                                        "تأكد أنك اخترت ملف نسخة احتياطية صحيح (.db أو .adbak).", "ملف غير صالح",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                    Dim targetFile As String = DatabaseModule.CurrentDatabasePath
                    If String.IsNullOrEmpty(targetFile) Then
                        MessageBox.Show("لم يتم العثور على مسار قاعدة البيانات الحالية", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                    Cursor = Cursors.WaitCursor

                    ' 🌟 1. إجبار قاعدة البيانات على تفريغ الـ WAL في الملف الرئيسي
                    DatabaseModule.VacuumDatabase() ' أو ExecuteNonQuery("PRAGMA wal_checkpoint(TRUNCATE);")

                    ' 🌟 2. إغلاق جميع الاتصالات المفتوحة في الـ Pool لتحرير الملف من الويندوز
                    SQLiteConnection.ClearAllPools()

                    Dim autoBackup As String = Path.Combine(Path.GetDirectoryName(targetFile), $"AutoBackup_BeforeRestore_{DateTime.Now:yyyy-MM-dd_HH-mm}.db")
                    File.Copy(targetFile, autoBackup, True)
                    File.Copy(restoreSource, targetFile, True)

                    ' 🔴 H-02: النسخة المحمية المفكوكة مؤقتاً لم تعد لازمة بعد الاستبدال
                    If plainTemp <> "" Then
                        Try
                            File.Delete(plainTemp)
                        Catch
                        End Try
                    End If

                    ' 🌟 M-06: حذف ملفات WAL/SHM/journal للقاعدة القديمة —
                    ' بقاءها بعد استبدال الملف قد يفسد القاعدة الجديدة (نفس إصلاح frmSecureBackup)
                    For Each side As String In New String() {"-wal", "-shm", "-journal"}
                        Dim sidePath As String = targetFile & side
                        If File.Exists(sidePath) Then File.Delete(sidePath)
                    Next

                    ' 🌟 [أولوية 4] توثيق الاستعادة في سجل التدقيق
                    Dim auditRestore As New AuditService()
                    auditRestore.Log(AuditService.Act_Restore, "قاعدة بيانات", "", System.IO.Path.GetFileName(ofd.FileName))

                    Cursor = Cursors.Default
                    MessageBox.Show("تم استعادة قاعدة البيانات بنجاح!" & vbCrLf & $"نسخة احتياطية تلقائية: {autoBackup}" & vbCrLf & vbCrLf & "سيتم إعادة تشغيل البرنامج الآن.", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Application.Restart()
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("RestoreDatabaseUI", ex)
            MessageBox.Show($"خطأ في الاستعادة: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ChangepassToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangepassToolStripMenuItem.Click
        ' 🌟 استخدام Using لضمان التخلص من النموذج بعد الإغلاق
        Using frm As New frmChangePassword()
            frm.CurrentUsername = UserSession.CurrentUsername
            frm.ShowDialog()
        End Using
    End Sub
    Private Sub UserManegToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserManegToolStripMenuItem.Click
        If UserSession.CurrentUsername.ToLower() <> "admin" Then
            MessageBox.Show("هذه الشاشة متاحة لمدير النظام (admin) فقط", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 🌟 استخدام Using لضمان التخلص من النموذج بعد الإغلاق
        Using frm As New frmManageUsers()
            frm.ShowDialog()
        End Using
    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Dim dialogResult As DialogResult = MessageBox.Show(
            "ماذا تريد أن تفعل؟" & vbCrLf & vbCrLf &
            "نعم = تسجيل الخروج والعودة لشاشة الدخول" & vbCrLf &
            "لا = إغلاق البرنامج بالكامل" & vbCrLf &
            "إلغاء = العودة للبرنامج",
            "خروج من النظام",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button3)

        If dialogResult = DialogResult.Yes Then
            UserSession.IsLoggingOut = True
            Me.Close()
        ElseIf dialogResult = DialogResult.No Then
            UserSession.IsLoggingOut = False
            Me.Close()
        End If
    End Sub

#End Region

#Region "Priority 4 — Security & Reports Menu (🌟 أولوية 4)"

    ''' <summary>
    ''' إضافة بنود قائمة المزايا الأمنية الجديدة ديناميكياً داخل قائمة
    ''' نسخ قاعدة البيانات الموجودة (لا يعتمد على تعديل ملف Designer غير المتوفر).
    ''' تُوضع بعد بند الاستعادة الموجود في نفس القائمة المنسدلة.
    ''' </summary>
    Private Sub AddSecurityMenuItems()
        Try
            ' نقطة الارتكاز: القائمة المنسدلة التي يحويها بند "نسخ قاعدة البيانات" الموجود أصلاً
            If CopyDBToolStripMenuItem Is Nothing OrElse CopyDBToolStripMenuItem.Owner Is Nothing Then
                DatabaseModule.LogInfo("AddSecurityMenuItems", "لم تُعثر القائمة المنسدلة — بنود الأولوية 4 لن تظهر")
                Return
            End If
            Dim owner As ToolStripDropDown = CType(CopyDBToolStripMenuItem.Owner, ToolStripDropDown)

            ' حماية من التكرار لو استُدعيت أكثر من مرة
            For Each existing As ToolStripItem In owner.Items
                If existing.Name = "MiSecureBackup" Then Return
            Next

            owner.Items.Add(New ToolStripSeparator())

            Dim miSecure As New ToolStripMenuItem("نسخة احتياطية مشفّرة (AES)...")
            miSecure.Name = "MiSecureBackup"
            AddHandler miSecure.Click, AddressOf MiSecureBackup_Click
            owner.Items.Add(miSecure)

            Dim miAudit As New ToolStripMenuItem("سجل التدقيق (من فعل ماذا ومتى)...")
            miAudit.Name = "MiAuditLog"
            AddHandler miAudit.Click, AddressOf MiAuditLog_Click
            owner.Items.Add(miAudit)

            Dim miDelinq As New ToolStripMenuItem("تقرير المتأخرات وغرامات التأخير...")
            miDelinq.Name = "MiDelinquency"
            AddHandler miDelinq.Click, AddressOf MiDelinquency_Click
            owner.Items.Add(miDelinq)

            Dim miLock As New ToolStripMenuItem("قفل الشاشة الآن")
            miLock.Name = "MiLockNow"
            AddHandler miLock.Click, AddressOf MiLockNow_Click
            owner.Items.Add(miLock)

            Dim miIdle As New ToolStripMenuItem("إعداد قفل الخمول...")
            miIdle.Name = "MiIdleSettings"
            AddHandler miIdle.Click, AddressOf MiIdleSettings_Click
            owner.Items.Add(miIdle)

            Dim miHtml As New ToolStripMenuItem("تصدير قائمة اللاعبين HTML (طباعة/PDF)")
            miHtml.Name = "MiExportHtml"
            AddHandler miHtml.Click, AddressOf MiExportHtml_Click
            owner.Items.Add(miHtml)
        Catch ex As Exception
            DatabaseModule.LogError("frmPlayers.AddSecurityMenuItems", ex)
        End Try
    End Sub

    Private Sub MiSecureBackup_Click(sender As Object, e As EventArgs)
        Using frm As New frmSecureBackup()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub MiAuditLog_Click(sender As Object, e As EventArgs)
        Using frm As New frmAuditLog()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub MiDelinquency_Click(sender As Object, e As EventArgs)
        Using frm As New frmDelinquency()
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub MiLockNow_Click(sender As Object, e As EventArgs)
        IdleLockManager.Instance.ShowLockScreen()
    End Sub

    ''' <summary>ضبط مدة قفل الخمول (0 = تعطيل) — تُحفظ في جدول الإعدادات</summary>
    Private Sub MiIdleSettings_Click(sender As Object, e As EventArgs)
        Try
            Dim current As Integer = AppSettingsStore.GetIntSetting(AppSettingsStore.Key_IdleLockMinutes, 15, 0, 720)
            Dim raw As String = InputBox(
                "عدد دقائق الخمول قبل قفل الجلسة تلقائياً:" & vbCrLf & vbCrLf &
                "• أدخل رقماً بين 1 و 720 لتفعيل القفل" & vbCrLf &
                "• أدخل 0 لتعطيل القفل" & vbCrLf & vbCrLf &
                "(إذا كان القفل معطلاً منذ إقلاع البرنامج، سيعمل بعد إعادة التشغيل)",
                "إعداد قفل الخمول", current.ToString())

            If raw Is Nothing OrElse raw.Trim() = "" Then Return

            Dim minutes As Integer
            If Not Integer.TryParse(raw.Trim(), minutes) OrElse minutes < 0 OrElse minutes > 720 Then
                MessageBox.Show("أدخل رقماً بين 0 و 720", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            AppSettingsStore.SetIntSetting(AppSettingsStore.Key_IdleLockMinutes, minutes)
            IdleLockManager.Instance.RefreshConfiguration()   ' 🌟 M-03: القيمة الجديدة تسري فوراً بدون انتظار إعادة تشغيل

            If minutes > 0 Then
                IdleLockManager.Instance.Startup()   ' فعّل فوراً إن كان معطلاً
                MessageBox.Show($"تم التفعيل: ستُقفل الجلسة بعد {minutes} دقيقة خمول." & vbCrLf &
                                "يمكنك تجربة «قفل الشاشة الآن» للتأكد.",
                                "تم الحفظ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("تم تعطيل قفل الخمول." & vbCrLf &
                                "(يتوقف الفحص فوراً — وإزالة المراقبة كاملة تتم عند إعادة التشغيل)",
                                "تم الحفظ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmPlayers.MiIdleSettings_Click", ex)
        End Try
    End Sub

    ''' <summary>تصدير قائمة اللاعبين المعروضة إلى تقرير HTML جاهز للطباعة/الـ PDF</summary>
    Private Sub MiExportHtml_Click(sender As Object, e As EventArgs)
        Try
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("لا توجد بيانات معروضة للتصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' استثناء أعمدة التحكم والصورة (البايتات لا تُعرض كنص)
            Dim path As String = ReportExporter.ExportGridToHtml(
                DataGridView1, "قائمة اللاعبين", "اللاعبون النشطون كما تظهر في الشاشة الرئيسية",
                landscape:=True,
                excludedColumnNames:=New String() {AppConstants.Grid_SelectColumn, AppConstants.Grid_SeqColumn,
                                                   AppConstants.Grid_DataSource, AppConstants.Col_PlayerPhoto})

            Dim audit As New AuditService()
            audit.Log(AuditService.Act_ReportExport, "تقرير", "", "قائمة اللاعبين HTML")

            MessageBox.Show("فُتح التقرير في المتصفح — اضغط زر الطباعة لحفظه PDF." & vbCrLf & path,
                            "تم", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            DatabaseModule.LogError("frmPlayers.MiExportHtml_Click", ex)
            MessageBox.Show("خطأ في التصدير: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

End Class
