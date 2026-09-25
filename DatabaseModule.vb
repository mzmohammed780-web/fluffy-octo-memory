Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.IO
Imports System.Text
Imports System.Security.AccessControl
Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Linq
Imports System.Security.Cryptography

Public Module DatabaseModule

#Region "Variables and Properties"

    Private _dbPath As String = ""
    Private _connectionString As String = ""
    Private _isInitialized As Boolean = False
    Private _lastError As String = ""

    ' كائن قفل مخصص لمنع تعارض خيوط المعالجة عند الكتابة في ملف السجل
    Private ReadOnly _logLock As New Object()

    Private Const ELNADY As String = "PlayerManagementSystem"
    Private Const DB_FILENAME As String = "playersdata.db"

    ''' <summary>هل تم تهيئة قاعدة البيانات بنجاح؟</summary>
    Public ReadOnly Property IsInitialized As Boolean
        Get
            Return _isInitialized
        End Get
    End Property

    ''' <summary>آخر خطأ حدث في قاعدة البيانات</summary>
    Public ReadOnly Property LastError As String
        Get
            Return _lastError
        End Get
    End Property

    ''' <summary>المسار الكامل لملف قاعدة البيانات الحالي</summary>
    Public ReadOnly Property CurrentDatabasePath As String
        Get
            Return _dbPath
        End Get
    End Property

#End Region

#Region "Path and Initialization"

    ''' <summary>الحصول على مسار مجلد بيانات التطبيق (ProgramData)</summary>
    Public Function GetAppDataPath() As String
        Try
            Dim appDataPath As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                ELNADY)

            If Not Directory.Exists(appDataPath) Then
                Directory.CreateDirectory(appDataPath)
                SetProgramDataPermissions(appDataPath)
            End If

            Return appDataPath
        Catch ex As Exception
            LogError("GetAppDataPath", ex)
            Return Application.StartupPath
        End Try
    End Function

    ''' <summary>المسار الافتراضي لملف قاعدة البيانات</summary>
    Public Function GetDefaultDatabasePath() As String
        Return Path.Combine(GetAppDataPath(), DB_FILENAME)
    End Function

    ''' <summary>تهيئة قاعدة البيانات (إنشاء الملف أو نسخه أو فتحه)</summary>
    Public Function Initialize(Optional ByVal showMessages As Boolean = True) As Boolean
        Try
            _lastError = ""

            Dim programDataPath As String = GetDefaultDatabasePath()
            Dim appPath As String = Application.StartupPath
            Dim sourceDbPath As String = Path.Combine(appPath, DB_FILENAME)

            ' التحقق من صلاحيات الكتابة في المجلد
            Dim dirPath As String = Path.GetDirectoryName(programDataPath)
            If Not HasWritePermission(dirPath) Then
                _lastError = "لا توجد صلاحيات كتابة في المجلد المحدد لقاعدة البيانات."
                If showMessages Then
                    MessageBox.Show(_lastError, "خطأ في الصلاحيات", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return False
            End If

            ' قاعدة البيانات موجودة مسبقاً
            If File.Exists(programDataPath) Then
                _dbPath = programDataPath
                _connectionString = BuildConnectionString(_dbPath)
                _isInitialized = True
                ApplyDatabasePragmas()
                CleanOldLogs()
                Return True
            End If

            ' نسخها من مجلد التطبيق إذا وُجدت
            If File.Exists(sourceDbPath) Then
                Try
                    File.Copy(sourceDbPath, programDataPath, overwrite:=True)
                    _dbPath = programDataPath
                    _connectionString = BuildConnectionString(_dbPath)
                    _isInitialized = True
                    LogInfo("Initialize", $"تم نسخ قاعدة البيانات من {sourceDbPath} إلى {programDataPath}")
                    ApplyDatabasePragmas()
                    CleanOldLogs()
                    Return True
                Catch ex As Exception
                    LogError("Initialize - Copy", ex)
                    _lastError = "فشل نسخ قاعدة البيانات: " & ex.Message
                    If showMessages Then
                        MessageBox.Show(_lastError, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    Return False
                End Try
            End If

            ' إنشاء قاعدة بيانات جديدة فارغة
            Try
                SQLiteConnection.CreateFile(programDataPath)
                _dbPath = programDataPath
                _connectionString = BuildConnectionString(_dbPath)
                _isInitialized = True
                LogInfo("Initialize", $"تم إنشاء قاعدة بيانات جديدة في: {programDataPath}")
                ApplyDatabasePragmas()
                CleanOldLogs()
                Return True
            Catch ex As Exception
                LogError("Initialize - CreateFile", ex)
                _lastError = "فشل إنشاء قاعدة البيانات: " & ex.Message
                If showMessages Then
                    MessageBox.Show(_lastError, "خطأ فادح", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return False
            End Try

        Catch ex As Exception
            _lastError = ex.Message
            LogError("Initialize", ex)
            If showMessages Then
                MessageBox.Show($"خطأ في تهيئة قاعدة البيانات: {ex.Message}", "خطأ",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return False
        End Try
    End Function

    ''' <summary>تطبيق إعدادات الأداء (PRAGMAS) على مستوى الملف</summary>
    Private Sub ApplyDatabasePragmas()
        Try
            Using conn As SQLiteConnection = GetConnection()
                ' وضع WAL يحسن الأداء المتزامن ويبقى مفعلاً على الملف بشكل دائم
                Using cmd As New SQLiteCommand("PRAGMA journal_mode = WAL;", conn)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            LogError("ApplyDatabasePragmas", ex)
        End Try
    End Sub

    ''' <summary>بناء سلسلة الاتصال مع تفعيل Connection Pooling وإعدادات الأداء لكل اتصال</summary>
    Private Function BuildConnectionString(dbPath As String) As String
        ' 🌟 تفعيل Pooling يقلل من وقت فتح/إغلاق الاتصالات بنسبة 90%
        ' 🌟 إضافة Cache Size و Synchronous هنا تضمن أن كل اتصال جديد يرث الأداء العالي
        Return $"Data Source={dbPath};Version=3;Pooling=True;Max Pool Size=100;BusyTimeout=5000;Synchronous=Normal;Cache Size=10000;"
    End Function

    ''' <summary>التحقق من صلاحية الكتابة في مجلد</summary>
    Private Function HasWritePermission(folderPath As String) As Boolean
        Try
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If
            Dim testFile As String = Path.Combine(folderPath, Guid.NewGuid().ToString() & ".tmp")
            File.WriteAllText(testFile, "test")
            File.Delete(testFile)
            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>الحصول على كائن اتصال مفتوح بقاعدة البيانات</summary>
    Public Function GetConnection() As SQLiteConnection
        Try
            _lastError = ""

            If Not _isInitialized Then
                If Not Initialize(False) Then
                    Throw New InvalidOperationException("لم يتم تهيئة قاعدة البيانات: " & _lastError)
                End If
            End If

            Dim conn As New SQLiteConnection(_connectionString)
            conn.Open()
            Return conn

        Catch ex As Exception
            _lastError = ex.Message
            LogError("GetConnection", ex)
            Throw New Exception("فشل فتح اتصال بقاعدة البيانات: " & ex.Message, ex)
        End Try
    End Function

    ''' <summary>إغلاق الاتصال (اختياري)</summary>
    Public Sub CloseConnection(Optional ByVal conn As SQLiteConnection = Nothing)
        Try
            If conn IsNot Nothing Then
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
                conn.Dispose()
            End If
        Catch ex As Exception
            LogError("CloseConnection", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 🔴 H-02 — منح صلاحيات مجلد البيانات بنسخة مقيدة:
    ''' كان يُمنح "Users" كامل الصلاحيات (قراءة+كتابة+تعديل) فأي حساب على الجهاز —
    ''' أو أي برنامج تحت أي حساب — يستطيع نسخ/تعديل/حذف القاعدة كاملة.
    ''' الآن صلاحية التعديل فقط لـ:
    '''   1) المستخدم الحالي   2) Administrators (كامل)   3) SYSTEM (كامل)
    '''   4) حسابات إضافية من الإعداد DataFolderAccounts (مفصولة بفواصل) — لوضع تعدد المستخدمين
    ''' ولا تُمنح مجموعة Users أي صلاحية صريحة جديدة (تبقى للوراثة الافتراضية من ProgramData قراءة فقط).
    ''' </summary>
    Public Sub SetProgramDataPermissions(Optional folderPath As String = "")
        Try
            If String.IsNullOrEmpty(folderPath) Then
                folderPath = GetAppDataPath()
            End If

            If Directory.Exists(folderPath) Then
                Dim di As New DirectoryInfo(folderPath)
                Dim ds As DirectorySecurity = di.GetAccessControl()

                Dim modifyRights As FileSystemRights =
                    FileSystemRights.ReadAndExecute Or FileSystemRights.Write Or FileSystemRights.Modify
                Dim inheritAll As InheritanceFlags =
                    InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit

                ds.AddAccessRule(New FileSystemAccessRule(
                    Environment.UserName, modifyRights, inheritAll, PropagationFlags.None, AccessControlType.Allow))
                ds.AddAccessRule(New FileSystemAccessRule(
                    "SYSTEM", FileSystemRights.FullControl, inheritAll, PropagationFlags.None, AccessControlType.Allow))
                ds.AddAccessRule(New FileSystemAccessRule(
                    "Administrators", FileSystemRights.FullControl, inheritAll, PropagationFlags.None, AccessControlType.Allow))

                ' حسابات مخوّلة إضافية — مثال: "ahmad, sales1" حين يستعمل البرنامج أكثر من حساب ويندوز
                Dim extraAccounts As String = AppSettingsStore.GetSetting(AppSettingsStore.Key_DataFolderAccounts, "")
                For Each acct As String In extraAccounts.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim cleanAccount As String = acct.Trim()
                    If cleanAccount = "" Then Continue For
                    ds.AddAccessRule(New FileSystemAccessRule(
                        cleanAccount, modifyRights, inheritAll, PropagationFlags.None, AccessControlType.Allow))
                Next

                di.SetAccessControl(ds)
                LogInfo("SetProgramDataPermissions", $"🔴 H-02: صلاحيات مقيدة (تعديل للمخوّلين فقط) — {folderPath}")
            End If
        Catch ex As Exception
            LogError("SetProgramDataPermissions", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 🔴 H-02 — تضييق صلاحيات مجلد بيانات موجود مسبقاً (مرة واحدة عبر علم في AppSettings).
    ''' التثبيتات القديمة أنشأت المجلد قبل هذا الإصلاح وورثت القاعدة الواسعة —
    ''' هذا الاستدعاء يطبّق التقييد عليها في أول إقلاع بعد التحديث.
    ''' الفشل لا يكتب العلم → يُعاد المحاولة تلقائياً (نفس سلوك M-01 ذاتي الشفاء).
    ''' يُستدعى من frmPlayers_Load بعد EnsureAppSettingsTableExists.
    ''' </summary>
    Public Sub TightenDataFolderPermissionsOnce()
        Try
            If AppSettingsStore.GetSetting(AppSettingsStore.Key_DataFolderAclTightened, "0") = "1" Then Return

            SetProgramDataPermissions(GetAppDataPath())
            AppSettingsStore.SetSetting(AppSettingsStore.Key_DataFolderAclTightened, "1")
            LogInfo("TightenDataFolderPermissionsOnce", "🔴 H-02: تم تقييد صلاحيات مجلد البيانات الموجود مسبقاً")
        Catch ex As Exception
            LogError("TightenDataFolderPermissionsOnce", ex)
        End Try
    End Sub

#End Region

#Region "Synchronous Query Execution"

    ''' <summary>تحديد نوع البارامترات لتجنب أخطاء Int64</summary>
    Private Sub FixParameterTypes(parameters As SQLiteParameter())
        If parameters Is Nothing Then Return
        For Each p As SQLiteParameter In parameters
            If p.Value IsNot Nothing AndAlso Not IsDBNull(p.Value) Then
                If TypeOf p.Value Is Long Then
                    p.DbType = DbType.Int64
                ElseIf TypeOf p.Value Is Integer Then
                    p.DbType = DbType.Int32
                ElseIf TypeOf p.Value Is Decimal Then
                    p.DbType = DbType.Decimal
                ElseIf TypeOf p.Value Is Byte() Then
                    p.DbType = DbType.Binary
                End If
            End If
        Next
    End Sub

    ''' <summary>تنفيذ استعلام وإرجاع DataTable (متزامن)</summary>
    Public Function ExecuteQuery(sql As String, ParamArray parameters As SQLiteParameter()) As DataTable
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim dt As New DataTable()
        Try
            _lastError = ""
            Using conn As SQLiteConnection = GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandTimeout = 60
                    If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                        FixParameterTypes(parameters)
                        cmd.Parameters.AddRange(parameters)
                    End If
                    Using da As New SQLiteDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            _lastError = ex.Message
            LogError("ExecuteQuery", ex, sql)
            Throw
        Finally
            sw.Stop()
            If sw.ElapsedMilliseconds > 500 Then
                LogInfo("ExecuteQuery", $"استعلام بطيء ({sw.ElapsedMilliseconds}ms): {sql}")
            End If
        End Try
        Return dt
    End Function

    ''' <summary>تنفيذ استعلام إدراج/تحديث/حذف وإرجاع عدد الصفوف (متزامن)</summary>
    Public Function ExecuteNonQuery(sql As String, ParamArray parameters As SQLiteParameter()) As Integer
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim result As Integer = 0
        Try
            _lastError = ""
            Using conn As SQLiteConnection = GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandTimeout = 60
                    If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                        FixParameterTypes(parameters)
                        cmd.Parameters.AddRange(parameters)
                    End If
                    result = cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            _lastError = ex.Message
            LogError("ExecuteNonQuery", ex, sql)
            Throw
        Finally
            sw.Stop()
            If sw.ElapsedMilliseconds > 500 Then
                LogInfo("ExecuteNonQuery", $"استعلام بطيء ({sw.ElapsedMilliseconds}ms): {sql}")
            End If
        End Try
        Return result
    End Function

    ''' <summary>تنفيذ استعلام وإرجاع قيمة أول عمود في أول صف (متزامن)</summary>
    Public Function ExecuteScalar(sql As String, ParamArray parameters As SQLiteParameter()) As Object
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Try
            _lastError = ""
            Using conn As SQLiteConnection = GetConnection()
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandTimeout = 60
                    If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                        FixParameterTypes(parameters)
                        cmd.Parameters.AddRange(parameters)
                    End If
                    Return cmd.ExecuteScalar()
                End Using
            End Using
        Catch ex As Exception
            _lastError = ex.Message
            LogError("ExecuteScalar", ex, sql)
            Throw
        Finally
            sw.Stop()
            If sw.ElapsedMilliseconds > 500 Then
                LogInfo("ExecuteScalar", $"استعلام بطيء ({sw.ElapsedMilliseconds}ms): {sql}")
            End If
        End Try
    End Function

    ''' <summary>تنفيذ عدة استعلامات في معاملة واحدة (متزامن)</summary>
    Public Function ExecuteTransaction(queries As List(Of Tuple(Of String, SQLiteParameter()))) As Boolean
        If queries Is Nothing OrElse queries.Count = 0 Then Return True

        Using conn As SQLiteConnection = GetConnection()
            Dim trans As SQLiteTransaction = Nothing
            Try
                trans = conn.BeginTransaction(IsolationLevel.Serializable)

                For Each query In queries
                    Using cmd As New SQLiteCommand(query.Item1, conn, trans)
                        cmd.CommandType = CommandType.Text
                        cmd.CommandTimeout = 60
                        If query.Item2 IsNot Nothing AndAlso query.Item2.Length > 0 Then
                            FixParameterTypes(query.Item2)
                            cmd.Parameters.AddRange(query.Item2)
                        End If
                        cmd.ExecuteNonQuery()
                    End Using
                Next

                trans.Commit()
                Return True

            Catch ex As Exception
                Try
                    If trans IsNot Nothing Then trans.Rollback()
                Catch rollbackEx As Exception
                    LogError("ExecuteTransaction - Rollback", rollbackEx)
                End Try

                _lastError = ex.Message
                LogError("ExecuteTransaction", ex)
                Return False

            Finally
                If trans IsNot Nothing Then trans.Dispose()
            End Try
        End Using
    End Function

#End Region

#Region "Asynchronous Query Execution"

    ''' <summary>تنفيذ استعلام وإرجاع DataTable (غير متزامن مع إلغاء)</summary>
    Public Async Function ExecuteQueryAsync(sql As String,
                                             cancellationToken As CancellationToken,
                                             ParamArray parameters As SQLiteParameter()) As Task(Of DataTable)
        cancellationToken.ThrowIfCancellationRequested()
        Return Await Task.Run(Function() ExecuteQuery(sql, parameters), cancellationToken)
    End Function

    ''' <summary>تنفيذ استعلام إدراج/تحديث/حذف (غير متزامن مع إلغاء)</summary>
    Public Async Function ExecuteNonQueryAsync(sql As String,
                                                cancellationToken As CancellationToken,
                                                ParamArray parameters As SQLiteParameter()) As Task(Of Integer)
        cancellationToken.ThrowIfCancellationRequested()
        Return Await Task.Run(Function() ExecuteNonQuery(sql, parameters), cancellationToken)
    End Function

    ''' <summary>تنفيذ استعلام وإرجاع قيمة مفردة (غير متزامن مع إلغاء)</summary>
    Public Async Function ExecuteScalarAsync(sql As String,
                                              cancellationToken As CancellationToken,
                                              ParamArray parameters As SQLiteParameter()) As Task(Of Object)
        cancellationToken.ThrowIfCancellationRequested()
        Return Await Task.Run(Function() ExecuteScalar(sql, parameters), cancellationToken)
    End Function

    ''' <summary>تنفيذ استعلام وإرجاع DataTable (غير متزامن)</summary>
    Public Async Function ExecuteQueryAsync(sql As String,
                                             ParamArray parameters As SQLiteParameter()) As Task(Of DataTable)
        Return Await Task.Run(Function() ExecuteQuery(sql, parameters))
    End Function

    ''' <summary>تنفيذ استعلام إدراج/تحديث/حذف (غير متزامن)</summary>
    Public Async Function ExecuteNonQueryAsync(sql As String,
                                                ParamArray parameters As SQLiteParameter()) As Task(Of Integer)
        Return Await Task.Run(Function() ExecuteNonQuery(sql, parameters))
    End Function

    ''' <summary>تنفيذ استعلام وإرجاع قيمة مفردة (غير متزامن)</summary>
    Public Async Function ExecuteScalarAsync(sql As String,
                                              ParamArray parameters As SQLiteParameter()) As Task(Of Object)
        Return Await Task.Run(Function() ExecuteScalar(sql, parameters))
    End Function

#End Region

#Region "Database Information"

    ''' <summary>التحقق من وجود جدول معين</summary>
    Public Function TableExists(tableName As String) As Boolean
        Try
            Dim sql As String = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@name"
            Dim result As Object = ExecuteScalar(sql, New SQLiteParameter("@name", tableName))
            Return Convert.ToInt32(result) > 0
        Catch ex As Exception
            LogError("TableExists", ex)
            Return False
        End Try
    End Function

    ''' <summary>التحقق من وجود عمود معين في جدول</summary>
    Public Function ColumnExists(tableName As String, columnName As String) As Boolean
        Try
            Using conn As SQLiteConnection = GetConnection()
                Dim sql As String = $"PRAGMA table_info([{tableName.Replace("'", "''")}])"
                Using cmd As New SQLiteCommand(sql, conn)
                    Using reader As SQLiteDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If reader("name").ToString().Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                                Return True
                            End If
                        End While
                    End Using
                End Using
            End Using
            Return False
        Catch ex As Exception
            LogError("ColumnExists", ex)
            Return False
        End Try
    End Function

    ''' <summary>🌟 M-02: هل الفهرس موجود؟ — استعلام خفيف على sqlite_master</summary>
    Public Function IndexExists(indexName As String) As Boolean
        Try
            Dim result As Object = ExecuteScalar(
                "SELECT COUNT(*) FROM sqlite_master WHERE type = 'index' AND name = @n",
                New SQLiteParameter("@n", indexName))
            Return result IsNot Nothing AndAlso Convert.ToInt32(result) > 0
        Catch ex As Exception
            LogError("IndexExists", ex)
            Return True   ' عند الشك نتصرف كأنه موجود — لا نلمس المخطط بلا داعٍ
        End Try
    End Function

    ''' <summary>🌟 M-02: يحذف فهارس قديمة فقط إن وُجدت فعلاً — بعد أول تنظيف لا يلمس شيئاً</summary>
    Private Sub DropLegacyIndexesIfExist(names As String())
        For Each name As String In names
            If IndexExists(name) Then
                ExecuteNonQuery($"DROP INDEX IF EXISTS [{name}]")
                LogInfo("DropLegacyIndexes", $"حُذف الفهرس القديم [{name}]")
            End If
        Next
    End Sub

    ''' <summary>الحصول على مخطط جدول (PRAGMA table_info)</summary>
    Public Function GetTableSchema(tableName As String) As DataTable
        Dim sql = $"PRAGMA table_info([{tableName.Replace("'", "''")}])"
        Return ExecuteQuery(sql)
    End Function

#End Region

#Region "Table Creation and Migration"

    ''' <summary>إنشاء جدول اللاعبين إذا لم يكن موجوداً</summary>
    Public Sub EnsurePlayersTableExists()
        Try
            Dim sql As String =
                "CREATE TABLE IF NOT EXISTS Players (" &
                "Playerid   INTEGER PRIMARY KEY, " &
                "Playername  TEXT, " &
                "birthdate   TEXT, " &
                "fathername  TEXT, " &
                "fatherid    INTEGER, " &
                "wifename    TEXT, " &
                "wifeid      INTEGER, " &
                "fone        TEXT, " &
                "altfone     TEXT, " &
                "address     TEXT, " &
                "datein      TEXT, " &
                "Notes       TEXT, " &
                "gender      TEXT, " &
                "MaritalStatus TEXT, " &
                "Team        TEXT, " &
                "Rolle       TEXT, " &
                "JobTitle    TEXT, " &
                "playerphoto BLOB, " &
                $"{AppConstants.Col_LastModifiedBy} TEXT)"
            ExecuteNonQuery(sql)

            ' ترحيل آمن: إضافة العمود إذا كان الجدول قديماً لا يملكه
            AlterTableAddColumnIfMissing("Players", AppConstants.Col_LastModifiedBy, "TEXT")

            LogInfo("EnsurePlayersTableExists", "تم التحقق من جدول Players")
        Catch ex As Exception
            LogError("EnsurePlayersTableExists", ex)
            Throw
        End Try
    End Sub

    ''' <summary>إنشاء جدول الأرشيف إذا لم يكن موجوداً</summary>
    Public Sub EnsureArchiveTableExists()
        Try
            Dim sql As String =
                "CREATE TABLE IF NOT EXISTS archive (" &
                "Playerid    INTEGER PRIMARY KEY, " &
                "Playername  TEXT, " &
                "birthdate   TEXT, " &
                "fathername  TEXT, " &
                "fatherid    INTEGER, " &
                "wifename    TEXT, " &
                "wifeid      INTEGER, " &
                "fone        TEXT, " &
                "altfone     TEXT, " &
                "address     TEXT, " &
                "datein      TEXT, " &
                "Notes       TEXT, " &
                "gender      TEXT, " &
                "MaritalStatus TEXT, " &
                "Team        TEXT, " &
                "Rolle       TEXT, " &
                "JobTitle    TEXT, " &
                "playerphoto BLOB, " &
                "DeletedDate TEXT, " &
                $"{AppConstants.Col_LastModifiedBy} TEXT)"
            ExecuteNonQuery(sql)

            ' ترحيل آمن: إضافة العمود إذا كان الجدول قديماً لا يملكه
            AlterTableAddColumnIfMissing("archive", AppConstants.Col_LastModifiedBy, "TEXT")

            LogInfo("EnsureArchiveTableExists", "تم التحقق من جدول archive")
        Catch ex As Exception
            LogError("EnsureArchiveTableExists", ex)
            Throw
        End Try
    End Sub

    ''' <summary>إنشاء جدول المصروفات إذا لم يكن موجوداً (ترحيل ذري داخل معاملة واحدة)</summary>
    Public Sub EnsureExpensesTableExists()
        Try
            ' 1. إذا كان الجدول غير موجود أصلاً، ننشئه بالهيكلة الصحيحة
            If Not TableExists("Expenses") Then
                Dim createSql As String = GetExpenseTableCreateSql()
                ExecuteNonQuery(createSql)
                CreateExpenseIndexes()   ' 🌟 الفهارس تنشأ مع الجدول الجديد أيضاً
                LogInfo("EnsureExpensesTableExists", "تم إنشاء جدول Expenses جديد.")
                Return
            End If

            ' 2. إذا كان الجدول موجوداً، نفحص هل يوجد عمود ID وهو Primary Key؟
            Dim schemaDt As DataTable = ExecuteQuery("PRAGMA table_info(Expenses)")
            Dim hasProperID As Boolean = False
            Dim oldColumns As New List(Of String)

            For Each row As DataRow In schemaDt.Rows
                Dim colName As String = row("name").ToString()
                Dim isPk As Integer = Convert.ToInt32(row("pk"))
                oldColumns.Add(colName)
                If colName.Equals("ID", StringComparison.OrdinalIgnoreCase) AndAlso isPk > 0 Then
                    hasProperID = True
                End If
            Next

            ' 3. إذا لم يوجد عمود ID سليم، نقوم بترحيل الجدول
            If Not hasProperID Then
                LogInfo("EnsureExpensesTableExists", "جاري ترحيل جدول Expenses لإضافة عمود ID...")

                ' 🌟 [تعديل 3-ب] نسخ فقط الأعمدة المشتركة بين الجدول القديم والجديد —
                ' نفس حماية جدول الإيرادات: أي عمود قديم زائد لن يفجّر الترحيل مستقبلاً
                Dim newExpenseCols As New List(Of String) From {
                    "VoucherNumber", "Description", "Amount", "CurrencyType", "ExpenseDate",
                    "Category", "Notes", "LastModifiedBy", "LastModifiedDate"}

                Dim colsToCopy As New List(Of String)
                For Each col In oldColumns
                    If Not col.Equals("ID", StringComparison.OrdinalIgnoreCase) AndAlso
                       newExpenseCols.Contains(col, StringComparer.OrdinalIgnoreCase) Then
                        colsToCopy.Add(col)
                    End If
                Next

                ' 🌟 استعلام إنشاء الجدول الجديد
                Dim createSql As String = GetExpenseTableCreateSql()

                ' 🌟 بناء كل خطوات الترحيل في قائمة واحدة
                Dim migrationQueries As New List(Of Tuple(Of String, SQLiteParameter()))

                ' الخطوة 1: إعادة تسمية الجدول القديم
                migrationQueries.Add(Tuple.Create("ALTER TABLE Expenses RENAME TO Old_Expenses_Temp",
                                                  New SQLiteParameter() {}))

                ' الخطوة 2: إنشاء الجدول الجديد بهيكلة صحيحة
                migrationQueries.Add(Tuple.Create(createSql, New SQLiteParameter() {}))

                ' الخطوة 3: نسخ البيانات (سيتم توليد ID جديد تلقائياً لكل سند)
                If colsToCopy.Count > 0 Then
                    Dim colNames As String = String.Join(", ", colsToCopy)
                    migrationQueries.Add(Tuple.Create(
                        $"INSERT INTO Expenses ({colNames}) SELECT {colNames} FROM Old_Expenses_Temp",
                        New SQLiteParameter() {}))
                End If

                ' الخطوة 4: حذف الجدول القديم
                migrationQueries.Add(Tuple.Create("DROP TABLE Old_Expenses_Temp",
                                                  New SQLiteParameter() {}))

                ' 🌟 تنفيذ كل الخطوات كمعاملة واحدة: إما نجاح كامل أو تراجع كامل — البيانات لا تضيع أبداً
                If Not ExecuteTransaction(migrationQueries) Then
                    Throw New Exception("فشل ترحيل جدول Expenses — تم التراجع عن العملية بالكامل والبيانات القديمة سليمة.")
                End If

                LogInfo("EnsureExpensesTableExists", "تم ترحيل جدول Expenses وإضافة عمود ID بنجاح.")
            End If

            ' ترحيل آمن لباقي الأعمدة إذا كانت قاعدة البيانات قديمة جداً
            AlterTableAddColumnIfMissing("Expenses", "LastModifiedBy", "TEXT")
            AlterTableAddColumnIfMissing("Expenses", "LastModifiedDate", "TEXT")
            ' ── الفهارس (التكرار المقصود: رقم السند الشامل يتكرر لعدة مستفيدين) ──
            CreateExpenseIndexes()

            ' 🌟 فحص تكرارات سندات المصروفات وتسجيلها باللوج فقط (كانت الدالة موجودة معطّلة أبداً)
            CheckDuplicateExpenseVoucherNumbers()

            ' ── 🌟 L-08: عمود Amount من REAL إلى NUMERIC (إعادة بناء ذرية تُنفذ مرة واحدة) ──
            RebuildAmountColumnIfReal("Expenses",
                New List(Of String) From {"VoucherNumber", "Description", "Amount", "CurrencyType",
                                          "ExpenseDate", "Category", "Notes", "LastModifiedBy", "LastModifiedDate"},
                GetExpenseTableCreateSql(), AddressOf CreateExpenseIndexes)

        Catch ex As Exception
            LogError("EnsureExpensesTableExists", ex)
            Throw
        End Try
    End Sub

    ''' <summary>فهارس جدول المصروفات — التكرار المقصود: رقم السند الشامل يتكرر لعدة مستفيدين
    ''' 🌟 M-02: تشغيل مشروط — لا DROP/CREATE إذا كل شيء بمكانه</summary>
    Private Sub CreateExpenseIndexes()
        Try
            If Not IndexExists("idx_expense_date") Then
                ' 🌟 فهرس التاريخ — للتقارير
                ExecuteNonQuery("CREATE INDEX IF NOT EXISTS idx_expense_date ON Expenses(ExpenseDate)")
            End If
            If Not IndexExists("idx_expense_voucher") Then
                ' 🌟 فهرس رقم السند (عادي مش فريد!) — لتسريع البحث برقم السند الشامل
                ExecuteNonQuery("CREATE INDEX IF NOT EXISTS idx_expense_voucher ON Expenses(VoucherNumber)")
            End If

            ' 🌟 إزالة الفهارس القديمة غير المطابقة للتصميم الجديد (مرة واحدة ثم يتخطى نفسه)
            DropLegacyIndexesIfExist(New String() {"idx_expense_voucher_unique", "idx_expense_beneficiary"})
        Catch idxEx As Exception
            LogError("CreateExpenseIndexes", idxEx)
        End Try
    End Sub
    ''' <summary>
    ''' 🌟 فحص أرقام سندات المصروفات المتكررة وتسجيلها باللوج فقط — دون أي تعديل على البيانات.
    ''' (كانت الدالة القديمة تُعيد تسمية التكرارات تلقائياً لكنها لم تُستدعَ قط؛ ومع سندات الصرف
    ''' الشاملة التكرار قد يكون مقصوداً، لذا النهج الصحيح هو التسجيل للمراجعة وليس التعديل التلقائي)
    ''' </summary>
    Private Sub CheckDuplicateExpenseVoucherNumbers()
        Try
            Dim dupDt As DataTable = ExecuteQuery(
                "SELECT VoucherNumber, COUNT(*) AS Cnt FROM Expenses " &
                "WHERE VoucherNumber IS NOT NULL AND TRIM(VoucherNumber) <> '' " &
                "GROUP BY VoucherNumber HAVING COUNT(*) > 1")

            If dupDt.Rows.Count = 0 Then Return

            For Each dupRow As DataRow In dupDt.Rows
                Dim voucher As String = dupRow("VoucherNumber").ToString().Trim()
                Dim count As Integer = Convert.ToInt32(dupRow("Cnt"))
                LogInfo("CheckDuplicateExpenseVoucherNumbers",
                        $"تكرار يحتاج مراجعة دون تعديل: سند مصروف [{voucher}] مستخدم {count} مرة")
            Next

        Catch ex As Exception
            LogError("CheckDuplicateExpenseVoucherNumbers", ex)
        End Try
    End Sub

    ''' <summary>إنشاء جدول المدفوعات (مع إعادة بناء ذرية للجداول القديمة داخل معاملة واحدة)</summary>
    Public Sub EnsurePaymentsTableExists()
        Try
            ' 🌟 حارس مرحلة الانتقال: لو الجدول القديم Payment موجود والجديد Revenues غير موجود
            ' → التسمية لم تتم بعد → ننفذها الآن فوراً (نفس المنطق الذري، مضمون التشغيل مرة واحدة)
            If Not TableExists(AppConstants.Table_Payments) AndAlso TableExists("Payment") Then
                LogInfo("EnsurePaymentsTableExists", "يوجد جدول Payment قديم — تنفيذ إعادة التسمية إلى Revenues أولاً...")
                RenamePaymentToRevenues()
            End If

            ' ── 1. الجدول غير موجود أصلاً: إنشاء بالهيكلة الكاملة ──
            If Not TableExists(AppConstants.Table_Payments) Then
                CreatePaymentTable()
                CreatePaymentIndexes()
                LogInfo("EnsurePaymentsTableExists", "تم إنشاء جدول Payment جديد بهيكلة كاملة.")
                Return
            End If

            ' ── 2. فحص هل الجدول الحالي يملك عمود ID وهو Primary Key ──
            Dim schemaDt As DataTable = ExecuteQuery($"PRAGMA table_info({AppConstants.Table_Payments})")
            Dim hasProperID As Boolean = False
            Dim oldColumns As New List(Of String)

            For Each row As DataRow In schemaDt.Rows
                Dim colName As String = row("name").ToString()
                Dim isPk As Integer = Convert.ToInt32(row("pk"))
                oldColumns.Add(colName)
                If colName.Equals("ID", StringComparison.OrdinalIgnoreCase) AndAlso isPk = 1 Then
                    hasProperID = True
                End If
            Next

            ' ── 3. إعادة بناء الجدول إذا كان قديماً بلا ID سليم (🌟 داخل معاملة واحدة) ──
            If Not hasProperID Then
                LogInfo("EnsurePaymentsTableExists", "جاري إعادة بناء جدول Payment لإضافة عمود ID...")

                ' 🌟 [تعديل 3-أ] نسخ فقط الأعمدة الموجودة في الجدول القديم والجديد معاً —
                ' النسخ الأعمى كان ينسخ عمود Status القديم فيفشل الإدراج
                ' (خطأ: table Revenues has no column named Status → تراجع كامل → تعطل الإقلاع)
                Dim newRevenueCols As New List(Of String) From {
                    "PlayerId", "Playername", "PaymentDate", "VoucherNumber", "Amount",
                    "PaymentMethod", "TransferorName", "Notes", "DueDate", "CurrencyType",
                    "CreatedDate", "LastModifiedBy", "LastModifiedDate", "VoucherImage"}

                Dim colsToCopy As New List(Of String)
                For Each col In oldColumns
                    If Not col.Equals("ID", StringComparison.OrdinalIgnoreCase) AndAlso
                       newRevenueCols.Contains(col, StringComparer.OrdinalIgnoreCase) Then
                        colsToCopy.Add(col)
                    End If
                Next

                ' 🌟 بناء كل خطوات الترحيل في قائمة واحدة
                Dim migrationQueries As New List(Of Tuple(Of String, SQLiteParameter()))

                ' الخطوة 1: حذف أي بقايا من محاولة سابقة + إعادة تسمية الجدول القديم
                migrationQueries.Add(Tuple.Create("DROP TABLE IF EXISTS Old_Payment_Temp",
                                                  New SQLiteParameter() {}))
                migrationQueries.Add(Tuple.Create(
                    $"ALTER TABLE {AppConstants.Table_Payments} RENAME TO Old_Payment_Temp",
                    New SQLiteParameter() {}))

                ' الخطوة 2: إنشاء الجدول الجديد بهيكلة صحيحة
                migrationQueries.Add(Tuple.Create(GetPaymentTableCreateSql(),
                                                  New SQLiteParameter() {}))

                ' الخطوة 3: نسخ البيانات (سيتم توليد ID جديد تلقائياً لكل دفعة)
                If colsToCopy.Count > 0 Then
                    Dim colNames As String = String.Join(", ", colsToCopy.Select(Function(c) "[" & c & "]"))
                    migrationQueries.Add(Tuple.Create(
                        $"INSERT INTO {AppConstants.Table_Payments} ({colNames}) SELECT {colNames} FROM Old_Payment_Temp",
                        New SQLiteParameter() {}))
                End If

                ' الخطوة 4: حذف الجدول القديم (مع فهارسه) — بعدها تُنشأ الفهارس الجديدة بأسماء حرة
                migrationQueries.Add(Tuple.Create("DROP TABLE Old_Payment_Temp",
                                                  New SQLiteParameter() {}))

                ' 🌟 تنفيذ كل الخطوات كمعاملة واحدة — إما نجاح كامل أو تراجع كامل
                If Not ExecuteTransaction(migrationQueries) Then
                    Throw New Exception("فشل إعادة بناء جدول Payment — تم التراجع عن العملية بالكامل والبيانات القديمة سليمة.")
                End If

                LogInfo("EnsurePaymentsTableExists", "تمت إعادة بناء جدول Payment وإضافة عمود ID بنجاح.")
            End If

            ' ── 4. ترحيل آمن للأعمدة إذا كانت قاعدة البيانات قديمة ──
            ' يجب إضافة الأعمدة قبل إنشاء الفهارس؛ وإلا يفشل إنشاء الفهرس
            ' بصمت في قواعد البيانات القديمة ولا تتم المحاولة مرة أخرى.
            AlterTableAddColumnIfMissing(AppConstants.Table_Payments, "DueDate", "TEXT")
            AlterTableAddColumnIfMissing(AppConstants.Table_Payments, "CurrencyType", "TEXT DEFAULT 'شيكل'")
            AlterTableAddColumnIfMissing(AppConstants.Table_Payments, "CreatedDate", "TEXT")
            AlterTableAddColumnIfMissing(AppConstants.Table_Payments, "LastModifiedBy", "TEXT")
            AlterTableAddColumnIfMissing(AppConstants.Table_Payments, "LastModifiedDate", "TEXT")
            AlterTableAddColumnIfMissing(AppConstants.Table_Payments, "VoucherImage", "BLOB")
            ' 🌟 [تعديل 5] حماية إضافية: لو وُجد جدول إيرادات قديم بلا PlayerId (سيناريو نظري) يُضاف تلقائياً
            AlterTableAddColumnIfMissing(AppConstants.Table_Payments, "PlayerId", "INTEGER")

            ' ── 5. فحص التكرارات دون تغيير أرقام السندات ──
            ' السند الشامل قد يستخدم الرقم نفسه لعدة لاعبين أو عملات،
            ' لذلك لا يجوز إعادة تسمية التكرارات اعتماداً على VoucherNumber وحده.
            DedupePaymentVoucherNumbers()

            ' ── 6. الفهارس (بما فيها الفهرس الفريد لرقم السند واللاعب والعملة) ──
            CreatePaymentIndexes()

            ' ── 7. ترحيل بيانات: ربط الدفعات القديمة برقم هوية اللاعب (آمن للتشغيل المتكرر)
            LinkPaymentsToPlayerIds()

            ' ── 8. 🔴 H-04: جدول صور السندات المستقل + ترحيل لمرة واحدة (صورة واحدة لكل سند)
            EnsureVoucherImagesTableExists()
            MigrateVoucherImagesToTable()

            ' ── 9. 🌟 L-08: عمود Amount من REAL إلى NUMERIC (إعادة بناء ذرية تُنفذ مرة واحدة) ──
            RebuildAmountColumnIfReal(AppConstants.Table_Payments,
                New List(Of String) From {"PlayerId", "Playername", "PaymentDate", "VoucherNumber",
                                          "Amount", "PaymentMethod", "TransferorName", "Notes", "DueDate",
                                          "CurrencyType", "CreatedDate", "LastModifiedBy", "LastModifiedDate", "VoucherImage"},
                GetPaymentTableCreateSql(), AddressOf CreatePaymentIndexes)

            LogInfo("EnsurePaymentsTableExists", "تم التحقق من جدول Payment")

        Catch ex As Exception
            LogError("EnsurePaymentsTableExists", ex)
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' 🌟 ترحيل لمرة واحدة: إعادة تسمية جدول Payment إلى Revenues + فهرسة الأسماء الجديدة
    ''' ذرية بالكامل — آمنة للتشغيل المتكرر (تتخطى إذا التسمية تمت سابقاً)
    ''' </summary>
    Public Sub RenamePaymentToRevenues()
        Try
            ' الحالتان: الجدول القديم موجود → نعيد التسمية / أو الجديد موجود → تمت سابقاً
            Dim paymentExists As Boolean = TableExists("Payment")
            Dim revenuesExists As Boolean = TableExists("Revenues")

            If Not paymentExists AndAlso revenuesExists Then
                ' تمت سابقاً — نتأكد فقط من أسماء الفهارس الجديدة
                RenamePaymentIndexes()
                LogInfo("RenamePaymentToRevenues", "الجدول Revenues موجود مسبقاً — تم التحقق من الفهارس")
                Return
            End If

            If Not paymentExists Then
                Throw New Exception("لا يوجد جدول Payment ولا Revenues — حالة غير متوقعة!")
            End If

            ' 🌟 إعادة التسمية داخل معاملة — تنتقل الفهارس مع الجدول تلقائياً
            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))
            queries.Add(Tuple.Create("ALTER TABLE Payment RENAME TO Revenues", New SQLiteParameter() {}))

            If Not ExecuteTransaction(queries) Then
                Throw New Exception("فشل إعادة تسمية Payment → Revenues — تم التراجع والبيانات سليمة.")
            End If

            ' إعادة تسمية الفهارس بأسماء جديدة (اختياري — للنظافة)
            RenamePaymentIndexes()

            LogInfo("RenamePaymentToRevenues", "تم إعادة تسمية جدول Payment → Revenues بنجاح مع فهارسه")

        Catch ex As Exception
            LogError("RenamePaymentToRevenues", ex)
            Throw
        End Try
    End Sub

    ''' <summary>🌟 M-04 — التعريف المرجعي الوحيد لفهارس جدول الإيرادات.
    ''' كانت الفهارس معرّفة بموضعين متضاربين بينهما حذف وإنشاء متقاطع —
    ''' الآن كلا الاسمن يمر من هنا حصراً (صفر تغيير على مواضع الاستدعاء).
    ''' الفريد على (رقم السند + اللاعب + العملة):
    '''   نفس الشخص مرتين بنفس العملة = مرفوض (تكرار غلط)
    '''   نفس الشخص مرتين بعملتين مختلفتين = مسموح (سند شامل متعدد العملات)
    ''' 🌟 M-02 — تشغيل مشروط: إذا كل الفهارس بمكانها لا يلمس المخطط إطلاقاً.</summary>
    Private Sub EnsureRevenueIndexes()
        Try
            ' ── فهارس البحث العادية ──
            If Not IndexExists("idx_revenue_playername") Then
                ExecuteNonQuery($"CREATE INDEX IF NOT EXISTS idx_revenue_playername ON {AppConstants.Table_Payments}(Playername)")
            End If
            If Not IndexExists("idx_revenue_date") Then
                ExecuteNonQuery($"CREATE INDEX IF NOT EXISTS idx_revenue_date ON {AppConstants.Table_Payments}(PaymentDate)")
            End If
            If Not IndexExists("idx_revenue_voucher") Then
                ExecuteNonQuery($"CREATE INDEX IF NOT EXISTS idx_revenue_voucher ON {AppConstants.Table_Payments}(VoucherNumber)")
            End If

            ' ── الفهارس الفريدة (التعريف الحديث بـ COALESCE العملة) ──
            If Not IndexExists("idx_revenue_voucher_player_cur_unique") Then
                ExecuteNonQuery(
                    $"CREATE UNIQUE INDEX IF NOT EXISTS idx_revenue_voucher_player_cur_unique ON {AppConstants.Table_Payments}(VoucherNumber, PlayerId, COALESCE(CurrencyType, '')) " &
                    "WHERE VoucherNumber IS NOT NULL AND TRIM(VoucherNumber) <> '' AND PlayerId IS NOT NULL")
            End If
            If Not IndexExists("idx_revenue_voucher_unlinked_cur_unique") Then
                ExecuteNonQuery(
                    $"CREATE UNIQUE INDEX IF NOT EXISTS idx_revenue_voucher_unlinked_cur_unique ON {AppConstants.Table_Payments}(VoucherNumber, COALESCE(CurrencyType, '')) " &
                    "WHERE VoucherNumber IS NOT NULL AND TRIM(VoucherNumber) <> '' AND PlayerId IS NULL")
            End If

            ' ── تنظيف أسماء قديمة (مرة واحدة تاريخياً ثم يتخطى نفسه) ──
            ' ملاحظة: idx_revenue_voucher_player_unique هو التعريف القديم قبل إضافة العملة — يُحذف عمداً
            DropLegacyIndexesIfExist(New String() {
                "idx_payment_playername", "idx_payment_date", "idx_payment_voucher",
                "idx_payment_voucher_player_unique", "idx_payment_voucher_unique",
                "idx_revenue_voucher_player_unique"})
        Catch ex As Exception
            LogError("EnsureRevenueIndexes", ex)
        End Try
    End Sub

    ''' <summary>🌟 M-04: كانت تعريفاً مستقلاً متضارباً — أصبحت تمر من التعريف المرجعي الوحيد</summary>
    Private Sub RenamePaymentIndexes()
        EnsureRevenueIndexes()
    End Sub
    ''' <summary>الهيكلة القياسية لجدول الإيرادات (نص SQL للاستخدام في الإنشاء والترحيل)</summary>
    Private Function GetPaymentTableCreateSql() As String
        Return $"CREATE TABLE {AppConstants.Table_Payments} (" &
               "ID            INTEGER PRIMARY KEY AUTOINCREMENT, " &
               "PlayerId      INTEGER, " &
               "Playername    TEXT, " &
               "PaymentDate   TEXT, " &
               "VoucherNumber TEXT, " &
               "Amount        NUMERIC, " &
               "PaymentMethod TEXT, " &
               "TransferorName TEXT, " &
               "Notes         TEXT, " &
               "DueDate       TEXT, " &
               "CurrencyType  TEXT DEFAULT 'شيكل', " &
               "CreatedDate   TEXT, " &
               "LastModifiedBy TEXT, " &
               "LastModifiedDate TEXT, " &
               "VoucherImage  BLOB)"
    End Function
    ''' <summary>إنشاء جدول المدفوعات</summary>
    Private Sub CreatePaymentTable()
        ExecuteNonQuery(GetPaymentTableCreateSql())
    End Sub

    ''' <summary>🌟 M-04: كانت تعريفاً مستقلاً متضارباً — أصبحت تمر من التعريف المرجعي الوحيد</summary>
    Private Sub CreatePaymentIndexes()
        EnsureRevenueIndexes()
    End Sub
    ''' <summary>الهيكلة القياسية لجدول المصروفات (نص SQL للاستخدام في الإنشاء والترحيل)
    ''' 🌟 L-08: تعريف مرجعي وحيد بدل نسختين مكررتين + عمود Amount بـ NUMERIC بدل REAL</summary>
    Private Function GetExpenseTableCreateSql() As String
        Return "CREATE TABLE Expenses (" &
               "ID            INTEGER PRIMARY KEY AUTOINCREMENT, " &
               "VoucherNumber TEXT, " &
               "Description   TEXT, " &
               "Amount        NUMERIC, " &
               "CurrencyType  TEXT, " &
               "ExpenseDate   TEXT, " &
               "Category      TEXT, " &
               "Notes         TEXT, " &
               "LastModifiedBy TEXT, " &
               "LastModifiedDate TEXT)"
    End Function

    ''' <summary>
    ''' 🌟 L-08 — دقة المبالغ المالية: تحويل عمود Amount من REAL إلى NUMERIC.
    ''' عمود REAL يخزّن كل المبالغ كأعداد فاصلة عائمة حتى الصحيحة منها (5000 → 5000.0)
    ''' مع تراكم أخطاء تقريب في الجمع والمقارنات عبر السنين؛ NUMERIC يخزّن المبلغ الصحيح
    ''' كعدد صحيح حقيقي والكسري كفاصلة عائمة — القراءة والكتابة من كل الكود الحالي
    ''' تعمل بلا أي تغيير (نفس القيم، دقة أفضل).
    ''' التنفيذ: إعادة بناء ذرية بنفس نمط الترحيل المجرَّب — إعادة تسمية، إنشاء،
    ''' نسخ الأعمدة المشتركة، حذف القديم — كلها معاملة واحدة؛ إما نجاح كامل أو تراجع كامل،
    ''' وتتخطى نفسها تماماً إذا تمت سابقاً (فحص نوع العمود من PRAGMA table_info).
    ''' </summary>
    Private Sub RebuildAmountColumnIfReal(tableName As String,
                                          newColumnList As List(Of String),
                                          createSql As String,
                                          indexRebuilder As Action)
        Try
            Dim dt As DataTable = ExecuteQuery($"PRAGMA table_info({tableName})")
            Dim amountIsReal As Boolean = False
            Dim oldColumns As New List(Of String)

            For Each row As DataRow In dt.Rows
                Dim colName As String = row("name").ToString()
                Dim colType As String = row("type").ToString().Trim().ToUpperInvariant()
                oldColumns.Add(colName)
                If colName.Equals("Amount", StringComparison.OrdinalIgnoreCase) AndAlso colType = "REAL" Then
                    amountIsReal = True
                End If
            Next

            ' تمت سابقاً (أو الجدول جديد أصلاً بـ NUMERIC) — نتخطى بلا أي لمسة
            If Not amountIsReal Then Return

            Dim countBefore As Long = Convert.ToInt64(ExecuteScalar($"SELECT COUNT(*) FROM [{tableName}]"))
            LogInfo("RebuildAmountColumnIfReal",
                    $"ترحيل L-08: تحويل عمود Amount بالجدول {tableName} من REAL إلى NUMERIC ({countBefore} سطراً)...")

            ' نسخ الأعمدة المشتركة فقط — نفس حماية ترحيلات ID القديمة
            Dim colsToCopy As New List(Of String)
            For Each col In oldColumns
                If Not col.Equals("ID", StringComparison.OrdinalIgnoreCase) AndAlso
                   newColumnList.Contains(col, StringComparer.OrdinalIgnoreCase) Then
                    colsToCopy.Add(col)
                End If
            Next

            Dim tempName As String = $"Old_{tableName}_Amount_Temp"
            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))

            queries.Add(Tuple.Create($"DROP TABLE IF EXISTS [{tempName}]", New SQLiteParameter() {}))
            queries.Add(Tuple.Create($"ALTER TABLE [{tableName}] RENAME TO [{tempName}]", New SQLiteParameter() {}))
            queries.Add(Tuple.Create(createSql, New SQLiteParameter() {}))
            If colsToCopy.Count > 0 Then
                Dim colNames As String = String.Join(", ", colsToCopy.Select(Function(c) "[" & c & "]"))
                queries.Add(Tuple.Create(
                    $"INSERT INTO [{tableName}] ({colNames}) SELECT {colNames} FROM [{tempName}]",
                    New SQLiteParameter() {}))
            End If
            queries.Add(Tuple.Create($"DROP TABLE [{tempName}]", New SQLiteParameter() {}))

            If Not ExecuteTransaction(queries) Then
                Throw New Exception($"فشل ترحيل L-08 للجدول {tableName} — تم التراجع بالكامل والبيانات القديمة سليمة.")
            End If

            ' الفهارس هُدمت مع الجدول القديم — يعاد إنشاؤها بأسماء حرة
            If indexRebuilder IsNot Nothing Then indexRebuilder()

            ' تحقق نهائي: عدد السطور بعد الترحيل يجب أن يطابق ما قبله تماماً
            Dim countAfter As Long = Convert.ToInt64(ExecuteScalar($"SELECT COUNT(*) FROM [{tableName}]"))
            If countAfter <> countBefore Then
                Throw New Exception($"ترحيل L-08 للجدول {tableName}: عدد السطور بعد الترحيل ({countAfter}) لا يطابق ما قبله ({countBefore}) — راجع اللوج فوراً.")
            End If

            ' تحقق من النوع الجديد فعلياً (حزام أمان إضافي)
            For Each chkRow As DataRow In ExecuteQuery($"PRAGMA table_info({tableName})").Rows
                If chkRow("name").ToString().Equals("Amount", StringComparison.OrdinalIgnoreCase) Then
                    Dim newType As String = chkRow("type").ToString().Trim().ToUpperInvariant()
                    If newType <> "NUMERIC" Then
                        Throw New Exception($"ترحيل L-08 للجدول {tableName}: نوع العمود بعد الترحيل [{newType}] ليس NUMERIC كما هو متوقع.")
                    End If
                End If
            Next

            LogInfo("RebuildAmountColumnIfReal",
                    $"تم تحويل عمود Amount بالجدول {tableName} إلى NUMERIC بنجاح ({countAfter} سطراً متطابقاً).")

        Catch ex As Exception
            LogError("RebuildAmountColumnIfReal", ex)
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' ربط الدفعات القديمة (بلا PlayerId) برقم هوية اللاعب عبر مطابقة الاسم —
    ''' يبحث في اللاعبين النشطين أولاً ثم الأرشيف. آمنة للتشغيل المتكرر (يفحص الفارغ فقط)
    ''' </summary>
    Private Sub LinkPaymentsToPlayerIds()
        Try
            If Not ColumnExists(AppConstants.Table_Payments, "PlayerId") Then Return

            ' تخطٍ سريع إذا لم يوجد أي سند غير مرتبط
            Dim nullCountObj As Object = ExecuteScalar(
                $"SELECT COUNT(*) FROM {AppConstants.Table_Payments} WHERE PlayerId IS NULL AND Playername IS NOT NULL")
            If Convert.ToInt32(nullCountObj) = 0 Then Return

            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))

            ' 1. الربط من جدول اللاعبين النشطين (الأولوية الأولى) — 🌟 ترتيب حتمي بأصغر رقم هوية
            ' (كان LIMIT 1 بدون ORDER BY يعطي نتيجة عشوائية مع الأسماء المكررة)
            queries.Add(Tuple.Create(
                $"UPDATE {AppConstants.Table_Payments} SET PlayerId = " &
                $"(SELECT P.{AppConstants.Col_PlayerId} FROM {AppConstants.Table_Players} P " &
                $" WHERE P.{AppConstants.Col_PlayerName} = {AppConstants.Table_Payments}.Playername ORDER BY P.{AppConstants.Col_PlayerId} ASC LIMIT 1) " &
                "WHERE PlayerId IS NULL AND Playername IS NOT NULL",
                New SQLiteParameter() {}))

            ' 2. الربط من الأرشيف لمن لم يُربط من النشطين — 🌟 ترتيب حتمي بأصغر رقم هوية
            queries.Add(Tuple.Create(
                $"UPDATE {AppConstants.Table_Payments} SET PlayerId = " &
                $"(SELECT A.{AppConstants.Col_PlayerId} FROM {AppConstants.Table_Archive} A " &
                $" WHERE A.{AppConstants.Col_PlayerName} = {AppConstants.Table_Payments}.Playername ORDER BY A.{AppConstants.Col_PlayerId} ASC LIMIT 1) " &
                "WHERE PlayerId IS NULL AND Playername IS NOT NULL",
                New SQLiteParameter() {}))

            If ExecuteTransaction(queries) Then
                Dim linkedObj As Object = ExecuteScalar(
                    $"SELECT COUNT(*) FROM {AppConstants.Table_Payments} WHERE PlayerId IS NOT NULL")
                LogInfo("LinkPaymentsToPlayerIds",
                        $"اكتمل ربط الدفعات القديمة — إجمالي الدفعات المرتبطة بلاعب: {Convert.ToInt32(linkedObj)}")
            Else
                LogError("LinkPaymentsToPlayerIds", New Exception("فشل ربط الدفعات القديمة — تم التراجع"))
            End If

        Catch ex As Exception
            ' لا نوقف تشغيل البرنامج بسبب الترحيل — السندات غير المرتبطة تعمل بالاسم المخزن
            LogError("LinkPaymentsToPlayerIds", ex)
        End Try
    End Sub

    ''' <summary>
    ''' فحص أرقام السندات المتكررة وتسجيلها فقط.
    ''' لا تتم إعادة تسمية أي سند تلقائياً؛ لأن رقم السند قد يكون سنداً شاملاً
    ''' صحيحاً مستخدماً لعدة لاعبين أو عملات.
    ''' </summary>
    Private Sub DedupePaymentVoucherNumbers()
        Try
            Dim dupDt As DataTable = ExecuteQuery(
                $"SELECT VoucherNumber, PlayerId, COALESCE(CurrencyType, '') AS CurrencyType, COUNT(*) AS Cnt FROM {AppConstants.Table_Payments} " &
                "WHERE VoucherNumber IS NOT NULL AND TRIM(VoucherNumber) <> '' " &
                "GROUP BY VoucherNumber, PlayerId, COALESCE(CurrencyType, '') HAVING COUNT(*) > 1")

            If dupDt.Rows.Count = 0 Then Return

            For Each dupRow As DataRow In dupDt.Rows
                Dim voucher As String = dupRow("VoucherNumber").ToString().Trim()
                Dim playerId As String = If(IsDBNull(dupRow("PlayerId")), "NULL", dupRow("PlayerId").ToString())
                Dim currency As String = dupRow("CurrencyType").ToString()
                Dim count As Integer = Convert.ToInt32(dupRow("Cnt"))

                LogInfo("DedupePaymentVoucherNumbers",
                        $"تكرار يحتاج مراجعة دون تعديل: السند [{voucher}]، اللاعب [{playerId}]، العملة [{currency}]، العدد [{count}]")
            Next

        Catch ex As Exception
            LogError("DedupePaymentVoucherNumbers", ex)
        End Try
    End Sub

    ''' <summary>إنشاء جدول المستخدمين (وإنشاء أدمن افتراضي إذا لم يكن موجوداً)</summary>
    Public Sub EnsureUsersTableExists()
        Try
            Dim sql As String =
                $"CREATE TABLE IF NOT EXISTS {AppConstants.Table_Users} (" &
                $"{AppConstants.Col_Username} TEXT PRIMARY KEY, " &
                $"{AppConstants.Col_PasswordHash} TEXT, " &
                $"{AppConstants.Col_CanEdit} INTEGER DEFAULT 1, " &
                $"{AppConstants.Col_CanDelete} INTEGER DEFAULT 1, " &
                $"{AppConstants.Col_MustChangePassword} INTEGER DEFAULT 0)"
            ExecuteNonQuery(sql)

            ' ترحيل آمن للأعمدة إذا كانت قاعدة البيانات قديمة
            AlterTableAddColumnIfMissing(AppConstants.Table_Users, AppConstants.Col_CanEdit, "INTEGER DEFAULT 1")
            AlterTableAddColumnIfMissing(AppConstants.Table_Users, AppConstants.Col_CanDelete, "INTEGER DEFAULT 1")
            AlterTableAddColumnIfMissing(AppConstants.Table_Users, AppConstants.Col_MustChangePassword, "INTEGER DEFAULT 0")

            ' التحقق من وجود مستخدم الأدمن
            Dim countObj As Object = ExecuteScalar($"SELECT COUNT(*) FROM {AppConstants.Table_Users} WHERE LOWER({AppConstants.Col_Username}) = 'admin'")

            If Convert.ToInt32(countObj) = 0 Then
                ' الحساب الأول مؤقت فقط، ويُجبر على تغيير كلمة المرور بعد الدخول الأول.
                Dim adminHash As String = UtilityModule.HashPassword("123456")
                ExecuteNonQuery(
                    $"INSERT INTO {AppConstants.Table_Users} ({AppConstants.Col_Username}, {AppConstants.Col_PasswordHash}, {AppConstants.Col_CanEdit}, {AppConstants.Col_CanDelete}, {AppConstants.Col_MustChangePassword}) VALUES (@u, @p, 1, 1, 1)",
                    New SQLiteParameter("@u", "admin"),
                    New SQLiteParameter("@p", adminHash))
                LogInfo("EnsureUsersTableExists", "تم إنشاء حساب المدير (admin) الافتراضي.")
            Else
                ' ترقية آمنة للحسابات القديمة التي ما زالت تستخدم كلمة المرور الافتراضية.
                Dim adminPassword As Object = ExecuteScalar(
                    $"SELECT {AppConstants.Col_PasswordHash} FROM {AppConstants.Table_Users} WHERE LOWER({AppConstants.Col_Username}) = 'admin' LIMIT 1")
                If adminPassword IsNot Nothing AndAlso Not IsDBNull(adminPassword) AndAlso
                   UtilityModule.VerifyPassword("123456", adminPassword.ToString()) Then
                    ExecuteNonQuery(
                        $"UPDATE {AppConstants.Table_Users} SET {AppConstants.Col_MustChangePassword} = 1 WHERE LOWER({AppConstants.Col_Username}) = 'admin'")
                End If
            End If
            ' ملاحظة: قمنا بحذف جملة الـ UPDATE التي كانت تعيد تعيين كلمة المرور في كل تشغيل!

        Catch ex As Exception
            LogError("EnsureUsersTableExists", ex)
            Throw
        End Try
    End Sub
    ''' <summary>جدول عناصر القوائم المنسدلة المخصصة (الفريق/الدور/المسمى/الحالة الاجتماعية...)</summary>
    Public Sub EnsureDropdownItemsTableExists()
        Try
            Dim sql As String =
                "CREATE TABLE IF NOT EXISTS DropdownItems (" &
                "FieldName  TEXT NOT NULL, " &
                "ItemValue  TEXT NOT NULL, " &
                "PRIMARY KEY (FieldName, ItemValue))"
            ExecuteNonQuery(sql)
            LogInfo("EnsureDropdownItemsTableExists", "تم التحقق من جدول DropdownItems")
        Catch ex As Exception
            LogError("EnsureDropdownItemsTableExists", ex)
            Throw
        End Try
    End Sub
    ''' <summary>إضافة عمود إلى جدول إذا كان مفقوداً (ترحيل آمن)</summary>
    Private Sub AlterTableAddColumnIfMissing(tableName As String, columnName As String, columnDefinition As String)
        Try
            If Not ColumnExists(tableName, columnName) Then
                ExecuteNonQuery($"ALTER TABLE [{tableName}] ADD COLUMN [{columnName}] {columnDefinition}")
                LogInfo("AlterTableAddColumnIfMissing", $"تمت إضافة العمود [{columnName}] إلى [{tableName}]")
            End If
        Catch ex As Exception
            If Not ex.Message.Contains("duplicate column") Then
                LogError("AlterTableAddColumnIfMissing", ex)
            End If
        End Try
    End Sub

#End Region

#Region "Date Migration"

    ''' <summary>
    ''' ترحيل آمن: تحويل التواريخ المخزنة dd/MM/yyyy إلى yyyy-MM-dd
    ''' آمن للتشغيل المتكرر — إذا لم يجد شيئاً لا يفعل شيئاً
    ''' </summary>
    Public Sub EnsureVoucherImagesTableExists()
        Try
            Dim existsObj As Object = ExecuteScalar(
                "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = 'VoucherImages'")
            If existsObj IsNot Nothing AndAlso Convert.ToInt32(existsObj) > 0 Then Return

            ExecuteNonQuery(
                "CREATE TABLE IF NOT EXISTS VoucherImages (" &
                "VoucherNumber    TEXT PRIMARY KEY, " &
                "ImageData        BLOB NOT NULL, " &
                "CreatedDate      TEXT, " &
                "LastModifiedDate TEXT)")
            LogInfo("EnsureVoucherImagesTableExists", "🔴 H-04: تم إنشاء جدول VoucherImages المستقل لصور السندات")
        Catch ex As Exception
            LogError("EnsureVoucherImagesTableExists", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 🔴 H-04 — ترحيل لمرة واحدة: صور السندات من عمود Payments.VoucherImage
    ''' إلى الجدول المستقل VoucherImages (نسخة واحدة لكل رقم سند بدل تكرار الـBLOB مع كل سطر).
    ''' آمنة للتشغيل المتكرر: علم إتمام في AppSettings + فحص فارغ سريع قبل أي شغل.
    ''' المجموعات المتضاربة (نفس رقم السند بصور مختلفة) لا تُلمس وتُسجّل تحذيراً — لا رأي صامت.
    ''' </summary>
    Public Sub MigrateVoucherImagesToTable()
        Try
            AppSettingsStore.EnsureAppSettingsTableExists()
            If AppSettingsStore.GetSetting(AppSettingsStore.Key_Migration_VoucherImages_Done, "0") = "1" Then Return

            ' لا صور أصلاً → سجّل الإتمام واخرج (شغل صفر بالإقلاعات القادمة)
            Dim anyCountObj As Object = ExecuteScalar(
                "SELECT COUNT(*) FROM " & AppConstants.Table_Payments &
                " WHERE VoucherNumber IS NOT NULL AND TRIM(VoucherNumber) <> '' AND VoucherImage IS NOT NULL")
            If Convert.ToInt32(anyCountObj) = 0 Then
                AppSettingsStore.SetSetting(AppSettingsStore.Key_Migration_VoucherImages_Done, "1")
                LogInfo("MigrateVoucherImages", "لا توجد صور سندات للترحيل — تم تسجيل الإتمام")
                Return
            End If

            ' المجموعات المكررة فعلاً: أكثر من سطر لنفس رقم السند وعليها صورة
            Dim groups As DataTable = ExecuteQuery(
                "SELECT VoucherNumber, COUNT(*) AS RowsCount, COUNT(DISTINCT VoucherImage) AS DistinctImages " &
                "FROM " & AppConstants.Table_Payments &
                " WHERE VoucherNumber IS NOT NULL AND TRIM(VoucherNumber) <> '' AND VoucherImage IS NOT NULL " &
                " GROUP BY VoucherNumber HAVING COUNT(*) > 1")

            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))
            Dim migratedCount As Integer = 0
            Dim nowStr As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)

            For Each g As DataRow In groups.Rows
                Dim voucher As String = Convert.ToString(g("VoucherNumber")).Trim()

                If Convert.ToInt32(g("DistinctImages")) > 1 Then
                    ' تعارض: نفس رقم السند بصور مختلفة — لا نختار رأياً صامتاً؛ نبقي الأعمدة كما هي
                    LogWarn("MigrateVoucherImages", $"تخطي رقم السند [{voucher}] — سطوره تحمل صوراً مختلفة، يلزم مراجعة يدوية")
                    Continue For
                End If

                ' نسخة الصورة الأولى (أصغر ID) إلى الجدول المستقل
                queries.Add(Tuple.Create(
                    "INSERT OR REPLACE INTO VoucherImages (VoucherNumber, ImageData, CreatedDate, LastModifiedDate) " &
                    "SELECT VoucherNumber, VoucherImage, @now, @now2 FROM " & AppConstants.Table_Payments &
                    " WHERE TRIM(VoucherNumber) = @v AND VoucherImage IS NOT NULL ORDER BY ID LIMIT 1",
                    New SQLiteParameter() {New SQLiteParameter("@v", voucher),
                                           New SQLiteParameter("@now", nowStr),
                                           New SQLiteParameter("@now2", nowStr)}))
                ' تفريغ عمود BLOB من كل سطور المجموعة — القراءة تصير من الجدول المستقل
                queries.Add(Tuple.Create(
                    "UPDATE " & AppConstants.Table_Payments & " SET VoucherImage = NULL WHERE TRIM(VoucherNumber) = @v AND VoucherImage IS NOT NULL",
                    New SQLiteParameter() {New SQLiteParameter("@v", voucher)}))
                migratedCount += 1
            Next

            If migratedCount > 0 Then
                Dim txOk As Boolean = ExecuteTransaction(queries)
                If Not txOk Then
                    LogWarn("MigrateVoucherImages", "فشلت معاملة ترحيل صور السندات — سيُعاد المحاولة بالإقلاع التالي")
                    Return   ' لا علم إتمام — سلوك ذاتي الشفاء مثل M-01
                End If
            End If

            AppSettingsStore.SetSetting(AppSettingsStore.Key_Migration_VoucherImages_Done, "1")
            LogInfo("MigrateVoucherImages", $"🔴 H-04: تم ترحيل صور {migratedCount} سنداً إلى الجدول المستقل VoucherImages")
        Catch ex As Exception
            LogError("MigrateVoucherImagesToTable", ex)
        End Try
    End Sub

    Public Sub MigrateDatesToIsoFormat()
        Try
            ' (الجدول، عمود التاريخ) — المفتاح دائماً rowid باسم مستعار ثابت
            Dim jobs As New List(Of Tuple(Of String, String)) From {
                Tuple.Create(AppConstants.Table_Players, AppConstants.Col_BirthDate),
                Tuple.Create(AppConstants.Table_Players, AppConstants.Col_DateIn),
                Tuple.Create(AppConstants.Table_Archive, AppConstants.Col_BirthDate),
                Tuple.Create(AppConstants.Table_Archive, AppConstants.Col_DateIn),
                Tuple.Create(AppConstants.Table_Payments, AppConstants.Col_Pay_PaymentDate),
                Tuple.Create(AppConstants.Table_Payments, AppConstants.Col_Pay_DueDate),
                Tuple.Create(AppConstants.Table_Expenses, AppConstants.Col_Exp_ExpenseDate)
            }

            Dim totalConverted As Integer = 0
            Dim hasFailure As Boolean = False   ' 🌟 M-01: لا نسجل الإتمام إذا فشل أي جدول

            For Each job In jobs
                Dim tableName As String = job.Item1
                Dim dateCol As String = job.Item2

                If Not TableExists(tableName) OrElse Not ColumnExists(tableName, dateCol) Then Continue For

                ' 🌟 كل مهمة معزولة — فشل جدول لا يوقف ترحيل الباقي
                Try
                    ' 🌟 rowid باسم مستعار "__rid": إذا كان الجدول يملك INTEGER PRIMARY KEY
                    ' (مثل Expenses)، SQLite يعيد rowid باسم المفتاح الأساسي بدل "rowid"
                    Dim dt As DataTable = ExecuteQuery(
                        $"SELECT [rowid] AS __rid, [{dateCol}] FROM [{tableName}] " &
                        $"WHERE [{dateCol}] IS NOT NULL AND TRIM([{dateCol}]) <> ''")

                    Dim updates As New List(Of Tuple(Of String, SQLiteParameter()))

                    For Each row As DataRow In dt.Rows
                        Dim isoDate As String = ConvertDmyToIso(UtilityModule.SafeString(row(dateCol)))
                        If isoDate IsNot Nothing Then
                            Dim rid As Long = Convert.ToInt64(row("__rid"))
                            Dim sql = $"UPDATE [{tableName}] SET [{dateCol}] = @v WHERE rowid = @k"
                            Dim prms = {
                                New SQLiteParameter("@v", isoDate),
                                New SQLiteParameter("@k", rid)
                            }
                            updates.Add(Tuple.Create(sql, prms))
                        End If
                    Next

                    If updates.Count > 0 Then
                        If ExecuteTransaction(updates) Then
                            totalConverted += updates.Count
                            LogInfo("MigrateDatesToIsoFormat", $"تم تحويل {updates.Count} قيمة في [{tableName}].[{dateCol}]")
                        Else
                            LogError("MigrateDatesToIsoFormat", New Exception($"فشل تحويل [{tableName}].[{dateCol}] — تم التراجع"))
                            hasFailure = True
                        End If
                    End If
                Catch ex As Exception
                    LogError("MigrateDatesToIsoFormat", New Exception($"خطأ أثناء معالجة [{tableName}].[{dateCol}]: {ex.Message}"))
                    hasFailure = True
                End Try
            Next

            If totalConverted > 0 Then
                LogInfo("MigrateDatesToIsoFormat", $"اكتمل الترحيل — تم تحويل {totalConverted} قيمة إجمالاً")
            End If

            ' 🌟 M-01: علم الإتمام — يُكتب فقط عند نجاح كامل بلا أي فشل
            ' الفشل يعني إعادة المحاولة تلقائياً في الإقلاع التالي (سلوك ذاتي الشفاء)
            If Not hasFailure Then
                AppSettingsStore.EnsureAppSettingsTableExists()
                AppSettingsStore.SetSetting(AppSettingsStore.Key_Migration_DatesIso_Done, "1")
                LogInfo("MigrateDatesToIsoFormat", "اكتمل ترحيل التواريخ نهائياً — لن يتكرر الفحص في الإقلاعات القادمة")
            End If

        Catch ex As Exception
            LogError("MigrateDatesToIsoFormat", ex)
        End Try
    End Sub

    ''' <summary>إذا كان النص بصيغة dd/MM/yyyy يعيد yyyy-MM-dd، وإلا يعيد Nothing (يتركه كما هو)</summary>
    Private Function ConvertDmyToIso(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Dim s As String = value.Trim()
        ' نمط dd/MM/yyyy: طول 10 وفواصل / في الموقعين 2 و5
        If s.Length = 10 AndAlso s(2) = "/"c AndAlso s(5) = "/"c Then
            Dim d As Date
            If Date.TryParseExact(s, "dd/MM/yyyy", Nothing, Globalization.DateTimeStyles.None, d) Then
                Return d.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture) ' 🌟 ثقافة ثابتة
            End If
        End If
        Return Nothing
    End Function

#End Region

#Region "Logging"

    ''' <summary>الحصول على مسار مجلد السجلات</summary>
    Private Function GetLogPath() As String
        Return Path.Combine(GetAppDataPath(), "Logs")
    End Function

    ''' <summary>تسجيل خطأ في ملف السجل</summary>
    Public Sub LogError(methodName As String, ex As Exception, Optional sql As String = "")
        Try
            Dim logPath As String = GetLogPath()
            Directory.CreateDirectory(logPath)
            Dim logFile As String = Path.Combine(logPath, $"ErrorLog_{DateTime.Now:yyyy-MM}.txt")

            ' استخدام كائن قفل مخصص لضمان سلامة الكتابة في بيئة متعددة المسارات (Thread-Safe)
            SyncLock _logLock
                Using writer As New StreamWriter(logFile, append:=True, encoding:=Encoding.UTF8)
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [خطأ] في {methodName}")
                    writer.WriteLine($"  الرسالة    : {ex.Message}")
                    If Not String.IsNullOrEmpty(sql) Then
                        writer.WriteLine($"  SQL        : {sql}")
                    End If
                    If ex.InnerException IsNot Nothing Then
                        writer.WriteLine($"  خطأ داخلي : {ex.InnerException.Message}")
                    End If
                    If ex.StackTrace IsNot Nothing Then
                        writer.WriteLine($"  Stack Trace: {ex.StackTrace}")
                    End If
                    writer.WriteLine(New String("-"c, 80))
                End Using
            End SyncLock
        Catch
            ' تجاهل أخطاء كتابة السجل نفسها لمنع توقف البرنامج
        End Try
    End Sub

    ''' <summary>تسجيل معلومات عامة في ملف السجل</summary>
    Public Sub LogInfo(methodName As String, message As String)
        Try
            Dim logPath As String = GetLogPath()
            Directory.CreateDirectory(logPath)
            Dim logFile As String = Path.Combine(logPath, $"AppLog_{DateTime.Now:yyyy-MM}.txt")

            SyncLock _logLock
                Using writer As New StreamWriter(logFile, append:=True, encoding:=Encoding.UTF8)
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [معلومة] في {methodName}")
                    writer.WriteLine($"  الرسالة: {message}")
                    writer.WriteLine(New String("-"c, 80))
                End Using
            End SyncLock
        Catch
        End Try
    End Sub

    ''' <summary>🌟 M-05: قناة تحذيرات مستقلة — تنبيهات دقة البيانات لا تختلط بأخطاء التشغيل
    ''' فيبقى ErrorLog نظيفاً: أي سطر فيه يعني "مشكلة فعلية تحتاج انتباه"</summary>
    Public Sub LogWarn(methodName As String, message As String)
        Try
            Dim logPath As String = GetLogPath()
            Directory.CreateDirectory(logPath)
            Dim logFile As String = Path.Combine(logPath, $"WarnLog_{DateTime.Now:yyyy-MM}.txt")

            SyncLock _logLock
                Using writer As New StreamWriter(logFile, append:=True, encoding:=Encoding.UTF8)
                    writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [تحذير] في {methodName}")
                    writer.WriteLine($"  الرسالة: {message}")
                    writer.WriteLine(New String("-"c, 80))
                End Using
            End SyncLock
        Catch
            ' تجاهل أخطاء كتابة السجل نفسها
        End Try
    End Sub

    ''' <summary>حذف السجلات القديمة (أكبر من 3 أشهر)</summary>
    Private Sub CleanOldLogs()
        Try
            Dim logPath As String = GetLogPath()
            If Not Directory.Exists(logPath) Then Return

            Dim cutoffDate As DateTime = DateTime.Now.AddMonths(-3)
            Dim files As String() = Directory.GetFiles(logPath, "*.txt")

            For Each filePath As String In files
                Dim fileInfo As New FileInfo(filePath)
                If fileInfo.LastWriteTime < cutoffDate Then
                    Try
                        fileInfo.Delete()
                    Catch
                    End Try
                End If
            Next
        Catch ex As Exception
            LogError("CleanOldLogs", ex)
        End Try
    End Sub

#End Region

#Region "Backup, Maintenance, and Utilities"

    ''' <summary>إنشاء نسخة احتياطية لقاعدة البيانات</summary>
    Public Function BackupDatabase(backupPath As String) As Boolean
        Try
            If Not _isInitialized Then
                If Not Initialize(False) Then Return False
            End If

            If String.IsNullOrEmpty(_dbPath) OrElse Not File.Exists(_dbPath) Then
                _lastError = "ملف قاعدة البيانات غير موجود"
                LogError("BackupDatabase", New FileNotFoundException(_lastError))
                Return False
            End If

            Dim backupDir As String = Path.GetDirectoryName(backupPath)
            If Not String.IsNullOrEmpty(backupDir) AndAlso Not Directory.Exists(backupDir) Then
                Directory.CreateDirectory(backupDir)
            End If

            ' تفريغ الـ WAL لضمان نسخ كل البيانات للملف الرئيسي
            ExecuteNonQuery("PRAGMA wal_checkpoint(TRUNCATE);")

            Using source As SQLiteConnection = GetConnection()
                Using dest As New SQLiteConnection($"Data Source={backupPath};Version=3;")
                    dest.Open()
                    source.BackupDatabase(dest, "main", "main", -1, Nothing, 0)
                End Using
            End Using

            LogInfo("BackupDatabase", $"تم إنشاء نسخة احتياطية في: {backupPath}")
            Return True

        Catch ex As Exception
            _lastError = ex.Message
            LogError("BackupDatabase", ex)
            Return False
        End Try
    End Function

    ''' <summary>تنظيف قاعدة البيانات وإزالة المساحة الزائدة (VACUUM)</summary>
    Public Sub VacuumDatabase()
        Try
            ExecuteNonQuery("VACUUM")
            LogInfo("VacuumDatabase", "تم تنفيذ VACUUM بنجاح")
        Catch ex As Exception
            LogError("VacuumDatabase", ex)
            Throw
        End Try
    End Sub

    ''' <summary>التحقق من سلامة قاعدة البيانات (PRAGMA integrity_check)</summary>
    Public Function ValidateDatabase() As Boolean
        Try
            Dim result As Object = ExecuteScalar("PRAGMA integrity_check")
            Dim isOk As Boolean = (result IsNot Nothing AndAlso
                                   result.ToString().Equals("ok", StringComparison.OrdinalIgnoreCase))
            If Not isOk Then
                LogError("ValidateDatabase",
                         New Exception($"فشل فحص السلامة: {result}"))
            End If
            Return isOk
        Catch ex As Exception
            _lastError = ex.Message
            LogError("ValidateDatabase", ex)
            Return False
        End Try
    End Function

    ''' <summary>الحصول على حجم قاعدة البيانات (بالبايت)</summary>
    Public Function GetDatabaseSize() As Long
        Try
            If File.Exists(_dbPath) Then
                Return New FileInfo(_dbPath).Length
            End If
        Catch ex As Exception
            LogError("GetDatabaseSize", ex)
        End Try
        Return 0
    End Function

    ''' <summary>الحصول على حجم قاعدة البيانات (نص منسق)</summary>
    Public Function GetDatabaseSizeFormatted() As String
        Dim size As Long = GetDatabaseSize()
        Select Case size
            Case Is >= 1073741824L : Return $"{size / 1073741824.0:F2} GB"
            Case Is >= 1048576L : Return $"{size / 1048576.0:F2} MB"
            Case Is >= 1024L : Return $"{size / 1024.0:F2} KB"
            Case Is > 0 : Return $"{size} B"
            Case Else : Return "0 B"
        End Select
    End Function

    ''' <summary>إعادة تعيين حالة الوحدة</summary>
    Public Sub Reset()
        _dbPath = ""
        _connectionString = ""
        _isInitialized = False
        _lastError = ""
    End Sub

    ' ═══════════════ 🔴 H-02: النسخ التلقائية المشفّرة (.adbak) ═══════════════
    ' كانت النسخ التلقائية ملفات .db نص صريح — نسخ مجلد AutoBackups = أخذ كل البيانات الحساسة.
    ' الصيغة الجديدة: [ELAB][1][IV 16][AES-256-CBC] بمفتاح مشتق من هوية الجهاز —
    '   * الملف المنقول لجهاز آخر (بريد/فلاشة) لا يُفتح إطلاقاً — الحماية الأساسية ضد السرقة الخارجية
    '   * على نفس الجهاز أي حساب يستطيع فكها نظرياً (نفس نطاق حماية DPAPI-LocalMachine) —
    '     الحل الكامل ضد التحليل المتعمد محلياً يبقى SQLCipher/EFS (خارطة الطريق متوسطة المدى)
    Private ReadOnly AutoBakMagic As Byte() = {&H45, &H4C, &H41, &H42}   ' "ELAB"
    Private Const AutoBakVersion As Byte = 1
    Private Const AutoBakIvSize As Integer = 16
    Private Const AutoBakChunk As Integer = 1048576

    ''' <summary>مفتاح النسخ التلقائية — مشتق من هوية الجهاز (لا يُخزن بالملف ولا بالإعدادات)</summary>
    Private Function GetAutoBackupKey() As Byte()
        Using sha As SHA256 = SHA256.Create()
            Dim raw As String = Environment.MachineName & "|ELNADY-AUTOBAK-v1|PlayerManagementSystem"
            Return sha.ComputeHash(Encoding.UTF8.GetBytes(raw))
        End Using
    End Function

    ''' <summary>هل الملف نسخة تلقائية محمية بصيغة ELAB؟</summary>
    Public Function IsProtectedAutoBackup(filePath As String) As Boolean
        Try
            If Not File.Exists(filePath) Then Return False
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                If fs.Length < AutoBakMagic.Length + 1 + AutoBakIvSize Then Return False
                Dim head(AutoBakMagic.Length - 1) As Byte
                If fs.Read(head, 0, head.Length) <> head.Length Then Return False
                For i As Integer = 0 To AutoBakMagic.Length - 1
                    If head(i) <> AutoBakMagic(i) Then Return False
                Next
                Return True
            End Using
        Catch
            Return False
        End Try
    End Function

    ''' <summary>تشفير نسخة تلقائية: قاعدة SQLite سليمة → ملف .adbak محمي</summary>
    Public Function WriteProtectedAutoBackup(plainSourceDb As String, destPath As String) As Boolean
        Try
            Dim iv As Byte() = New Byte(AutoBakIvSize - 1) {}
            Using rng As RandomNumberGenerator = RandomNumberGenerator.Create()
                rng.GetBytes(iv)
            End Using

            Using fsOut As New FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None)
                fsOut.Write(AutoBakMagic, 0, AutoBakMagic.Length)
                fsOut.WriteByte(AutoBakVersion)
                fsOut.Write(iv, 0, iv.Length)

                Using aes As Aes = Aes.Create()
                    aes.KeySize = 256
                    aes.Key = GetAutoBackupKey()
                    aes.IV = iv
                    aes.Mode = CipherMode.CBC
                    aes.Padding = PaddingMode.PKCS7
                    Using enc As ICryptoTransform = aes.CreateEncryptor()
                        Using cs As New CryptoStream(fsOut, enc, CryptoStreamMode.Write)
                            Using fsIn As New FileStream(plainSourceDb, FileMode.Open, FileAccess.Read, FileShare.Read)
                                Dim buffer(AutoBakChunk - 1) As Byte
                                Dim read As Integer
                                Do
                                    read = fsIn.Read(buffer, 0, buffer.Length)
                                    If read > 0 Then cs.Write(buffer, 0, read)
                                Loop While read > 0
                            End Using
                            cs.FlushFinalBlock()
                        End Using
                    End Using
                End Using
            End Using
            Return True
        Catch ex As Exception
            LogError("WriteProtectedAutoBackup", ex)
            Try
                If File.Exists(destPath) Then File.Delete(destPath)
            Catch
            End Try
            Return False
        End Try
    End Function

    ''' <summary>فك حماية نسخة تلقائية إلى ملف SQLite نصي مؤقت — المتصل يتحقق من سلامته بعدها (نمط M-06)</summary>
    Public Function UnprotectAutoBackup(encryptedPath As String, plainDestPath As String) As Boolean
        Try
            Using fsIn As New FileStream(encryptedPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim head(AutoBakMagic.Length - 1) As Byte
                If fsIn.Read(head, 0, head.Length) <> head.Length Then Return False
                For i As Integer = 0 To AutoBakMagic.Length - 1
                    If head(i) <> AutoBakMagic(i) Then Return False
                Next
                Dim version As Integer = fsIn.ReadByte()
                If version <> AutoBakVersion Then Return False

                Dim iv(AutoBakIvSize - 1) As Byte
                If fsIn.Read(iv, 0, iv.Length) <> iv.Length Then Return False

                Using aes As Aes = Aes.Create()
                    aes.KeySize = 256
                    aes.Key = GetAutoBackupKey()
                    aes.IV = iv
                    aes.Mode = CipherMode.CBC
                    aes.Padding = PaddingMode.PKCS7
                    Using dec As ICryptoTransform = aes.CreateDecryptor()
                        Using cs As New CryptoStream(fsIn, dec, CryptoStreamMode.Read)
                            Using fsOut As New FileStream(plainDestPath, FileMode.Create, FileAccess.Write, FileShare.None)
                                Dim buffer(AutoBakChunk - 1) As Byte
                                Dim read As Integer
                                Do
                                    read = cs.Read(buffer, 0, buffer.Length)
                                    If read > 0 Then fsOut.Write(buffer, 0, read)
                                Loop While read > 0
                            End Using
                        End Using
                    End Using
                End Using
            End Using
            Return True
        Catch ex As Exception
            LogError("UnprotectAutoBackup", ex)
            Try
                If File.Exists(plainDestPath) Then File.Delete(plainDestPath)
            Catch
            End Try
            Return False
        End Try
    End Function

    ''' <summary>
    ''' إنشاء نسخة احتياطية تلقائية يومياً في مجلد AutoBackups وحذف النسخ القديمة
    ''' 🔴 H-02: النسخة تُكتب مشفّرة (.adbak) — كانت .db نص صريح يُفتح من أي حساب أو جهاز
    ''' </summary>
    Public Sub PerformAutoBackup(Optional keepDays As Integer = 14)
        Try
            Dim backupDir As String = Path.Combine(GetAppDataPath(), "AutoBackups")
            If Not Directory.Exists(backupDir) Then
                Directory.CreateDirectory(backupDir)
            End If

            ' اسم ملف النسخة الاحتياطية لليوم (الصيغة المحمية الجديدة)
            Dim todayFile As String = Path.Combine(backupDir, $"AutoBackup_{DateTime.Today:yyyy-MM-dd}.adbak")

            ' إذا لم يتم أخذ نسخة اليوم: نسخة SQLite سليمة → تشفير → إزالة الوسيطة النصية
            If Not File.Exists(todayFile) Then
                Dim tempDb As String = Path.Combine(Path.GetTempPath(),
                    $"AutoBak_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.tmpdb")
                Try
                    If BackupDatabase(tempDb) Then
                        If WriteProtectedAutoBackup(tempDb, todayFile) Then
                            LogInfo("PerformAutoBackup", $"🔴 H-02: أُنشئت النسخة التلقائية المشفّرة لليوم: {Path.GetFileName(todayFile)}")
                        Else
                            LogWarn("PerformAutoBackup", "فشل تشفير النسخة التلقائية — سيُعاد المحاولة لاحقاً")
                        End If
                    End If
                Finally
                    Try
                        If File.Exists(tempDb) Then File.Delete(tempDb)
                    Catch
                    End Try
                End Try
            End If

            ' تنظيف النسخ القديمة التي تجاوزت عدد الأيام (الجديدة .adbak والقديمة .db معاً)
            Dim cutoffDate As DateTime = DateTime.Today.AddDays(-keepDays)
            Dim backupFiles As String() = Directory.GetFiles(backupDir, "AutoBackup_*.adbak")
            backupFiles = backupFiles.Concat(Directory.GetFiles(backupDir, "AutoBackup_*.db")).ToArray()

            For Each filePath As String In backupFiles
                Dim fileInfo As New FileInfo(filePath)
                If fileInfo.CreationTime < cutoffDate OrElse fileInfo.LastWriteTime < cutoffDate Then
                    Try
                        fileInfo.Delete()
                        LogInfo("PerformAutoBackup", $"تم حذف نسخة قديمة: {fileInfo.Name}")
                    Catch
                    End Try
                End If
            Next

        Catch ex As Exception
            LogError("PerformAutoBackup", ex)
        End Try
    End Sub

#End Region

End Module
