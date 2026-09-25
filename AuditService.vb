Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Data
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading.Tasks

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — خدمة سجل التدقيق (AuditLog)
' يوثّق تلقائياً: تسجيل الدخول/الخروج، إضافة/تعديل/حذف اللاعبين والدفعات
' والمصروفات والمستخدمين، النسخ الاحتياطي والاستعادة.
' ── قواعد التصميم الذهبية:
'    * التسجيل لا يُسقط العملية أبداً: أي فشل يُلتهم بصمت مع LogError فقط
'    * لا يُسجل أبداً: كلمات المرور، الصور، أو نص الملاحظات الكامل (خصوصية)
'    * الجدول يُنشأ تلقائياً عند أول استخدام
' ═══════════════════════════════════════════════════════════════════
Public Class AuditService

    Public Const Table_AuditLog As String = "AuditLog"

    ' ── أنواع الأحداث المعروفة (توحيد النصوص حتى تبقى الفلاتر دقيقة) ──
    Public Const Act_LoginSuccess As String = "تسجيل دخول"
    Public Const Act_LoginFailed As String = "محاولة دخول فاشلة"
    Public Const Act_Logout As String = "تسجيل خروج"
    Public Const Act_AppClosed As String = "إغلاق البرنامج"
    Public Const Act_SessionLocked As String = "قفل الجلسة (خمول)"
    Public Const Act_SessionUnlocked As String = "فتح الجلسة"
    Public Const Act_PlayerAdd As String = "إضافة لاعب"
    Public Const Act_PlayerUpdate As String = "تعديل لاعب"
    Public Const Act_PlayerArchive As String = "أرشفة لاعب"
    Public Const Act_PlayerRestore As String = "استرجاع لاعب من الأرشيف"
    Public Const Act_PaymentAdd As String = "إضافة دفعة"
    Public Const Act_PaymentUpdate As String = "تعديل دفعة"
    Public Const Act_PaymentDelete As String = "حذف دفعة"
    Public Const Act_ExpenseAdd As String = "إضافة مصروف"
    Public Const Act_ExpenseUpdate As String = "تعديل مصروف"
    Public Const Act_ExpenseDelete As String = "حذف مصروف"
    Public Const Act_UserAdd As String = "إضافة مستخدم"
    Public Const Act_UserDelete As String = "حذف مستخدم"
    Public Const Act_UserPerms As String = "تعديل صلاحيات"
    Public Const Act_PasswordChanged As String = "تغيير كلمة مرور"
    Public Const Act_Backup As String = "نسخ احتياطي"
    Public Const Act_BackupEncrypted As String = "نسخ احتياطي مشفّر"
    Public Const Act_Restore As String = "استعادة نسخة احتياطية"
    Public Const Act_RestoreEncrypted As String = "استعادة نسخة مشفّرة"
    Public Const Act_ReportExport As String = "تصدير تقرير"

    ' 🌟 إصلاح: كان العَلَم لكل نسخة (وهناك New AuditService() مع كل عملية) —
    ' لذلك كانت CREATE TABLE/INDEX تُنفَّذ مع كل تسجيل. الآن Shared ومحمي بقفل
    Private Shared _tableChecked As Boolean = False
    Private Shared ReadOnly _tableCheckLock As New Object()

    Public Sub EnsureAuditLogTableExists()
        If _tableChecked Then Return
        SyncLock _tableCheckLock
            If _tableChecked Then Return
        Try
            DatabaseModule.ExecuteNonQuery(
                "CREATE TABLE IF NOT EXISTS " & Table_AuditLog & " (" & vbCrLf &
                "  [Id]        INTEGER PRIMARY KEY AUTOINCREMENT," & vbCrLf &
                "  [Timestamp] TEXT    NOT NULL," & vbCrLf &
                "  [Username]  TEXT," & vbCrLf &
                "  [Action]    TEXT    NOT NULL," & vbCrLf &
                "  [EntityType] TEXT," & vbCrLf &
                "  [EntityId]  TEXT," & vbCrLf &
                "  [Details]   TEXT" & vbCrLf &
                ")")

            Dim existing As String = "|"
            Dim info As System.Data.DataTable = DatabaseModule.ExecuteQuery("PRAGMA table_info(" & Table_AuditLog & ")")
            For Each r As System.Data.DataRow In info.Rows
                existing &= Convert.ToString(r("name")).ToLowerInvariant() & "|"
            Next
            For Each colName As String In New String() {"Timestamp", "Username", "Action", "EntityType", "EntityId", "Details", "EntryHash"}
                If Not existing.Contains("|" & colName.ToLowerInvariant() & "|") Then
                    DatabaseModule.ExecuteNonQuery("ALTER TABLE " & Table_AuditLog & " ADD COLUMN [" & colName & "] TEXT")
                End If
            Next

            DatabaseModule.ExecuteNonQuery(
                "CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON " & Table_AuditLog & " ([Timestamp] DESC)")
            _tableChecked = True
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.EnsureAuditLogTableExists", ex)
        End Try
        End SyncLock
    End Sub
    ''' <summary>تسجيل حدث بشكل متزامن (للتحويلات غير الـ Async)</summary>
    ''' <remarks>آمن تماماً: لا يرمي استثناءات، والفشل لا يوقف البرنامج</remarks>
    Public Sub Log(action As String, Optional entityType As String = "",
                   Optional entityId As String = "", Optional details As String = "")
        Try
            If String.IsNullOrWhiteSpace(action) Then Return
            EnsureAuditLogTableExists()

            Dim user As String = If(String.IsNullOrWhiteSpace(UserSession.CurrentUsername), "غير معروف", UserSession.CurrentUsername)

            ' 🌟 L-12: سلسلة هاش — كل سطر يحمل SHA256 لـ(هاش السطر السابق + محتواه) لكشف أي تعديل خارجي صامت.
            ' فشل حساب الهاش لا يُسقط التسجيل أبداً (القاعدة الذهبية): السطر يُخزن بلا هاش ويكشفه فحص السلامة
            Dim tsText As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
            Dim entryHash As Object = DBNull.Value
            Try
                Dim prevHash As String = ""
                Dim prevObj As Object = DatabaseModule.ExecuteScalar(
                    $"SELECT [EntryHash] FROM {Table_AuditLog} ORDER BY [Id] DESC LIMIT 1")
                If prevObj IsNot Nothing AndAlso Not IsDBNull(prevObj) Then prevHash = Convert.ToString(prevObj)
                entryHash = CObj(ComputeEntryHash(prevHash, tsText, user, action, entityType, entityId, details))
            Catch
            End Try

            DatabaseModule.ExecuteNonQuery(
                $"INSERT INTO {Table_AuditLog} ([Timestamp], [Username], [Action], [EntityType], [EntityId], [Details], [EntryHash]) " &
                "VALUES (@ts, @u, @a, @et, @ei, @d, @h)",
                New SQLiteParameter("@ts", tsText),
                New SQLiteParameter("@u", user),
                New SQLiteParameter("@a", action),
                New SQLiteParameter("@et", If(String.IsNullOrWhiteSpace(entityType), DBNull.Value, CObj(entityType))),
                New SQLiteParameter("@ei", If(String.IsNullOrWhiteSpace(entityId), DBNull.Value, CObj(entityId))),
                New SQLiteParameter("@d", If(String.IsNullOrWhiteSpace(details), DBNull.Value, CObj(details))),
                New SQLiteParameter("@h", entryHash))
        Catch ex As Exception
            ' سجل التدقيق لا يُسقط العملية الحالية أبداً — يُكتفى بتسجيل الفشل في اللوج
            DatabaseModule.LogError("AuditService.Log", ex)
        End Try
    End Sub

    ''' <summary>تسجيل حدث بشكل غير متزامن (للتحويلات الـ Async في الخدمات)</summary>
    Public Async Function LogAsync(action As String, Optional entityType As String = "",
                                   Optional entityId As String = "", Optional details As String = "") As Task
        Try
            If String.IsNullOrWhiteSpace(action) Then Return
            EnsureAuditLogTableExists()

            Dim user As String = If(String.IsNullOrWhiteSpace(UserSession.CurrentUsername), "غير معروف", UserSession.CurrentUsername)

            ' 🌟 L-12: سلسلة هاش — نفس منطق المسار المتزامن (انظر التعليق أعلاه في Log)
            Dim tsText As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
            Dim entryHash As Object = DBNull.Value
            Try
                Dim prevHash As String = ""
                Dim prevObj As Object = Await DatabaseModule.ExecuteScalarAsync(
                    $"SELECT [EntryHash] FROM {Table_AuditLog} ORDER BY [Id] DESC LIMIT 1")
                If prevObj IsNot Nothing AndAlso Not IsDBNull(prevObj) Then prevHash = Convert.ToString(prevObj)
                entryHash = CObj(ComputeEntryHash(prevHash, tsText, user, action, entityType, entityId, details))
            Catch
            End Try

            Await DatabaseModule.ExecuteNonQueryAsync(
                $"INSERT INTO {Table_AuditLog} ([Timestamp], [Username], [Action], [EntityType], [EntityId], [Details], [EntryHash]) " &
                "VALUES (@ts, @u, @a, @et, @ei, @d, @h)",
                New SQLiteParameter("@ts", tsText),
                New SQLiteParameter("@u", user),
                New SQLiteParameter("@a", action),
                New SQLiteParameter("@et", If(String.IsNullOrWhiteSpace(entityType), DBNull.Value, CObj(entityType))),
                New SQLiteParameter("@ei", If(String.IsNullOrWhiteSpace(entityId), DBNull.Value, CObj(entityId))),
                New SQLiteParameter("@d", If(String.IsNullOrWhiteSpace(details), DBNull.Value, CObj(details))),
                New SQLiteParameter("@h", entryHash))
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.LogAsync", ex)
        End Try
    End Function

    ''' <summary>استعلام السجل مع فلاتر اختيارية — يعيد Nothing عند الفشل</summary>
    ''' <param name="fromDate">تاريخ بداية (yyyy-MM-dd) أو "" للاستغناء</param>
    ''' <param name="toDate">تاريخ نهاية (yyyy-MM-dd) أو ""</param>
    ''' <param name="username">اسم مستخدم للفلترة أو ""</param>
    ''' <param name="action">نوع حدث للفلترة أو ""</param>
    ''' <param name="searchText">بحث حر في التفاصيل/المعرّف أو ""</param>
    ''' <param name="limit">أقصى عدد صفوف (حماية من تعليق الواجهة بالسجلات الضخمة)</param>
    Public Async Function QueryAsync(fromDate As String, toDate As String,
                                     username As String, action As String,
                                     searchText As String, limit As Integer) As Task(Of DataTable)
        Try
            EnsureAuditLogTableExists()

            Dim conditions As New List(Of String)
            Dim parameters As New List(Of SQLiteParameter)

            If Not String.IsNullOrWhiteSpace(fromDate) Then
                conditions.Add("[Timestamp] >= @f")
                parameters.Add(New SQLiteParameter("@f", fromDate.Trim() & " 00:00:00"))
            End If
            If Not String.IsNullOrWhiteSpace(toDate) Then
                conditions.Add("[Timestamp] <= @t")
                parameters.Add(New SQLiteParameter("@t", toDate.Trim() & " 23:59:59"))
            End If
            If Not String.IsNullOrWhiteSpace(username) Then
                conditions.Add("[Username] = @u")
                parameters.Add(New SQLiteParameter("@u", username.Trim()))
            End If
            If Not String.IsNullOrWhiteSpace(action) Then
                conditions.Add("[Action] = @a")
                parameters.Add(New SQLiteParameter("@a", action.Trim()))
            End If
            If Not String.IsNullOrWhiteSpace(searchText) Then
                conditions.Add("([Details] LIKE @s OR [EntityId] LIKE @s OR [EntityType] LIKE @s)")
                parameters.Add(New SQLiteParameter("@s", "%" & searchText.Trim() & "%"))
            End If

            Dim cappedLimit As Integer = limit
            If cappedLimit <= 0 Then cappedLimit = 2000

            Dim sql As String =
                $"SELECT [Id], [Timestamp], [Username], [Action], [EntityType], [EntityId], [Details] " &
                $"FROM {Table_AuditLog}"
            If conditions.Count > 0 Then
                sql &= " WHERE " & String.Join(" AND ", conditions)
            End If
            sql &= $" ORDER BY [Id] DESC LIMIT {cappedLimit}"

            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, parameters.ToArray())
            Return dt
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.QueryAsync", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>أسماء المستخدمين الموجودين في السجل (لفلتر العارض)</summary>
    Public Async Function GetDistinctUsernamesAsync() As Task(Of List(Of String))
        Dim list As New List(Of String)
        Try
            EnsureAuditLogTableExists()
            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(
                $"SELECT DISTINCT [Username] FROM {Table_AuditLog} WHERE [Username] IS NOT NULL ORDER BY [Username]")
            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    If Not IsDBNull(row("Username")) Then list.Add(Convert.ToString(row("Username")))
                Next
            End If
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.GetDistinctUsernamesAsync", ex)
        End Try
        Return list
    End Function

    ''' <summary>أنواع الأحداث الموجودة في السجل (لفلتر العارض)</summary>
    Public Async Function GetDistinctActionsAsync() As Task(Of List(Of String))
        Dim list As New List(Of String)
        Try
            EnsureAuditLogTableExists()
            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(
                $"SELECT DISTINCT [Action] FROM {Table_AuditLog} ORDER BY [Action]")
            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    If Not IsDBNull(row("Action")) Then list.Add(Convert.ToString(row("Action")))
                Next
            End If
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.GetDistinctActionsAsync", ex)
        End Try
        Return list
    End Function

    ''' <summary>حذف السجلات الأقدم من عدد أيام محدد — يعيد عدد الصفوف المحذوفة (-1 عند الفشل)</summary>
    Public Async Function PurgeOlderThanAsync(keepDays As Integer) As Task(Of Integer)
        Try
            If keepDays <= 0 Then Return 0
            EnsureAuditLogTableExists()
            Dim cutoff As String = DateTime.Now.AddDays(-keepDays).ToString("yyyy-MM-dd 00:00:00", Globalization.CultureInfo.InvariantCulture)
            Return Await DatabaseModule.ExecuteNonQueryAsync(
                $"DELETE FROM {Table_AuditLog} WHERE [Timestamp] < @c",
                New SQLiteParameter("@c", cutoff))
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.PurgeOlderThanAsync", ex)
            Return -1
        End Try
    End Function

    ''' <summary>عدد إجمالي للسجلات (تُظهره شاشة العارض)</summary>
    Public Async Function GetTotalCountAsync() As Task(Of Long)
        Try
            EnsureAuditLogTableExists()
            Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(
                $"SELECT COUNT(*) FROM {Table_AuditLog}")
            If result Is Nothing OrElse IsDBNull(result) Then Return 0
            Return Convert.ToInt64(result)
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.GetTotalCountAsync", ex)
            Return -1
        End Try
    End Function

    ''' <summary>حساب هاش سلسلة سطر التدقيق (L-12): SHA256 لـ(هاش السابق | الطابع | المستخدم | الحدث | النوع | المعرف | التفاصيل)</summary>
    Private Shared Function ComputeEntryHash(prevHash As String, timestamp As String,
                                             username As String, action As String,
                                             entityType As String, entityId As String,
                                             details As String) As String
        Dim canonical As String = prevHash & "|" & timestamp & "|" & username & "|" &
                                  action & "|" & entityType & "|" & entityId & "|" & details
        Using sha As SHA256 = SHA256.Create()
            Dim bytes As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(canonical))
            Dim sb As New StringBuilder(bytes.Length * 2)
            For Each b As Byte In bytes
                sb.Append(b.ToString("x2", Globalization.CultureInfo.InvariantCulture))
            Next
            Return sb.ToString()
        End Using
    End Function

    ''' <summary>فحص سلامة سلسلة التدقيق (L-12) — يعيد (سليم؟، رسالة تفصيلية)</summary>
    ''' <remarks>يتجول على السطور بترتيب Id ويعيد حساب الهاش ويقارنه بالمخزن، وأول سطر لا يطابق يُبلّغ عنه.
    ''' السطور بلا هاش (قديمة سابقة للتحديث أو فشل حساب) تُعد بداية سلسلة جديدة ولا تُعد خللاً.
    ''' الفحص للقراءة فقط ولا يعدل أي بيانات.</remarks>
    Public Async Function VerifyAuditChainAsync() As Task(Of Tuple(Of Boolean, String))
        Try
            EnsureAuditLogTableExists()
            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(
                $"SELECT [Id], [Timestamp], [Username], [Action], [EntityType], [EntityId], [Details], [EntryHash] " &
                $"FROM {Table_AuditLog} ORDER BY [Id]")
            If dt Is Nothing Then
                Return Tuple.Create(False, "تعذر قراءة السجل من قاعدة البيانات — راجع ملف اللوج")
            End If

            Dim prevHash As String = ""
            Dim verified As Integer = 0
            Dim legacy As Integer = 0
            For Each row As DataRow In dt.Rows
                Dim idText As String = Convert.ToString(row("Id"))
                Dim stored As String = ""
                If row("EntryHash") IsNot Nothing AndAlso Not IsDBNull(row("EntryHash")) Then
                    stored = Convert.ToString(row("EntryHash"))
                End If
                If String.IsNullOrWhiteSpace(stored) Then
                    legacy += 1
                    prevHash = ""
                    Continue For
                End If
                Dim recomputed As String = ComputeEntryHash(prevHash,
                    Convert.ToString(row("Timestamp")),
                    If(IsDBNull(row("Username")), "", Convert.ToString(row("Username"))),
                    If(IsDBNull(row("Action")), "", Convert.ToString(row("Action"))),
                    If(IsDBNull(row("EntityType")), "", Convert.ToString(row("EntityType"))),
                    If(IsDBNull(row("EntityId")), "", Convert.ToString(row("EntityId"))),
                    If(IsDBNull(row("Details")), "", Convert.ToString(row("Details"))))
                If Not String.Equals(recomputed, stored, StringComparison.OrdinalIgnoreCase) Then
                    Return Tuple.Create(False,
                        $"أول خلل عند السطر رقم {idText} — محتواه أو ما قبله تعرّض لتعديل خارج السجل." & vbCrLf &
                        $"سطور متحققة: {verified} / سطور قديمة بلا هاش: {legacy}")
                End If
                prevHash = stored
                verified += 1
            Next
            Dim okMsg As String = $"فُحص {dt.Rows.Count} سطراً: {verified} ضمن سلسلة سليمة / {legacy} قديمة بلا هاش (سابقة للتحديث)."
            Return Tuple.Create(True, okMsg)
        Catch ex As Exception
            DatabaseModule.LogError("AuditService.VerifyAuditChainAsync", ex)
            Return Tuple.Create(False, "فشل الفحص — راجع ملف اللوج")
        End Try
    End Function

End Class
