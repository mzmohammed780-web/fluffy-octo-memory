Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Data ' 🌟 تمت الإضافة لتعريف DataTable و DataRow
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class UserService

    ''' <summary>نتيجة التحقق من بيانات الدخول — بيانات خام بدون أي تعديل لجلسة المستخدم</summary>
    Private Class LoginResult
        Public Success As Boolean = False
        Public RealUsername As String = ""
        Public CanEdit As Boolean = False
        Public CanDelete As Boolean = False
        Public MustChangePassword As Boolean = False
    End Class

    ''' <summary>التحقق من بيانات الدخول فقط بدون تعديل الجلسة (للاستخدام الداخلي وللتغييرات الإدارية)</summary>
    ''' <remarks>🌟 إصلاح أمني: كانت ValidateLoginAsync تعدّل جلسة المستخدم كأثر جانبي، فحين يغيّر الأدمن
    ''' كلمة مرور مستخدم آخر كانت صلاحيات الأدمن نفسه تتبدّل لصلاحيات ذلك المستخدم!</remarks>
    Private Async Function ValidateCredentialsAsync(username As String, password As String) As Task(Of LoginResult)
        Dim result As New LoginResult()
        Try
            If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then Return result

            ' 🌟 L-11 — قرار موثق: الاقتطاع مقصود ومتسق في كل مسارات كلمات المرور (تخزين/تحقق/تغيير).
            ' إزالته يتطلب ترحيلاً واعياً لأن الهاشات المحفوظة أُنشئت بعد الاقتطاع؛
            ' الثمن المقبول: فضاء الأحرف الفعلي ينقص عند طرفي الكلمة فقط — التفصيل في اقرأني_أولاً
            Dim cleanUser As String = username.Trim()
            Dim cleanPass As String = password.Trim()

            ' 1. فحص المستخدمين من قاعدة البيانات (بما فيهم الأدمن)
            ' 🌟 إصلاح: قراءة MustChangePassword من القاعدة — كانت الجلسة لا تستقبلها أبداً
            ' فآلية الإجبار على تغيير كلمة المرور المؤقتة (مثل 123456) ما كانت تعمل أبداً
            Dim sql As String = $"SELECT {AppConstants.Col_Username}, {AppConstants.Col_PasswordHash}, {AppConstants.Col_CanEdit}, {AppConstants.Col_CanDelete}, {AppConstants.Col_MustChangePassword} " &
                                $"FROM {AppConstants.Table_Users} WHERE LOWER({AppConstants.Col_Username}) = LOWER(@u)"

            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, New SQLiteParameter("@u", cleanUser))

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim storedHash As String = UtilityModule.SafeString(dt.Rows(0)(AppConstants.Col_PasswordHash))

                ' التحقق من كلمة المرور باستخدام التشفير
                If UtilityModule.VerifyPassword(cleanPass, storedHash) Then
                    result.Success = True
                    result.RealUsername = UtilityModule.SafeString(dt.Rows(0)(AppConstants.Col_Username))
                    result.CanEdit = (Convert.ToInt32(dt.Rows(0)(AppConstants.Col_CanEdit)) = 1)
                    result.CanDelete = (Convert.ToInt32(dt.Rows(0)(AppConstants.Col_CanDelete)) = 1)
                    ' 🌟 قراءة آمنة (NULL يعامل 0) — العمود أحدث من غيره فقد يحتوي NULL ببيانات قديمة
                    Dim mcpValue As Object = dt.Rows(0)(AppConstants.Col_MustChangePassword)
                    Dim mcp As Integer = If(IsDBNull(mcpValue), 0, Convert.ToInt32(mcpValue))
                    result.MustChangePassword = (mcp = 1)

                    ' 🌟 ترقية صيغة التشفير القديمة إلى الصيغة المحصّنة v2 بشفافية عند أول دخول ناجح
                    If UtilityModule.GetHashFormatVersion(storedHash) < 2 Then
                        Await UpgradePasswordHashAsync(cleanUser, cleanPass)
                    End If
                End If
            End If

            Return result
        Catch ex As Exception
            DatabaseModule.LogError("UserService.ValidateCredentialsAsync", ex)
            Return New LoginResult()
        End Try
    End Function

    ''' <summary>التحقق من صحة اسم المستخدم وكلمة المرور وتعبئة الجلسة عند النجاح (لشاشة تسجيل الدخول)</summary>
    Public Async Function ValidateLoginAsync(username As String, password As String) As Task(Of Boolean)
        Dim result As LoginResult = Await ValidateCredentialsAsync(username, password)
        If result.Success Then
            UserSession.CurrentUsername = result.RealUsername
            UserSession.CanEdit = result.CanEdit
            UserSession.CanDelete = result.CanDelete
            UserSession.MustChangePassword = result.MustChangePassword

            ' 🌟 [أولوية 4] توثيق الدخول الناجح في سجل التدقيق
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_LoginSuccess, "جلسة", result.RealUsername, "")
        Else
            ' 🌟 [أولوية 4] توثيق المحاولات الفاشلة (لا تُسجل كلمة المرور أبداً)
            Dim auditFail As New AuditService()
            Await auditFail.LogAsync(AuditService.Act_LoginFailed, "جلسة",
                If(String.IsNullOrWhiteSpace(username), "غير معروف", username.Trim()), "")
        End If
        Return result.Success
    End Function

    ''' <summary>ترقية صيغة تخزين كلمة المرور إلى v2 (تكرارات PBKDF2 أعلى) بعد دخول ناجح — شفافة تماماً للمستخدم</summary>
    Private Async Function UpgradePasswordHashAsync(username As String, plainPassword As String) As Task
        Try
            Dim newHash As String = UtilityModule.HashPassword(plainPassword)
            Dim sql As String = $"UPDATE {AppConstants.Table_Users} SET {AppConstants.Col_PasswordHash} = @p WHERE LOWER({AppConstants.Col_Username}) = LOWER(@u)"
            Await DatabaseModule.ExecuteNonQueryAsync(sql,
                New SQLiteParameter("@p", newHash),
                New SQLiteParameter("@u", username.Trim()))
        Catch ex As Exception
            DatabaseModule.LogError("UserService.UpgradePasswordHashAsync", ex)
        End Try
    End Function

    ''' <summary>إضافة مستخدم جديد بصلاحيات محددة</summary>
    Public Async Function AddUserAsync(username As String, password As String, canEdit As Boolean, canDelete As Boolean) As Task(Of Boolean)
        Try
            If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrWhiteSpace(password) Then Return False

            ' 🌟 حد أدنى لطول كلمة المرور عند إنشاء مستخدم جديد (موحّد على 8 مع شاشة تغيير كلمة المرور)
            If password.Trim().Length < 8 Then Return False

            Dim checkSql As String = $"SELECT COUNT(*) FROM {AppConstants.Table_Users} WHERE LOWER({AppConstants.Col_Username}) = LOWER(@u)"
            Dim checkResult As Object = Await DatabaseModule.ExecuteScalarAsync(checkSql, New SQLiteParameter("@u", username.Trim()))
            If Convert.ToInt32(checkResult) > 0 Then Return False

            Dim hash As String = UtilityModule.HashPassword(password.Trim())
            Dim sql As String = $"INSERT INTO {AppConstants.Table_Users} ({AppConstants.Col_Username}, {AppConstants.Col_PasswordHash}, {AppConstants.Col_CanEdit}, {AppConstants.Col_CanDelete}) VALUES (@u, @p, @e, @d)"
            Dim result As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql,
                New SQLiteParameter("@u", username.Trim()),
                New SQLiteParameter("@p", hash),
                New SQLiteParameter("@e", If(canEdit, 1, 0)),
                New SQLiteParameter("@d", If(canDelete, 1, 0)))

            ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
            If result > 0 Then
                Dim audit As New AuditService()
                Await audit.LogAsync(AuditService.Act_UserAdd, "مستخدم", username.Trim(),
                    $"تعديل: {If(canEdit, 1, 0)} — حذف: {If(canDelete, 1, 0)}")
            End If

            Return result > 0
        Catch ex As Exception
            DatabaseModule.LogError("UserService.AddUserAsync", ex)
            Return False
        End Try
    End Function

    ''' <summary>تغيير كلمة المرور لمستخدم معين</summary>
    Public Async Function ChangePasswordAsync(username As String, oldPassword As String, newPassword As String) As Task(Of Boolean)
        Try
            ' 🌟 إصلاح حرج: التحقق من كلمة المرور الجديدة أولاً — كانت كلمة مرور فارغة تُحفظ
            ' وتجعل الحساب مستحيل الدخول إليه نهائياً (VerifyPassword يفشل دائماً على hash فارغ)
            If String.IsNullOrWhiteSpace(newPassword) Then Return False
            If newPassword.Trim().Length < 8 Then Return False

            ' 🌟 منع "التغيير" إلى نفس كلمة المرور القديمة (يُبطل غرض الإجبار على التغيير)
            If String.Equals(newPassword.Trim(), If(oldPassword, "").Trim()) Then Return False

            ' 🌟 إصلاح أمني: التحقق من كلمة المرور القديمة الآن لا يمس جلسة المستخدم الحالية إطلاقاً
            ' 🌟 إصلاح ترجمي: كان المتغير اسمه result ويتعارض مع result As Integer أدناه — أعيدت تسميته cred
            Dim cred As LoginResult = Await ValidateCredentialsAsync(username, oldPassword)
            If Not cred.Success Then Return False

            Dim newHash As String = UtilityModule.HashPassword(newPassword.Trim())

            ' 🌟 إصلاح حرج: تصفير علم MustChangePassword بعد التغيير الناجح —
            ' كان يبقى مرفوعاً فيُجبر المستخدم على تغيير كلمة مروره عند كل تسجيل دخول للأبد
            Dim sql As String = $"UPDATE {AppConstants.Table_Users} SET {AppConstants.Col_PasswordHash} = @p, {AppConstants.Col_MustChangePassword} = 0 WHERE LOWER({AppConstants.Col_Username}) = LOWER(@u)"
            Dim result As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql,
                New SQLiteParameter("@p", newHash),
                New SQLiteParameter("@u", username.Trim()))

            ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق (بدون أي كلمات مرور)
            If result > 0 Then
                Dim audit As New AuditService()
                Await audit.LogAsync(AuditService.Act_PasswordChanged, "مستخدم", username.Trim(), "")
            End If

            Return result > 0
        Catch ex As Exception
            DatabaseModule.LogError("UserService.ChangePasswordAsync", ex)
            Return False
        End Try
    End Function

    ''' <summary>جلب كل المستخدمين لعرضهم في الجدول</summary>
    Public Async Function GetAllUsersAsync() As Task(Of DataTable)
        Dim sql As String = $"SELECT {AppConstants.Col_Username}, {AppConstants.Col_CanEdit}, {AppConstants.Col_CanDelete} FROM {AppConstants.Table_Users} ORDER BY {AppConstants.Col_Username}"
        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql)

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            dt.Columns.Add("EditBool", GetType(Boolean))
            dt.Columns.Add("DeleteBool", GetType(Boolean))

            For Each row As DataRow In dt.Rows
                row("EditBool") = (Convert.ToInt32(row(AppConstants.Col_CanEdit)) = 1)
                row("DeleteBool") = (Convert.ToInt32(row(AppConstants.Col_CanDelete)) = 1)
            Next

            dt.Columns.Remove(AppConstants.Col_CanEdit)
            dt.Columns.Remove(AppConstants.Col_CanDelete)

            dt.Columns("EditBool").ColumnName = AppConstants.Col_CanEdit
            dt.Columns("DeleteBool").ColumnName = AppConstants.Col_CanDelete
        End If

        Return dt
    End Function

    ''' <summary>تحديث صلاحيات مستخدم فردي</summary>
    Public Async Function UpdateUserPermissionsAsync(username As String, canEdit As Boolean, canDelete As Boolean) As Task(Of Integer)
        ' 🌟 حماية خلفية: صلاحيات المدير الرئيسي لا تُنقص أبداً — حتى لو استُدعيت من مكان آخر
        If username.Trim().ToLower() = "admin" Then
            Return Await UpdateUserPermissionsAdminProtected()
        End If
        Dim sql As String = $"UPDATE {AppConstants.Table_Users} SET {AppConstants.Col_CanEdit} = @e, {AppConstants.Col_CanDelete} = @d WHERE LOWER({AppConstants.Col_Username}) = LOWER(@u)"
        Dim params As SQLiteParameter() = {
            New SQLiteParameter("@e", If(canEdit, 1, 0)),
            New SQLiteParameter("@d", If(canDelete, 1, 0)),
            New SQLiteParameter("@u", username.Trim())
        }
        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, params)

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_UserPerms, "مستخدم", username.Trim(),
                $"تعديل: {If(canEdit, 1, 0)} — حذف: {If(canDelete, 1, 0)}")
        End If

        Return affected
    End Function

    ''' <summary>تحديث صلاحيات عدة مستخدمين</summary>
    Public Async Function UpdateMultipleUsersPermissions(usersList As List(Of Tuple(Of String, Boolean, Boolean))) As Task(Of Boolean)
        If usersList Is Nothing OrElse usersList.Count = 0 Then Return True
        Try
            Dim queries As New List(Of Tuple(Of String, SQLiteParameter()))
            For Each item In usersList
                Dim sql As String = $"UPDATE {AppConstants.Table_Users} SET {AppConstants.Col_CanEdit} = @e, {AppConstants.Col_CanDelete} = @d WHERE LOWER({AppConstants.Col_Username}) = LOWER(@u)"
                Dim params As SQLiteParameter() = {
                    New SQLiteParameter("@e", If(item.Item2, 1, 0)),
                    New SQLiteParameter("@d", If(item.Item3, 1, 0)),
                    New SQLiteParameter("@u", item.Item1.Trim())
                }
                queries.Add(Tuple.Create(sql, params))
            Next

            ' 🌟 تنفيذ المعاملة بشكل متزامن داخل Task.Run لتجنب Anti-pattern
            Dim committed As Boolean = Await Task.Run(Function() DatabaseModule.ExecuteTransaction(queries))

            ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
            If committed Then
                Dim audit As New AuditService()
                Await audit.LogAsync(AuditService.Act_UserPerms, "مستخدم", "",
                    $"تحديث صلاحيات {usersList.Count} مستخدماً دفعة واحدة")
            End If

            Return committed
        Catch ex As Exception
            DatabaseModule.LogError("UserService.UpdateMultipleUsersPermissions", ex)
            Return False
        End Try
    End Function

    ''' <summary>حذف مستخدم</summary>
    Public Async Function DeleteUserAsync(username As String) As Task(Of Integer)
        If username.Trim().ToLower() = "admin" Then Return -1

        Dim sql As String = $"DELETE FROM {AppConstants.Table_Users} WHERE LOWER({AppConstants.Col_Username}) = LOWER(@u)"
        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, New SQLiteParameter("@u", username.Trim()))

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_UserDelete, "مستخدم", username.Trim(), "")
        End If

        Return affected
    End Function
    ''' <summary>الأدمن يُحفظ دائماً بصلاحيات كاملة (1,1) — حماية قسرية على مستوى قاعدة البيانات</summary>
    Private Async Function UpdateUserPermissionsAdminProtected() As Task(Of Integer)
        Dim sql As String = $"UPDATE {AppConstants.Table_Users} SET {AppConstants.Col_CanEdit} = 1, {AppConstants.Col_CanDelete} = 1 WHERE LOWER({AppConstants.Col_Username}) = 'admin'"
        Return Await DatabaseModule.ExecuteNonQueryAsync(sql)
    End Function
End Class
