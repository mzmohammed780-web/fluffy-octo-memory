Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Threading.Tasks

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — مخزن إعدادات التطبيق (جدول AppSettings مفتاح/قيمة)
' يستخدمه: قفل الخمول (المدة)، وتقرير المتأخرات (الاشتراك الشهري/الغرامة/العملة)
' ── قواعد التصميم:
'    * لا يرمي استثناءات أبداً — أي فشل يعيد القيمة الافتراضية ويسجل في ملف اللوج
'    * الجدول يُنشأ تلقائياً عند أول استخدام (CREATE TABLE IF NOT EXISTS)
' ═══════════════════════════════════════════════════════════════════
Public Module AppSettingsStore

    Private Const Table_AppSettings As String = "AppSettings"
    Private _tableChecked As Boolean = False

    Public Sub EnsureAppSettingsTableExists()
        If _tableChecked Then Return
        Try
            DatabaseModule.ExecuteNonQuery(
                "CREATE TABLE IF NOT EXISTS " & Table_AppSettings & " (" & vbCrLf &
                "  [Key]   TEXT PRIMARY KEY," & vbCrLf &
                "  [Value] TEXT" & vbCrLf &
                ")")

            Dim existing As String = "|"
            Dim info As System.Data.DataTable = DatabaseModule.ExecuteQuery("PRAGMA table_info(" & Table_AppSettings & ")")
            For Each r As System.Data.DataRow In info.Rows
                existing &= Convert.ToString(r("name")).ToLowerInvariant() & "|"
            Next
            If Not existing.Contains("|value|") Then
                DatabaseModule.ExecuteNonQuery("ALTER TABLE " & Table_AppSettings & " ADD COLUMN [Value] TEXT")
            End If

            _tableChecked = True
        Catch ex As Exception
            DatabaseModule.LogError("AppSettingsStore.EnsureAppSettingsTableExists", ex)
        End Try
    End Sub
    ''' <summary>قراءة قيمة إعداد — تعيد fallback إذا لم توجد أو فشلت القراءة</summary>
    Public Function GetSetting(key As String, fallback As String) As String
        Try
            EnsureAppSettingsTableExists()
            Dim result As Object = DatabaseModule.ExecuteScalar(
                $"SELECT [Value] FROM {Table_AppSettings} WHERE [Key] = @k",
                New SQLiteParameter("@k", key))

            If result Is Nothing OrElse IsDBNull(result) Then Return fallback
            Return Convert.ToString(result)
        Catch ex As Exception
            DatabaseModule.LogError("AppSettingsStore.GetSetting(" & key & ")", ex)
            Return fallback
        End Try
    End Function

    ''' <summary>حفظ قيمة إعداد (INSERT أو UPDATE) — يعيد True عند النجاح</summary>
    Public Function SetSetting(key As String, value As String) As Boolean
        Try
            EnsureAppSettingsTableExists()
            ' 🌟 إصلاح: UPDATE ثم INSERT بدون معاملة — سباق صغير يسمح بفشل PK.
            ' INSERT OR REPLACE ذرية بعبارة واحدة
            DatabaseModule.ExecuteNonQuery(
                $"INSERT OR REPLACE INTO {Table_AppSettings} ([Key], [Value]) VALUES (@k, @v)",
                New SQLiteParameter("@k", key),
                New SQLiteParameter("@v", value))
            Return True
        Catch ex As Exception
            DatabaseModule.LogError("AppSettingsStore.SetSetting(" & key & ")", ex)
            Return False
        End Try
    End Function

    ''' <summary>قراءة عدد صحيح من الإعدادات مع حدّ أدنى وأقصى اختياريين</summary>
    Public Function GetIntSetting(key As String, fallback As Integer,
                                  Optional minValue As Integer = Integer.MinValue,
                                  Optional maxValue As Integer = Integer.MaxValue) As Integer
        Dim raw As String = GetSetting(key, fallback.ToString(System.Globalization.CultureInfo.InvariantCulture))
        Dim parsed As Integer
        If Not Integer.TryParse(raw, parsed) Then Return fallback
        If parsed < minValue Then Return minValue
        If parsed > maxValue Then Return maxValue
        Return parsed
    End Function

    ''' <summary>قراءة عدد عشري (Decimal) من الإعدادات مع حدّ أدنى اختياري</summary>
    Public Function GetDecimalSetting(key As String, fallback As Decimal,
                                      Optional minValue As Decimal = Decimal.MinValue) As Decimal
        Dim raw As String = GetSetting(key, fallback.ToString(System.Globalization.CultureInfo.InvariantCulture))
        Dim parsed As Decimal
        ' 🌟 InvariantCulture — نفس سياسة تحليل المبالغ الموحدة في المشروع
        If Not Decimal.TryParse(raw, System.Globalization.NumberStyles.Number,
                                System.Globalization.CultureInfo.InvariantCulture, parsed) Then Return fallback
        If parsed < minValue Then Return minValue
        Return parsed
    End Function

    ''' <summary>حفظ عدد صحيح</summary>
    Public Function SetIntSetting(key As String, value As Integer) As Boolean
        Return SetSetting(key, value.ToString(System.Globalization.CultureInfo.InvariantCulture))
    End Function

    ''' <summary>حفظ عدد عشري (بثقافة ثابتة لضمان قراءة صحيحة على أي جهاز)</summary>
    Public Function SetDecimalSetting(key As String, value As Decimal) As Boolean
        Return SetSetting(key, value.ToString(System.Globalization.CultureInfo.InvariantCulture))
    End Function

    ' ── مفاتيح الإعدادات المعروفة (توحيد لتجنب أخطاء إملائية) ──
    Public Const Key_IdleLockMinutes As String = "Security_IdleLockMinutes"
    Public Const Key_Delinq_MonthlyDue As String = "Delinq_MonthlyDue"
    Public Const Key_Delinq_FinePerMonth As String = "Delinq_FinePerMonth"
    Public Const Key_Delinq_Currency As String = "Delinq_Currency"
    Public Const Key_AuditRetentionDays As String = "Audit_RetentionDays"
    Public Const Key_Migration_DatesIso_Done As String = "Migration_DatesIso_Done"   ' 🌟 M-01: علم إتمام ترحيل التواريخ
    ' 🔴 H-04: علم إتمام ترحيل صور السندات إلى الجدول المستقل VoucherImages
    Public Const Key_Migration_VoucherImages_Done As String = "Migration_VoucherImages_Done"
    ' 🔴 H-02: حسابات ويندوز إضافية تُمنح صلاحية التعديل على مجلد البيانات (مفصولة بفواصل)
    ' ملاحظة: لكتابة "Users" هنا تعود الصلاحيات الواسعة القديمة لكل من يريدها (خيار استرجاع صريح)
    Public Const Key_DataFolderAccounts As String = "DataFolderAccounts"
    ' 🔴 H-02: علم تطبيق تقييد الصلاحيات على مجلد موجود مسبقاً (مرة واحدة — ذاتي الشفاء)
    Public Const Key_DataFolderAclTightened As String = "DataFolderAclTightened"

End Module
