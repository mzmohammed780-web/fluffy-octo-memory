Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data.SQLite
Imports System.Threading
Imports System.Threading.Tasks

Public Class PaymentService


    ''' <summary>تحميل كل الدفعات — الربط برقم الهوية، والاسم/الجوال يُجلبان من اللاعب أينما كان (نشط/مؤرشف)</summary>
    Public Async Function GetAllPaymentsAsync() As Task(Of DataTable)
        Dim sql As String =
            "SELECT Pay.ID, " &
            "COALESCE(P.Playername, A.Playername, Pay.Playername) AS Playername, " &
            "Pay.PlayerId, Pay.TransferorName, Pay.VoucherNumber, " &
            "Pay.Amount, Pay.CurrencyType, Pay.PaymentMethod, " &
            "Pay.PaymentDate, Pay.DueDate, Pay.Notes, " &
            "COALESCE(P.Fone, A.Fone) AS MobileNumber, " &
            "Pay.LastModifiedBy, Pay.LastModifiedDate " &
            "FROM " & AppConstants.Table_Payments & " AS Pay " &
            "LEFT JOIN Players AS P ON P.Playerid = Pay.PlayerId " &
            "LEFT JOIN archive AS A ON A.Playerid = Pay.PlayerId " &
            "ORDER BY Playername, Pay.PaymentDate DESC"
        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql)
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_Pay_PaymentDate, AppConstants.Col_Pay_DueDate)
        Return dt
    End Function

    Public Async Function InsertPaymentAsync(
        playerName As String, playerId As Object, paymentDate As String, voucherNumber As String,
        amount As Decimal, paymentMethod As String, transferorName As String,
        notes As String, dueDate As String, currencyType As String,
        modifiedBy As String, modifiedDate As String, voucherImage As Object) As Task(Of Integer)


        ' 🌟 حماية صلاحيات: إضافة الدفعات تتطلب صلاحية التعديل (توحيد مع ExpenseService/PlayerService)
        If Not UserSession.CanEdit Then Return 0

        Dim params As New List(Of SQLiteParameter)
        params.Add(New SQLiteParameter("@PlayerId", If(playerId Is Nothing, DBNull.Value, playerId)))
        params.Add(New SQLiteParameter("@Playername", If(String.IsNullOrWhiteSpace(playerName), DBNull.Value, CObj(playerName))))
        params.Add(New SQLiteParameter("@PaymentDate", If(String.IsNullOrWhiteSpace(paymentDate), DBNull.Value, CObj(paymentDate))))
        params.Add(New SQLiteParameter("@VoucherNumber", UtilityModule.ToDBValue(voucherNumber)))
        params.Add(New SQLiteParameter("@Amount", amount))
        params.Add(New SQLiteParameter("@PaymentMethod", UtilityModule.ToDBValue(paymentMethod)))
        params.Add(New SQLiteParameter("@TransferorName", UtilityModule.ToDBValue(transferorName)))
        params.Add(New SQLiteParameter("@Notes", UtilityModule.ToDBValue(notes)))
        params.Add(New SQLiteParameter("@DueDate", If(String.IsNullOrWhiteSpace(dueDate), DBNull.Value, CObj(dueDate))))
        params.Add(New SQLiteParameter("@CurrencyType", UtilityModule.ToDBValue(currencyType)))
        ' ⭐ توحيد مع السند الشامل: كتابة CreatedDate (وقت إنشاء السجل بالنظام) — كان يبقى NULL في السند المفرد
        params.Add(New SQLiteParameter("@CreatedDate", If(String.IsNullOrWhiteSpace(modifiedDate), DBNull.Value, CObj(modifiedDate))))
        params.Add(New SQLiteParameter("@LastModifiedBy", If(String.IsNullOrWhiteSpace(modifiedBy), DBNull.Value, CObj(modifiedBy))))
        params.Add(New SQLiteParameter("@LastModifiedDate", If(String.IsNullOrWhiteSpace(modifiedDate), DBNull.Value, CObj(modifiedDate))))
        params.Add(New SQLiteParameter("@VoucherImage", If(voucherImage Is Nothing, DBNull.Value, voucherImage)))

        Dim sql As String =
            "INSERT INTO " & AppConstants.Table_Payments & " " &
            "(PlayerId, Playername, PaymentDate, VoucherNumber, Amount, PaymentMethod, TransferorName, Notes, DueDate, CurrencyType, CreatedDate, LastModifiedBy, LastModifiedDate, VoucherImage) " &
            "VALUES (@PlayerId, @Playername, @PaymentDate, @VoucherNumber, @Amount, @PaymentMethod, @TransferorName, @Notes, @DueDate, @CurrencyType, @CreatedDate, @LastModifiedBy, @LastModifiedDate, @VoucherImage)"

        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, params.ToArray())

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Dim invCult As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture
            Dim details As String =
                "اللاعب: " & playerName &
                " — المبلغ: " & amount.ToString("0.##", invCult) &
                " " & currencyType
            Await audit.LogAsync(AuditService.Act_PaymentAdd, "دفعة",
                If(String.IsNullOrWhiteSpace(voucherNumber), "-", voucherNumber),
                details)
        End If

        Return affected
    End Function

    ''' <summary>تحديث دفعة موجودة بالمعرف الفريد ID (مع تحديث ربط اللاعب بالـ ID)</summary>
    Public Async Function UpdatePaymentAsync(
        paymentId As Long,
        playerName As String, playerId As Object, paymentDate As String, voucherNumber As String,
        amount As Decimal, paymentMethod As String, transferorName As String,
        notes As String, dueDate As String, currencyType As String,
        modifiedBy As String, modifiedDate As String,
        voucherImage As Object, imageChanged As Boolean) As Task(Of Integer)


        ' 🌟 حماية صلاحيات: تعديل الدفعات يتطلب صلاحية التعديل
        If Not UserSession.CanEdit Then Return 0

        Dim parameters As New List(Of SQLiteParameter)
        parameters.Add(New SQLiteParameter("@ID", paymentId))
        parameters.Add(New SQLiteParameter("@PlayerId", If(playerId Is Nothing, DBNull.Value, playerId)))
        parameters.Add(New SQLiteParameter("@Playername", If(String.IsNullOrWhiteSpace(playerName), DBNull.Value, CObj(playerName))))
        parameters.Add(New SQLiteParameter("@PaymentDate", If(String.IsNullOrWhiteSpace(paymentDate), DBNull.Value, CObj(paymentDate))))
        parameters.Add(New SQLiteParameter("@VoucherNumber", UtilityModule.ToDBValue(voucherNumber)))
        parameters.Add(New SQLiteParameter("@Amount", amount))
        parameters.Add(New SQLiteParameter("@PaymentMethod", UtilityModule.ToDBValue(paymentMethod)))
        parameters.Add(New SQLiteParameter("@TransferorName", UtilityModule.ToDBValue(transferorName)))
        parameters.Add(New SQLiteParameter("@Notes", UtilityModule.ToDBValue(notes)))
        parameters.Add(New SQLiteParameter("@DueDate", If(String.IsNullOrWhiteSpace(dueDate), DBNull.Value, CObj(dueDate))))
        parameters.Add(New SQLiteParameter("@CurrencyType", UtilityModule.ToDBValue(currencyType)))
        parameters.Add(New SQLiteParameter("@LastModifiedBy", If(String.IsNullOrWhiteSpace(modifiedBy), DBNull.Value, CObj(modifiedBy))))
        parameters.Add(New SQLiteParameter("@LastModifiedDate", modifiedDate))

        Dim sql As String =
            "UPDATE " & AppConstants.Table_Payments & " SET PlayerId=@PlayerId, Playername=@Playername, PaymentDate=@PaymentDate, " &
            "VoucherNumber=@VoucherNumber, Amount=@Amount, PaymentMethod=@PaymentMethod, " &
            "TransferorName=@TransferorName, Notes=@Notes, " &
            "DueDate=@DueDate, CurrencyType=@CurrencyType, " &
            "LastModifiedBy=@LastModifiedBy, LastModifiedDate=@LastModifiedDate "
        If imageChanged Then
            sql &= ", VoucherImage=@VoucherImage "
            parameters.Add(New SQLiteParameter("@VoucherImage", If(voucherImage Is Nothing, DBNull.Value, voucherImage)))
        End If

        sql &= "WHERE ID=@ID"

        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, parameters.ToArray())

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Dim invCult As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture
            Dim details As String =
                "المعرف: " & paymentId.ToString(invCult) &
                " — اللاعب: " & playerName &
                " — المبلغ: " & amount.ToString("0.##", invCult) &
                " " & currencyType
            Await audit.LogAsync(AuditService.Act_PaymentUpdate, "دفعة",
                If(String.IsNullOrWhiteSpace(voucherNumber), "-", voucherNumber),
                details)
        End If

        Return affected
    End Function

    ''' <summary>L-09: بناء شروط الفلترة المشتركة (بحث/ترقيم/مجاميع) — مصدر وحيد يمنع الانحراف.
    ''' الدالة متزامنة، لذا ByRef هنا قانوني (المنع خاص بالدوال غير المتزامنة فقط)</summary>
    Private Function BuildPaymentFilters(
        playerName As String, voucherNumber As String, transferorName As String,
        notes As String, paymentMethod As String, currencyType As String,
        ByRef parameters As List(Of SQLiteParameter)) As List(Of String)

        Dim conditions As New List(Of String)
        parameters = New List(Of SQLiteParameter)

        If Not String.IsNullOrWhiteSpace(playerName) AndAlso playerName <> "-- اختر اللاعب --" Then
            conditions.Add("(Pay.Playername LIKE @Playername OR P.Playername LIKE @Playername)")
            parameters.Add(New SQLiteParameter("@Playername", "%" & playerName.Trim() & "%"))
        End If
        If Not String.IsNullOrWhiteSpace(voucherNumber) Then
            conditions.Add("Pay.VoucherNumber LIKE @VoucherNumber")
            parameters.Add(New SQLiteParameter("@VoucherNumber", "%" & voucherNumber.Trim() & "%"))
        End If
        If Not String.IsNullOrWhiteSpace(transferorName) Then
            conditions.Add("Pay.TransferorName LIKE @TransferorName")
            parameters.Add(New SQLiteParameter("@TransferorName", "%" & transferorName.Trim() & "%"))
        End If
        If Not String.IsNullOrWhiteSpace(notes) Then
            conditions.Add("Pay.Notes LIKE @Notes")
            parameters.Add(New SQLiteParameter("@Notes", "%" & notes.Trim() & "%"))
        End If
        If Not String.IsNullOrWhiteSpace(paymentMethod) Then
            conditions.Add("Pay.PaymentMethod = @PaymentMethod")
            parameters.Add(New SQLiteParameter("@PaymentMethod", paymentMethod.Trim()))
        End If
        If Not String.IsNullOrWhiteSpace(currencyType) Then
            conditions.Add("Pay.CurrencyType = @CurrencyType")
            parameters.Add(New SQLiteParameter("@CurrencyType", currencyType.Trim()))
        End If

        Return conditions
    End Function

    Public Async Function SearchPaymentsAsync(
        playerName As String, voucherNumber As String, transferorName As String,
        notes As String, paymentMethod As String, currencyType As String,
        ct As Threading.CancellationToken) As Task(Of DataTable)

        ' 🌟 L-09: الشروط الستة السابقة نفسها حرفياً — انتقلت إلى BuildPaymentFilters (مصدر وحيد)
        Dim parameters As List(Of SQLiteParameter) = Nothing
        Dim conditions As List(Of String) = BuildPaymentFilters(
            playerName, voucherNumber, transferorName,
            notes, paymentMethod, currencyType, parameters)

        Dim sql As String =
            "SELECT Pay.ID, " &
            "COALESCE(P.Playername, A.Playername, Pay.Playername) AS Playername, " &
            "Pay.PlayerId, Pay.TransferorName, Pay.VoucherNumber, Pay.Amount, " &
            "Pay.CurrencyType, Pay.PaymentMethod, Pay.PaymentDate, " &
            "Pay.DueDate, Pay.Notes, COALESCE(P.Fone, A.Fone) AS MobileNumber, " &
            "Pay.LastModifiedBy, Pay.LastModifiedDate " &
            "FROM " & AppConstants.Table_Payments & " AS Pay " &
            "LEFT JOIN Players AS P ON P.Playerid = Pay.PlayerId " &
            "LEFT JOIN archive AS A ON A.Playerid = Pay.PlayerId " &
            If(conditions.Count > 0, " WHERE " & String.Join(" AND ", conditions), "") &
            " ORDER BY Playername, Pay.PaymentDate DESC"
        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, ct, parameters.ToArray())
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_Pay_PaymentDate, AppConstants.Col_Pay_DueDate)
        Return dt
    End Function

    ''' <summary>L-09 — ترقيم تزايدي: نفس استعلام البحث تماماً (نفس الأعمدة والفلاتر والترتيب)
    ''' مع LIMIT/OFFSET عبر وسائط مُعرّفة. الفلاتر الفارغة = كل الدفعات.</summary>
    Public Async Function SearchPaymentsPagedAsync(
        playerName As String, voucherNumber As String, transferorName As String,
        notes As String, paymentMethod As String, currencyType As String,
        ct As Threading.CancellationToken, limit As Integer, offset As Integer) As Task(Of DataTable)

        Dim parameters As List(Of SQLiteParameter) = Nothing
        Dim conditions As List(Of String) = BuildPaymentFilters(
            playerName, voucherNumber, transferorName,
            notes, paymentMethod, currencyType, parameters)

        Dim sql As String =
            "SELECT Pay.ID, " &
            "COALESCE(P.Playername, A.Playername, Pay.Playername) AS Playername, " &
            "Pay.PlayerId, Pay.TransferorName, Pay.VoucherNumber, Pay.Amount, " &
            "Pay.CurrencyType, Pay.PaymentMethod, Pay.PaymentDate, " &
            "Pay.DueDate, Pay.Notes, COALESCE(P.Fone, A.Fone) AS MobileNumber, " &
            "Pay.LastModifiedBy, Pay.LastModifiedDate " &
            "FROM " & AppConstants.Table_Payments & " AS Pay " &
            "LEFT JOIN Players AS P ON P.Playerid = Pay.PlayerId " &
            "LEFT JOIN archive AS A ON A.Playerid = Pay.PlayerId " &
            If(conditions.Count > 0, " WHERE " & String.Join(" AND ", conditions), "") &
            " ORDER BY Playername, Pay.PaymentDate DESC" &
            " LIMIT @limit OFFSET @offset"
        parameters.Add(New SQLiteParameter("@limit", limit))
        parameters.Add(New SQLiteParameter("@offset", offset))

        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, ct, parameters.ToArray())
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_Pay_PaymentDate, AppConstants.Col_Pay_DueDate)
        Return dt
    End Function

    ''' <summary>L-09 — المجاميع الحقيقية لكل النتائج المطابقة (لا الصفوف المعروضة فقط).
    ''' تصنيف العملات مطابق تماماً لمنطق SumCurrencyTotalsFromRows: الدولار/الدينار بالتطابق التام،
    ''' وكل ما عدا ذلك (NULL/''/شيكل/أي قيمة أخرى) يذهب لعمود الشيكل.
    ''' Count = سطور ذات Amount غير NULL (مطابق لتخطي DBNull بالشبكة)،
    ''' Item2 = إجمالي السطور المطابقة COUNT(*) بغض النظر عن خلوّها من مبلغ.</summary>
    Public Async Function GetPaymentsTotalsAsync(
        playerName As String, voucherNumber As String, transferorName As String,
        notes As String, paymentMethod As String, currencyType As String,
        Optional ct As Threading.CancellationToken = Nothing) As Task(Of Tuple(Of CurrencyTotals, Integer))

        Dim parameters As List(Of SQLiteParameter) = Nothing
        Dim conditions As List(Of String) = BuildPaymentFilters(
            playerName, voucherNumber, transferorName,
            notes, paymentMethod, currencyType, parameters)

        Dim sql As String =
            "SELECT COUNT(*) AS TotalRows, COUNT(Pay.Amount) AS CountedAmounts, " &
            "IFNULL(SUM(CASE WHEN Pay.CurrencyType = '" & AppConstants.Currency_USD & "' THEN Pay.Amount ELSE 0 END), 0) AS TotalUSD, " &
            "IFNULL(SUM(CASE WHEN Pay.CurrencyType = '" & AppConstants.Currency_JOD & "' THEN Pay.Amount ELSE 0 END), 0) AS TotalJOD, " &
            "IFNULL(SUM(CASE WHEN Pay.CurrencyType IS NULL OR (Pay.CurrencyType <> '" & AppConstants.Currency_USD & "' AND Pay.CurrencyType <> '" & AppConstants.Currency_JOD & "') THEN Pay.Amount ELSE 0 END), 0) AS TotalILS " &
            "FROM " & AppConstants.Table_Payments & " AS Pay " &
            "LEFT JOIN Players AS P ON P.Playerid = Pay.PlayerId " &
            "LEFT JOIN archive AS A ON A.Playerid = Pay.PlayerId " &
            If(conditions.Count > 0, " WHERE " & String.Join(" AND ", conditions), "")
        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, ct, parameters.ToArray())

        Dim t As New CurrencyTotals()
        Dim totalRows As Integer = 0
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim r As DataRow = dt.Rows(0)
            totalRows = Convert.ToInt32(r("TotalRows"))
            t.Count = Convert.ToInt32(r("CountedAmounts"))
            t.USD = Convert.ToDecimal(r("TotalUSD"))
            t.JOD = Convert.ToDecimal(r("TotalJOD"))
            t.ILS = Convert.ToDecimal(r("TotalILS"))
        End If
        Return Tuple.Create(t, totalRows)
    End Function

    ''' <summary>تحميل أسماء اللاعبين النشطين والمؤرشفين لقائمة المنسدلة (مع المعرف الفريد)</summary>
    Public Async Function LoadPlayersComboDataAsync() As Task(Of Tuple(Of DataTable, DataTable))
        Dim activeSql As String =
            "SELECT " & AppConstants.Col_PlayerId & ", " & AppConstants.Col_PlayerName & ", " & AppConstants.Col_Fone &
            " FROM " & AppConstants.Table_Players &
            " ORDER BY " & AppConstants.Col_PlayerName
        Dim archiveSql As String =
            "SELECT " & AppConstants.Col_PlayerId & ", " & AppConstants.Col_PlayerName & ", " & AppConstants.Col_Fone &
            " FROM " & AppConstants.Table_Archive &
            " ORDER BY " & AppConstants.Col_PlayerName
        Dim activeDt As DataTable = Await DatabaseModule.ExecuteQueryAsync(activeSql)
        Dim archiveDt As DataTable = Await DatabaseModule.ExecuteQueryAsync(archiveSql)
        Return Tuple.Create(activeDt, archiveDt)
    End Function

    ''' <summary>
    ''' هل (رقم السند + اللاعب + العملة) مكرر؟ — نفس الشخص بعملتين مختلفتين مسموح.
    ''' excludePaymentId لاستثناء الدفعة قيد التعديل
    ''' </summary>
    Public Async Function IsVoucherNumberExistsAsync(voucherNumber As String, Optional excludePaymentId As Long = -1, Optional playerId As Object = Nothing, Optional currency As String = "") As Task(Of Boolean)
        Try
            If String.IsNullOrWhiteSpace(voucherNumber) Then Return False

            Dim conditions As New List(Of String)
            Dim params As New List(Of SQLiteParameter)

            conditions.Add("TRIM(" & AppConstants.Col_Pay_VoucherNumber & ") = @v")
            conditions.Add("PlayerId IS NOT NULL")
            params.Add(New SQLiteParameter("@v", voucherNumber.Trim()))

            ' 🌟 إصلاح: شرط العملة يُضاف فقط عند تمرير عملة فعلياً —
            ' كان يُضاف دائماً بقيمة '' فيصبح الفحص شبه معطّل (لا يطابق أي صف طبيعي عملته شيكل/دولار/دينار)
            If Not String.IsNullOrWhiteSpace(currency) Then
                conditions.Add("COALESCE(CurrencyType, '') = @cur")
                params.Add(New SQLiteParameter("@cur", currency.Trim()))
            End If

            If playerId IsNot Nothing AndAlso Not IsDBNull(playerId) Then
                conditions.Add("PlayerId = @pid")
                params.Add(New SQLiteParameter("@pid", playerId))
            End If

            If excludePaymentId > 0 Then
                conditions.Add(AppConstants.Col_Pay_ID & " <> @id")
                params.Add(New SQLiteParameter("@id", excludePaymentId))
            End If

            Dim sql As String = "SELECT COUNT(*) FROM " & AppConstants.Table_Payments & " WHERE " & String.Join(" AND ", conditions)
            Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(sql, params.ToArray())
            Return Convert.ToInt32(result) > 0
        Catch ex As Exception
            DatabaseModule.LogError("PaymentService.IsVoucherNumberExistsAsync", ex)
            Throw
        End Try
    End Function

    ''' <summary>جلب أسماء كل من يستخدم رقم السند (للعرض في فحص التكرار) — قائمة قد تكون متعددة (سند شامل)</summary>
    Public Async Function GetVoucherOwnersAsync(voucherNumber As String) As Task(Of List(Of String))
        Dim owners As New List(Of String)
        Try
            If String.IsNullOrWhiteSpace(voucherNumber) Then Return owners

            Dim sql As String =
                "SELECT DISTINCT COALESCE(P.Playername, A.Playername, Pay.Playername) AS OwnerName " &
                "FROM " & AppConstants.Table_Payments & " Pay " &
                "LEFT JOIN Players P ON P.Playerid = Pay.PlayerId " &
                "LEFT JOIN archive A ON A.Playerid = Pay.PlayerId " &
                "WHERE TRIM(Pay.VoucherNumber) = @v AND Pay.Playername IS NOT NULL " &
                "ORDER BY OwnerName"

            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, New SQLiteParameter("@v", voucherNumber.Trim()))
            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    Dim n As String = UtilityModule.SafeString(row("OwnerName"))
                    If n <> "" Then owners.Add(n)
                Next
            End If
        Catch ex As Exception
            DatabaseModule.LogError("PaymentService.GetVoucherOwnersAsync", ex)
        End Try
        Return owners
    End Function

    ''' <summary>حذف دفعة بالمعرف الفريد ID</summary>
    Public Async Function DeletePaymentAsync(paymentId As Long) As Task(Of Integer)
        ' 🌟 حماية صلاحيات: حذف الدفعات يتطلب صلاحية الحذف
        If Not UserSession.CanDelete Then Return 0
        Dim sql As String = "DELETE FROM " & AppConstants.Table_Payments & " WHERE " & AppConstants.Col_Pay_ID & " = @id"
        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, New SQLiteParameter("@id", paymentId))

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_PaymentDelete, "دفعة",
                paymentId.ToString(System.Globalization.CultureInfo.InvariantCulture), "")
        End If

        Return affected
    End Function

    ''' <summary>جلب صورة السند بالمعرف الفريد ID</summary>
    Public Async Function GetVoucherImageAsync(paymentId As Long) As Task(Of Byte())
        Try
            If paymentId <= 0 Then Return Nothing

            ' 1) العمود القديم أولاً — الصفوف الفردية والسجلات غير المُرحّلة
            Dim sql As String = "SELECT VoucherImage FROM " & AppConstants.Table_Payments & " WHERE " & AppConstants.Col_Pay_ID & " = @id"
            Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(sql, New SQLiteParameter("@id", paymentId))
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CType(result, Byte())
            End If

            ' 🔴 H-04: الصور المُرحّلة — من الجدول المستقل عبر رقم سند الصف نفسه
            Dim sql2 As String =
                "SELECT vi.ImageData FROM VoucherImages vi WHERE vi.VoucherNumber = " &
                "(SELECT TRIM(p.VoucherNumber) FROM " & AppConstants.Table_Payments & " p WHERE p." & AppConstants.Col_Pay_ID & " = @id " &
                " AND p.VoucherNumber IS NOT NULL AND TRIM(p.VoucherNumber) <> '')"
            result = Await DatabaseModule.ExecuteScalarAsync(sql2, New SQLiteParameter("@id", paymentId))
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                Return CType(result, Byte())
            End If
            Return Nothing
        Catch ex As Exception
            DatabaseModule.LogError("PaymentService.GetVoucherImageAsync", ex)
            Throw ' 🌟 الفشل الصامت يسمح بمرور سندات مكررة — نرمي الخطأ ليفشل الحفظ بأمان
        End Try
    End Function

    ''' <summary>جلب رقم هوية اللاعب بالاسم (النشطون أولاً ثم الأرشيف) — يعيد Nothing إذا لم يوجد</summary>
    ''' <remarks>🌟 إصلاح: مع الأسماء المكررة كان LIMIT 1 بدون ترتيب يُرجع سجلاً عشوائياً (غير حتمي)،
    ''' فالدفعة قد تُربط للاعب مختلف في كل مرة. الآن يُختار أصغر رقم هوية بشكل حتمي دائماً،
    ''' ويُسجّل تنبيه في اللوج عند وجود تكرار لمراجعته</remarks>
    Public Async Function GetPlayerIdByNameAsync(playerName As String) As Task(Of Object)
        Try
            If String.IsNullOrWhiteSpace(playerName) Then Return Nothing
            Dim nameParam As New SQLiteParameter("@n", playerName.Trim())

            ' 🌟 تنبيه عند وجود أكثر من لاعب بنفس الاسم (يُسجّل باللوج للمراجعة)
            Dim dupSql As String =
                "SELECT COUNT(*) FROM " & AppConstants.Table_Players &
                " WHERE " & AppConstants.Col_PlayerName & " = @n"
            Dim dupCountObj As Object = Await DatabaseModule.ExecuteScalarAsync(dupSql, New SQLiteParameter("@n", playerName.Trim()))
            If dupCountObj IsNot Nothing AndAlso Convert.ToInt32(dupCountObj) > 1 Then
                Dim warnMsg As String =
                    "تنبيه دقة بيانات: يوجد أكثر من لاعب باسم '" & playerName.Trim() &
                    "' — تم الربط بأصغر رقم هوية. يُنصح بتمييز الأسماء المتشابهة."
                ' 🌟 M-05: تنبيه دقة بيانات — ليس خطأ تشغيل، يُسجّل بقناة التحذيرات بدل ErrorLog
                DatabaseModule.LogWarn("PaymentService.GetPlayerIdByNameAsync", warnMsg)
            End If

            Dim activeSql As String =
                "SELECT " & AppConstants.Col_PlayerId & " FROM " & AppConstants.Table_Players &
                " WHERE " & AppConstants.Col_PlayerName & " = @n" &
                " ORDER BY " & AppConstants.Col_PlayerId & " ASC LIMIT 1"
            Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(activeSql, nameParam)

            If result IsNot Nothing AndAlso Not IsDBNull(result) Then Return result

            Dim archiveSql As String =
                "SELECT " & AppConstants.Col_PlayerId & " FROM " & AppConstants.Table_Archive &
                " WHERE " & AppConstants.Col_PlayerName & " = @n" &
                " ORDER BY " & AppConstants.Col_PlayerId & " ASC LIMIT 1"
            result = Await DatabaseModule.ExecuteScalarAsync(archiveSql, New SQLiteParameter("@n", playerName.Trim()))

            Return If(result IsNot Nothing AndAlso Not IsDBNull(result), result, Nothing)
        Catch ex As Exception
            DatabaseModule.LogError("PaymentService.GetPlayerIdByNameAsync", ex)
            Throw
        End Try
    End Function

    ''' <summary>هل رقم السند مستخدم مسبقاً بأي سجل بالجدول؟ — لفحص رقم السند الشامل الجديد</summary>
    Public Async Function IsVoucherNumberUsedAsync(voucherNumber As String) As Task(Of Boolean)
        Try
            If String.IsNullOrWhiteSpace(voucherNumber) Then Return False
            Dim sql As String = "SELECT COUNT(*) FROM " & AppConstants.Table_Payments &
                " WHERE TRIM(" & AppConstants.Col_Pay_VoucherNumber & ") = @v"
            Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(sql, New SQLiteParameter("@v", voucherNumber.Trim()))
            Return Convert.ToInt32(result) > 0
        Catch ex As Exception
            DatabaseModule.LogError("PaymentService.IsVoucherNumberUsedAsync", ex)
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 🌟 هل رقم السند مستخدم في سطور "محول فقط" (بلا لاعب - PlayerId IS NULL)؟
    ''' هذه السطور خارج الفهرس الفريد (الذي يشمل PlayerId غير الفارغ فقط)
    ''' لذا تحتاج فحصاً مستقلاً يمنع تكرارها بصمت.
    ''' excludePaymentId لاستثناء الدفعة قيد التعديل.
    ''' </summary>
    Public Async Function IsVoucherNumberExistsForTransferorAsync(voucherNumber As String, Optional excludePaymentId As Long = -1) As Task(Of Boolean)
        Try
            If String.IsNullOrWhiteSpace(voucherNumber) Then Return False

            Dim conditions As New List(Of String)
            Dim params As New List(Of SQLiteParameter)

            conditions.Add("TRIM(" & AppConstants.Col_Pay_VoucherNumber & ") = @v")
            conditions.Add("PlayerId IS NULL")
            params.Add(New SQLiteParameter("@v", voucherNumber.Trim()))

            If excludePaymentId > 0 Then
                conditions.Add(AppConstants.Col_Pay_ID & " <> @id")
                params.Add(New SQLiteParameter("@id", excludePaymentId))
            End If

            Dim sql As String = "SELECT COUNT(*) FROM " & AppConstants.Table_Payments & " WHERE " & String.Join(" AND ", conditions)
            Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(sql, params.ToArray())
            Return Convert.ToInt32(result) > 0
        Catch ex As Exception
            DatabaseModule.LogError("PaymentService.IsVoucherNumberExistsForTransferorAsync", ex)
            Throw
        End Try
    End Function
    ''' <summary>جلب كل مطابقات الاسم (نشط + مؤرشف) مع الهاتف — للحوار الصريح عند الأسماء المكررة</summary>
    Public Async Function GetAllPlayerMatchesByNameAsync(playerName As String) As Task(Of List(Of Tuple(Of Long, String)))
        Dim result As New List(Of Tuple(Of Long, String))
        Try
            If String.IsNullOrWhiteSpace(playerName) Then Return result

            Dim sql As String =
                "SELECT Playerid, COALESCE(Fone, '') AS Fone FROM " & AppConstants.Table_Players &
                " WHERE Playername = @n " &
                "UNION ALL " &
                "SELECT Playerid, COALESCE(Fone, '') || ' — مؤرشف' AS Fone FROM " & AppConstants.Table_Archive &
                " WHERE Playername = @n " &
                "ORDER BY Playerid ASC"

            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql,
                New SQLiteParameter("@n", playerName.Trim()))

            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    Dim pid As Long = Convert.ToInt64(row("Playerid"))
                    Dim phone As String = UtilityModule.SafeString(row("Fone"))
                    result.Add(Tuple.Create(pid, phone))
                Next
            End If
        Catch ex As Exception
            DatabaseModule.LogError("PaymentService.GetAllPlayerMatchesByNameAsync", ex)
        End Try
        Return result
    End Function
End Class
