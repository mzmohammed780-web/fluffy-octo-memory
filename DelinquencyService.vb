Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Data
Imports System.Threading.Tasks
Imports System.Globalization

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — خدمة تقرير المتأخرات وغرامات التأخير
' الفكرة: بناءً على اشتراك شهري متفق عليه، نحسب لكل لاعب نشط:
'    * الأشهر المستحقة في السنة المختارة (من شهر الالتحاق حتى ديسمبر
'      أو حتى الشهر الحالي إن كانت السنة هي السنة الحالية)
'    * مجموع ما دفعه فعلاً في هذه السنة بالعملة المختارة
'    * العجز = المستحق - المدفوع، وعدد الأشهر المتأخرة = تقريب العجز للشهر الأعلى
'    * الغرامة التقديرية = الأشهر المتأخرة × غرامة الشهر
' ── حدود صدق التقرير (مهمة):
'    * الحساب يقارن بالمبالغ المدفوعة في العملة المختارة فقط — المدفوعات
'      بالعملات الأخرى تُعرض في عمود منفصل دون دمج (لا يوجد سعر صرف في النظام)
'    * الغرامة "تقديرية" معروضة للتفاوض — لا تُسجل في القاعدة تلقائياً
'    * من صادفه عجز يُعد متأخراً حتى لو دفع مبالغ متفرقة — الشهر المنطق واحد
' ═══════════════════════════════════════════════════════════════════
Public Class DelinquencyService

    ''' <summary>سطر نتيجة للاعب واحد</summary>
    Public Class DelinquencyRow
        Public PlayerId As String = ""
        Public PlayerName As String = ""
        Public Phone As String = ""
        Public JoinDate As String = ""
        Public MonthsDue As Integer = 0
        Public PaidSelected As Decimal = 0D
        Public PaidOtherCurrencies As String = ""
        Public ExpectedTotal As Decimal = 0D
        Public Shortfall As Decimal = 0D
        Public MonthsBehind As Integer = 0
        Public EstimatedFine As Decimal = 0D
    End Class

    Public Class ReportResult
        Public Success As Boolean = False
        Public ErrorMessage As String = ""
        Public Rows As New List(Of DelinquencyRow)
        Public Currency As String = ""
        Public Year As Integer = 0
        Public MonthlyDue As Decimal = 0D
        Public FinePerMonth As Decimal = 0D
        Public TotalShortfall As Decimal = 0D
        Public TotalFine As Decimal = 0D
        Public DelinquentCount As Integer = 0
    End Class

    ''' <summary>حساب أشهر السنة المستحقة للاعب التحق في تاريخ معين</summary>
    Private Function ComputeMonthsDue(joinDateText As String, year As Integer) As Integer
        ' 🌟 إصلاح: تحليل بثقافة ثابتة عبر SafeDate (يدعم ISO و dd/MM/yyyy القديمة) —
        ' كان Date.TryParse بثقافة النظام فينقلب الشهر/اليوم على أجهزة en-US وتصير المستحقات خاطئة
        Dim joinDate As Date = UtilityModule.SafeDate(joinDateText)
        If joinDate = Date.MinValue Then
            ' تاريخ تحقق غير مقروء؟ نعتبره التحق منذ بداية السنة (تقدير متحفظ لصالح النادي)
            joinDate = New Date(year, 1, 1)
        End If

        Dim startMonth As Integer = 1
        If joinDate.Year >= year Then
            If joinDate.Year > year Then Return 0        ' التحق بعد هذه السنة — لا مستحقات فيها
            startMonth = joinDate.Month                  ' التحق خلال هذه السنة
        End If

        Dim endMonth As Integer = 12
        If year = DateTime.Now.Year Then
            endMonth = DateTime.Now.Month                ' السنة الحالية: حتى الشهر الحالي فقط
        End If

        If endMonth < startMonth Then Return 0
        Return endMonth - startMonth + 1
    End Function

    ''' <summary>توليد التقرير الكامل — عملية قراءة فقط لا تعدل أي بيانات</summary>
    ''' <param name="year">السنة المحللة</param>
    ''' <param name="currency">العملة المرجعية للاشتراك والغرامة</param>
    ''' <param name="monthlyDue">الاشتراك الشهري المستحق</param>
    ''' <param name="finePerMonth">غرامة التأخير عن كل شهر</param>
    Public Async Function GenerateReportAsync(year As Integer, currency As String,
                                              monthlyDue As Decimal, finePerMonth As Decimal) As Task(Of ReportResult)
        Dim result As New ReportResult()
        result.Year = year
        result.Currency = currency
        result.MonthlyDue = monthlyDue
        result.FinePerMonth = finePerMonth

        Try
            ' ── 1. اللاعبون النشطون فقط (غير المؤرشفين) ──
            Dim playersSql As String =
                $"SELECT {AppConstants.Col_PlayerId}, {AppConstants.Col_PlayerName}, " &
                $"{AppConstants.Col_Fone}, {AppConstants.Col_DateIn} " &
                $"FROM {AppConstants.Table_Players} ORDER BY {AppConstants.Col_PlayerName}"

            Dim players As DataTable = Await DatabaseModule.ExecuteQueryAsync(playersSql)
            If players Is Nothing Then
                result.ErrorMessage = "تعذرت قراءة جدول اللاعبين"
                Return result
            End If

            ' ── 2. مدفوعات السنة كاملة مرة واحدة ثم تجميعها بالذاكرة (أسرع من استعلام لكل لاعب) ──
            Dim yearPrefix As String = year.ToString(CultureInfo.InvariantCulture)
            Dim paymentsSql As String =
                $"SELECT {AppConstants.Col_Pay_PlayerId}, {AppConstants.Col_Pay_PlayerName}, " &
                $"{AppConstants.Col_Pay_CurrencyType}, {AppConstants.Col_Pay_Amount} " &
                $"FROM {AppConstants.Table_Payments} " &
                $"WHERE substr({AppConstants.Col_Pay_PaymentDate}, 1, 4) = @y"

            Dim payments As DataTable = Await DatabaseModule.ExecuteQueryAsync(
                paymentsSql, New SQLiteParameter("@y", yearPrefix))
            If payments Is Nothing Then payments = New DataTable()

            ' 🌟 تجميع حسب معرف اللاعب (أو الاسم عند غياب المعرف) — مقاوم لغياب الربط
            ' key: "id:123" أو "name:أحمد"  →  value: قاموس (عملة → مجموع)
            Dim totals As New Dictionary(Of String, Dictionary(Of String, Decimal))
            Dim idToName As New Dictionary(Of String, String)

            For Each row As DataRow In payments.Rows
                Dim playerId As String = If(IsDBNull(row(AppConstants.Col_Pay_PlayerId)), "", Convert.ToString(row(AppConstants.Col_Pay_PlayerId))).Trim()
                Dim playerName As String = If(IsDBNull(row(AppConstants.Col_Pay_PlayerName)), "", Convert.ToString(row(AppConstants.Col_Pay_PlayerName))).Trim()
                Dim currencyVal As String = If(IsDBNull(row(AppConstants.Col_Pay_CurrencyType)), "", Convert.ToString(row(AppConstants.Col_Pay_CurrencyType))).Trim()
                Dim amount As Decimal = UtilityModule.ToDecimalSafe(row(AppConstants.Col_Pay_Amount))

                Dim key As String = If(playerId <> "", "id:" & playerId, If(playerName <> "", "name:" & playerName, ""))
                ' 🌟 إصلاح: العملة الفارغة (قيم قديمة) تعامل شيكل — افتراضي العمود في القاعدة —
                ' كانت تُستبعد كلياً فيظهر اللاعب متأخراً رغم أنه دفع
                If currencyVal = "" Then currencyVal = AppConstants.Currency_ILS
                If key = "" OrElse amount = 0D Then Continue For

                If Not totals.ContainsKey(key) Then
                    totals(key) = New Dictionary(Of String, Decimal)
                End If
                If Not totals(key).ContainsKey(currencyVal) Then
                    totals(key)(currencyVal) = 0D
                End If
                totals(key)(currencyVal) += amount

                If playerId <> "" AndAlso playerName <> "" AndAlso Not idToName.ContainsKey(playerId) Then
                    idToName(playerId) = playerName
                End If
            Next

            ' ── 3. الحساب لكل لاعب نشط ──
            For Each pRow As DataRow In players.Rows
                Dim playerId As String = If(IsDBNull(pRow(AppConstants.Col_PlayerId)), "", Convert.ToString(pRow(AppConstants.Col_PlayerId))).Trim()
                Dim playerName As String = If(IsDBNull(pRow(AppConstants.Col_PlayerName)), "", Convert.ToString(pRow(AppConstants.Col_PlayerName))).Trim()
                If playerName = "" Then Continue For

                Dim info As New DelinquencyRow()
                info.PlayerId = playerId
                info.PlayerName = playerName
                info.Phone = If(IsDBNull(pRow(AppConstants.Col_Fone)), "", Convert.ToString(pRow(AppConstants.Col_Fone))).Trim()
                info.JoinDate = If(IsDBNull(pRow(AppConstants.Col_DateIn)), "", Convert.ToString(pRow(AppConstants.Col_DateIn))).Trim()
                info.MonthsDue = ComputeMonthsDue(info.JoinDate, year)

                If monthlyDue > 0D Then
                    info.ExpectedTotal = CDec(info.MonthsDue) * monthlyDue
                End If

                ' مدفوعاته هذه السنة: بالعملة المرجعية + تعريف موجز للعملات الأخرى
                ' 🌟 دمج المصدرين: المدفوعات المربوطة بمعرف اللاعب + القديمة المربوطة باسمه فقط
                '    (قد يكون نفس اللاعب مدفوعاته موزعة بين المصدرين)
                Dim selectedCurrencyPaid As Decimal = 0D
                Dim otherParts As New List(Of String)

                For Each lookupKey As String In New String() {If(playerId <> "", "id:" & playerId, ""), If(playerName <> "", "name:" & playerName, "")}
                    If lookupKey = "" Then Continue For
                    Dim perCurrency As Dictionary(Of String, Decimal) = Nothing
                    If totals.TryGetValue(lookupKey, perCurrency) Then
                        For Each kv As KeyValuePair(Of String, Decimal) In perCurrency
                            If kv.Key = currency Then
                                selectedCurrencyPaid += kv.Value
                            Else
                                ' لا نجمع عملات مختلفة أبداً — نعرضها نصاً فقط
                                otherParts.Add($"{kv.Key}: {kv.Value.ToString("0.##", CultureInfo.InvariantCulture)}")
                            End If
                        Next
                    End If
                Next

                info.PaidSelected = selectedCurrencyPaid
                info.PaidOtherCurrencies = String.Join(" ، ", otherParts)

                ' العجز والأشهر المتأخرة والغرامة التقديرية
                If info.ExpectedTotal > info.PaidSelected AndAlso monthlyDue > 0D Then
                    info.Shortfall = info.ExpectedTotal - info.PaidSelected
                    Dim monthsBehindRaw As Decimal = info.Shortfall / monthlyDue
                    info.MonthsBehind = CInt(Math.Ceiling(monthsBehindRaw))
                    If finePerMonth > 0D Then
                        info.EstimatedFine = CDec(info.MonthsBehind) * finePerMonth
                    End If
                    result.TotalShortfall += info.Shortfall
                    result.TotalFine += info.EstimatedFine
                    result.DelinquentCount += 1
                End If

                ' نعرض الجميع (بما فيهم المسددون) حتى يتضح الغياب — الفرز يليق بالمشكلة
                result.Rows.Add(info)
            Next

            result.Success = True
            Return result

        Catch ex As Exception
            DatabaseModule.LogError("DelinquencyService.GenerateReportAsync", ex)
            result.ErrorMessage = "خطأ في توليد التقرير: " & ex.Message
            Return result
        End Try
    End Function

End Class
