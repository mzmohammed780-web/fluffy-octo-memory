Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Threading.Tasks

Public Class DashboardService

    Public Structure DashboardData
        Public ActivePlayers As Integer
        Public ArchivedPlayers As Integer
        Public MonthRevenuesILS As Decimal
        Public MonthExpensesILS As Decimal
        Public MonthNetILS As Decimal
    End Structure

    ''' <summary>جلب ملخص الإحصائيات للشهر الحالي في استعلام واحد سريع جداً</summary>
    ''' <remarks>🌟 إصلاح: كان الفلترة بحد أدنى فقط (بدون حد أقصى) — أي دفعة أو مصروف بتاريخ مستقبلي
    ''' (خطأ إدخال أو تأجيل) يظل يُحسب ضمن الشهر الحالي للأبد. الآن النطاق مقيد بين أول الشهر وأوله من الشهر التالي</remarks>
    Public Async Function GetStatsAsync() As Task(Of DashboardData)
        Dim data As New DashboardData()
        Try
            Dim firstDayOfMonth As String = New Date(Date.Today.Year, Date.Today.Month, 1).ToString("yyyy-MM-dd")
            ' 🌟 الحد الأقصى = أول يوم من الشهر التالي (يستثني التواريخ المستقبلية)
            Dim firstDayOfNextMonth As String = New Date(Date.Today.Year, Date.Today.Month, 1).AddMonths(1).ToString("yyyy-MM-dd")

            ' 🌟 العملة الفارغة/NULL تنحسب شيكل — نفس منطق التقرير العام
            Dim sql As String =
                "SELECT " &
                $"(SELECT COUNT(*) FROM {AppConstants.Table_Players}) AS ActivePlayers, " &
                $"(SELECT COUNT(*) FROM {AppConstants.Table_Archive}) AS ArchivedPlayers, " &
                $"(SELECT IFNULL(SUM({AppConstants.Col_Pay_Amount}), 0) FROM {AppConstants.Table_Payments} " &
                $"WHERE {AppConstants.Col_Pay_PaymentDate} >= @fromDate AND {AppConstants.Col_Pay_PaymentDate} < @toDate " &
                $"AND ({AppConstants.Col_Pay_CurrencyType} = @cur OR {AppConstants.Col_Pay_CurrencyType} IS NULL OR {AppConstants.Col_Pay_CurrencyType} = '')) AS RevILS, " &
                $"(SELECT IFNULL(SUM({AppConstants.Col_Exp_Amount}), 0) FROM {AppConstants.Table_Expenses} " &
                $"WHERE {AppConstants.Col_Exp_ExpenseDate} >= @fromDate AND {AppConstants.Col_Exp_ExpenseDate} < @toDate " &
                $"AND ({AppConstants.Col_Exp_CurrencyType} = @cur OR {AppConstants.Col_Exp_CurrencyType} IS NULL OR {AppConstants.Col_Exp_CurrencyType} = '')) AS ExpILS"

            Dim params As SQLiteParameter() = {
                New SQLiteParameter("@fromDate", firstDayOfMonth),
                New SQLiteParameter("@toDate", firstDayOfNextMonth),
                New SQLiteParameter("@cur", AppConstants.Currency_ILS)
            }

            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, params)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim row As DataRow = dt.Rows(0)
                data.ActivePlayers = Convert.ToInt32(row("ActivePlayers"))
                data.ArchivedPlayers = Convert.ToInt32(row("ArchivedPlayers"))
                data.MonthRevenuesILS = Convert.ToDecimal(row("RevILS"))
                data.MonthExpensesILS = Convert.ToDecimal(row("ExpILS"))
                data.MonthNetILS = data.MonthRevenuesILS - data.MonthExpensesILS
            End If

        Catch ex As Exception
            DatabaseModule.LogError("DashboardService.GetStatsAsync", ex)
        End Try
        Return data
    End Function
End Class
