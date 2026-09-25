Option Explicit On
Option Strict On

Imports System.Data.SQLite
Imports System.Threading
Imports System.Threading.Tasks

Public Class ExpenseService
    Public Async Function GetAllExpensesAsync() As Task(Of DataTable)
        Dim sql As String = $"SELECT ID, VoucherNumber, Description, Amount, CurrencyType, " &
                  $"ExpenseDate, Category, Notes, LastModifiedBy, LastModifiedDate FROM Expenses ORDER BY ExpenseDate DESC"
        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql)
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_Exp_ExpenseDate)
        Return dt
    End Function

    Public Async Function InsertExpenseAsync(
        voucherNumber As String, description As String, amount As Decimal,
        currencyType As String, expenseDate As String, category As String, notes As String,
        modifiedBy As String, modifiedDate As String) As Task(Of Integer)

        If Not UserSession.CanEdit Then Return 0

        Dim params As SQLiteParameter() = {
            New SQLiteParameter("@VoucherNumber", UtilityModule.ToDBValue(voucherNumber)),
            New SQLiteParameter("@Description", UtilityModule.ToDBValue(description)),
            New SQLiteParameter("@Amount", amount),
            New SQLiteParameter("@CurrencyType", UtilityModule.ToDBValue(currencyType)),
            New SQLiteParameter("@ExpenseDate", expenseDate),
            New SQLiteParameter("@Category", UtilityModule.ToDBValue(category)),
            New SQLiteParameter("@Notes", UtilityModule.ToDBValue(notes)),
            New SQLiteParameter("@LastModifiedBy", If(String.IsNullOrWhiteSpace(modifiedBy), DBNull.Value, CObj(modifiedBy))),
            New SQLiteParameter("@LastModifiedDate", modifiedDate)
        }

        Dim sql As String = $"INSERT INTO Expenses (VoucherNumber, Description, Amount, CurrencyType, ExpenseDate, Category, Notes, LastModifiedBy, LastModifiedDate) " &
                  "VALUES (@VoucherNumber, @Description, @Amount, @CurrencyType, @ExpenseDate, @Category, @Notes, @LastModifiedBy, @LastModifiedDate)"

        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, params)

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_ExpenseAdd, "مصروف",
                If(String.IsNullOrWhiteSpace(voucherNumber), "-", voucherNumber),
                $"{description} — المبلغ: {amount.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)} {currencyType}")
        End If

        Return affected
    End Function

    Public Async Function UpdateExpenseAsync(
        voucherNumber As String, description As String, amount As Decimal,
        currencyType As String, expenseDate As String, category As String, notes As String,
        expenseID As Integer, modifiedBy As String, modifiedDate As String) As Task(Of Integer)

        If Not UserSession.CanEdit Then Return 0

        Dim params As SQLiteParameter() = {
            New SQLiteParameter("@VoucherNumber", UtilityModule.ToDBValue(voucherNumber)),
            New SQLiteParameter("@Description", UtilityModule.ToDBValue(description)),
            New SQLiteParameter("@Amount", amount),
            New SQLiteParameter("@CurrencyType", UtilityModule.ToDBValue(currencyType)),
            New SQLiteParameter("@ExpenseDate", expenseDate),
            New SQLiteParameter("@Category", UtilityModule.ToDBValue(category)),
            New SQLiteParameter("@Notes", UtilityModule.ToDBValue(notes)),
            New SQLiteParameter("@LastModifiedBy", If(String.IsNullOrWhiteSpace(modifiedBy), DBNull.Value, CObj(modifiedBy))),
            New SQLiteParameter("@LastModifiedDate", modifiedDate),
            New SQLiteParameter("@ID", expenseID)
        }

        Dim sql = $"UPDATE Expenses SET VoucherNumber=@VoucherNumber, Description=@Description, Amount=@Amount, " &
                  "CurrencyType=@CurrencyType, ExpenseDate=@ExpenseDate, Category=@Category, Notes=@Notes, " &
                  "LastModifiedBy=@LastModifiedBy, LastModifiedDate=@LastModifiedDate " &
                  "WHERE ID=@ID"

        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, params)

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_ExpenseUpdate, "مصروف",
                If(String.IsNullOrWhiteSpace(voucherNumber), "-", voucherNumber),
                $"المعرف: {expenseID} — {description} — المبلغ: {amount.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)} {currencyType}")
        End If

        Return affected
    End Function

    Public Async Function SearchExpensesAsync(
        voucherNumber As String, description As String,
        currencyType As String, category As String, notes As String,
        ct As Threading.CancellationToken) As Task(Of DataTable)

        Dim conditions As New List(Of String)
        Dim parameters As New List(Of SQLiteParameter)

        If Not String.IsNullOrWhiteSpace(voucherNumber) Then
            conditions.Add("VoucherNumber LIKE @VoucherNumber")
            parameters.Add(New SQLiteParameter("@VoucherNumber", "%" & voucherNumber.Trim() & "%"))
        End If
        If Not String.IsNullOrWhiteSpace(description) Then
            conditions.Add("Description LIKE @Description")
            parameters.Add(New SQLiteParameter("@Description", "%" & description.Trim() & "%"))
        End If
        If Not String.IsNullOrWhiteSpace(currencyType) Then
            conditions.Add("CurrencyType = @CurrencyType")
            parameters.Add(New SQLiteParameter("@CurrencyType", currencyType.Trim()))
        End If
        If Not String.IsNullOrWhiteSpace(category) Then
            conditions.Add("Category = @Category")
            parameters.Add(New SQLiteParameter("@Category", category.Trim()))
        End If
        If Not String.IsNullOrWhiteSpace(notes) Then
            conditions.Add("Notes LIKE @Notes")
            parameters.Add(New SQLiteParameter("@Notes", "%" & notes.Trim() & "%"))
        End If

        Dim sql As String = "SELECT ID, VoucherNumber, Description, Amount, CurrencyType, ExpenseDate, Category, Notes, LastModifiedBy, LastModifiedDate FROM Expenses"
        If conditions.Count > 0 Then sql &= " WHERE " & String.Join(" AND ", conditions)
        sql &= " ORDER BY ExpenseDate DESC"

        Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, ct, parameters.ToArray())
        UtilityModule.FormatDateColumnsForDisplay(dt, AppConstants.Col_Exp_ExpenseDate)
        Return dt
    End Function

    ''' <summary>التحقق من وجود رقم السند مسبقاً (للإدخال اليدوي المفرد — السند الشامل يدخل من النافذة الخاصة)</summary>
    ''' <remarks>🌟 إصلاح: المقارنة الآن بـ TRIM لتطابق أسلوب PaymentService — كانت فراغات بداية/نهاية
    ''' رقم السند تُخفي التكرار (" 123 " لا تطابق "123") فتسلّل السندات المكررة بصمت</remarks>
    Public Async Function IsVoucherNumberExistsAsync(voucherNumber As String, Optional excludeVoucher As String = "") As Task(Of Boolean)
        Try
            If String.IsNullOrWhiteSpace(voucherNumber) Then Return False

            If String.IsNullOrEmpty(excludeVoucher) Then
                Dim sql As String = $"SELECT COUNT(*) FROM {AppConstants.Table_Expenses} WHERE TRIM({AppConstants.Col_Exp_VoucherNumber}) = @VoucherNumber"
                Dim param As SQLiteParameter = New SQLiteParameter("@VoucherNumber", voucherNumber.Trim())
                Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(sql, param)
                Return Convert.ToInt32(result) > 0
            Else
                Dim sql As String = $"SELECT COUNT(*) FROM {AppConstants.Table_Expenses} WHERE TRIM({AppConstants.Col_Exp_VoucherNumber}) = @VoucherNumber AND TRIM({AppConstants.Col_Exp_VoucherNumber}) <> @ExcludeVoucher"
                Dim params As SQLiteParameter() = {
                    New SQLiteParameter("@VoucherNumber", voucherNumber.Trim()),
                    New SQLiteParameter("@ExcludeVoucher", excludeVoucher.Trim())
                }
                Dim result As Object = Await DatabaseModule.ExecuteScalarAsync(sql, params)
                Return Convert.ToInt32(result) > 0
            End If
        Catch ex As Exception
            DatabaseModule.LogError("ExpenseService.IsVoucherNumberExistsAsync", ex)
            Throw
        End Try
    End Function

    ''' <summary>حذف مصروف بالمعرف الفريد ID (آمن حتى لو كان رقم السند مشتركاً بين أسطر)</summary>
    Public Async Function DeleteExpenseAsync(expenseId As Integer) As Task(Of Integer)
        If Not UserSession.CanDelete Then Return 0

        Dim sql As String = $"DELETE FROM {AppConstants.Table_Expenses} WHERE {AppConstants.Col_Exp_ID} = @id"
        Dim affected As Integer = Await DatabaseModule.ExecuteNonQueryAsync(sql, New SQLiteParameter("@id", expenseId))

        ' 🌟 [أولوية 4] توثيق العملية في سجل التدقيق
        If affected > 0 Then
            Dim audit As New AuditService()
            Await audit.LogAsync(AuditService.Act_ExpenseDelete, "مصروف",
                expenseId.ToString(System.Globalization.CultureInfo.InvariantCulture), "")
        End If

        Return affected
    End Function
End Class
