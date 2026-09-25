Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Module UtilityModule

    Private Const APP_SALT As String = "Academy_Secured_Salt_2026#@!"

    ''' <summary>تحويل آمن للقيمة إلى نص</summary>
    Public Function SafeString(value As Object) As String
        If value Is Nothing OrElse IsDBNull(value) Then Return ""
        Return value.ToString().Trim()
    End Function

    ''' <summary>تحويل آمن للتاريخ. يعيد Date.MinValue إذا كان فارغاً أو غير صالح</summary>
    Public Function SafeDate(value As Object) As Date
        If value Is Nothing OrElse IsDBNull(value) Then Return Date.MinValue
        Try
            Dim s As String = value.ToString().Trim()
            If String.IsNullOrWhiteSpace(s) Then Return Date.MinValue

            If s.Length >= 10 AndAlso s(4) = "-"c AndAlso s(7) = "-"c Then
                Dim result As Date
                If Date.TryParseExact(s.Substring(0, 10), "yyyy-MM-dd", Nothing, Globalization.DateTimeStyles.None, result) Then
                    Return result
                End If
            End If
            If s.Length = 10 AndAlso s(2) = "/"c AndAlso s(5) = "/"c Then
                Dim dmy As Date
                If Date.TryParseExact(s, "dd/MM/yyyy", Nothing, Globalization.DateTimeStyles.None, dmy) Then
                    Return dmy
                End If
            End If
            Dim dtResult As Date
            If Date.TryParse(s, dtResult) Then
                Return dtResult
            End If

            Return Date.MinValue
        Catch
            Return Date.MinValue
        End Try
    End Function

    ''' <summary>تنسيق قيمة لقاعدة البيانات مع دعم القيم الرقمية والنصية</summary>
    Public Function ToDBValue(value As Object, Optional isNumeric As Boolean = False) As Object
        If value Is Nothing OrElse String.IsNullOrWhiteSpace(value.ToString()) Then Return DBNull.Value
        If isNumeric Then
            Dim n As Long
            Return If(Long.TryParse(value.ToString(), n), CObj(n), DBNull.Value)
        End If
        Return value.ToString().Trim()
    End Function

    ''' <summary>تهريب النصوص لملفات CSV + حماية من حقن معادلات Excel (CSV Injection)</summary>
    ''' <remarks>🌟 إصلاح v2: الآن كل ما يبدأ بـ = + @ Tab يُهرّب، وكل ما يبدأ بـ - يُفحص:
    '''   * إن كان رقمياً صحيحاً (بما فيه الكسور) → يُترك رقماً (Excel يقرأه كرقم)
    '''   * إن كان نصاً (مثل "-cmd|...") → يُسبق بفاصلة عليا
    ''' كما يحافظ على تهريب الفواصل وعلامات التنصيص وسطور الأسطر</remarks>
    Public Function EscapeCsvValue(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""

        Dim firstChar As Char = value(0)
        Dim needsProtection As Boolean = False

        If firstChar = "="c OrElse firstChar = "+"c OrElse firstChar = "@"c OrElse firstChar = ControlChars.Tab Then
            needsProtection = True
        ElseIf firstChar = "-"c Then
            ' رقم سالب سليم فقط إذا كان يتكون من رقم/نقطة عشرية/فاصلة آلاف
            Dim asNumber As Decimal
            If Not Decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, asNumber) Then
                needsProtection = True
            End If
        End If

        If needsProtection Then
            value = "'" & value
        End If

        If value.Contains(",") OrElse value.Contains("""") OrElse
           value.Contains(vbCrLf) OrElse value.Contains(vbLf) OrElse value.Contains(vbCr) Then
            Return """" & value.Replace("""", """""") & """"
        End If
        Return value
    End Function

    ' =========================================================================
    ' 🔐 دوال التشفير والحماية
    ' =========================================================================

    Private Const PBKDF2_ITERATIONS_V2 As Integer = 100000
    Private Const HASH_LENGTH As Integer = 32

    Public Function GenerateSalt(Optional size As Integer = 16) As String
        Dim saltBytes(size - 1) As Byte
        Using rng As RandomNumberGenerator = RandomNumberGenerator.Create()
            rng.GetBytes(saltBytes)
        End Using
        Return Convert.ToBase64String(saltBytes)
    End Function

    Public Function GetHashFormatVersion(storedHash As String) As Integer
        If String.IsNullOrWhiteSpace(storedHash) Then Return 0
        If storedHash.StartsWith("v2|") Then Return 2
        Return 0
    End Function

    Public Function HashPassword(password As String, Optional salt As String = "") As String
        If String.IsNullOrWhiteSpace(password) Then Return ""

        Dim saltBytes As Byte()
        If String.IsNullOrWhiteSpace(salt) Then
            saltBytes = New Byte(15) {}
            Using rng As RandomNumberGenerator = RandomNumberGenerator.Create()
                rng.GetBytes(saltBytes)
            End Using
        Else
            Try
                saltBytes = Convert.FromBase64String(salt)
            Catch
                saltBytes = Encoding.UTF8.GetBytes(salt)
            End Try
        End If

        Using pbkdf2 As New Rfc2898DeriveBytes(password & APP_SALT, saltBytes, PBKDF2_ITERATIONS_V2)
            Dim hashBytes As Byte() = pbkdf2.GetBytes(HASH_LENGTH)
            Return "v2|" & Convert.ToBase64String(saltBytes) & "|" & Convert.ToBase64String(hashBytes)
        End Using
    End Function

    Public Function VerifyPassword(enteredPassword As String, storedHash As String) As Boolean
        If String.IsNullOrWhiteSpace(enteredPassword) OrElse String.IsNullOrWhiteSpace(storedHash) Then Return False

        Try
            If storedHash.StartsWith("v2|") Then
                Dim parts As String() = storedHash.Split("|"c)
                If parts.Length <> 3 Then Return False

                Dim saltBytes As Byte() = Convert.FromBase64String(parts(1))
                Dim storedHashBytes As Byte() = Convert.FromBase64String(parts(2))
                If saltBytes Is Nothing OrElse saltBytes.Length = 0 OrElse
                   storedHashBytes Is Nothing OrElse storedHashBytes.Length = 0 Then Return False

                Using pbkdf2 As New Rfc2898DeriveBytes(enteredPassword & APP_SALT, saltBytes, PBKDF2_ITERATIONS_V2)
                    Dim hashBytes As Byte() = pbkdf2.GetBytes(storedHashBytes.Length)
                    Dim areEqual As Boolean = True
                    For i As Integer = 0 To storedHashBytes.Length - 1
                        If hashBytes(i) <> storedHashBytes(i) Then areEqual = False
                    Next
                    Return areEqual
                End Using
            End If

            Dim combinedBytes As Byte() = Convert.FromBase64String(storedHash)
            If combinedBytes.Length >= 36 Then
                Dim legacySalt(15) As Byte
                Array.Copy(combinedBytes, 0, legacySalt, 0, 16)

                Using pbkdf2 As New Rfc2898DeriveBytes(enteredPassword & APP_SALT, legacySalt, 10000)
                    Dim hashBytes As Byte() = pbkdf2.GetBytes(20)
                    Dim areEqual As Boolean = True
                    For i As Integer = 0 To 19
                        If combinedBytes(i + 16) <> hashBytes(i) Then areEqual = False
                    Next
                    Return areEqual
                End Using
            End If
        Catch
            Return False
        End Try

        Return False
    End Function

    ' =========================================================================
    ' 🖼️ معالجة الصور
    ' =========================================================================

    Public Function ImageToDBValue(img As Image) As Object
        If img Is Nothing Then Return DBNull.Value
        Try
            Dim bytes As Byte() = CompressAndResizeImage(img, 500, 80L)
            If bytes IsNot Nothing AndAlso bytes.Length > 0 Then
                Return DirectCast(bytes, Object)
            End If
            Return DBNull.Value
        Catch ex As Exception
            DatabaseModule.LogError("UtilityModule.ImageToDBValue", ex)
            Return DBNull.Value
        End Try
    End Function

    Public Function ByteArrayToImage(bytes As Byte()) As Image
        If bytes Is Nothing OrElse bytes.Length = 0 Then Return Nothing
        Try
            Using ms As New MemoryStream(bytes)
                Return New Bitmap(Image.FromStream(ms))
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("UtilityModule.ByteArrayToImage", ex)
            Return Nothing
        End Try
    End Function

    Public Function CompressAndResizeImage(originalImage As Image, Optional maxDimension As Integer = 500, Optional quality As Long = 80L) As Byte()
        If originalImage Is Nothing Then Return Nothing

        Try
            Dim origW As Integer = originalImage.Width
            Dim origH As Integer = originalImage.Height
            Dim newW As Integer = origW
            Dim newH As Integer = origH

            If origW > maxDimension OrElse origH > maxDimension Then
                If origW > origH Then
                    newW = maxDimension
                    newH = Math.Max(1, CInt((origH * maxDimension) / origW))
                Else
                    newH = maxDimension
                    newW = Math.Max(1, CInt((origW * maxDimension) / origH))
                End If
            End If

            Using newBmp As New Bitmap(newW, newH)
                Using g As Graphics = Graphics.FromImage(newBmp)
                    g.CompositingQuality = CompositingQuality.HighQuality
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic
                    g.SmoothingMode = SmoothingMode.HighQuality
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality
                    g.DrawImage(originalImage, 0, 0, newW, newH)
                End Using

                Using ms As New MemoryStream()
                    Dim jpgEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)
                    If jpgEncoder IsNot Nothing Then
                        Using encoderParams As New EncoderParameters(1)
                            Using encParam As New EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality)
                                encoderParams.Param(0) = encParam
                                newBmp.Save(ms, jpgEncoder, encoderParams)
                            End Using
                        End Using
                    Else
                        newBmp.Save(ms, ImageFormat.Jpeg)
                    End If
                    Return ms.ToArray()
                End Using
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("CompressAndResizeImage", ex)
            Return Nothing
        End Try
    End Function

    Private Function GetEncoder(format As ImageFormat) As ImageCodecInfo
        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
        For Each codec In codecs
            If codec.FormatID = format.Guid Then Return codec
        Next
        Return Nothing
    End Function

    Public Sub FormatDateColumnsForDisplay(dt As DataTable, ParamArray columnNames As String())
        If dt Is Nothing OrElse columnNames Is Nothing OrElse columnNames.Length = 0 Then Return
        Try
            For Each row As DataRow In dt.Rows
                For Each colName In columnNames
                    If dt.Columns.Contains(colName) AndAlso Not IsDBNull(row(colName)) Then
                        Dim s As String = row(colName).ToString().Trim()
                        If s.Length >= 10 AndAlso s(4) = "-"c AndAlso s(7) = "-"c Then
                            Dim d As Date
                            If Date.TryParseExact(s.Substring(0, 10), "yyyy-MM-dd", Nothing, Globalization.DateTimeStyles.None, d) Then
                                row(colName) = d.ToString("dd/MM/yyyy")
                            End If
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            DatabaseModule.LogError("UtilityModule.FormatDateColumnsForDisplay", ex)
        End Try
    End Sub

    Public Function FormatDisplayDateSafe(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return value
        Dim d As Date
        If Date.TryParseExact(value.Trim(), "dd/MM/yyyy", Nothing, Globalization.DateTimeStyles.None, d) Then
            Return d.ToString("dd/MM/yyyy")
        End If
        Return value
    End Function

    ' =========================================================================
    ' 🌟 توحيد الكود المكرر: تصدير CSV، مجاميع العملات، السندات الشاملة
    ' =========================================================================

    Public Function ToDecimalSafe(value As Object) As Decimal
        If value Is Nothing OrElse IsDBNull(value) Then Return 0D
        If TypeOf value Is Decimal Then Return CDec(value)
        If TypeOf value Is Double OrElse TypeOf value Is Single OrElse TypeOf value Is Long OrElse
           TypeOf value Is Integer OrElse TypeOf value Is Short OrElse TypeOf value Is Byte Then
            Return Convert.ToDecimal(value)
        End If
        Dim s As String = value.ToString().Trim()
        If s = "" Then Return 0D
        Dim d As Decimal
        Decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, d)
        Return d
    End Function

    Public Structure CurrencyTotals
        Public ILS As Decimal
        Public USD As Decimal
        Public JOD As Decimal
        Public Count As Integer
    End Structure

    Public Function SumCurrencyTotalsFromRows(rows As IEnumerable(Of DataGridViewRow),
                                              amountColumn As String,
                                              currencyColumn As String) As CurrencyTotals
        Dim t As New CurrencyTotals()
        If rows Is Nothing Then Return t
        For Each row As DataGridViewRow In rows
            If row Is Nothing OrElse row.IsNewRow Then Continue For
            Dim cell As DataGridViewCell = row.Cells(amountColumn)
            If cell.Value Is Nothing OrElse IsDBNull(cell.Value) Then Continue For
            Dim amount As Decimal = ToDecimalSafe(cell.Value)
            Select Case SafeString(row.Cells(currencyColumn).Value)
                Case AppConstants.Currency_USD : t.USD += amount
                Case AppConstants.Currency_JOD : t.JOD += amount
                Case Else : t.ILS += amount
            End Select
            t.Count += 1
        Next
        Return t
    End Function

    Public Function SumCurrencyTotalsFromGrid(grid As DataGridView,
                                              amountColumn As String,
                                              currencyColumn As String) As CurrencyTotals
        Dim t As New CurrencyTotals()
        If grid Is Nothing OrElse Not grid.Columns.Contains(amountColumn) OrElse
           Not grid.Columns.Contains(currencyColumn) Then Return t
        Dim allRows As New List(Of DataGridViewRow)
        For Each row As DataGridViewRow In grid.Rows
            allRows.Add(row)
        Next
        Return SumCurrencyTotalsFromRows(allRows, amountColumn, currencyColumn)
    End Function

    Public Function SumCurrencyTotalsFromTable(dt As DataTable,
                                               amountColumn As String,
                                               currencyColumn As String) As CurrencyTotals
        Dim t As New CurrencyTotals()
        If dt Is Nothing OrElse Not dt.Columns.Contains(amountColumn) OrElse
           Not dt.Columns.Contains(currencyColumn) Then Return t
        For Each row As DataRow In dt.Rows
            If IsDBNull(row(amountColumn)) Then Continue For
            Dim amount As Decimal = ToDecimalSafe(row(amountColumn))
            Select Case SafeString(row(currencyColumn))
                Case AppConstants.Currency_USD : t.USD += amount
                Case AppConstants.Currency_JOD : t.JOD += amount
                Case Else : t.ILS += amount
            End Select
            t.Count += 1
        Next
        Return t
    End Function

    Public Function FormatCurrencyPart(amount As Decimal, currency As String) As String
        Select Case currency
            Case AppConstants.Currency_USD : Return $"دولار: {amount:#,##0.##} $"
            Case AppConstants.Currency_JOD : Return $"دينار: {amount:#,##0.##} د.أ"
            Case Else : Return $"شيكل: {amount:#,##0.##} ₪"
        End Select
    End Function

    Public Sub AmountKeyPressFilter(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "."c AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
        If e.KeyChar = "."c AndAlso DirectCast(sender, TextBox).Text.Contains("."c) Then
            e.Handled = True
        End If
    End Sub

    Public Function ParseDeclaredTotal(currency As String,
                                       rawILS As String, rawUSD As String, rawJOD As String,
                                       ByRef parsed As Decimal) As Boolean
        Dim raw As String
        Select Case currency
            Case AppConstants.Currency_ILS : raw = rawILS
            Case AppConstants.Currency_USD : raw = rawUSD
            Case AppConstants.Currency_JOD : raw = rawJOD
            Case Else : parsed = 0D : Return False
        End Select

        If String.IsNullOrWhiteSpace(raw) Then
            parsed = 0D
            Return False
        End If
        Return Decimal.TryParse(raw, NumberStyles.Number, CultureInfo.InvariantCulture, parsed)
    End Function

    ''' <summary>توحيد تصدير CSV — الكتابة تشمل تهريب العناوين الآن (إصلاح CSV Injection في الترويسات)</summary>
    Public Sub WriteGridCsv(grid As DataGridView,
                            writer As StreamWriter,
                            excludedColumnNames As IEnumerable(Of String),
                            Optional rowsToExport As IEnumerable(Of DataGridViewRow) = Nothing,
                            Optional dateColumnNames As IEnumerable(Of String) = Nothing,
                            Optional amountColumnNames As IEnumerable(Of String) = Nothing)
        If grid Is Nothing OrElse writer Is Nothing Then Return

        Dim excl As New HashSet(Of String)(If(excludedColumnNames, New String() {}), StringComparer.OrdinalIgnoreCase)
        Dim dateCols As New HashSet(Of String)(If(dateColumnNames, New String() {}), StringComparer.OrdinalIgnoreCase)
        Dim amtCols As New HashSet(Of String)(If(amountColumnNames, New String() {}), StringComparer.OrdinalIgnoreCase)

        ' 🌟 إصلاح: صف العناوين يُهرّب الآن — كان خاماً (CSV Injection محتمل من ترويسة عمود معدّلة)
        Dim headers As New List(Of String)
        For Each col As DataGridViewColumn In grid.Columns
            If col.Visible AndAlso Not excl.Contains(col.Name) Then
                headers.Add(EscapeCsvValue(col.HeaderText))
            End If
        Next
        writer.WriteLine(String.Join(",", headers))

        Dim rows As New List(Of DataGridViewRow)
        If rowsToExport IsNot Nothing Then
            For Each r As DataGridViewRow In rowsToExport
                rows.Add(r)
            Next
        Else
            For Each r As DataGridViewRow In grid.Rows
                rows.Add(r)
            Next
        End If

        For Each row As DataGridViewRow In rows
            If row.IsNewRow Then Continue For
            Dim values As New List(Of String)
            For Each col As DataGridViewColumn In grid.Columns
                If col.Visible AndAlso Not excl.Contains(col.Name) Then
                    values.Add(EscapeCsvValue(FormatGridCellForCsv(row.Cells(col.Name), col, dateCols, amtCols)))
                End If
            Next
            writer.WriteLine(String.Join(",", values))
        Next
    End Sub

    Private Function FormatGridCellForCsv(cell As DataGridViewCell,
                                          col As DataGridViewColumn,
                                          dateCols As HashSet(Of String),
                                          amtCols As HashSet(Of String)) As String
        If cell.Value Is Nothing OrElse IsDBNull(cell.Value) Then Return ""
        Dim v As String = cell.Value.ToString()
        If dateCols.Contains(col.Name) Then
            v = FormatDisplayDateSafe(v)
        ElseIf amtCols.Contains(col.Name) Then
            Dim dec As Decimal
            If Decimal.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, dec) Then
                v = dec.ToString("0.##", CultureInfo.InvariantCulture)
            End If
        End If
        Return v
    End Function
End Module

' ═══════════════════════════════════════════════════════════════════
' 🌟 لوحة مجاميع السند الشاملة المشتركة
' ═══════════════════════════════════════════════════════════════════
Public Class BulkTotalsPanel
    Private ReadOnly _declared As New Dictionary(Of String, TextBox)
    Private ReadOnly _running As New Dictionary(Of String, Label)
    Private ReadOnly _remaining As New Dictionary(Of String, Label)

    Public Sub Register(currency As String, declared As TextBox, running As Label, remaining As Label)
        _declared(currency) = declared
        _running(currency) = running
        _remaining(currency) = remaining
    End Sub

    Public Function ParseDeclared(currency As String, ByRef parsed As Decimal) As Boolean
        Dim txt As TextBox = Nothing
        If Not _declared.TryGetValue(currency, txt) OrElse txt Is Nothing Then
            parsed = 0D
            Return False
        End If
        If String.IsNullOrWhiteSpace(txt.Text) Then
            parsed = 0D
            Return False
        End If
        Return Decimal.TryParse(txt.Text, NumberStyles.Number, CultureInfo.InvariantCulture, parsed)
    End Function

    Public Sub Refresh(getEntered As Func(Of String, Decimal))
        If getEntered Is Nothing Then Return
        For Each cur As String In New String() {AppConstants.Currency_ILS, AppConstants.Currency_USD, AppConstants.Currency_JOD}
            Dim entered As Decimal = getEntered(cur)

            Dim valLabel As Label = Nothing
            Dim remLabel As Label = Nothing
            _running.TryGetValue(cur, valLabel)
            _remaining.TryGetValue(cur, remLabel)
            If valLabel Is Nothing OrElse remLabel Is Nothing Then Continue For

            valLabel.Text = $"{entered:#,##0.##}"

            Dim declared As Decimal
            If ParseDeclared(cur, declared) AndAlso declared > 0 Then
                Dim remaining As Decimal = declared - entered
                remLabel.Text = $"{remaining:#,##0.##}"

                If remaining = 0 AndAlso entered > 0 Then
                    valLabel.ForeColor = Color.Green
                    remLabel.ForeColor = Color.Green
                ElseIf remaining > 0 Then
                    valLabel.ForeColor = Color.FromArgb(0, 90, 160)
                    remLabel.ForeColor = Color.FromArgb(0, 90, 160)
                Else
                    valLabel.ForeColor = Color.Red
                    remLabel.ForeColor = Color.Red
                End If
            Else
                remLabel.Text = "—"
                valLabel.ForeColor = Color.Black
                remLabel.ForeColor = Color.Black
            End If
        Next
    End Sub
End Class