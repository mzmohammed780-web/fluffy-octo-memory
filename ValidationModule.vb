Option Explicit On
Option Strict On

Imports System.Linq

Public Module ValidationModule

    ''' <summary>التحقق من صحة رقم الهوية باستخدام خوارزمية Luhn</summary>
    Public Function IsLuhnValid(number As String) As Boolean
        If String.IsNullOrEmpty(number) Then Return False
        Try
            Dim sum As Integer = 0
            Dim alternate As Boolean = False
            For i As Integer = number.Length - 1 To 0 Step -1
                Dim digit As Integer = Asc(number(i)) - Asc("0"c)
                If alternate Then
                    digit *= 2
                    If digit > 9 Then digit -= 9
                End If
                sum += digit
                alternate = Not alternate
            Next
            Return (sum Mod 10 = 0)
        Catch ex As Exception
            DatabaseModule.LogError("IsLuhnValid", ex)
            Return False
        End Try
    End Function

    ''' <summary>التحقق من رقم الهوية وتنسيقه (يعيد رسالة الخطأ إن وجدت)</summary>
    Public Function IsIdValid(id As String, ByRef errorMessage As String) As Boolean
        If String.IsNullOrWhiteSpace(id) Then
            errorMessage = "رقم الهوية مطلوب!"
            Return False
        End If

        Dim clean As String = New String(id.Where(AddressOf Char.IsDigit).ToArray())
        If clean.Length <> 9 Then
            errorMessage = "يجب أن يكون 9 أرقام!"
            Return False
        End If

        If Not IsLuhnValid(clean) Then
            errorMessage = "رقم الهوية غير صحيح (تحقق من الأرقام)"
            Return False
        End If

        Return True
    End Function

    ''' <summary>التحقق من تنسيق رقم الهوية فقط دون فحص Luhn (L-01 — للحقول الثانوية)</summary>
    ''' <remarks>هويات قديمة صحيحة قد لا تجتاز خوارزمية Luhn؛ لذا التنسيق إجباري هنا،
    ''' وفحص Luhn يُدار من شاشة اللاعبين كتحذير تأكيد اختياري لهوية الأب والزوجة/الأم</remarks>
    Public Function IsIdFormatValid(id As String, ByRef errorMessage As String) As Boolean
        If String.IsNullOrWhiteSpace(id) Then
            errorMessage = "رقم الهوية مطلوب!"
            Return False
        End If

        Dim clean As String = New String(id.Where(AddressOf Char.IsDigit).ToArray())
        If clean.Length <> 9 Then
            errorMessage = "يجب أن يكون 9 أرقام!"
            Return False
        End If

        Return True
    End Function

    ''' <summary>التحقق من رقم الهاتف (يعيد رسالة الخطأ إن وجدت)</summary>
    ''' <remarks>🌟 إصلاح: كان الوصول إلى phone بدون فحص Nothing يرمي NullReferenceException
    ''' بدلاً من إرجاع نتيجة تحقق</remarks>
    Public Function IsPhoneValid(phone As String, ByRef errorMessage As String) As Boolean
        If String.IsNullOrWhiteSpace(phone) Then
            errorMessage = "رقم الهاتف يجب أن يكون 10 أرقام!"
            Return False
        End If
        Dim clean As String = New String(phone.Where(AddressOf Char.IsDigit).ToArray())
        If clean.Length <> 10 Then
            errorMessage = "رقم الهاتف يجب أن يكون 10 أرقام!"
            Return False
        End If
        Return True
    End Function
    ''' <summary>التحقق من الاسم وتاريخ الميلاد قبل الحفظ (يعيد رسالة الخطأ إن وجدت)</summary>
    Public Function IsPlayerBasicDataValid(playerName As String, birthDate As Date?, ByRef errorMessage As String) As Boolean
        ' 1. التحقق من الاسم
        If String.IsNullOrWhiteSpace(playerName) Then
            errorMessage = "الرجاء إدخال اسم اللاعب"
            Return False
        End If

        ' 2. التحقق من تاريخ الميلاد (إن كان قد تم تحديده)
        If birthDate.HasValue Then
            If birthDate.Value > Date.Now Then
                errorMessage = "تاريخ الميلاد لا يمكن أن يكون في المستقبل"
                Return False
            End If

            If birthDate.Value > Date.Now.AddYears(-5) Then
                errorMessage = "عمر اللاعب يجب أن يكون 5 سنوات على الأقل"
                Return False
            End If
        End If

        Return True
    End Function
    ''' <summary>التحقق من صحة بيانات الدفعة الأساسية قبل الحفظ (يعيد رسالة الخطأ إن وجدت)</summary>
    Public Function IsPaymentBasicDataValid(
        voucherNumber As String, amountStr As String,
        playerName As String, transferorName As String,
        paymentMethod As String, currencyType As String,
        ByRef errorMessage As String) As Boolean

        ' 1. التحقق من رقم السند
        If String.IsNullOrWhiteSpace(voucherNumber) Then
            errorMessage = "الرجاء إدخال رقم السند"
            Return False
        End If

        ' 2. التحقق من المبلغ
        If String.IsNullOrWhiteSpace(amountStr) Then
            errorMessage = "الرجاء إدخال المبلغ"
            Return False
        End If

        Dim amountValue As Decimal
        ' 🌟 تحليل بثقافة ثابتة — كان بثقافة النظام فيتعارض مع الحفظ النهائي بـ InvariantCulture
        If Not Decimal.TryParse(amountStr, Globalization.NumberStyles.Number, Globalization.CultureInfo.InvariantCulture, amountValue) OrElse amountValue <= 0 Then
            errorMessage = "الرجاء إدخال مبلغ صحيح أكبر من صفر"
            Return False
        End If

        ' 3. التحقق من وجود لاعب أو محول (واحد على الأقل)
        Dim hasPlayer As Boolean = Not String.IsNullOrWhiteSpace(playerName) AndAlso playerName <> "-- اختر اللاعب --"
        Dim hasTransferor As Boolean = Not String.IsNullOrWhiteSpace(transferorName)
        If Not hasPlayer AndAlso Not hasTransferor Then
            errorMessage = "الرجاء إدخال اسم اللاعب أو اسم المحول (واحد على الأقل)"
            Return False
        End If

        ' 4. التحقق من طريقة الدفع
        If String.IsNullOrWhiteSpace(paymentMethod) Then
            errorMessage = "الرجاء اختيار طريقة الدفع"
            Return False
        End If

        ' 5. التحقق من نوع العملة
        If String.IsNullOrWhiteSpace(currencyType) Then
            errorMessage = "الرجاء اختيار نوع العملة"
            Return False
        End If

        Return True
    End Function
    ''' <summary>التحقق من بيانات المصروف الأساسية قبل الحفظ (يعيد رسالة الخطأ إن وجدت)</summary>
    Public Function IsExpenseDataValid(
        description As String, amountStr As String, voucherNumber As String,
        ByRef errorMessage As String) As Boolean

        ' 1. التحقق من بيان المصروف
        If String.IsNullOrWhiteSpace(description) Then
            errorMessage = "الرجاء إدخال بيان المصروف"
            Return False
        End If

        ' 2. التحقق من المبلغ
        If String.IsNullOrWhiteSpace(amountStr) Then
            errorMessage = "الرجاء إدخال المبلغ"
            Return False
        End If

        Dim amountValue As Decimal
        ' 🌟 تحليل بثقافة ثابتة — كان بثقافة النظام فيتعارض مع الحفظ النهائي بـ InvariantCulture
        If Not Decimal.TryParse(amountStr, Globalization.NumberStyles.Number, Globalization.CultureInfo.InvariantCulture, amountValue) OrElse amountValue <= 0 Then
            errorMessage = "الرجاء إدخال مبلغ صحيح أكبر من صفر"
            Return False
        End If

        ' 3. التحقق من رقم السند
        If String.IsNullOrWhiteSpace(voucherNumber) Then
            errorMessage = "الرجاء إدخال رقم السند"
            Return False
        End If

        Return True
    End Function
End Module
