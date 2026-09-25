Option Explicit On
Option Strict On

Imports Microsoft.Win32
Imports System.IO
Imports System.Globalization ' 🌟 لضبط تحليل التواريخ مستقل عن إعدادات الويندوز
Imports System.Security.Cryptography
Imports System.Text

Public Module LicenseManager

    Private Const RegPath As String = "SOFTWARE\Academy2026_Secured"
    Private Const Key_Data As String = "LicenseToken"
    Private Const ValidityDays As Integer = 365
    ' 🌟 مهلة السماح لتراجع ساعة الجهاز (يومين) — قفزات الساعة/التوقيت/بطارية BIOS
    ' كانت تقفل الترخيص نهائياً بدون رجعة لأن today أصبح أصغر من lastRunDate للأبد
    Private Const RollbackGraceDays As Integer = 2

    ' ═══════════════════════════════════════════════════════════════════
    ' 🌟 M-09 — سياسة حماية الترخيص (قرار موثق):
    '    * الحماية من جهة العميل مقصودة وكافية لنطاق التوزيع الداخلي الحالي.
    '    * المفاتيح مشتقة من قيم الجهاز — تخمينها ممكن لمن لديه نية ودقة، وهذا مقبول.
    '    * إذا تغير نمط التوزيع مستقبلاً (بيع خارجي / توزيع واسع):
    '      يُضاف توقيع خادمي خفيف عند التفعيل — HMAC على معرف الجهاز بمفتاح خادمي،
    '      ويتحقق العميل من التوقيع بدل الاشتقاق المحلي.
    '    * ResetLicense لم تعد Public — لا يستدعيها أي كود (الاستدعاء الوحيد في frmLogin مُعطَّل).
    ' ═══════════════════════════════════════════════════════════════════

    Private ReadOnly LicenseFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Academy2026_Secured",
        "syscfg.dat")

    ''' <summary>تصفير الترخيص — أداة دعم داخلية (Private): للاستخدام المؤقت من كود التطوير فقط.
    ''' إذا احتجت تصفيراً يدوياً مؤقتاً: غيّر Private إلى Public واستدعها من frmLogin ثم أعِدها Private</summary>
    Private Sub ResetLicense()
        Try
            ' 1. حذف الريجستري نهائياً
            Registry.CurrentUser.DeleteSubKeyTree(RegPath, False)

            ' 2. حذف الملف نهائياً
            If File.Exists(LicenseFilePath) Then
                Dim fi As New FileInfo(LicenseFilePath)
                fi.Attributes = FileAttributes.Normal ' إزالة الحماية قبل الحذف
                File.Delete(LicenseFilePath)
            End If
        Catch ex As Exception
            DatabaseModule.LogError("LicenseManager.ResetLicense", ex)
        End Try
    End Sub

    Public Function IsExpired() As Boolean
        Try
            Dim hwId As String = GetHardwareId()
            Dim today As DateTime = DateTime.Today

            Dim regToken As String = Nothing
            Dim fileToken As String = Nothing

            Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey(RegPath)
                regToken = regKey?.GetValue(Key_Data)?.ToString()
            End Using

            If File.Exists(LicenseFilePath) Then
                fileToken = File.ReadAllText(LicenseFilePath)
            End If

            ' 3. أول تشغيل للبرنامج (كلاهما فارغ)
            If String.IsNullOrEmpty(regToken) AndAlso String.IsNullOrEmpty(fileToken) Then
                SaveLicenseData(hwId, today, today)
                Return False
            End If

            ' 4. المزامنة الذكية
            If String.IsNullOrEmpty(regToken) AndAlso Not String.IsNullOrEmpty(fileToken) Then
                regToken = fileToken
                SaveTokenToRegistry(regToken)
            ElseIf Not String.IsNullOrEmpty(regToken) AndAlso String.IsNullOrEmpty(fileToken) Then
                fileToken = regToken
                SaveTokenToFile(regToken)
            End If

            ' 5. فحص التطابق — 🌟 إصلاح: عدم التطابق لم يعد يقفل المستخدم نهائياً
            ' كان فشل حفظ أحد المخزنين (مثلاً برنامج حماية قفل الملف) يترك توكِناً قديماً
            ' فيقفل المستخدم الشرعي للأبد. الآن: ننقذ توكِناً صالحاً ونعيد مزامنة المخزنين
            If regToken <> fileToken Then
                Dim savedParts As String() = DecryptLicenseToken(regToken, hwId)
                Dim savedToken As String = regToken

                If savedParts Is Nothing OrElse savedParts.Length < 2 Then
                    savedParts = DecryptLicenseToken(fileToken, hwId)
                    savedToken = fileToken
                End If

                If savedParts Is Nothing OrElse savedParts.Length < 2 Then
                    Return True ' كلا النسختين تالفتان فعلاً — لا يوجد ما يُنقذ
                End If

                regToken = savedToken
                fileToken = savedToken
                SaveTokenToRegistry(regToken)
                SaveTokenToFile(regToken)
            End If

            ' 6. فك التشفير
            Dim parts As String() = DecryptLicenseToken(regToken, hwId)
            If parts Is Nothing OrElse parts.Length < 2 Then
                Return True
            End If

            ' 🌟 ثقافة ثابتة لتحليل التواريخ (كانت تعتمد على إعدادات الجهاز)
            Dim firstRunDate As DateTime = DateTime.ParseExact(parts(0), "yyyy-MM-dd", CultureInfo.InvariantCulture)
            Dim lastRunDate As DateTime = DateTime.ParseExact(parts(1), "yyyy-MM-dd", CultureInfo.InvariantCulture)

            ' 7. فحص التلاعب بالتاريخ — 🌟 مع مهلة سماح يومين لقفزات الساعة العادية
            If today < lastRunDate Then
                If lastRunDate.Subtract(today).Days > RollbackGraceDays Then
                    Return True ' تراجع كبير = محاولة تلاعب حقيقية
                End If
                ' تراجع بسيط: نبقي التاريخ الأحدث المخزّن (لا نحدّث lastRun) ونجري فحص الانتهاء فقط
                Dim expDateGrace As DateTime = firstRunDate.AddDays(ValidityDays)
                Return today > expDateGrace
            End If

            ' تحديث آخر تشغيل
            SaveLicenseData(hwId, firstRunDate, today)

            ' 8. فحص انتهاء المدة
            Dim expirationDate As DateTime = firstRunDate.AddDays(ValidityDays)
            If today > expirationDate Then
                Return True
            End If

            Return False
        Catch ex As Exception
            DatabaseModule.LogError("LicenseManager.IsExpired", ex)
            Return True
        End Try
    End Function

    Public Function GetRemainingDays() As Integer
        Try
            Dim hwId As String = GetHardwareId()
            Dim token As String = Nothing

            ' 🌟 إصلاح: ترتيب المصادر كان معاكساً لـ IsExpired (ملف أولاً هنا، ريجستري أولاً هناك) —
            ' تلف syscfg.dat كان يعطي "0 يوم" بينما IsExpired تقول سليم. الآن نفس الترتيب في الموضعين
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(RegPath)
                token = key?.GetValue(Key_Data)?.ToString()
            End Using
            If String.IsNullOrEmpty(token) AndAlso File.Exists(LicenseFilePath) Then
                token = File.ReadAllText(LicenseFilePath)
            End If

            If String.IsNullOrEmpty(token) Then Return ValidityDays

            Dim parts As String() = DecryptLicenseToken(token, hwId)
            If parts Is Nothing OrElse parts.Length < 2 Then Return 0

            ' 🌟 ثقافة ثابتة لتحليل التواريخ
            Dim firstRunDate As DateTime = DateTime.ParseExact(parts(0), "yyyy-MM-dd", CultureInfo.InvariantCulture)
            Dim expirationDate As DateTime = firstRunDate.AddDays(ValidityDays)
            Dim remaining As Integer = (expirationDate - DateTime.Today).Days
            Return If(remaining > 0, remaining, 0)
        Catch ex As Exception
            ' 🌟 إصلاح: كان يبتلع الاستثناءات بصمت — تشخيص "ليش يعرض 0 يوم؟" كان مستحيلاً
            DatabaseModule.LogError("LicenseManager.GetRemainingDays", ex)
            Return 0
        End Try
    End Function

    Private Sub EnsureLicenseDirectoryExists()
        Dim dirPath As String = Path.GetDirectoryName(LicenseFilePath)
        If Not Directory.Exists(dirPath) Then
            Directory.CreateDirectory(dirPath)
        End If
    End Sub

    Private Sub SaveTokenToRegistry(token As String)
        Using key As RegistryKey = Registry.CurrentUser.CreateSubKey(RegPath)
            key.SetValue(Key_Data, token)
        End Using
    End Sub

    Private Sub SaveTokenToFile(token As String)
        Try
            EnsureLicenseDirectoryExists()
            If File.Exists(LicenseFilePath) Then
                Dim fiPre As New FileInfo(LicenseFilePath)
                fiPre.Attributes = FileAttributes.Normal
            End If
            File.WriteAllText(LicenseFilePath, token)
            Dim fiPost As New FileInfo(LicenseFilePath)
            fiPost.Attributes = FileAttributes.Hidden
        Catch ex As Exception
            DatabaseModule.LogError("LicenseManager.SaveTokenToFile", ex)
        End Try
    End Sub

    Private Sub SaveLicenseData(hwId As String, firstRun As DateTime, lastRun As DateTime)
        Dim raw As String = $"{firstRun:yyyy-MM-dd}|{lastRun:yyyy-MM-dd}"
        Dim encryptedToken As String = EncryptLicenseToken(raw, hwId)
        SaveTokenToRegistry(encryptedToken)
        SaveTokenToFile(encryptedToken)
    End Sub

    Private Function GetHardwareId() As String
        Try
            Using rk As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                Using cryptoKey As RegistryKey = rk.OpenSubKey("SOFTWARE\Microsoft\Cryptography")
                    Dim guid = cryptoKey?.GetValue("MachineGuid")?.ToString()
                    If Not String.IsNullOrEmpty(guid) Then Return guid
                End Using
            End Using
        Catch
        End Try
        Return Environment.MachineName & "_DefaultID"
    End Function

    Private Function EncryptLicenseToken(plainText As String, keySecret As String) As String
        Using aes As Aes = Aes.Create()
            Dim keyBytes As Byte()
            Using sha As SHA256 = SHA256.Create() ' 🌟 إصلاح: إدارة عمر الكائن (كان بدون Dispose)
                keyBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(keySecret & "L1c3ns3_K3y!"))
            End Using
            aes.Key = keyBytes
            aes.GenerateIV()

            Using ms As New MemoryStream()
                ms.Write(aes.IV, 0, aes.IV.Length)
                Using cs As New CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write)
                    Dim plainBytes As Byte() = Encoding.UTF8.GetBytes(plainText)
                    cs.Write(plainBytes, 0, plainBytes.Length)
                    cs.FlushFinalBlock()
                End Using
                Return Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
    End Function

    Private Function DecryptLicenseToken(cipherText As String, keySecret As String) As String()
        Try
            Dim allBytes As Byte() = Convert.FromBase64String(cipherText)
            Using aes As Aes = Aes.Create()
                Dim keyBytes As Byte()
                Using sha As SHA256 = SHA256.Create() ' 🌟 إصلاح: إدارة عمر الكائن (كان بدون Dispose)
                    keyBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(keySecret & "L1c3ns3_K3y!"))
                End Using
                aes.Key = keyBytes

                Dim iv(15) As Byte
                Array.Copy(allBytes, 0, iv, 0, 16)
                aes.IV = iv

                Using ms As New MemoryStream()
                    Using cs As New CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Write)
                        cs.Write(allBytes, 16, allBytes.Length - 16)
                        cs.FlushFinalBlock()
                    End Using
                    Dim decryptedStr As String = Encoding.UTF8.GetString(ms.ToArray())
                    Return decryptedStr.Split("|"c)
                End Using
            End Using
        Catch
            Return Nothing
        End Try
    End Function

End Module
