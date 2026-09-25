Option Explicit On
Option Strict On

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading.Tasks

' ═══════════════════════════════════════════════════════════════════
' 🔴 H-01 — نسخ احتياطي مشفّر مع مصادقة سلامة (Encrypt-then-MAC)
' ── صيغة الملف الناتج (الإصدار 2):
'      [4 بايت  : التوقيع NDBK] [1 بايت : إصدار الصيغة = 2]
'      [16 بايت : ملح PBKDF2  ] [16 بايت : IV]
'      [النص المشفّر (AES-256-CBC + PKCS7)]
'      [آخر 32 بايت : HMAC-SHA256 على (الرأس + النص المشفّر)]
' ── ما الجديد عن الإصدار 1؟
'    * CBC وحده كان يقبل التعديل الصامت (Bit Flipping) — النسخة تُستعاد
'      وتبدو سليمة وهي تحوي بيانات مزوّرة. الآن أي تعديل ولو بايت واحد
'      يُكتشف قبل فك التشفير ولا يُستعاد شيء من ملف مشكوك في سلامته.
'    * مفتاح الـ HMAC مستقل عن مفتاح التشفير — كلاهما مشتق من نفس كلمة
'      المرور والملح عبر PBKDF2 (بايتات متتالية من نفس التدفق).
'    * المقارنة بزمن ثابت (FixedTimeEquals) — لا تتسرب معلومات عبر التوقيت.
'    * فك نسخ الإصدار 1 (قبل الإصلاح) ما زال مدعوماً مع تحذير صريح من الواجهة —
'      حتى لا يفقد المستخدم نسخه القديمة، مع العلم أنها بلا حماية سلامة.
' ── قواعد التصميم (كما هي):
'    * المفتاح يُشتق من كلمة مرور فقط — لا تُخزن أي مفاتيح على القرص
'    * 150,000 جولة PBKDF2 (SHA1) — توازن آمن مع أجهزة قديمة
'    * التشفير/فك التشفير بالتدفقات على أجزاء 1MB — الملف مهما كبر لا يُحمّل كله بالذاكرة
' ═══════════════════════════════════════════════════════════════════
Public Class EncryptedBackupService

    ' توقيع الملف + رقم الإصدار
    Private Shared ReadOnly MagicBytes As Byte() = {&H4E, &H44, &H42, &H4B}   ' "NDBK"
    Private Const FormatVersion As Byte = 2          ' 🔴 H-01: الإصدار 2 = مع بصمة HMAC
    Private Const LegacyFormatVersion As Byte = 1    ' الإصدار 1 = بلا HMAC (يُفك بتحذير)
    Private Const SaltSize As Integer = 16
    Private Const IvSize As Integer = 16
    Private Const KeySizeBits As Integer = 256
    Private Const HmacSize As Integer = 32           ' HMAC-SHA256 = 256 بت
    Private Const Pbkdf2Iterations As Integer = 150000
    Private Const ChunkSize As Integer = 1048576   ' 1MB

    Public Class BackupResult
        Public Success As Boolean = False
        Public ErrorMessage As String = ""
        Public OutputPath As String = ""
        Public DecryptedTempPath As String = ""
        ' 🔴 H-01: النسخة بالصيغة القديمة (1) — فُكّت بنجاح لكنها بلا بصمة سلامة
        Public LegacyNoHmac As Boolean = False
    End Class

    ' حاوية المفاتيح المشتقة: مفتاح التشفير + مفتاح المصادقة (مستقلان)
    Private Class DerivedKeys
        Public EncryptionKey As Byte()
        Public MacKey As Byte()
    End Class

    ''' <summary>
    ''' إنشاء نسخة احتياطية مشفّرة: نسخ SQLite سليم أولاً (BackupDatabase) ثم تشفيره إلى ملف المستخدم.
    ''' 🔴 H-01: بعد التشفير تُحسب بصمة HMAC-SHA256 على (الرأس + النص المشفّر) وتُكتب بنهاية الملف.
    ''' عملية متزامنة — تُستدعى من الواجهة داخل Task.Run كي لا تتجمد الشاشة.
    ''' </summary>
    Public Function CreateEncryptedBackup(outputPath As String, password As String) As BackupResult
        Dim result As New BackupResult()
        Dim tempDb As String = ""
        Dim cipherTemp As String = ""

        Try
            ' ── 1. التحقق من المدخلات ──
            If String.IsNullOrWhiteSpace(outputPath) Then
                result.ErrorMessage = "لم يُحدد مسار ملف الحفظ"
                Return result
            End If
            If String.IsNullOrEmpty(password) Then
                result.ErrorMessage = "كلمة المرور مطلوبة"
                Return result
            End If
            If Not DatabaseModule.IsInitialized OrElse
               String.IsNullOrEmpty(DatabaseModule.CurrentDatabasePath) OrElse
               Not File.Exists(DatabaseModule.CurrentDatabasePath) Then
                result.ErrorMessage = "قاعدة البيانات غير موجودة أو غير مهيأة"
                Return result
            End If

            ' ── 2. نسخة SQLite سليمة أولاً (نفس آلية النسخ الأصلي للبرنامج) ──
            tempDb = Path.Combine(Path.GetTempPath(),
                $"EncBak_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.tmpdb")

            If Not DatabaseModule.BackupDatabase(tempDb) Then
                result.ErrorMessage = "فشل إنشاء النسخة الداخلية قبل التشفير"
                Return result
            End If

            ' ── 3. التشفير إلى ملف وسيط (النص المشفّر فقط) ──
            Dim salt As Byte() = GenerateRandomBytes(SaltSize)
            Dim iv As Byte() = GenerateRandomBytes(IvSize)
            Dim keys As DerivedKeys = DeriveKeys(password, salt)

            cipherTemp = Path.Combine(Path.GetTempPath(),
                $"EncCipher_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.tmp")

            EncryptFileToCipher(tempDb, cipherTemp, keys.EncryptionKey, iv)

            ' ── 4. 🔴 H-01: تجميع الملف النهائي — رأس + نص مشفّر + بصمة HMAC ──
            ' الرأس كامل (توقيع + إصدار + ملح + IV) يدخل في حساب الـ HMAC
            Dim header(MagicBytes.Length + 1 + SaltSize + IvSize - 1) As Byte
            Dim pos As Integer = 0
            Array.Copy(MagicBytes, 0, header, pos, MagicBytes.Length)
            pos += MagicBytes.Length
            header(pos) = FormatVersion
            pos += 1
            Array.Copy(salt, 0, header, pos, SaltSize)
            pos += SaltSize
            Array.Copy(iv, 0, header, pos, IvSize)

            Dim hmac As Byte() = ComputeHmacOverFile(header, cipherTemp, keys.MacKey)

            Using fsOut As New FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None)
                fsOut.Write(header, 0, header.Length)
                CopyFileToStream(cipherTemp, fsOut)
                fsOut.Write(hmac, 0, hmac.Length)
                fsOut.Flush()
            End Using

            result.Success = True
            result.OutputPath = outputPath
            Return result

        Catch ex As Exception
            DatabaseModule.LogError("EncryptedBackupService.CreateEncryptedBackup", ex)
            ' 🌟 إصلاح: ملف نصف مشفر كان يبقى عند الهدف بعد أي فشل — يُحذف
            Try
                If File.Exists(outputPath) Then File.Delete(outputPath)
            Catch
            End Try
            result.ErrorMessage = "فشل التشفير: " & ex.Message
            Return result
        Finally
            ' ── تنظيف الملفات المؤقتة فوراً ──
            Try
                If tempDb <> "" AndAlso File.Exists(tempDb) Then
                    ' 🌟 استبدال المحتوى بأصفار قبل الحذف (يبقى أثر الملف صعب القراءة)
                    OverwriteFileWithZeros(tempDb)
                    File.Delete(tempDb)
                End If
                ' النص المشفّر الوسيط: لا يحتوي بيانات صريحة — حذف عادي يكفي
                If cipherTemp <> "" AndAlso File.Exists(cipherTemp) Then
                    File.Delete(cipherTemp)
                End If
            Catch
                ' فشل التنظيف لا يُبطل نجاح العملية
            End Try
        End Try
    End Function

    ''' <summary>
    ''' فك تشفير نسخة احتياطية إلى ملف مؤقت + التحقق من أنه ملف SQLite سليم.
    ''' 🔴 H-01: للصيغة 2 يُتحقق من بصمة HMAC أولاً قبل أي محاولة فك تشفير —
    ''' أي تعديل بالملف (ولو بايت واحد) يُرفض فوراً برسالة صريحة.
    ''' نسخ الصيغة 1 القديمة تُفك كما هي مع وسم LegacyNoHmac لتحذير المستخدم.
    ''' عملية متزامنة — تُستدعى من الواجهة داخل Task.Run كي لا تتجمد الشاشة.
    ''' </summary>
    Public Function DecryptBackup(encryptedPath As String, password As String) As BackupResult
        Dim result As New BackupResult()

        Try
            If String.IsNullOrWhiteSpace(encryptedPath) OrElse Not File.Exists(encryptedPath) Then
                result.ErrorMessage = "الملف المحدد غير موجود"
                Return result
            End If
            If String.IsNullOrEmpty(password) Then
                result.ErrorMessage = "كلمة المرور مطلوبة"
                Return result
            End If

            Dim fileInfo As New FileInfo(encryptedPath)
            Dim headerLen As Integer = MagicBytes.Length + 1 + SaltSize + IvSize
            Dim minSizeV1 As Long = headerLen
            Dim minSizeV2 As Long = headerLen + HmacSize
            If fileInfo.Length < minSizeV1 Then
                result.ErrorMessage = "الملف صغير جداً — لا يبدو نسخة احتياطية مشفّرة لهذا البرنامج"
                Return result
            End If

            ' ── 1. قراءة الرأس والتحقق من التوقيع والإصدار ──
            Using fsIn As New FileStream(encryptedPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim magic(MagicBytes.Length - 1) As Byte
                If fsIn.Read(magic, 0, magic.Length) <> magic.Length Then
                    result.ErrorMessage = "فشل قراءة توقيع الملف"
                    Return result
                End If
                For i As Integer = 0 To MagicBytes.Length - 1
                    If magic(i) <> MagicBytes(i) Then
                        result.ErrorMessage = "هذا الملف ليس نسخة احتياطية مشفّرة من هذا البرنامج"
                        Return result
                    End If
                Next

                Dim version As Integer = fsIn.ReadByte()
                If version <> FormatVersion AndAlso version <> LegacyFormatVersion Then
                    result.ErrorMessage = $"إصدار صيغة التشفير غير مدعوم ({version})"
                    Return result
                End If

                Dim salt(SaltSize - 1) As Byte
                Dim iv(IvSize - 1) As Byte
                If fsIn.Read(salt, 0, salt.Length) <> salt.Length OrElse
                   fsIn.Read(iv, 0, iv.Length) <> iv.Length Then
                    result.ErrorMessage = "الملف تالف (رأس غير مكتمل)"
                    Return result
                End If

                ' إعادة بناء الرأس كاملاً — يدخل في حساب الـ HMAC
                Dim header(headerLen - 1) As Byte
                Dim pos As Integer = 0
                Array.Copy(MagicBytes, 0, header, pos, MagicBytes.Length)
                pos += MagicBytes.Length
                header(pos) = CByte(version)
                pos += 1
                Array.Copy(salt, 0, header, pos, SaltSize)
                pos += SaltSize
                Array.Copy(iv, 0, header, pos, IvSize)

                Dim keys As DerivedKeys = DeriveKeys(password, salt)

                If version = FormatVersion Then
                    ' ── 2. 🔴 H-01: التحقق من بصمة السلامة قبل أي فك تشفير ──
                    If fileInfo.Length < minSizeV2 Then
                        result.ErrorMessage = "الملف تالف — ناقص بصمة السلامة (HMAC)"
                        Return result
                    End If

                    Dim storedHmac(HmacSize - 1) As Byte
                    fsIn.Seek(fileInfo.Length - HmacSize, SeekOrigin.Begin)
                    If fsIn.Read(storedHmac, 0, HmacSize) <> HmacSize Then
                        result.ErrorMessage = "الملف تالف (فشل قراءة بصمة السلامة)"
                        Return result
                    End If

                    Dim cipherLen As Long = fileInfo.Length - headerLen - HmacSize
                    If cipherLen <= 0 Then
                        result.ErrorMessage = "الملف تالف (لا يحتوي بيانات مشفّرة)"
                        Return result
                    End If

                    Dim actualHmac As Byte() = ComputeHmacOverStreamRegion(
                        header, encryptedPath, headerLen, cipherLen, keys.MacKey)

                    ' 🌟 مقارنة بزمن ثابت — لا تُفشل مبكراً على أول بايت مختلف
                    If Not FixedTimeEquals(actualHmac, storedHmac) Then
                        result.ErrorMessage =
                            "فشل التحقق من سلامة الملف." & vbCrLf &
                            "السبب: كلمة المرور غير صحيحة، أو الملف مُعدّل/تالف بعد إنشائه." & vbCrLf &
                            "لن تُستعاد أي بيانات من ملف مشكوك في سلامته."
                        Return result
                    End If

                    ' ── 3. فك التشفير من منطقة النص المشفّر فقط (بدون بصمة النهاية) ──
                    fsIn.Seek(headerLen, SeekOrigin.Begin)
                    If Not DecryptStreamRegion(fsIn, cipherLen, keys.EncryptionKey, iv, result) Then Return result
                Else
                    ' ── الصيغة 1 (بلا HMAC): فك مباشر مع وسم التحذير ──
                    result.LegacyNoHmac = True
                    If Not DecryptStreamRegion(fsIn, fileInfo.Length - headerLen, keys.EncryptionKey, iv, result) Then Return result
                End If

                ' ── 4. تحقق السلامة: هل الناتج ملف SQLite فعلاً؟ ──
                If Not IsSqliteFile(result.DecryptedTempPath) Then
                    Try
                        File.Delete(result.DecryptedTempPath)
                    Catch
                    End Try
                    result.DecryptedTempPath = ""
                    result.ErrorMessage = "الملف فُك تشفيره لكنه ليس قاعدة بيانات سليمة"
                    Return result
                End If

                result.Success = True
                Return result

            End Using

        Catch ex As Exception
            DatabaseModule.LogError("EncryptedBackupService.DecryptBackup", ex)
            result.ErrorMessage = "خطأ غير متوقع: " & ex.Message
            Return result
        End Try
    End Function

    ''' <summary>
    ''' 🔴 H-01: فك تشفير منطقة مشفّرة (بطول محدد) من تدفق مفتوح إلى ملف مؤقت.
    ''' قراءة يدوية محكومة بالطول — حتى لا يبتلع فك التشفير بايتات الـ HMAC بعد منطقة النص المشفّر.
    ''' </summary>
    Private Function DecryptStreamRegion(source As FileStream, cipherLen As Long,
                                         key As Byte(), iv As Byte(),
                                         result As BackupResult) As Boolean
        Dim tempOut As String = Path.Combine(Path.GetTempPath(),
            $"EncRest_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.tmpdb")

        Try
            Using aes As Aes = Aes.Create()
                aes.KeySize = KeySizeBits
                aes.Key = key
                aes.IV = iv
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7

                Using decryptor As ICryptoTransform = aes.CreateDecryptor()
                    Using fsOut As New FileStream(tempOut, FileMode.Create, FileAccess.Write, FileShare.None)
                        Dim buffer(ChunkSize - 1) As Byte
                        Dim remaining As Long = cipherLen

                        While remaining > 0
                            Dim toRead As Integer = CInt(Math.Min(remaining, CLng(buffer.Length)))
                            Dim read As Integer = source.Read(buffer, 0, toRead)
                            If read <= 0 Then Exit While
                            remaining -= read

                            Dim outBuf(read + 16) As Byte
                            Dim written As Integer = decryptor.TransformBlock(buffer, 0, read, outBuf, 0)
                            If written > 0 Then fsOut.Write(outBuf, 0, written)
                        End While

                        ' الكتلة الأخيرة: تُطبق إزالة الـ PKCS7
                        Dim finalBytes As Byte() = decryptor.TransformFinalBlock(New Byte() {}, 0, 0)
                        If finalBytes IsNot Nothing AndAlso finalBytes.Length > 0 Then
                            fsOut.Write(finalBytes, 0, finalBytes.Length)
                        End If
                    End Using
                End Using
            End Using

            result.DecryptedTempPath = tempOut
            Return True
        Catch
            ' 🌟 فشل فك التشفير هنا يعني غالباً كلمة مرور خاطئة (PKCS7 padding fail)
            Try
                If File.Exists(tempOut) Then File.Delete(tempOut)
            Catch
            End Try
            result.ErrorMessage = "فشل فك التشفير — الأرجح أن كلمة المرور غير صحيحة"
            Return False
        End Try
    End Function

    ''' <summary>حذف الملف المؤقت المفكوك بعد انتهاء الاستعادة (يستدعيه المستوى الأعلى)</summary>
    Public Sub CleanupDecryptedTemp(tempPath As String)
        Try
            If tempPath <> "" AndAlso File.Exists(tempPath) Then
                OverwriteFileWithZeros(tempPath)
                File.Delete(tempPath)
            End If
        Catch ex As Exception
            ' الفشل لا يبطل العملية لكن يُسجل — الملف المتبقي نص صريح في %TEMP%
            DatabaseModule.LogError("EncryptedBackupService.CleanupDecryptedTemp", ex)
        End Try
    End Sub

    ' ═══════════════ أدوات داخلية ═══════════════

    ''' <summary>🔴 H-01: اشتقاق مفتاحين مستقلين (تشفير + مصادقة) من كلمة المرور — بايتات متتالية من نفس PBKDF2</summary>
    Private Function DeriveKeys(password As String, salt As Byte()) As DerivedKeys
        ' 🌟 صيغة الـ 3 معاملات متوافقة مع كل أهداف .NET Framework 4.x
        Using derive As New Rfc2898DeriveBytes(password, salt, Pbkdf2Iterations)
            Dim keys As New DerivedKeys()
            keys.EncryptionKey = derive.GetBytes(KeySizeBits \ 8)
            keys.MacKey = derive.GetBytes(HmacSize)
            Return keys
        End Using
    End Function

    ''' <summary>🔴 H-01: بصمة HMAC-SHA256 على (رأس الملف + كل بايتات ملف النص المشفّر)</summary>
    Private Function ComputeHmacOverFile(header As Byte(), cipherFilePath As String, macKey As Byte()) As Byte()
        Using h As New HMACSHA256(macKey)
            h.TransformBlock(header, 0, header.Length, Nothing, 0)
            PumpFileThroughHmac(h, cipherFilePath)
            h.TransformFinalBlock(New Byte() {}, 0, 0)
            Return h.Hash
        End Using
    End Function

    ''' <summary>🔴 H-01: بصمة HMAC على (رأس + منطقة محددة بالطول من ملف) — للاستعادة</summary>
    Private Function ComputeHmacOverStreamRegion(header As Byte(), filePath As String,
                                                 startOffset As Long, regionLength As Long,
                                                 macKey As Byte()) As Byte()
        Using h As New HMACSHA256(macKey)
            h.TransformBlock(header, 0, header.Length, Nothing, 0)
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                fs.Seek(startOffset, SeekOrigin.Begin)
                Dim buffer(ChunkSize - 1) As Byte
                Dim remaining As Long = regionLength
                While remaining > 0
                    Dim toRead As Integer = CInt(Math.Min(remaining, CLng(buffer.Length)))
                    Dim read As Integer = fs.Read(buffer, 0, toRead)
                    If read <= 0 Then Exit While
                    h.TransformBlock(buffer, 0, read, Nothing, 0)
                    remaining -= read
                End While
            End Using
            h.TransformFinalBlock(New Byte() {}, 0, 0)
            Return h.Hash
        End Using
    End Function

    ''' <summary>تمرير ملف كامل عبر حالة HMAC (بدون تحويل نهائي — يستدعيه المُجمّع)</summary>
    Private Sub PumpFileThroughHmac(h As HMACSHA256, filePath As String)
        Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
            Dim buffer(ChunkSize - 1) As Byte
            Dim read As Integer
            Do
                read = fs.Read(buffer, 0, buffer.Length)
                If read > 0 Then h.TransformBlock(buffer, 0, read, Nothing, 0)
            Loop While read > 0
        End Using
    End Sub

    ''' <summary>تشفير ملف SQLite سليم إلى ملف نص مشفّر خالص (بلا رأس)</summary>
    Private Sub EncryptFileToCipher(plainSourcePath As String, cipherDestPath As String,
                                    key As Byte(), iv As Byte())
        Using fsOut As New FileStream(cipherDestPath, FileMode.Create, FileAccess.Write, FileShare.None)
            Using aes As Aes = Aes.Create()
                aes.KeySize = KeySizeBits
                aes.Key = key
                aes.IV = iv
                aes.Mode = CipherMode.CBC
                aes.Padding = PaddingMode.PKCS7

                Using encryptor As ICryptoTransform = aes.CreateEncryptor()
                    Using cs As New CryptoStream(fsOut, encryptor, CryptoStreamMode.Write)
                        Using fsIn As New FileStream(plainSourcePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                            Dim buffer(ChunkSize - 1) As Byte
                            Dim read As Integer
                            Do
                                read = fsIn.Read(buffer, 0, buffer.Length)
                                If read > 0 Then cs.Write(buffer, 0, read)
                            Loop While read > 0
                        End Using
                        ' إغلاق CryptoStream يطبق الـ PKCS7 ويكتب الـ ciphertext كاملاً
                        cs.FlushFinalBlock()
                    End Using
                End Using
            End Using
        End Using
    End Sub

    ''' <summary>نسخ ملف كامل إلى تدفق مفتوح (للتجميع النهائي)</summary>
    Private Sub CopyFileToStream(filePath As String, target As Stream)
        Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
            Dim buffer(ChunkSize - 1) As Byte
            Dim read As Integer
            Do
                read = fs.Read(buffer, 0, buffer.Length)
                If read > 0 Then target.Write(buffer, 0, read)
            Loop While read > 0
        End Using
    End Sub

    ''' <summary>🔴 H-01: مقارنة بايتات بزمن ثابت — لا تتسرب معلومات عبر التوقيت (غير متوفرة في .NET Framework)</summary>
    Private Shared Function FixedTimeEquals(a As Byte(), b As Byte()) As Boolean
        If a Is Nothing OrElse b Is Nothing OrElse a.Length <> b.Length Then Return False
        Dim diff As Integer = 0
        For i As Integer = 0 To a.Length - 1
            diff = diff Or (CInt(a(i)) Xor CInt(b(i)))
        Next
        Return diff = 0
    End Function

    ''' <summary>توليد بايتات عشوائية مشفّرياً (ملح / IV)</summary>
    Private Function GenerateRandomBytes(size As Integer) As Byte()
        Using rng As RandomNumberGenerator = RandomNumberGenerator.Create()
            Dim data(size - 1) As Byte
            rng.GetBytes(data)
            Return data
        End Using
    End Function

    ''' <summary>هل الملف يبدأ بتوقيع SQLite القياسي؟</summary>
    Private Function IsSqliteFile(path As String) As Boolean
        Try
            Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim header(15) As Byte
                If fs.Read(header, 0, 16) < 16 Then Return False
                Dim expected As String = "SQLite format 3" & Chr(0)
                Dim actual As String = Encoding.ASCII.GetString(header)
                Return actual = expected
            End Using
        Catch
            Return False
        End Try
    End Function

    ''' <summary>الكتابة فوق ملف مؤقت بأصفار قبل حذفه (حماية إضافية للبيانات الحساسة)</summary>
    Private Sub OverwriteFileWithZeros(path As String)
        Try
            Dim info As New FileInfo(path)
            If info.Length = 0 Then Return
            Using fs As New FileStream(path, FileMode.Open, FileAccess.Write, FileShare.None)
                fs.Seek(0, SeekOrigin.Begin)
                Dim zeros(ChunkSize - 1) As Byte
                Dim remaining As Long = info.Length
                While remaining > 0
                    Dim toWrite As Integer = CInt(Math.Min(remaining, CLng(zeros.Length)))
                    fs.Write(zeros, 0, toWrite)
                    remaining -= toWrite
                End While
                fs.Flush()
            End Using
        Catch
            ' فشل الاستبدال لا يمنع الحذف
        End Try
    End Sub

End Class
