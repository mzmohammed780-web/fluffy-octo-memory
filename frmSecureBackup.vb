Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Threading.Tasks
Imports System.Data.SQLite

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — واجهة النسخ الاحتياطي المشفّر والاستعادة
' 🌟 إصلاح v2: الملف المؤقت المفكوك يُحذف فوراً بعد الاستبدال،
'              حتى لو فشل شيء آخر — لا يبقى نص صريح على القرص.
' ═══════════════════════════════════════════════════════════════════
Public Class frmSecureBackup
    Inherits System.Windows.Forms.Form

    Private grpBackup As New GroupBox()
    Private btnBackup As New Button()
    Private grpRestore As New GroupBox()
    Private btnRestore As New Button()
    Private lblWarning As New Label()
    Private _isBusy As Boolean = False

    Public Sub New()
        MyBase.New()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Try
            Me.Text = "النسخ الاحتياطي المشفّر (AES-256)"
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Size = New Size(560, 400)
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.RightToLeft = RightToLeft.Yes
            Me.BackColor = Color.FromArgb(243, 244, 246)
            Me.Font = New Font("Arial", 9.5!)

            lblWarning.Text =
                "النسخة المشفّرة لا تُفتح إلا بكلمة المرور التي أنشأتها — " &
                "فاقد كلمة المرور = فاقد النسخة نهائياً." & vbCrLf &
                "احفظ كلمة المرور في مكان آمن منفصل عن ملف النسخة."
            lblWarning.AutoSize = False
            lblWarning.Size = New Size(520, 42)
            lblWarning.Location = New Point(20, 12)
            lblWarning.Font = New Font("Arial", 9.0!, FontStyle.Bold)
            lblWarning.ForeColor = Color.FromArgb(150, 80, 0)
            lblWarning.TextAlign = ContentAlignment.MiddleRight

            grpBackup.Text = "إنشاء نسخة احتياطية مشفّرة"
            grpBackup.Size = New Size(520, 120)
            grpBackup.Location = New Point(20, 62)

            btnBackup.Text = "إنشاء نسخة مشفّرة..."
            btnBackup.Size = New Size(200, 40)
            btnBackup.Location = New Point(280, 55)
            btnBackup.BackColor = Color.FromArgb(100, 210, 130)
            btnBackup.FlatStyle = FlatStyle.Flat
            btnBackup.Font = New Font("Arial", 10.0!, FontStyle.Bold)
            AddHandler btnBackup.Click, AddressOf DoBackup

            Dim lblBackupHint As New Label()
            lblBackupHint.Text =
                "يُنشأ ملف .ndbak يحوي قاعدة بيانات كاملة مشفّرة AES-256." & vbCrLf &
                "التشفير يتم بعد أخذ نسخة SQLite سليمة، ولن يتأثر البرنامج أثناءها."
            lblBackupHint.AutoSize = False
            lblBackupHint.Size = New Size(230, 55)
            lblBackupHint.Location = New Point(30, 45)
            lblBackupHint.TextAlign = ContentAlignment.MiddleRight

            grpBackup.Controls.AddRange({btnBackup, lblBackupHint})

            grpRestore.Text = "استعادة من نسخة مشفّرة"
            grpRestore.Size = New Size(520, 120)
            grpRestore.Location = New Point(20, 195)

            btnRestore.Text = "استعادة من ملف مشفّر..."
            btnRestore.Size = New Size(200, 40)
            btnRestore.Location = New Point(280, 55)
            btnRestore.BackColor = Color.FromArgb(240, 200, 130)
            btnRestore.FlatStyle = FlatStyle.Flat
            btnRestore.Font = New Font("Arial", 10.0!, FontStyle.Bold)
            AddHandler btnRestore.Click, AddressOf DoRestore

            Dim lblRestoreHint As New Label()
            lblRestoreHint.Text =
                "تكتب فوق قاعدة البيانات الحالية بالكامل!" & vbCrLf &
                "يُتخذ تلقائياً نسخ أمان من الوضع الحالي قبل الاستبدال،" & vbCrLf &
                "ويُعاد تشغيل البرنامج بعد النجاح."
            lblRestoreHint.AutoSize = False
            lblRestoreHint.Size = New Size(230, 65)
            lblRestoreHint.Location = New Point(30, 40)
            lblRestoreHint.TextAlign = ContentAlignment.MiddleRight

            grpRestore.Controls.AddRange({btnRestore, lblRestoreHint})

            Me.Controls.AddRange({lblWarning, grpBackup, grpRestore})
        Catch ex As Exception
            DatabaseModule.LogError("frmSecureBackup.BuildUI", ex)
        End Try
    End Sub

    Private Async Sub DoBackup(sender As Object, e As EventArgs)
        If _isBusy Then Return

        Try
            Dim pass1 As String = ""
            Using passPrompt As New frmPasswordPrompt(
                    "كلمة مرور التشفير",
                    "أدخل كلمة مرور لحماية النسخة المشفّرة:",
                    True, 8)
                If passPrompt.ShowDialog(Me) <> DialogResult.OK Then Return
                pass1 = passPrompt.EnteredPassword
            End Using

            Using sfd As New SaveFileDialog()
                sfd.Title = "حفظ النسخة الاحتياطية المشفّرة"
                sfd.Filter = "نسخة احتياطية مشفّرة|*.ndbak"
                sfd.DefaultExt = "ndbak"
                sfd.FileName = $"SecureBackup_{DateTime.Now:yyyy-MM-dd_HH-mm}.ndbak"
                If sfd.ShowDialog() <> DialogResult.OK Then Return

                _isBusy = True
                btnBackup.Enabled = False
                btnBackup.Text = "جاري التشفير..."
                Me.Cursor = Cursors.WaitCursor

                Dim chosenPath As String = sfd.FileName
                Dim chosenPass As String = pass1

                Dim backupSvc As New EncryptedBackupService()
                Dim result As EncryptedBackupService.BackupResult =
                    Await Task.Run(Function() backupSvc.CreateEncryptedBackup(chosenPath, chosenPass))

                If result.Success Then
                    Dim audit As New AuditService()
                    audit.Log(AuditService.Act_BackupEncrypted, "قاعدة بيانات", "", Path.GetFileName(chosenPath))
                    MessageBox.Show("تم إنشاء النسخة الاحتياطية المشفّرة بنجاح!" & vbCrLf &
                                    "الموقع: " & chosenPath & vbCrLf & vbCrLf &
                                    "تأكد من حفظ كلمة المرور في مكان آمن.",
                                    "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(result.ErrorMessage, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("frmSecureBackup.DoBackup", ex)
            MessageBox.Show("خطأ في النسخ المشفّر: " & ex.Message, "خطأ",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isBusy = False
            btnBackup.Enabled = True
            btnBackup.Text = "إنشاء نسخة مشفّرة..."
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Async Sub DoRestore(sender As Object, e As EventArgs)
        If _isBusy Then Return

        Dim decryptedTempPath As String = ""

        Try
            Dim userInput As String = InputBox(
                "تحذير: الاستعادة تكتب فوق قاعدة البيانات الحالية بالكامل!" & vbCrLf &
                "جميع البيانات الحالية غير المحفوظة ستُستبدل." & vbCrLf & vbCrLf &
                "للتأكيد، اكتب الكلمة التالية: استعادة" & vbCrLf & vbCrLf &
                "أو اضغط إلغاء للتراجع.",
                "تأكيد الاستعادة من نسخة مشفّرة", "")
            If userInput.Trim() <> "استعادة" Then Return

            Using ofd As New OpenFileDialog()
                ofd.Title = "اختر ملف النسخة المشفّرة"
                ofd.Filter = "نسخة احتياطية مشفّرة|*.ndbak|جميع الملفات|*.*"
                If ofd.ShowDialog() <> DialogResult.OK Then Return
                If Not File.Exists(ofd.FileName) Then
                    MessageBox.Show("الملف المختار غير موجود", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim pass As String = ""
                Using restorePrompt As New frmPasswordPrompt(
                        "كلمة مرور فك التشفير",
                        "أدخل كلمة مرور النسخة المشفّرة:",
                        False, 1)
                    If restorePrompt.ShowDialog(Me) <> DialogResult.OK Then Return
                    pass = restorePrompt.EnteredPassword
                End Using

                _isBusy = True
                btnRestore.Enabled = False
                btnRestore.Text = "جاري فك التشفير..."
                Me.Cursor = Cursors.WaitCursor

                Dim chosenFile As String = ofd.FileName
                Dim chosenPass As String = pass

                Dim decryptSvc As New EncryptedBackupService()
                Dim decryptResult As EncryptedBackupService.BackupResult =
                    Await Task.Run(Function() decryptSvc.DecryptBackup(chosenFile, chosenPass))

                If Not decryptResult.Success Then
                    MessageBox.Show(decryptResult.ErrorMessage, "فشل", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                decryptedTempPath = decryptResult.DecryptedTempPath

                If decryptResult.LegacyNoHmac Then
                    MessageBox.Show(
                        "ملاحظة أمنية: هذه النسخة بالصيغة القديمة (قبل إضافة بصمة السلامة HMAC)." & vbCrLf &
                        "فُكّت بنجاح لكن ملفها لم يكن محمياً من التعديل الصامت." & vbCrLf &
                        "يُنصح بأخذ نسخة مشفّرة جديدة بالصيغة المحصّنة بعد هذه الاستعادة.",
                        "صيغة قديمة", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                Dim targetFile As String = DatabaseModule.CurrentDatabasePath
                If String.IsNullOrEmpty(targetFile) Then
                    MessageBox.Show("لم يتم العثور على مسار قاعدة البيانات الحالية", "خطأ",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' 🌟 إصلاح v2: الاستبدال + حذف الملف المؤقت داخل نفس البلوك الذري
                Dim tempPath As String = decryptedTempPath
                Dim plainBackupFailed As Boolean = False
                Await Task.Run(Sub()
                                   Try
                                       DatabaseModule.VacuumDatabase()
                                       SQLiteConnection.ClearAllPools()

                                       Dim autoBackup As String = Path.Combine(
                                           Path.GetDirectoryName(targetFile),
                                           $"AutoBackup_BeforeSecureRestore_{DateTime.Now:yyyy-MM-dd_HH-mm}.db")
                                       File.Copy(targetFile, autoBackup, True)
                                       File.Copy(tempPath, targetFile, True)

                                       For Each side As String In New String() {"-wal", "-shm", "-journal"}
                                           Dim sidePath As String = targetFile & side
                                           If File.Exists(sidePath) Then File.Delete(sidePath)
                                       Next
                                   Finally
                                       ' 🌟 حذف فوري: لا ننتظر Finally الخارجي
                                       If tempPath <> "" AndAlso File.Exists(tempPath) Then
                                           Try
                                               Dim cleanupSvc As New EncryptedBackupService()
                                               cleanupSvc.CleanupDecryptedTemp(tempPath)
                                           Catch
                                           End Try
                                       End If
                                   End Try
                               End Sub)
                ' بعد الحذف الفوري، نُفرغ المسار حتى لا يعيد Finally الخارجي المحاولة
                decryptedTempPath = ""

                Dim audit As New AuditService()
                audit.Log(AuditService.Act_RestoreEncrypted, "قاعدة بيانات", "", Path.GetFileName(chosenFile))

                MessageBox.Show("تمت الاستعادة بنجاح!" & vbCrLf &
                                "سيتم إعادة تشغيل البرنامج الآن.",
                                "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Application.Restart()
            End Using
        Catch ex As Exception
            DatabaseModule.LogError("frmSecureBackup.DoRestore", ex)
            MessageBox.Show("خطأ في الاستعادة: " & ex.Message, "خطأ",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If decryptedTempPath <> "" Then
                Try
                    Dim cleanupSvc As New EncryptedBackupService()
                    cleanupSvc.CleanupDecryptedTemp(decryptedTempPath)
                Catch
                End Try
            End If
            _isBusy = False
            btnRestore.Enabled = True
            btnRestore.Text = "استعادة من ملف مشفّر..."
            Me.Cursor = Cursors.Default
        End Try
    End Sub

End Class