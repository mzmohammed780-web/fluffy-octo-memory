Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — شاشة القفل عند الخمول
' تُعرض حوارياً فوق النافذة النشطة، وتتطلب إعادة إدخال كلمة مرور
' صاحب الجلسة الحالية لفتحها (أو تسجيل خروج كامل).
' واجهة مبنية بالكود بالكامل (لا يعتمد على ملف Designer) — نفس نمط frmLogin
' ═══════════════════════════════════════════════════════════════════
Public Class frmLock
    Inherits System.Windows.Forms.Form

    Private lblIcon As New Label()
    Private lblTitle As New Label()
    Private lblUser As New Label()
    Private lblPass As New Label()
    Private txtPassword As New TextBox()
    Private btnUnlock As New Button()
    Private btnLogout As New Button()
    Private _isBusy As Boolean = False

    Public Sub New()
        MyBase.New()
        BuildUI()
    End Sub

    Private Sub BuildUI()
        Try
            Me.Text = "الجلسة مقفلة"
            Me.FormBorderStyle = FormBorderStyle.None          ' إطار لا يمكن تحريكه أو إغلاقه
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.Size = New Size(460, 330)
            Me.TopMost = True
            Me.RightToLeft = RightToLeft.Yes
            Me.BackColor = Color.FromArgb(34, 42, 60)          ' أزرق داكن (نفس لوحة الداشبورد)

            ' أيقونة القفل
            lblIcon.Text = "🔒"
            lblIcon.AutoSize = False
            lblIcon.Size = New Size(420, 60)
            lblIcon.Location = New Point(20, 25)
            lblIcon.Font = New Font("Segoe UI Symbol", 28.0!, FontStyle.Regular)
            lblIcon.TextAlign = ContentAlignment.MiddleCenter
            lblIcon.ForeColor = Color.White

            ' العنوان
            lblTitle.Text = "قُفلت الجلسة تلقائياً بسبب عدم النشاط"
            lblTitle.AutoSize = False
            lblTitle.Size = New Size(420, 35)
            lblTitle.Location = New Point(20, 88)
            lblTitle.Font = New Font("Arial", 13.0!, FontStyle.Bold)
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            lblTitle.ForeColor = Color.White

            ' اسم المستخدم الحالي
            lblUser.Text = "المستخدم الحالي: " & If(String.IsNullOrWhiteSpace(UserSession.CurrentUsername), "غير معروف", UserSession.CurrentUsername)
            lblUser.AutoSize = False
            lblUser.Size = New Size(420, 28)
            lblUser.Location = New Point(20, 128)
            lblUser.Font = New Font("Arial", 11.0!, FontStyle.Bold)
            lblUser.TextAlign = ContentAlignment.MiddleCenter
            lblUser.ForeColor = Color.FromArgb(160, 200, 255)

            ' حقل كلمة المرور
            lblPass.Text = "كلمة المرور:"
            lblPass.AutoSize = True
            lblPass.Location = New Point(330, 182)
            lblPass.Font = New Font("Arial", 10.0!)
            lblPass.ForeColor = Color.White

            txtPassword.Location = New Point(70, 178)
            txtPassword.Width = 250
            txtPassword.Font = New Font("Arial", 11.0!)
            txtPassword.PasswordChar = "*"c
            txtPassword.RightToLeft = RightToLeft.No

            ' زر الفتح
            btnUnlock.Text = "فتح الجلسة"
            btnUnlock.Size = New Size(150, 40)
            btnUnlock.Location = New Point(170, 228)
            btnUnlock.BackColor = Color.FromArgb(100, 210, 130)
            btnUnlock.FlatStyle = FlatStyle.Flat
            btnUnlock.Font = New Font("Arial", 10.0!, FontStyle.Bold)
            AddHandler btnUnlock.Click, AddressOf DoUnlock

            ' زر تسجيل الخروج (لمن لا يملك كلمة المرور)
            btnLogout.Text = "تسجيل الخروج"
            btnLogout.Size = New Size(120, 40)
            btnLogout.Location = New Point(40, 228)
            btnLogout.BackColor = Color.FromArgb(210, 110, 110)
            btnLogout.FlatStyle = FlatStyle.Flat
            btnLogout.Font = New Font("Arial", 9.0!, FontStyle.Bold)
            AddHandler btnLogout.Click, AddressOf DoLogout

            Me.Controls.AddRange({lblIcon, lblTitle, lblUser, lblPass, txtPassword, btnUnlock, btnLogout})
            Me.AcceptButton = btnUnlock
        Catch ex As Exception
            DatabaseModule.LogError("frmLock.BuildUI", ex)
        End Try
    End Sub

    ''' <summary>التحقق من كلمة المرور وفتح الجلسة</summary>
    Private Async Sub DoUnlock(sender As Object, e As EventArgs)
        If _isBusy Then Return

        Dim entered As String = txtPassword.Text
        If String.IsNullOrWhiteSpace(entered) Then
            MessageBox.Show("الرجاء إدخال كلمة المرور", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        _isBusy = True
        btnUnlock.Enabled = False
        btnUnlock.Text = "جاري التحقق..."
        txtPassword.Enabled = False

        Try
            ' 🌟 التحقق عبر نفس مسار الدخول الرسمي (نفس مستخدم الجلسة)
            Dim service As New UserService()
            Dim ok As Boolean = Await service.ValidateLoginAsync(UserSession.CurrentUsername, entered)

            If ok Then
                ' نجاح: تسجيل الحدث ثم إغلاق شاشة القفل
                Dim audit As New AuditService()
                audit.Log(AuditService.Act_SessionUnlocked, "جلسة", "", "فتح بعد قفل الخمول")
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                Dim audit As New AuditService()
                audit.Log(AuditService.Act_LoginFailed, "جلسة", "", "كلمة مرور خاطئة في شاشة القفل")
                MessageBox.Show("كلمة المرور غير صحيحة", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPassword.Clear()
                txtPassword.Focus()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmLock.DoUnlock", ex)
            MessageBox.Show("خطأ أثناء التحقق: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isBusy = False
            btnUnlock.Enabled = True
            btnUnlock.Text = "فتح الجلسة"
            txtPassword.Enabled = True
        End Try
    End Sub

    ''' <summary>تسجيل خروج كامل: يغلق شاشة القفل ثم الشاشة الرئيسية (فيعود لمسار تسجيل الخروج المعتاد)</summary>
    Private Sub DoLogout(sender As Object, e As EventArgs)
        Try
            Dim confirm As DialogResult = MessageBox.Show(
                "هل تريد تسجيل الخروج من البرنامج بالكامل؟",
                "تأكيد الخروج", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2)
            If confirm <> DialogResult.Yes Then Return

            Dim audit As New AuditService()
            audit.Log(AuditService.Act_Logout, "جلسة", "", "خروج من شاشة القفل")

            ' 🌟 المسار الآمن: نعلم الجلسة بخروج ثم نغلق النافذة المالكة لشاشة القفل (frmPlayers)
            ' — فيمر البرنامج على نفس منطق تسجيل الخروج الموجود أصلاً في frmPlayers و frmLogin
            UserSession.IsLoggingOut = True

            If Me.Owner IsNot Nothing Then
                Me.Owner.Close()
            End If

            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            DatabaseModule.LogError("frmLock.DoLogout", ex)
        End Try
    End Sub

    ''' <summary>منع الإغلاق بالـ Alt+F4 — يجب إدخال كلمة المرور أو تسجيل الخروج</summary>
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_SYSCOMMAND As Integer = &H112
        Const SC_CLOSE As Integer = &HF060
        If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt64() = SC_CLOSE Then
            Return   ' تجاهل أمر الإغلاق
        End If
        MyBase.WndProc(m)
    End Sub

End Class
