Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO

Public Class frmLogin
    Inherits System.Windows.Forms.Form

    ' تعريف عناصر الواجهة برمجياً
    Private txtUsername As New TextBox()
    Private txtPassword As New TextBox()
    Private lblUser As New Label()
    Private lblPass As New Label()
    Private btnLogin As New Button()
    Private btnExit As New Button()
    Private lblTitle As New Label()
    Private Async Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 🌟 إصلاح: تعطيل زر الدخول أثناء التهيئة — كان المستخدم يقدر يدخل قبل اكتمال
        ' تهيئة قاعدة البيانات فيوصل لرسائل خطأ مربكة
        btnLogin.Enabled = False
        ' ==========================================
        ' === وضع المطور: يتفعل فقط بوجود ملف devmode.key بجانب الملف التنفيذي ===
        Dim BYPASS_LOGIN As Boolean = IsDeveloperBypassEnabled()
        ' ==========================================
        ' LicenseManager.ResetLicense()

        ' 🌟 شاشة البداية — تظهر فوراً
        Using splash As New frmSplash()
            splash.ShowSplash()

            ' --- 1. فحص الترخيص ---
            splash.SetProgress(10, "فحص ترخيص النظام...")
            Await splash.PauseAsync()

            ' 🌟 إصلاح انهيار: كان فحص الترخيص خارج أي Try — أي استثناء هنا كان يُسقط البرنامج بالكامل
            Dim licenseExpired As Boolean = False
            Dim daysLeft As Integer = 0
            Try
                licenseExpired = LicenseManager.IsExpired()
                If Not licenseExpired Then daysLeft = LicenseManager.GetRemainingDays()
            Catch licEx As Exception
                splash.CloseSplash()
                DatabaseModule.LogError("frmLogin_Load - LicenseCheck", licEx)
                MessageBox.Show("حدث خطأ أثناء فحص الترخيص: " & licEx.Message & vbCrLf &
                                "يرجى إعادة تشغيل البرنامج أو التواصل مع مبرمج النظام.",
                                "خطأ في الترخيص", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End Try

            If licenseExpired Then
                splash.CloseSplash()
                MessageBox.Show("انتهت صلاحية استخدام هذا البرنامج." & vbCrLf & "يرجى التواصل مع مبرمج النظام لتجديد الاشتراك.", "صلاحية منتهية", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Application.Exit()
                Return
            End If

            ' 🌟 إصلاح: فشل قراءة الأيام يعيد 0 — كان يخوّف مستخدم مرخص بـ"تبقى 0 يوم!"
            If daysLeft > 0 AndAlso daysLeft <= 7 Then
                Me.Text = $"تسجيل الدخول - تبقى {daysLeft} يوم على انتهاء الصلاحية!"
            End If
            splash.SetProgress(25, "الترخيص سليم ✅")
            Await splash.PauseAsync()

            ' --- 2. قاعدة البيانات ---
            splash.SetProgress(40, "تهيئة قاعدة البيانات...")
            Try
                If Not DatabaseModule.IsInitialized Then
                    DatabaseModule.Initialize(True)
                End If
                splash.SetProgress(60, "إنشاء الجداول والفهارس...")
                Await splash.PauseAsync()

                DatabaseModule.EnsureUsersTableExists()

                splash.SetProgress(80, "ترحيل البيانات القديمة...")
                Await splash.PauseAsync()

                splash.SetProgress(95, "اكتمل التحميل ✅")
                Await splash.PauseAsync()
            Catch ex As Exception
                splash.CloseSplash()
                DatabaseModule.LogError("frmLogin_Load", ex)
                MessageBox.Show("فشل في تهيئة قاعدة البيانات: " & ex.Message, "خطأ فادح", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End Try

            ' --- 3. إغلاق الشاشة والانتقال ---
            splash.SetProgress(100, "مرحباً بك 👋")
            Await splash.PauseAsync()
            splash.CloseSplash()
        End Using

        ' 🌟 إعادة تفعيل زر الدخول بعد اكتمال التهيئة
        btnLogin.Enabled = True

        ' --- 4. تخطي الدخول للمطور ---
        If BYPASS_LOGIN Then
            DatabaseModule.LogInfo("frmLogin_Load", "وضع المطور مفعل (devmode.key موجود) — دخول تلقائي كـ admin")
            UserSession.CurrentUsername = "admin"
            UserSession.CanEdit = True
            UserSession.CanDelete = True
            UserSession.MustChangePassword = False

            Me.Hide()
            Using mainForm As New frmPlayers()
                mainForm.ShowDialog()
            End Using
            Me.Close()
            Return
        End If
    End Sub
    ''' <summary>
    ''' وضع المطور: تُفعّل فقط إذا وُجد ملف devmode.key بجانب الملف التنفيذي.
    ''' أنشئ الملف يدوياً (حتى فارغاً) على جهاز التطوير فقط — لا يُوزّع أبداً مع البرنامج.
    ''' </summary>
    Private Function IsDeveloperBypassEnabled() As Boolean
        Try
            ' لا يُسمح بتجاوز الدخول في البناء الإنتاجي حتى لو وُجد الملف.
#If Not DEBUG Then
            Return False
#Else
            Return File.Exists(Path.Combine(Application.StartupPath, "devmode.key"))
#End If
        Catch
            Return False
        End Try
    End Function
    Private Sub SetupUI()
        Me.Text = "تسجيل الدخول"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(350, 280)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.RightToLeft = RightToLeft.Yes
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ' العنوان
        lblTitle.Text = "نظام إدارة اللاعبين"
        lblTitle.AutoSize = False
        lblTitle.Size = New Size(320, 40)
        lblTitle.Location = New Point(15, 15)
        lblTitle.Font = New Font("Arial", 16, FontStyle.Bold)
        lblTitle.ForeColor = Color.DarkBlue
        lblTitle.TextAlign = ContentAlignment.MiddleCenter

        ' اسم المستخدم
        lblUser.Text = "اسم المستخدم:"
        lblUser.AutoSize = True
        lblUser.Location = New Point(220, 80)
        lblUser.Font = New Font("Arial", 10)
        txtUsername.Location = New Point(30, 77)
        txtUsername.Width = 180
        txtUsername.Font = New Font("Arial", 10)

        ' كلمة المرور
        lblPass.Text = "كلمة المرور:"
        lblPass.AutoSize = True
        lblPass.Location = New Point(220, 120)
        lblPass.Font = New Font("Arial", 10)
        txtPassword.Location = New Point(30, 117)
        txtPassword.Width = 180
        txtPassword.Font = New Font("Arial", 10)
        txtPassword.PasswordChar = "*"c

        ' أزرار الدخول والخروج
        btnLogin.Text = "دخول"
        btnLogin.Size = New Size(100, 40)
        btnLogin.Location = New Point(30, 170)
        btnLogin.BackColor = Color.FromArgb(100, 210, 130)   ' 🌟 توحيد أخضر الحفظ مع باقي الشاشات
        btnLogin.FlatStyle = FlatStyle.Flat
        btnLogin.Font = New Font("Arial", 10, FontStyle.Bold)

        btnExit.Text = "خروج"
        btnExit.Size = New Size(100, 40)
        btnExit.Location = New Point(140, 170)
        btnExit.BackColor = Color.LightCoral
        btnExit.FlatStyle = FlatStyle.Flat
        btnExit.Font = New Font("Arial", 10, FontStyle.Bold)

        AddHandler btnLogin.Click, AddressOf DoLogin
        AddHandler btnExit.Click, Sub() Me.Close()

        Me.Controls.AddRange({lblTitle, lblUser, txtUsername, lblPass, txtPassword, btnLogin, btnExit})
        Me.AcceptButton = btnLogin
    End Sub

    ' 🌟 عداد المحاولات الفاشلة لتأخير التخمين الآلي
    Private _failedAttempts As Integer = 0

    Private Async Sub DoLogin(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtUsername.Text) OrElse String.IsNullOrWhiteSpace(txtPassword.Text) Then
            MessageBox.Show("الرجاء إدخال اسم المستخدم وكلمة المرور", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        btnLogin.Enabled = False
        btnLogin.Text = "جاري التحقق..."

        Try
            Dim service As New UserService()
            Dim isValid As Boolean = Await service.ValidateLoginAsync(txtUsername.Text, txtPassword.Text)
            If isValid Then
                If UserSession.MustChangePassword Then
                    Using changeForm As New frmChangePassword()
                        changeForm.CurrentUsername = UserSession.CurrentUsername
                        If changeForm.ShowDialog() <> DialogResult.OK Then
                            UserSession.ClearSession()
                            btnLogin.Enabled = True
                            btnLogin.Text = "دخول"
                            Return
                        End If
                    End Using
                End If

                _failedAttempts = 0

                ' 1. إخفاء شاشة الدخول
                Me.Hide()

                ' 2. فتح الشاشة الرئيسية وتدميرها من الذاكرة بعد الإغلاق
                Using mainForm As New frmPlayers()
                    mainForm.ShowDialog()
                End Using

                ' 3. عندما يغلق المستخدم frmPlayers، نتحقق هل هو تسجيل خروج أم إغلاق؟
                If UserSession.IsLoggingOut Then
                    ' تنظيف بيانات المستخدم السابق
                    UserSession.ClearSession()

                    ' إعادة إظهار شاشة الدخول ومسح الحقول
                    txtUsername.Clear()
                    txtPassword.Clear()

                    ' --- إعادة تفعيل زر الدخول وتغيير نصه ---
                    btnLogin.Enabled = True
                    btnLogin.Text = "دخول"
                    ' -------------------------------------

                    Me.Show()
                    txtUsername.Focus()
                Else
                    ' إذا لم يكن تسجيل خروج، فهذا يعني إغلاق البرنامج كلياً
                    Me.Close()
                End If
            Else
                ' 🌟 إصلاح: تأخير متصاعد بعد المحاولات الفاشلة — كان التخمين الآلي غير محدود
                _failedAttempts += 1
                If _failedAttempts >= 2 Then
                    Await Task.Delay(Math.Min(_failedAttempts * 500, 5000))
                End If
                MessageBox.Show("اسم المستخدم أو كلمة المرور غير صحيحة", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                btnLogin.Enabled = True
                btnLogin.Text = "دخول"
                txtPassword.Clear() ' 🌟 مسح كلمة المرور الفاشلة (كانت تبقى بالحقل)
                txtPassword.Focus()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmLogin.DoLogin", ex)
            ' 🌟 إصلاح: إذا رمى فتح الشاشة الرئيسية استثناءً بعد Me.Hide() كانت الشاشة
            ' تبقى مخفية بلا أي نافذة (تطبيق زومبي) — نعيد إظهار شاشة الدخول
            If Not Me.Visible Then Me.Show()
            MessageBox.Show("خطأ في النظام: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnLogin.Enabled = True
            btnLogin.Text = "دخول"
        End Try
    End Sub

    ' دالة البناء (Constructor)
    Public Sub New()
        ' هذا السطر مطلوب من مصمم الواجهات
        InitializeComponent()

        ' بعد التهيئة، قم ببناء بقية الواجهة برمجياً
        SetupUI()
    End Sub

    ' دالة تهيئة المكونات
    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'frmLogin
        '
        Me.ClientSize = New System.Drawing.Size(350, 280)
        Me.Name = "frmLogin"
        Me.ResumeLayout(False)
    End Sub

End Class
