Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Public Class frmChangePassword
    Inherits System.Windows.Forms.Form

    ' خاصية لحفظ اسم المستخدم الحالي (سنرسلها من frmPlayers)
    Public Property CurrentUsername As String = ""

    Private txtOldPass As New TextBox()
    Private txtNewPass As New TextBox()
    Private txtConfirmPass As New TextBox()
    Private lblOld As New Label()
    Private lblNew As New Label()
    Private lblConfirm As New Label()
    Private btnChange As New Button()
    Private btnCancel As New Button()

    Public Sub New()
        ' 🌟 هذا السطر مطلوب من مصمم الواجهات ويجب أن يكون الأول
        InitializeComponent()
        SetupUI()
    End Sub
    Private Sub SetupUI()
        Me.Text = "تغيير كلمة المرور"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(350, 280)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.RightToLeft = RightToLeft.Yes
        Me.BackColor = Color.FromArgb(240, 240, 240)

        lblOld.Text = "كلمة المرور الحالية:"
        lblOld.AutoSize = True
        lblOld.Location = New Point(200, 25)
        txtOldPass.Location = New Point(20, 22)
        txtOldPass.Width = 170
        txtOldPass.PasswordChar = "*"c

        lblNew.Text = "كلمة المرور الجديدة:"
        lblNew.AutoSize = True
        lblNew.Location = New Point(200, 65)
        txtNewPass.Location = New Point(20, 62)
        txtNewPass.Width = 170
        txtNewPass.PasswordChar = "*"c

        lblConfirm.Text = "تأكيد كلمة المرور:"
        lblConfirm.AutoSize = True
        lblConfirm.Location = New Point(200, 105)
        txtConfirmPass.Location = New Point(20, 102)
        txtConfirmPass.Width = 170
        txtConfirmPass.PasswordChar = "*"c

        btnChange.Text = "تغيير"
        btnChange.Size = New Size(100, 35)
        btnChange.Location = New Point(20, 150)
        btnChange.BackColor = Color.FromArgb(100, 210, 130)   ' 🌟 توحيد أخضر الحفظ مع باقي الشاشات
        btnChange.FlatStyle = FlatStyle.Flat

        btnCancel.Text = "إلغاء"
        btnCancel.Size = New Size(100, 35)
        btnCancel.Location = New Point(130, 150)
        btnCancel.BackColor = Color.LightCoral
        btnCancel.FlatStyle = FlatStyle.Flat
        btnCancel.DialogResult = DialogResult.Cancel

        AddHandler btnChange.Click, AddressOf DoChange

        Me.Controls.AddRange({lblOld, txtOldPass, lblNew, txtNewPass, lblConfirm, txtConfirmPass, btnChange, btnCancel})
        Me.AcceptButton = btnChange
        Me.CancelButton = btnCancel
    End Sub

    Private Async Sub DoChange(sender As Object, e As EventArgs)
        ' التحقق من الحد الأدنى داخل الواجهة، مع تكراره داخل UserService.
        If String.IsNullOrWhiteSpace(txtOldPass.Text) OrElse String.IsNullOrWhiteSpace(txtNewPass.Text) OrElse String.IsNullOrWhiteSpace(txtConfirmPass.Text) Then
            MessageBox.Show("الرجاء تعبئة جميع الحقول", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If txtNewPass.Text.Length < 8 Then
            MessageBox.Show("كلمة المرور الجديدة يجب أن تكون 8 أحرف على الأقل", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNewPass.Focus()
            Return
        End If

        ' 🌟 إصلاح: منع "التغيير" إلى نفس كلمة المرور القديمة — كان مسموحاً ويُبطل غرض الإجبار على التغيير
        If txtNewPass.Text = txtOldPass.Text Then
            MessageBox.Show("كلمة المرور الجديدة يجب أن تكون مختلفة عن الحالية", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNewPass.Focus()
            Return
        End If

        If txtNewPass.Text <> txtConfirmPass.Text Then
            MessageBox.Show("كلمة المرور الجديدة وتأكيدها غير متطابقين", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        btnChange.Enabled = False
        btnChange.Text = "جاري التغيير..."

        Try
            Dim service As New UserService()
            Dim success As Boolean = Await service.ChangePasswordAsync(CurrentUsername, txtOldPass.Text, txtNewPass.Text)

            If success Then
                MessageBox.Show("تم تغيير كلمة المرور بنجاح", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                ' 🌟 رسالة أدق: الفشل صار ممكن لثلاثة أسباب (كلمة قديمة خاطئة / جديدة قصيرة أو فارغة / مطابقة للقديمة)
                MessageBox.Show("فشل تغيير كلمة المرور — تأكد أن كلمة المرور الحالية صحيحة وأن الجديدة " & vbCrLf &
                                "8 أحرف على الأقل ومختلفة عن الحالية", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                btnChange.Enabled = True
                btnChange.Text = "تغيير"
                txtOldPass.Focus()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("frmChangePassword.DoChange", ex)
            MessageBox.Show("خطأ في النظام: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnChange.Enabled = True
            btnChange.Text = "تغيير"
        End Try
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'frmChangePassword
        '
        Me.ClientSize = New System.Drawing.Size(350, 280)
        Me.Name = "frmChangePassword"
        Me.ResumeLayout(False)

    End Sub


End Class
