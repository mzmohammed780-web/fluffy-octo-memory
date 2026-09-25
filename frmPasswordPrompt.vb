Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

' ═══════════════════════════════════════════════════════════════════
' 🔴 H-05 — نموذج إدخال كلمة مرور مقنّع (بديل InputBox المكشوف)
'   * الأحرف تظهر نقاطاً — لا مشاهدة كتفية
'   * زر "إظهار الأحرف مؤقتاً" للمراجعة البصرية القصيرة
'   * وضع تأكيد اختياري: حقلان يجب أن يتطابقا (لإنشاء كلمة مرور جديدة)
'   * حد أدنى قابل للضبط — الإنشاء 8 أحرف (موحّد مع سياسة المستخدمين)،
'     فك التشفير حد أدنى 1 (كلمات النسخ القديمة 6 أحرف يجب أن تُقبل)
'   * إرشاد عبارة مرور (Passphrase) — ملف النسخة لا يُكتب يومياً فيتحمّل كلمة أطول
' نمط واجهة مبنية بالكود — نفس نمط frmSecureBackup وfrmLogin
' ═══════════════════════════════════════════════════════════════════
Public Class frmPasswordPrompt
    Inherits System.Windows.Forms.Form

    Private lblPrompt As New Label()
    Private txtPass As New TextBox()
    Private txtConfirm As New TextBox()
    Private chkShow As New CheckBox()
    Private btnOk As New Button()
    Private btnCancel As New Button()
    Private lblHint As New Label()

    Private ReadOnly _requireConfirm As Boolean
    Private ReadOnly _minLength As Integer

    ''' <summary>كلمة المرور المُدخلة (تُقرأ بعد DialogResult.OK فقط)</summary>
    Public ReadOnly Property EnteredPassword As String
        Get
            Return txtPass.Text
        End Get
    End Property

    Public Sub New(title As String, promptText As String, requireConfirm As Boolean, minLength As Integer)
        MyBase.New()
        _requireConfirm = requireConfirm
        _minLength = Math.Max(1, minLength)
        BuildUI(title, promptText)
    End Sub

    Private Sub BuildUI(title As String, promptText As String)
        Me.Text = title
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(480, If(_requireConfirm, 300, 255))
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.RightToLeft = RightToLeft.Yes
        Me.BackColor = Color.FromArgb(243, 244, 246)
        Me.Font = New Font("Arial", 9.5!)

        lblPrompt.Text = promptText
        lblPrompt.AutoSize = False
        lblPrompt.Size = New Size(420, 40)
        lblPrompt.Location = New Point(30, 15)
        lblPrompt.TextAlign = ContentAlignment.MiddleRight
        lblPrompt.Font = New Font("Arial", 9.5!, FontStyle.Bold)

        txtPass.Location = New Point(30, 60)
        txtPass.Size = New Size(420, 26)
        txtPass.UseSystemPasswordChar = True
        txtPass.MaxLength = 128

        If _requireConfirm Then
            Dim lblConfirm As New Label()
            lblConfirm.Text = "تأكيد كلمة المرور:"
            lblConfirm.AutoSize = True
            lblConfirm.Location = New Point(30, 95)
            lblConfirm.TextAlign = ContentAlignment.MiddleRight

            txtConfirm.Location = New Point(30, 118)
            txtConfirm.Size = New Size(420, 26)
            txtConfirm.UseSystemPasswordChar = True
            txtConfirm.MaxLength = 128

            chkShow.Location = New Point(30, 150)
        Else
            chkShow.Location = New Point(30, 95)
        End If

        chkShow.Text = "إظهار الأحرف مؤقتاً"
        chkShow.AutoSize = True
        AddHandler chkShow.CheckedChanged,
            Sub(s As Object, e As EventArgs)
                txtPass.UseSystemPasswordChar = Not chkShow.Checked
                txtConfirm.UseSystemPasswordChar = Not chkShow.Checked
            End Sub

        lblHint.Text = $"{_minLength} أحرف على الأقل — يُنصح بعبارة مرور طويلة سهلة الحفظ" & vbCrLf &
                       "(جملة أو أرقام مفهومة لك وحدك) — فاقد الكلمة = فاقد النسخة نهائياً."
        lblHint.AutoSize = False
        lblHint.Size = New Size(420, 36)
        lblHint.Location = New Point(30, If(_requireConfirm, 176, 124))
        lblHint.ForeColor = Color.FromArgb(100, 100, 100)

        btnOk.Text = "موافق"
        btnOk.Size = New Size(120, 34)
        btnOk.Location = New Point(30, If(_requireConfirm, 220, 170))
        btnOk.BackColor = Color.FromArgb(100, 210, 130)
        btnOk.FlatStyle = FlatStyle.Flat
        AddHandler btnOk.Click, AddressOf ValidateAndClose

        btnCancel.Text = "إلغاء"
        btnCancel.Size = New Size(100, 34)
        btnCancel.Location = New Point(165, If(_requireConfirm, 220, 170))
        btnCancel.DialogResult = DialogResult.Cancel

        Me.Controls.AddRange({lblPrompt, txtPass, chkShow, lblHint, btnOk, btnCancel})
        If _requireConfirm Then
            Me.Controls.Add(txtConfirm)
        End If

        Me.AcceptButton = btnOk
        Me.CancelButton = btnCancel
    End Sub

    Private Sub ValidateAndClose(sender As Object, e As EventArgs)
        If txtPass.Text.Length < _minLength Then
            MessageBox.Show($"كلمة المرور يجب أن تكون {_minLength} أحرف على الأقل.", "تنبيه",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPass.Focus()
            Return
        End If

        If _requireConfirm AndAlso txtConfirm.Text <> txtPass.Text Then
            MessageBox.Show("الكلمتان غير متطابقتين — أعد الإدخال.", "تنبيه",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtConfirm.Focus()
            Return
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
