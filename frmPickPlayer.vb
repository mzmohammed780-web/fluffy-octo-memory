Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

' ═══════════════════════════════════════════════════════════════════
' 🔴 H-03 — حوار اختيار صريح عند وجود أكثر من لاعب بنفس الاسم
'   كان النظام يربط الدفعة بأصغر رقم هوية صامتاً — قد تُحسب دفعة مالية
'   على اللاعب الخطأ ويتأثر ذلك بالتقارير والمتأخرات ولوحة الإحصاء كلها.
'   الخيارات تُعرض بـ(رقم الهوية + الهاتف) ليتعرف المستخدم على المقصود.
' نمط واجهة مبنية بالكود — نفس نمط frmSecureBackup وfrmLogin
' ═══════════════════════════════════════════════════════════════════
Public Class frmPickPlayer
    Inherits System.Windows.Forms.Form

    Private lblTitle As New Label()
    Private lstPlayers As New ListBox()
    Private btnOk As New Button()
    Private btnCancel As New Button()

    Private ReadOnly _options As List(Of Tuple(Of Long, String))

    ''' <summary>المعرف المختار — 0 إذا لم يُختر شيء</summary>
    Public ReadOnly Property SelectedId As Long
        Get
            If lstPlayers IsNot Nothing AndAlso lstPlayers.SelectedIndex >= 0 AndAlso
               lstPlayers.SelectedIndex < _options.Count Then
                Return _options(lstPlayers.SelectedIndex).Item1
            End If
            Return 0
        End Get
    End Property

    Public Sub New(playerName As String, options As List(Of Tuple(Of Long, String)))
        MyBase.New()
        _options = options
        BuildUI(playerName)
    End Sub

    Private Sub BuildUI(playerName As String)
        Me.Text = "تأكيد اللاعب المقصود"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(470, 330)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.RightToLeft = RightToLeft.Yes
        Me.BackColor = Color.FromArgb(243, 244, 246)
        Me.Font = New Font("Arial", 9.5!)

        lblTitle.Text = $"يوجد {_options.Count} لاعبين بنفس الاسم ""{playerName}""." & vbCrLf &
                        "اختر اللاعب المقصود قبل الحفظ:"
        lblTitle.AutoSize = False
        lblTitle.Size = New Size(420, 44)
        lblTitle.Location = New Point(25, 12)
        lblTitle.TextAlign = ContentAlignment.MiddleRight
        lblTitle.Font = New Font("Arial", 9.5!, FontStyle.Bold)

        lstPlayers.Location = New Point(25, 62)
        lstPlayers.Size = New Size(420, 160)
        lstPlayers.Font = New Font("Consolas", 10.0!)

        For Each opt As Tuple(Of Long, String) In _options
            Dim phonePart As String = If(String.IsNullOrWhiteSpace(opt.Item2), "بلا هاتف مسجل", $"هاتف: {opt.Item2}")
            lstPlayers.Items.Add($"هوية #{opt.Item1} — {phonePart}")
        Next
        If lstPlayers.Items.Count > 0 Then lstPlayers.SelectedIndex = 0

        AddHandler lstPlayers.DoubleClick, Sub(s As Object, e As EventArgs) ConfirmSelection()

        btnOk.Text = "اختيار"
        btnOk.Size = New Size(120, 34)
        btnOk.Location = New Point(25, 238)
        btnOk.BackColor = Color.FromArgb(100, 210, 130)
        btnOk.FlatStyle = FlatStyle.Flat
        AddHandler btnOk.Click, Sub(s As Object, e As EventArgs) ConfirmSelection()

        btnCancel.Text = "إلغاء الحفظ"
        btnCancel.Size = New Size(120, 34)
        btnCancel.Location = New Point(160, 238)
        btnCancel.DialogResult = DialogResult.Cancel

        Me.Controls.AddRange({lblTitle, lstPlayers, btnOk, btnCancel})
        Me.AcceptButton = btnOk
        Me.CancelButton = btnCancel
    End Sub

    Private Sub ConfirmSelection()
        If lstPlayers.SelectedIndex < 0 Then
            MessageBox.Show("اختر لاعباً من القائمة أولاً.", "تنبيه",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
