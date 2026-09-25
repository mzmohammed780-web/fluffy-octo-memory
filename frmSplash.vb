Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

''' <summary>
''' شاشة البداية (Splash Screen): لوجو + اسم النظام + شريط تقدم بخطوات التحميل الفعلية
''' بظل حقيقي (CS_DROPSHADOW) — تُغلق تلقائياً بعد اكتمال التحميل
''' </summary>
Public Class frmSplash

    ' 🌟 التوقيت الأدنى لعرض كل خطوة (ميلي ثانية) — حتى لا تختفي الشاشة بسرعة مخلة
    Private Const MinStepDelay As Integer = 250

    Private ReadOnly lblTitle As New Label()
    Private ReadOnly lblSubtitle As New Label()
    Private ReadOnly lblStatus As New Label()
    Private ReadOnly lblVersion As New Label()
    Private ReadOnly progressBar As New ProgressBar()

    Public Sub New()
        SetupUI()
    End Sub

    Private Sub SetupUI()
        ' ── خصائص النافذة ──
        Me.FormBorderStyle = FormBorderStyle.None          ' بلا حدود — الظل بيرسم من الـ CreateParams
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(520, 340)
        Me.TopMost = True
        Me.ShowInTaskbar = False
        Me.BackColor = Color.FromArgb(32, 38, 58)          ' كحلي غامق (نفس هوية الداشبورد)

        ' ── اللوجو (رمز نصي مؤقت — يمكنك استبداله بـ PictureBox بصورة لاحقاً) ──
        Dim lblLogo As New Label() With {
            .Text = "🏆",
            .Font = New Font("Segoe UI Emoji", 42),
            .AutoSize = True,
            .Location = New Point(230, 30),
            .BackColor = Color.Transparent
        }
        Me.Controls.Add(lblLogo)

        ' ── العنوان الرئيسي ──
        lblTitle.Text = "نظام إدارة اللاعبين"
        lblTitle.Font = New Font("Arial", 22, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(95, 105)
        Me.Controls.Add(lblTitle)

        ' ── العنوان الفرعي ──
        lblSubtitle.Text = "إدارة اللاعبين • الإيرادات • المصروفات"
        lblSubtitle.Font = New Font("Arial", 10)
        lblSubtitle.ForeColor = Color.FromArgb(140, 160, 200)
        lblSubtitle.AutoSize = True
        lblSubtitle.Location = New Point(150, 150)
        Me.Controls.Add(lblSubtitle)

        ' ── شريط التقدم ──
        progressBar.Location = New Point(60, 215)
        progressBar.Size = New Size(400, 18)
        progressBar.Style = ProgressBarStyle.Continuous
        progressBar.Minimum = 0
        progressBar.Maximum = 100
        Me.Controls.Add(progressBar)

        ' ── نص الخطوة الحالية ──
        lblStatus.Text = "جاري البدء..."
        lblStatus.Font = New Font("Arial", 9)
        lblStatus.ForeColor = Color.FromArgb(170, 190, 220)
        lblStatus.AutoSize = True
        lblStatus.Location = New Point(60, 243)
        Me.Controls.Add(lblStatus)

        ' ── الإصدار (يسار أسفل) ──
        lblVersion.Text = "الإصدار 2.0"
        lblVersion.Font = New Font("Arial", 8)
        lblVersion.ForeColor = Color.FromArgb(120, 140, 170)
        lblVersion.AutoSize = True
        lblVersion.Location = New Point(25, 300)
        Me.Controls.Add(lblVersion)
    End Sub

    ' ─────────────────────────────────────────
    ' 🌟 الظل الحقيقي (CS_DROPSHADOW — على ClassStyle وليس ExStyle!)
    ' ─────────────────────────────────────────
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = cp.ClassStyle Or &H20000   ' CS_DROPSHADOW — الظل تحت النافذة
            Return cp
        End Get
    End Property

    ' ─────────────────────────────────────────
    ' 🌟 حواف مستديرة خفيفة (مظهر عصري)
    ' ─────────────────────────────────────────
    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        Try
            Dim radius As Integer = 18
            Dim path As New GraphicsPath()
            path.AddArc(0, 0, radius, radius, 180, 90)
            path.AddArc(Me.Width - radius - 1, 0, radius, radius, 270, 90)
            path.AddArc(Me.Width - radius - 1, Me.Height - radius - 1, radius, radius, 0, 90)
            path.AddArc(0, Me.Height - radius - 1, radius, radius, 90, 90)
            path.CloseFigure()
            Me.Region = New Region(path)
        Catch
            ' لو فشل تقريب الحواف — نكمل بالمستطيل العادي
        End Try
    End Sub

    ' ─────────────────────────────────────────
    ' 🌟 دوال العرض العامة — تُستدعى من frmLogin أثناء التحميل
    ' ─────────────────────────────────────────

    ''' <summary>عرض الشاشة</summary>
    Public Sub ShowSplash()
        Me.Show()
        Me.Refresh()   ' 🌟 L-05: إعادة رسم متزامنة بلا مضخ رسائل — تمنع إعادة الدخول
    End Sub

    ''' <summary>تحديث الخطوة والنسبة (0-100)</summary>
    Public Sub SetProgress(percent As Integer, statusText As String)
        If percent < 0 Then percent = 0
        If percent > 100 Then percent = 100
        progressBar.Value = percent
        lblStatus.Text = statusText
        Me.Refresh()   ' 🌟 L-05: تحديث مرئي بلا مضخ رسائل
    End Sub

    ''' <summary>خطوة: انتظار قصير حتى تظهر بوضوح</summary>
    Public Async Function PauseAsync() As Task
        Await Task.Delay(MinStepDelay)
    End Function

    ''' <summary>إغلاق الشاشة مع تلاشٍ خفيف</summary>
    Public Sub CloseSplash()
        Try
            ' تلاشٍ خفيف (Fade Out)
            For opacity As Double = 1.0 To 0 Step -0.1
                Me.Opacity = opacity
                Threading.Thread.Sleep(30)
                Me.Refresh()   ' 🌟 L-05: إطار التلاشي يعاد رسمه متزامناً بلا مضخ رسائل
            Next
        Catch
        End Try
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frmSplash_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
