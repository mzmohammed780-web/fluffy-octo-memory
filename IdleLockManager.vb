Option Explicit On
Option Strict On

Imports System.Windows.Forms

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — مدير القفل التلقائي عند الخمول
' 🌟 إصلاح v2: إذا فشل تثبيت مرشّح الرسائل عند Startup (بيئة محدودة)،
'              تُعاد المحاولة تلقائياً داخل TimerTick حتى ينجح.
' ═══════════════════════════════════════════════════════════════════
Public Class IdleLockManager
    Implements IMessageFilter

    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_MBUTTONDOWN As Integer = &H207
    Private Const WM_MOUSEWHEEL As Integer = &H20A
    Private Const WM_KEYDOWN As Integer = &H100

    Private ReadOnly _activityLock As New Object()
    Private _lastActivity As DateTime = DateTime.Now
    Private _timer As Windows.Forms.Timer
    Private _isLockShown As Boolean = False
    Private _installed As Boolean = False
    Private _startupRequested As Boolean = False   ' 🌟 يُرفع في Startup مهما كانت النتيجة

    Private Shared ReadOnly _instance As New IdleLockManager()
    Public Shared ReadOnly Property Instance As IdleLockManager
        Get
            Return _instance
        End Get
    End Property

    Private Sub New()
    End Sub

    Private _cachedIdleMinutes As Integer = -1
    Private ReadOnly _configLock As New Object()

    Public ReadOnly Property IdleMinutesConfigured As Integer
        Get
            SyncLock _configLock
                If _cachedIdleMinutes < 0 Then
                    _cachedIdleMinutes = AppSettingsStore.GetIntSetting(AppSettingsStore.Key_IdleLockMinutes, 15, 0, 720)
                End If
                Return _cachedIdleMinutes
            End SyncLock
        End Get
    End Property

    Public Sub RefreshConfiguration()
        SyncLock _configLock
            _cachedIdleMinutes = AppSettingsStore.GetIntSetting(AppSettingsStore.Key_IdleLockMinutes, 15, 0, 720)
        End SyncLock
    End Sub

    Public Sub Startup()
        _startupRequested = True
        Try
            If _installed Then Return
            If IdleMinutesConfigured <= 0 Then Return

            TryInstallMessageFilter()

            ResetActivity()

            If _timer Is Nothing Then
                _timer = New Windows.Forms.Timer()
                _timer.Interval = 20000
                AddHandler _timer.Tick, AddressOf TimerTick
                _timer.Start()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("IdleLockManager.Startup", ex)
        End Try
    End Sub

    ''' <summary>🌟 إصلاح v2: محاولة تثبيت مستقلة قابلة لإعادة الاستدعاء</summary>
    Private Sub TryInstallMessageFilter()
        If _installed Then Return
        Try
            Application.AddMessageFilter(Me)
            _installed = True
        Catch ex As Exception
            ' يُترك _installed = False — TimerTick سيعيد المحاولة تلقائياً
            DatabaseModule.LogError("IdleLockManager.TryInstallMessageFilter", ex)
        End Try
    End Sub

    Public Sub Shutdown()
        Try
            If _timer IsNot Nothing Then
                _timer.Stop()
                _timer.Dispose()
                _timer = Nothing
            End If
            If _installed Then
                Application.RemoveMessageFilter(Me)
                _installed = False
            End If
            _startupRequested = False
        Catch ex As Exception
            DatabaseModule.LogError("IdleLockManager.Shutdown", ex)
        End Try
    End Sub

    Public Sub ResetActivity()
        SyncLock _activityLock
            _lastActivity = DateTime.Now
        End SyncLock
    End Sub

    Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
        Select Case m.Msg
            Case WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_MOUSEWHEEL, WM_KEYDOWN
                If Not _isLockShown Then
                    SyncLock _activityLock
                        _lastActivity = DateTime.Now
                    End SyncLock
                End If
        End Select
        Return False
    End Function

    Private Sub TimerTick(sender As Object, e As EventArgs)
        Try
            ' 🌟 إصلاح v2: إن لم يُثبَّت المرشّح سابقاً لأي سبب، نحاول الآن
            If _startupRequested AndAlso Not _installed AndAlso IdleMinutesConfigured > 0 Then
                TryInstallMessageFilter()
            End If

            If _isLockShown Then Return
            If IdleMinutesConfigured <= 0 Then Return

            Dim idleMinutes As Double
            SyncLock _activityLock
                idleMinutes = (DateTime.Now - _lastActivity).TotalMinutes
            End SyncLock

            If idleMinutes >= CDbl(IdleMinutesConfigured) Then
                ShowLockScreen()
            End If
        Catch ex As Exception
            DatabaseModule.LogError("IdleLockManager.TimerTick", ex)
        End Try
    End Sub

    Public Sub ShowLockScreen()
        Try
            If _isLockShown Then Return
            _isLockShown = True

            Dim audit As New AuditService()
            audit.Log(AuditService.Act_SessionLocked, "جلسة", "", $"خمول لمدة {IdleMinutesConfigured} دقيقة")

            Dim owner As Form = Form.ActiveForm

            Using lockForm As New frmLock()
                lockForm.ShowDialog(owner)
            End Using

            ResetActivity()
        Catch ex As Exception
            DatabaseModule.LogError("IdleLockManager.ShowLockScreen", ex)
        Finally
            _isLockShown = False
        End Try
    End Sub

End Class