Option Explicit On
Option Strict On

Public Module UserSession
    Public Property CurrentUsername As String = ""
    Public Property CanEdit As Boolean = False
    Public Property CanDelete As Boolean = False

    ' 🌟 إضافة: هل يجب على المستخدم تغيير كلمة مروره؟ (تُقرأ من قاعدة البيانات عند تسجيل الدخول)
    ' كانت تُستخدم في frmLogin و UserService بدون تعريف — وهذا يمنع المشروع من التجميع أصلاً
    Public Property MustChangePassword As Boolean = False

    ' متغير لتحديد ما إذا كان المستخدم يريد تسجيل الخروج أم إغلاق البرنامج
    Public Property IsLoggingOut As Boolean = False

    ''' <summary>تنظيف بيانات الجلسة الحالية</summary>
    Public Sub ClearSession()
        CurrentUsername = ""
        CanEdit = False
        CanDelete = False
        MustChangePassword = False
        IsLoggingOut = False
    End Sub
End Module
