Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.Threading.Tasks
Imports System.Data ' 🌟 تمت الإضافة لتعريف DataTable

Public Class frmManageUsers
    Inherits System.Windows.Forms.Form

    Private dgvUsers As New DataGridView()
    Private btnSaveChanges As New Button()
    Private btnDeleteUser As New Button()
    Private btnClose As New Button()
    Private lblTitle As New Label()

    ' عناصر إضافة مستخدم جديد في الأسفل
    Private txtNewUser As New TextBox()
    Private txtNewPass As New TextBox()
    Private chkNewEdit As New CheckBox()
    Private chkNewDelete As New CheckBox()
    Private btnAddNew As New Button()

    Public Sub New()
        ' 🌟 هذا السطر مطلوب من مصمم الواجهات ويجب أن يكون الأول
        InitializeComponent()
        SetupUI()
    End Sub

    Private Async Sub frmManageUsers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 🌟 إصلاح أمني: الحماية كانت في معالج القائمة فقط — أي مسار مستقبلي يفتح النموذج
        ' يمنح إدارة كاملة للمستخدمين. بوابة داخلية الآن
        If Not String.Equals(UserSession.CurrentUsername, "admin", StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("هذه الشاشة لمدير النظام فقط", "مرفوض", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Close()
            Return
        End If
        Await LoadUsersAsync()
    End Sub

    Private Sub SetupUI()
        Me.Text = "إدارة المستخدمين والصلاحيات"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(600, 550)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.RightToLeft = RightToLeft.Yes
        Me.RightToLeftLayout = True
        Me.BackColor = Color.FromArgb(240, 240, 240)

        ' العنوان
        lblTitle.Text = "قائمة المستخدمين (ضع علامة صح لإعطاء الصلاحية أو أزلها لإلغائها)"
        lblTitle.AutoSize = False
        lblTitle.Size = New Size(560, 30)
        lblTitle.Location = New Point(15, 10)
        lblTitle.Font = New Font("Arial", 10, FontStyle.Bold)
        lblTitle.TextAlign = ContentAlignment.MiddleCenter
        Me.Controls.Add(lblTitle)

        ' إعداد الجدول
        dgvUsers.Location = New Point(15, 45)
        dgvUsers.Size = New Size(560, 300)
        dgvUsers.AllowUserToAddRows = False
        dgvUsers.AllowUserToDeleteRows = False
        dgvUsers.ReadOnly = False
        dgvUsers.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvUsers.RightToLeft = RightToLeft.Yes
        Me.Controls.Add(dgvUsers)

        ' أزرار الإدارة (تحت الجدول)
        btnSaveChanges.Text = "💾 حفظ التعديلات"
        btnSaveChanges.Size = New Size(150, 40)
        btnSaveChanges.Location = New Point(15, 355)
        btnSaveChanges.BackColor = Color.FromArgb(100, 210, 130)   ' 🌟 توحيد أخضر الحفظ مع باقي الشاشات
        btnSaveChanges.FlatStyle = FlatStyle.Flat
        AddHandler btnSaveChanges.Click, AddressOf DoSaveChanges
        Me.Controls.Add(btnSaveChanges)

        btnDeleteUser.Text = "🗑️ حذف المستخدم المحدد"
        btnDeleteUser.Size = New Size(150, 40)
        btnDeleteUser.Location = New Point(175, 355)
        btnDeleteUser.BackColor = Color.LightCoral
        btnDeleteUser.FlatStyle = FlatStyle.Flat
        AddHandler btnDeleteUser.Click, AddressOf DoDeleteUser
        Me.Controls.Add(btnDeleteUser)

        btnClose.Text = "إغلاق"
        btnClose.Size = New Size(100, 40)
        btnClose.Location = New Point(335, 355)
        btnClose.FlatStyle = FlatStyle.Flat
        AddHandler btnClose.Click, Sub() Me.Close()
        Me.Controls.Add(btnClose)

        ' --- قسم إضافة مستخدم جديد في الأسفل ---
        Dim lblAddTitle As New Label() With {.Text = "إضافة مستخدم جديد:", .Location = New Point(15, 410), .Font = New Font("Arial", 10, FontStyle.Bold), .AutoSize = True}
        Me.Controls.Add(lblAddTitle)

        Dim lblUser As New Label() With {.Text = "الاسم:", .Location = New Point(420, 445), .AutoSize = True}
        txtNewUser.Location = New Point(300, 442)
        txtNewUser.Width = 115
        Me.Controls.AddRange({lblUser, txtNewUser})

        Dim lblPass As New Label() With {.Text = "المرور:", .Location = New Point(420, 475), .AutoSize = True}
        txtNewPass.Location = New Point(300, 472)
        txtNewPass.Width = 115
        txtNewPass.PasswordChar = "*"c
        Me.Controls.AddRange({lblPass, txtNewPass})

        chkNewEdit.Text = "صلاحية إضافة/تعديل"
        chkNewEdit.Location = New Point(150, 445)
        chkNewEdit.Checked = True
        Me.Controls.Add(chkNewEdit)

        chkNewDelete.Text = "صلاحية حذف"
        chkNewDelete.Location = New Point(150, 475)
        chkNewDelete.Checked = True
        Me.Controls.Add(chkNewDelete)

        btnAddNew.Text = "➕ إضافة"
        btnAddNew.Size = New Size(100, 40)
        btnAddNew.Location = New Point(15, 445)
        btnAddNew.BackColor = Color.LightBlue
        btnAddNew.FlatStyle = FlatStyle.Flat
        AddHandler btnAddNew.Click, AddressOf DoAddNewUser
        Me.Controls.Add(btnAddNew)
    End Sub

    Private Async Function LoadUsersAsync() As Task
        Try
            Dim service As New UserService()
            Dim dt As DataTable = Await service.GetAllUsersAsync()
            dgvUsers.DataSource = dt

            ' تنسيق أعمدة الجدول
            If dgvUsers.Columns.Contains(AppConstants.Col_Username) Then
                dgvUsers.Columns(AppConstants.Col_Username).HeaderText = "اسم المستخدم"
                dgvUsers.Columns(AppConstants.Col_Username).ReadOnly = True ' لا يمكن تعديل الاسم من هنا
            End If
            If dgvUsers.Columns.Contains(AppConstants.Col_CanEdit) Then
                dgvUsers.Columns(AppConstants.Col_CanEdit).HeaderText = "صلاحية الإضافة/التعديل"
            End If
            If dgvUsers.Columns.Contains(AppConstants.Col_CanDelete) Then
                dgvUsers.Columns(AppConstants.Col_CanDelete).HeaderText = "صلاحية الحذف"
            End If

            ' 🌟 حماية حساب المدير: صف الأدمن لا يقبل تعديل الصلاحيات إطلاقاً
            If dgvUsers.Columns.Contains(AppConstants.Col_Username) Then
                For Each row As DataGridViewRow In dgvUsers.Rows
                    If String.Equals(UtilityModule.SafeString(row.Cells(AppConstants.Col_Username).Value), "admin", StringComparison.OrdinalIgnoreCase) Then
                        row.Cells(AppConstants.Col_CanEdit).ReadOnly = True
                        row.Cells(AppConstants.Col_CanDelete).ReadOnly = True
                        row.Cells(AppConstants.Col_CanEdit).ToolTipText = "صلاحيات المدير الرئيسي محمية"
                        row.Cells(AppConstants.Col_CanDelete).ToolTipText = "صلاحيات المدير الرئيسي محمية"
                    End If
                Next
            End If

        Catch ex As Exception
            MessageBox.Show("خطأ في تحميل المستخدمين: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function

    ' 🌟 إصلاح: حاجز تكرار — بعد أول Await تعود التحكم لمضخة الرسائل والأزرار مفعّلة،
    ' فنقرة ثانية كانت تطلق معاملة ثانية متزامنة (حفظان/حذفان متداخلان)
    Private _isBusy As Boolean = False

    Private Async Sub DoSaveChanges(sender As Object, e As EventArgs)
        If _isBusy Then Return
        _isBusy = True
        Try
            ' 🌟 إجبار الـ DataGridView على حفظ التعديلات في الخلية الحالية قبل القراءة
            dgvUsers.EndEdit()

            Dim usersToUpdate As New List(Of Tuple(Of String, Boolean, Boolean))
            ' تجميع كل البيانات المعدلة في قائمة واحدة — مع ضمان بقاء صلاحيات الأدمن كاملة
            For Each row As DataGridViewRow In dgvUsers.Rows
                Dim username As String = UtilityModule.SafeString(row.Cells(AppConstants.Col_Username).Value)

                If String.Equals(username, "admin", StringComparison.OrdinalIgnoreCase) Then
                    ' 🌟 الأدمن دائماً بصلاحيات كاملة — نتجاهل أي حالة معروضة بالجدول
                    usersToUpdate.Add(Tuple.Create(username, True, True))
                Else
                    ' 🌟 إصلاح انهيار: كان Convert.ToBoolean يرمي InvalidCastException إذا كانت
                    ' الخلية NULL في قاعدة البيانات فيفشل الحفظ كله في كل مرة
                    Dim canEdit As Boolean = SafeBoolCell(row.Cells(AppConstants.Col_CanEdit).Value)
                    Dim canDelete As Boolean = SafeBoolCell(row.Cells(AppConstants.Col_CanDelete).Value)
                    usersToUpdate.Add(Tuple.Create(username, canEdit, canDelete))
                End If
            Next

            If usersToUpdate.Count = 0 Then
                MessageBox.Show("لا يوجد مستخدمين للتحديث.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' --- تنفيذ التحديث دفعة واحدة ---
            Dim service As New UserService()
            ' 🌟 تم الإصلاح: استدعاء الدالة Async مباشرة بدون Task.Run
            Dim success As Boolean = Await service.UpdateMultipleUsersPermissions(usersToUpdate)

            If success Then
                MessageBox.Show($"تم تحديث صلاحيات {usersToUpdate.Count} مستخدم بنجاح.", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("خطأ أثناء الحفظ: فشل تحديث الصلاحيات (تم إلغاء العملية).", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            DatabaseModule.LogError("frmManageUsers.DoSaveChanges", ex)
            MessageBox.Show("خطأ أثناء الحفظ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _isBusy = False
        End Try
    End Sub


    Private Async Sub DoDeleteUser(sender As Object, e As EventArgs)
        If _isBusy Then Return
        _isBusy = True
        Try
            If dgvUsers.CurrentRow Is Nothing Then
            MessageBox.Show("الرجاء تحديد مستخدم من الجدول أولاً.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim username As String = UtilityModule.SafeString(dgvUsers.CurrentRow.Cells(AppConstants.Col_Username).Value)

        If String.Equals(username.Trim(), "admin", StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("لا يمكن حذف حساب المدير الرئيسي (admin).", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 🌟 إصلاح: منع المستخدم من حذف حسابه الحالي المسجّل به — كان يترك جلسة يتيمة
        If String.Equals(username.Trim(), UserSession.CurrentUsername.Trim(), StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("لا يمكنك حذف الحساب الذي تسجّل الدخول به حالياً.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If MessageBox.Show($"هل أنت متأكد من حذف المستخدم: {username}؟", "تأكيد الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Dim service As New UserService()
                Dim result As Integer = Await service.DeleteUserAsync(username)

                If result > 0 Then
                    MessageBox.Show("تم حذف المستخدم بنجاح.", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Await LoadUsersAsync() ' تحديث الجدول
                Else
                    MessageBox.Show("فشل الحذف.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Catch ex As Exception
                MessageBox.Show("خطأ أثناء الحذف: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

        Finally
            _isBusy = False
        End Try
    End Sub

    ''' <summary>قراءة خلية إذن من الجدول بأمان (NULL/فارغ/نص = False)</summary>
    Private Function SafeBoolCell(value As Object) As Boolean
        If value Is Nothing OrElse IsDBNull(value) Then Return False
        If TypeOf value Is Boolean Then Return CBool(value)
        Dim s As String = value.ToString().Trim()
        Return s = "1" OrElse s = "True" OrElse s = "true"
    End Function

    Private Async Sub DoAddNewUser(sender As Object, e As EventArgs)
        If _isBusy Then Return
        _isBusy = True
        Try
        If String.IsNullOrWhiteSpace(txtNewUser.Text) OrElse String.IsNullOrWhiteSpace(txtNewPass.Text) Then
            MessageBox.Show("الرجاء إدخال اسم المستخدم وكلمة المرور للمستخدم الجديد.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 🌟 فرض سياسة طول كلمة المرور
        ' 🌟 توحيد السياسة: 8 أحرف مثل شاشة تغيير كلمة المرور — كان يقبل 6
        If txtNewPass.Text.Length < 8 Then
            MessageBox.Show("كلمة المرور يجب أن تكون 8 أحرف على الأقل.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNewPass.Focus()
            Return
        End If

        Try
            Dim service As New UserService()
            Dim success As Boolean = Await service.AddUserAsync(txtNewUser.Text, txtNewPass.Text, chkNewEdit.Checked, chkNewDelete.Checked)

            If success Then
                MessageBox.Show("تمت إضافة المستخدم بنجاح.", "نجاح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtNewUser.Clear()
                txtNewPass.Clear()
                chkNewEdit.Checked = True
                chkNewDelete.Checked = True
                Await LoadUsersAsync() ' تحديث الجدول
            Else
                MessageBox.Show("فشل الإضافة (ربما اسم المستخدم موجود مسبقاً).", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("خطأ: " & ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Finally
            _isBusy = False
        End Try
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'frmManageUsers
        '
        Me.ClientSize = New System.Drawing.Size(600, 550) ' 🌟 تم تعديل الحجم ليتطابق مع SetupUI
        Me.Name = "frmManageUsers"
        Me.ResumeLayout(False)
    End Sub

End Class
