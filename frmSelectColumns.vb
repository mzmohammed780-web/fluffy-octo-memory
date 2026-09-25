Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms



' 🌟 ملاحظة توثيق (أولوية 3): اسم الملف "Form2.vb" لا يطابق الكلاس "frmSelectColumns" —
'    أعد تسمية الملف إلى frmSelectColumns.vb من Visual Studio
'    (كليك يمين على الملف ← Rename) دون أي تأثير على الكود.
Public Class frmSelectColumns

    ' 🌟 L-03: عنصر قائمة يحمل الاسم التقني للعمود ويعرض الترويسة —
    ' الاختيار بالاسم التقني يمنع الترويسات المكررة من تضليل الطباعة والتصدير
    Private Class ColumnItem
        Public Property ColName As String = ""
        Public Property Header As String = ""
        Public Overrides Function ToString() As String
            Return Header
        End Function
    End Class

#Region "Public Properties"

    ''' <summary>قائمة الأعمدة المختارة بأسمائها التقنية (col.Name) — المطابقة بالاسم أدق من الترويسة (L-03)</summary>
    Public SelectedColumns As New List(Of String)

    ''' <summary>"Print" أو "Export" — يحدد نص الأزرار وعنوان النموذج</summary>
    Public OperationType As String = "Export"

    ''' <summary>DataGridView المصدر الذي نقرأ أسماء أعمدته منه</summary>
    Public SourceDataGridView As DataGridView = Nothing

    ''' <summary>
    ''' قائمة الأعمدة التقنية التي لا يجب أن يراها المستخدم.
    ''' إذا تركتها فارغة، سيستخدم القائمة الافتراضية لشاشة اللاعبين.
    ''' </summary>
    Public Property ExcludedColumns As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

#End Region

#Region "Private Fields"

    Private _chkColumns As CheckedListBox

    ' الأعمدة الافتراضية المحددة عند فتح النموذج
    Private ReadOnly _defaultColumns As New HashSet(Of String) From {
        "رقم الهوية",
        "الاسم الكامل",
        "تاريخ الميلاد",
        "اسم الأب",
        "رقم الأب",
        "اسم الزوجة/الأم",
        "رقم الزوجة/الأم",
        "رقم الهاتف",
        "العنوان"
    }

#End Region

#Region "Form Load"

    Private Sub frmSelectColumns_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' التحقق من صحة المصدر
            If SourceDataGridView Is Nothing Then
                MessageBox.Show("لم يتم تحديد مصدر البيانات!", "خطأ",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.DialogResult = DialogResult.Cancel
                Return
            End If

            If SourceDataGridView.Columns.Count = 0 Then
                MessageBox.Show("لا توجد أعمدة في مصدر البيانات!", "تنبيه",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.DialogResult = DialogResult.Cancel
                Return
            End If

            ' إذا لم يمرر المبرمج قائمة مستثناة، نستخدم القيمة الافتراضية لشاشة اللاعبين
            If ExcludedColumns.Count = 0 Then
                ExcludedColumns = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "playerphoto", "DataSource", "SelectColumn", "SeqColumn",
                    "altfone", "Talented", "Exempted", "DeletedDate", "ID"
                }
            End If

            ConfigureForm()
            BuildUI()

            Me.DialogResult = DialogResult.None

        Catch ex As Exception
            DatabaseModule.LogError("frmSelectColumns_Load", ex)
            MessageBox.Show("خطأ في تحميل النموذج: " & ex.Message, "خطأ",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.DialogResult = DialogResult.Cancel
        End Try
    End Sub

    Private Sub ConfigureForm()
        Me.Text = If(OperationType = "Print", "اختر الأعمدة للطباعة", "اختر الأعمدة للتصدير")
        Me.Size = New Size(470, 620)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.White
        Me.RightToLeft = RightToLeft.Yes
        Me.RightToLeftLayout = True
        Me.Font = New Font("Arial", 9)
        Me.KeyPreview = True
    End Sub

    Private Sub BuildUI()
        Me.SuspendLayout()

        Const MARGIN As Integer = 20
        Const FORM_W As Integer = 430

        ' ── عنوان ──────────────────────────────────────────────
        Dim lblTitle As New Label() With {
            .Text = If(OperationType = "Print",
                       "✓ حدد الأعمدة التي تريد طباعتها:",
                       "✓ حدد الأعمدة التي تريد تصديرها:"),
            .Location = New Point(MARGIN, 18),
            .Size = New Size(FORM_W, 26),
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .ForeColor = Color.DarkBlue,
            .TextAlign = ContentAlignment.MiddleRight
        }
        Me.Controls.Add(lblTitle)

        ' ── قائمة الأعمدة (CheckedListBox) ─────────────────────
        _chkColumns = New CheckedListBox() With {
            .Name = "chkColumns",
            .Location = New Point(MARGIN, 52),
            .Size = New Size(FORM_W, 390),
            .Font = New Font("Arial", 9),
            .CheckOnClick = True,
            .BorderStyle = BorderStyle.FixedSingle,
            .RightToLeft = RightToLeft.Yes,
            .SelectionMode = SelectionMode.One
        }
        PopulateColumns()
        Me.Controls.Add(_chkColumns)

        ' ── لوحة أزرار الاختيار (تحديد الكل / إلغاء / عكس) ──
        Dim panelSelection As New Panel() With {
            .Location = New Point(MARGIN, 452),
            .Size = New Size(FORM_W, 40),
            .BorderStyle = BorderStyle.None
        }

        Dim btnSelectAll As New Button() With {
            .Text = "تحديد الكل",
            .Location = New Point(0, 0),
            .Size = New Size(130, 36),
            .BackColor = Color.FromArgb(173, 216, 230),
            .FlatStyle = FlatStyle.Flat
        }
        btnSelectAll.FlatAppearance.BorderSize = 0
        AddHandler btnSelectAll.Click, AddressOf BtnSelectAll_Click
        panelSelection.Controls.Add(btnSelectAll)

        Dim btnDeselectAll As New Button() With {
            .Text = "إلغاء الكل",
            .Location = New Point(140, 0),
            .Size = New Size(130, 36),
            .BackColor = Color.FromArgb(250, 180, 180),
            .FlatStyle = FlatStyle.Flat
        }
        btnDeselectAll.FlatAppearance.BorderSize = 0
        AddHandler btnDeselectAll.Click, AddressOf BtnDeselectAll_Click
        panelSelection.Controls.Add(btnDeselectAll)

        Dim btnInvert As New Button() With {
            .Text = "عكس الاختيار",
            .Location = New Point(280, 0),
            .Size = New Size(130, 36),
            .BackColor = Color.FromArgb(255, 255, 180),
            .FlatStyle = FlatStyle.Flat
        }
        btnInvert.FlatAppearance.BorderSize = 0
        AddHandler btnInvert.Click, AddressOf BtnInvert_Click
        panelSelection.Controls.Add(btnInvert)

        Me.Controls.Add(panelSelection)

        ' ── لوحة أزرار التأكيد والإلغاء ──────────────────────
        Dim panelConfirm As New Panel() With {
            .Location = New Point(MARGIN, 500),
            .Size = New Size(FORM_W, 50),
            .BorderStyle = BorderStyle.None
        }

        Dim btnConfirm As New Button() With {
            .Text = If(OperationType = "Print", "▶  معاينة الطباعة", "💾  تصدير"),
            .Location = New Point(0, 0),
            .Size = New Size(200, 44),
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .BackColor = If(OperationType = "Print",
                            Color.FromArgb(100, 180, 230),
                            Color.FromArgb(100, 210, 130)),
            .ForeColor = Color.Black,
            .FlatStyle = FlatStyle.Flat
        }
        btnConfirm.FlatAppearance.BorderSize = 0
        AddHandler btnConfirm.Click, AddressOf BtnConfirm_Click
        Me.AcceptButton = btnConfirm
        panelConfirm.Controls.Add(btnConfirm)

        Dim btnCancel As New Button() With {
            .Text = "✖  إلغاء",
            .Location = New Point(215, 0),
            .Size = New Size(200, 44),
            .Font = New Font("Arial", 10),
            .BackColor = Color.FromArgb(200, 200, 200),
            .FlatStyle = FlatStyle.Flat,
            .DialogResult = DialogResult.Cancel
        }
        btnCancel.FlatAppearance.BorderSize = 0
        Me.CancelButton = btnCancel
        panelConfirm.Controls.Add(btnCancel)

        Me.Controls.Add(panelConfirm)

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Private Sub PopulateColumns()
        Try
            _chkColumns.Items.Clear()

            ' 🌟 التحقق هل نحن في شاشة اللاعبين أم شاشة أخرى
            Dim isPlayerGrid As Boolean = SourceDataGridView.Columns.Contains(AppConstants.Col_PlayerId)

            For Each col As DataGridViewColumn In SourceDataGridView.Columns
                If col.Visible AndAlso
                   Not ExcludedColumns.Contains(col.Name) AndAlso
                   Not String.IsNullOrWhiteSpace(col.HeaderText) Then

                    ' إذا كان شبكة اللاعبين، نستخدم القائمة الافتراضية. إذا لا، نحدد الكل (True)
                    Dim isChecked As Boolean = If(isPlayerGrid, _defaultColumns.Contains(col.HeaderText), True)

                    _chkColumns.Items.Add(New ColumnItem With {.ColName = col.Name, .Header = col.HeaderText}, isChecked)
                End If
            Next

            If _chkColumns.Items.Count = 0 Then
                _chkColumns.Items.Add("لا توجد أعمدة للعرض", False)
                _chkColumns.Enabled = False
            End If

        Catch ex As Exception
            DatabaseModule.LogError("frmSelectColumns.PopulateColumns", ex)
            MessageBox.Show("خطأ في تحميل قائمة الأعمدة: " & ex.Message, "خطأ",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
            _chkColumns.Items.Clear()
            _chkColumns.Items.Add("حدث خطأ أثناء تحميل الأعمدة", False)
            _chkColumns.Enabled = False
        End Try
    End Sub
#End Region

#Region "Button Click Handlers"

    Private Sub BtnSelectAll_Click(sender As Object, e As EventArgs)
        If Not _chkColumns.Enabled Then Return
        For i As Integer = 0 To _chkColumns.Items.Count - 1
            _chkColumns.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub BtnDeselectAll_Click(sender As Object, e As EventArgs)
        If Not _chkColumns.Enabled Then Return
        For i As Integer = 0 To _chkColumns.Items.Count - 1
            _chkColumns.SetItemChecked(i, False)
        Next
    End Sub

    Private Sub BtnInvert_Click(sender As Object, e As EventArgs)
        If Not _chkColumns.Enabled Then Return
        For i As Integer = 0 To _chkColumns.Items.Count - 1
            _chkColumns.SetItemChecked(i, Not _chkColumns.GetItemChecked(i))
        Next
    End Sub

    Private Sub BtnConfirm_Click(sender As Object, e As EventArgs)
        Try
            SelectedColumns.Clear()
            For i As Integer = 0 To _chkColumns.CheckedItems.Count - 1
                Dim itemObj As Object = _chkColumns.CheckedItems(i)
                Dim colItem As ColumnItem = TryCast(itemObj, ColumnItem)
                If colItem IsNot Nothing Then
                    SelectedColumns.Add(colItem.ColName)       ' الاسم التقني — L-03
                Else
                    SelectedColumns.Add(itemObj.ToString())    ' احتياط لعناصر نصية استثنائية
                End If
            Next

            If SelectedColumns.Count = 0 Then
                Dim action As String = If(OperationType = "Print", "الطباعة", "التصدير")
                MessageBox.Show($"الرجاء تحديد عمود واحد على الأقل لـ{action}", "تنبيه",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            DatabaseModule.LogError("frmSelectColumns.BtnConfirm_Click", ex)
            MessageBox.Show("خطأ: " & ex.Message, "خطأ",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Keyboard Handler"

    Private Sub frmSelectColumns_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
            e.Handled = True
        End If
    End Sub

#End Region

End Class
