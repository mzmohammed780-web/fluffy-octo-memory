Option Explicit On
Option Strict On

Imports System.Windows.Forms

''' <summary>
''' دوال مشتركة للتعامل مع DataGridView — تُستخدم من frmPlayers وfrmPayments وfrmExpenses
''' لتوحيد منطق التحديد والترقيم وحالة رأس عمود التحديد
''' </summary>
Public Module GridHelper

    ''' <summary>جلب الصفوف المحددة بعلامة ✔ (مع تجاهل القيم الفارغة/DBNull بأمان)</summary>
    Public Function GetCheckedRows(grid As DataGridView) As List(Of DataGridViewRow)
        Dim result As New List(Of DataGridViewRow)
        If grid Is Nothing OrElse Not grid.Columns.Contains(AppConstants.Grid_SelectColumn) Then Return result

        For Each row As DataGridViewRow In grid.Rows
            If Not row.IsNewRow Then
                Dim val = row.Cells(AppConstants.Grid_SelectColumn).Value
                If val IsNot Nothing AndAlso Not IsDBNull(val) AndAlso Convert.ToBoolean(val) Then
                    result.Add(row)
                End If
            End If
        Next
        Return result
    End Function

    ''' <summary>حساب عدد الصفوف المحددة بدون بناء قائمة (أسرع للعرض فقط)</summary>
    Public Function CountCheckedRows(grid As DataGridView) As Integer
        If grid Is Nothing OrElse Not grid.Columns.Contains(AppConstants.Grid_SelectColumn) Then Return 0

        Dim count As Integer = 0
        For Each row As DataGridViewRow In grid.Rows
            If Not row.IsNewRow Then
                Dim val = row.Cells(AppConstants.Grid_SelectColumn).Value
                If val IsNot Nothing AndAlso Not IsDBNull(val) AndAlso Convert.ToBoolean(val) Then
                    count += 1
                End If
            End If
        Next
        Return count
    End Function

    ''' <summary>تحديث رأس عمود التحديد: □ (لا شيء) / ◑ (جزئي) / ✔ (الكل)</summary>
    Public Sub UpdateSelectHeaderState(grid As DataGridView)
        Try
            If grid Is Nothing OrElse Not grid.Columns.Contains(AppConstants.Grid_SelectColumn) Then Return

            Dim totalRows As Integer = 0
            Dim checkedRows As Integer = 0

            For Each row As DataGridViewRow In grid.Rows
                If Not row.IsNewRow Then
                    totalRows += 1
                    Dim val = row.Cells(AppConstants.Grid_SelectColumn).Value
                    If val IsNot Nothing AndAlso Not IsDBNull(val) AndAlso Convert.ToBoolean(val) Then
                        checkedRows += 1
                    End If
                End If
            Next

            If totalRows = 0 OrElse checkedRows = 0 Then
                grid.Columns(AppConstants.Grid_SelectColumn).HeaderText = "□"
            ElseIf checkedRows = totalRows Then
                grid.Columns(AppConstants.Grid_SelectColumn).HeaderText = "✔"
            Else
                grid.Columns(AppConstants.Grid_SelectColumn).HeaderText = "◑"
            End If
        Catch ex As Exception
            DatabaseModule.LogError("GridHelper.UpdateSelectHeaderState", ex)
        End Try
    End Sub

    ''' <summary>تحديد/إلغاء تحديد كل الصفوف (نقر رأس العمود) ويعيد الحالة الجديدة</summary>
    Public Function ToggleAllRows(grid As DataGridView) As Boolean
        If grid Is Nothing OrElse Not grid.Columns.Contains(AppConstants.Grid_SelectColumn) Then Return False

        Dim newValue As Boolean = (grid.Columns(AppConstants.Grid_SelectColumn).HeaderText <> "✔")
        For Each row As DataGridViewRow In grid.Rows
            If Not row.IsNewRow Then
                row.Cells(AppConstants.Grid_SelectColumn).Value = newValue
            End If
        Next
        grid.Columns(AppConstants.Grid_SelectColumn).HeaderText = If(newValue, "✔", "□")
        Return newValue
    End Function

    ''' <summary>تعبئة عمود الترقيم التسلسلي</summary>
    Public Sub FillSequenceNumbers(grid As DataGridView)
        Try
            If grid Is Nothing OrElse Not grid.Columns.Contains(AppConstants.Grid_SeqColumn) Then Return
            For i As Integer = 0 To grid.Rows.Count - 1
                grid.Rows(i).Cells(AppConstants.Grid_SeqColumn).Value = i + 1
            Next
        Catch ex As Exception
            DatabaseModule.LogError("GridHelper.FillSequenceNumbers", ex)
        End Try
    End Sub

    ''' <summary>
    ''' إعداد عام موحد للشبكة (الاتجاه، الرؤوس، الألوان، الـ Double Buffering)
    ''' مع معاملات للاختلافات المقصودة بين الشاشات الثلاث
    ''' </summary>
    Public Sub ApplyCommonGridStyle(grid As DataGridView,
                                    Optional multiSelect As Boolean = True,
                                    Optional readOnlyGrid As Boolean = False,
                                    Optional cellFontSize As Single = 8.0F)
        If grid Is Nothing Then Return
        grid.RightToLeft = RightToLeft.Yes
        grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.MultiSelect = multiSelect
        grid.ReadOnly = readOnlyGrid
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        grid.EnableHeadersVisualStyles = False
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(64, 64, 64)
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        grid.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
        grid.ColumnHeadersHeight = 35
        grid.DefaultCellStyle.Font = New Font("Arial", cellFontSize)
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(135, 206, 250)
        grid.DefaultCellStyle.SelectionForeColor = Color.Black
        grid.RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(173, 216, 230)
        grid.RowsDefaultCellStyle.BackColor = Color.White
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255)

        ' تفعيل Double Buffering عبر الانعكاس (يمنع الوميض أثناء التمرير)
        ' 🌟 L-06: فحص القيمة الراجعة قبل الاستخدام — يمنع استثناء نظري في بيئات محدودة
        Dim dblBufProp As Reflection.PropertyInfo = grid.GetType().GetProperty("DoubleBuffered",
            Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        If dblBufProp IsNot Nothing Then
            dblBufProp.SetValue(grid, True, Nothing)
        End If
    End Sub
End Module
