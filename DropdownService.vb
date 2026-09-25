Option Explicit On
Option Strict On

Imports System.Data
Imports System.Data.SQLite
Imports System.Threading.Tasks

''' <summary>
''' إدارة عناصر القوائم المنسدلة المخصصة — تُخزن بقاعدة البيانات نفسها
''' فتنتقل مع النسخ الاحتياطي والاستعادة والنقل بين الأجهزة
''' </summary>
Public Class DropdownService

    ''' <summary>جلب عناصر قائمة معينة (مثلاً "Team") مرتبة أبجدياً</summary>
    Public Async Function GetItemsAsync(fieldName As String) As Task(Of List(Of String))
        Dim result As New List(Of String)
        Try
            Dim sql As String = "SELECT ItemValue FROM DropdownItems WHERE FieldName = @f ORDER BY ItemValue"
            Dim dt As DataTable = Await DatabaseModule.ExecuteQueryAsync(sql, New SQLiteParameter("@f", fieldName))

            If dt IsNot Nothing Then
                For Each row As DataRow In dt.Rows
                    Dim v As String = UtilityModule.SafeString(row("ItemValue"))
                    If v <> "" Then result.Add(v)
                Next
            End If
        Catch ex As Exception
            DatabaseModule.LogError("DropdownService.GetItemsAsync", ex)
        End Try
        Return result
    End Function

    ''' <summary>إضافة عنصر جديد لقائمة — يتجاهل بصمت إذا كان موجوداً (مفتاح مركب يحمي من التكرار)</summary>
    Public Async Function AddItemAsync(fieldName As String, itemValue As String) As Task
        Try
            If String.IsNullOrWhiteSpace(itemValue) Then Return

            Dim sql As String = "INSERT OR IGNORE INTO DropdownItems (FieldName, ItemValue) VALUES (@f, @v)"
            Await DatabaseModule.ExecuteNonQueryAsync(sql,
                New SQLiteParameter("@f", fieldName),
                New SQLiteParameter("@v", itemValue.Trim()))
        Catch ex As Exception
            DatabaseModule.LogError("DropdownService.AddItemAsync", ex)
        End Try
    End Function

    ''' <summary>حذف عنصر من قائمة (لاستخدام مستقبلي إن أضفت شاشة إدارة للقوائم)</summary>
    Public Async Function DeleteItemAsync(fieldName As String, itemValue As String) As Task
        Try
            Dim sql As String = "DELETE FROM DropdownItems WHERE FieldName = @f AND ItemValue = @v"
            Await DatabaseModule.ExecuteNonQueryAsync(sql,
                New SQLiteParameter("@f", fieldName),
                New SQLiteParameter("@v", itemValue.Trim()))
        Catch ex As Exception
            DatabaseModule.LogError("DropdownService.DeleteItemAsync", ex)
        End Try
    End Function

End Class
