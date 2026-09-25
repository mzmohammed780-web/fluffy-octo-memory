Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports System.Diagnostics

' ═══════════════════════════════════════════════════════════════════
' 🌟 أولوية 4 — مصدّر تقارير HTML جاهز للطباعة/الـ PDF
' لماذا HTML وليس PDF مباشرة؟
'   * توليد PDF عربي RTL سليم يتطلب مكتبات خارجية (iTextSharp وغيرها) ودمج
'     خطوط عربية وتشكيل حروف معقد — إضافة ثقيلة ومحفوفة بالمخاطر لمشروع WinForms قائم.
'   * HTML المكتوب يدعم العربية 100% في أي متصفح، وزر الطباعة المدمج في
'     التقرير يفتح حوار "حفظ كـ PDF" — النتيجة نفسها بلا أي مكتبة إضافية.
' الاستخدام: ExportGridToHtml(grid, title) ثم يفتح التقرير في المتصفح تلقائياً.
' ═══════════════════════════════════════════════════════════════════
Public Class ReportExporter

    ''' <summary>هل النص يحتوي محارف تتطلب ترميز HTML؟ (تحقق سريع)</summary>
    Private Shared Function NeedsHtmlEscape(text As String) As Boolean
        Return text.IndexOf("<"c) >= 0 OrElse text.IndexOf(">"c) >= 0 OrElse
               text.IndexOf("&"c) >= 0 OrElse text.IndexOf(""""c) >= 0
    End Function

    ''' <summary>تهريب محارف HTML الخاصة لمنع كسر الترميز</summary>
    Private Shared Function EscapeHtml(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If Not NeedsHtmlEscape(text) Then Return text
        Return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;")
    End Function

    ''' <summary>
    ''' تصدير محتوى داتا جريد إلى تقرير HTML أنيق RTL جاهز للطباعة/الـ PDF، ثم فتحه في المتصفح.
    ''' </summary>
    ''' <param name="grid">الشبكة المصدّرة (الأعمدة المرئية فقط)</param>
    ''' <param name="title">عنوان التقرير الرئيسي</param>
    ''' <param name="subtitle">سطر فرعي اختياري (مثل نطاق التاريخ أو اسم اللاعب)</param>
    ''' <param name="footerNote">ملاحظة تذييل اختيارية (مثل شرح "كل الفترات")</param>
    ''' <param name="landscape">اتجاه الطباعة: True = عرضي (جداول عريضة)</param>
    ''' <param name="excludedColumnNames">أعمدة مستثناة (مثل عمود التحديد)</param>
    ''' <returns>مسار ملف HTML المُنشأ</returns>
    Public Shared Function ExportGridToHtml(grid As DataGridView,
                                            title As String,
                                            Optional subtitle As String = "",
                                            Optional footerNote As String = "",
                                            Optional landscape As Boolean = False,
                                            Optional excludedColumnNames As IEnumerable(Of String) = Nothing) As String
        If grid Is Nothing Then
            Throw New ArgumentNullException(NameOf(grid))
        End If
        If String.IsNullOrWhiteSpace(title) Then title = "تقرير"

        Dim excl As New HashSet(Of String)(If(excludedColumnNames, New String() {}), StringComparer.OrdinalIgnoreCase)

        ' ── 1. مسار الملف: مجلد التقارير بجانب قاعدة البيانات ──
        Dim reportsFolder As String = Path.Combine(DatabaseModule.GetAppDataPath(), "Reports")
        If Not Directory.Exists(reportsFolder) Then
            Directory.CreateDirectory(reportsFolder)
        End If
        Dim safeTitle As String = title
        For Each bad As Char In Path.GetInvalidFileNameChars()
            safeTitle = safeTitle.Replace(bad, "_"c)
        Next
        Dim filePath As String = Path.Combine(reportsFolder,
            $"{safeTitle}_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.html")

        ' ── 2. بناء جسم الجدول (أعمدة مرئية غير مستثناة فقط) ──
        Dim columns As New List(Of DataGridViewColumn)
        For Each col As DataGridViewColumn In grid.Columns
            If col.Visible AndAlso Not excl.Contains(col.Name) Then
                columns.Add(col)
            End If
        Next

        Dim headersSb As New StringBuilder()
        For Each col As DataGridViewColumn In columns
            headersSb.Append("<th>").Append(EscapeHtml(col.HeaderText)).Append("</th>")
        Next

        Dim rowsSb As New StringBuilder()
        Dim rowCount As Integer = 0
        For Each row As DataGridViewRow In grid.Rows
            If row.IsNewRow Then Continue For
            rowCount += 1
            rowsSb.Append("<tr>")
            For Each col As DataGridViewColumn In columns
                Dim cellText As String = ""
                Dim cell As DataGridViewCell = row.Cells(col.Name)
                If cell.Value IsNot Nothing AndAlso Not IsDBNull(cell.Value) Then
                    cellText = cell.Value.ToString()
                End If
                rowsSb.Append("<td>").Append(EscapeHtml(cellText)).Append("</td>")
            Next
            rowsSb.Append("</tr>")
        Next

        ' ── 3. القالب الكامل مع CSS طباعة احترافي ──
        Dim orientation As String = If(landscape, "landscape", "portrait")
        Dim subtitleHtml As String = If(String.IsNullOrWhiteSpace(subtitle), "",
            $"<div class='subtitle'>{EscapeHtml(subtitle)}</div>")
        Dim footerHtml As String = If(String.IsNullOrWhiteSpace(footerNote), "",
            $"<div class='footer-note'>{EscapeHtml(footerNote)}</div>")

        Dim html As New StringBuilder()
        html.AppendLine("<!DOCTYPE html>")
        html.AppendLine("<html dir='rtl' lang='ar'>")
        html.AppendLine("<head>")
        html.AppendLine("<meta charset='utf-8'>")
        html.AppendLine($"<title>{EscapeHtml(title)}</title>")
        html.AppendLine("<style>")
        html.AppendLine("  @page { size: A4 " & orientation & "; margin: 12mm; }")
        html.AppendLine("  * { box-sizing: border-box; }")
        html.AppendLine("  body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; direction: rtl; color: #1f2937; margin: 0; padding: 16px; background: #fff; }")
        html.AppendLine("  .report-header { text-align: center; border-bottom: 3px solid #283044; padding-bottom: 10px; margin-bottom: 14px; }")
        html.AppendLine("  h1 { margin: 0 0 4px 0; font-size: 20pt; color: #283044; }")
        html.AppendLine("  .subtitle { font-size: 11pt; color: #4b5563; }")
        html.AppendLine("  .meta { font-size: 9pt; color: #6b7280; margin-top: 6px; }")
        html.AppendLine("  table { width: 100%; border-collapse: collapse; font-size: 9.5pt; }")
        html.AppendLine("  th { background: #283044; color: #fff; padding: 7px 6px; border: 1px solid #1a1f2e; text-align: right; }")
        html.AppendLine("  td { padding: 6px; border: 1px solid #d1d5db; text-align: right; }")
        html.AppendLine("  tr:nth-child(even) td { background: #f3f4f6; }")
        html.AppendLine("  thead { display: table-header-group; }")   ' تكرار الترويسة في كل صفحة
        html.AppendLine("  tr { page-break-inside: avoid; }")
        html.AppendLine("  .footer-note { margin-top: 14px; font-size: 9pt; color: #6b7280; border-top: 1px solid #d1d5db; padding-top: 8px; }")
        html.AppendLine("  .print-bar { text-align: center; margin: 0 0 14px 0; }")
        html.AppendLine("  .print-btn { font-family: inherit; font-size: 12pt; font-weight: bold; padding: 10px 42px; background: #283044; color: #fff; border: none; border-radius: 8px; cursor: pointer; }")
        html.AppendLine("  @media print { .print-bar { display: none; } body { padding: 0; } }")
        html.AppendLine("</style>")
        html.AppendLine("</head>")
        html.AppendLine("<body>")
        html.AppendLine("<div class='print-bar'><button class='print-btn' onclick='window.print()'>طباعة / حفظ كـ PDF</button></div>")
        html.AppendLine("<div class='report-header'>")
        html.AppendLine($"<h1>{EscapeHtml(title)}</h1>")
        html.AppendLine(subtitleHtml)
        html.AppendLine($"<div class='meta'>تاريخ الإنشاء: {DateTime.Now:yyyy-MM-dd HH:mm} — المستخدم: {EscapeHtml(UserSession.CurrentUsername)} — عدد السجلات: {rowCount}</div>")
        html.AppendLine("</div>")
        html.AppendLine("<table>")
        html.AppendLine("<thead><tr>").Append(headersSb.ToString()).Append("</tr></thead>")
        html.AppendLine("<tbody>").Append(rowsSb.ToString()).Append("</tbody>")
        html.AppendLine("</table>")
        html.AppendLine(footerHtml)
        html.AppendLine("</body>")
        html.AppendLine("</html>")

        ' ── 4. الكتابة بترميز UTF-8 مع BOM ──
        File.WriteAllText(filePath, html.ToString(), New UTF8Encoding(True))

        ' ── 5. فتح التقرير في المتصفح الافتراضي ──
        Try
            Dim psi As New ProcessStartInfo(filePath) With {
                .UseShellExecute = True
            }
            Process.Start(psi)
        Catch ex As Exception
            ' فشل فتح المتصفح لا يبطل نجاح التصدير — المستخدم سيجد الملف في مجلد التقارير
            DatabaseModule.LogInfo("ReportExporter", $"تم إنشاء التقرير لكن فشل فتح المتصفح: {filePath}")
        End Try

        Return filePath
    End Function

End Class
