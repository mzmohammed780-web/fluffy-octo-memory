Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Linq

Public Class PrintHelper

    Implements IDisposable

    Public Property SourceGrid As DataGridView
    Public Property RowsToPrint As List(Of DataGridViewRow)
    Public Property ColumnsToPrint As List(Of String)
    Public Property ReportTitle As String = "تقرير"
    Public Property FooterSummary As String = ""

    Private _currentPrintRowIndex As Integer = 0
    Private _currentPageNumber As Integer = 0

    ' متغيرات مستوى الكلاس لحفظ الأبعاد وحسابها مرة واحدة فقط
    Private _colWidths As Dictionary(Of String, Integer)
    Private _totalWidth As Integer = 0
    Private Const SeqWidth As Integer = 40

    Private ReadOnly _titleFont As New Font("Arial", 16, FontStyle.Bold)
    Private ReadOnly _headerFont As New Font("Arial", 10, FontStyle.Bold)
    Private ReadOnly _cellFont As New Font("Arial", 9)
    Private ReadOnly _smallFont As New Font("Arial", 8)
    Private ReadOnly _totalFont As New Font("Arial", 10, FontStyle.Bold)

    Public Sub New()
        _currentPrintRowIndex = 0
        _currentPageNumber = 0
        _colWidths = New Dictionary(Of String, Integer)
    End Sub

    ''' <summary>تصفير حالة الطباعة (يضمن بداية نظيفة عند كل معاينة)</summary>
    Public Sub ResetPrintState()
        _currentPrintRowIndex = 0
        _currentPageNumber = 0
        _colWidths.Clear()
        _totalWidth = 0
    End Sub

    ''' <summary>دالة موحدة لفتح نافذة المعاينة قبل الطباعة وإدارة الموارد بأمان</summary>
    Public Sub ShowPreview()
        ResetPrintState()   ' 🌟 بداية نظيفة

        Using prnDoc As New Printing.PrintDocument()
            AddHandler prnDoc.PrintPage, AddressOf Me.PrintPage
            prnDoc.DefaultPageSettings.Landscape = True
            prnDoc.DefaultPageSettings.Margins = New Printing.Margins(50, 50, 50, 50)

            Using prnPreview As New PrintPreviewDialog()
                prnPreview.Document = prnDoc
                prnPreview.WindowState = FormWindowState.Maximized
                prnPreview.Text = "معاينة قبل الطباعة"
                prnPreview.ShowDialog()
            End Using

            RemoveHandler prnDoc.PrintPage, AddressOf Me.PrintPage
        End Using
    End Sub

    ''' <summary>إيجاد العمود بالاسم التقني أولاً ثم بالترويسة (L-03) — تكرار الترويسة لا يضلل الطباعة</summary>
    Private Function ResolveColumn(key As String) As DataGridViewColumn
        If String.IsNullOrWhiteSpace(key) OrElse SourceGrid Is Nothing Then Return Nothing
        If SourceGrid.Columns.Contains(key) Then Return SourceGrid.Columns(key)
        For Each col As DataGridViewColumn In SourceGrid.Columns
            If col.HeaderText = key Then Return col
        Next
        Return Nothing
    End Function

    ''' <summary>ترويسة العرض الفعلية للعمود (بالاسم التقني، أو النص نفسه إن لم يوجد عمود مطابق)</summary>
    Private Function HeaderOf(key As String) As String
        Dim col As DataGridViewColumn = ResolveColumn(key)
        Return If(col IsNot Nothing, col.HeaderText, key)
    End Function

    ''' <summary>قراءة نص خلية مع تنسيق التواريخ والمبالغ (توحيد بدل تكرار الكود)</summary>
    ''' <remarks>L-03: يقبل الاسم التقني أو الترويسة — الحل بالاسم أولاً والتنسيق بحسب ترويسة العرض</remarks>
    Private Function GetCellText(row As DataGridViewRow, colHeader As String) As String
        Dim col As DataGridViewColumn = ResolveColumn(colHeader)
        If col Is Nothing Then Return ""
        Dim val = row.Cells(col.Name).Value
        If val IsNot Nothing AndAlso Not IsDBNull(val) Then
            Dim header As String = col.HeaderText
            Dim s As String = val.ToString()
            If header.Contains("تاريخ") Then
                Dim d As Date
                If Date.TryParse(s, d) Then Return d.ToString("dd/MM/yyyy")
            ElseIf header.Contains("مبلغ") OrElse header.Contains("إجمالي") Then
                Dim dec As Decimal
                If Decimal.TryParse(s, dec) Then Return dec.ToString("0.##")
            End If
            Return s
        End If
        Return ""
    End Function

    Public Sub PrintPage(sender As Object, e As PrintPageEventArgs)
        If RowsToPrint Is Nothing OrElse RowsToPrint.Count = 0 Then
            e.HasMorePages = False
            Return
        End If

        Try
            Dim yPos As Integer = e.MarginBounds.Top
            Const defaultRowHeight As Integer = 25
            Dim leftMargin As Integer = e.MarginBounds.Left
            Dim rightMargin As Integer = e.MarginBounds.Right
            Dim pageWidth As Integer = e.MarginBounds.Width
            Dim pageHeight As Integer = e.MarginBounds.Bottom

            ' 🌟 كل استدعاء PrintPage = صفحة واحدة — رقم الصفحة يزداد دائماً
            _currentPageNumber += 1

            ' 🌟 حساب عرض الأعمدة مرة واحدة فقط (أول صفحة)
            If _colWidths.Count = 0 Then
                _totalWidth = SeqWidth
                _colWidths.Add("ت", SeqWidth)

                For Each colHeader As String In ColumnsToPrint
                    ' 🌟 L-03: القادم من شاشة الاختيار اسم تقني — قرارات العرض تُبنى على ترويسة العرض الفعلية
                    Dim displayHeader As String = HeaderOf(colHeader)
                    Dim w As Integer = 100
                    If displayHeader.Contains("اسم") OrElse displayHeader.Contains("البيان") OrElse displayHeader.Contains("عنوان") OrElse displayHeader.Contains("ملاحظات") Then
                        w = 180
                    ElseIf displayHeader.Contains("رقم") OrElse displayHeader.Contains("جوال") Then
                        w = 95
                    ElseIf displayHeader.Contains("تاريخ") Then
                        w = 90
                    ElseIf displayHeader.Contains("مبلغ") OrElse displayHeader.Contains("إجمالي") Then
                        w = 85
                    End If
                    _colWidths(colHeader) = w
                    _totalWidth += w
                Next

                If _totalWidth > pageWidth Then
                    Dim ratio As Double = pageWidth / _totalWidth
                    For Each key In _colWidths.Keys.ToList()
                        _colWidths(key) = CInt(Math.Max(40, _colWidths(key) * ratio))
                    Next
                    _totalWidth = _colWidths.Values.Sum()
                End If
            End If

            ' 🌟 العنوان وتاريخ الطباعة على الصفحة الأولى فقط
            If _currentPageNumber = 1 Then
                Dim titleSize As SizeF = e.Graphics.MeasureString(ReportTitle, _titleFont)
                e.Graphics.DrawString(ReportTitle, _titleFont, Brushes.DarkBlue, rightMargin - titleSize.Width, yPos)
                yPos += 35

                Dim infoText As String = "تاريخ الطباعة: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm")
                Dim infoSize As SizeF = e.Graphics.MeasureString(infoText, _smallFont)
                e.Graphics.DrawString(infoText, _smallFont, Brushes.Gray, rightMargin - infoSize.Width, yPos)
                yPos += 25
                e.Graphics.DrawLine(Pens.Black, leftMargin, yPos - 5, rightMargin, yPos - 5)
            Else
                yPos += 5
            End If

            ' 🌟 رؤوس الأعمدة تُطبع على كل صفحة
            Dim xPos As Integer = rightMargin - SeqWidth
            e.Graphics.FillRectangle(Brushes.LightGray, xPos, yPos, SeqWidth, defaultRowHeight)
            e.Graphics.DrawRectangle(Pens.Black, xPos, yPos, SeqWidth, defaultRowHeight)
            e.Graphics.DrawString("ت", _headerFont, Brushes.Black, xPos + 12, yPos + 5)

            For Each colHeader As String In ColumnsToPrint
                Dim w As Integer = _colWidths(colHeader)
                xPos -= w
                e.Graphics.FillRectangle(Brushes.LightGray, xPos, yPos, w, defaultRowHeight)
                e.Graphics.DrawRectangle(Pens.Black, xPos, yPos, w, defaultRowHeight)
                ' 🌟 L-03: يُطبع نص الترويسة الظاهر حتى لو كان معرف القائمة اسمه التقني
                Dim columnHeader As String = HeaderOf(colHeader)
                Dim ts As SizeF = e.Graphics.MeasureString(columnHeader, _headerFont)
                e.Graphics.DrawString(columnHeader, _headerFont, Brushes.Black, xPos + w - ts.Width - 5, yPos + 5)
            Next
            yPos += defaultRowHeight

            Dim pageFull As Boolean = False

            Using sf As New StringFormat()
                sf.FormatFlags = StringFormatFlags.DirectionRightToLeft
                sf.Alignment = StringAlignment.Near
                sf.LineAlignment = StringAlignment.Center
                sf.Trimming = StringTrimming.Word

                While _currentPrintRowIndex < RowsToPrint.Count
                    Dim row As DataGridViewRow = RowsToPrint(_currentPrintRowIndex)

                    ' حساب ارتفاع الصف ديناميكياً بناءً على أطول نص
                    Dim currentRowHeight As Integer = defaultRowHeight
                    For Each colHeader As String In ColumnsToPrint
                        Dim w As Integer = _colWidths(colHeader)
                        Dim cellValue As String = GetCellText(row, colHeader)
                        Dim layoutSize As New SizeF(w - 10, 0)
                        Dim ts As SizeF = e.Graphics.MeasureString(cellValue, _cellFont, layoutSize, sf)
                        If ts.Height > currentRowHeight Then
                            currentRowHeight = CInt(ts.Height) + 10
                        End If
                    Next

                    ' 🌟 إذا لم يعد الصف يتسع → نخرج من الحلقة (بدون Return مبكر)
                    ' حتى يُرسم رقم الصفحة ويُضبط HasMorePages في النهاية
                    If yPos + currentRowHeight > pageHeight - 50 Then
                        pageFull = True
                        Exit While
                    End If

                    ' رسم الصف
                    If (_currentPrintRowIndex Mod 2) = 0 Then
                        Using altBrush As New SolidBrush(Color.FromArgb(240, 248, 255))
                            e.Graphics.FillRectangle(altBrush, rightMargin - _totalWidth, yPos, _totalWidth, currentRowHeight)
                        End Using
                    End If

                    Dim xPosRow As Integer = rightMargin - SeqWidth
                    e.Graphics.DrawRectangle(Pens.LightGray, xPosRow, yPos, SeqWidth, currentRowHeight)
                    e.Graphics.DrawString((_currentPrintRowIndex + 1).ToString(), _cellFont, Brushes.Black, New RectangleF(xPosRow + 2, yPos, SeqWidth - 4, currentRowHeight), sf)

                    For Each colHeader As String In ColumnsToPrint
                        Dim w As Integer = _colWidths(colHeader)
                        xPosRow -= w
                        Dim cellValue As String = GetCellText(row, colHeader)
                        e.Graphics.DrawRectangle(Pens.LightGray, xPosRow, yPos, w, currentRowHeight)
                        Dim rect As New RectangleF(xPosRow + 5, yPos, w - 10, currentRowHeight)
                        e.Graphics.DrawString(cellValue, _cellFont, Brushes.Black, rect, sf)
                    Next

                    yPos += currentRowHeight
                    _currentPrintRowIndex += 1
                End While

                ' الإجماليات أسفل التقرير (الصفحة الأخيرة فقط)
                If Not pageFull AndAlso Not String.IsNullOrEmpty(FooterSummary) Then
                    yPos += 10
                    e.Graphics.DrawLine(Pens.Black, leftMargin, yPos - 5, rightMargin, yPos - 5)
                    Dim totalSize As SizeF = e.Graphics.MeasureString(FooterSummary, _totalFont)
                    e.Graphics.DrawString(FooterSummary, _totalFont, Brushes.DarkRed, rightMargin - totalSize.Width, yPos)
                End If

                ' 🌟 رقم الصفحة على كل الصفحات (كان مفقوداً في الصفحات الممتلئة)
                Dim pageNumText As String = $"صفحة {_currentPageNumber}"
                Dim pageNumSize As SizeF = e.Graphics.MeasureString(pageNumText, _smallFont)
                e.Graphics.DrawString(pageNumText, _smallFont, Brushes.Gray, rightMargin - pageNumSize.Width, pageHeight - 20)
            End Using

            e.HasMorePages = pageFull

        Catch ex As Exception
            DatabaseModule.LogError("PrintHelper.PrintPage", ex)
            e.HasMorePages = False
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        _titleFont.Dispose()
        _headerFont.Dispose()
        _cellFont.Dispose()
        _smallFont.Dispose()
        _totalFont.Dispose()
    End Sub

End Class
