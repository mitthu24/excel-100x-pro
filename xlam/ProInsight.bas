Attribute VB_Name = "ProInsight"
Option Explicit

' ===== HIGHLIGHT =====
Public Sub HlDupes()
    With SelU.FormatConditions.AddUniqueValues
        .DupeUnique = xlDuplicate
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
End Sub

Public Sub HlBlanks()
    With Tgt.FormatConditions.Add(Type:=xlBlanksCondition)
        .Interior.Color = RGB(255, 232, 163)
    End With
End Sub

Public Sub HlErrors()
    With Tgt.FormatConditions.Add(Type:=xlErrorsCondition)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
End Sub

Public Sub Outliers()
    Dim r As Range, c As Range, q1 As Double, q3 As Double, lo As Double, hi As Double, n As Long
    Set r = SelU
    If Application.WorksheetFunction.Count(r) < 4 Then Say "Select at least 4 numbers": Exit Sub
    q1 = Application.WorksheetFunction.Quartile(r, 1)
    q3 = Application.WorksheetFunction.Quartile(r, 3)
    lo = q1 - 1.5 * (q3 - q1)
    hi = q3 + 1.5 * (q3 - q1)
    For Each c In r.Cells
        If IsNumeric(c.Value) And Not IsEmpty(c.Value) And VarType(c.Value) <> vbString Then
            If c.Value < lo Or c.Value > hi Then
                c.Interior.Color = RGB(255, 214, 165)
                c.Font.Color = RGB(185, 28, 28)
                n = n + 1
            End If
        End If
    Next
    Say n & " outlier(s) marked (outside " & Format(lo, "0.##") & " to " & Format(hi, "0.##") & ")"
End Sub

Public Sub TopBottom(ByVal top As Boolean)
    Dim n As Long
    n = Val(InputBox("How many?", IIf(top, "Top N", "Bottom N"), "10"))
    If n < 1 Then Exit Sub
    With SelU.FormatConditions.AddTop10
        .TopBottom = IIf(top, xlTop10Top, xlTop10Bottom)
        .Rank = n
        .Interior.Color = IIf(top, RGB(198, 239, 206), RGB(255, 199, 206))
    End With
End Sub

Public Sub AboveAverage()
    With SelU.FormatConditions.AddAboveAverage
        .AboveBelow = xlAboveAverage
        .Interior.Color = RGB(198, 239, 206)
    End With
End Sub

Public Sub HeatMap()
    With SelU.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
    End With
End Sub

Public Sub DataBars()
    SelU.FormatConditions.AddDatabar.BarColor.Color = RGB(99, 102, 241)
End Sub

Public Sub IconSets()
    With SelU.FormatConditions.AddIconSetCondition
        .IconSet = ActiveWorkbook.IconSets(xl3Arrows)
    End With
End Sub

Public Sub CompareCols()
    Dim s As Range, a As Range, b As Range, da As Object, db As Object, c As Range, na As Long, nb As Long
    Set s = Sel
    If s.Columns.Count < 2 Then Err.Raise vbObjectError + 7, , "Select two columns."
    Set a = Intersect(s.Columns(1), ActiveSheet.UsedRange)
    Set b = Intersect(s.Columns(2), ActiveSheet.UsedRange)
    Set da = CreateObject("Scripting.Dictionary")
    Set db = CreateObject("Scripting.Dictionary")
    da.CompareMode = vbTextCompare
    db.CompareMode = vbTextCompare
    For Each c In a.Cells
        If CStr(c.Value) <> "" Then da(Trim$(CStr(c.Value))) = 1
    Next
    For Each c In b.Cells
        If CStr(c.Value) <> "" Then db(Trim$(CStr(c.Value))) = 1
    Next
    For Each c In a.Cells
        If CStr(c.Value) <> "" Then
            If Not db.Exists(Trim$(CStr(c.Value))) Then c.Interior.Color = RGB(255, 214, 165): na = na + 1
        End If
    Next
    For Each c In b.Cells
        If CStr(c.Value) <> "" Then
            If Not da.Exists(Trim$(CStr(c.Value))) Then c.Interior.Color = RGB(191, 219, 254): nb = nb + 1
        End If
    Next
    Say na & " only in column 1 (orange), " & nb & " only in column 2 (blue)"
End Sub

' ===== ANALYZE =====
Private Function IsNumCol(ByVal r As Range, ByVal j As Long) As Boolean
    Dim body As Range
    If r.Rows.Count < 2 Then Exit Function
    Set body = r.Columns(j).Offset(1).Resize(r.Rows.Count - 1)
    IsNumCol = (Application.WorksheetFunction.Count(body) * 2 > Application.WorksheetFunction.CountA(body)) And Application.WorksheetFunction.Count(body) > 0
End Function

Public Sub StatsSheet()
    Dim r As Range, ws As Worksheet, j As Long, k As Long, body As String, src As String, lbl As Variant, fn As Variant, i As Long
    Set r = Tgt
    src = "'" & r.Worksheet.Name & "'!"
    Set ws = NewWS("Statistics")
    StyleTitle ws, "Statistics", "Descriptive statistics of every numeric column on " & r.Worksheet.Name
    lbl = Array("Column", "Count", "Sum", "Mean", "Median", "Min", "Max", "Std Dev", "Q1", "Q3", "Blanks")
    For i = 0 To UBound(lbl)
        ws.Cells(4, i + 1).Value = lbl(i)
    Next
    StyleHeader ws.Range(ws.Cells(4, 1), ws.Cells(4, 11))
    k = 5
    For j = 1 To r.Columns.Count
        If IsNumCol(r, j) Then
            body = src & r.Columns(j).Offset(1).Resize(r.Rows.Count - 1).Address
            ws.Cells(k, 1).Value = CStr(r.Cells(1, j).Value)
            ws.Cells(k, 2).Formula = "=COUNT(" & body & ")"
            ws.Cells(k, 3).Formula = "=SUM(" & body & ")"
            ws.Cells(k, 4).Formula = "=AVERAGE(" & body & ")"
            ws.Cells(k, 5).Formula = "=MEDIAN(" & body & ")"
            ws.Cells(k, 6).Formula = "=MIN(" & body & ")"
            ws.Cells(k, 7).Formula = "=MAX(" & body & ")"
            ws.Cells(k, 8).Formula = "=STDEV(" & body & ")"
            ws.Cells(k, 9).Formula = "=QUARTILE(" & body & ",1)"
            ws.Cells(k, 10).Formula = "=QUARTILE(" & body & ",3)"
            ws.Cells(k, 11).Formula = "=COUNTBLANK(" & body & ")"
            k = k + 1
        End If
    Next
    If k = 5 Then Say "No numeric columns found" Else ws.Range(ws.Cells(5, 3), ws.Cells(k - 1, 10)).NumberFormat = "#,##0.00"
    ws.Columns.AutoFit
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub FrequencyTable()
    Dim r As Range, d As Object, c As Range, ws As Worksheet, k As Variant, i As Long, n As Long
    Set r = ColSel
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For Each c In r.Cells
        If CStr(c.Value) <> "" Then d(CStr(c.Value)) = d(CStr(c.Value)) + 1: n = n + 1
    Next
    Set ws = NewWS("Frequency")
    StyleTitle ws, "Frequency Table", "How often each value appears"
    ws.Range("A4:C4").Value = Array("Value", "Count", "% of total")
    StyleHeader ws.Range("A4:C4")
    i = 5
    For Each k In d.Keys
        ws.Cells(i, 1).NumberFormat = "@"
        ws.Cells(i, 1).Value = CStr(k)
        ws.Cells(i, 2).Value = d(k)
        ws.Cells(i, 3).Formula = "=B" & i & "/SUM($B$5:$B$" & (4 + d.Count) & ")"
        i = i + 1
    Next
    ws.Range(ws.Cells(5, 3), ws.Cells(i - 1, 3)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(4, 1), ws.Cells(i - 1, 3)).Sort Key1:=ws.Range("B4"), Order1:=xlDescending, Header:=xlYes
    ws.Range(ws.Cells(5, 2), ws.Cells(i - 1, 2)).FormatConditions.AddDatabar
    ws.Columns.AutoFit
    ws.Activate
End Sub

Public Sub ParetoABC()
    Dim s As Range, ws As Worksheet, n As Long
    Set s = SelU
    If s.Columns.Count < 2 Then Err.Raise vbObjectError + 8, , "Select two columns: label and value."
    n = s.Rows.Count
    Set ws = NewWS("Pareto ABC")
    StyleTitle ws, "Pareto / ABC analysis", "A = first 80% of value, B = next 15%, C = last 5%"
    ws.Range("A4:E4").Value = Array("Item", "Value", "Cumulative %", "Class", "")
    ws.Range("A5").Resize(n, 2).Value = s.Resize(n, 2).Value
    ws.Range("A4:B4").Value = Array("Item", "Value")
    ws.Range("A4:D4").Value = Array("Item", "Value", "Cumulative %", "Class")
    StyleHeader ws.Range("A4:D4")
    ws.Range(ws.Cells(4, 1), ws.Cells(4 + n, 2)).Sort Key1:=ws.Range("B4"), Order1:=xlDescending, Header:=xlYes
    ws.Range(ws.Cells(5, 3), ws.Cells(4 + n, 3)).Formula = "=SUM($B$5:B5)/SUM($B$5:$B$" & (4 + n) & ")"
    ws.Range(ws.Cells(5, 3), ws.Cells(4 + n, 3)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(5, 4), ws.Cells(4 + n, 4)).Formula = "=IF(C5<=0.8,""A"",IF(C5<=0.95,""B"",""C""))"
    ws.Columns.AutoFit
    ws.Activate
End Sub

Public Sub CorrMatrix()
    Dim r As Range, ws As Worksheet, idx() As Long, n As Long, i As Long, j As Long, src As String
    Set r = Tgt
    ReDim idx(1 To r.Columns.Count)
    For j = 1 To r.Columns.Count
        If IsNumCol(r, j) Then n = n + 1: idx(n) = j
    Next
    If n < 2 Then Err.Raise vbObjectError + 9, , "Need at least two numeric columns."
    src = "'" & r.Worksheet.Name & "'!"
    Set ws = NewWS("Correlation")
    StyleTitle ws, "Correlation matrix", "+1 = move together, -1 = move opposite, 0 = unrelated"
    For i = 1 To n
        ws.Cells(4, i + 1).Value = CStr(r.Cells(1, idx(i)).Value)
        ws.Cells(i + 4, 1).Value = CStr(r.Cells(1, idx(i)).Value)
        For j = 1 To n
            ws.Cells(i + 4, j + 1).Formula = "=CORREL(" & src & r.Columns(idx(i)).Offset(1).Resize(r.Rows.Count - 1).Address & "," & src & r.Columns(idx(j)).Offset(1).Resize(r.Rows.Count - 1).Address & ")"
        Next
    Next
    StyleHeader ws.Range(ws.Cells(4, 1), ws.Cells(4, n + 1))
    ws.Range(ws.Cells(5, 2), ws.Cells(n + 4, n + 1)).NumberFormat = "0.00"
    With ws.Range(ws.Cells(5, 2), ws.Cells(n + 4, n + 1)).FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 255)
        .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
    End With
    ws.Columns.AutoFit
    ws.Activate
End Sub

Public Sub ForecastCol()
    Dim r As Range, o As Range, nxt As Double, ys As String
    Set r = ColSel
    If Application.WorksheetFunction.Count(r) < 3 Then Err.Raise vbObjectError + 10, , "Need at least 3 numbers in the column."
    ys = r.Address
    Set o = NewColAfter(r, "Trend")
    o.Formula = "=FORECAST(ROW()," & ys & ",ROW(" & ys & "))"
    o.NumberFormat = "#,##0.00"
    o.Font.Color = RGB(79, 70, 229)
    nxt = Application.Evaluate("FORECAST(" & (r.Row + r.Rows.Count) & "," & r.Address(External:=True) & ",ROW(" & r.Address(External:=True) & "))")
    MsgBox "Linear trend added next to the column." & vbCrLf & "Forecast for the next row: " & Format(nxt, "#,##0.00"), vbInformation, "Excel 100x Pro"
End Sub

Public Sub QuickPivot()
    Dim r As Range, ws As Worksheet, pc As PivotCache, pt As PivotTable, j As Long, tc As Long, nc As Long
    Set r = Tgt
    If r.Rows.Count < 2 Then Err.Raise vbObjectError + 11, , "Need a header row and data."
    For j = 1 To r.Columns.Count
        If IsNumCol(r, j) Then
            If nc = 0 Then nc = j
        Else
            If tc = 0 Then tc = j
        End If
    Next
    If tc = 0 Or nc = 0 Then Err.Raise vbObjectError + 11, , "Need one text column and one numeric column."
    Set ws = NewWS("Pivot")
    Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=r)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="Pivot" & Format(Now, "hhnnss"))
    pt.PivotFields(CStr(r.Cells(1, tc).Value)).Orientation = xlRowField
    pt.AddDataField pt.PivotFields(CStr(r.Cells(1, nc).Value)), "Total " & r.Cells(1, nc).Value, xlSum
    ws.Activate
End Sub

Public Sub ShowInsights()
    Dim r As Range, n As Long, m As String, w As WorksheetFunction
    Set w = Application.WorksheetFunction
    Set r = SelU
    n = w.Count(r)
    m = "Selection: " & r.Address(False, False) & vbCrLf & vbCrLf & _
        "Cells with data:  " & w.CountA(r) & vbCrLf & "Blank cells:  " & w.CountBlank(r) & vbCrLf & "Numbers:  " & n
    If n > 0 Then
        m = m & vbCrLf & vbCrLf & "Sum:  " & Format(w.Sum(r), "#,##0.00") & vbCrLf & "Average:  " & Format(w.Average(r), "#,##0.00") & vbCrLf & _
            "Median:  " & Format(w.Median(r), "#,##0.00") & vbCrLf & "Min:  " & Format(w.Min(r), "#,##0.00") & vbCrLf & "Max:  " & Format(w.Max(r), "#,##0.00")
        If n > 1 Then m = m & vbCrLf & "Std dev:  " & Format(w.StDev(r), "#,##0.00")
    End If
    MsgBox m, vbInformation, "Live Insights"
End Sub

' ===== CHARTS =====
Public Sub AddChart(ByVal t As Long, ByVal nm As String)
    Dim r As Range, sh As Shape
    Set r = Tgt
    Set sh = ActiveSheet.Shapes.AddChart2(201, t, r.Columns(r.Columns.Count).Offset(0, 1).Left + 12, r.Top, 420, 260)
    sh.Chart.SetSourceData r
    sh.Chart.HasTitle = True
    sh.Chart.ChartTitle.Text = nm & " overview"
End Sub

Public Sub KpiDashboard()
    Dim src As Worksheet, r As Range, ws As Worksheet, nc As Long, tc As Long, j As Long, col As String, i As Long, m As Long
    Dim lab As Variant, fm As Variant, clr As Variant, nf As Variant, c As Range, ch As Shape
    Set src = ActiveSheet
    Set r = ActiveSheet.UsedRange
    If r.Rows.Count < 3 Then Err.Raise vbObjectError + 12, , "Need a header row and at least 2 data rows."
    For j = 1 To r.Columns.Count
        If IsNumCol(r, j) Then
            If nc = 0 Then nc = j
        ElseIf tc = 0 Then
            tc = j
        End If
    Next
    If nc = 0 Then Err.Raise vbObjectError + 12, , "No numeric column found."
    If tc = 0 Then tc = 1
    col = "'" & src.Name & "'!" & r.Columns(nc).Offset(1).Resize(r.Rows.Count - 1).Address
    Set ws = NewWS("Dashboard")
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1:K30").Interior.Color = RGB(245, 247, 251)
    With ws.Range("B1")
        .Value = "Dashboard - " & CStr(r.Cells(1, nc).Value)
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(49, 46, 129)
    End With
    lab = Array("RECORDS", "TOTAL", "AVERAGE", "HIGHEST")
    fm = Array("=COUNT(" & col & ")", "=SUM(" & col & ")", "=AVERAGE(" & col & ")", "=MAX(" & col & ")")
    clr = Array(RGB(79, 70, 229), RGB(5, 150, 105), RGB(217, 119, 6), RGB(220, 38, 38))
    nf = Array("#,##0", "#,##0", "#,##0.0", "#,##0")
    ws.Columns(1).ColumnWidth = 2
    For i = 0 To 3
        ws.Columns(2 + i * 2).ColumnWidth = 20
        ws.Columns(3 + i * 2).ColumnWidth = 2.5
        With ws.Range(ws.Cells(3, 2 + i * 2), ws.Cells(4, 2 + i * 2))
            .Interior.Color = vbWhite
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .Borders(xlEdgeTop).Weight = xlThick
            .Borders(xlEdgeTop).Color = clr(i)
        End With
        ws.Cells(3, 2 + i * 2).Value = lab(i)
        ws.Cells(3, 2 + i * 2).Font.Size = 9
        ws.Cells(3, 2 + i * 2).Font.Bold = True
        ws.Cells(3, 2 + i * 2).Font.Color = RGB(107, 114, 128)
        ws.Cells(4, 2 + i * 2).Formula = fm(i)
        ws.Cells(4, 2 + i * 2).NumberFormat = nf(i)
        ws.Cells(4, 2 + i * 2).Font.Size = 24
        ws.Cells(4, 2 + i * 2).Font.Bold = True
        ws.Cells(4, 2 + i * 2).Font.Color = clr(i)
    Next
    ws.Rows(4).RowHeight = 40
    m = r.Rows.Count
    If m > 13 Then m = 13
    Set ch = ws.Shapes.AddChart2(201, xlColumnClustered, ws.Range("B7").Left, ws.Range("B7").Top, ws.Range("B7:I7").Width, 280)
    ch.Chart.SetSourceData src.Range(r.Cells(1, nc), r.Cells(m, nc))
    ch.Chart.SeriesCollection(1).XValues = src.Range(r.Cells(2, tc), r.Cells(m, tc))
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = CStr(r.Cells(1, nc).Value) & " by " & CStr(r.Cells(1, tc).Value) & " (first " & (m - 1) & ")"
    ch.Chart.HasLegend = False
    ws.Activate
End Sub

' ===== SELECT =====
Public Sub SelSpecial(ByVal kind As String)
    Dim x As Range
    On Error Resume Next
    Select Case kind
        Case "blank": Set x = Tgt.SpecialCells(xlCellTypeBlanks)
        Case "formula": Set x = Tgt.SpecialCells(xlCellTypeFormulas)
        Case "const": Set x = Tgt.SpecialCells(xlCellTypeConstants)
        Case "error": Set x = Tgt.SpecialCells(xlCellTypeFormulas, xlErrors)
        Case "visible": Set x = Tgt.SpecialCells(xlCellTypeVisible)
    End Select
    On Error GoTo 0
    If x Is Nothing Then Say "Nothing found" Else x.Select
End Sub
