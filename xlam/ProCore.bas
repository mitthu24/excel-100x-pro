Attribute VB_Name = "ProCore"
Option Explicit

' ===== Shared helpers =====
Public gHdr As Boolean   ' True when the last ColSel() skipped a header row

Public Function Brand() As Long
    Brand = RGB(79, 70, 229)
End Function

Public Function INRFormat() As String
    INRFormat = "[$" & ChrW(8377) & "-4009]#,##0.00"
End Function

Public Function Tgt() As Range
    Dim r As Range
    If TypeName(Selection) <> "Range" Then Err.Raise vbObjectError + 1, , "Select some cells first."
    If Selection.Cells.CountLarge = 1 Then Set r = ActiveSheet.UsedRange Else Set r = Intersect(Selection, ActiveSheet.UsedRange)
    If r Is Nothing Then Err.Raise vbObjectError + 2, , "No data in the selection."
    Set Tgt = r
End Function

Public Function Sel() As Range
    If TypeName(Selection) <> "Range" Then Err.Raise vbObjectError + 1, , "Select some cells first."
    Set Sel = Selection
End Function

Public Function SelU() As Range
    Dim r As Range
    Set r = Intersect(Sel, ActiveSheet.UsedRange)
    If r Is Nothing Then Err.Raise vbObjectError + 2, , "No data in the selection."
    Set SelU = r
End Function

' First selected column, limited to data. One cell selected = that whole column, header row skipped.
Public Function ColSel() As Range
    Dim r As Range, s As Range
    Set s = Sel
    gHdr = False
    If s.Cells.CountLarge = 1 Then
        Set r = Intersect(s.EntireColumn, ActiveSheet.UsedRange)
        If Not r Is Nothing Then
            If r.Rows.Count > 1 Then
                Set r = r.Offset(1).Resize(r.Rows.Count - 1)
                gHdr = True
            End If
        End If
    Else
        Set r = Intersect(s.Columns(1).EntireColumn, s, ActiveSheet.UsedRange)
    End If
    If r Is Nothing Then Err.Raise vbObjectError + 2, , "No data in the selection."
    Set ColSel = r
End Function

' Insert an empty column right of r and return the matching cells in it (header cell gets lbl).
Public Function NewColAfter(ByVal r As Range, ByVal lbl As String) As Range
    Dim c As Range
    Set c = r.Columns(r.Columns.Count).Offset(0, 1)
    c.EntireColumn.Insert
    Set c = r.Columns(r.Columns.Count).Offset(0, 1)
    If gHdr Then c.Cells(1, 1).Offset(-1, 0).Value = lbl
    Set NewColAfter = c
End Function

Public Sub PageFit(ByVal ws As Worksheet, ByVal landscape As Boolean)
    On Error Resume Next
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = IIf(landscape, xlLandscape, xlPortrait)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub

Public Sub FreezeRows(ByVal n As Long)
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = n
    ActiveWindow.FreezePanes = True
End Sub

Public Sub Say(ByVal msg As String)
    Application.StatusBar = "Excel 100x Pro: " & msg
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 8), "'" & ThisWorkbook.Name & "'!ProCore.ClearBar"
End Sub

Public Sub ClearBar()
    Application.StatusBar = False
End Sub

Public Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Object
    For Each ws In ActiveWorkbook.Sheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next
End Function

Public Function SafeName(ByVal s As String) As String
    Dim i As Long, c As String, o As String, b As String, n As Long
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr("\/?*[]:", c) = 0 Then o = o & c
    Next
    o = Trim$(Left$(o, 31))
    If o = "" Then o = "Sheet"
    b = o
    n = 1
    Do While SheetExists(o)
        n = n + 1
        o = Left$(b, 31 - Len(CStr(n)) - 3) & " (" & n & ")"
    Loop
    SafeName = o
End Function

Public Function NewWS(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    ws.Name = SafeName(nm)
    Set NewWS = ws
End Function

Public Sub StyleHeader(ByVal r As Range, Optional ByVal clr As Long = -1)
    If clr = -1 Then clr = Brand()
    r.Interior.Color = clr
    r.Font.Color = vbWhite
    r.Font.Bold = True
    r.VerticalAlignment = xlCenter
End Sub

Public Sub StyleTitle(ByVal ws As Worksheet, ByVal t As String, ByVal s As String)
    With ws.Range("A1")
        .Value = t
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = Brand()
    End With
    With ws.Range("A2")
        .Value = s
        .Font.Italic = True
        .Font.Color = RGB(107, 114, 128)
    End With
End Sub

' ===== Number to words =====
Private Function Three(ByVal n As Long) As String
    Dim o As Variant, t As Variant, r As String
    o = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    t = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    If n >= 100 Then
        r = o(n \ 100) & " Hundred "
        n = n Mod 100
    End If
    If n >= 20 Then
        r = r & t(n \ 10) & " "
        n = n Mod 10
    End If
    If n > 0 Then r = r & o(n)
    Three = Trim$(r)
End Function

Private Function WordsOf(ByVal n As Double, ByVal sys As String) As String
    Dim u As Variant, nm As Variant, i As Long, q As Double, r As String
    If n = 0 Then WordsOf = "Zero": Exit Function
    If sys = "in" Then
        u = Array(10000000#, 100000#, 1000#)
        nm = Array("Crore", "Lakh", "Thousand")
    Else
        u = Array(1000000000000#, 1000000000#, 1000000#, 1000#)
        nm = Array("Trillion", "Billion", "Million", "Thousand")
    End If
    For i = LBound(u) To UBound(u)
        q = Int(n / u(i))
        If q > 0 Then
            r = r & WordsOf(q, sys) & " " & nm(i) & " "
            n = n - q * u(i)
        End If
    Next
    If n > 0 Then r = r & Three(CLng(n))
    WordsOf = Trim$(r)
End Function

' Usable as a worksheet function too:  =AmountInWords(A1)  or  =AmountInWords(A1,"intl")
Public Function AmountInWords(ByVal x As Double, Optional ByVal sys As String = "in") As String
    Dim cents As Double, whole As Double, frac As Long, big As String, small As String, tail As String
    cents = Int(Abs(x) * 100 + 0.5)
    whole = Int(cents / 100)
    frac = CLng(cents - whole * 100)
    If sys = "in" Then
        big = "Rupees": small = "Paise": tail = " Only"
    Else
        big = "Dollars": small = "Cents": tail = ""
    End If
    AmountInWords = IIf(x < 0, "Minus ", "") & big & " " & WordsOf(whole, sys) & IIf(frac > 0, " and " & WordsOf(CDbl(frac), sys) & " " & small, "") & tail
End Function

Public Function MakeGuid() As String
    Dim i As Long, s As String
    For i = 1 To 32
        s = s & Hex(Int(Rnd * 16))
    Next
    MakeGuid = LCase$(Left$(s, 8) & "-" & Mid$(s, 9, 4) & "-4" & Mid$(s, 14, 3) & "-a" & Mid$(s, 18, 3) & "-" & Mid$(s, 21, 12))
End Function

Private Function ToNum(ByVal s As String, ByRef d As Double) As Boolean
    Dim t As String, pct As Boolean
    t = Replace(Replace(Replace(Replace(Replace(Trim$(s), ",", ""), "$", ""), ChrW(8377), ""), " ", ""), ChrW(163), "")
    If Len(t) = 0 Then Exit Function
    If Right$(t, 1) = "%" Then pct = True: t = Left$(t, Len(t) - 1)
    If Not t Like "*[0-9]*" Then Exit Function
    If t Like "*[!0-9.+-]*" Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    d = CDbl(t)
    If pct Then d = d / 100
    ToNum = True
End Function

' ===== CLEAN =====
Public Sub MapText(ByVal mode As String)
    Dim rng As Range, c As Range, s As String, o As String, d As Double, n As Long, i As Long, p As String
    On Error Resume Next
    Set rng = Tgt.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If rng Is Nothing Then Say "No text cells found": Exit Sub
    For Each c In rng.Cells
        o = CStr(c.Value)
        s = o
        If mode = "trim" Or mode = "magic" Then s = Application.WorksheetFunction.Trim(Replace(Replace(s, Chr(160), " "), ChrW(8203), ""))
        If mode = "junk" Or mode = "magic" Then s = Application.WorksheetFunction.Clean(s)
        If mode = "upper" Then s = UCase$(s)
        If mode = "lower" Then s = LCase$(s)
        If mode = "proper" Then s = StrConv(s, vbProperCase)
        If mode = "sentence" Then s = UCase$(Left$(LCase$(s), 1)) & Mid$(LCase$(s), 2)
        If (mode = "textnum" Or mode = "magic") And ToNum(s, d) Then
            If c.NumberFormat = "@" Then c.NumberFormat = "General"
            c.Value = d
            n = n + 1
        ElseIf s <> o Then
            c.Value = s
            n = n + 1
        End If
    Next
    Say n & " cell(s) updated"
End Sub

Public Sub FillBlanksDown()
    Dim r As Range, b As Range
    Set r = Tgt
    If r.Row = 1 Then Set r = r.Offset(1).Resize(r.Rows.Count - 1)
    On Error Resume Next
    Set b = r.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    If b Is Nothing Then Say "No blank cells": Exit Sub
    b.FormulaR1C1 = "=R[-1]C"
    b.Value = b.Value
    Say "Blanks filled"
End Sub

Public Sub UnmergeFill()
    Dim c As Range, a As Range, v As Variant
    For Each c In Tgt.Cells
        If c.MergeCells Then
            Set a = c.MergeArea
            v = a.Cells(1, 1).Value
            a.UnMerge
            a.Value = v
        End If
    Next
    Say "Unmerged and filled"
End Sub

Public Sub RemoveDupes()
    Dim r As Range, v As Variant, d As Object, i As Long, j As Long, k As String, n As Long, dup() As Boolean
    Set r = Tgt
    v = r.Value
    If r.Cells.CountLarge = 1 Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    ReDim dup(1 To UBound(v, 1))
    For i = 2 To UBound(v, 1)
        k = ""
        For j = 1 To UBound(v, 2)
            k = k & Chr(1) & CStr(v(i, j))
        Next
        If d.Exists(k) Then dup(i) = True Else d(k) = 1
    Next
    For i = UBound(v, 1) To 2 Step -1
        If dup(i) Then r.Rows(i).Delete Shift:=xlUp: n = n + 1
    Next
    Say n & " duplicate row(s) removed (first row treated as header)"
End Sub

Public Sub DeleteBlankRows()
    Dim r As Range, i As Long, n As Long
    Set r = Tgt
    For i = r.Rows.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(r.Rows(i)) = 0 Then r.Rows(i).EntireRow.Delete: n = n + 1
    Next
    Say n & " blank row(s) deleted"
End Sub

Public Sub DeleteBlankCols()
    Dim r As Range, i As Long, n As Long
    Set r = Tgt
    For i = r.Columns.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(r.Columns(i)) = 0 Then r.Columns(i).EntireColumn.Delete: n = n + 1
    Next
    Say n & " blank column(s) deleted"
End Sub

Public Sub RemoveLinks()
    Tgt.Hyperlinks.Delete
    Say "Hyperlinks removed"
End Sub

' ===== TRANSFORM =====
Public Sub SplitColumn()
    Dim d As String, r As Range, c As Range, mx As Long, k As Long, isTab As Boolean, isSp As Boolean
    d = InputBox("Delimiter (type a character, or the word SPACE or TAB):", "Split column", ",")
    If d = "" Then Exit Sub
    If UCase$(d) = "SPACE" Then d = " "
    If UCase$(d) = "TAB" Then d = vbTab
    Set r = ColSel
    For Each c In r.Cells
        k = UBound(Split(CStr(c.Value), d)) + 1
        If k > mx Then mx = k
    Next
    If mx < 2 Then Say "Delimiter not found": Exit Sub
    isTab = (d = vbTab)
    isSp = (d = " ")
    r.Offset(0, 1).Resize(1, mx - 1).EntireColumn.Insert
    r.TextToColumns Destination:=r.Cells(1, 1), DataType:=xlDelimited, TextQualifier:=xlTextQualifierNone, ConsecutiveDelimiter:=False, Tab:=isTab, Space:=isSp, Other:=Not (isTab Or isSp), OtherChar:=IIf(isTab Or isSp, "", Left$(d, 1))
    Say "Split into " & mx & " columns"
End Sub

Public Sub MergeColumns()
    Dim s As Range, d As String, i As Long, j As Long, t As String, o As Range
    Set s = SelU
    gHdr = False
    If s.Columns.Count < 2 Then Err.Raise vbObjectError + 3, , "Select two or more columns."
    d = InputBox("Separator between values:", "Merge columns", " ")
    Set o = NewColAfter(s, "Merged")
    For i = 1 To s.Rows.Count
        t = ""
        For j = 1 To s.Columns.Count
            If CStr(s.Cells(i, j).Text) <> "" Then t = t & IIf(t = "", "", d) & s.Cells(i, j).Text
        Next
        o.Cells(i, 1).Value = t
    Next
    Say "Merged into a new column"
End Sub

Public Sub TransposeSel()
    Dim r As Range, ws As Worksheet
    Set r = Tgt
    r.Copy
    Set ws = NewWS("Transposed")
    ws.Range("A1").PasteSpecial Paste:=xlPasteAll, Transpose:=True
    Application.CutCopyMode = False
    ws.Columns.AutoFit
    Say "Transposed into a new sheet"
End Sub

Public Sub SortCol(ByVal asc As Boolean)
    Dim r As Range, whole As Boolean, k As Long
    whole = (TypeName(Selection) = "Range" And Selection.Cells.CountLarge = 1)
    Set r = Tgt
    k = 1
    If whole Then
        k = ActiveCell.Column - r.Column + 1
        If k < 1 Or k > r.Columns.Count Then k = 1
    End If
    r.Sort Key1:=r.Cells(1, k), Order1:=IIf(asc, xlAscending, xlDescending), Header:=IIf(whole, xlYes, xlNo), Orientation:=xlTopToBottom
    Say "Sorted"
End Sub

Public Sub ShuffleRows()
    Dim r As Range, h As Range, whole As Boolean
    whole = (Selection.Cells.CountLarge = 1)
    Set r = Tgt
    If whole Then Set r = r.Offset(1).Resize(r.Rows.Count - 1)
    Set h = r.Columns(r.Columns.Count).Offset(0, 1)
    If Application.WorksheetFunction.CountA(h) > 0 Then Err.Raise vbObjectError + 4, , "The column right of the data must be empty."
    h.Formula = "=RAND()"
    h.Value = h.Value
    r.Resize(, r.Columns.Count + 1).Sort Key1:=h.Cells(1, 1), Order1:=xlAscending, Header:=xlNo
    h.ClearContents
    Say "Rows shuffled"
End Sub

Public Sub FillSerial()
    Dim s As Range, i As Long
    Set s = Sel
    For i = 1 To s.Rows.Count
        s.Cells(i, 1).Value = i
    Next
    Say "Serial numbers added"
End Sub

Public Sub FillGuid()
    Dim c As Range
    Randomize
    For Each c In Sel.Cells
        c.Value = MakeGuid()
    Next
    Say "IDs generated"
End Sub

Public Sub StampDate(ByVal withTime As Boolean)
    If withTime Then
        ActiveCell.Value = Now
        ActiveCell.NumberFormat = "dd-mmm-yyyy hh:mm"
    Else
        ActiveCell.Value = Date
        ActiveCell.NumberFormat = "dd-mmm-yyyy"
    End If
End Sub

Public Sub AmountWords(ByVal sys As String)
    Dim s As Range, o As Range, i As Long, v As Variant
    Set s = SelU
    gHdr = False
    Set o = NewColAfter(s, "")
    For i = 1 To s.Rows.Count
        v = s.Cells(i, s.Columns.Count).Value
        If Not IsEmpty(v) And Not IsError(v) Then
            If IsNumeric(v) Then o.Cells(i, 1).Value = AmountInWords(CDbl(v), sys)
        End If
    Next
    o.EntireColumn.AutoFit
    Say "Words added in the next column"
End Sub

' ===== TEXT TOOLS =====
Public Sub ExtractText(ByVal kind As String)
    Dim r As Range, o As Range, re As Object, i As Long, m As Object, t As String, j As Long
    Set r = ColSel
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    Select Case kind
        Case "email": re.Pattern = "[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"
        Case "digits": re.Pattern = "\D"
        Case "nodigits": re.Pattern = "\d"
        Case "phone": re.Pattern = "(?:\+?91[\- ]?)?[6-9]\d{9}"
    End Select
    Set o = NewColAfter(r, "Extracted " & kind)
    For i = 1 To r.Rows.Count
        t = CStr(r.Cells(i, 1).Value)
        If kind = "digits" Or kind = "nodigits" Then
            o.Cells(i, 1).Value = re.Replace(t, "")
        Else
            Set m = re.Execute(t)
            t = ""
            For j = 0 To m.Count - 1
                t = t & IIf(t = "", "", ", ") & m(j).Value
            Next
            o.Cells(i, 1).Value = t
        End If
    Next
    Say "Result placed in the next column"
End Sub

Public Sub AddAffix(ByVal isPrefix As Boolean)
    Dim txt As String, c As Range, rng As Range
    txt = InputBox("Text to add:", IIf(isPrefix, "Add prefix", "Add suffix"))
    If txt = "" Then Exit Sub
    On Error Resume Next
    Set rng = Tgt.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        c.NumberFormat = "@"
        c.Value = IIf(isPrefix, txt & c.Value, c.Value & txt)
    Next
    Say "Done"
End Sub

Public Sub PadZeros()
    Dim n As Long, c As Range, rng As Range, s As String
    n = Val(InputBox("Total length (e.g. 6 makes 123 into 000123):", "Pad with zeros", "6"))
    If n < 1 Then Exit Sub
    On Error Resume Next
    Set rng = Tgt.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        s = CStr(c.Value)
        c.NumberFormat = "@"
        If Len(s) < n Then c.Value = String(n - Len(s), "0") & s
    Next
    Say "Padded to " & n & " characters"
End Sub

Public Sub ReverseText()
    Dim c As Range, rng As Range
    On Error Resume Next
    Set rng = Tgt.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        c.Value = StrReverse(CStr(c.Value))
    Next
End Sub

Public Sub FindReplaceSel()
    Dim f As String, r As String
    f = InputBox("Find what:", "Find and replace (selection only)")
    If f = "" Then Exit Sub
    r = InputBox("Replace with (empty = delete):", "Find and replace")
    Tgt.Replace What:=f, Replacement:=r, LookAt:=xlPart, MatchCase:=False
    Say "Replaced"
End Sub

Public Sub RoundValues()
    Dim n As Long, c As Range, rng As Range
    n = Val(InputBox("Decimal places:", "Round values", "2"))
    On Error Resume Next
    Set rng = Tgt.SpecialCells(xlCellTypeConstants, xlNumbers)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        c.Value = Round(c.Value, n)
    Next
    Say "Rounded to " & n & " decimals"
End Sub

Public Sub TextToDates()
    Dim c As Range, rng As Range, k As Long
    On Error Resume Next
    Set rng = Tgt.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If IsDate(c.Value) Then
            c.NumberFormat = "dd-mmm-yyyy"
            c.Value = CDate(c.Value)
            k = k + 1
        End If
    Next
    Say k & " text date(s) converted"
End Sub

' ===== DATES & CALC (new column to the right) =====
Public Sub DateCol(ByVal kind As String)
    Dim r As Range, o As Range, a As String, f As String
    Set r = ColSel
    a = r.Cells(1, 1).Address(False, False)
    Select Case kind
        Case "dddd", "mmmm": f = "TEXT(" & a & ",""" & kind & """)"
        Case "q": f = """Q""&ROUNDUP(MONTH(" & a & ")/3,0)"
        Case "fy": f = """FY ""&IF(MONTH(" & a & ")>=4,YEAR(" & a & "),YEAR(" & a & ")-1)&""-""&RIGHT(IF(MONTH(" & a & ")>=4,YEAR(" & a & ")+1,YEAR(" & a & ")),2)"
        Case "age": f = "DATEDIF(" & a & ",TODAY(),""Y"")"
    End Select
    Set o = NewColAfter(r, "Result")
    o.Formula = "=IF(" & a & "="""",""""," & f & ")"
    o.EntireColumn.AutoFit
End Sub

Public Sub CalcCol(ByVal kind As String)
    Dim r As Range, o As Range, a As String, f As String, nf As String
    Set r = ColSel
    If kind = "chg" And r.Row = 1 Then Set r = r.Offset(1).Resize(r.Rows.Count - 1)
    a = r.Cells(1, 1).Address(False, False)
    Select Case kind
        Case "pct": f = a & "/SUM(" & r.Address & ")": nf = "0.0%"
        Case "run": f = "SUM(" & r.Cells(1, 1).Address(True, True) & ":" & a & ")": nf = "General"
        Case "rank": f = "RANK(" & a & "," & r.Address & ")": nf = "0"
        Case "chg": f = "IFERROR(" & a & "/" & r.Cells(1, 1).Offset(-1, 0).Address(False, False) & "-1,"""")": nf = "0.0%"
    End Select
    Set o = NewColAfter(r, UCase$(Left$(kind, 1)) & Mid$(kind, 2))
    o.Formula = "=IF(ISNUMBER(" & a & ")," & f & ","""")"
    o.NumberFormat = nf
    o.EntireColumn.AutoFit
End Sub
