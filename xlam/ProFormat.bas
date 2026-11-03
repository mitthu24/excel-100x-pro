Attribute VB_Name = "ProFormat"
Option Explicit

' ===== FORMAT =====
Public Sub Beautify(ByVal dark As Long, ByVal light As Long)
    Dim r As Range
    Set r = Tgt
    r.FormatConditions.Delete
    r.VerticalAlignment = xlCenter
    With r.Borders
        .LineStyle = xlContinuous
        .Color = RGB(213, 219, 229)
    End With
    With r.Rows(1)
        .Interior.Color = dark
        .Font.Color = vbWhite
        .Font.Bold = True
        .RowHeight = 24
    End With
    With r.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ROW()>" & r.Row & ",MOD(ROW()-" & r.Row & ",2)=0)")
        .Interior.Color = light
    End With
    r.Columns.AutoFit
    Say "Table styled"
End Sub

Public Sub ZebraRows()
    With Tgt.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0")
        .Interior.Color = RGB(241, 243, 248)
    End With
End Sub

Public Sub AutoFitAll()
    Dim r As Range
    Set r = Tgt
    r.Columns.AutoFit
    r.Rows.AutoFit
End Sub

Public Sub ToggleFreeze()
    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
        Say "Panes unfrozen"
    Else
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
        Say "Header row frozen"
    End If
End Sub

Public Sub NumFmt(ByVal f As String)
    SelU.NumberFormat = f
End Sub

Public Sub MakeTable()
    Dim lo As ListObject
    Set lo = ActiveSheet.ListObjects.Add(xlSrcRange, Tgt, , xlYes)
    lo.TableStyle = "TableStyleMedium2"
End Sub

Public Sub ToggleGrid()
    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
End Sub

Public Sub AllBorders()
    With SelU.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 186, 200)
    End With
End Sub

Public Sub WrapText()
    With SelU
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    SelU.Rows.AutoFit
End Sub

Public Sub ClearFmt()
    SelU.ClearFormats
End Sub

' ===== FORMULAS =====
Public Sub ToValues()
    Dim r As Range
    Set r = SelU
    r.Value = r.Value
    Say "Converted to values"
End Sub

Public Sub WrapIfError()
    Dim c As Range, f As Range, n As Long
    On Error Resume Next
    Set f = SelU.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If f Is Nothing Then Say "No formulas found": Exit Sub
    For Each c In f.Cells
        If UCase$(Left$(c.Formula, 9)) <> "=IFERROR(" Then
            c.Formula = "=IFERROR(" & Mid$(c.Formula, 2) & ","""")"
            n = n + 1
        End If
    Next
    Say n & " formula(s) wrapped"
End Sub

Public Sub HighlightFormulas()
    Dim f As Range
    On Error Resume Next
    Set f = Tgt.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If f Is Nothing Then Say "No formulas found": Exit Sub
    f.Interior.Color = RGB(255, 244, 194)
End Sub

Public Sub ConvRef(ByVal absolute As Boolean)
    Dim c As Range, f As Range
    On Error Resume Next
    Set f = SelU.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If f Is Nothing Then Say "No formulas found": Exit Sub
    For Each c In f.Cells
        On Error Resume Next
        c.Formula = Application.ConvertFormula(c.Formula, xlA1, xlA1, IIf(absolute, xlAbsolute, xlRelative), c)
        On Error GoTo 0
    Next
End Sub

Public Sub NamesFromHeaders()
    On Error Resume Next
    Tgt.CreateNames Top:=True
    If Err.Number <> 0 Then
        MsgBox "Some headers are not valid range names (spaces, numbers or symbols). Rename them and try again.", vbInformation, "Excel 100x Pro"
        Exit Sub
    End If
    Say "Named ranges created from the header row"
End Sub

Public Sub InsertF(ByVal f As String)
    On Error Resume Next
    CallByName ActiveCell, "Formula2", VbLet, f
    If Err.Number <> 0 Then
        Err.Clear
        ActiveCell.Formula = f
    End If
End Sub

' ===== SHEETS =====
Public Sub MakeIndex()
    Dim ws As Worksheet, ix As Worksheet, i As Long
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("Index").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ix = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Sheets(1))
    ix.Name = "Index"
    ActiveWindow.DisplayGridlines = False
    With ix.Range("B2")
        .Value = "Workbook Index"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = Brand()
    End With
    i = 4
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "Index" Then
            ix.Hyperlinks.Add Anchor:=ix.Cells(i, 2), Address:="", SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=(i - 3) & ".  " & ws.Name
            i = i + 1
        End If
    Next
    ix.Columns(2).ColumnWidth = 40
End Sub

Public Sub SortSheets()
    Dim i As Long, j As Long, wb As Workbook
    Set wb = ActiveWorkbook
    For i = 1 To wb.Worksheets.Count - 1
        For j = i + 1 To wb.Worksheets.Count
            If UCase$(wb.Worksheets(j).Name) < UCase$(wb.Worksheets(i).Name) Then wb.Worksheets(j).Move Before:=wb.Worksheets(i)
        Next
    Next
End Sub

Public Sub UnhideAll()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next
End Sub

Public Sub RainbowTabs()
    Dim ws As Worksheet, p As Variant, i As Long
    p = Array(RGB(239, 68, 68), RGB(245, 158, 11), RGB(16, 185, 129), RGB(59, 130, 246), RGB(139, 92, 246), RGB(236, 72, 153), RGB(20, 184, 166), RGB(249, 115, 22))
    For Each ws In ActiveWorkbook.Worksheets
        ws.Tab.Color = p(i Mod 8)
        i = i + 1
    Next
End Sub

Public Sub ProtectAll(ByVal doProtect As Boolean)
    Dim ws As Worksheet, pw As String
    pw = InputBox("Password (leave empty for none):", "Sheet protection")
    On Error Resume Next
    For Each ws In ActiveWorkbook.Worksheets
        If doProtect Then ws.Protect pw Else ws.Unprotect pw
    Next
    Say IIf(doProtect, "All sheets protected", "Protection removed where the password matched")
End Sub

Public Sub SheetsFromList()
    Dim c As Range, s As Range, ws As Worksheet, n As Long
    Set s = SelU
    For Each c In s.Cells
        If Trim$(CStr(c.Value)) <> "" Then
            Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
            ws.Name = SafeName(CStr(c.Value))
            n = n + 1
        End If
    Next
    Say n & " sheet(s) created"
End Sub

Public Sub SplitByColumn()
    Dim src As Worksheet, r As Range, k As Long, d As Object, c As Range, key As Variant, ws As Worksheet
    Set src = ActiveSheet
    Set r = src.UsedRange
    k = ActiveCell.Column - r.Column + 1
    If k < 1 Or k > r.Columns.Count Then k = 1
    Set d = CreateObject("Scripting.Dictionary")
    For Each c In r.Columns(k).Offset(1).Resize(r.Rows.Count - 1).Cells
        d(CStr(c.Value)) = 1
    Next
    If d.Count = 0 Then Err.Raise vbObjectError + 5, , "No data rows."
    If d.Count > 100 Then Err.Raise vbObjectError + 5, , d.Count & " unique values - choose a column with fewer than 100."
    For Each key In d.Keys
        r.AutoFilter Field:=k, Criteria1:="=" & key
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        ws.Name = SafeName(IIf(CStr(key) = "", "Blank", CStr(key)))
        r.SpecialCells(xlCellTypeVisible).Copy ws.Range("A1")
        ws.Columns.AutoFit
    Next
    src.AutoFilterMode = False
    Application.CutCopyMode = False
    src.Activate
    Say "Split into " & d.Count & " sheets"
End Sub

Public Sub MergeSheets()
    Dim ws As Worksheet, t As Worksheet, r As Range, nextRow As Long, cols As Long, n As Long
    Set t = NewWS("Merged")
    nextRow = 1
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> t.Name And ws.Name <> "Index" Then
            Set r = ws.UsedRange
            If Application.WorksheetFunction.CountA(r) > 0 Then
                If nextRow = 1 Then
                    r.Copy t.Cells(1, 1)
                    cols = r.Columns.Count
                    t.Cells(1, cols + 1).Value = "Source"
                    If r.Rows.Count > 1 Then t.Range(t.Cells(2, cols + 1), t.Cells(r.Rows.Count, cols + 1)).Value = ws.Name
                    nextRow = r.Rows.Count + 1
                ElseIf r.Rows.Count > 1 Then
                    Set r = r.Offset(1).Resize(r.Rows.Count - 1)
                    r.Copy t.Cells(nextRow, 1)
                    t.Range(t.Cells(nextRow, cols + 1), t.Cells(nextRow + r.Rows.Count - 1, cols + 1)).Value = ws.Name
                    nextRow = nextRow + r.Rows.Count
                End If
                n = n + 1
            End If
        End If
    Next
    StyleHeader t.Range(t.Cells(1, 1), t.Cells(1, cols + 1))
    t.Columns.AutoFit
    Application.CutCopyMode = False
    Say n & " sheet(s) merged"
End Sub

Public Sub CopySheetN()
    Dim n As Long, i As Long
    n = Val(InputBox("How many copies?", "Copy sheet", "1"))
    For i = 1 To n
        ActiveSheet.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Next
End Sub

Public Sub BackupCopy()
    Dim p As String, n As String, e As String, i As Long
    If ActiveWorkbook.Path = "" Then Err.Raise vbObjectError + 6, , "Save the workbook first."
    n = ActiveWorkbook.Name
    i = InStrRev(n, ".")
    e = Mid$(n, i)
    p = ActiveWorkbook.Path & "\" & Left$(n, i - 1) & "_backup_" & Format(Now, "yyyymmdd_hhnn") & e
    ActiveWorkbook.SaveCopyAs p
    Say "Backup saved: " & p
End Sub

' ===== PRINT & EXPORT =====
Public Sub PrintFit()
    On Error GoTo NoPrinter
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    Say "Fit to one page wide, landscape"
    Exit Sub
NoPrinter:
    MsgBox "Excel could not change the page setup. Install or select a default printer (even Microsoft Print to PDF) and try again.", vbInformation, "Excel 100x Pro"
End Sub

Public Sub PrintTitles()
    On Error GoTo NoPrinter
    ActiveSheet.PageSetup.PrintTitleRows = "$1:$1"
    Say "Row 1 repeats on every page"
    Exit Sub
NoPrinter:
    MsgBox "Excel could not change the page setup. Install or select a default printer (even Microsoft Print to PDF) and try again.", vbInformation, "Excel 100x Pro"
End Sub

Public Sub PrintFooter()
    On Error GoTo NoPrinter
    With ActiveSheet.PageSetup
        .LeftFooter = "&F"
        .CenterFooter = "Page &P of &N"
        .RightFooter = "&D"
    End With
    Say "Footer added"
    Exit Sub
NoPrinter:
    MsgBox "Excel could not change the page setup. Install or select a default printer (even Microsoft Print to PDF) and try again.", vbInformation, "Excel 100x Pro"
End Sub

Private Function SaveName(ByVal base As String) As String
    SaveName = Application.GetSaveAsFilename(InitialFileName:=base & ".pdf", FileFilter:="PDF (*.pdf),*.pdf")
End Function

Public Sub ExportPDF(ByVal mode As String)
    Dim fn As Variant, ws As Worksheet, folder As String
    Select Case mode
        Case "sel"
            fn = SaveName("Selection")
            If fn = "False" Then Exit Sub
            Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(fn), OpenAfterPublish:=True
        Case "book"
            fn = SaveName(Replace(ActiveWorkbook.Name, ".", "_"))
            If fn = "False" Then Exit Sub
            ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(fn), OpenAfterPublish:=True
        Case "each"
            folder = IIf(ActiveWorkbook.Path = "", Environ$("USERPROFILE") & "\Documents", ActiveWorkbook.Path) & "\"
            For Each ws In ActiveWorkbook.Worksheets
                If ws.Visible = xlSheetVisible Then ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & ws.Name & ".pdf"
            Next
            Say "PDFs saved in " & folder
    End Select
End Sub
