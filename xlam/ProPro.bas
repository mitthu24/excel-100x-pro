Attribute VB_Name = "ProPro"
Option Explicit

Public Sub WorkbookAudit()
    Dim ws As Worksheet, a As Worksheet, k As Long, u As Range, f As Long, e As Long, x As Range
    Set a = NewWS("Audit Report")
    ActiveWindow.DisplayGridlines = False
    StyleTitle a, "Workbook Audit", "Size, formulas, errors and hidden sheets"
    a.Range("A4:H4").Value = Array("Sheet", "Visibility", "Used range", "Cells", "Formulas", "Errors", "Tables", "Charts")
    StyleHeader a.Range("A4:H4")
    k = 5
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> a.Name Then
            f = 0: e = 0
            Set u = ws.UsedRange
            On Error Resume Next
            Set x = Nothing
            Set x = u.SpecialCells(xlCellTypeFormulas)
            If Not x Is Nothing Then f = x.CountLarge
            Set x = Nothing
            Set x = u.SpecialCells(xlCellTypeFormulas, xlErrors)
            If Not x Is Nothing Then e = x.CountLarge
            On Error GoTo 0
            a.Cells(k, 1).Value = ws.Name
            a.Cells(k, 2).Value = IIf(ws.Visible = xlSheetVisible, "Visible", "Hidden")
            a.Cells(k, 3).Value = u.Address(False, False)
            a.Cells(k, 4).Value = u.CountLarge
            a.Cells(k, 5).Value = f
            a.Cells(k, 6).Value = e
            a.Cells(k, 7).Value = ws.ListObjects.Count
            a.Cells(k, 8).Value = ws.ChartObjects.Count
            k = k + 1
        End If
    Next
    If k > 5 Then
        With a.Range(a.Cells(5, 6), a.Cells(k - 1, 6)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
    End If
    a.Columns.AutoFit
    Say (k - 5) & " sheet(s) audited"
End Sub

Public Sub GoToSheet()
    Dim s As String, ws As Worksheet, l As String
    For Each ws In ActiveWorkbook.Worksheets
        l = l & ws.Index & ".  " & ws.Name & vbCrLf
    Next
    s = InputBox("Type a sheet name or number:" & vbCrLf & vbCrLf & l, "Go to sheet")
    If s = "" Then Exit Sub
    If IsNumeric(s) Then ActiveWorkbook.Worksheets(CLng(s)).Activate Else ActiveWorkbook.Worksheets(s).Activate
End Sub

Public Sub ZoomAll()
    Dim z As Long, ws As Worksheet, cur As Object
    z = Val(InputBox("Zoom % for every sheet (10-400):", "Zoom all sheets", "100"))
    If z < 10 Or z > 400 Then Exit Sub
    Set cur = ActiveSheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.Zoom = z
        End If
    Next
    cur.Activate
End Sub

Public Sub MaskText()
    Dim c As Range, rng As Range, s As String
    On Error Resume Next
    Set rng = Tgt.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    If MsgBox("Replace the selected values with ****1234 style masks? This changes the data (Ctrl+Z to undo).", vbYesNo + vbQuestion, "Mask text") <> vbYes Then Exit Sub
    For Each c In rng.Cells
        s = CStr(c.Value)
        If Len(s) > 4 Then
            c.NumberFormat = "@"
            c.Value = String(Len(s) - 4, "*") & Right$(s, 4)
        End If
    Next
End Sub

Public Sub ScrubInfo()
    On Error Resume Next
    ActiveWorkbook.RemoveDocumentInformation xlRDIDocumentProperties
    ActiveWorkbook.RemoveDocumentInformation xlRDIRemovePersonalInformation
    On Error GoTo 0
    Say "Personal information removed (save the file to keep it)"
End Sub

Public Sub BreakLinks()
    Dim l As Variant, n As Long
    l = ActiveWorkbook.LinkSources(xlExcelLinks)
    If IsEmpty(l) Then Say "No external links": Exit Sub
    For n = LBound(l) To UBound(l)
        ActiveWorkbook.BreakLink Name:=l(n), Type:=xlLinkTypeExcelLinks
    Next
    Say "External links converted to values"
End Sub

Public Sub BookToValues()
    Dim ws As Worksheet
    If MsgBox("Convert ALL formulas in ALL sheets to values? This cannot be undone with Ctrl+Z.", vbYesNo + vbExclamation, "Formulas to values") <> vbYes Then Exit Sub
    For Each ws In ActiveWorkbook.Worksheets
        If ws.ProtectContents = False Then ws.UsedRange.Value = ws.UsedRange.Value
    Next
    Say "Workbook converted to values"
End Sub

Public Sub DeleteEmptySheets()
    Dim i As Long, n As Long
    If MsgBox("Delete every sheet that has no data?", vbYesNo + vbQuestion, "Delete empty sheets") <> vbYes Then Exit Sub
    Application.DisplayAlerts = False
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
        If ActiveWorkbook.Worksheets.Count > 1 Then
            If Application.WorksheetFunction.CountA(ActiveWorkbook.Worksheets(i).Cells) = 0 And ActiveWorkbook.Worksheets(i).ChartObjects.Count = 0 Then ActiveWorkbook.Worksheets(i).Delete: n = n + 1
        End If
    Next
    Application.DisplayAlerts = True
    Say n & " empty sheet(s) deleted"
End Sub

Public Sub OpenWeb()
    ActiveWorkbook.FollowHyperlink "https://excel-100x-pro.vercel.app/"
End Sub

Public Sub AboutPro()
    MsgBox "Excel 100x Pro  -  version 1.0" & vbCrLf & vbCrLf & _
        "5 tabs, 180+ tools: clean, transform, format, analyse and ready-made business templates." & vbCrLf & _
        "Also available as a web add-in (Excel on the web / Mac):" & vbCrLf & "https://excel-100x-pro.vercel.app/" & vbCrLf & vbCrLf & _
        "Everything runs locally. No data leaves your computer.", vbInformation, "Excel 100x Pro"
End Sub

' Used by the build script to force a full compile of every module.
Public Function SelfTest() As String
    SelfTest = "ok " & AmountInWords(1234567.5, "in")
End Function
