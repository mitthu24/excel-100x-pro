Attribute VB_Name = "ProExtra"
Option Explicit

Private Function PickFolder(ByVal ttl As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = ttl
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

Private Function OutFolder(ByVal sub1 As String) As String
    Dim p As String
    p = IIf(ActiveWorkbook.Path = "", Environ$("USERPROFILE") & "\Documents", ActiveWorkbook.Path) & "\" & sub1
    If Dir(p, vbDirectory) = "" Then MkDir p
    OutFolder = p
End Function

Private Function NumCol(ByVal r As Range, ByVal j As Long) As Boolean
    Dim body As Range
    If r.Rows.Count < 2 Then Exit Function
    Set body = r.Columns(j).Offset(1).Resize(r.Rows.Count - 1)
    NumCol = (Application.WorksheetFunction.Count(body) > 0) And (Application.WorksheetFunction.Count(body) * 2 > Application.WorksheetFunction.CountA(body))
End Function

' ===== COMPILE FROM FOLDER =====
Public Sub CompileBooks()
    Dim fld As String, f As String, wb As Workbook, host As Workbook, ws As Worksheet, t As Worksheet, r As Range
    Dim nxt As Long, cols As Long, allSheets As Boolean, n As Long, ext As Variant, v As Variant, rc As Long, i As Long
    fld = PickFolder("Select the folder with the workbooks / CSV files to compile")
    If fld = "" Then Exit Sub
    allSheets = (MsgBox("Include ALL sheets of every file?" & vbCrLf & "Yes = every sheet   No = first sheet only", vbYesNo + vbQuestion, "Compile workbooks") = vbYes)
    Set host = ActiveWorkbook
    Set t = NewWS("Compiled")
    nxt = 1
    Application.DisplayAlerts = False
    For Each ext In Array("*.xlsx", "*.xlsm", "*.xls", "*.csv")
        f = Dir(fld & "\" & ext)
        Do While f <> ""
            If Left$(f, 2) <> "~$" And StrComp(f, host.Name, vbTextCompare) <> 0 Then
                Set wb = Workbooks.Open(fld & "\" & f, ReadOnly:=True, UpdateLinks:=0)
                For Each ws In wb.Worksheets
                    If allSheets Or ws.Index = 1 Then
                        Set r = ws.UsedRange
                        If Application.WorksheetFunction.CountA(r) > 0 And r.Cells.CountLarge > 1 Then
                            v = r.Value
                            rc = UBound(v, 1)
                            If r.Cells.CountLarge = 1 Then
                                rc = 0
                            ElseIf nxt = 1 Then
                                cols = UBound(v, 2)
                                t.Cells(1, 1).Resize(rc, cols).Value = v
                                t.Cells(1, cols + 1).Value = "Source File"
                                t.Cells(1, cols + 2).Value = "Sheet"
                                If rc > 1 Then
                                    t.Range(t.Cells(2, cols + 1), t.Cells(rc, cols + 1)).Value = f
                                    t.Range(t.Cells(2, cols + 2), t.Cells(rc, cols + 2)).Value = ws.Name
                                End If
                                nxt = rc + 1
                            ElseIf rc > 1 Then
                                i = IIf(UBound(v, 2) < cols, UBound(v, 2), cols)
                                t.Cells(nxt, 1).Resize(rc - 1, i).Value = r.Offset(1).Resize(rc - 1, i).Value
                                t.Range(t.Cells(nxt, cols + 1), t.Cells(nxt + rc - 2, cols + 1)).Value = f
                                t.Range(t.Cells(nxt, cols + 2), t.Cells(nxt + rc - 2, cols + 2)).Value = ws.Name
                                nxt = nxt + rc - 1
                            End If
                        End If
                    End If
                Next
                wb.Close False
                n = n + 1
            End If
            f = Dir
        Loop
    Next
    Application.DisplayAlerts = True
    host.Activate
    t.Activate
    If nxt > 1 Then
        StyleHeader t.Range(t.Cells(1, 1), t.Cells(1, cols + 2))
        t.Columns.AutoFit
        t.Rows(1).RowHeight = 24
    End If
    Say n & " file(s) compiled into sheet '" & t.Name & "' (" & (nxt - 1) & " rows)"
End Sub

Public Sub FolderToSheets()
    Dim fld As String, f As String, wb As Workbook, host As Workbook, ws As Worksheet, ext As Variant, n As Long, nm As String
    fld = PickFolder("Select the folder - each workbook's first sheet becomes a sheet here")
    If fld = "" Then Exit Sub
    Set host = ActiveWorkbook
    Application.DisplayAlerts = False
    For Each ext In Array("*.xlsx", "*.xlsm", "*.xls", "*.csv")
        f = Dir(fld & "\" & ext)
        Do While f <> ""
            If Left$(f, 2) <> "~$" And StrComp(f, host.Name, vbTextCompare) <> 0 Then
                Set wb = Workbooks.Open(fld & "\" & f, ReadOnly:=True, UpdateLinks:=0)
                nm = Left$(f, InStrRev(f, ".") - 1)
                wb.Worksheets(1).Copy After:=host.Sheets(host.Sheets.Count)
                Set ws = host.Sheets(host.Sheets.Count)
                ws.Name = SafeName(nm)
                ws.UsedRange.Value = ws.UsedRange.Value
                wb.Close False
                n = n + 1
            End If
            f = Dir
        Loop
    Next
    Application.DisplayAlerts = True
    Say n & " workbook(s) imported as sheets"
End Sub

Public Sub ListFiles()
    Dim fld As String, fso As Object, ws As Worksheet, k As Long, deep As Boolean
    fld = PickFolder("Select the folder to list")
    If fld = "" Then Exit Sub
    deep = (MsgBox("Include files in sub-folders too?", vbYesNo + vbQuestion, "List files") = vbYes)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = NewWS("File List")
    StyleTitle ws, "Files in " & fld, "Click a name to open the file"
    ws.Range("A4:F4").Value = Array("Name", "Type", "Size (KB)", "Modified", "Folder", "Path")
    StyleHeader ws.Range("A4:F4")
    k = 5
    ListDir fso, fso.GetFolder(fld), ws, k, deep
    ws.Range("C5:C" & k).NumberFormat = "#,##0.0"
    ws.Range("D5:D" & k).NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Columns("A:F").AutoFit
    If ws.Columns(1).ColumnWidth > 50 Then ws.Columns(1).ColumnWidth = 50
    Say (k - 5) & " file(s) listed"
End Sub

Private Sub ListDir(ByVal fso As Object, ByVal folder As Object, ByVal ws As Worksheet, ByRef k As Long, ByVal deep As Boolean)
    Dim f As Object, sf As Object
    For Each f In folder.Files
        If k > 50000 Then Exit Sub
        ws.Hyperlinks.Add Anchor:=ws.Cells(k, 1), Address:=f.Path, TextToDisplay:=f.Name
        ws.Cells(k, 2).Value = f.Type
        ws.Cells(k, 3).Value = f.Size / 1024
        ws.Cells(k, 4).Value = f.DateLastModified
        ws.Cells(k, 5).Value = folder.Path
        ws.Cells(k, 6).Value = f.Path
        k = k + 1
    Next
    If deep Then
        For Each sf In folder.SubFolders
            ListDir fso, sf, ws, k, deep
        Next
    End If
End Sub

Public Sub SplitBook()
    Dim ws As Worksheet, fld As String, nb As Workbook, host As Workbook, n As Long, conv As Boolean, names As Collection, nm As Variant
    fld = PickFolder("Select the folder to save one workbook per sheet")
    If fld = "" Then Exit Sub
    conv = (MsgBox("Convert formulas to values in the new files?" & vbCrLf & "(recommended - keeps numbers when sheets depend on each other)", vbYesNo + vbQuestion, "Save each sheet as a workbook") = vbYes)
    Set host = ActiveWorkbook
    Set names = New Collection
    For Each ws In host.Worksheets
        If ws.Visible = xlSheetVisible Then names.Add ws.Name
    Next
    Application.DisplayAlerts = False
    For Each nm In names
        host.Worksheets(CStr(nm)).Copy
        Set nb = ActiveWorkbook
        If conv Then nb.Sheets(1).UsedRange.Value = nb.Sheets(1).UsedRange.Value
        nb.SaveAs fld & "\" & Replace(CStr(nm), "/", "-") & ".xlsx", xlOpenXMLWorkbook
        nb.Close False
        n = n + 1
    Next
    Application.DisplayAlerts = True
    host.Activate
    Say n & " workbook(s) saved in " & fld
End Sub

' ===== ADVANCED DATA =====
Public Sub Unpivot()
    Dim r As Range, v As Variant, o() As Variant, i As Long, j As Long, k As Long, ws As Worksheet
    Set r = Tgt
    If r.Columns.Count < 3 Or r.Rows.Count < 2 Then Err.Raise vbObjectError + 30, , "Select a table: first column = ID, other columns = values."
    v = r.Value
    ReDim o(1 To (UBound(v, 1) - 1) * (UBound(v, 2) - 1), 1 To 3)
    For i = 2 To UBound(v, 1)
        For j = 2 To UBound(v, 2)
            If CStr(v(i, j)) <> "" Then
                k = k + 1
                o(k, 1) = v(i, 1): o(k, 2) = v(1, j): o(k, 3) = v(i, j)
            End If
        Next
    Next
    If k = 0 Then Say "Nothing to unpivot": Exit Sub
    Set ws = NewWS("Unpivoted")
    ws.Range("A1:C1").Value = Array(v(1, 1), "Attribute", "Value")
    ws.Range("A2").Resize(UBound(o, 1), 3).Value = o
    If UBound(o, 1) > k Then ws.Range(ws.Cells(k + 2, 1), ws.Cells(UBound(o, 1) + 1, 3)).ClearContents
    StyleHeader ws.Range("A1:C1")
    ws.Columns.AutoFit
    Say k & " rows created"
End Sub

Public Sub GroupSum()
    Dim r As Range, v As Variant, d As Object, ws As Worksheet, i As Long, j As Long, k As Long, n As Long, isN() As Boolean, o() As Variant
    Set r = Tgt
    If r.Rows.Count < 2 Or r.Columns.Count < 2 Then Err.Raise vbObjectError + 31, , "Need a key column and at least one numeric column."
    v = r.Value
    ReDim isN(1 To UBound(v, 2))
    For j = 2 To UBound(v, 2)
        isN(j) = NumCol(r, j)
    Next
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    ReDim o(1 To UBound(v, 1), 1 To UBound(v, 2))
    For j = 1 To UBound(v, 2)
        o(1, j) = v(1, j)
    Next
    n = 1
    For i = 2 To UBound(v, 1)
        If Not d.Exists(CStr(v(i, 1))) Then
            n = n + 1
            d(CStr(v(i, 1))) = n
            o(n, 1) = v(i, 1)
        End If
        k = d(CStr(v(i, 1)))
        For j = 2 To UBound(v, 2)
            If isN(j) Then
                If IsNumeric(v(i, j)) And Not IsEmpty(v(i, j)) Then o(k, j) = o(k, j) + v(i, j)
            End If
        Next
    Next
    Set ws = NewWS("Grouped")
    ws.Range("A1").Resize(n, UBound(v, 2)).Value = o
    StyleHeader ws.Range("A1").Resize(1, UBound(v, 2))
    ws.Columns.AutoFit
    Say (n - 1) & " unique keys; numeric columns summed"
End Sub

Public Sub DropdownList()
    Dim s As String
    s = InputBox("Drop-down items, separated by commas:", "Create drop-down", "Yes,No")
    If s = "" Then Exit Sub
    With Sel.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=s
        .InCellDropdown = True
    End With
    Say "Drop-down created"
End Sub

Public Sub CompareSheets()
    Dim a As Worksheet, b As Worksheet, nm As String, va As Variant, vb As Variant, ra As Long, ca As Long, i As Long, j As Long, n As Long
    Set a = ActiveSheet
    nm = InputBox("Compare the active sheet with which sheet? (name)", "Compare sheets", IIf(a.Index < ActiveWorkbook.Sheets.Count, ActiveWorkbook.Sheets(a.Index + 1).Name, ActiveWorkbook.Sheets(1).Name))
    If nm = "" Then Exit Sub
    Set b = ActiveWorkbook.Worksheets(nm)
    ra = Application.WorksheetFunction.Max(a.UsedRange.Row + a.UsedRange.Rows.Count - 1, b.UsedRange.Row + b.UsedRange.Rows.Count - 1)
    ca = Application.WorksheetFunction.Max(a.UsedRange.Column + a.UsedRange.Columns.Count - 1, b.UsedRange.Column + b.UsedRange.Columns.Count - 1)
    For i = 1 To ra
        For j = 1 To ca
            If CStr(a.Cells(i, j).Value) <> CStr(b.Cells(i, j).Value) Then
                a.Cells(i, j).Interior.Color = RGB(255, 199, 206)
                b.Cells(i, j).Interior.Color = RGB(255, 199, 206)
                n = n + 1
            End If
        Next
    Next
    Say n & " difference(s) highlighted in red on both sheets"
End Sub

Public Sub LockFormulas()
    Dim ws As Worksheet, f As Range, pw As String
    Set ws = ActiveSheet
    pw = InputBox("Password (optional):", "Protect formulas only")
    ws.Unprotect pw
    ws.Cells.Locked = False
    On Error Resume Next
    Set f = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not f Is Nothing Then f.Locked = True
    ws.Protect pw
    Say "Formula cells locked; other cells stay editable"
End Sub

Public Sub DataForm()
    ActiveSheet.ShowDataForm
End Sub

Public Sub MakeCalendar()
    Dim s As String, p() As String, m As Long, y As Long, ws As Worksheet, d As Date, i As Long, c As Long, rw As Long, last As Long
    s = InputBox("Month and year (e.g. 11-2026):", "Calendar", Format(Date, "mm-yyyy"))
    If s = "" Then Exit Sub
    p = Split(Replace(s, "/", "-"), "-")
    m = Val(p(0)): y = Val(p(1))
    If m < 1 Or m > 12 Or y < 1900 Then Err.Raise vbObjectError + 32, , "Use the format mm-yyyy."
    Set ws = NewWS("Cal " & Format(DateSerial(y, m, 1), "mmm yyyy"))
    ActiveWindow.DisplayGridlines = False
    StyleTitle ws, Format(DateSerial(y, m, 1), "mmmm yyyy"), "Type notes in each day"
    ws.Range("A4:G4").Value = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    StyleHeader ws.Range("A4:G4")
    ws.Range("A4:G4").HorizontalAlignment = xlCenter
    d = DateSerial(y, m, 1)
    c = Weekday(d, vbMonday)
    rw = 5
    last = Day(DateSerial(y, m + 1, 0))
    For i = 1 To last
        With ws.Cells(rw, c)
            .Value = i
            .Font.Bold = True
            .Font.Color = Brand()
        End With
        c = c + 1
        If c > 7 Then c = 1: rw = rw + 1
    Next
    With ws.Range("A5:G10")
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(213, 219, 229)
        .RowHeight = 70
    End With
    ws.Range("F5:G10").Interior.Color = RGB(245, 247, 251)
    ws.Columns("A:G").ColumnWidth = 16
    ws.Range("A5").Select
End Sub

' ===== DOCUMENTS =====
Public Sub MakeVoucher(ByVal kind As String)
    Dim ws As Worksheet
    If kind = "receipt" Then
        Set ws = NewWS("Receipt")
        ActiveWindow.DisplayGridlines = False
        ws.Range("A1").Value = "Your Company Name": ws.Range("A2").Value = "PAYMENT RECEIPT"
        ws.Range("A4").Value = "Receipt No": ws.Range("B4").Value = "RC-0001"
        ws.Range("A5").Value = "Date": ws.Range("B5").Formula = "=DATE(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")": ws.Range("B5").NumberFormat = "dd-mmm-yyyy"
        ws.Range("A6").Value = "Received from": ws.Range("B6").Value = "Customer name"
        ws.Range("A7").Value = "Amount": ws.Range("B7").Value = 25000: ws.Range("B7").NumberFormat = INRFormat()
        ws.Range("A8").Value = "In words": ws.Range("B8").Formula = "=AmountInWords(B7)"
        ws.Range("A9").Value = "Payment mode": ws.Range("B9").Value = "UPI"
        ws.Range("B9").Validation.Add Type:=xlValidateList, Formula1:="Cash,UPI,Card,Cheque,Bank transfer"
        ws.Range("A10").Value = "Towards": ws.Range("B10").Value = "Invoice INV-0001"
        ws.Range("A13").Value = "Authorised signatory"
        ws.Range("A4:A10").Font.Bold = True
        ws.Range("B4:B10").HorizontalAlignment = xlLeft
        ws.Range("A12:B12").Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Columns(1).ColumnWidth = 20: ws.Columns(2).ColumnWidth = 48
    Else
        Set ws = NewWS("Delivery Challan")
        ActiveWindow.DisplayGridlines = False
        ws.Range("A1").Value = "Your Company Name": ws.Range("A2").Value = "DELIVERY CHALLAN"
        ws.Range("A4").Value = "Challan No": ws.Range("B4").Value = "DC-0001"
        ws.Range("A5").Value = "Date": ws.Range("B5").Formula = "=DATE(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")": ws.Range("B5").NumberFormat = "dd-mmm-yyyy"
        ws.Range("A6").Value = "Deliver to": ws.Range("B6").Value = "Customer name and address"
        ws.Range("A7").Value = "Vehicle No": ws.Range("B7").Value = "GJ-00-AA-0000"
        ws.Range("A4:A7").Font.Bold = True
        ws.Range("A9:D9").Value = Array("S.No", "Item", "Qty", "Remarks")
        StyleHeader ws.Range("A9:D9")
        ws.Range("A10:A24").Formula = "=IF(B10="""","""",ROW()-9)"
        ws.Range("A9:D24").Borders.LineStyle = xlContinuous
        ws.Range("A9:D24").Borders.Color = RGB(213, 219, 229)
        ws.Range("A27").Value = "Received by (name / sign):"
        ws.Columns(1).ColumnWidth = 12: ws.Columns(2).ColumnWidth = 40: ws.Columns(3).ColumnWidth = 10: ws.Columns(4).ColumnWidth = 30
    End If
    ws.Range("A1").Font.Size = 20: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = Brand()
    ws.Range("A2").Font.Size = 13: ws.Range("A2").Font.Bold = True: ws.Range("A2").Font.Color = RGB(107, 114, 128)
    PageFit ws, False
    ws.Range("A4").Select
End Sub

Public Sub MakeLetter(ByVal kind As String)
    Dim ws As Worksheet, subj As String, body As String, ttl As String
    Select Case kind
        Case "offer"
            ttl = "Offer Letter": subj = "Offer of Employment - [Designation]"
            body = "Dear [Candidate Name]," & vbLf & vbLf & "We are pleased to offer you the position of [Designation] at [Company Name]. Your annual CTC will be Rs. [Amount] and your date of joining will be [Joining Date]." & vbLf & vbLf & "Please sign and return a copy of this letter as your acceptance. We look forward to welcoming you to the team."
        Case "experience"
            ttl = "Experience Letter": subj = "Experience Certificate"
            body = "TO WHOMSOEVER IT MAY CONCERN" & vbLf & vbLf & "This is to certify that [Employee Name] was employed with [Company Name] as [Designation] from [Start Date] to [End Date]." & vbLf & vbLf & "During this period the employee was sincere, hardworking and diligent. We wish them every success in the future."
        Case Else
            ttl = "Certificate": subj = "CERTIFICATE OF APPRECIATION"
            body = "This certificate is proudly presented to" & vbLf & vbLf & "[Name]" & vbLf & vbLf & "in recognition of outstanding performance, dedication and contribution to [Company Name]."
    End Select
    Set ws = NewWS(ttl)
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A:H").ColumnWidth = 12
    ws.Range("A1").Value = "Your Company Name"
    ws.Range("A1").Font.Size = 20: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = Brand()
    ws.Range("A2").Value = "Address line  |  Phone  |  Email"
    ws.Range("A2").Font.Color = RGB(107, 114, 128)
    ws.Range("A3:H3").Borders(xlEdgeBottom).Color = Brand()
    ws.Range("A3:H3").Borders(xlEdgeBottom).Weight = xlMedium
    If kind = "cert" Then
        ws.Range("A6:H6").Merge
        ws.Range("A6").Value = subj
        ws.Range("A6").Font.Size = 28: ws.Range("A6").Font.Bold = True: ws.Range("A6").Font.Color = Brand()
        ws.Range("A6").HorizontalAlignment = xlCenter
        ws.Rows(6).RowHeight = 60
        ws.Range("A9:H16").Merge
        ws.Range("A9").Value = body
        ws.Range("A9").Font.Size = 16
        ws.Range("A9").HorizontalAlignment = xlCenter: ws.Range("A9").VerticalAlignment = xlCenter: ws.Range("A9").WrapText = True
        ws.Range("A19").Value = "Date: ": ws.Range("B19").Formula = "=TODAY()": ws.Range("B19").NumberFormat = "dd-mmm-yyyy"
        ws.Range("F19").Value = "Authorised signatory"
        With ws.Range("A5:H20").Borders
            .LineStyle = xlContinuous: .Color = Brand(): .Weight = xlThick
        End With
    Else
        ws.Range("A5").Value = "Date:"
        ws.Range("B5").Formula = "=TODAY()": ws.Range("B5").NumberFormat = "dd mmmm yyyy": ws.Range("B5").HorizontalAlignment = xlLeft
        ws.Range("A7").Value = "To,": ws.Range("A8").Value = "[Name]": ws.Range("A9").Value = "[Address]"
        ws.Range("A11").Value = "Subject: " & subj
        ws.Range("A11").Font.Bold = True
        ws.Range("A13:H22").Merge
        ws.Range("A13").Value = body
        ws.Range("A13").WrapText = True: ws.Range("A13").VerticalAlignment = xlTop: ws.Range("A13").HorizontalAlignment = xlLeft
        ws.Range("A24").Value = "Yours sincerely,"
        ws.Range("A28").Value = "[Name]": ws.Range("A29").Value = "[Designation], [Company Name]"
        ws.Range("A28").Font.Bold = True
    End If
    PageFit ws, (kind = "cert")
    ws.Range("A1").Select
End Sub

' ===== PRINT / PDF / EMAIL =====
Public Sub ExportPDFActive()
    Dim p As String, fld As String
    fld = OutFolder("PDF")
    p = fld & "\" & ActiveSheet.Name & "_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=p, Quality:=xlQualityStandard, OpenAfterPublish:=True
    Say "PDF saved: " & p
End Sub

Public Sub PayslipsBatch()
    Dim src As Worksheet, ws As Worksheet, i As Long, n As Long, fld As String, nm As String, mth As String
    Set src = ActiveSheet
    If CStr(src.Range("A5").Value) <> "Employee" Or CStr(src.Range("H5").Value) <> "Net Pay" Then Err.Raise vbObjectError + 33, , "Open the Payroll Sheet first (Business > HR > Payroll Sheet), fill employees, then run this."
    fld = OutFolder("Payslips")
    mth = InputBox("Month label for the payslips:", "Payslips", Format(Date, "mmmm yyyy"))
    If mth = "" Then Exit Sub
    Application.DisplayAlerts = False
    For i = 6 To src.Cells(src.Rows.Count, 1).End(xlUp).Row
        nm = CStr(src.Cells(i, 1).Value)
        If nm <> "" Then
            MakePayslip
            Set ws = ActiveSheet
            ws.Range("B4").Value = nm
            ws.Range("B6").Value = mth
            ws.Range("B9").Value = src.Cells(i, 2).Value
            ws.Range("B10").Value = src.Cells(i, 3).Value
            ws.Range("B11").Value = src.Cells(i, 4).Value
            ws.Range("E9").Value = src.Cells(i, 6).Value
            ws.Range("E10").Value = src.Cells(i, 7).Value
            ws.Range("E11").Value = 0
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fld & "\Payslip_" & Replace(nm, " ", "_") & "_" & Replace(mth, " ", "_") & ".pdf"
            ws.Delete
            n = n + 1
        End If
    Next
    Application.DisplayAlerts = True
    src.Activate
    Say n & " payslip PDF(s) saved in " & fld
End Sub

Public Sub EmailPDF()
    Dim p As String, ol As Object, m As Object
    p = Environ$("TEMP") & "\" & ActiveSheet.Name & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=p, OpenAfterPublish:=False
    On Error GoTo NoMail
    Set ol = CreateObject("Outlook.Application")
    Set m = ol.CreateItem(0)
    m.Subject = ActiveSheet.Name
    m.Body = "Please find the attached " & ActiveSheet.Name & "."
    m.Attachments.Add p
    m.Display
    Exit Sub
NoMail:
    MsgBox "Outlook is not available. The PDF was saved here:" & vbCrLf & p, vbInformation, "Excel 100x Pro"
End Sub

Public Sub PrintAreaSel()
    If Selection.Cells.CountLarge = 1 Then
        ActiveSheet.PageSetup.PrintArea = ""
        Say "Print area cleared"
    Else
        ActiveSheet.PageSetup.PrintArea = Selection.Address
        Say "Print area set to the selection"
    End If
End Sub

Public Sub A4Setup()
    On Error GoTo NoPrinter
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA4
        .LeftMargin = Application.CentimetersToPoints(1.2)
        .RightMargin = Application.CentimetersToPoints(1.2)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
    End With
    Say "A4 page setup applied"
    Exit Sub
NoPrinter:
    MsgBox "Excel could not change the page setup. Install or select a default printer (even 'Microsoft Print to PDF') and try again.", vbInformation, "Excel 100x Pro"
End Sub
