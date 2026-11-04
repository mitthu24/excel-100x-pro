Attribute VB_Name = "ProLicence"
Option Explicit

' Licensing: yearly subscription, N computers per key, seat transfer when a PC dies.
' Seats are counted on the server, so a released computer stops working at its next check.

Private Const API_URL   As String = "https://excel-100x-pro.vercel.app/api/licence"
Private Const BUY_URL   As String = "https://excel-100x-pro.vercel.app/buy.html"
Private Const APP       As String = "Excel100xPro"
Private Const SECTION   As String = "Licence"
Private Const GRACE_DAYS As Long = 14      ' days the add-in keeps working with no internet
Private Const PRODUCT   As String = "Excel 100x Pro"

Private mChecked As Boolean      ' re-checked at most once per Excel session
Private mValid   As Boolean

' ---------------------------------------------------------------- stored state
Private Function Cfg(ByVal k As String) As String
    Cfg = GetSetting(APP, SECTION, k, "")
End Function

Private Sub SetCfg(ByVal k As String, ByVal v As String)
    SaveSetting APP, SECTION, k, v
End Sub

Private Sub ForgetLicence()
    On Error Resume Next
    DeleteSetting APP, SECTION
    mChecked = False
    mValid = False
End Sub

' ---------------------------------------------------------------- machine identity
' Hardware UUID where available, else the system drive serial. Hashed so the raw
' hardware id never leaves the PC.
Public Function MachineId() As String
    Static cached As String
    Dim raw As String, o As Object, it As Object
    If cached <> "" Then MachineId = cached: Exit Function
    On Error Resume Next
    Set o = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select UUID from Win32_ComputerSystemProduct")
    For Each it In o
        raw = CStr(it.UUID)
        Exit For
    Next
    If raw = "" Or InStr(raw, "FFFFFFFF") > 0 Then
        Set o = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select VolumeSerialNumber from Win32_LogicalDisk where DeviceID='C:'")
        For Each it In o
            raw = CStr(it.VolumeSerialNumber)
            Exit For
        Next
    End If
    On Error GoTo 0
    raw = raw & "|" & Environ$("COMPUTERNAME")
    cached = Fnv(raw)
    MachineId = cached
End Function

' FNV-1a style hash in two 32-bit halves. Doubles hold the values because VBA has no
' unsigned 32-bit type, and Hex8 formats them because Hex$ overflows above 2^31.
Private Function Fnv(ByVal s As String) As String
    Dim h1 As Double, h2 As Double, i As Long, c As Long
    h1 = 2166136261#: h2 = 16777619#
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        h1 = Mul32(XorLow(h1, c And &HFF&), 16777619#)
        h2 = Mul32(h2 + c * (i + 7), 16777619#)
    Next
    Fnv = Hex8(h1) & Hex8(h2)
End Function

' v Xor b for a byte b. Only the low byte changes, so the Xor runs on two small Longs:
' VBA coerces Xor operands to Long and v goes above 2^31, which would overflow.
Private Function XorLow(ByVal v As Double, ByVal b As Long) As Double
    Dim low As Double
    low = v - Int(v / 256#) * 256#
    XorLow = v - low + (CLng(low) Xor b)
End Function

' (v * m) mod 2^32, keeping everything inside the exact range of a Double
Private Function Mul32(ByVal v As Double, ByVal m As Double) As Double
    Dim lo As Double, hi As Double
    v = v - Int(v / 4294967296#) * 4294967296#
    hi = Int(v / 65536#)
    lo = v - hi * 65536#
    Mul32 = (lo * m + (hi * m - Int(hi * m / 65536#) * 65536#) * 65536#)
    Mul32 = Mul32 - Int(Mul32 / 4294967296#) * 4294967296#
End Function

Private Function Hex8(ByVal v As Double) As String
    Dim s As String, d As Long, i As Long
    v = Int(Abs(v))
    For i = 1 To 8
        d = CLng(v - Int(v / 16#) * 16#)
        s = Mid$("0123456789ABCDEF", d + 1, 1) & s
        v = Int(v / 16#)
    Next
    Hex8 = s
End Function

' ---------------------------------------------------------------- tiny JSON reader
' The server replies with a flat object we control, so string scanning is enough.
Private Function JVal(ByVal json As String, ByVal field As String) As String
    Dim p As Long, q As Long, tag As String
    tag = """" & field & """:"
    p = InStr(1, json, tag, vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(tag)
    Do While p <= Len(json) And Mid$(json, p, 1) = " "
        p = p + 1
    Loop
    If Mid$(json, p, 1) = """" Then
        p = p + 1
        q = InStr(p, json, """")
        If q = 0 Then Exit Function
        JVal = Replace(Mid$(json, p, q - p), "\""", """")
    Else
        q = p
        Do While q <= Len(json) And InStr(",}", Mid$(json, q, 1)) = 0
            q = q + 1
        Loop
        JVal = Trim$(Mid$(json, p, q - p))
    End If
End Function

Private Function JOk(ByVal json As String) As Boolean
    JOk = (LCase$(JVal(json, "ok")) = "true")
End Function

' ---------------------------------------------------------------- server call
' Returns "" when the server could not be reached (so callers can apply the grace period).
Private Function Api(ByVal action As String, ByVal key As String, Optional ByVal machine As String = "") As String
    Dim http As Object, body As String
    If machine = "" Then machine = MachineId()
    body = "{""action"":""" & action & """,""key"":""" & JEsc(key) & """,""machine"":""" & JEsc(machine) & _
           """,""machine_name"":""" & JEsc(Environ$("COMPUTERNAME")) & """,""user"":""" & JEsc(Environ$("USERNAME")) & """}"
    On Error GoTo Offline
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 4000, 8000, 8000, 12000
    http.Open "POST", API_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    Api = http.responseText
    Exit Function
Offline:
    Api = ""
End Function

Private Function JEsc(ByVal s As String) As String
    JEsc = Replace(Replace(s, "\", "\\"), """", "\""")
End Function

' ---------------------------------------------------------------- the gate
' Every ribbon tool calls this first. True = allowed to run.
Public Function LicAllow() As Boolean
    Dim key As String, r As String, exp As String, last As String, daysOff As Long

    If mChecked Then LicAllow = mValid: Exit Function

    key = Cfg("Key")
    If key = "" Then
        mChecked = True
        NotLicensed "No licence on this computer yet."
        Exit Function
    End If

    ' hard stop on a known expiry, even with no internet
    exp = Cfg("Expires")
    If exp <> "" And IsDate(exp) Then
        If CDate(exp) < Date Then
            mChecked = True
            NotLicensed "Your licence expired on " & Format(CDate(exp), "dd mmm yyyy") & "."
            Exit Function
        End If
    End If

    ' one online check per day
    last = Cfg("LastCheck")
    If last <> "" And IsDate(last) Then
        If CDate(last) = Date Then
            mChecked = True
            mValid = True
            LicAllow = True
            WarnIfEnding
            Exit Function
        End If
    End If

    r = Api("check", key)
    mChecked = True

    If r = "" Then                                   ' offline: allow within the grace window
        If last <> "" And IsDate(last) Then daysOff = DateDiff("d", CDate(last), Date) Else daysOff = GRACE_DAYS + 1
        If daysOff <= GRACE_DAYS Then
            mValid = True
            LicAllow = True
            Say "Working offline - licence re-checks when the internet is back (" & (GRACE_DAYS - daysOff) & " days left)."
        Else
            NotLicensed "This computer has been offline for more than " & GRACE_DAYS & " days, so the licence could not be re-checked."
        End If
        Exit Function
    End If

    If JOk(r) Then
        SetCfg "LastCheck", Format(Date, "yyyy-mm-dd")
        If JVal(r, "expires") <> "" Then SetCfg "Expires", JVal(r, "expires")
        mValid = True
        LicAllow = True
        WarnIfEnding
    Else
        If JVal(r, "error") = "not_activated" Then ForgetLicence   ' seat released elsewhere
        NotLicensed JVal(r, "message")
    End If
End Function

Private Sub WarnIfEnding()
    Dim exp As String, d As Long
    exp = Cfg("Expires")
    If exp = "" Or Not IsDate(exp) Then Exit Sub
    d = DateDiff("d", Date, CDate(exp))
    If d <= 14 And d >= 0 And Cfg("Warned") <> Format(Date, "yyyy-mm-dd") Then
        SetCfg "Warned", Format(Date, "yyyy-mm-dd")
        MsgBox "Your " & PRODUCT & " licence ends in " & d & " day(s), on " & Format(CDate(exp), "dd mmm yyyy") & "." & vbCrLf & vbCrLf & _
               "Please renew to keep using it.", vbInformation, PRODUCT
    End If
End Sub

Private Sub NotLicensed(ByVal why As String)
    mValid = False
    If MsgBox(why & vbCrLf & vbCrLf & "Enter a licence key now?", vbYesNo + vbExclamation, PRODUCT & " - licence needed") = vbYes Then
        LicManage
    End If
End Sub

' ---------------------------------------------------------------- activation / transfer
Public Sub LicManage()
    Dim key As String, r As String, msg As String

    key = Cfg("Key")
    If key <> "" Then
        msg = "Licence:  " & key & vbCrLf & _
              "Computer: " & Environ$("COMPUTERNAME") & vbCrLf & _
              "Valid to: " & Cfg("Expires") & vbCrLf & _
              "Last check: " & Cfg("LastCheck") & vbCrLf & vbCrLf & _
              "Yes = enter a different key" & vbCrLf & "No = remove the licence from this computer" & vbCrLf & "Cancel = close"
        Select Case MsgBox(msg, vbYesNoCancel + vbInformation, PRODUCT & " - licence")
            Case vbNo
                If MsgBox("Free this computer's seat so another computer can use it?", vbYesNo + vbQuestion, PRODUCT) = vbYes Then
                    r = Api("release", key)
                    If r = "" Then
                        MsgBox "Could not reach the licence server. Please try again when online.", vbExclamation, PRODUCT
                        Exit Sub
                    End If
                    ForgetLicence
                    MsgBox JVal(r, "message"), vbInformation, PRODUCT
                End If
                Exit Sub
            Case vbCancel
                Exit Sub
        End Select
    End If

    key = UCase$(Trim$(InputBox("Enter your " & PRODUCT & " licence key:" & vbCrLf & vbCrLf & _
                                "It looks like  EXL-XXXX-XXXX-XXXX", PRODUCT & " - activate", key)))
    If key = "" Then Exit Sub
    MsgBox LicActivateKey(key), vbInformation, PRODUCT
End Sub

' Activates this computer on the given key. Returns a message to show the user.
' Offers the transfer list when every seat is taken.
Public Function LicActivateKey(ByVal key As String, Optional ByVal silent As Boolean = False) As String
    Dim r As String
    key = UCase$(Trim$(key))
    r = Api("activate", key)
    If r = "" Then
        LicActivateKey = "Could not reach the licence server. Check the internet connection and try again."
        Exit Function
    End If

    If JOk(r) Then
        LicActivateKey = Stored(key, r)
        Exit Function
    End If

    If JVal(r, "error") = "seat_limit" And Not silent Then
        If OfferTransfer(key, JVal(r, "message")) Then
            r = Api("activate", key)
            If JOk(r) Then LicActivateKey = Stored(key, r) Else LicActivateKey = JVal(r, "message")
            Exit Function
        End If
    End If
    LicActivateKey = JVal(r, "message")
End Function

Private Function Stored(ByVal key As String, ByVal r As String) As String
    SetCfg "Key", key
    SetCfg "Expires", JVal(r, "expires")
    SetCfg "LastCheck", Format(Date, "yyyy-mm-dd")
    mChecked = True
    mValid = True
    Stored = "Activated." & vbCrLf & vbCrLf & _
             "Plan:      " & JVal(r, "plan") & vbCrLf & _
             "Computers: " & JVal(r, "seats_used") & " of " & JVal(r, "seats_total") & vbCrLf & _
             "Valid to:  " & JVal(r, "expires") & vbCrLf & vbCrLf & _
             "The 100x Pro Tools tab is ready to use."
End Function

' Support / diagnostics: frees this computer and forgets the key.
Public Function LicReleaseThis() As String
    Dim r As String, key As String
    key = Cfg("Key")
    If key = "" Then LicReleaseThis = "No licence on this computer.": Exit Function
    r = Api("release", key)
    If r = "" Then LicReleaseThis = "Could not reach the licence server.": Exit Function
    ForgetLicence
    LicReleaseThis = JVal(r, "message")
End Function

Public Function LicStatus() As String
    LicStatus = "key=" & Cfg("Key") & "; expires=" & Cfg("Expires") & _
                "; lastcheck=" & Cfg("LastCheck") & "; machine=" & MachineId()
End Function

' Lists the computers holding a seat and lets the user free one (the old PC need not be alive).
Private Function OfferTransfer(ByVal key As String, ByVal why As String) As Boolean
    Dim r As String, ids() As String, names() As String, n As Long, lst As String, pick As String, p As Long, blk As String
    r = Api("machines", key)
    If Not JOk(r) Then
        MsgBox why, vbExclamation, PRODUCT
        Exit Function
    End If

    ReDim ids(1 To 50): ReDim names(1 To 50)
    p = 1
    Do
        p = InStr(p, r, """machine_id"":")
        If p = 0 Or n >= 50 Then Exit Do
        blk = Mid$(r, p, 400)
        n = n + 1
        ids(n) = JVal(blk, "machine_id")
        names(n) = JVal(blk, "name") & IIf(JVal(blk, "user") <> "", " (" & JVal(blk, "user") & ")", "") & _
                   "  - last used " & JVal(blk, "last_seen")
        If ids(n) = MachineId() Then names(n) = names(n) & "  <- this computer"
        lst = lst & n & ".  " & names(n) & vbCrLf
        p = p + 12
    Loop
    If n = 0 Then
        MsgBox why, vbExclamation, PRODUCT
        Exit Function
    End If

    pick = InputBox(why & vbCrLf & vbCrLf & "Computers using this licence:" & vbCrLf & vbCrLf & lst & vbCrLf & _
                    "Type the number of the computer to remove (for example an old or dead PC)," & vbCrLf & _
                    "or leave empty to cancel.", PRODUCT & " - free a computer")
    If Not IsNumeric(pick) Then Exit Function
    p = CLng(pick)
    If p < 1 Or p > n Then Exit Function

    If MsgBox("Remove this computer from the licence?" & vbCrLf & vbCrLf & names(p) & vbCrLf & vbCrLf & _
              "It will stop working at its next daily check.", vbYesNo + vbExclamation, PRODUCT) <> vbYes Then Exit Function

    r = Api("release", key, ids(p))
    If JOk(r) Then
        OfferTransfer = True
    Else
        MsgBox JVal(r, "message"), vbExclamation, PRODUCT
    End If
End Function

Public Sub LicBuy()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink BUY_URL
End Sub

Public Sub LicAbout()
    Dim s As String
    s = PRODUCT & "  -  version 1.1" & vbCrLf & vbCrLf & _
        "202 tools in one ribbon tab." & vbCrLf & vbCrLf
    If Cfg("Key") = "" Then
        s = s & "Not licensed on this computer." & vbCrLf & "Use Licence > Activate to enter your key."
    Else
        s = s & "Licence:   " & Cfg("Key") & vbCrLf & _
                "Valid to:  " & Cfg("Expires") & vbCrLf & _
                "Computer:  " & Environ$("COMPUTERNAME") & vbCrLf & _
                "Machine id: " & MachineId()
    End If
    MsgBox s, vbInformation, PRODUCT
End Sub
