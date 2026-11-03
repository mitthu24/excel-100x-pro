Attribute VB_Name = "ProBiz"
Option Explicit

' ===== Template engine: LoadSpec() describes a sheet, Template() builds it =====
Private mTitle As String, mSub As String, mHdr As String, mW As String, mSample As String
Private mList As String, mFmt As String, mCalc As String, mSum As String, mCF As String, mBar As String, mRows As Long

Private Sub ColRange(ByVal spec As String, ByRef a As Long, ByRef b As Long)
    Dim p As Long
    p = InStr(spec, "-")
    If p > 0 Then
        a = Val(Left$(spec, p - 1))
        b = Val(Mid$(spec, p + 1))
    Else
        a = Val(spec)
        b = a
    End If
End Sub

Private Sub LoadSpec(ByVal key As String)
    Dim d As String, i As Long, h As String, w As String
    mRows = 30: mSample = "": mList = "": mFmt = "": mCalc = "": mSum = "": mCF = "": mBar = ""
    d = "=DATE(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")"
    Select Case key
    Case "expenses"
        mTitle = "Expense Tracker": mSub = "Log every expense. Pick Category, Mode and Status from the drop-downs."
        mHdr = "Date|Category|Description|Paid By|Mode|Amount|Status": mW = "12|16|34|16|10|14|14"
        mSample = d & "|Travel|Client visit cab|Rahul|UPI|1250|Pending~" & d & "|Office|Stationery|Priya|Cash|640|Approved"
        mList = "2=Travel,Food,Office,Software,Marketing,Utilities,Other;5=Cash,UPI,Card,Bank;7=Pending,Approved,Reimbursed,Rejected"
        mFmt = "1=dd-mmm-yy;6=INR": mSum = "6": mCF = "7=Rejected=R;7=Reimbursed=G;7=Pending=Y"
    Case "budget"
        mTitle = "Monthly Budget": mSub = "Variance and % used calculate automatically.": mRows = 20
        mHdr = "Category|Budget|Actual|Variance|% Used": mW = "22|16|16|16|12"
        mSample = "Salaries|500000|480000~Marketing|80000|95000~Rent|60000|60000"
        mCalc = "4==IF(B6='','',B6-C6)~5==IF(B6='','',IFERROR(C6/B6,0))"
        mFmt = "2-4=INR;5=0%": mSum = "2,3,4": mBar = "5"
    Case "todo"
        mTitle = "To-Do List": mSub = "Prioritise, assign and tick off.": mRows = 40
        mHdr = "Task|Owner|Priority|Due|Status|Notes": mW = "36|16|12|12|14|34"
        mSample = "Send quotation|Rahul|High|" & d & "|To Do|~Update stock sheet|Priya|Medium|" & d & "|In Progress|"
        mList = "3=High,Medium,Low;5=To Do,In Progress,Blocked,Done": mFmt = "4=dd-mmm-yy": mCF = "5=Done=G;5=Blocked=R"
    Case "timesheet"
        mTitle = "Timesheet": mSub = "Hours x rate for billable work.": mRows = 40
        mHdr = "Date|Project|Task|Hours|Billable|Rate|Amount": mW = "12|22|30|10|10|14|14"
        mSample = d & "|Website|Design review|3.5|Yes|1500~" & d & "|Internal|Team meeting|1|No|0"
        mCalc = "7==IF(E6='Yes',D6*F6,0)": mList = "5=Yes,No": mFmt = "1=dd-mmm-yy;6-7=INR": mSum = "4,7"
    Case "minutes"
        mTitle = "Meeting Minutes": mSub = "Decisions and action items.": mRows = 30
        mHdr = "Date|Topic|Discussion / Decision|Action Item|Owner|Due|Status": mW = "12|24|40|30|14|12|14"
        mSample = d & "|Q4 plan|Approved budget|Prepare PO|Rahul|" & d & "|Open"
        mList = "7=Open,In Progress,Done": mFmt = "1=dd-mmm-yy;6=dd-mmm-yy": mCF = "7=Done=G"
    Case "attendance"
        mTitle = "Attendance Sheet": mSub = "P = Present, A = Absent, L = Leave, H = Half day. Totals are automatic.": mRows = 40
        h = "Employee|Dept": w = "22|14"
        For i = 1 To 31
            h = h & "|" & i: w = w & "|4.5"
        Next
        mHdr = h & "|Present|Absent|Leave": mW = w & "|9|9|9"
        mSample = "Rahul|Sales~Priya|HR"
        mCalc = "34==IF(A6='','',COUNTIF(C6:AG6,'P')+COUNTIF(C6:AG6,'H')/2)~35==IF(A6='','',COUNTIF(C6:AG6,'A'))~36==IF(A6='','',COUNTIF(C6:AG6,'L'))"
        mList = "3-33=P,A,L,H": mCF = "3-33=A=R;3-33=P=G;3-33=L=Y"
    Case "project"
        mTitle = "Project Tracker": mSub = "Owners, deadlines and progress at a glance.": mRows = 40
        mHdr = "Task|Owner|Start|Due|Priority|Status|% Complete|Notes": mW = "30|14|12|12|11|14|13|34"
        mSample = "Kick-off meeting|Rahul|" & d & "|" & d & "|High|Done|1~Prepare proposal|Priya|" & d & "|" & d & "|Medium|In Progress|0.4"
        mList = "5=High,Medium,Low;6=To Do,In Progress,Blocked,Done": mFmt = "3-4=dd-mmm-yy;7=0%": mBar = "7": mCF = "6=Blocked=R;6=Done=G"
    Case "pipeline"
        mTitle = "Sales Pipeline": mSub = "Leads by stage with probability-weighted value.": mRows = 40
        mHdr = "Lead|Company|Contact|Phone|Source|Stage|Value|Probability|Weighted Value|Next Follow-up": mW = "18|22|18|14|13|14|14|12|16|14"
        mSample = "L-001|Acme Traders|Mr. Sharma|9876543210|Referral|Proposal|250000|0.5||" & d
        mCalc = "9==IF(G6='','',G6*H6)": mList = "5=Website,Referral,Cold Call,Social,Event;6=New,Contacted,Proposal,Negotiation,Won,Lost"
        mFmt = "7=INR;8=0%;9=INR;10=dd-mmm-yy": mSum = "7,9": mCF = "6=Won=G;6=Lost=R"
    Case "sales"
        mTitle = "Sales Tracker": mSub = "Quantity x rate less discount.": mRows = 40
        mHdr = "Date|Customer|Product|Qty|Rate|Discount %|Amount|Salesperson|Payment": mW = "12|22|22|8|13|11|15|16|12"
        mSample = d & "|Acme Traders|Plywood 18mm|10|1450|0.05||Rahul|Paid"
        mCalc = "7==IF(D6='','',D6*E6*(1-F6))": mList = "9=Paid,Pending,Partial": mFmt = "1=dd-mmm-yy;5=INR;6=0%;7=INR": mSum = "4,7": mCF = "9=Pending=Y;9=Paid=G"
    Case "commission"
        mTitle = "Sales Commission": mSub = "Slabs: 1% below 1L, 2% from 1L, 3% from 5L, 5% from 10L.": mRows = 20
        mHdr = "Salesperson|Sales|Slab %|Commission": mW = "24|18|12|18"
        mSample = "Rahul|650000~Priya|1200000~Amit|80000"
        mCalc = "3==IF(B6='','',IF(B6>=1000000,0.05,IF(B6>=500000,0.03,IF(B6>=100000,0.02,0.01))))~4==IF(B6='','',B6*C6)"
        mFmt = "2=INR;3=0%;4=INR": mSum = "2,4"
    Case "customers"
        mTitle = "Customer List": mSub = "Your customer master.": mRows = 50
        mHdr = "ID|Customer|Contact|Phone|Email|City|Segment|Status": mW = "10|26|20|14|28|16|13|11"
        mSample = "C-001|Acme Traders|Mr. Sharma|9876543210|sharma@acme.com|Mumbai|Wholesale|Active"
        mList = "7=Retail,Wholesale,Corporate;8=Active,Inactive": mCF = "8=Inactive=R"
    Case "content"
        mTitle = "Content Calendar": mSub = "Plan posts by channel.": mRows = 40
        mHdr = "Date|Channel|Topic|Owner|Status|Link": mW = "12|14|36|14|13|30"
        mSample = d & "|Instagram|Product reel|Priya|Drafting"
        mList = "2=Instagram,Facebook,LinkedIn,YouTube,Blog,Email;5=Idea,Drafting,Scheduled,Published": mFmt = "1=dd-mmm-yy": mCF = "5=Published=G"
    Case "campaign"
        mTitle = "Campaign ROI": mSub = "ROI = (revenue - spend) / spend. Cost per lead = spend / leads.": mRows = 20
        mHdr = "Campaign|Spend|Leads|Revenue|ROI|Cost per Lead": mW = "28|15|10|15|10|15"
        mSample = "Diwali offer|50000|120|300000~Google Ads|80000|200|240000"
        mCalc = "5==IF(B6='','',IFERROR((D6-B6)/B6,0))~6==IF(C6='','',IFERROR(B6/C6,0))": mFmt = "2=INR;4=INR;5=0%;6=INR": mSum = "2,3,4"
    Case "directory"
        mTitle = "Employee Directory": mSub = "Employee master data.": mRows = 60
        mHdr = "Emp ID|Name|Department|Designation|Join Date|Phone|Email|Salary|Status": mW = "10|24|16|18|12|14|28|14|11"
        mSample = "E-001|Rahul Verma|Sales|Manager|" & d & "|9876543210|rahul@company.com|60000|Active"
        mList = "9=Active,On Leave,Resigned": mFmt = "5=dd-mmm-yy;8=INR": mCF = "9=Resigned=R"
    Case "leave"
        mTitle = "Leave Tracker": mSub = "Working days are counted automatically (Mon-Fri).": mRows = 50
        mHdr = "Employee|Leave Type|From|To|Days|Reason|Status": mW = "22|14|12|12|8|30|12"
        mSample = "Rahul|Casual|" & d & "|" & d & "||Personal|Pending"
        mCalc = "5==IF(OR(C6='',D6=''),'',NETWORKDAYS(C6,D6))": mList = "2=Casual,Sick,Earned,Unpaid;7=Pending,Approved,Rejected"
        mFmt = "3-4=dd-mmm-yy": mCF = "7=Approved=G;7=Rejected=R"
    Case "payroll"
        mTitle = "Payroll Sheet": mSub = "HRA 40% of basic, PF 12% of basic. Edit the formulas to suit your policy.": mRows = 40
        mHdr = "Employee|Basic|HRA (40%)|Allowances|Gross|PF (12%)|Tax|Net Pay": mW = "24|14|14|14|14|14|12|15"
        mSample = "Rahul|30000||5000||||~Priya|45000||7000|||2000"
        mCalc = "3==IF(B6='','',B6*0.4)~5==IF(B6='','',B6+C6+D6)~6==IF(B6='','',B6*0.12)~8==IF(B6='','',E6-F6-G6)"
        mFmt = "2-8=INR": mSum = "5,8"
    Case "recruit"
        mTitle = "Recruitment Tracker": mSub = "Candidates by stage.": mRows = 50
        mHdr = "Candidate|Position|Source|Phone|Stage|Interview Date|Expected CTC|Remarks": mW = "22|22|14|14|13|14|14|30"
        mSample = "Neha Singh|Accountant|Referral|9876543210|Interview|" & d & "|420000"
        mList = "5=Applied,Screening,Interview,Offer,Joined,Rejected": mFmt = "6=dd-mmm-yy;7=INR": mCF = "5=Joined=G;5=Rejected=R"
    Case "invreg"
        mTitle = "Invoice Register": mSub = "GST, total, paid and balance - status is automatic.": mRows = 50
        mHdr = "Invoice No|Date|Client|Taxable|GST %|GST Amount|Total|Paid|Balance|Status": mW = "13|12|24|15|8|14|15|15|15|11"
        mSample = "INV-001|" & d & "|Acme Traders|100000|0.18|||50000"
        mCalc = "6==IF(D6='','',D6*E6)~7==IF(D6='','',D6+F6)~9==IF(D6='','',G6-H6)~10==IF(D6='','',IF(I6<=0,'Paid',IF(H6>0,'Partial','Unpaid')))"
        mFmt = "2=dd-mmm-yy;4=INR;5=0%;6-9=INR": mSum = "4,6,7,8,9": mCF = "10=Paid=G;10=Unpaid=R;10=Partial=Y"
    Case "cashflow"
        mTitle = "Cash Flow": mSub = "Net and cumulative cash by month.": mRows = 12
        mHdr = "Month|Inflow|Outflow|Net|Cumulative": mW = "14|18|18|18|18"
        mSample = "Apr|500000|420000~May|520000|450000~Jun|480000|470000~Jul||~Aug||~Sep||~Oct||~Nov||~Dec||~Jan||~Feb||~Mar||"
        mCalc = "4==IF(B6='','',B6-C6)~5==IF(B6='','',SUM($D$6:D6))": mFmt = "2-5=INR": mSum = "2,3,4"
    Case "pnl"
        mTitle = "Profit and Loss": mSub = "Monthly P&L with margin.": mRows = 12
        mHdr = "Month|Revenue|COGS|Gross Profit|Expenses|Net Profit|Margin %": mW = "14|16|16|16|16|16|12"
        mSample = "Apr|900000|520000||150000~May|950000|540000||160000~Jun|870000|500000||155000~Jul||~Aug||~Sep||~Oct||~Nov||~Dec||~Jan||~Feb||~Mar||"
        mCalc = "4==IF(B6='','',B6-C6)~6==IF(B6='','',D6-E6)~7==IF(B6='','',IFERROR(F6/B6,0))": mFmt = "2-6=INR;7=0.0%": mSum = "2,3,4,5,6"
    Case "receivable"
        mTitle = "Receivables Ageing": mSub = "Days overdue and ageing bucket update daily.": mRows = 50
        mHdr = "Customer|Invoice No|Invoice Date|Due Date|Amount|Days Overdue|Bucket": mW = "26|13|13|13|16|13|12"
        mSample = "Acme Traders|INV-001|" & d & "|" & d & "|100000"
        mCalc = "6==IF(D6='','',MAX(0,TODAY()-D6))~7==IF(F6='','',IF(F6=0,'Current',IF(F6<=30,'1-30',IF(F6<=60,'31-60',IF(F6<=90,'61-90','90+')))))"
        mFmt = "3-4=dd-mmm-yy;5=INR": mSum = "5": mCF = "7=90+=R;7=61-90=Y"
    Case "inventory"
        mTitle = "Inventory": mSub = "Status flips to REORDER when stock reaches the reorder level.": mRows = 50
        mHdr = "SKU|Item|Category|Qty|Reorder Level|Unit Cost|Stock Value|Status": mW = "12|28|16|9|13|14|16|11"
        mSample = "SKU-001|Plywood 18mm|Board|120|50|1450~SKU-002|Laminate 1mm|Surface|30|40|980"
        mCalc = "7==IF(D6='','',D6*F6)~8==IF(D6='','',IF(D6<=E6,'REORDER','OK'))": mFmt = "6-7=INR": mSum = "4,7": mCF = "8=REORDER=R"
    Case "purchase"
        mTitle = "Purchase Log": mSub = "Orders to vendors and delivery status.": mRows = 50
        mHdr = "PO No|Date|Vendor|Item|Qty|Rate|Amount|Delivery Date|Status": mW = "12|12|24|26|8|13|15|14|13"
        mSample = "PO-001|" & d & "|Shree Timber|Teak plank|40|2200||" & d & "|Ordered"
        mCalc = "7==IF(E6='','',E6*F6)": mList = "9=Ordered,In Transit,Received,Cancelled": mFmt = "2=dd-mmm-yy;6-7=INR;8=dd-mmm-yy": mSum = "7": mCF = "9=Received=G;9=Cancelled=R"
    Case "vendors"
        mTitle = "Vendor List": mSub = "Supplier master.": mRows = 50
        mHdr = "Vendor|Contact|Phone|Email|City|GSTIN|Payment Terms|Rating": mW = "26|18|14|28|14|18|14|8"
        mSample = "Shree Timber|Mr. Patel|9876543210|sales@shree.com|Surat|24ABCDE1234F1Z5|30 days|4"
        mList = "8=1,2,3,4,5"
    Case "production"
        mTitle = "Production Planner": mSub = "Orders by machine with progress.": mRows = 40
        mHdr = "Order|Product|Qty|Start|End|Machine|Status|% Done": mW = "12|26|9|12|12|14|12|11"
        mSample = "ORD-01|Door shutter|200|" & d & "|" & d & "|CNC-1|Running|0.6"
        mList = "7=Planned,Running,Done,Delayed": mFmt = "4-5=dd-mmm-yy;8=0%": mBar = "8": mCF = "7=Delayed=R;7=Done=G"
    Case "bom"
        mTitle = "Bill of Materials": mSub = "Components and cost roll-up per unit.": mRows = 30
        mHdr = "Component|Description|Qty per Unit|Unit Cost|Line Cost": mW = "22|34|13|14|14"
        mSample = "Panel|Plywood 18mm|2|1450~Screw|M4 x 30|12|2"
        mCalc = "5==IF(C6='','',C6*D6)": mFmt = "4-5=INR": mSum = "5"
    Case "shipment"
        mTitle = "Shipment Tracker": mSub = "AWB, route and status.": mRows = 50
        mHdr = "AWB|Customer|From|To|Dispatch|ETA|Status|Freight": mW = "14|24|14|14|12|12|13|13"
        mSample = "AWB1001|Acme Traders|Mumbai|Delhi|" & d & "|" & d & "|In Transit|3500"
        mList = "7=Booked,In Transit,Delivered,Delayed,Returned": mFmt = "5-6=dd-mmm-yy;8=INR": mSum = "8": mCF = "7=Delivered=G;7=Delayed=R"
    Case "maint"
        mTitle = "Maintenance Log": mSub = "Next due and status are automatic.": mRows = 40
        mHdr = "Asset|Location|Last Service|Interval (days)|Next Due|Status": mW = "26|18|13|14|13|12"
        mSample = "CNC Machine 1|Plant A|" & d & "|90"
        mCalc = "5==IF(C6='','',C6+D6)~6==IF(E6='','',IF(E6<TODAY(),'OVERDUE',IF(E6-TODAY()<=7,'Due soon','OK')))"
        mFmt = "3=dd-mmm-yy;5=dd-mmm-yy": mCF = "6=OVERDUE=R;6=Due soon=Y;6=OK=G"
    Case "grade"
        mTitle = "Gradebook": mSub = "Total, percentage and grade (out of 100 per subject).": mRows = 50
        mHdr = "Student|Roll No|Maths|Science|English|Social|Total|Percentage|Grade": mW = "24|9|9|9|9|9|10|12|8"
        mSample = "Aarav Shah|1|88|92|79|85~Diya Patel|2|67|71|80|74"
        mCalc = "7==IF(C6='','',SUM(C6:F6))~8==IF(C6='','',G6/400)~9==IF(H6='','',IF(H6>=0.9,'A+',IF(H6>=0.8,'A',IF(H6>=0.7,'B',IF(H6>=0.6,'C',IF(H6>=0.4,'D','F'))))))"
        mFmt = "8=0.0%": mCF = "9=F=R;9=A+=G"
    Case "patient"
        mTitle = "Patient Register": mSub = "Visits, doctors and fees.": mRows = 60
        mHdr = "Patient ID|Name|Age|Gender|Phone|Doctor|Visit Date|Diagnosis|Fee|Status": mW = "11|24|7|10|14|18|12|26|12|12"
        mSample = "P-001|Sunita Rao|42|Female|9876543210|Dr. Mehta|" & d & "|Fever|500|Visited"
        mList = "4=Male,Female,Other;10=Scheduled,Visited,Follow-up,Cancelled": mFmt = "7=dd-mmm-yy;9=INR": mSum = "9": mCF = "10=Cancelled=R"
    Case "property"
        mTitle = "Property Listing": mSub = "Price per sqft is automatic.": mRows = 50
        mHdr = "ID|Type|Location|Area (sqft)|Price|Price per sqft|Status|Owner Contact": mW = "9|11|24|12|16|14|12|16"
        mSample = "PR-001|Flat|Andheri West|950|18500000"
        mCalc = "6==IF(D6='','',IFERROR(E6/D6,0))": mList = "2=Flat,House,Plot,Shop,Office;7=Available,Booked,Sold": mFmt = "5-6=INR": mCF = "7=Sold=G;7=Booked=Y"
    Case "pos"
        mTitle = "Daily Sales Report": mSub = "Item-wise sales for the day.": mRows = 60
        mHdr = "Date|Item|Category|Qty|Price|Total|Payment": mW = "12|26|16|8|13|15|11"
        mSample = d & "|Notebook|Stationery|10|60"
        mCalc = "6==IF(D6='','',D6*E6)": mList = "7=Cash,UPI,Card": mFmt = "1=dd-mmm-yy;5-6=INR": mSum = "4,6"
    Case "menu"
        mTitle = "Menu Costing": mSub = "Margin per dish.": mRows = 40
        mHdr = "Dish|Category|Food Cost|Selling Price|Margin|Margin %": mW = "28|14|13|14|13|11"
        mSample = "Paneer Tikka|Starter|90|260~Masala Dosa|Main|40|140"
        mCalc = "5==IF(C6='','',D6-C6)~6==IF(D6='','',IFERROR(E6/D6,0))": mList = "2=Starter,Main,Dessert,Beverage": mFmt = "3-5=INR;6=0%": mBar = "6"
    Case "site"
        mTitle = "Site Daily Log": mSub = "Work done, labour and cost per day.": mRows = 60
        mHdr = "Date|Site|Work Done|Labour Count|Labour Cost|Material Cost|Total Cost|Remarks": mW = "12|18|34|12|14|14|14|28"
        mSample = d & "|Tower A|Slab casting|18|27000|65000"
        mCalc = "7==IF(COUNT(E6:F6)=0,'',E6+F6)": mFmt = "1=dd-mmm-yy;5-7=INR": mSum = "4,5,6,7"
    Case Else
        Err.Raise vbObjectError + 20, , "Unknown template: " & key
    End Select
    mCalc = Replace(mCalc, "'", Chr(34))
End Sub

Public Sub Template(ByVal key As String)
    Dim ws As Worksheet, h() As String, w() As String, rw() As String, f() As String, it() As String
    Dim i As Long, j As Long, nc As Long, n As Long, a As Long, b As Long, p As Long, fm As String
    Dim lo As ListObject, rg As Range, parts() As String, col As Variant, cl() As String
    LoadSpec key
    n = mRows
    h = Split(mHdr, "|")
    nc = UBound(h) + 1
    Set ws = NewWS(mTitle)
    StyleTitle ws, mTitle, mSub
    ws.Cells(5, 1).Resize(1, nc).Value = h
    w = Split(mW, "|")
    For j = 0 To UBound(w)
        ws.Columns(j + 1).ColumnWidth = Val(w(j))
    Next
    If mSample <> "" Then
        rw = Split(mSample, "~")
        For i = 0 To UBound(rw)
            f = Split(rw(i), "|")
            For j = 0 To UBound(f)
                If f(j) <> "" Then ws.Cells(6 + i, j + 1).Value = f(j)
            Next
        Next
    End If
    If mFmt <> "" Then
        it = Split(mFmt, ";")
        For i = 0 To UBound(it)
            p = InStr(it(i), "=")
            ColRange Left$(it(i), p - 1), a, b
            fm = Mid$(it(i), p + 1)
            If fm = "INR" Then fm = INRFormat()
            ws.Range(ws.Cells(6, a), ws.Cells(5 + n, b)).NumberFormat = fm
        Next
    End If
    If mCalc <> "" Then
        it = Split(mCalc, "~")
        For i = 0 To UBound(it)
            p = InStr(it(i), "=")
            a = Val(Left$(it(i), p - 1))
            ws.Range(ws.Cells(6, a), ws.Cells(5 + n, a)).Formula = Mid$(it(i), p + 1)
        Next
    End If
    Set rg = ws.Range(ws.Cells(5, 1), ws.Cells(5 + n, nc))
    Set lo = ws.ListObjects.Add(xlSrcRange, rg, , xlYes)
    lo.TableStyle = "TableStyleLight9"
    StyleHeader rg.Rows(1)
    rg.Rows(1).RowHeight = 26
    If mList <> "" Then
        it = Split(mList, ";")
        For i = 0 To UBound(it)
            p = InStr(it(i), "=")
            ColRange Left$(it(i), p - 1), a, b
            With ws.Range(ws.Cells(6, a), ws.Cells(5 + n, b))
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Mid$(it(i), p + 1)
                If b > a Then .HorizontalAlignment = xlCenter
            End With
        Next
    End If
    If mCF <> "" Then
        it = Split(mCF, ";")
        For i = 0 To UBound(it)
            parts = Split(it(i), "=")
            ColRange parts(0), a, b
            With ws.Range(ws.Cells(6, a), ws.Cells(5 + n, b)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & parts(1) & """")
                Select Case parts(2)
                    Case "R": .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
                    Case "G": .Interior.Color = RGB(209, 250, 229): .Font.Color = RGB(6, 95, 70)
                    Case "Y": .Interior.Color = RGB(254, 243, 199): .Font.Color = RGB(146, 64, 14)
                End Select
            End With
        Next
    End If
    If mBar <> "" Then
        cl = Split(mBar, ",")
        For i = 0 To UBound(cl)
            ws.Range(ws.Cells(6, Val(cl(i))), ws.Cells(5 + n, Val(cl(i)))).FormatConditions.AddDatabar.BarColor.Color = RGB(99, 102, 241)
        Next
    End If
    If mSum <> "" Then
        ws.Cells(3, 1).Value = "TOTAL"
        ws.Cells(3, 1).Font.Bold = True
        cl = Split(mSum, ",")
        For i = 0 To UBound(cl)
            col = Val(cl(i))
            With ws.Cells(3, col)
                .Formula = "=SUBTOTAL(109," & ws.Range(ws.Cells(6, col), ws.Cells(5 + n, col)).Address(False, False) & ")"
                .NumberFormat = ws.Cells(6, col).NumberFormat
                .Font.Bold = True
                .Font.Size = 13
                .Font.Color = Brand()
            End With
        Next
    End If
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    FreezeRows 5
    ws.Cells(6, 1).Select
End Sub

' ===== Documents =====
Public Sub MakeDocument(ByVal heading As String, ByVal nm As String)
    Dim ws As Worksheet, i As Long, r As Long, pre As String, lbl As String
    Set ws = NewWS(nm)
    ActiveWindow.DisplayGridlines = False
    pre = IIf(heading = "INVOICE", "INV", IIf(heading = "QUOTATION", "QUO", "PO"))
    lbl = IIf(heading = "INVOICE", "Invoice No", IIf(heading = "QUOTATION", "Quote No", "PO No"))
    ws.Range("A1").Value = "Your Company Name"
    ws.Range("A2").Value = "Address line  |  GSTIN  |  Phone  |  Email"
    ws.Range("F1").Value = heading
    ws.Range("A4").Value = IIf(heading = "PURCHASE ORDER", "VENDOR", "BILL TO")
    ws.Range("A5").Value = IIf(heading = "PURCHASE ORDER", "Vendor name", "Client name")
    ws.Range("A6").Value = "Address"
    ws.Range("E4").Value = lbl: ws.Range("F4").Value = pre & "-0001"
    ws.Range("E5").Value = "Date": ws.Range("F5").Formula = "=DATE(" & Year(Date) & "," & Month(Date) & "," & Day(Date) & ")"
    ws.Range("E6").Value = IIf(heading = "QUOTATION", "Valid Until", "Due Date"): ws.Range("F6").Formula = "=F5+15"
    ws.Range("F5:F6").NumberFormat = "dd-mmm-yyyy"
    With ws.Range("A1")
        .Font.Size = 20: .Font.Bold = True: .Font.Color = Brand()
    End With
    With ws.Range("F1")
        .Font.Size = 24: .Font.Bold = True: .Font.Color = Brand()
    End With
    ws.Range("A2").Font.Color = RGB(107, 114, 128)
    ws.Range("F1:F6").HorizontalAlignment = xlRight
    ws.Range("A4").Font.Bold = True: ws.Range("A4").Font.Color = RGB(107, 114, 128)
    ws.Range("E4:E6").Font.Bold = True
    ws.Range("A8:F8").Value = Array("Item", "Description", "Qty", "Rate", "GST %", "Amount")
    StyleHeader ws.Range("A8:F8")
    ws.Rows(8).RowHeight = 26
    ws.Range("A9:E9").Value = Array("Product A", "", 2, 1500, 0.18)
    ws.Range("A10:E10").Value = Array("Service B", "", 1, 5000, 0.18)
    ws.Range("F9:F18").Formula = "=IF(C9="""","""",C9*D9*(1+E9))"
    ws.Range("D9:D18,F9:F18,F20:F22").NumberFormat = INRFormat()
    ws.Range("E9:E18").NumberFormat = "0%"
    With ws.Range("A9:F18").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Color = RGB(229, 231, 235)
    End With
    With ws.Range("A9:F18").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(229, 231, 235)
    End With
    ws.Range("E20").Value = "Subtotal": ws.Range("F20").Formula = "=SUMPRODUCT(C9:C18,D9:D18)"
    ws.Range("E21").Value = "GST": ws.Range("F21").Formula = "=SUMPRODUCT(C9:C18,D9:D18,E9:E18)"
    ws.Range("E22").Value = "TOTAL": ws.Range("F22").Formula = "=SUM(F9:F18)"
    ws.Range("E20:F22").Font.Bold = True
    With ws.Range("E22:F22")
        .Interior.Color = RGB(238, 242, 255): .Font.Color = Brand(): .Font.Size = 14
    End With
    ws.Range("A23").Value = "Amount in words:"
    ws.Range("B23").Formula = "=AmountInWords(F22)"
    ws.Range("A23").Font.Bold = True
    ws.Range("A25").Value = IIf(heading = "QUOTATION", "Terms: Prices valid for 15 days. Taxes extra as applicable.", "Terms: Payment due within 15 days. Thank you for your business!")
    ws.Range("A25").Font.Italic = True: ws.Range("A25").Font.Color = RGB(107, 114, 128)
    ws.Columns(1).ColumnWidth = 18: ws.Columns(2).ColumnWidth = 30: ws.Columns(3).ColumnWidth = 8
    ws.Columns(4).ColumnWidth = 14: ws.Columns(5).ColumnWidth = 10: ws.Columns(6).ColumnWidth = 18
    PageFit ws, False
    ws.Range("A5").Select
End Sub

Public Sub MakePayslip()
    Dim ws As Worksheet
    Set ws = NewWS("Salary Slip")
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Value = "Your Company Name"
    ws.Range("A2").Value = "SALARY SLIP"
    ws.Range("A1").Font.Size = 20: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = Brand()
    ws.Range("A2").Font.Size = 12: ws.Range("A2").Font.Bold = True: ws.Range("A2").Font.Color = RGB(107, 114, 128)
    ws.Range("A4").Value = "Employee": ws.Range("B4").Value = "Employee name"
    ws.Range("A5").Value = "Department": ws.Range("B5").Value = "Sales"
    ws.Range("A6").Value = "Month": ws.Range("B6").Formula = "=TEXT(TODAY(),""mmmm yyyy"")"
    ws.Range("D4").Value = "Employee ID": ws.Range("E4").Value = "EMP-001"
    ws.Range("D5").Value = "Designation": ws.Range("E5").Value = "Executive"
    ws.Range("A4:A6,D4:D6").Font.Bold = True
    ws.Range("A8:B8").Value = Array("Earnings", "Amount")
    ws.Range("D8:E8").Value = Array("Deductions", "Amount")
    StyleHeader ws.Range("A8:B8"): StyleHeader ws.Range("D8:E8")
    ws.Range("A9").Value = "Basic": ws.Range("B9").Value = 30000
    ws.Range("A10").Value = "HRA": ws.Range("B10").Formula = "=B9*0.4"
    ws.Range("A11").Value = "Allowances": ws.Range("B11").Value = 5000
    ws.Range("A12").Value = "Bonus": ws.Range("B12").Value = 0
    ws.Range("D9").Value = "Provident Fund": ws.Range("E9").Formula = "=B9*0.12"
    ws.Range("D10").Value = "Income Tax": ws.Range("E10").Value = 0
    ws.Range("D11").Value = "Professional Tax": ws.Range("E11").Value = 200
    ws.Range("D12").Value = "Other": ws.Range("E12").Value = 0
    ws.Range("A13").Value = "Gross Earnings": ws.Range("B13").Formula = "=SUM(B9:B12)"
    ws.Range("D13").Value = "Total Deductions": ws.Range("E13").Formula = "=SUM(E9:E12)"
    ws.Range("A13:E13").Font.Bold = True
    ws.Range("A13:E13").Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Range("A15").Value = "NET PAY": ws.Range("B15").Formula = "=B13-E13"
    With ws.Range("A15:B15")
        .Interior.Color = RGB(238, 242, 255): .Font.Bold = True: .Font.Size = 14: .Font.Color = Brand()
    End With
    ws.Range("A16").Value = "In words:": ws.Range("B16").Formula = "=AmountInWords(B15)"
    ws.Range("B9:B13,E9:E13,B15").NumberFormat = INRFormat()
    ws.Columns(1).ColumnWidth = 22: ws.Columns(2).ColumnWidth = 22: ws.Columns(3).ColumnWidth = 4
    ws.Columns(4).ColumnWidth = 22: ws.Columns(5).ColumnWidth = 18
    ws.Range("A4").Select
End Sub

Public Sub LoanSchedule()
    Dim ws As Worksheet, amt As Double, rt As Double, yr As Long, n As Long
    amt = Val(InputBox("Loan amount:", "Loan EMI schedule", "500000"))
    rt = Val(InputBox("Annual interest rate (%):", "Loan EMI schedule", "9"))
    yr = Val(InputBox("Tenure in years:", "Loan EMI schedule", "5"))
    If amt <= 0 Or rt <= 0 Or yr < 1 Then Exit Sub
    n = yr * 12
    If n > 360 Then n = 360
    Set ws = NewWS("Loan Schedule")
    ActiveWindow.DisplayGridlines = False
    StyleTitle ws, "Loan EMI Schedule", "Change the blue inputs - the schedule updates."
    ws.Range("A4").Value = "Loan amount": ws.Range("B4").Value = amt
    ws.Range("A5").Value = "Annual rate": ws.Range("B5").Value = rt / 100
    ws.Range("A6").Value = "Tenure (years)": ws.Range("B6").Value = n / 12
    ws.Range("A7").Value = "Monthly EMI": ws.Range("B7").Formula = "=PMT(B5/12,B6*12,-B4)"
    ws.Range("A8").Value = "Total interest": ws.Range("B8").Formula = "=B7*B6*12-B4"
    ws.Range("B4:B6").Font.Color = RGB(37, 99, 235)
    ws.Range("B4,B7:B8").NumberFormat = INRFormat()
    ws.Range("B5").NumberFormat = "0.00%"
    ws.Range("A4:A8").Font.Bold = True
    ws.Range("A11:F11").Value = Array("No", "Opening", "EMI", "Interest", "Principal", "Closing")
    StyleHeader ws.Range("A11:F11")
    ws.Range("A12").Value = 1
    ws.Range("B12").Formula = "=B4"
    ws.Range("A13:A" & (11 + n)).Formula = "=A12+1"
    ws.Range("B13:B" & (11 + n)).Formula = "=F12"
    ws.Range("C12:C" & (11 + n)).Formula = "=$B$7"
    ws.Range("D12:D" & (11 + n)).Formula = "=B12*$B$5/12"
    ws.Range("E12:E" & (11 + n)).Formula = "=C12-D12"
    ws.Range("F12:F" & (11 + n)).Formula = "=B12-E12"
    ws.Range("B12:F" & (11 + n)).NumberFormat = INRFormat()
    ws.Range("A11:F" & (11 + n)).Borders(xlInsideHorizontal).Color = RGB(229, 231, 235)
    ws.Columns("A:F").ColumnWidth = 16
    ws.Range("A12").Select
    FreezeRows 11
End Sub
