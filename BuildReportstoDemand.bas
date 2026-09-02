Option Explicit

Private Const FIRST_DATA_ROW As Long = 6
Private Const NAME_COL As Long = 3
Private Const CAT_ROW As Long = 2
Private Const PROD_ROW As Long = 3
Private Const UNIT_ROW As Long = 4
Private Const UID_ROW As Long = 5
Private Const MAX_ROWS As Long = 40000

Private dSr() As Variant, dDist() As String, dDate() As Variant, dLoc() As String
Private dUid() As String, dVal() As Double, dCat() As String, dProd() As String
Private dUnit() As String, dRate() As Double, dWt() As Double
Private dCount As Long

Public Sub BuildBothReports()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ReDim dSr(1 To MAX_ROWS): ReDim dDist(1 To MAX_ROWS): ReDim dDate(1 To MAX_ROWS)
    ReDim dLoc(1 To MAX_ROWS): ReDim dUid(1 To MAX_ROWS): ReDim dVal(1 To MAX_ROWS)
    ReDim dCat(1 To MAX_ROWS): ReDim dProd(1 To MAX_ROWS): ReDim dUnit(1 To MAX_ROWS)
    ReDim dRate(1 To MAX_ROWS): ReDim dWt(1 To MAX_ROWS)
    dCount = 0

    CollectDetail
    EnsurePMMaster
    WriteReportA
    WriteReportB

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Reports built." & vbLf & "Rows: " & dCount & vbLf & vbLf & _
           "'Report A' = flat detail + location summary." & vbLf & _
           "'Report B' = detail + discount/GST columns." & vbLf & _
           "Fill the 'PM Master' sheet for Report B's numbers.", _
           vbInformation, "Demand Sheet Reports"
End Sub

Private Sub CollectDetail()
    Dim s As Variant, ws As Worksheet, wsR As Worksheet
    Dim lastRow As Long, lastCol As Long, distEnd As Long, wtRow As Long
    Dim r As Long, c As Long, rr As Long

    For Each s In CitySheets()
        If Not SheetExists(CStr(s)) Then GoTo NextSheet
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        Set wsR = Nothing
        If SheetExists(ws.Name & " R") Then Set wsR = ThisWorkbook.Worksheets(ws.Name & " R")

        lastRow = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).row
        lastCol = MeasureLastCol(ws)
        If lastRow < FIRST_DATA_ROW Or lastCol < 4 Then GoTo NextSheet

        distEnd = lastRow: wtRow = 0
        For rr = FIRST_DATA_ROW To lastRow
            Dim lb As String: lb = Norm(ws.Cells(rr, 2).Value) & " " & Norm(ws.Cells(rr, 3).Value)
            If distEnd = lastRow And InStr(lb, "TOTAL") > 0 Then distEnd = rr - 1
            If InStr(lb, "WEIGHT MEASUREMENT") > 0 Then wtRow = rr
        Next rr

        Dim loc As String, dt As Variant
        loc = CStr(ws.Cells(CAT_ROW, 3).Value)
        If Len(Trim$(loc)) = 0 Then loc = ws.Name
        dt = ws.Cells(3, 2).Value

        Dim cats() As String, prods() As String, units() As String, uids() As String
        ReDim cats(1 To lastCol): ReDim prods(1 To lastCol)
        ReDim units(1 To lastCol): ReDim uids(1 To lastCol)
        Dim curCat As String, curProd As String
        curCat = "": curProd = ""
        For c = 1 To lastCol
            Dim vc As Variant, vp As Variant
            vc = ws.Cells(CAT_ROW, c).Value
            vp = ws.Cells(PROD_ROW, c).Value
            If Not IsEmpty(vc) And Len(CStr(vc)) > 0 And c >= 4 Then curCat = Trim$(CleanTxt(vc))
            If Not IsEmpty(vp) And Len(CStr(vp)) > 0 Then curProd = Trim$(CleanTxt(vp))
            cats(c) = curCat: prods(c) = curProd
            units(c) = Trim$(CleanTxt(ws.Cells(UNIT_ROW, c).Value))
            uids(c) = Trim$(CleanTxt(ws.Cells(UID_ROW, c).Value))
        Next c

        For r = FIRST_DATA_ROW To distEnd
            Dim nm As String: nm = Trim$(CStr(ws.Cells(r, NAME_COL).Value & ""))
            If Len(nm) = 0 Then GoTo NextR
            If IsReserved(Norm(nm)) Then GoTo NextR
            For c = 4 To lastCol
                Dim cv As Variant: cv = ws.Cells(r, c).Value
                If IsEmpty(cv) Then GoTo NextC
                If Not IsNumeric(cv) Then GoTo NextC
                If Len(units(c)) = 0 Then GoTo NextC

                dCount = dCount + 1
                If dCount > MAX_ROWS Then dCount = MAX_ROWS: Exit Sub
                dSr(dCount) = ws.Cells(r, 2).Value
                dDist(dCount) = nm
                dDate(dCount) = dt
                dLoc(dCount) = loc
                dUid(dCount) = uids(c)
                dVal(dCount) = CDbl(cv)
                dCat(dCount) = cats(c)
                dProd(dCount) = prods(c)
                dUnit(dCount) = units(c)
                If Not wsR Is Nothing Then
                    If IsNumeric(wsR.Cells(r, c).Value) Then dRate(dCount) = CDbl(wsR.Cells(r, c).Value)
                End If
                If wtRow > 0 Then
                    If IsNumeric(ws.Cells(wtRow, c).Value) Then dWt(dCount) = CDbl(ws.Cells(wtRow, c).Value)
                End If
NextC:
            Next c
NextR:
        Next r
NextSheet:
    Next s
End Sub

Private Sub WriteReportA()
    Dim ws As Worksheet
    Set ws = FreshSheet("Report A")
    ws.Range("A1:M1").Value = Array("Sr. No.", "DISTRIBUTOR NAME", "Date", "Location", _
        "Unique Id", "Value", "Product Category", "Pruduct", "Text After Delimiter", _
        "Unique Id for rate", "Rate", "Weight M", "Sales Revenue")
    ws.Range("A1:M1").Font.Bold = True

    ws.Range("P1").Value = "Parameter": ws.Range("Q1").Value = "Source"
    ws.Range("P1:Q1").Font.Bold = True
    ws.Range("P2").Value = "File Name"
    ws.Range("Q2").Value = ThisWorkbook.FullName

    If dCount = 0 Then Exit Sub

    Dim locTot As Object: Set locTot = CreateObject("Scripting.Dictionary")
    Dim out() As Variant, i As Long, rev As Double
    ReDim out(1 To dCount, 1 To 13)
    For i = 1 To dCount
        rev = dVal(i) * dRate(i)
        out(i, 1) = dSr(i)
        out(i, 2) = dDist(i)
        out(i, 3) = dDate(i)
        out(i, 4) = dLoc(i)
        out(i, 5) = dUid(i)
        out(i, 6) = dVal(i)
        out(i, 7) = dCat(i)
        out(i, 8) = Trim$(dCat(i) & " " & dProd(i))
        out(i, 9) = dUnit(i)
        out(i, 10) = dUid(i) & dDist(i)
        out(i, 11) = dRate(i)
        out(i, 12) = dWt(i)
        out(i, 13) = rev
        If locTot.Exists(dLoc(i)) Then
            locTot(dLoc(i)) = locTot(dLoc(i)) + rev
        Else
            locTot.Add dLoc(i), rev
        End If
    Next i
    ws.Range("A2").Resize(dCount, 13).Value = out

    ws.Range("S1").Value = "Row Labels": ws.Range("T1").Value = "Sum of Sales Revenue"
    ws.Range("S1:T1").Font.Bold = True
    Dim k As Variant, n As Long, grand As Double
    n = 2
    For Each k In locTot.keys
        ws.Cells(n, 19).Value = k
        ws.Cells(n, 20).Value = locTot(k)
        grand = grand + locTot(k)
        n = n + 1
    Next k
    ws.Cells(n, 19).Value = "Grand Total"
    ws.Cells(n, 20).Value = grand
    ws.Range(ws.Cells(n, 19), ws.Cells(n, 20)).Font.Bold = True

    ws.Columns.AutoFit
End Sub

Private Sub WriteReportB()
    Dim ws As Worksheet
    Set ws = FreshSheet("Report B")
    ws.Range("A1:N1").Value = Array("Sr. No.", "DISTRIBUTOR NAME", "Date", "Location", _
        "Unique Id", "Value", "Product Category", "Text After Delimiter", "Rate", "Weight M", _
        "Discount Schm1", "Selected Location and Group for Milk", "Product", "PM.Pkt")
    ws.Range("O1:AA1").Value = Array("PM.GST Rate", "PM.Discount Thresh Hold", _
        "PM.Flat Discount", "PM.Discount if Conditions Meet", "Qty_Pcs", "Sales_Revenue", _
        "Sales_Rev Excl GST", "UniqueProduct", "Total Qty Sold Pcs", "Threshhold Lot", _
        "Rate Per Pc", "Product Discount", "Flat Discount")
    ws.Range("A1:AA1").Font.Bold = True

    ws.Range("AD1").Value = "Parameter": ws.Range("AE1").Value = "Source"
    ws.Range("AD1:AE1").Font.Bold = True
    ws.Range("AD2").Value = "File Name"
    ws.Range("AE2").Value = ThisWorkbook.FullName

    If dCount = 0 Then Exit Sub

    Dim pm As Object: Set pm = LoadPMMaster()
    Dim schm As Object: Set schm = LoadDiscountScheme()
    Dim tot As Object: Set tot = CreateObject("Scripting.Dictionary")

    Dim i As Long, key As String, uProd As String
    Dim pkt#, gst#, thr#, flat#, disc#, qtyPcs#
    Dim out() As Variant
    ReDim out(1 To dCount, 1 To 27)

    For i = 1 To dCount
        key = Norm(dProd(i)) & "|" & Norm(dUnit(i))
        pkt = PMVal(pm, key, 1): gst = PMVal(pm, key, 2)
        thr = PMVal(pm, key, 3): flat = PMVal(pm, key, 4): disc = PMVal(pm, key, 5)
        If pkt = 0 Then pkt = 1
        qtyPcs = dVal(i) * pkt
        uProd = dDist(i) & "-" & Replace(dProd(i), " ", "")

        out(i, 1) = dSr(i): out(i, 2) = dDist(i): out(i, 3) = dDate(i): out(i, 4) = dLoc(i)
        out(i, 5) = dUid(i): out(i, 6) = dVal(i): out(i, 7) = dCat(i): out(i, 8) = dUnit(i)
        out(i, 9) = dRate(i): out(i, 10) = dWt(i)
        out(i, 11) = SchemeFlag(schm, dDist(i), dProd(i))
        out(i, 12) = ""
        out(i, 13) = Replace(dProd(i), " ", "")
        out(i, 14) = pkt: out(i, 15) = gst: out(i, 16) = thr
        out(i, 17) = flat: out(i, 18) = disc
        out(i, 19) = qtyPcs
        out(i, 20) = dVal(i) * dRate(i)
        out(i, 21) = (dVal(i) * dRate(i)) / (1 + gst / 100)
        out(i, 22) = uProd
        If tot.Exists(uProd) Then
            tot(uProd) = tot(uProd) + qtyPcs
        Else
            tot.Add uProd, qtyPcs
        End If
    Next i

    Dim totPcs#, lot#, ratePc#
    For i = 1 To dCount
        uProd = CStr(out(i, 22))
        totPcs = tot(uProd)
        thr = CDbl(out(i, 16)): pkt = CDbl(out(i, 14))
        flat = CDbl(out(i, 17)): disc = CDbl(out(i, 18))
        lot = 0: If thr > 0 Then lot = Int(totPcs / thr)
        ratePc = 0: If pkt > 0 Then ratePc = dRate(i) / pkt
        out(i, 23) = totPcs
        out(i, 24) = lot
        out(i, 25) = ratePc
        out(i, 26) = lot * thr * ratePc * disc
        out(i, 27) = CDbl(out(i, 19)) * flat
    Next i

    ws.Range("A2").Resize(dCount, 27).Value = out
    ws.Columns.AutoFit
End Sub

Private Sub EnsurePMMaster()
    If SheetExists("PM Master") Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "PM Master"
    ws.Range("A1:G1").Value = Array("Product", "Unit", "PM.Pkt (pcs per unit)", _
        "PM.GST Rate %", "PM.Discount Thresh Hold", "PM.Flat Discount", _
        "PM.Discount if Conditions Meet")
    ws.Range("A1:G1").Font.Bold = True

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String, n As Long: n = 2
    For i = 1 To dCount
        k = dProd(i) & "|" & dUnit(i)
        If Not seen.Exists(k) Then
            seen.Add k, 1
            ws.Cells(n, 1).Value = dProd(i)
            ws.Cells(n, 2).Value = dUnit(i)
            n = n + 1
        End If
    Next i
    ws.Columns.AutoFit
End Sub

Private Function LoadPMMaster() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set LoadPMMaster = d
    If Not SheetExists("PM Master") Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("PM Master")
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For r = 2 To lastRow
        Dim k As String
        k = Norm(ws.Cells(r, 1).Value) & "|" & Norm(ws.Cells(r, 2).Value)
        If Len(k) > 1 And Not d.Exists(k) Then
            d.Add k, Array(Nz(ws.Cells(r, 3).Value), Nz(ws.Cells(r, 4).Value), _
                           Nz(ws.Cells(r, 5).Value), Nz(ws.Cells(r, 6).Value), _
                           Nz(ws.Cells(r, 7).Value))
        End If
    Next r
End Function

Private Function PMVal(ByVal pm As Object, ByVal key As String, ByVal idx As Long) As Double
    If pm Is Nothing Then Exit Function
    If Not pm.Exists(key) Then Exit Function
    Dim a As Variant: a = pm(key)
    PMVal = CDbl(a(idx - 1))
End Function

Private Function LoadDiscountScheme() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set LoadDiscountScheme = d
    If Not SheetExists("Discount Scheme") Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Discount Scheme")
    Dim lastRow As Long, r As Long, k As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For r = 2 To lastRow
        k = Norm(ws.Cells(r, 1).Value) & "|" & Norm(ws.Cells(r, 2).Value)
        If Len(Replace(k, "|", "")) > 0 And Not d.Exists(k) Then
            d.Add k, CStr(ws.Cells(r, 3).Value)
        End If
    Next r
End Function

Private Function SchemeFlag(ByVal schm As Object, ByVal dist As String, ByVal prod As String) As String
    If schm Is Nothing Then Exit Function
    If schm.Count = 0 Then Exit Function
    Dim k As String
    k = Norm(dist) & "|" & Norm(prod)
    If schm.Exists(k) Then SchemeFlag = schm(k): Exit Function
    k = Norm(dist) & "|"
    If schm.Exists(k) Then SchemeFlag = schm(k)
End Function

Private Function FreshSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    If SheetExists(nm) Then
        Set ws = ThisWorkbook.Worksheets(nm)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set FreshSheet = ws
End Function

Private Function CitySheets() As Variant
    CitySheets = Array("Raipur-1", "Raipur-2", "Raipur-3", "Bhilai", "Bilaspur", _
                       "Raigarh", "Odisha", "Mahasamund", "Rajnandgaon", "Korba", _
                       "Nagpur", "Other States", "Jagdalpur")
End Function

Private Function IsReserved(ByVal nm As String) As Boolean
    Select Case nm
        Case "RATE", "DB RATE", "MRP RATE", "ESPACIAL RATE", "SPECIAL RATE", _
             "ATP", "INSITUTION", "INSTITUTION"
            IsReserved = True
    End Select
End Function

Private Function MeasureLastCol(ByVal ws As Worksheet) As Long
    Dim lastCol As Long, j As Long, lab As String
    lastCol = ws.Cells(UNIT_ROW, ws.Columns.Count).End(xlToLeft).Column
    For j = 4 To lastCol
        lab = Norm(ws.Cells(PROD_ROW, j).Value)
        If InStr(lab, "TOTAL RUN") > 0 Or InStr(lab, "DEMAND SHEET AREA") > 0 Then
            MeasureLastCol = j - 1: Exit Function
        End If
    Next j
    MeasureLastCol = lastCol
End Function

Private Function CleanTxt(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Dim s As String
    s = CStr(v)
    s = Replace(s, vbLf, " "): s = Replace(s, vbCr, " "): s = Replace(s, vbTab, " ")
    On Error Resume Next
    s = Application.WorksheetFunction.Trim(s)
    On Error GoTo 0
    CleanTxt = Trim$(s)
End Function

Private Function Norm(ByVal v As Variant) As String
    Norm = UCase$(CleanTxt(v))
End Function

Private Function Nz(ByVal v As Variant) As Double
    If IsNumeric(v) Then Nz = CDbl(v)
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

