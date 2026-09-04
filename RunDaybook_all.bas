Option Explicit

Private Const STARTING_VCH_NO As Long = 12034

Private Const COL_FIRST As Long = 4
Private Const COL_LAST As Long = 138
Private Const OUT_SHEET As String = "Day Book"
Private Const UNIT_COL As Long = 6
Private Const BATCH_COL As Long = 11
Private Const MAX_DROPDOWN_CHARS As Long = 255
Private Const STOCK_FLAG_COL As Long = 20
Private Const PROD_ID_COL As Long = 21
Private Const PCS_QTY_COL As Long = 22

Private Const API_BASE As String = "http://103.171.96.225:3000"

Private mProdId(1 To 200) As String
Private mItem(1 To 200) As String
Private mLedger(1 To 200) As String
Private mUnit(1 To 200) As String
Private mDays(1 To 200) As Long
Private mPackSize(1 To 200) As Double

Private Const MAXL As Long = 300
Private bProdId(1 To MAXL) As String
Private bItem(1 To MAXL) As String
Private bLedger(1 To MAXL) As String
Private bQty(1 To MAXL) As Double
Private bQtyPcs(1 To MAXL) As Double
Private bUnit(1 To MAXL) As String
Private bRate(1 To MAXL) As Double
Private bAmt(1 To MAXL) As Double
Private bGst(1 To MAXL) As Double
Private bShelf(1 To MAXL) As Long
Private bBatch(1 To MAXL) As String

Private Const SV_COMPANY As String = "Vamaa Dairy Private Limited"
Private Const SV_STATE As String = "Chhattisgarh"

Private colParty As Collection
Private colPartyState As Collection
Private colLedgerNames As Collection
Private colPartyRow As Collection
Private colBatchMap As Collection
Private colPartyType As Collection

Private gAPIStockLoaded As Boolean
Private gAPIProductBatches As Object
Private gAPIPcsPerCrt As Object
Private gMFSShortfallMsg As String
Private gMFSLineCount As Long
Private gOtherLineCount As Long

' =====================================================================
' 1. LIVE API STOCK & BATCH ALLOCATION (FEFO IN EXACT PIECES,
'    MFS STRICT >75% UBD RULE, LIVE SHELF LIFE FROM API)
' =====================================================================

Public Sub LoadStockFromAPI()
    Set gAPIProductBatches = CreateObject("Scripting.Dictionary")
    Set gAPIPcsPerCrt = CreateObject("Scripting.Dictionary")
    gAPIStockLoaded = False

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    On Error GoTo APIFail
    http.Open "GET", API_BASE & "/api/stock/current", False
    http.setRequestHeader "Accept", "application/json"
    http.send

    If http.Status <> 200 Then
        MsgBox "Stock API error: HTTP " & http.Status & vbCrLf & _
               "URL: " & API_BASE & "/api/stock/current", vbExclamation, "API Warning"
        Exit Sub
    End If

    ParseAndStoreStockBatches http.responseText
    gAPIStockLoaded = True
    Exit Sub

APIFail:
    MsgBox "API se connect nahi ho paya: " & Err.Description & vbCrLf & _
           "Server check karo: " & API_BASE, vbExclamation, "API Connection Error"
End Sub

Public Sub AllocateBatchesForItem( _
        ByVal pid As String, _
        ByVal itemName As String, _
        ByVal qtyNeededPcs As Double, _
        ByVal shelfDays As Long, _
        ByVal vchDateVal As Date, _
        ByVal isMFS As Boolean, _
        ByRef allocBatch() As String, _
        ByRef allocQtyPcs() As Double, _
        ByRef allocCount As Long)

    On Error GoTo SafeFail

    If isMFS Then
        gMFSLineCount = gMFSLineCount + 1
    Else
        gOtherLineCount = gOtherLineCount + 1
    End If

    ReDim allocBatch(1 To 30)
    ReDim allocQtyPcs(1 To 30)
    allocCount = 0

    If Not gAPIStockLoaded Or gAPIProductBatches Is Nothing Then GoTo UseFallback

    Dim key As String
    key = pid
    If Not gAPIProductBatches.Exists(key) Then key = NormalizeName(itemName)
    If Not gAPIProductBatches.Exists(key) Then GoTo UseFallback

    Dim pd As Object
    Set pd = gAPIProductBatches(key)

    Dim totalBatches As Long
    totalBatches = val(pd("count"))
    If totalBatches = 0 Then GoTo UseFallback

    Dim remaining As Double
    remaining = qtyNeededPcs

    Dim i As Long, parts() As String
    Dim bc As String, ubdS As String, availPcs As Double, givePcs As Double
    Dim remPct As Double, ubdDate As Date
    Dim liveShelfDays As Double, effShelfDays As Double

    For i = 0 To totalBatches - 1
        If remaining <= 0.0001 Then Exit For
        If allocCount >= 30 Then Exit For

        If Not pd.Exists(CStr(i)) Then GoTo SkipThisBatch

        parts = Split(pd(CStr(i)), "|")
        If UBound(parts) < 2 Then GoTo SkipThisBatch

        bc = parts(0)
        ubdS = parts(1)
        availPcs = val(parts(2))
        If Len(bc) = 0 Then GoTo SkipThisBatch

        liveShelfDays = 0
        If UBound(parts) >= 3 Then liveShelfDays = val(parts(3))
        If liveShelfDays > 0 Then
            effShelfDays = liveShelfDays
        Else
            effShelfDays = shelfDays
        End If

        If availPcs > 0.0001 Then
            ' ---- MFS strict >75% shelf-life-remaining rule ----
            If isMFS And effShelfDays > 0 Then
                If IsDate(ubdS) Then
                    ubdDate = CDate(ubdS)
                    remPct = (CDbl(ubdDate) - CDbl(vchDateVal)) / CDbl(effShelfDays) * 100
                    If remPct <= 75 Then GoTo SkipThisBatch
                Else
                    GoTo SkipThisBatch
                End If
            End If

            givePcs = IIf(availPcs >= remaining, remaining, availPcs)

            allocCount = allocCount + 1
            allocBatch(allocCount) = bc
            allocQtyPcs(allocCount) = givePcs
            remaining = Round(remaining - givePcs, 4)

            pd(CStr(i)) = bc & "|" & ubdS & "|" & CStr(availPcs - givePcs) & "|" & CStr(liveShelfDays)
        End If
SkipThisBatch:
    Next i

    If allocCount = 0 Then GoTo UseFallback

    If remaining > 0.0001 Then
        allocCount = allocCount + 1
        allocBatch(allocCount) = "Primary Batch"
        allocQtyPcs(allocCount) = remaining

        If isMFS Then
            gMFSShortfallMsg = gMFSShortfallMsg & "- " & itemName & ": " & FmtQty(remaining) & _
                " pcs ke liye koi batch strict >75% UBD criteria pura nahi kar payi (Primary Batch mein daala)" & vbCrLf
        End If
    End If
    Exit Sub

UseFallback:
    ReDim allocBatch(1 To 30)
    ReDim allocQtyPcs(1 To 30)
    allocCount = 1
    allocBatch(1) = "Primary Batch"
    allocQtyPcs(1) = qtyNeededPcs

    If isMFS And shelfDays > 0 Then
        gMFSShortfallMsg = gMFSShortfallMsg & "- " & itemName & ": live batch data hi nahi mila (Primary Batch, >75% UBD check skip)" & vbCrLf
    End If
    Exit Sub

SafeFail:
    Dim errDesc As String
    errDesc = Err.Description
    ReDim allocBatch(1 To 30)
    ReDim allocQtyPcs(1 To 30)
    allocCount = 1
    allocBatch(1) = "Primary Batch"
    allocQtyPcs(1) = qtyNeededPcs
    gMFSShortfallMsg = gMFSShortfallMsg & "- " & itemName & ": batch allocation error (" & errDesc & "), Primary Batch mein daala" & vbCrLf
End Sub

Private Sub ShowMFSShortfallWarningIfAny()
    If Len(gMFSShortfallMsg) > 0 Then
        MsgBox "MFS (strict >75% UBD) criteria kuch lines ke liye pura nahi ho paya:" & vbCrLf & vbCrLf & _
               gMFSShortfallMsg & vbCrLf & "Stock team se fresh batch confirm kar lo.", _
               vbExclamation, "MFS Shelf-Life Warning"
    End If
End Sub

Public Sub ShowAllocationSummary()
    MsgBox "Batch allocation classification:" & vbCrLf & vbCrLf & _
           "MFS lines (strict >75% UBD rule applied): " & gMFSLineCount & vbCrLf & _
           "Other lines (total stock, FEFO only): " & gOtherLineCount, _
           vbInformation, "MFS vs Others - Allocation Summary"
End Sub

Private Sub ParseAndStoreStockBatches(ByVal json As String)
    Dim dataArr As String
    Dim catElems() As String, catCnt As Long
    Dim prodElems() As String, prodCnt As Long
    Dim batchElems() As String, batchCnt As Long
    Dim i As Long, j As Long, k As Long
    Dim pid As String, prodName As String, normName As String
    Dim batchNum As String, ubdVal As String, expiryVal As String
    Dim closingTotalPcs As Double
    Dim np As Long
    Dim liveShelfDays As Double

    Dim dPos As Long
    dPos = InStr(1, json, """data"":")
    If dPos = 0 Then Exit Sub
    dataArr = JSON_ExtractBlock(json, dPos, "[", "]")
    If Len(dataArr) = 0 Then Exit Sub

    JSON_ArrayElements dataArr, catElems, catCnt

    For i = 1 To catCnt
        Dim pPos As Long
        pPos = InStr(1, catElems(i), """products"":")
        If pPos = 0 Then GoTo NextCat

        Dim prodArr As String
        prodArr = JSON_ExtractBlock(catElems(i), pPos, "[", "]")
        If Len(prodArr) = 0 Then GoTo NextCat

        JSON_ArrayElements prodArr, prodElems, prodCnt

        For j = 1 To prodCnt
            On Error GoTo ProdErr

            Dim pJson As String
            pJson = prodElems(j)

            pid = JStr(pJson, "id", 1, np)
            prodName = JStr(pJson, "name", 1, np)
            If Len(pid) = 0 And Len(prodName) = 0 Then GoTo NextProd
            normName = NormalizeName(prodName)

            liveShelfDays = JNum(pJson, "shelfLifeDays", 1)

            Dim pcsPerCrtVal As Double, pcsDictKey As String
            pcsPerCrtVal = JNum(pJson, "pcsPerCrt", 1)
            pcsDictKey = IIf(Len(pid) > 0, pid, normName)
            If pcsPerCrtVal > 0 And Len(pcsDictKey) > 0 Then
                gAPIPcsPerCrt(pcsDictKey) = pcsPerCrtVal
                If Len(normName) > 0 Then gAPIPcsPerCrt(normName) = pcsPerCrtVal
            End If

            Dim blPos As Long
            blPos = InStr(1, pJson, """batchesList"":")
            If blPos = 0 Then GoTo NextProd

            Dim batchArr As String
            batchArr = JSON_ExtractBlock(pJson, blPos, "[", "]")
            If Len(batchArr) = 0 Then GoTo NextProd

            JSON_ArrayElements batchArr, batchElems, batchCnt

            Dim rawBatches() As String
            ReDim rawBatches(1 To batchCnt + 1)
            Dim bWithStock As Long
            bWithStock = 0

            For k = 1 To batchCnt
                Dim bJson As String
                bJson = batchElems(k)

                batchNum = JStr(bJson, "batchNumber", 1, np)
                ubdVal = JStr(bJson, "ubd", 1, np)
                expiryVal = JStr(bJson, "expiryDate", 1, np)

                If Len(ubdVal) = 0 Then ubdVal = expiryVal

                Dim clPos As Long
                clPos = InStr(1, bJson, """closing"":")
                Dim clBlock As String
                clBlock = JSON_ExtractBlock(bJson, clPos, "{", "}")
                closingTotalPcs = JNum(clBlock, "total", 1)

                If closingTotalPcs > 0 And Len(batchNum) > 0 Then
                    bWithStock = bWithStock + 1
                    rawBatches(bWithStock) = batchNum & "|" & ubdVal & "|" & CStr(closingTotalPcs) & "|" & CStr(liveShelfDays)
                End If
            Next k

            If bWithStock = 0 Then GoTo NextProd

            Dim m As Long, n As Long, swp As String
            Dim dM As Double, dN As Double
            For m = 1 To bWithStock - 1
                For n = m + 1 To bWithStock
                    dM = UbdSortValue(SplitPart(rawBatches(m), "|", 2))
                    dN = UbdSortValue(SplitPart(rawBatches(n), "|", 2))
                    If dN < dM Then
                        swp = rawBatches(m)
                        rawBatches(m) = rawBatches(n)
                        rawBatches(n) = swp
                    End If
                Next n
            Next m

            Dim dictKey As String
            dictKey = IIf(Len(pid) > 0, pid, normName)

            If Not gAPIProductBatches.Exists(dictKey) Then
                Set gAPIProductBatches(dictKey) = CreateObject("Scripting.Dictionary")
            End If
            Dim pd2 As Object
            Set pd2 = gAPIProductBatches(dictKey)
            For k = 1 To bWithStock
                pd2(CStr(k - 1)) = rawBatches(k)
            Next k
            pd2("count") = CStr(bWithStock)

            If Len(normName) > 0 And Not gAPIProductBatches.Exists(normName) Then
                Set gAPIProductBatches(normName) = pd2
            End If

NextProd:
            On Error GoTo 0
        Next j
        GoTo NextCat
ProdErr:
        Resume NextProd
NextCat:
    Next i
End Sub

Private Function BuildBatchCsvForItem(ByVal pid As String, ByVal itemName As String) As String
    Dim key As String, pd As Object
    Dim i As Long, totalBatches As Long, parts() As String
    Dim csv As String, bc As String, tmp As String

    BuildBatchCsvForItem = ""
    If Not gAPIStockLoaded Or gAPIProductBatches Is Nothing Then Exit Function

    key = pid
    If Not gAPIProductBatches.Exists(key) Then key = NormalizeName(itemName)
    If Not gAPIProductBatches.Exists(key) Then Exit Function

    Set pd = gAPIProductBatches(key)
    totalBatches = val(pd("count"))
    If totalBatches = 0 Then Exit Function

    csv = ""
    For i = 0 To totalBatches - 1
        parts = Split(pd(CStr(i)), "|")
        bc = parts(0)
        If Len(bc) > 0 Then
            If Len(csv) = 0 Then
                tmp = bc
            Else
                tmp = csv & "," & bc
            End If
            If Len(tmp) > MAX_DROPDOWN_CHARS Then Exit For
            csv = tmp
        End If
    Next i

    BuildBatchCsvForItem = csv
End Function

Private Sub ApplyItemBatchDropdown(ByVal o As Worksheet, ByVal r As Long, ByVal pid As String, ByVal itemName As String)
    Dim csv As String
    Dim cell As Range

    Set cell = o.Cells(r, BATCH_COL)
    cell.Validation.Delete

    csv = BuildBatchCsvForItem(pid, itemName)

    If Len(csv) > 0 Then
        cell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:=csv
        cell.Validation.IgnoreBlank = True
        cell.Validation.InCellDropdown = True
        cell.Validation.ShowInput = False
        cell.Validation.ShowError = True
    End If
End Sub

' =====================================================================
' 2. JSON PARSER HELPERS
' =====================================================================

Private Function JStr(ByVal json As String, ByVal key As String, ByVal fromPos As Long, ByRef nextPos As Long) As String
    Dim searchFor As String, p As Long, q As Long
    nextPos = fromPos
    searchFor = Chr(34) & key & Chr(34) & ":"
    p = InStr(fromPos, json, searchFor)
    If p = 0 Then Exit Function
    p = p + Len(searchFor)

    Do While p <= Len(json)
        Dim ch1 As String: ch1 = Mid$(json, p, 1)
        If ch1 <> " " And ch1 <> vbTab And ch1 <> vbLf And ch1 <> vbCr Then Exit Do
        p = p + 1
    Loop

    If Mid$(json, p, 1) = Chr(34) Then
        p = p + 1
        q = InStr(p, json, Chr(34))
        If q > 0 Then
            JStr = Mid$(json, p, q - p)
            nextPos = q + 1
        End If
    ElseIf Mid$(json, p, 4) = "null" Then
        JStr = ""
        nextPos = p + 4
    End If
End Function

Private Function JNum(ByVal json As String, ByVal key As String, ByVal fromPos As Long) As Double
    Dim searchFor As String, p As Long, q As Long, ch As String
    searchFor = Chr(34) & key & Chr(34) & ":"
    p = InStr(fromPos, json, searchFor)
    If p = 0 Then Exit Function
    p = p + Len(searchFor)

    Do While p <= Len(json)
        ch = Mid$(json, p, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        p = p + 1
    Loop

    q = p
    Do While q <= Len(json)
        ch = Mid$(json, q, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then
            q = q + 1
        Else
            Exit Do
        End If
    Loop

    If q > p Then JNum = val(Mid$(json, p, q - p))
End Function

Private Function JSON_ExtractBlock(ByVal json As String, ByVal fromPos As Long, ByVal openCh As String, ByVal closeCh As String) As String
    Dim p As Long, depth As Long, ch As String
    Dim isInQuote As Boolean, isEscaped As Boolean, startP As Long

    If fromPos < 1 Then fromPos = 1
    p = InStr(fromPos, json, openCh)
    If p = 0 Then Exit Function
    startP = p: depth = 0: isInQuote = False: isEscaped = False

    Do While p <= Len(json)
        ch = Mid$(json, p, 1)
        If isEscaped Then
            isEscaped = False
        ElseIf ch = "\" And isInQuote Then
            isEscaped = True
        ElseIf ch = Chr(34) Then
            isInQuote = Not isInQuote
        ElseIf Not isInQuote Then
            If ch = openCh Then
                depth = depth + 1
            ElseIf ch = closeCh Then
                depth = depth - 1
                If depth = 0 Then
                    JSON_ExtractBlock = Mid$(json, startP, p - startP + 1)
                    Exit Function
                End If
            End If
        End If
        p = p + 1
    Loop
End Function

Private Sub JSON_ArrayElements(ByVal arrJson As String, ByRef elems() As String, ByRef cnt As Long)
    ReDim elems(1 To 1000)
    cnt = 0
    Dim p As Long, depth As Long, ch As String
    Dim isInQuote As Boolean, isEscaped As Boolean, startP As Long

    p = 1
    Do While p <= Len(arrJson) And Mid$(arrJson, p, 1) <> "["
        p = p + 1
    Loop
    If p > Len(arrJson) Then Exit Sub
    p = p + 1

    depth = 0: isInQuote = False: isEscaped = False: startP = 0

    Do While p <= Len(arrJson)
        ch = Mid$(arrJson, p, 1)
        If isEscaped Then
            isEscaped = False
        ElseIf ch = "\" And isInQuote Then
            isEscaped = True
        ElseIf ch = Chr(34) Then
            isInQuote = Not isInQuote
        ElseIf Not isInQuote Then
            If ch = "{" Or ch = "[" Then
                If depth = 0 Then startP = p
                depth = depth + 1
            ElseIf ch = "}" Or ch = "]" Then
                depth = depth - 1
                If depth = 0 And startP > 0 Then
                    cnt = cnt + 1
                    If cnt <= UBound(elems) Then
                        elems(cnt) = Mid$(arrJson, startP, p - startP + 1)
                    End If
                    startP = 0
                End If
            End If
        End If
        p = p + 1
    Loop
End Sub

Private Function SplitPart(ByVal s As String, ByVal sep As String, ByVal part As Long) As String
    Dim arr() As String
    arr = Split(s, sep)
    If part - 1 <= UBound(arr) Then SplitPart = arr(part - 1)
End Function

Private Function UbdSortValue(ByVal ubdS As String) As Double
    If Len(Trim$(ubdS)) = 0 Then
        UbdSortValue = 2958465
    ElseIf IsDate(ubdS) Then
        UbdSortValue = CDbl(CDate(ubdS))
    Else
        UbdSortValue = 2958465
    End If
End Function

' =====================================================================
' 3. HELPERS & PACK SIZE ENGINE
' =====================================================================

Private Function GetExclusiveRate(ByVal inclRate As Double, ByVal gstPct As Double) As Double
    If gstPct > 0 Then
        GetExclusiveRate = Round(inclRate / (1 + gstPct / 100), 2)
    Else
        GetExclusiveRate = inclRate
    End If
End Function

Private Function IsDirectBillItem(ByVal leadCol As Long) As Boolean
    Select Case leadCol
        Case 83, 87, 89, 92, 93: IsDirectBillItem = True
        Case Else: IsDirectBillItem = False
    End Select
End Function

Private Sub ComputeLineQtyAndRate(ByVal c As Long, ByVal leadCol As Long, ByVal qvRaw As Double, ByVal rvIncl As Double, rws As Worksheet, ByVal r As Long, ByVal gstPct As Double, ByRef outQty As Double, ByRef outRate As Double)
    Dim caseRateRaw As Double, packSize As Double

    If c = leadCol Or IsDirectBillItem(leadCol) Then
        outQty = qvRaw
        outRate = GetExclusiveRate(rvIncl, gstPct)
    Else
        caseRateRaw = val(rws.Cells(r, leadCol).Value)

        If caseRateRaw > 0 And rvIncl > 0 Then
            packSize = caseRateRaw / rvIncl
            outQty = qvRaw / packSize
            outRate = GetExclusiveRate(caseRateRaw, gstPct)
        ElseIf mPackSize(leadCol) > 1 Then
            outQty = qvRaw / mPackSize(leadCol)
            outRate = GetExclusiveRate(rvIncl * mPackSize(leadCol), gstPct)
        Else
            outQty = qvRaw
            outRate = GetExclusiveRate(rvIncl, gstPct)
        End If
    End If
End Sub

Private Function GetPcsPerUnit(ByVal leadCol As Long) As Double
    If leadCol >= 1 And leadCol <= 200 Then
        If mPackSize(leadCol) > 0 Then
            GetPcsPerUnit = mPackSize(leadCol)
            Exit Function
        End If
    End If
    GetPcsPerUnit = 1
End Function

Private Function LocationToState(ByVal loc As String) As String
    Dim l As String: l = LCase$(Trim$(loc))
    Select Case l
        Case "odisha": LocationToState = "Odisha"
        Case "nagpur": LocationToState = "Maharashtra"
        Case "raipur", "raipur 1", "raipur 2", "raipur 3", "raipur-1", "raipur-2", "raipur-3", "bhilai", "bilaspur", "korba", "raigarh", "mahasamund", "rajnandgaon", "jagdalpur", "narayanpur", "rajim": LocationToState = SV_STATE
        Case Else: LocationToState = ""
    End Select
End Function

Private Function GetVchDate(ByVal wb As Workbook) As String
    Dim ws As Worksheet, d As Variant
    d = Empty

    For Each ws In wb.Worksheets
        If ws.Name = OUT_SHEET Then GoTo NextWs
        If Not SheetExists(wb, ws.Name & " R") Then GoTo NextWs
        If ws.Name Like "* R" Then GoTo NextWs

        d = ws.Cells(3, 2).Value
        If Not IsEmpty(d) Then
            If Len(Trim$(CStr(d))) > 0 Then Exit For
        End If
NextWs:
    Next ws

    If IsEmpty(d) Or Len(Trim$(CStr(d))) = 0 Then
        GetVchDate = Format(Date, "dd-mmm-yy")
    ElseIf IsDate(d) Then
        GetVchDate = Format(CDate(d), "dd-mmm-yy")
    Else
        GetVchDate = CStr(d)
    End If
End Function

Private Function IsExcludedOutlet(ByVal outlet As String) As Boolean
    Select Case LCase$(Trim$(outlet))
        Case "rate", "mrp rate", "espacial rate", "db rate", "institution", "insitution", "individual": IsExcludedOutlet = True
        Case Else: IsExcludedOutlet = False
    End Select
End Function

Private Function IsValidSrNo(ByVal srno As Variant) As Boolean
    If VarType(srno) = vbEmpty Then
        IsValidSrNo = False
    ElseIf Len(Trim$(CStr(srno))) = 0 Then
        IsValidSrNo = False
    ElseIf Not IsNumeric(srno) Then
        IsValidSrNo = False
    Else
        IsValidSrNo = True
    End If
End Function

Private Function NormalizeName(ByVal s As String) As String
    Dim i As Long, ch As String, outp As String
    outp = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            outp = outp & UCase$(ch)
        End If
    Next i
    NormalizeName = outp
End Function

Private Function IsSuperStockistParty(ByVal party As String) As Boolean
    Dim np As String: np = NormalizeName(party)
    Select Case True
        Case InStr(np, "UNIVERSALCORP") > 0: IsSuperStockistParty = True
        Case InStr(np, "TAPASWINI") > 0: IsSuperStockistParty = True
        Case InStr(np, "MPSAGENCY") > 0: IsSuperStockistParty = True
        Case InStr(np, "ANNAPURNA") > 0: IsSuperStockistParty = True
        Case Else: IsSuperStockistParty = False
    End Select
End Function

Private Function IsMilkLedger(ByVal ledger As String) As Boolean
    IsMilkLedger = (InStr(1, ledger, "Milk", vbTextCompare) > 0)
End Function

Private Function IsMFSOutlet(ByVal outlet As String) As Boolean
    Dim t As String
    On Error Resume Next
    t = colPartyType(LCase$(Trim$(outlet)))
    On Error GoTo 0
    IsMFSOutlet = (StrComp(Trim$(t), "MFS", vbTextCompare) = 0)
End Function

Private Sub LoadPartyStates(wb As Workbook)
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long, c As Long, colName As Long, colState As Long
    Dim nm As String, st As String

    Set colPartyState = New Collection: Set colLedgerNames = New Collection

    On Error Resume Next
    Set ws = wb.Worksheets("LEDGER")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    colName = 0: colState = 0
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case Trim$(CStr(ws.Cells(1, c).Value))
            Case "NAME": colName = c
            Case "OLDLEDSTATENAME": colState = c
        End Select
    Next c

    If colName = 0 Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).row

    For r = 2 To lastRow
        nm = Trim$(CStr(ws.Cells(r, colName).Value))
        If Len(nm) > 0 Then
            On Error Resume Next
            colLedgerNames.Add nm, LCase$(nm)
            On Error GoTo 0
        End If
        If colState > 0 Then
            st = Trim$(CStr(ws.Cells(r, colState).Value))
            If Len(nm) > 0 And Len(st) > 0 Then
                On Error Resume Next
                colPartyState.Add st, LCase$(nm)
                On Error GoTo 0
            End If
        End If
    Next r
End Sub

Private Function PartyState(ByVal party As String) As String
    On Error Resume Next
    PartyState = colPartyState(LCase$(Trim$(party)))
    On Error GoTo 0
End Function

Private Function LedgerMasterName(ByVal nm As String) As String
    On Error Resume Next
    LedgerMasterName = colLedgerNames(LCase$(Trim$(nm)))
    On Error GoTo 0
End Function

Private Function IsInterstate(ByVal party As String) As Boolean
    Dim st As String: st = PartyState(party)
    If Len(st) = 0 Then
        IsInterstate = False
    Else
        IsInterstate = (UCase$(st) <> UCase$(SV_STATE))
    End If
End Function

Private Sub LoadDistributors(wb As Workbook, Optional ByVal WarnIfMissing As Boolean = False)
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long, colLocation As Long, colOutlet As Long, colPartyLedger As Long
    Dim colType As Long
    Dim outlet As String, party As String, matched As String, loc As String, st As String, ptype As String
    Dim missingList As String, unknownStateList As String, missingCount As Long, unknownStateCount As Long

    Set colParty = New Collection: Set colPartyRow = New Collection
    Set colPartyType = New Collection

    On Error Resume Next
    Set ws = wb.Worksheets("Distributor Master")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    colLocation = FindHeaderColumn(ws, 4, "Location")
    colOutlet = FindHeaderColumn(ws, 4, "Distributor")
    colPartyLedger = FindHeaderColumn(ws, 4, "Distributor in Tally")
    colType = FindHeaderColumn(ws, 4, "Party in Tally")

    If colLocation = 0 Then colLocation = 1
    If colOutlet = 0 Then colOutlet = 2
    If colPartyLedger = 0 Then colPartyLedger = 4

    lastRow = ws.Cells(ws.Rows.Count, colOutlet).End(xlUp).row
    missingList = "": unknownStateList = "": missingCount = 0: unknownStateCount = 0

    For r = 5 To lastRow
        loc = Trim$(CStr(ws.Cells(r, colLocation).Value))
        outlet = Trim$(CStr(ws.Cells(r, colOutlet).Value))
        party = Trim$(CStr(ws.Cells(r, colPartyLedger).Value))
        If colType > 0 Then
            ptype = Trim$(CStr(ws.Cells(r, colType).Value))
        Else
            ptype = ""
        End If

        If Len(outlet) > 0 And Len(party) > 0 Then
            matched = LedgerMasterName(party)
            If Len(matched) > 0 Then
                party = matched
            Else
                missingCount = missingCount + 1
                If missingCount <= 40 Then missingList = missingList & "- " & outlet & " -> " & party & vbCrLf
            End If

            On Error Resume Next
            colParty.Add party, LCase$(outlet)
            colPartyRow.Add CStr(r), LCase$(outlet)
            If Len(ptype) > 0 Then colPartyType.Add ptype, LCase$(outlet)
            On Error GoTo 0

            st = LocationToState(loc)
            If Len(st) > 0 Then
                On Error Resume Next
                colPartyState.Add st, LCase$(party)
                On Error GoTo 0
            Else
                unknownStateCount = unknownStateCount + 1
                If unknownStateCount <= 40 Then unknownStateList = unknownStateList & "- " & outlet & " (" & party & ")" & vbCrLf
            End If
        End If
    Next r

    If WarnIfMissing And missingCount > 0 Then
        MsgBox missingCount & " party name(s) match nahi hue:" & vbCrLf & missingList, vbExclamation, "Party Name Mismatch"
    End If
End Sub

Private Function LookupParty(ByVal outlet As String) As String
    On Error Resume Next
    LookupParty = colParty(LCase$(Trim$(outlet)))
    On Error GoTo 0
End Function

Private Function PartyRow(ByVal outlet As String) As Long
    On Error Resume Next
    PartyRow = CLng(colPartyRow(LCase$(Trim$(outlet))))
    On Error GoTo 0
End Function

Private Sub GetPartyDetails(ByVal wb As Workbook, ByVal outlet As String, ByRef consName As String, ByRef consAddress As String, ByRef consState As String, ByRef consPin As String, ByRef consGST As String, ByRef buyerAddress As String, ByRef buyerGST As String, ByRef buyerPin As String)
    Dim ws As Worksheet, r As Long
    Dim cName As Long, cAddress As Long, cState As Long, cPin As Long, cgst As Long
    Dim cBuyerAddress As Long, cBuyerGST As Long, cBuyerPin As Long

    consName = "": consAddress = "": consState = "": consPin = "": consGST = ""
    buyerAddress = "": buyerGST = "": buyerPin = ""

    r = PartyRow(outlet)
    If r = 0 Then Exit Sub

    Set ws = wb.Worksheets("Distributor Master")
    cName = FindHeaderColumn(ws, 4, "Consignee Name")
    cAddress = FindHeaderColumn(ws, 4, "Consignee Address")
    cState = FindHeaderColumn(ws, 4, "Consignee State")
    cPin = FindHeaderColumn(ws, 4, "Consignee Pincode")
    cgst = FindHeaderColumn(ws, 4, "Consignee GSTIN")
    cBuyerAddress = FindHeaderColumn(ws, 4, "Buyer Address")
    cBuyerGST = FindHeaderColumn(ws, 4, "Buyer GSTIN")
    cBuyerPin = FindHeaderColumn(ws, 4, "Buyer Pincode")

    If cName > 0 Then consName = Trim$(CStr(ws.Cells(r, cName).Value))
    If cAddress > 0 Then consAddress = CStr(ws.Cells(r, cAddress).Value)
    If cState > 0 Then consState = Trim$(CStr(ws.Cells(r, cState).Value))
    If cPin > 0 Then consPin = Trim$(CStr(ws.Cells(r, cPin).Value))
    If cgst > 0 Then consGST = Trim$(CStr(ws.Cells(r, cgst).Value))
    If cBuyerAddress > 0 Then buyerAddress = CStr(ws.Cells(r, cBuyerAddress).Value)
    If cBuyerGST > 0 Then buyerGST = Trim$(CStr(ws.Cells(r, cBuyerGST).Value))
    If cBuyerPin > 0 Then buyerPin = Trim$(CStr(ws.Cells(r, cBuyerPin).Value))
End Sub

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim c As Long, lastC As Long
    lastC = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).Value)), headerText, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function

Private Sub LoadBatchMap(wb As Workbook)
    Dim dbws As Worksheet
    Dim r As Long, lastR As Long, batchCol As Long, cc As Long
    Dim curParty As String, cellB As String, bv As String, key As String

    Set colBatchMap = New Collection

    On Error Resume Next
    Set dbws = wb.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If dbws Is Nothing Then Exit Sub

    For cc = 1 To 30
        If Trim$(CStr(dbws.Cells(1, cc).Value)) = "Batch No." Then
            batchCol = cc: Exit For
        End If
    Next cc
    If batchCol = 0 Then Exit Sub

    lastR = dbws.Cells(dbws.Rows.Count, 2).End(xlUp).row
    curParty = ""

    For r = 3 To lastR
        If Len(Trim$(CStr(dbws.Cells(r, 7).Value))) > 0 Then
            curParty = Trim$(CStr(dbws.Cells(r, 2).Value))
        ElseIf Left$(CStr(dbws.Cells(r, 2).Value), 2) = "  " Then
            cellB = Trim$(CStr(dbws.Cells(r, 2).Value))
            bv = Trim$(CStr(dbws.Cells(r, batchCol).Value))
            If Len(bv) > 0 And Len(curParty) > 0 And Len(cellB) > 0 Then
                key = LCase$(curParty & "|" & cellB)
                On Error Resume Next
                colBatchMap.Add bv, key
                On Error GoTo 0
            End If
        End If
    Next r
End Sub

' =====================================================================
' 4. BUILD DAY BOOK & XML EXPORT
' =====================================================================

Public Sub RunDayBookAndTallyExport()
    Application.ScreenUpdating = False
    Call BuildDayBookFormat(ShowMsg:=False)
    Call ExportDayBookToTallyXML
    Application.ScreenUpdating = True
    MsgBox "Dono kaam ho gaye!" & vbCrLf & "1) 'Day Book' sheet ban gayi." & vbCrLf & "2) Tally Import XML file save ho gayi.", vbInformation, "Demand -> Day Book + Tally XML"
End Sub

Public Sub BuildDayBookFormat(Optional ByVal ShowMsg As Boolean = True)
    Dim wb As Workbook, ws As Worksheet, rws As Worksheet, o As Worksheet
    Dim r As Long, c As Long, lastRow As Long, oRow As Long
    Dim srno As Variant, outlet As String, party As String
    Dim curStart As Long, startOf() As Long
    Dim nL As Long, i As Long, vno As Long, vchDateStr As String, vchDateVal As Date
    Dim qv As Double, rv As Double, outQty As Double, outRate As Double
    Dim consName As String, consAddress As String, consState As String
    Dim consPin As String, consGST As String, buyerAddress As String
    Dim buyerGST As String, buyerPin As String

    Dim spLines As Object, spHeader As Object, spOrder As Collection
    Dim spKey As Variant, lineDict As Object, lk As Variant
    Dim arr As Variant, hdr As Variant, spCat As String, spFullKey As String, spCatSeq As Variant

    Set wb = ActiveWorkbook
    LoadMapping
    LoadPackSizes wb
    LoadPartyStates wb
    LoadDistributors wb
    LoadShelfLife
    vchDateStr = GetVchDate(wb)
    If IsDate(vchDateStr) Then
        vchDateVal = DateValue(vchDateStr)
    Else
        vchDateVal = Date
    End If
    LoadBatchMap wb
    LoadStockFromAPI
    ValidatePackSizesAgainstAPI
    gMFSShortfallMsg = ""
    gMFSLineCount = 0
    gOtherLineCount = 0

    Set spLines = CreateObject("Scripting.Dictionary")
    Set spHeader = CreateObject("Scripting.Dictionary")
    Set spOrder = New Collection

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(OUT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set o = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    o.Name = OUT_SHEET

    o.Range("A1:S1").Value = Array("Date", "Particulars", "Qty", "Rate", "Amount", "Unit", "Vch Type", "Vch No.", "Debit Amount", "Credit Amount", "Batch No.", "Consignee Name", "Consignee Address", "Consignee State", "Consignee Pincode", "Consignee GSTIN", "Buyer Address", "Buyer GSTIN", "Buyer Pincode")
    o.Cells(1, STOCK_FLAG_COL).Value = "Stock Deducted"
    o.Cells(1, PROD_ID_COL).Value = "Product ID"
    o.Cells(1, PCS_QTY_COL).Value = "PCS Qty (App)"

    o.Range("I2").Value = "Inwards Qty"
    o.Range("J2").Value = "Outwards Qty"
    o.Rows(1).Font.Bold = True
    oRow = 3

    Application.ScreenUpdating = False
    vno = STARTING_VCH_NO - 1

    For Each ws In wb.Worksheets
        If ws.Name = OUT_SHEET Then GoTo NextSheet
        If Not SheetExists(wb, ws.Name & " R") Then GoTo NextSheet
        If ws.Name Like "* R" Then GoTo NextSheet

        Set rws = wb.Worksheets(ws.Name & " R")

        ReDim startOf(COL_FIRST To COL_LAST)
        curStart = 0

        ' --- FIX: a product header can live in row 3 (normal case) OR, for a
        ' handful of products such as Masala Chaas, only in row 2. The old
        ' loop below checked row 3 only, so those columns silently inherited
        ' the previous product's block (this is what turned Masala Chaas
        ' orders into "Strawberry Lassi" lines in the Day Book). Now it
        ' checks both rows. ---
        For c = COL_FIRST To COL_LAST
            If Len(Trim$(CStr(ws.Cells(3, c).Value))) > 0 Then
                curStart = c
            ElseIf Len(Trim$(CStr(ws.Cells(2, c).Value))) > 0 Then
                curStart = c
            End If
            startOf(c) = curStart
        Next c

        lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row

        For r = 6 To lastRow
            srno = ws.Cells(r, 2).Value
            outlet = Trim$(CStr(ws.Cells(r, 3).Value))

            If Len(outlet) = 0 Then GoTo NextRow
            If Not IsValidSrNo(srno) Then GoTo NextRow
            If IsExcludedOutlet(outlet) Then GoTo NextRow

            party = LookupParty(outlet)
            If Len(party) = 0 Then party = outlet

            GetPartyDetails wb, outlet, consName, consAddress, consState, consPin, consGST, buyerAddress, buyerGST, buyerPin

            Dim isMFSLocal As Boolean
            isMFSLocal = IsMFSOutlet(outlet)

            nL = 0

            For c = COL_FIRST To COL_LAST
                qv = val(ws.Cells(r, c).Value)
                If qv = 0 Then GoTo NextCol
                If startOf(c) = 0 Then GoTo NextCol
                If Len(mItem(startOf(c))) = 0 Then GoTo NextCol

                rv = val(rws.Cells(r, c).Value)
                If rv = 0 Then GoTo NextCol
                If InStr(LCase$(Trim$(CStr(ws.Cells(4, c).Value))), "free") > 0 Then GoTo NextCol
                If nL >= MAXL Then GoTo NextCol

                Dim pidLocal As String, itemNameLocal As String, ledgerNameLocal As String, unitNameLocal As String, gstRateLocal As Double
                pidLocal = mProdId(startOf(c))
                itemNameLocal = mItem(startOf(c))
                ledgerNameLocal = mLedger(startOf(c))
                If Len(ledgerNameLocal) = 0 Then ledgerNameLocal = "Sales"
                unitNameLocal = mUnit(startOf(c))
                gstRateLocal = GstForLedger(ledgerNameLocal)

                Call ComputeLineQtyAndRate(c, startOf(c), qv, rv, rws, r, gstRateLocal, outQty, outRate)

                Dim pcsPerUnitLocal As Double, pcsNeededLocal As Double
                pcsPerUnitLocal = GetPcsPerUnit(startOf(c))
                pcsNeededLocal = outQty * pcsPerUnitLocal

                Dim aB() As String, aQPcs() As Double, aC As Long, allocIndex As Long
                Call AllocateBatchesForItem(pidLocal, itemNameLocal, pcsNeededLocal, mDays(startOf(c)), vchDateVal, isMFSLocal, aB, aQPcs, aC)

                For allocIndex = 1 To aC
                    If nL >= MAXL Then Exit For
                    nL = nL + 1
                    bProdId(nL) = pidLocal
                    bItem(nL) = itemNameLocal
                    bLedger(nL) = ledgerNameLocal
                    bUnit(nL) = unitNameLocal
                    bGst(nL) = gstRateLocal
                    bQtyPcs(nL) = aQPcs(allocIndex)

                    If pcsPerUnitLocal > 0 Then
                        bQty(nL) = aQPcs(allocIndex) / pcsPerUnitLocal
                    Else
                        bQty(nL) = aQPcs(allocIndex)
                    End If

                    bRate(nL) = outRate
                    bAmt(nL) = Round(bQty(nL) * outRate, 2)
                    bBatch(nL) = aB(allocIndex)
                    bShelf(nL) = mDays(startOf(c))
                Next allocIndex

NextCol:
            Next c

            If nL = 0 Then GoTo NextRow

            If IsSuperStockistParty(party) Then
                If Not spHeader.Exists(party) Then
                    spHeader(party) = Array(buyerAddress, buyerGST, buyerPin)
                    spOrder.Add party
                End If

                For i = 1 To nL
                    spCat = IIf(IsMilkLedger(bLedger(i)), "MILK", "OTHER")
                    spFullKey = party & Chr(1) & spCat

                    If Not spLines.Exists(spFullKey) Then
                        Set spLines(spFullKey) = CreateObject("Scripting.Dictionary")
                    End If

                    Set lineDict = spLines(spFullKey)
                    lk = bItem(i) & "|" & Format(Round(bRate(i), 2), "0.00")

                    If lineDict.Exists(lk) Then
                        arr = lineDict(lk)
                        arr(4) = arr(4) + bQty(i)
                        arr(5) = Round(arr(5) + bAmt(i), 2)
                        arr(9) = arr(9) + bQtyPcs(i)
                        lineDict(lk) = arr
                    Else
                        lineDict.Add lk, Array(bItem(i), bLedger(i), bUnit(i), bRate(i), bQty(i), bAmt(i), bGst(i), bBatch(i), bProdId(i), bQtyPcs(i))
                    End If
                Next i

                GoTo NextRow
            End If

            vno = vno + 1
            Call WriteDayBookVoucher(o, oRow, vno, party, nL, vchDateStr, consName, consAddress, consState, consPin, consGST, buyerAddress, buyerGST, buyerPin, False)

NextRow:
        Next r
NextSheet:
    Next ws

    For Each spKey In spOrder
        For Each spCatSeq In Array("MILK", "OTHER")
            spFullKey = CStr(spKey) & Chr(1) & CStr(spCatSeq)

            If spLines.Exists(spFullKey) Then
                Set lineDict = spLines(spFullKey)
                nL = 0

                For Each lk In lineDict.keys
                    nL = nL + 1
                    arr = lineDict(lk)
                    bItem(nL) = arr(0)
                    bLedger(nL) = arr(1)
                    bUnit(nL) = arr(2)
                    bRate(nL) = arr(3)
                    bQty(nL) = arr(4)
                    bAmt(nL) = Round(arr(5), 2)
                    bGst(nL) = arr(6)
                    bBatch(nL) = arr(7)
                    bProdId(nL) = arr(8)
                    bQtyPcs(nL) = arr(9)
                Next lk

                If nL > 0 Then
                    vno = vno + 1
                    hdr = spHeader(CStr(spKey))
                    Call WriteDayBookVoucher(o, oRow, vno, CStr(spKey), nL, vchDateStr, "", "", "", "", "", CStr(hdr(0)), CStr(hdr(1)), CStr(hdr(2)), True)
                End If
            End If
        Next spCatSeq
    Next spKey

    o.Columns("A").ColumnWidth = 13
    o.Columns("B").ColumnWidth = 40
    o.Columns("C:E").ColumnWidth = 12
    o.Columns("F").ColumnWidth = 10
    o.Columns("G:H").ColumnWidth = 12
    o.Columns("I:J").ColumnWidth = 13
    o.Columns("K").ColumnWidth = 16
    o.Columns("L").ColumnWidth = 24
    o.Columns("M").ColumnWidth = 42
    o.Columns("N").ColumnWidth = 18
    o.Columns("O").ColumnWidth = 16
    o.Columns("P").ColumnWidth = 20
    o.Columns("Q").ColumnWidth = 42
    o.Columns("R:S").ColumnWidth = 18
    o.Columns("T:V").ColumnWidth = 14

    o.Columns("M").WrapText = True
    o.Columns("Q").WrapText = True
    o.Rows("1:2").Font.Bold = True

    Call AddUpdateStockButton(o)
    Application.ScreenUpdating = True

    Call ShowMFSShortfallWarningIfAny
    Call ShowAllocationSummary

    If ShowMsg Then
        MsgBox "Ho gaya!" & vbCrLf & vbCrLf & (vno - STARTING_VCH_NO + 1) & " vouchers (Day Book format)." & vbCrLf & "'" & OUT_SHEET & "' sheet ban gayi.", vbInformation, "Demand -> Day Book"
    End If
End Sub

Private Sub AddUpdateStockButton(ByVal o As Worksheet)
    Dim btn As Button: Dim btnName As String
    btnName = "btnUpdateStock"

    On Error Resume Next
    o.Buttons(btnName).Delete
    On Error GoTo 0

    Set btn = o.Buttons.Add(o.Range("W1").Left, o.Range("W1").Top, 160, 28)
    btn.Name = btnName
    btn.Caption = "Post Sales to App"
    btn.OnAction = "UpdateStockFromDayBook"
End Sub

Private Sub WriteDayBookVoucher(ByVal o As Worksheet, ByRef oRow As Long, ByVal vno As Long, ByVal party As String, ByVal nL As Long, ByVal vchDateStr As String, ByVal consName As String, ByVal consAddress As String, ByVal consState As String, ByVal consPin As String, ByVal consGST As String, ByVal buyerAddress As String, ByVal buyerGST As String, ByVal buyerPin As String, ByVal skipConsignee As Boolean)

    Dim i As Long, k As Long
    Dim taxable As Double, cgst As Double, sgst As Double
    Dim gross As Double, billtotal As Double, roundoff As Double
    Dim curLed As String, doneLed As String, ledsum As Double
    Dim bLook As String

    taxable = 0: cgst = 0

    For i = 1 To nL
        taxable = taxable + bAmt(i)
        cgst = cgst + bAmt(i) * bGst(i) / 200
    Next i

    cgst = Round(cgst, 2): sgst = cgst
    gross = taxable + cgst + sgst
    billtotal = Round(gross, 0)
    roundoff = Round(billtotal - gross, 2)

    o.Cells(oRow, 1).Value = vchDateStr
    o.Cells(oRow, 2).Value = party
    o.Cells(oRow, 2).Font.Bold = True
    o.Cells(oRow, 7).Value = "Sales"
    o.Cells(oRow, 8).Value = vno
    o.Cells(oRow, 9).Value = billtotal

    If skipConsignee Then
        o.Cells(oRow, 12).Value = "": o.Cells(oRow, 13).Value = ""
        o.Cells(oRow, 14).Value = "": o.Cells(oRow, 15).Value = "": o.Cells(oRow, 16).Value = ""
    Else
        o.Cells(oRow, 12).Value = consName: o.Cells(oRow, 13).Value = consAddress
        o.Cells(oRow, 14).Value = consState: o.Cells(oRow, 15).Value = consPin
        o.Cells(oRow, 16).Value = consGST
    End If

    o.Cells(oRow, 17).Value = buyerAddress: o.Cells(oRow, 18).Value = buyerGST: o.Cells(oRow, 19).Value = buyerPin
    oRow = oRow + 1

    o.Cells(oRow, 2).Value = "New Ref": o.Cells(oRow, 3).Value = vno
    o.Cells(oRow, 5).Value = billtotal: o.Cells(oRow, 5).NumberFormat = "0.00"" Dr"""
    oRow = oRow + 1

    doneLed = "|"

    For i = 1 To nL
        curLed = bLedger(i)
        If InStr(doneLed, "|" & curLed & "|") > 0 Then GoTo NextLed

        doneLed = doneLed & curLed & "|"
        ledsum = 0

        For k = 1 To nL
            If bLedger(k) = curLed Then ledsum = ledsum + bAmt(k)
        Next k

        o.Cells(oRow, 2).Value = curLed: o.Cells(oRow, 10).Value = Round(ledsum, 2)
        oRow = oRow + 1

        For k = 1 To nL
            If bLedger(k) = curLed Then
                o.Cells(oRow, 2).Value = "  " & bItem(k)
                o.Cells(oRow, 3).Value = bQty(k)
                o.Cells(oRow, 4).Value = bRate(k)
                o.Cells(oRow, 5).Value = bAmt(k)
                o.Cells(oRow, UNIT_COL).Value = bUnit(k)

                bLook = bBatch(k)
                If Len(bLook) = 0 Then bLook = "Primary Batch"
                o.Cells(oRow, BATCH_COL).Value = bLook
                o.Cells(oRow, PROD_ID_COL).Value = bProdId(k)
                o.Cells(oRow, PCS_QTY_COL).Value = bQtyPcs(k)

                ApplyItemBatchDropdown o, oRow, bProdId(k), bItem(k)
                oRow = oRow + 1
            End If
        Next k

NextLed:
    Next i

    If cgst > 0 Then
        If IsInterstate(party) Then
            o.Cells(oRow, 2).Value = "Output IGST 5%"
            o.Cells(oRow, 10).Value = Round(cgst + sgst, 2)
            oRow = oRow + 1
        Else
            o.Cells(oRow, 2).Value = "Out Put CGST 2.5%"
            o.Cells(oRow, 10).Value = cgst
            oRow = oRow + 1
            o.Cells(oRow, 2).Value = "Out Put SGST 2.5%"
            o.Cells(oRow, 10).Value = sgst
            oRow = oRow + 1
        End If
    End If

    If Abs(roundoff) >= 0.005 Then
        o.Cells(oRow, 2).Value = "Roundoff"
        If roundoff < 0 Then
            o.Cells(oRow, 9).Value = Round(-roundoff, 2)
        Else
            o.Cells(oRow, 10).Value = Round(roundoff, 2)
        End If
        oRow = oRow + 1
    End If
End Sub

Private Function GstForLedger(ByVal ledger As String) As Double
    Dim l As String: l = UCase$(Trim$(ledger))
    Select Case l
        Case "MILK SALES", "CREAM SALES", "GAIA PANEER SALES", "LOOSE PANEER SALES", "UHT MILK SALE": GstForLedger = 0
        Case Else: GstForLedger = 5
    End Select
End Function

Private Sub AddMap(ByVal col As Long, ByVal pid As String, ByVal item As String, ByVal ledger As String, ByVal unit As String)
    mProdId(col) = pid: mItem(col) = item: mLedger(col) = ledger: mUnit(col) = unit
End Sub

Private Sub LoadMapping()
    Dim i As Long
    For i = 1 To 200: mProdId(i) = "": mItem(i) = "": mLedger(i) = "": mUnit(i) = "": Next i

    AddMap 4, "1", "Gaia Doodh Toned 500ml- Pkt", "Milk Sales", "CRT"
    AddMap 7, "2", "Gaia Cow Milk 500 ML (Crt)", "Milk Sales", "CRT"
    AddMap 10, "3", "Gaia Maxx UHT 110ml", "UHT Milk Sale", "CBX"
    AddMap 13, "4", "Gaia Maxx UHT 400ml", "UHT Milk Sale", "CBX"
    AddMap 16, "5", "Gaia Premium Milk UHT Tetra Pack 1 Ltr", "UHT Milk Sale", "CBX"
    AddMap 19, "6", "Gaia Lite Milk UHT Tetra Pack 1Ltr", "UHT Milk Sale", "CBX"
    AddMap 22, "7", "Gaia Gold Milk UHT Tetra Pack 1Ltr", "UHT Milk Sale", "CBX"
    AddMap 25, "16", "Gaia Mishti Doi 80g", "Dahi Sales", "CBX"
    AddMap 28, "14", "Gaia Dahi Cup 200 Gms Box(6Pc)", "Dahi Sales", "CBX"
    AddMap 31, "15", "Gaia Dahi (Cup) 200 Gms", "Dahi Sales", "CBX"
    AddMap 34, "8", "Gaia Dahi (Crt) 200 Gms 60 Pkt", "Dahi Sales", "CRT"
    AddMap 37, "9", "Gaia Dahi 400 Gms (Crt) 30 Pkt", "Dahi Sales", "CRT"
    AddMap 40, "10", "Gaia Dahi 1 Kg- 12 Packets", "Dahi Sales", "CRT"
    AddMap 43, "11", "Gaia Dahi 5kg Pouch (2pcs)", "Dahi Sales", "CRT"
    AddMap 46, "12", "Gaia Dahi  Bucket 5 Kg", "Dahi Sales", "PCS"
    AddMap 47, "13", "Gaia Dahi  Bucket 15 Kg", "Dahi Sales", "PCS"
    AddMap 48, "17", "Gaia Premium Sweet Dahi (Lychee) Cup 90 Gms", "Dahi Sales", "CBX"
    AddMap 51, "18", "Gaia Premium Sweet Dahi (Muskmelon) Cup 90 Gms  O", "Dahi Sales", "PCS"
    AddMap 54, "19", "Gaia Kadhi Dahi (Crt) 200 Gm 60 Pkt", "Dahi Sales", "CRT"
    AddMap 57, "20", "Gaia Kadhi Dahi 1 Kg Crt (12 Pcs)", "Dahi Sales", "CRT"
    AddMap 60, "21", "Gaia Kadhi Dahi 5kg Pouch (2 Pcs)", "Dahi Sales", "CRT"
    AddMap 64, "23", "Gaia Kadhi Dahi Bucket 15 Kg", "Dahi Sales", "PCS"
    AddMap 68, "27", "Gaia Sweet Lassi 180ml Pouch", "Lassi Sales", "CRT"
    AddMap 71, "24", "Gaia Plain Lassi 180 Ml (Glass Cup 10)", "Lassi Sales", "CBX"
    AddMap 74, "25", "Mango Lassi 180 Ml - 10Pcs", "Flavoured Lassi", "CBX"
    AddMap 77, "28", "Strawberry Lassi 180 Ml - 10 Pcs", "Flavoured Lassi", "CBX"

    ' *** ADDED - this is the second half of the fix ***
    ' Columns 65 and 80 (Masala Chaas 200 Grms / Masala Chaas Glaas 180 Grms)
    ' never had an AddMap entry at all. Even with the row-2 header fix above,
    ' these columns would now correctly start their OWN block - but mItem()
    ' for that block would still be blank, so the line would just be silently
    ' dropped (GoTo NextCol) instead of appearing in the Day Book.
    '
    ' >>> CONFIRM BEFORE USE <<<
    ' - "22" and "26" below are placeholder product IDs (the only ids in the
    '   1-53 range NOT already used elsewhere in this list). Replace them with
    '   the real Masala Chaas product IDs from your stock API / product master.
    ' - Confirm "Masala Chaas 200 Grms" / "Masala Chaas Glaas 180 Grms" match
    '   your Tally stock item names exactly (case/spacing matters for XML import).
    ' - Confirm "Chaas Sales" is the correct Tally ledger name (vs e.g. "Lassi Sales").
    ' - Pack sizes/shelf life below (in LoadPackSizes / LoadShelfLife) are guesses
    '   too - confirm against your other Chaas SKUs.
    AddMap 65, "22", "Masala Chaas 200 Grms", "Chaas Sales", "CRT"
    AddMap 80, "26", "Masala Chaas Glaas 180 Grms", "Chaas Sales", "CBX"

    AddMap 83, "29", "Gaia Paneer 200gms", "Gaia Paneer Sales", "PCS"
    AddMap 87, "30", "Gaia Paneer 500Gms", "Gaia Paneer Sales", "PCS"
    AddMap 89, "31", "Gaia Paneer 1kg", "Gaia Paneer Sales", "PCS"
    AddMap 92, "32", "Loose 1 Kg", "Loose Paneer Sales", "PCS"
    AddMap 93, "33", "Loose 5 Kg", "Loose Paneer Sales", "PCS"
    AddMap 97, "35", "Gaia Ghee 200 ML Jar", "Ghee Sales", "CBX"
    AddMap 99, "36", "Gaia Ghee 500 ML Jar", "Ghee Sales", "CBX"
    AddMap 101, "37", "Gaia  Ghee 1 Ltr Jar 18", "Ghee Sales", "CBX"
    AddMap 103, "38", "Gaia Ghee Ceka Pack 1ltr", "Ghee Sales", "CBX"
    AddMap 105, "45", "Gaia Cow Ghee Ceka Pack 1ltr", "Ghee Sales", "CBX"
    AddMap 107, "39", "Gaia Premium Desi Ghee Ceka Pack 900ml", "Ghee Sales", "CBX"
    AddMap 109, "44", "Gaia Pure Cow Ghee Ceka Pack 900ml", "Ghee Sales", "CBX"
    AddMap 111, "40", "Gaia Ghee 5 Ltr.Jar", "Ghee Sales", "CBX"
    AddMap 113, "41", "Gaia Cow Ghee 200 Ml Jar", "Ghee Sales", "CBX"
    AddMap 115, "42", "Gaia Cow Ghee 500 Ml Jar", "Ghee Sales", "CBX"
    AddMap 117, "43", "Gaia Cow Ghee 1Ltr Jar", "Ghee Sales", "CBX"
    AddMap 119, "34", "Gaia Premium Desi Ghee 20ml", "Ghee Sales", "CBX"
    AddMap 122, "46", "Gaia Ghee 15 Kilogram", "Ghee Sales", "PCS"
    AddMap 123, "48", "Gaia Shrikhand KE 80g", "Shrikhand Sale", "CBX"
    AddMap 126, "49", "Gaia Shrikhand KE 80g-Box (6Pc)", "Shrikhand Sale", "CBX"
    AddMap 128, "47", "Gaia Shahi Rabdi 80g  (1cup*6pcs)", "Rabdi Sales", "CBX"
    AddMap 130, "47", "Gaia Shahi Rabdi 80g  (1cup*6pcs)", "Rabdi Sales", "CBX"
    AddMap 132, "47", "Gaia Shahi Rabdi 80g-Box(6Pc)", "Rabdi Sales", "CBX"
    AddMap 134, "50", "Gaia Peda 200 Gm Pcs", "Peda Sales", "CBX"
    AddMap 136, "51", "Gaia Kesar Peda 200 Gm Pcs", "Peda Sales", "CBX"
    AddMap 138, "53", "Khowa Brown ( Unsweetened )", "Khowa Sales", "KG"
End Sub

Private Sub LoadPackSizes(wb As Workbook)
    Dim i As Long
    For i = 1 To 200: mPackSize(i) = 1: Next i

    SetPackSize 4, 24     ' id1  Gaia Doodh Toned 500ml- Pkt        (CRT)
    SetPackSize 7, 24     ' id2  Gaia Cow Milk 500 ML (Crt)         (CRT)
    SetPackSize 10, 40    ' id3  Gaia Maxx UHT 110ml                (CBX)
    SetPackSize 13, 20    ' id4  Gaia Maxx UHT 400ml                (CBX)
    SetPackSize 16, 12    ' id5  Gaia Premium Milk UHT 1 Ltr        (CBX)
    SetPackSize 19, 12    ' id6  Gaia Lite Milk UHT 1Ltr             (CBX)
    SetPackSize 22, 12    ' id7  Gaia Gold Milk UHT 1Ltr             (CBX)
    SetPackSize 25, 12    ' id16 Gaia Mishti Doi 80g                 (CBX)
    SetPackSize 28, 6     ' id14 Gaia Dahi Cup 200 Gms Box(6Pc)      (CBX)
    SetPackSize 31, 30    ' id15 Gaia Dahi (Cup) 200 Gms             (CBX)
    SetPackSize 34, 60    ' id8  Gaia Dahi (Crt) 200 Gms 60 Pkt      (CRT)
    SetPackSize 37, 30    ' id9  Gaia Dahi 400 Gms (Crt) 30 Pkt      (CRT)
    SetPackSize 40, 12    ' id10 Gaia Dahi 1 Kg- 12 Packets          (CRT)
    SetPackSize 43, 2     ' id11 Gaia Dahi 5kg Pouch (2pcs)          (CRT)
    SetPackSize 46, 1     ' id12 Gaia Dahi Bucket 5 Kg               (PCS)
    SetPackSize 47, 1     ' id13 Gaia Dahi Bucket 15 Kg              (PCS)
    SetPackSize 48, 12    ' id17 Gaia Prem Sweet Dahi (Lychee) 90g   (CBX)
    SetPackSize 51, 1     ' id18 Gaia Prem Sweet Dahi (Muskmelon)    (PCS)
    SetPackSize 54, 60    ' id19 Gaia Kadhi Dahi (Crt) 200 Gm 60 Pkt (CRT)
    SetPackSize 57, 12    ' id20 Gaia Kadhi Dahi 1 Kg Crt (12 Pcs)   (CRT)
    SetPackSize 60, 2     ' id21 Gaia Kadhi Dahi 5kg Pouch (2 Pcs)   (CRT)
    SetPackSize 64, 1     ' id23 Gaia Kadhi Dahi Bucket 15 Kg        (PCS)
    SetPackSize 68, 24    ' id27 Gaia Sweet Lassi 180ml Pouch        (CRT)
    SetPackSize 71, 10    ' id24 Gaia Plain Lassi 180 Ml Glass Cup10 (CBX)
    SetPackSize 74, 10    ' id25 Mango Lassi 180 Ml - 10Pcs          (CBX)
    SetPackSize 77, 10    ' id28 Strawberry Lassi 180 Ml - 10 Pcs    (CBX)

    ' *** ADDED - CONFIRM real pack sizes (pcs per Crt/Box) before using ***
    SetPackSize 65, 24    ' id22 Masala Chaas 200 Grms         (CRT) - GUESS, confirm actual
    SetPackSize 80, 10    ' id26 Masala Chaas Glaas 180 Grms   (CBX) - GUESS, confirm actual

    SetPackSize 83, 1     ' id29 Gaia Paneer 200gms                  (PCS)
    SetPackSize 87, 1     ' id30 Gaia Paneer 500Gms                  (PCS)
    SetPackSize 89, 1     ' id31 Gaia Paneer 1kg                     (PCS)
    SetPackSize 92, 1     ' id32 Loose 1 Kg                          (PCS)
    SetPackSize 93, 1     ' id33 Loose 5 Kg                          (PCS)
    SetPackSize 97, 24    ' id35 Gaia Ghee 200 ML Jar                (CBX)
    SetPackSize 99, 12    ' id36 Gaia Ghee 500 ML Jar                (CBX)
    SetPackSize 101, 12   ' id37 Gaia Ghee 1 Ltr Jar 18              (CBX)
    SetPackSize 103, 12   ' id38 Gaia Ghee Ceka Pack 1ltr            (CBX)
    SetPackSize 105, 12   ' id45 Gaia Cow Ghee Ceka Pack 1ltr        (CBX)
    SetPackSize 107, 12   ' id39 Gaia Prem Desi Ghee Ceka Pack 900ml (CBX)
    SetPackSize 109, 12   ' id44 Gaia Pure Cow Ghee Ceka Pack 900ml  (CBX)
    SetPackSize 111, 4    ' id40 Gaia Ghee 5 Ltr.Jar                 (CBX)
    SetPackSize 113, 24   ' id41 Gaia Cow Ghee 200 Ml Jar            (CBX)
    SetPackSize 115, 12   ' id42 Gaia Cow Ghee 500 Ml Jar            (CBX)
    SetPackSize 117, 12   ' id43 Gaia Cow Ghee 1Ltr Jar              (CBX)
    SetPackSize 119, 100  ' id34 Gaia Premium Desi Ghee 20ml         (CBX)
    SetPackSize 122, 1    ' id46 Gaia Ghee 15 Kilogram               (PCS)
    SetPackSize 123, 12   ' id48 Gaia Shrikhand KE 80g               (CBX)
    SetPackSize 126, 6    ' id49 Gaia Shrikhand KE 80g-Box (6Pc)     (CBX)
    SetPackSize 128, 12   ' id47 Gaia Shahi Rabdi 80g (1cup*12pcs)   (CBX)
    SetPackSize 130, 6    ' id47 Gaia Shahi Rabdi 80g (1cup*6pcs)    (CBX)
    SetPackSize 132, 6    ' id47 Gaia Shahi Rabdi 80g-Box(6Pc)       (CBX)
    SetPackSize 134, 15   ' id50 Gaia Peda 200 Gm Pcs                (CBX)
    SetPackSize 136, 20   ' id51 Gaia Kesar Peda 200 Gm Pcs          (CBX)
    SetPackSize 138, 1    ' id53 Khowa Brown (Unsweetened)           (KG)
End Sub

Private Sub SetPackSize(ByVal col As Long, ByVal pcsPerUnit As Double)
    If pcsPerUnit > 0 Then mPackSize(col) = pcsPerUnit
End Sub

Private Sub ValidatePackSizesAgainstAPI()
    If Not gAPIStockLoaded Or gAPIPcsPerCrt Is Nothing Then Exit Sub

    Dim col As Long, apiVal As Double, msg As String, mismatchCount As Long
    msg = ""
    mismatchCount = 0

    For col = 1 To 200
        If Len(mProdId(col)) > 0 And Len(mItem(col)) > 0 Then
            If gAPIPcsPerCrt.Exists(mProdId(col)) Then
                apiVal = gAPIPcsPerCrt(mProdId(col))
                If apiVal > 0 And Abs(apiVal - mPackSize(col)) > 0.0001 Then
                    mismatchCount = mismatchCount + 1
                    If mismatchCount <= 25 Then
                        msg = msg & "- " & mItem(col) & ": macro=" & CStr(mPackSize(col)) & _
                              " pcs/unit, API=" & CStr(apiVal) & " pcs/unit" & vbCrLf
                    End If
                End If
            End If
        End If
    Next col

    If mismatchCount > 0 Then
        MsgBox mismatchCount & " item(s) ka pack size macro aur live API mein match nahi kar raha:" & vbCrLf & vbCrLf & _
               msg & vbCrLf & "Stock allocation live API values ke anusar chalega.", _
               vbExclamation, "Pack Size Sync"
    End If
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Sub ExportDayBookToTallyXML()
    Dim wb As Workbook, ws As Worksheet, rws As Worksheet
    Dim r As Long, c As Long, lastRow As Long
    Dim srno As Variant, outlet As String, party As String
    Dim curStart As Long, startOf() As Long
    Dim nL As Long, i As Long, vno As Long
    Dim vchDate As Date, qv As Double, rv As Double, outQty As Double, outRate As Double
    Dim filePath As Variant, stm As Object
    Dim consName As String, consAddress As String, consState As String
    Dim consPin As String, consGST As String, buyerAddress As String
    Dim buyerGST As String, buyerPin As String, buyerState As String

    Dim spLines As Object, spHeader As Object, spOrder As Collection
    Dim spKey As Variant, lineDict As Object, lk As Variant
    Dim arr As Variant, hdr As Variant, spBuyerState As String
    Dim spCat As String, spFullKey As String, spCatSeq As Variant

    Set wb = ActiveWorkbook
    LoadMapping
    LoadPackSizes wb
    LoadPartyStates wb
    LoadDistributors wb, WarnIfMissing:=True
    LoadShelfLife
    LoadBatchMap wb
    LoadStockFromAPI
    ValidatePackSizesAgainstAPI
    gMFSShortfallMsg = ""
    gMFSLineCount = 0
    gOtherLineCount = 0

    Set spLines = CreateObject("Scripting.Dictionary")
    Set spHeader = CreateObject("Scripting.Dictionary")
    Set spOrder = New Collection

    vchDate = DateValue(GetVchDate(wb))

    filePath = Application.GetSaveAsFilename(InitialFileName:="DayBook_Tally_Import.xml", FileFilter:="XML Files (*.xml), *.xml", Title:="Save Tally Import XML As")
    If filePath = False Then Exit Sub

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open

    stm.WriteText "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    stm.WriteText "<ENVELOPE>" & vbCrLf
    stm.WriteText "<HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>" & vbCrLf
    stm.WriteText "<BODY><IMPORTDATA>" & vbCrLf
    stm.WriteText "<REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME><STATICVARIABLES><SVCURRENTCOMPANY>" & XmlEsc(SV_COMPANY) & "</SVCURRENTCOMPANY></STATICVARIABLES></REQUESTDESC>" & vbCrLf
    stm.WriteText "<REQUESTDATA>" & vbCrLf

    WriteBuyerLedgerMasters stm, wb
    Application.ScreenUpdating = False
    vno = STARTING_VCH_NO - 1

    For Each ws In wb.Worksheets
        If ws.Name = OUT_SHEET Then GoTo NextSheet
        If Not SheetExists(wb, ws.Name & " R") Then GoTo NextSheet
        If ws.Name Like "* R" Then GoTo NextSheet

        Set rws = wb.Worksheets(ws.Name & " R")

        ReDim startOf(COL_FIRST To COL_LAST)
        curStart = 0

        ' --- SAME FIX applied here as in BuildDayBookFormat: fall back to
        ' row 2 when row 3 is blank, so Masala Chaas columns start their own
        ' block instead of inheriting Strawberry Lassi's. ---
        For c = COL_FIRST To COL_LAST
            If Len(Trim$(CStr(ws.Cells(3, c).Value))) > 0 Then
                curStart = c
            ElseIf Len(Trim$(CStr(ws.Cells(2, c).Value))) > 0 Then
                curStart = c
            End If
            startOf(c) = curStart
        Next c

        lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row

        For r = 6 To lastRow
            srno = ws.Cells(r, 2).Value
            outlet = Trim$(CStr(ws.Cells(r, 3).Value))

            If Len(outlet) = 0 Then GoTo NextRow
            If Not IsValidSrNo(srno) Then GoTo NextRow
            If IsExcludedOutlet(outlet) Then GoTo NextRow

            party = LookupParty(outlet)
            If Len(party) = 0 Then party = outlet

            GetPartyDetails wb, outlet, consName, consAddress, consState, consPin, consGST, buyerAddress, buyerGST, buyerPin

            buyerState = PartyState(party)
            If Len(buyerState) = 0 Then buyerState = SV_STATE

            Dim isMFSLocal2 As Boolean
            isMFSLocal2 = IsMFSOutlet(outlet)

            nL = 0

            For c = COL_FIRST To COL_LAST
                qv = val(ws.Cells(r, c).Value)
                If qv = 0 Then GoTo NextCol
                If startOf(c) = 0 Then GoTo NextCol
                If Len(mItem(startOf(c))) = 0 Then GoTo NextCol

                rv = val(rws.Cells(r, c).Value)
                If rv = 0 Then GoTo NextCol
                If InStr(LCase$(Trim$(CStr(ws.Cells(4, c).Value))), "free") > 0 Then GoTo NextCol
                If nL >= MAXL Then GoTo NextCol

                Dim pidLocal2 As String, itemNameLocal2 As String, ledgerNameLocal2 As String, unitNameLocal2 As String, gstRateLocal2 As Double
                pidLocal2 = mProdId(startOf(c))
                itemNameLocal2 = mItem(startOf(c))
                ledgerNameLocal2 = mLedger(startOf(c))
                If Len(ledgerNameLocal2) = 0 Then ledgerNameLocal2 = "Sales"
                unitNameLocal2 = mUnit(startOf(c))
                gstRateLocal2 = GstForLedger(ledgerNameLocal2)

                Call ComputeLineQtyAndRate(c, startOf(c), qv, rv, rws, r, gstRateLocal2, outQty, outRate)

                Dim pcsPerUnitLocal2 As Double, pcsNeededLocal2 As Double
                pcsPerUnitLocal2 = GetPcsPerUnit(startOf(c))
                pcsNeededLocal2 = outQty * pcsPerUnitLocal2

                Dim aB2() As String, aQPcs2() As Double, aC2 As Long, allocIndex2 As Long
                Call AllocateBatchesForItem(pidLocal2, itemNameLocal2, pcsNeededLocal2, mDays(startOf(c)), vchDate, isMFSLocal2, aB2, aQPcs2, aC2)

                For allocIndex2 = 1 To aC2
                    If nL >= MAXL Then Exit For
                    nL = nL + 1
                    bProdId(nL) = pidLocal2
                    bItem(nL) = itemNameLocal2
                    bLedger(nL) = ledgerNameLocal2
                    bUnit(nL) = unitNameLocal2
                    bGst(nL) = gstRateLocal2
                    bQtyPcs(nL) = aQPcs2(allocIndex2)

                    If pcsPerUnitLocal2 > 0 Then
                        bQty(nL) = aQPcs2(allocIndex2) / pcsPerUnitLocal2
                    Else
                        bQty(nL) = aQPcs2(allocIndex2)
                    End If

                    bRate(nL) = outRate
                    bAmt(nL) = Round(bQty(nL) * outRate, 2)

                    ' ---- USE USER'S MANUAL BATCH FROM DAY BOOK IF PRESENT ----
                    Dim mappedBatch As String
                    mappedBatch = ""
                    On Error Resume Next
                    mappedBatch = colBatchMap(LCase$(party & "|" & itemNameLocal2))
                    On Error GoTo 0

                    If Len(mappedBatch) > 0 Then
                        bBatch(nL) = mappedBatch
                    Else
                        bBatch(nL) = aB2(allocIndex2)
                    End If
                    ' -----------------------------------------------------------

                    bShelf(nL) = mDays(startOf(c))
                Next allocIndex2

NextCol:
            Next c

            If nL = 0 Then GoTo NextRow

            If IsSuperStockistParty(party) Then
                If Not spHeader.Exists(party) Then
                    spHeader(party) = Array(buyerAddress, buyerGST, buyerPin)
                    spOrder.Add party
                End If

                For i = 1 To nL
                    spCat = IIf(IsMilkLedger(bLedger(i)), "MILK", "OTHER")
                    spFullKey = party & Chr(1) & spCat

                    If Not spLines.Exists(spFullKey) Then
                        Set spLines(spFullKey) = CreateObject("Scripting.Dictionary")
                    End If

                    Set lineDict = spLines(spFullKey)
                    lk = bItem(i) & "|" & Format(Round(bRate(i), 2), "0.00")

                    If lineDict.Exists(lk) Then
                        arr = lineDict(lk)
                        arr(4) = arr(4) + bQty(i)
                        arr(5) = Round(arr(5) + bAmt(i), 2)
                        arr(10) = arr(10) + bQtyPcs(i)
                        lineDict(lk) = arr
                    Else
                        lineDict.Add lk, Array(bItem(i), bLedger(i), bUnit(i), bRate(i), bQty(i), bAmt(i), bGst(i), bBatch(i), bShelf(i), bProdId(i), bQtyPcs(i))
                    End If
                Next i

                GoTo NextRow
            End If

            vno = vno + 1
            Call WriteXMLVoucher(stm, vno, vchDate, party, nL, buyerState, buyerGST, buyerPin, buyerAddress, consName, consAddress, consState, consPin, consGST, False)

NextRow:
        Next r
NextSheet:
    Next ws

    For Each spKey In spOrder
        For Each spCatSeq In Array("MILK", "OTHER")
            spFullKey = CStr(spKey) & Chr(1) & CStr(spCatSeq)

            If spLines.Exists(spFullKey) Then
                Set lineDict = spLines(spFullKey)
                nL = 0

                For Each lk In lineDict.keys
                    nL = nL + 1
                    arr = lineDict(lk)
                    bItem(nL) = arr(0)
                    bLedger(nL) = arr(1)
                    bUnit(nL) = arr(2)
                    bRate(nL) = arr(3)
                    bQty(nL) = arr(4)
                    bAmt(nL) = Round(arr(5), 2)
                    bGst(nL) = arr(6)
                    bBatch(nL) = arr(7)
                    bShelf(nL) = arr(8)
                    bProdId(nL) = arr(9)
                    bQtyPcs(nL) = arr(10)
                Next lk

                If nL > 0 Then
                    vno = vno + 1
                    hdr = spHeader(CStr(spKey))
                    spBuyerState = PartyState(CStr(spKey))
                    If Len(spBuyerState) = 0 Then spBuyerState = SV_STATE
                    Call WriteXMLVoucher(stm, vno, vchDate, CStr(spKey), nL, spBuyerState, CStr(hdr(1)), CStr(hdr(2)), CStr(hdr(0)), "", "", "", "", "", True)
                End If
            End If
        Next spCatSeq
    Next spKey

    stm.WriteText "</REQUESTDATA>" & vbCrLf
    stm.WriteText "</IMPORTDATA></BODY>" & vbCrLf
    stm.WriteText "</ENVELOPE>"

    stm.SaveToFile CStr(filePath), 2
    stm.Close

    Application.ScreenUpdating = True

    Call ShowMFSShortfallWarningIfAny
    Call ShowAllocationSummary

    MsgBox "Ho gaya!" & vbCrLf & (vno - STARTING_VCH_NO + 1) & " vouchers XML mein likhe gaye." & vbCrLf & "File: " & filePath, vbInformation, "Demand -> Tally XML"
End Sub

Private Sub WriteBuyerLedgerMasters(ByVal stm As Object, ByVal wb As Workbook)
    Dim ws As Worksheet, r As Long, lastRow As Long, colOutlet As Long, colParty As Long
    Dim outlet As String, party As String, masterName As String
    Dim consName As String, consAddress As String, consState As String
    Dim consPin As String, consGST As String, buyerAddress As String
    Dim buyerGST As String, buyerPin As String, seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = wb.Worksheets("Distributor Master")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    colOutlet = FindHeaderColumn(ws, 4, "Distributor")
    colParty = FindHeaderColumn(ws, 4, "Distributor in Tally")
    If colOutlet = 0 Then colOutlet = 2
    If colParty = 0 Then colParty = 4

    lastRow = ws.Cells(ws.Rows.Count, colOutlet).End(xlUp).row

    For r = 5 To lastRow
        outlet = Trim$(CStr(ws.Cells(r, colOutlet).Value))
        party = Trim$(CStr(ws.Cells(r, colParty).Value))

        If Len(outlet) > 0 And Len(party) > 0 Then
            masterName = LedgerMasterName(party)

            If Len(masterName) > 0 Then
                If Not seen.Exists(LCase$(masterName)) Then
                    GetPartyDetails wb, outlet, consName, consAddress, consState, consPin, consGST, buyerAddress, buyerGST, buyerPin
                    If Len(Trim$(buyerAddress)) > 0 Then
                        WriteXMLBuyerLedger stm, masterName, buyerAddress, PartyState(masterName), buyerPin
                    End If
                    seen.Add LCase$(masterName), True
                End If
            End If
        End If
    Next r
End Sub

Private Sub WriteXMLBuyerLedger(ByVal stm As Object, ByVal party As String, ByVal buyerAddress As String, ByVal buyerState As String, ByVal buyerPin As String)
    Dim addrLine As Variant, mailingName As String
    mailingName = party
    If Len(Trim$(buyerState)) = 0 Then buyerState = SV_STATE

    stm.WriteText "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
    stm.WriteText "<LEDGER NAME=""" & XmlEsc(party) & """ ACTION=""Alter"">" & vbCrLf
    stm.WriteText "<NAME>" & XmlEsc(party) & "</NAME>" & vbCrLf
    stm.WriteText "<MAILINGNAME.LIST TYPE=""String"">" & vbCrLf
    stm.WriteText "<MAILINGNAME>" & XmlEsc(mailingName) & "</MAILINGNAME>" & vbCrLf
    stm.WriteText "</MAILINGNAME.LIST>" & vbCrLf
    stm.WriteText "<ADDRESS.LIST TYPE=""String"">" & vbCrLf
    For Each addrLine In Split(Replace(buyerAddress, vbCrLf, vbLf), vbLf)
        If Len(Trim$(CStr(addrLine))) > 0 Then
            stm.WriteText "<ADDRESS>" & XmlEsc(CStr(addrLine)) & "</ADDRESS>" & vbCrLf
        End If
    Next addrLine
    stm.WriteText "</ADDRESS.LIST>" & vbCrLf
    stm.WriteText "<COUNTRYNAME>India</COUNTRYNAME>" & vbCrLf
    stm.WriteText "<STATENAME>" & XmlEsc(buyerState) & "</STATENAME>" & vbCrLf
    If Len(Trim$(buyerPin)) > 0 Then stm.WriteText "<PINCODE>" & XmlEsc(buyerPin) & "</PINCODE>" & vbCrLf
    stm.WriteText "</LEDGER>" & vbCrLf
    stm.WriteText "</TALLYMESSAGE>" & vbCrLf
End Sub

Private Sub WriteXMLVoucher(ByVal stm As Object, ByVal vno As Long, ByVal vchDate As Date, ByVal party As String, ByVal nL As Long, ByVal buyerState As String, ByVal buyerGST As String, ByVal buyerPin As String, ByVal buyerAddress As String, ByVal consName As String, ByVal consAddress As String, ByVal consState As String, ByVal consPin As String, ByVal consGST As String, ByVal skipConsignee As Boolean)
    Dim i As Long
    Dim taxable As Double, cgst As Double, sgst As Double
    Dim gross As Double, billtotal As Double, roundoff As Double, total As Double
    Dim expDate As Date, addrLine As Variant

    taxable = 0: cgst = 0
    For i = 1 To nL
        taxable = taxable + bAmt(i)
        cgst = cgst + bAmt(i) * bGst(i) / 200
    Next i

    cgst = Round(cgst, 2): sgst = cgst
    gross = taxable + cgst + sgst
    billtotal = Round(gross, 0)
    roundoff = Round(billtotal - gross, 2)
    total = billtotal

    stm.WriteText "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
    stm.WriteText "<VOUCHER VCHTYPE=""Sales"" ACTION=""Create"" OBJVIEW=""Invoice Voucher View"">" & vbCrLf
    stm.WriteText "<DATE>" & Format(vchDate, "yyyymmdd") & "</DATE>" & vbCrLf
    stm.WriteText "<EFFECTIVEDATE>" & Format(vchDate, "yyyymmdd") & "</EFFECTIVEDATE>" & vbCrLf
    stm.WriteText "<VOUCHERNUMBER>" & XmlEsc(CStr(vno)) & "</VOUCHERNUMBER>" & vbCrLf
    stm.WriteText "<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>" & vbCrLf
    stm.WriteText "<PARTYLEDGERNAME>" & XmlEsc(party) & "</PARTYLEDGERNAME>" & vbCrLf
    stm.WriteText "<PARTYNAME>" & XmlEsc(party) & "</PARTYNAME>" & vbCrLf
    stm.WriteText "<BASICBASEPARTYNAME>" & XmlEsc(party) & "</BASICBASEPARTYNAME>" & vbCrLf
    stm.WriteText "<BASICBUYERNAME>" & XmlEsc(party) & "</BASICBUYERNAME>" & vbCrLf
    stm.WriteText "<PARTYMAILINGNAME>" & XmlEsc(party) & "</PARTYMAILINGNAME>" & vbCrLf
    stm.WriteText "<STATENAME>" & XmlEsc(buyerState) & "</STATENAME>" & vbCrLf
    stm.WriteText "<PLACEOFSUPPLY>" & XmlEsc(buyerState) & "</PLACEOFSUPPLY>" & vbCrLf

    If Len(Trim$(buyerGST)) > 0 Then stm.WriteText "<PARTYGSTIN>" & XmlEsc(buyerGST) & "</PARTYGSTIN>" & vbCrLf
    If Len(Trim$(buyerPin)) > 0 Then stm.WriteText "<PARTYPINCODE>" & XmlEsc(buyerPin) & "</PARTYPINCODE>" & vbCrLf

    If Len(Trim$(buyerAddress)) > 0 Then
        stm.WriteText "<BASICBUYERADDRESS.LIST TYPE=""String"">" & vbCrLf
        For Each addrLine In Split(Replace(buyerAddress, vbCrLf, vbLf), vbLf)
            If Len(Trim$(CStr(addrLine))) > 0 Then
                stm.WriteText "<BASICBUYERADDRESS>" & XmlEsc(CStr(addrLine)) & "</BASICBUYERADDRESS>" & vbCrLf
            End If
        Next addrLine
        stm.WriteText "</BASICBUYERADDRESS.LIST>" & vbCrLf
    End If

    If Not skipConsignee Then
        If Len(Trim$(consName)) > 0 Then stm.WriteText "<CONSIGNEEMAILINGNAME>" & XmlEsc(consName) & "</CONSIGNEEMAILINGNAME>" & vbCrLf
        If Len(Trim$(consState)) > 0 Then stm.WriteText "<CONSIGNEESTATENAME>" & XmlEsc(consState) & "</CONSIGNEESTATENAME>" & vbCrLf
        If Len(Trim$(consPin)) > 0 Then stm.WriteText "<CONSIGNEEPINCODE>" & XmlEsc(consPin) & "</CONSIGNEEPINCODE>" & vbCrLf
        If Len(Trim$(consGST)) > 0 Then stm.WriteText "<CONSIGNEEGSTIN>" & XmlEsc(consGST) & "</CONSIGNEEGSTIN>" & vbCrLf

        If Len(Trim$(consAddress)) > 0 Then
            stm.WriteText "<ADDRESS.LIST TYPE=""String"">" & vbCrLf
            For Each addrLine In Split(Replace(consAddress, vbCrLf, vbLf), vbLf)
                If Len(Trim$(CStr(addrLine))) > 0 Then
                    stm.WriteText "<ADDRESS>" & XmlEsc(CStr(addrLine)) & "</ADDRESS>" & vbCrLf
                End If
            Next addrLine
            stm.WriteText "</ADDRESS.LIST>" & vbCrLf
        End If
    End If

    stm.WriteText "<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>" & vbCrLf
    stm.WriteText "<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>" & vbCrLf
    stm.WriteText "<ISINVOICE>Yes</ISINVOICE>" & vbCrLf

    For i = 1 To nL
        expDate = vchDate + bShelf(i)
        stm.WriteText "<ALLINVENTORYENTRIES.LIST>" & vbCrLf
        stm.WriteText "<STOCKITEMNAME>" & XmlEsc(bItem(i)) & "</STOCKITEMNAME>" & vbCrLf
        stm.WriteText "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf
        stm.WriteText "<RATE>" & FmtMoney(bRate(i)) & "/" & XmlEsc(bUnit(i)) & "</RATE>" & vbCrLf
        stm.WriteText "<AMOUNT>" & FmtMoney(bAmt(i)) & "</AMOUNT>" & vbCrLf
        stm.WriteText "<ACTUALQTY> " & FmtQty(bQty(i)) & " " & XmlEsc(bUnit(i)) & "</ACTUALQTY>" & vbCrLf
        stm.WriteText "<BILLEDQTY> " & FmtQty(bQty(i)) & " " & XmlEsc(bUnit(i)) & "</BILLEDQTY>" & vbCrLf
        stm.WriteText "<BATCHALLOCATIONS.LIST>" & vbCrLf
        stm.WriteText "<GODOWNNAME>Plant</GODOWNNAME>" & vbCrLf
        stm.WriteText "<BATCHNAME>" & XmlEsc(bBatch(i)) & "</BATCHNAME>" & vbCrLf
        stm.WriteText "<EXPIRYDATE>" & Format(expDate, "yyyymmdd") & "</EXPIRYDATE>" & vbCrLf
        stm.WriteText "<AMOUNT>" & FmtMoney(bAmt(i)) & "</AMOUNT>" & vbCrLf
        stm.WriteText "<ACTUALQTY> " & FmtQty(bQty(i)) & " " & XmlEsc(bUnit(i)) & "</ACTUALQTY>" & vbCrLf
        stm.WriteText "<BILLEDQTY> " & FmtQty(bQty(i)) & " " & XmlEsc(bUnit(i)) & "</BILLEDQTY>" & vbCrLf
        stm.WriteText "</BATCHALLOCATIONS.LIST>" & vbCrLf
        stm.WriteText "<ACCOUNTINGALLOCATIONS.LIST>" & vbCrLf
        stm.WriteText "<LEDGERNAME>" & XmlEsc(bLedger(i)) & "</LEDGERNAME>" & vbCrLf
        stm.WriteText "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf
        stm.WriteText "<AMOUNT>" & FmtMoney(bAmt(i)) & "</AMOUNT>" & vbCrLf
        stm.WriteText "</ACCOUNTINGALLOCATIONS.LIST>" & vbCrLf
        stm.WriteText "</ALLINVENTORYENTRIES.LIST>" & vbCrLf
    Next i

    stm.WriteText "<LEDGERENTRIES.LIST>" & vbCrLf
    stm.WriteText "<LEDGERNAME>" & XmlEsc(party) & "</LEDGERNAME>" & vbCrLf
    stm.WriteText "<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf
    stm.WriteText "<ISPARTYLEDGER>Yes</ISPARTYLEDGER>" & vbCrLf
    stm.WriteText "<AMOUNT>" & FmtMoney(-total) & "</AMOUNT>" & vbCrLf
    stm.WriteText "<BILLALLOCATIONS.LIST>" & vbCrLf
    stm.WriteText "<NAME>" & XmlEsc(CStr(vno)) & "</NAME>" & vbCrLf
    stm.WriteText "<BILLTYPE>New Ref</BILLTYPE>" & vbCrLf
    stm.WriteText "<AMOUNT>" & FmtMoney(-total) & "</AMOUNT>" & vbCrLf
    stm.WriteText "</BILLALLOCATIONS.LIST>" & vbCrLf
    stm.WriteText "</LEDGERENTRIES.LIST>" & vbCrLf

    If cgst > 0 Then
        If IsInterstate(party) Then
            stm.WriteText "<LEDGERENTRIES.LIST>" & vbCrLf
            stm.WriteText "<LEDGERNAME>Output IGST 5%</LEDGERNAME>" & vbCrLf
            stm.WriteText "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf
            stm.WriteText "<ISPARTYLEDGER>No</ISPARTYLEDGER>" & vbCrLf
            stm.WriteText "<AMOUNT>" & FmtMoney(cgst + sgst) & "</AMOUNT>" & vbCrLf
            stm.WriteText "</LEDGERENTRIES.LIST>" & vbCrLf
        Else
            stm.WriteText "<LEDGERENTRIES.LIST>" & vbCrLf
            stm.WriteText "<LEDGERNAME>Out Put CGST 2.5%</LEDGERNAME>" & vbCrLf
            stm.WriteText "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf
            stm.WriteText "<ISPARTYLEDGER>No</ISPARTYLEDGER>" & vbCrLf
            stm.WriteText "<AMOUNT>" & FmtMoney(cgst) & "</AMOUNT>" & vbCrLf
            stm.WriteText "</LEDGERENTRIES.LIST>" & vbCrLf
            stm.WriteText "<LEDGERENTRIES.LIST>" & vbCrLf
            stm.WriteText "<LEDGERNAME>Out Put SGST 2.5%</LEDGERNAME>" & vbCrLf
            stm.WriteText "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf
            stm.WriteText "<ISPARTYLEDGER>No</ISPARTYLEDGER>" & vbCrLf
            stm.WriteText "<AMOUNT>" & FmtMoney(sgst) & "</AMOUNT>" & vbCrLf
            stm.WriteText "</LEDGERENTRIES.LIST>" & vbCrLf
        End If
    End If

    If Abs(roundoff) >= 0.005 Then
        stm.WriteText "<LEDGERENTRIES.LIST>" & vbCrLf
        stm.WriteText "<LEDGERNAME>Roundoff</LEDGERNAME>" & vbCrLf
        stm.WriteText "<ISPARTYLEDGER>No</ISPARTYLEDGER>" & vbCrLf
        If roundoff < 0 Then
            stm.WriteText "<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf
        Else
            stm.WriteText "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf
        End If
        stm.WriteText "<AMOUNT>" & FmtMoney(roundoff) & "</AMOUNT>" & vbCrLf
        stm.WriteText "</LEDGERENTRIES.LIST>" & vbCrLf
    End If

    stm.WriteText "</VOUCHER>" & vbCrLf
    stm.WriteText "</TALLYMESSAGE>" & vbCrLf
End Sub

Private Sub AddDays(ByVal col As Long, ByVal days As Long)
    mDays(col) = days
End Sub

Private Sub LoadShelfLife()
    Dim i As Long: For i = 1 To 200: mDays(i) = 7: Next i
    AddDays 4, 2: AddDays 7, 2
    AddDays 10, 90: AddDays 13, 90: AddDays 16, 90: AddDays 19, 90: AddDays 22, 90
    AddDays 25, 5: AddDays 28, 5: AddDays 31, 5: AddDays 34, 5: AddDays 37, 5
    AddDays 40, 5: AddDays 43, 5: AddDays 46, 5: AddDays 47, 5: AddDays 48, 7
    AddDays 51, 7: AddDays 54, 5: AddDays 57, 5: AddDays 60, 5: AddDays 64, 5
    AddDays 68, 3: AddDays 71, 3: AddDays 74, 3: AddDays 77, 3

    ' *** ADDED - CONFIRM real shelf life (days) before using ***
    AddDays 65, 3   ' id22 Masala Chaas 200 Grms       - GUESS (matched to other Lassi/Chaas items), confirm
    AddDays 80, 3   ' id26 Masala Chaas Glaas 180 Grms - GUESS, confirm

    AddDays 83, 5: AddDays 87, 5: AddDays 89, 5: AddDays 92, 5: AddDays 93, 5
    AddDays 97, 180: AddDays 99, 180: AddDays 101, 180: AddDays 103, 180: AddDays 105, 180
    AddDays 107, 180: AddDays 109, 180: AddDays 111, 180: AddDays 113, 180: AddDays 115, 180
    AddDays 117, 180: AddDays 119, 180: AddDays 122, 180
    AddDays 123, 7: AddDays 126, 7: AddDays 128, 7: AddDays 130, 7: AddDays 132, 7
    AddDays 134, 7: AddDays 136, 7: AddDays 138, 5
End Sub

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;"): s = Replace(s, "<", "&lt;"): s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;"): s = Replace(s, "'", "&apos;")
    XmlEsc = s
End Function

Private Function FmtMoney(ByVal v As Double) As String
    FmtMoney = Format(Round(v, 2), "0.00")
End Function

Private Function FmtQty(ByVal q As Double) As String
    If q = Int(q) Then
        FmtQty = CStr(CLng(q))
    Else
        FmtQty = Format(q, "0.####")
    End If
End Function

' =====================================================================
' 5. POST SALES DIRECTLY TO WEBSITE / API (UNAMBIGUOUS PIECES POSTING)
' =====================================================================

Public Sub UpdateStockFromDayBook()
    Dim wb As Workbook, dbws As Worksheet
    Dim r As Long, lastR As Long
    Dim prodId As String, itemName As String, batchCode As String, qtyPcs As Double
    Dim vchDateStr As String, isoDate As String
    Dim jsonItems As String, itemCount As Long
    Dim http As Object

    Set wb = ActiveWorkbook

    On Error Resume Next
    Set dbws = wb.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If dbws Is Nothing Then
        MsgBox "'" & OUT_SHEET & "' sheet nahi mili.", vbExclamation, "Update Sales"
        Exit Sub
    End If

    If Trim$(CStr(dbws.Cells(1, STOCK_FLAG_COL).Value)) <> "Stock Deducted" Then
        dbws.Cells(1, STOCK_FLAG_COL).Value = "Stock Deducted"
        dbws.Cells(1, STOCK_FLAG_COL).Font.Bold = True
    End If

    vchDateStr = GetVchDate(wb)
    If IsDate(vchDateStr) Then
        isoDate = Format(CDate(vchDateStr), "yyyy-mm-dd")
    Else
        isoDate = Format(Date, "yyyy-mm-dd")
    End If

    lastR = dbws.Cells(dbws.Rows.Count, 2).End(xlUp).row
    jsonItems = ""
    itemCount = 0

    Application.ScreenUpdating = False

    For r = 3 To lastR
        If Left$(CStr(dbws.Cells(r, 2).Value), 2) = "  " Then
            batchCode = Trim$(CStr(dbws.Cells(r, BATCH_COL).Value))
            itemName = Trim$(Mid$(CStr(dbws.Cells(r, 2).Value), 3))
            prodId = Trim$(CStr(dbws.Cells(r, PROD_ID_COL).Value))

            ' Always post exact PIECES (Col 22)
            qtyPcs = val(dbws.Cells(r, PCS_QTY_COL).Value)
            If qtyPcs <= 0 Then qtyPcs = val(dbws.Cells(r, 3).Value)

            If Len(batchCode) > 0 And LCase$(batchCode) <> "primary batch" And qtyPcs > 0 Then
                If Trim$(CStr(dbws.Cells(r, STOCK_FLAG_COL).Value)) <> "Y" Then
                    itemCount = itemCount + 1

                    If Len(jsonItems) > 0 Then jsonItems = jsonItems & ","
                    jsonItems = jsonItems & "{" & _
                        """productId"":""" & prodId & """," & _
                        """productName"":""" & Replace(itemName, """", "\""") & """," & _
                        """date"":""" & isoDate & """," & _
                        """batchNumber"":""" & batchCode & """," & _
                        """salePc"":" & CStr(qtyPcs) & "," & _
                        """saleTotal"":" & CStr(qtyPcs) & _
                        "}"

                    dbws.Cells(r, STOCK_FLAG_COL).Value = "Y"
                End If
            End If
        End If
    Next r

    Application.ScreenUpdating = True

    If itemCount = 0 Then
        MsgBox "Koi nayi sales line nahi mili (sab pehle se posted hain ya Primary Batch hain).", vbInformation, "Post Sales"
        Exit Sub
    End If

    ' Send Live Sales Entry directly to Web App API
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error GoTo PostFail

    http.Open "POST", API_BASE & "/api/metrics", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send "[" & jsonItems & "]"

    If http.Status = 200 Then
        MsgBox itemCount & " items ki Sale entry website par exact Pieces mein update ho gayi hai!" & vbCrLf & vbCrLf & _
               "1) Website ke Sales page (" & isoDate & ") par live record ho gaya." & vbCrLf & _
               "2) Current Stock se exactly minus ho gaya.", vbInformation, "Sales Posted Successfully"
    Else
        MsgBox "Server error: HTTP " & http.Status & vbCrLf & http.responseText, vbExclamation, "Sales Post Error"
    End If
    Exit Sub

PostFail:
    MsgBox "Server se connect nahi ho paya: " & Err.Description, vbExclamation, "Connection Error"
End Sub
