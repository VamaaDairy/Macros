Option Explicit

Private Const FIRST_DATA_ROW As Long = 6
Private Const NAME_COL As Long = 3
Private Const PROD_ROW As Long = 3
Private Const UNIT_ROW As Long = 4

Private gSkippedSku As String
Private gSkippedDist As String
Private gFileTotal As Double
Private gWrittenTotal As Double
Private gSkipTotal As Double
Private gLogWs As Worksheet
Private gLogRow As Long
Private gExtraWs As Worksheet
Private gExtraRow As Long

Public Sub ImportOrdersToDemand()
    Dim path As String
    path = PickOrderFile()
    If Len(path) = 0 Then Exit Sub
    If Dir(path) = "" Then
        MsgBox "File not found:" & vbLf & path, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    gSkippedSku = "": gSkippedDist = ""
    gFileTotal = 0: gWrittenTotal = 0: gSkipTotal = 0

    InitLog
    UnhideManaged
    ClearOldQuantities

    Dim written As Long
    written = FillFromOrders(path)

    FixWeightMeasurementQuiet
    HideEmpties

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True

    Dim bal As String
    If Abs(gWrittenTotal + gSkipTotal - gFileTotal) < 0.001 Then
        bal = "BALANCED (nothing lost)"
    Else
        bal = "NOT BALANCED - " & Format(gFileTotal - gWrittenTotal - gSkipTotal, "0.##") & " units unaccounted!"
    End If

    Dim msg As String
    msg = "Done. Wrote " & written & " quantity cells." & vbLf & vbLf & _
          "CONTROL TOTAL (qty)" & vbLf & _
          "  Order file : " & Format(gFileTotal, "0.##") & vbLf & _
          "  Written    : " & Format(gWrittenTotal, "0.##") & vbLf & _
          "  Skipped    : " & Format(gSkipTotal, "0.##") & vbLf & _
          "  " & bal & vbLf
    If Len(gSkippedSku) > 0 Then msg = msg & vbLf & "SKU codes not mapped (skipped):" & vbLf & gSkippedSku
    If Len(gSkippedDist) > 0 Then msg = msg & vbLf & "Distributors not found:" & vbLf & gSkippedDist
    If gExtraRow > 2 Then _
        msg = msg & vbLf & (gExtraRow - 2) & " order line(s) had no demand column -> see 'Extra Orders' sheet."
    msg = msg & vbLf & "Full details on the 'Import Log' sheet."
    MsgBox msg, vbInformation, "Demand Sheet Import"
End Sub

Public Sub ShowAllRowsCols()
    UnhideManaged
    MsgBox "All distributor rows and product columns are now visible on every city tab.", _
           vbInformation, "Show All"
End Sub

Private Function PickOrderFile() As String
    Dim startFolder As String
    startFolder = GetRememberedFolder()
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(3)
    On Error GoTo 0
    If Not fd Is Nothing Then
        With fd
            .Title = "Select the PRIMARY_ORDER_BOOKING file for today"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Excel files", "*.xlsx; *.xlsm; *.xls"
            If Len(startFolder) > 0 Then .InitialFileName = startFolder
            If .Show = -1 Then
                PickOrderFile = .SelectedItems(1)
                RememberFolder PickOrderFile
            End If
        End With
        Exit Function
    End If
    Dim v As Variant
    v = Application.GetOpenFilename("Excel files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
                                    , "Select today's PRIMARY_ORDER_BOOKING file")
    If VarType(v) = vbBoolean Then Exit Function
    PickOrderFile = CStr(v)
    RememberFolder PickOrderFile
End Function

Private Function GetRememberedFolder() As String
    If SheetExists("Config") Then _
        GetRememberedFolder = Trim$(CStr(ThisWorkbook.Worksheets("Config").Range("B1").Value))
End Function

Private Sub RememberFolder(ByVal fullPath As String)
    Dim ws As Worksheet, folder As String, p As Long
    p = InStrRev(fullPath, "\")
    If p = 0 Then p = InStrRev(fullPath, "/")
    If p > 0 Then folder = Left$(fullPath, p)
    If Not SheetExists("Config") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Config"
        ws.Range("A1").Value = "Last used order-file folder:"
        ws.Range("A1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 28
        ws.Columns("B").ColumnWidth = 80
    End If
    ThisWorkbook.Worksheets("Config").Range("B1").Value = folder
End Sub

Private Sub InitLog()
    If Not SheetExists("Import Log") Then
        Set gLogWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        gLogWs.Name = "Import Log"
    Else
        Set gLogWs = ThisWorkbook.Worksheets("Import Log")
        gLogWs.Cells.Clear
    End If
    gLogWs.Range("A1:G1").Value = Array("Tab / Route", "Distributor", "SKU Code", "SKU Name", "Unit", "Qty", "Reason skipped")
    gLogWs.Range("A1:G1").Font.Bold = True
    gLogRow = 2

    ' clean list of orders that have no demand column (Masala Chaas etc.)
    If Not SheetExists("Extra Orders") Then
        Set gExtraWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        gExtraWs.Name = "Extra Orders"
    Else
        Set gExtraWs = ThisWorkbook.Worksheets("Extra Orders")
        gExtraWs.Cells.Clear
    End If
    gExtraWs.Range("A1:G1").Value = Array("Route / City", "Distributor", "SKU Code", "SKU Name", "Product", "Unit", "Qty")
    gExtraWs.Range("A1:G1").Font.Bold = True
    gExtraRow = 2
End Sub

Private Sub LogExtra(ByVal tabN As String, ByVal dist As String, ByVal code As Variant, _
                     ByVal skuName As String, ByVal prod As String, ByVal unit As String, ByVal qty As Double)
    gExtraWs.Cells(gExtraRow, 1).Value = tabN
    gExtraWs.Cells(gExtraRow, 2).Value = dist
    gExtraWs.Cells(gExtraRow, 3).Value = code
    gExtraWs.Cells(gExtraRow, 4).Value = skuName
    gExtraWs.Cells(gExtraRow, 5).Value = prod
    gExtraWs.Cells(gExtraRow, 6).Value = unit
    gExtraWs.Cells(gExtraRow, 7).Value = qty
    gExtraRow = gExtraRow + 1
End Sub

Private Sub LogSkip(ByVal tabN As String, ByVal dist As String, ByVal code As Variant, _
                    ByVal skuName As String, ByVal unit As String, ByVal qty As Double, ByVal reason As String)
    gLogWs.Cells(gLogRow, 1).Value = tabN
    gLogWs.Cells(gLogRow, 2).Value = dist
    gLogWs.Cells(gLogRow, 3).Value = code
    gLogWs.Cells(gLogRow, 4).Value = skuName
    gLogWs.Cells(gLogRow, 5).Value = unit
    gLogWs.Cells(gLogRow, 6).Value = qty
    gLogWs.Cells(gLogRow, 7).Value = reason
    gLogRow = gLogRow + 1
    gSkipTotal = gSkipTotal + qty
End Sub

Private Function FillFromOrders(ByVal path As String) As Long
    Dim wbO As Workbook, wsO As Worksheet
    Set wbO = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    On Error Resume Next
    Set wsO = wbO.Worksheets("Order Booking")
    On Error GoTo 0
    If wsO Is Nothing Then Set wsO = wbO.Worksheets(1)

    Dim data As Variant
    data = wsO.UsedRange.Value

    Dim cRoute&, cDist&, cSku&, cUnit&, cQty&, cSkuName&, colDate&, cStatus&, j&
    Dim hdr As String
    For j = LBound(data, 2) To UBound(data, 2)
        hdr = Norm(data(LBound(data, 1), j))
        Select Case hdr
            Case "TRUCK ROUTE": cRoute = j
            Case "CUSTOMER NAME": cDist = j
            Case "SKU CODE": cSku = j
            Case "BOOKED UNIT": cUnit = j
            Case "QUANTITY": cQty = j
            Case "SKU": cSkuName = j
            Case "DATE": colDate = j
            Case "STATUS": cStatus = j
        End Select
    Next j
    If cRoute * cDist * cSku * cUnit * cQty = 0 Then
        wbO.Close SaveChanges:=False
        MsgBox "The order file is missing a required column " & _
               "(TRUCK ROUTE, Customer Name, SKU Code, Booked Unit, Quantity).", vbExclamation
        Exit Function
    End If

    Dim orderDate As Variant: orderDate = Empty
    Dim written As Long, i As Long
    Dim routeN$, distN$, unitN$, skuNameN$, prodLabel$, skuNameRaw$
    Dim code&, qty As Double
    Dim ws As Worksheet, targetRow As Long, targetCol As Long

    For i = LBound(data, 1) + 1 To UBound(data, 1)
        If IsEmpty(data(i, cSku)) And IsEmpty(data(i, cDist)) Then GoTo NextRow
        If colDate > 0 And IsEmpty(orderDate) Then orderDate = data(i, colDate)
        qty = val(CStr(data(i, cQty)))
        If qty = 0 Then GoTo NextRow

        gFileTotal = gFileTotal + qty
        code = ParseSkuCode(data(i, cSku))
        skuNameRaw = CStr(data(i, cSkuName))
        routeN = Norm(data(i, cRoute))
        distN = Norm(data(i, cDist))
        unitN = Norm(data(i, cUnit))
        skuNameN = Norm(data(i, cSkuName))

        If cStatus > 0 Then
            If InStr(Norm(data(i, cStatus)), "REJECT") > 0 Then
                LogSkip routeN, distN, code, skuNameRaw, unitN, qty, _
                        "Order status REJECTED - not imported"
                GoTo NextRow
            End If
        End If

        prodLabel = SkuToProduct(code)
        If Len(prodLabel) = 0 Then
            AddOnce gSkippedSku, code & " (" & skuNameRaw & ")"
            LogSkip routeN, distN, code, skuNameRaw, unitN, qty, "SKU code not in mapping list"
            GoTo NextRow
        End If

        Set ws = ResolveSheet(routeN, distN)
        If ws Is Nothing Then
            AddOnce gSkippedDist, distN & " [route " & routeN & "]"
            LogSkip routeN, distN, code, skuNameRaw, unitN, qty, "Distributor not found on any tab"
            GoTo NextRow
        End If

        targetRow = FindDistRow(ws, distN)
        If targetRow = 0 Then
            AddOnce gSkippedDist, distN & " [" & ws.Name & "]"
            LogSkip ws.Name, distN, code, skuNameRaw, unitN, qty, "Distributor row not found on tab"
            GoTo NextRow
        End If

        targetCol = FindColumn(ws, Norm(prodLabel), unitN, skuNameN)
        If targetCol = 0 Then
            ' No demand column for this product (e.g. Masala Chaas) ->
            ' capture it in the clean "Extra Orders" sheet, demand grid untouched
            LogExtra ws.Name, distN, code, skuNameRaw, prodLabel, unitN, qty
            LogSkip ws.Name, distN, code, skuNameRaw, unitN, qty, _
                    "No demand column - moved to 'Extra Orders' sheet"
            GoTo NextRow
        End If

        ws.Cells(targetRow, targetCol).Value = Nz(ws.Cells(targetRow, targetCol).Value) + qty
        gWrittenTotal = gWrittenTotal + qty
        written = written + 1
NextRow:
    Next i

    If Not IsEmpty(orderDate) Then
        Dim s As Variant
        For Each s In CitySheets()
            If SheetExists(CStr(s)) Then ThisWorkbook.Worksheets(CStr(s)).Cells(3, 2).Value = orderDate
        Next s
    End If

    wbO.Close SaveChanges:=False
    FillFromOrders = written
End Function

Private Sub ClearOldQuantities()
    Dim s As Variant, ws As Worksheet, row As Long, col As Long
    Dim lastRow As Long, lastMeasureCol As Long, distEnd As Long, rr As Long
    For Each s In CitySheets()
        If Not SheetExists(CStr(s)) Then GoTo NextSheet
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        lastRow = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).row
        lastMeasureCol = MeasureLastCol(ws)
        If lastMeasureCol < 4 Then GoTo NextSheet

        distEnd = lastRow
        For rr = FIRST_DATA_ROW To lastRow
            If InStr(Norm(ws.Cells(rr, 2).Value), "TOTAL") > 0 _
               Or InStr(Norm(ws.Cells(rr, 3).Value), "TOTAL") > 0 Then
                distEnd = rr - 1: Exit For
            End If
        Next rr

        For row = FIRST_DATA_ROW To distEnd
            Dim nm As String: nm = Norm(ws.Cells(row, NAME_COL).Value)
            If Len(nm) = 0 Then GoTo NextRow
            If IsReserved(nm) Then GoTo NextRow
            For col = 4 To lastMeasureCol
                Dim cell As Range: Set cell = ws.Cells(row, col)
                If Not cell.HasFormula Then
                    If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then cell.ClearContents
                End If
            Next col
NextRow:
        Next row
NextSheet:
    Next s
End Sub

Private Sub UnhideManaged()
    Dim s As Variant, ws As Worksheet, lastRow As Long, lastMeasureCol As Long
    For Each s In CitySheets()
        If Not SheetExists(CStr(s)) Then GoTo NextSheet
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        lastRow = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).row
        lastMeasureCol = MeasureLastCol(ws)
        If lastRow >= FIRST_DATA_ROW Then ws.Range(ws.Rows(FIRST_DATA_ROW), ws.Rows(lastRow)).EntireRow.Hidden = False
        If lastMeasureCol >= 4 Then ws.Range(ws.Columns(4), ws.Columns(lastMeasureCol)).EntireColumn.Hidden = False
NextSheet:
    Next s
End Sub

Private Sub HideEmpties()
    Dim s As Variant, ws As Worksheet
    Dim lastRow As Long, lastMeasureCol As Long, nRows As Long, nCols As Long
    Dim r As Long, c As Long, distEndRow As Long, rr As Long
    For Each s In CitySheets()
        If Not SheetExists(CStr(s)) Then GoTo NextSheet
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        lastRow = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).row
        lastMeasureCol = MeasureLastCol(ws)
        If lastRow < FIRST_DATA_ROW Or lastMeasureCol < 4 Then GoTo NextSheet

        distEndRow = lastRow
        For rr = FIRST_DATA_ROW To lastRow
            If InStr(Norm(ws.Cells(rr, 2).Value), "TOTAL") > 0 _
               Or InStr(Norm(ws.Cells(rr, 3).Value), "TOTAL") > 0 Then
                distEndRow = rr - 1: Exit For
            End If
        Next rr

        nRows = lastRow - FIRST_DATA_ROW + 1
        nCols = lastMeasureCol - 4 + 1
        Dim vals As Variant, names As Variant
        vals = ws.Range(ws.Cells(FIRST_DATA_ROW, 4), ws.Cells(lastRow, lastMeasureCol)).Value
        names = ws.Range(ws.Cells(FIRST_DATA_ROW, NAME_COL), ws.Cells(lastRow, NAME_COL)).Value
        If nRows = 1 Or nCols = 1 Then GoTo NextSheet

        For r = 1 To nRows
            Dim thisRow As Long: thisRow = FIRST_DATA_ROW + r - 1
            Dim nm As String: nm = Norm(names(r, 1))
            If thisRow <= distEndRow And Not IsReserved(nm) Then
                Dim rowHas As Boolean: rowHas = False
                For c = 1 To nCols
                    If IsNumeric(vals(r, c)) Then
                        If vals(r, c) <> 0 Then rowHas = True: Exit For
                    End If
                Next c
                ws.Rows(thisRow).Hidden = Not rowHas
            End If
        Next r

        For c = 1 To nCols
            Dim colHas As Boolean: colHas = False
            For r = 1 To nRows
                Dim nm2 As String: nm2 = Norm(names(r, 1))
                If Len(nm2) > 0 And Not IsReserved(nm2) Then
                    If IsNumeric(vals(r, c)) Then
                        If vals(r, c) <> 0 Then colHas = True: Exit For
                    End If
                End If
            Next r
            ws.Columns(4 + c - 1).Hidden = Not colHas
        Next c
NextSheet:
    Next s
End Sub

Public Sub FixWeightMeasurement()
    Dim filled As Long, sheets As Long
    FixWM filled, sheets
    Application.Calculate
    MsgBox "Weight Measurement checked on " & sheets & " sheets." & vbLf & _
           "Blank/zero cells filled: " & filled & vbLf & vbLf & _
           "Weight row auto-calculates (= Total Qty x Weight Measurement).", _
           vbInformation, "Weight Measurement"
End Sub

Private Sub FixWeightMeasurementQuiet()
    Dim f As Long, sh As Long
    FixWM f, sh
End Sub

Private Sub FixWM(ByRef filled As Long, ByRef sheetCount As Long)
    Dim std As Object, ovr As Object
    Set std = ParseWM(WMStandardData())
    Set ovr = ParseWM(WMOverrideData())

    Dim s As Variant, ws As Worksheet
    Dim wmRow As Long, lastCol As Long, c As Long
    For Each s In CitySheets()
        If Not SheetExists(CStr(s)) Then GoTo NextSheet
        Set ws = ThisWorkbook.Worksheets(CStr(s))
        wmRow = FindLabelRow(ws, "WEIGHT MEASUREMENT")
        If wmRow = 0 Then GoTo NextSheet
        lastCol = MeasureLastCol(ws)
        If lastCol < 4 Then GoTo NextSheet

        Dim useOvr As Boolean
        useOvr = (UCase$(ws.Name) = "NAGPUR" Or UCase$(ws.Name) = "JAGDALPUR")

        Dim curProd As String, plw As String
        curProd = ""
        For c = 1 To lastCol
            plw = ProdLabelAt(ws, c)
            If Len(plw) > 0 Then curProd = plw
            If c < 4 Then GoTo NextCol
            Dim u As String: u = Norm(ws.Cells(UNIT_ROW, c).Value)
            If Len(u) = 0 Or Len(curProd) = 0 Then GoTo NextCol

            Dim key As String: key = curProd & "|" & u
            Dim val As Variant: val = Empty
            If useOvr Then
                If ovr.Exists(key) Then val = ovr(key)
            End If
            If IsEmpty(val) Then
                If std.Exists(key) Then val = std(key)
            End If
            If IsEmpty(val) Then GoTo NextCol

            Dim cell As Range: Set cell = ws.Cells(wmRow, c)
            Dim isBlank As Boolean
            isBlank = IsEmpty(cell.Value)
            If Not isBlank Then
                If IsNumeric(cell.Value) Then isBlank = (CDbl(cell.Value) = 0)
            End If
            If isBlank Then
                cell.Value = CDbl(val)
                filled = filled + 1
            End If
NextCol:
        Next c
        sheetCount = sheetCount + 1
NextSheet:
    Next s
End Sub

Private Function ParseWM(ByVal data As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim parts() As String, i As Long, p As Long, k As String
    parts = Split(data, ";")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            p = InStrRev(parts(i), "=")
            If p > 0 Then
                k = UCase$(Trim$(Left$(parts(i), p - 1)))
                If Not d.Exists(k) Then d.Add k, CDbl(Mid$(parts(i), p + 1))
            End If
        End If
    Next i
    Set ParseWM = d
End Function

Private Function WMStandardData() As String
    Dim s As String
    s = s & "TM 500 ML|CRT=12;TM 500 ML|PKT=0.5;TM 500 ML|FREE PKT=0.5;CM 500 ML|CRT=12;CM 500 ML|PKT=0.5;CM 500 ML|FREE PKT=0.5;MAXX UHT 110 ML|CBX=3.6;MAXX UHT 110 ML|PC=0.12;MAXX UHT 110 ML|FREE PKT=0.12;GAIA MAXX UHT 400ML|CBX=9.6;GAIA MAXX UHT 400ML|PC=0.4;GAIA MAXX UHT 400ML|FREE PKT=0.4;GAIA PREMIUM MILK UHT TETRA PACK 1 LTR|CBX=12;GAIA PREMIUM MILK UHT TETRA PACK 1 LTR|PKT=1;"
    s = s & "GAIA PREMIUM MILK UHT TETRA PACK 1 LTR|FREE PKT=1;GAIA LITE MILK UHT TETRA PACK 1LTR|CBX=12;GAIA LITE MILK UHT TETRA PACK 1LTR|PKT=1;GAIA LITE MILK UHT TETRA PACK 1LTR|FREE PKT=1;GAIA GOLD MILK UHT TETRA PACK 1LTR|CBX=12;GAIA GOLD MILK UHT TETRA PACK 1LTR|PKT=1;GAIA GOLD MILK UHT TETRA PACK 1LTR|FREE PKT=1;GAIA MISHTI DOI 80G|BOX=0.96;GAIA MISHTI DOI 80G|CUPS=0.08;"
    s = s & "GAIA MISHTI DOI 80G|FREE BOX=0.08;DAHI CUP 200 GRMS BOX(6PC)|BOX=1.2;DAHI CUP 200 GRMS BOX(6PC)|CUPS=0.2;DAHI CUP 200 GRMS BOX(6PC)|FREE BOX=0.2;DAHI CUP 200 GRMS|BOX=2.4;DAHI CUP 200 GRMS|CUPS=0.2;DAHI CUP 200 GRMS|FREE BOX=0.2;PD 200 GRMS|CRT=12;PD 200 GRMS|PKT=0.2;PD 200 GRMS|FREE PKT=0.2;PD 400 GRMS|CRT=12;PD 400 GRMS|PKT=0.4;PD 400 GRMS|FREE PKT=0.4;PD 1 KG|CRT=12;"
    s = s & "PD 1 KG|PKT=1;PD 1 KG|FREE PKT=1;PD 5 KG|CRT=10;PD 5 KG|PKT=5;PD 5 KG|FREE PKT=5;PD 5 KG|BKT=5;PD 15 KG|BKT=15;LYCHEE 90 GRMS|BOX=1.08;LYCHEE 90 GRMS|CUPS=0.09;LYCHEE 90 GRMS|FREE BOX=0.09;MUSKMELON 90 GRMS|BOX=1.08;MUSKMELON 90 GRMS|CUPS=0.09;MUSKMELON 90 GRMS|FREE BOX=0.09;KD 200 GRMS|CRT=12;KD 200 GRMS|PKT=0.2;KD 200 GRMS|FREE PKT=0.2;KD 1 KG|CRT=12;KD 1 KG|PKT=1;"
    s = s & "KD 1 KG|FREE PKT=1;KD 5 KG|CRT=10;KD 5 KG|PKT=5;KD 5 KG|FREE PKT=5;KD 5 KG|BKT=5;KD 15KG|BKT=15;KD 15KG|CRT=12;KD 15KG|PKT=0.2;KD 15KG|FREE PKT=0.2;SWEET LASSI 180 GRMS|CRT=10.8;SWEET LASSI 180 GRMS|PKT=0.18;SWEET LASSI 180 GRMS|FREE PKT=0.18;PLAIN (GLASS) LASSI 180 GRMS|BOX3=1.8;PLAIN (GLASS) LASSI 180 GRMS|GLASS=0.18;PLAIN (GLASS) LASSI 180 GRMS|FREE BOX=0.18;"
    s = s & "MANGO LASSI 180 GRMS|BOX2=1.8;MANGO LASSI 180 GRMS|GLASS=0.18;MANGO LASSI 180 GRMS|FREE BOX=0.18;STRAWBERRY LASSI 180 GRMS|BOX2=1.8;STRAWBERRY LASSI 180 GRMS|GLASS=0.18;STRAWBERRY LASSI 180 GRMS|FREE BOX=0.18;STRAWBERRY LASSI 180 GRMS|BOX=1.8;PANEER 200 GRMS|BOX=1;PANEER 200 GRMS|PKT=0.2;PANEER 200 GRMS|FREE PKT=0.2;ZOFF GARAM MASHALA|FREE PKT=0.08;PANEER 500 GRMS|BRAND=0.5;"
    s = s & "PANEER 500 GRMS|FREE PKT=0.5;PANEER 1 KG|BRAND=1;PANEER 1 KG|FREE PKT=1;PANEER 5 KG|BRAND=5;1 KG|LOOSE=1;5 KG|LOOSE=5;5 KG|(IN KG)=0;GHEE 100 ML|BOX=12;GHEE 100 ML|JAR=0.1;GHEE 200 ML|BOX=18;GHEE 200 ML|JAR=0.2;GHEE 500 ML|BOX=18;GHEE 500 ML|JAR=0.5;GHEE 1 LTR|BOX=18;GHEE 1 LTR|JAR=1;GHEE 1 LTR CEKA PACK|BOX=18;GHEE 1 LTR CEKA PACK|JAR=1;COW GHEE 1 LTR CEKA PACK|BOX=18;"
    s = s & "COW GHEE 1 LTR CEKA PACK|JAR=1;GHEE 5 LTRS|BOX=20;GHEE 5 LTRS|JAR=5;COW GHEE 200 ML|BOX=16;COW GHEE 200 ML|JAR=0.2;COW GHEE 500 ML|BOX=18;COW GHEE 500 ML|JAR=0.5;COW GHEE 1LTR|BOX=18;COW GHEE 1LTR|JAR=1;GHEE 15 KG|TIN=15;SHRIKHAND KE 80G|BOX=0.96;SHRIKHAND KE 80G|CUP=0.08;SHRIKHAND KE 80G|FREE BOX=0.08;SHRIKHAND KE 80G-BOX(6PC)|BOX=17.28;SHRIKHAND KE 80G-BOX(6PC)|FREE BOX=2.88;"
    s = s & "SHAHI RABDI 80G|TREY=0.48;SHAHI RABDI 80G|FREE TRAY=0.48;SHAHI RABDI 80G|BOX=0.96;SHAHI RABDI 80G|FREE BOX=0.96;SHAHI RABDI 80G-BOX(6PC)|BOX=0.96;SHAHI RABDI 80G-BOX(6PC)|FREE BOX=0.96;PEDA 200G|BOX=0.2;PEDA 200G|FREE CBX=0.2;KESAR PEDA 200G|BOX=0.2;KESAR PEDA 200G|FREE CBX=0.2;BROWN 1KG|KG=1;BROWN 1KG|VALUE IN RS.=0;"
    s = s & "MASALA CHAAS 200 GRMS|CRT=12;MASALA CHAAS 200 GRMS|PKT=0.2;MASALA CHAAS 200 GRMS|FREE PKT=0.2;MASALA CHAAS GLAAS 180 GRMS|BOX=1.8;MASALA CHAAS GLAAS 180 GRMS|GLASS=0.18;MASALA CHAAS GLAAS 180 GRMS|FREE BOX=0.18;"
    WMStandardData = s
End Function

Private Function WMOverrideData() As String
    Dim s As String
    s = s & "CM 500 ML|CRT=11.28;CM 500 ML|PKT=0.47;CM 500 ML|FREE PKT=0.47;GAIA PREMIUM MILK UHT TETRA PACK 1 LTR|PKT=0.5;GAIA PREMIUM MILK UHT TETRA PACK 1 LTR|FREE PKT=0.5;GAIA LITE MILK UHT TETRA PACK 1LTR|PKT=0.5;GAIA LITE MILK UHT TETRA PACK 1LTR|FREE PKT=0.5;GAIA GOLD MILK UHT TETRA PACK 1LTR|PKT=0.2;GAIA GOLD MILK UHT TETRA PACK 1LTR|FREE PKT=0.2;DAHI CUP 200 GRMS BOX(6PC)|BOX=1.08;"
    s = s & "DAHI CUP 200 GRMS BOX(6PC)|CUPS=0.09;DAHI CUP 200 GRMS BOX(6PC)|FREE BOX=0.09;DAHI CUP 180 GRMS|BOX=1.44;DAHI CUP 180 GRMS|CUPS=0.18;DAHI CUP 180 GRMS|FREE BOX=0.18;SWEET LASSI 180 GRMS|CRT=12;SWEET LASSI 180 GRMS|PKT=0.2;SWEET LASSI 180 GRMS|FREE PKT=0.2;PLAIN (GLASS) LASSI 180 GRMS|BOX3=1.44;MANGO LASSI 180 GRMS|BOX2=1.44;STRAWBERRY LASSI 180 GRMS|BOX2=1.44;"
    s = s & "STRAWBERRY LASSI 180 GRMS|BOX=1.44;PANEER 200 GRMS|BOX=12;GHEE 200 ML|BOX=16;GHEE 500 ML|BOX=16;GHEE 1 LTR|BOX=16;GHEE 900 ML CEKA PACK|BOX=18;GHEE 900 ML CEKA PACK|JAR=1;COW GHEE 900 ML CEKA PACK|BOX=18;COW GHEE 900 ML CEKA PACK|JAR=1;COW GHEE 200 ML|BOX=18;GHEE 20 ML POUCH|BOX=12;GHEE 20 ML POUCH|JAR=0.1;GHEE 20 ML POUCH|FREE JAR=0.18;SHRIKHAND KE 80G-BOX(6PC)|BOX=69.12;"
    s = s & "SHRIKHAND KE 80G-BOX(6PC)|FREE BOX=5.76;SHAHI RABDI 80G|FREE TRAY=0.08;SHAHI RABDI 80G|FREE BOX=0.08;SHAHI RABDI 80G-BOX(6PC)|BOX=0.08;SHAHI RABDI 80G-BOX(6PC)|FREE BOX=0.08;"
    WMOverrideData = s
End Function

Private Function FindLabelRow(ByVal ws As Worksheet, ByVal label As String) As Long
    Dim lastRow As Long, r As Long, lb As String
    lastRow = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).row + 20
    For r = FIRST_DATA_ROW To lastRow
        lb = Norm(ws.Cells(r, 2).Value) & " " & Norm(ws.Cells(r, 3).Value)
        If InStr(lb, label) > 0 Then FindLabelRow = r: Exit Function
    Next r
End Function

Private Function ResolveSheet(ByVal routeN As String, ByVal distN As String) As Worksheet
    Dim target As String: target = RouteToSheet(routeN)
    If Len(target) > 0 Then
        If SheetExists(target) Then
            If FindDistRow(ThisWorkbook.Worksheets(target), distN) > 0 Then
                Set ResolveSheet = ThisWorkbook.Worksheets(target)
                Exit Function
            End If
        End If
    End If
    Dim s As Variant, hit As Worksheet, hits As Long
    For Each s In CitySheets()
        If SheetExists(CStr(s)) Then
            If FindDistRow(ThisWorkbook.Worksheets(CStr(s)), distN) > 0 Then
                Set hit = ThisWorkbook.Worksheets(CStr(s)): hits = hits + 1
            End If
        End If
    Next s
    If hits = 1 Then Set ResolveSheet = hit
End Function

Private Function FindDistRow(ByVal ws As Worksheet, ByVal distN As String) As Long
    Dim lastRow As Long, row As Long, nm As String
    lastRow = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).row
    For row = FIRST_DATA_ROW To lastRow
        nm = Norm(ws.Cells(row, NAME_COL).Value)
        If Len(nm) > 0 Then
            If Not IsReserved(nm) Then
                If nm = distN Then FindDistRow = row: Exit Function
            End If
        End If
    Next row
End Function

' Product name of a column: row 3 normally; if blank, row 2 (Masala Chaas etc.)
Private Function ProdLabelAt(ByVal ws As Worksheet, ByVal col As Long) As String
    Dim v As Variant
    v = ws.Cells(PROD_ROW, col).Value          ' row 3
    If Not IsEmpty(v) And Len(CStr(v)) > 0 Then
        ProdLabelAt = Norm(v)
    Else
        v = ws.Cells(2, col).Value             ' row 2 fallback
        If Not IsEmpty(v) And Len(CStr(v)) > 0 Then ProdLabelAt = Norm(v)
    End If
End Function

Private Function FindColumn(ByVal ws As Worksheet, ByVal prodN As String, _
                            ByVal unitN As String, ByVal skuNameN As String) As Long
    Dim lastCol As Long, j As Long
    lastCol = MeasureLastCol(ws)
    Dim curLabel As String, pl As String
    Dim blockCols() As Long, blockUnits() As String, nBlk As Long
    ReDim blockCols(1 To lastCol): ReDim blockUnits(1 To lastCol)
    curLabel = ""
    For j = 1 To lastCol
        pl = ProdLabelAt(ws, j)
        If Len(pl) > 0 Then curLabel = pl
        If curLabel = prodN Then
            Dim u As String: u = Norm(ws.Cells(UNIT_ROW, j).Value)
            If Len(u) > 0 Then
                nBlk = nBlk + 1: blockCols(nBlk) = j: blockUnits(nBlk) = u
            End If
        End If
    Next j
    If nBlk = 0 Then Exit Function

    Dim k As Long
    If InStr(skuNameN, "BUCKET") > 0 Then
        For k = 1 To nBlk
            If blockUnits(k) = "BKT" Then FindColumn = blockCols(k): Exit Function
        Next k
    End If
    For k = 1 To nBlk
        If blockUnits(k) = "TIN" And (InStr(skuNameN, "TIN") > 0 Or InStr(skuNameN, "15 KG") > 0) Then
            FindColumn = blockCols(k): Exit Function
        End If
    Next k

    Dim cand() As String, ci As Long
    cand = Split(UnitCandidates(unitN), "|")
    For ci = LBound(cand) To UBound(cand)
        If Len(cand(ci)) > 0 Then
            For k = 1 To nBlk
                If blockUnits(k) = cand(ci) Then FindColumn = blockCols(k): Exit Function
            Next k
        End If
    Next ci

    Dim only As Long, cnt As Long
    For k = 1 To nBlk
        If InStr(blockUnits(k), "FREE") = 0 Then only = blockCols(k): cnt = cnt + 1
    Next k
    If cnt = 1 Then FindColumn = only
End Function

Private Function RouteToSheet(ByVal r As String) As String
    Select Case r
        Case "RAIPUR 1": RouteToSheet = "Raipur-1"
        Case "RAIPUR 2": RouteToSheet = "Raipur-2"
        Case "RAIPUR 3": RouteToSheet = "Raipur-3"
        Case "BHILAI": RouteToSheet = "Bhilai"
        Case "BILASPUR": RouteToSheet = "Bilaspur"
        Case "RAIGARH": RouteToSheet = "Raigarh"
        Case "ODISHA", "ODISA": RouteToSheet = "Odisha"
        Case "MAHASAMUND", "MAHASMUND": RouteToSheet = "Mahasamund"
        Case "RAJNANDGAON": RouteToSheet = "Rajnandgaon"
        Case "KORBA": RouteToSheet = "Korba"
        Case "NAGPUR": RouteToSheet = "Nagpur"
        Case "OTHER STATES": RouteToSheet = "Other States"
        Case "JAGDALPUR": RouteToSheet = "Jagdalpur"
    End Select
End Function

Private Function CitySheets() As Variant
    CitySheets = Array("Raipur-1", "Raipur-2", "Raipur-3", "Bhilai", "Bilaspur", _
                       "Raigarh", "Odisha", "Mahasamund", "Rajnandgaon", "Korba", _
                       "Nagpur", "Other States", "Jagdalpur")
End Function

Private Function SkuToProduct(ByVal code As Long) As String
    Select Case code
        Case 999:  SkuToProduct = "CM 500 ML"
        Case 1003: SkuToProduct = "TM 500 ML"
        Case 1005: SkuToProduct = "PANEER  200 GRMS"
        Case 1006: SkuToProduct = "PANEER 500 GRMS"
        Case 1007: SkuToProduct = "PANEER 1 KG"
        Case 1008: SkuToProduct = "1 KG"
        Case 1009: SkuToProduct = "5 KG"
        Case 1011: SkuToProduct = "DAHI CUP 200 GRMS"
        Case 1012: SkuToProduct = "PD 200 GRMS"
        Case 1013: SkuToProduct = "PD 400 GRMS"
        Case 1014: SkuToProduct = "PD 1 KG"
        Case 1015: SkuToProduct = "PD 5 KG"
        Case 1016: SkuToProduct = "PD 15 KG"
        Case 1017: SkuToProduct = "LYCHEE 90 GRMS"
        Case 1018: SkuToProduct = "MUSKMELON   90 GRMS"
        Case 1019: SkuToProduct = "KD 1 KG"
        Case 1020: SkuToProduct = "KD 200 GRMS"
        Case 1021: SkuToProduct = "KD 5 KG"
        Case 1022: SkuToProduct = "KD 15KG"
        Case 1023: SkuToProduct = "PLAIN (GLASS) LASSI 180 GRMS"
        Case 1024: SkuToProduct = "MANGO LASSI 180 GRMS"
        Case 1025: SkuToProduct = "STRAWBERRY LASSI 180 GRMS"
        Case 1026: SkuToProduct = "SWEET LASSI 180 GRMS"
        Case 1027: SkuToProduct = "MASALA CHAAS 200 GRMS"
        Case 1028: SkuToProduct = "GHEE 100 ML"
        Case 1029: SkuToProduct = "GHEE 200 ML"
        Case 1030: SkuToProduct = "GHEE 500 ML"
        Case 1031: SkuToProduct = "GHEE 1 LTR"
        Case 1032: SkuToProduct = "GHEE 5 LTRS"
        Case 1033: SkuToProduct = "GHEE 15 KG"
        Case 1034: SkuToProduct = "Peda 200g"
        Case 1035: SkuToProduct = "BROWN 1KG"
        Case 1036: SkuToProduct = "Shahi Rabdi 80g"                  ' FIXED: was "Shahi Rabdi 80g-Box(6Pc)" (BOX unit -> Box column)
        Case 1037: SkuToProduct = "Cow GHEE 1 LTR Ceka Pack"
        Case 1038: SkuToProduct = "GHEE 1 LTR Ceka Pack"
        Case 1039: SkuToProduct = "COW GHEE 200 ML"
        Case 1040: SkuToProduct = "COW GHEE 500 ML"
        Case 1041: SkuToProduct = "COW GHEE 1LTR"
        Case 1042: SkuToProduct = "Maxx UHT 110 ml"
        Case 1043: SkuToProduct = "Shrikhand KE 80g"
        Case 1044: SkuToProduct = "Gaia Mishti Doi 80g"
        Case 1045: SkuToProduct = "Gaia Maxx UHT 400ml"
        Case 1046: SkuToProduct = "MASALA CHAAS GLAAS 180 GRMS"      ' exact column name (row 2), note GLAAS spelling
        Case 1047: SkuToProduct = "PD 5 KG"
        Case 1048: SkuToProduct = "Shahi Rabdi 80g"                  ' TRAY unit -> Trey column
        Case 1049: SkuToProduct = "Kesar Peda 200g"
        Case 1050: SkuToProduct = "Gaia Premium Milk UHT Tetra Pack 1 Ltr"
        Case 1051: SkuToProduct = "Gaia Lite Milk UHT Tetra Pack 1Ltr"
        Case 1052: SkuToProduct = "Gaia Gold Milk UHT Tetra Pack 1Ltr"
        Case 1053: SkuToProduct = "GHEE 900 ML Ceka Pack"
        Case 1054: SkuToProduct = "Cow GHEE 900 ML Ceka Pack"
        Case 1055: SkuToProduct = "GHEE 20 ML Pouch"
        Case 1056: SkuToProduct = "DAHI CUP 200 GRMS Box(6Pc)"
        Case 1057: SkuToProduct = "Shahi Rabdi 80g-Box(6Pc)"        ' -> Box(6Pc) column
        Case 1058: SkuToProduct = "Shrikhand KE 80g-Box(6Pc)"
    End Select
End Function

Private Function UnitCandidates(ByVal u As String) As String
    Select Case u
        Case "CRT": UnitCandidates = "CRT|CBX"
        Case "PCS": UnitCandidates = "PKT|PC|PCS|LOOSE|CUPS|CUP|GLASS|BRAND|(IN KG)|KG|JAR"
        Case "BOX": UnitCandidates = "BOX3|BOX2|BOX|CBX"
        Case "JAR": UnitCandidates = "JAR|BOX"
        Case "TRAY": UnitCandidates = "TREY|TRAY"
        Case "KG": UnitCandidates = "(IN KG)|KG|LOOSE"
        Case Else: UnitCandidates = ""
    End Select
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

Private Function Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Norm = "": Exit Function
    Dim s As String
    s = CStr(v)
    s = Replace(s, vbLf, " "): s = Replace(s, vbCr, " "): s = Replace(s, vbTab, " ")
    On Error Resume Next
    s = Application.WorksheetFunction.Trim(s)
    On Error GoTo 0
    Norm = UCase$(Trim$(s))
End Function

Private Function ParseSkuCode(ByVal v As Variant) As Long
    Dim s As String, i As Long, ch As String, out As String
    s = CStr(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    If Len(out) > 0 Then ParseSkuCode = CLng(out)
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

Private Sub AddOnce(ByRef bag As String, ByVal item As String)
    If InStr(bag, "[" & item & "]") = 0 Then bag = bag & "[" & item & "]" & vbLf
End Sub

