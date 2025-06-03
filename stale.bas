Option Explicit

'======================= MASTER ROUTINE ========================
Sub Refresh_PortfolioTable()
    Dim fTrig As String, fNon As String, fAll As String
    Dim wbTrig As Workbook, wbNon As Workbook, wbAll As Workbook
    Dim loTrig As ListObject, loNon As ListObject, loAll As ListObject
    Dim loPort As ListObject, loData As ListObject
    Dim dictAll As Object, dictData As Object
    Dim colMap As Object, arr(), ptr As Long, capacity As Long

    '––– 1. Prompt for files (TRIGGER, NON‐TRIGGER, ALL‐FUNDS) –––
    fTrig = PickFile("Select TRIGGER file"):        If fTrig = "" Then Exit Sub
    fNon  = PickFile("Select NON‐TRIGGER file"):    If fNon = "" Then Exit Sub
    fAll  = PickFile("Select ALL‐FUNDS file"):      If fAll = "" Then Exit Sub

    '––– 2. Open workbooks as ReadOnly –––
    Set wbTrig = Workbooks.Open(fTrig, ReadOnly:=True)
    Set wbNon  = Workbooks.Open(fNon,  ReadOnly:=True)
    Set wbAll  = Workbooks.Open(fAll,  ReadOnly:=True)

    '––– 3. Convert sheets to tables –––
    Set loTrig = EnsureTable(wbTrig.Worksheets(1), False, False)
    Set loNon  = EnsureTable(wbNon.Worksheets(1),  False, False)
    ' For All‐Funds: delete row 1 first, then keep only “Approved” in Review Status
    Set loAll = EnsureTable(wbAll.Worksheets(1), True, True)

    '––– 4. Identify PortfolioTable & DatasetTable in THIS workbook –––
    Set loPort = ThisWorkbook.Worksheets("Portfolio").ListObjects("PortfolioTable")
    Set loData = ThisWorkbook.Worksheets("Dataset").ListObjects("DatasetTable")

    '––– 5. Build dictionaries for lookups (Fund GCI → IA GCI, LEI, Code; Fund Mgr GCI → Family, ECA) –––
    Set dictAll  = BuildDict(loAll,  "Fund GCI",         Array("IA GCI", "Fund LEI", "Fund Code"))
    Set dictData = BuildDict(loData, "Fund Manager GCI", Array("Family", "ECA India Analyst"))

    '––– 6. Clear any filters and delete existing rows from PortfolioTable –––
    On Error Resume Next: loPort.Range.AutoFilter.ShowAllData: On Error GoTo 0
    If Not loPort.DataBodyRange Is Nothing Then loPort.DataBodyRange.Delete

    '––– 7. Turn OFF screen updating/events/calculation for speed –––
    With Application
        .ScreenUpdating = False
        .EnableEvents   = False
        .Calculation    = xlCalculationManual
    End With

    '––– 8. Pre‐allocate output array to size (trigger rows + non‐trigger rows) × number of columns –––
    Set colMap  = BuildIndex(loPort)
    capacity   = IIf(Not loTrig.DataBodyRange Is Nothing, loTrig.DataBodyRange.Rows.Count, 0) _
               + IIf(Not loNon.DataBodyRange Is Nothing, loNon.DataBodyRange.Rows.Count, 0)
    ReDim arr(1 To capacity, 1 To loPort.ListColumns.Count)

    '––– 9. Populate rows from Trigger file (flag = “Trigger”) –––
    ptr = FillArray(loTrig, "Trigger", "", "", dictAll, dictData, arr, 0, colMap)

    '––– 10. Populate rows from Non‐Trigger file (skip “FI‐ASIA”) –––
    ptr = FillArray(loNon, "Non‐Trigger", "Business Unit", "FI‐ASIA", dictAll, dictData, arr, ptr, colMap)

    '––– 11. Write array back underneath PortfolioTable’s header row –––
    If ptr > 0 Then
        'Use HeaderRowRange.Offset(1) even if DataBodyRange is Nothing
        loPort.HeaderRowRange.Offset(1, 0).Resize(ptr, UBound(arr, 2)).Value = arr
        'Resize table to include header + new rows
        loPort.Resize loPort.HeaderRowRange.Resize(ptr + 1)
    End If

    '––– 12. Remap Region codes (“US”→“AMRS”, “ASIA”→“APAC”) –––
    With loPort.ListColumns("Region").DataBodyRange
        .Replace "US",   "AMRS", xlWhole
        .Replace "ASIA", "APAC", xlWhole
    End With

CleanUp:
    '––– 13. Restore application settings –––
    With Application
        .Calculation    = xlCalculationAutomatic
        .EnableEvents   = True
        .ScreenUpdating = True
    End With
End Sub

'======================= HELPER: PICK A FILE =======================
Private Function PickFile(promptText As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = promptText
        .AllowMultiSelect = False
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

'================== HELPER: ENSURE A SHEET IS A TABLE =================
' deleteRow1 = True will delete the top row before resizing
' filterApproved = True will remove ANY row where "Review Status" <> "Approved"
Private Function EnsureTable(ws As Worksheet, deleteRow1 As Boolean, filterApproved As Boolean) As ListObject
    Dim tblRange As Range, ur As Range, lo As ListObject

    ' 1. If requested, delete the first row (often a spurious header)
    If deleteRow1 Then ws.Rows(1).Delete

    ' 2. Define the used‐range after deletion
    Set ur = ws.UsedRange
    If ur Is Nothing Then
        ' If sheet is completely empty, create an empty 1×1 table
        Set tblRange = ws.Range("A1")
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        Exit Function
    End If

    ' 3. Convert that UsedRange to a ListObject (always treat first row as header)
    Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ur, , xlYes)

    ' 4. If requested, filter out any row whose "Review Status" <> "Approved"
    If filterApproved Then
        If ColumnExists(EnsureTable, "Review Status") Then
            With EnsureTable.Range
                .AutoFilter Field:=EnsureTable.ListColumns("Review Status").Index, _
                            Criteria1:="<>Approved"
                On Error Resume Next
                ' Delete only the VISIBLE rows (i.e. non-“Approved”)
                EnsureTable.DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete
                On Error GoTo 0
                .AutoFilter ' Clear filter
            End With
        End If
    End If
End Function

'================== HELPER: BUILD DICTIONARY FOR LOOKUP =================
'   keyCol   = name of the column to key on (e.g., "Fund GCI")
'   valCols  = array of column names whose values get stored in a small array
Private Function BuildDict(lo As ListObject, keyCol As String, valCols As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As Variant, v(), i As Long

    If lo.DataBodyRange Is Nothing Then
        Set BuildDict = d: Exit Function
    End If

    For r = 1 To lo.DataBodyRange.Rows.Count
        k = lo.DataBodyRange(r, lo.ListColumns(keyCol).Index).Value
        If Len(Trim(k)) > 0 Then
            ReDim v(0 To UBound(valCols))
            For i = 0 To UBound(valCols)
                v(i) = lo.DataBodyRange(r, lo.ListColumns(valCols(i)).Index).Value
            Next i
            d(k) = v
        End If
    Next r

    Set BuildDict = d
End Function

'================== HELPER: MAP COLUMN NAMES TO INDEX =================
Private Function BuildIndex(lo As ListObject) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim c As Long

    For c = 1 To lo.ListColumns.Count
        m(lo.ListColumns(c).Name) = c
    Next c

    Set BuildIndex = m
End Function

'================ HELPER: FILL ARRAY FROM A SOURCE TABLE =================
'   loSrc       = source ListObject (Trigger or Non-Trigger)
'   flag        = either "Trigger" or "Non-Trigger"
'   skipCol     = name of column to check for skipping (e.g. "Business Unit")
'   skipVal     = if skipCol = skipVal, that row is omitted
'   dictAll     = dictionary from All-Funds (keyed on Fund GCI)
'   dictData    = dictionary from DatasetTable (keyed on Fund Mgr GCI)
'   arr         = pre-dim’d 2D array to hold output
'   startPtr    = row index in arr to begin filling (1-based)
'   colMap      = dictionary mapping table‐column names → index in arr
Private Function FillArray( _
    loSrc       As ListObject, _
    flag        As String, _
    skipCol     As String, _
    skipVal     As String, _
    dictAll     As Object, _
    dictData    As Object, _
    arr         As Variant, _
    startPtr    As Long, _
    colMap      As Object) As Long

    Dim r As Long, fGCI As String, fMgrGCI As String
    Dim srcRow As Range

    If loSrc.DataBodyRange Is Nothing Then
        FillArray = startPtr: Exit Function
    End If

    For r = 1 To loSrc.DataBodyRange.Rows.Count
        Set srcRow = loSrc.DataBodyRange.Rows(r)

        ' 1. Skip rows if skipCol = skipVal
        If skipCol <> "" Then
            If srcRow.Cells(1, loSrc.ListColumns(skipCol).Index).Value = skipVal Then GoTo NextRow
        End If

        ' 2. Copy common columns
        startPtr = startPtr + 1
        fGCI = srcRow.Cells(1, loSrc.ListColumns("Fund GCI").Index).Value

        With arr
            .Item(startPtr, colMap("Fund GCI"))        = fGCI
            .Item(startPtr, colMap("Fund Manager"))    = srcRow.Cells(1, loSrc.ListColumns("Fund Manager").Index).Value
            .Item(startPtr, colMap("Fund Name"))       = srcRow.Cells(1, loSrc.ListColumns("Fund Name").Index).Value
            .Item(startPtr, colMap("Credit Officer"))  = srcRow.Cells(1, loSrc.ListColumns("Credit Officer").Index).Value
            .Item(startPtr, colMap("WCA"))             = srcRow.Cells(1, loSrc.ListColumns("WCA").Index).Value
            .Item(startPtr, colMap("Region"))          = srcRow.Cells(1, loSrc.ListColumns("Region").Index).Value
            ' Weeks Missing vs Wks Missing alias
            .Item(startPtr, colMap("Wks Missing"))     = srcRow.Cells(1, loSrc.ListColumns( _
                                                         OptionalAlias(loSrc, "Wks Missing", "Weeks Missing")).Index).Value
            .Item(startPtr, colMap("Latest NAV Date")) = srcRow.Cells(1, loSrc.ListColumns("Latest NAV Date").Index).Value
            .Item(startPtr, colMap("Req NAV Date"))    = srcRow.Cells(1, loSrc.ListColumns( _
                                                         OptionalAlias(loSrc, "Req NAV Date", "Required NAV Date")).Index).Value
            .Item(startPtr, colMap("Trigger/Non-Trigger")) = flag
        End With

        ' 3. Lookup IA GCI, LEI, Code in All-Funds
        If dictAll.Exists(fGCI) Then
            fMgrGCI = dictAll(fGCI)(0)
            arr(startPtr, colMap("Fund Manager GCI")) = dictAll(fGCI)(0)
            arr(startPtr, colMap("Fund LEI"))         = dictAll(fGCI)(1)
            arr(startPtr, colMap("Fund Code"))        = dictAll(fGCI)(2)
        Else
            fMgrGCI = ""
        End If

        ' 4. Lookup Family/ECA via Fund Mgr GCI
        If Len(Trim(fMgrGCI)) > 0 Then
            If dictData.Exists(fMgrGCI) Then
                arr(startPtr, colMap("Family"))            = dictData(fMgrGCI)(0)
                arr(startPtr, colMap("ECA India Analyst")) = dictData(fMgrGCI)(1)
            End If
        End If

NextRow:
    Next r

    FillArray = startPtr
End Function

'=========== HELPER: RETURN column name if it exists, else fallback name ===========
Private Function OptionalAlias(lo As ListObject, nm1 As String, nm2 As String) As String
    OptionalAlias = IIf(ColumnExists(lo, nm1), nm1, nm2)
End Function

'=========== HELPER: CHECK IF A COLUMN EXISTS IN A TABLE ===========
Private Function ColumnExists(lo As ListObject, colName As String) As Boolean
    On Error Resume Next
    ColumnExists = (lo.ListColumns(colName).Index > 0)
    Err.Clear: On Error GoTo 0
End Function
'===================================================================================