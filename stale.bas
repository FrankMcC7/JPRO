Option Explicit

'===========================
'  MAIN ENTRY POINT
'===========================
Sub Refresh_PortfolioTable()
    Dim fTrig As String, fNon As String, fAll As String
    Dim wbTrig As Workbook, wbNon As Workbook, wbAll As Workbook
    Dim loTrig As ListObject, loNon As ListObject, loAll As ListObject
    Dim loPort As ListObject, loData As ListObject
    Dim dictAll As Object, dictData As Object
    Dim arrOut As Variant, ptr As Long, capacity As Long
    Dim colMap As Object

    '––– STEP 1: Ask user to pick the three files –––
    fTrig = PickFile("Select the TRIGGER file")
    If fTrig = "" Then Exit Sub
    fNon = PickFile("Select the NON-TRIGGER file")
    If fNon = "" Then Exit Sub
    fAll = PickFile("Select the ALL-FUNDS file")
    If fAll = "" Then Exit Sub

    '––– STEP 2: Open each workbook as ReadOnly –––
    Set wbTrig = Workbooks.Open(fTrig, ReadOnly:=True)
    Set wbNon  = Workbooks.Open(fNon,  ReadOnly:=True)
    Set wbAll  = Workbooks.Open(fAll,  ReadOnly:=True)

    '––– STEP 3: Convert each sheet to a ListObject (table) –––
    '    • TRIGGER and NON-TRIGGER: just convert entire UsedRange to a table
    '    • ALL-FUNDS: delete row 1 first, then convert, then immediately filter to keep only Review Status = "Approved"
    Set loTrig = EnsureTable(wbTrig.Worksheets(1), False, False)
    Set loNon  = EnsureTable(wbNon.Worksheets(1),  False, False)
    Set loAll  = EnsureTable(wbAll.Worksheets(1),  True,  True)   ' delete row 1; filter out non-"Approved"

    '––– STEP 4: Grab the two target tables in THIS workbook –––
    Set loPort = ThisWorkbook.Worksheets("Portfolio").ListObjects("PortfolioTable")
    Set loData = ThisWorkbook.Worksheets("Dataset").ListObjects("DatasetTable")

    '––– STEP 5: Build two lookup dictionaries via array‐based loops –––
    '  • dictAll : keyed on “Fund GCI” → array( “IA GCI”, “Fund LEI”, “Fund Code” ), but only where Review Status = "Approved"
    '  • dictData : keyed on “Fund Manager GCI” → array( “Family”, “ECA India Analyst” )
    Set dictAll  = BuildDictFromTable(loAll, "Fund GCI",         Array("IA GCI", "Fund LEI", "Fund Code"), _
                                      "Review Status", "Approved")
    Set dictData = BuildDictFromTable(loData, "Fund Manager GCI", Array("Family", "ECA India Analyst"), _
                                      "", "")

    '––– STEP 6: Clear any existing data (and filters) from PortfolioTable –––
    On Error Resume Next
    loPort.Range.AutoFilter.ShowAllData
    On Error GoTo 0
    If Not loPort.DataBodyRange Is Nothing Then loPort.DataBodyRange.Delete

    '––– STEP 7: Speed optimizations – turn off screen updates, events, and auto‐calculation –––
    With Application
        .ScreenUpdating = False
        .EnableEvents   = False
        .Calculation    = xlCalculationManual
    End With

    '––– STEP 8: Pre‐allocate an output array large enough to hold all Trigger + Non-Trigger rows –––
    Dim trigCount As Long, nonCount As Long
    If Not loTrig.DataBodyRange Is Nothing Then trigCount = loTrig.DataBodyRange.Rows.Count Else trigCount = 0
    If Not loNon.DataBodyRange Is Nothing  Then nonCount  = loNon.DataBodyRange.Rows.Count  Else nonCount  = 0
    capacity = trigCount + nonCount

    ' If there is no data at all, we can bail out immediately
    If capacity = 0 Then GoTo Finalize

    ' We dimension arrOut = (1 To capacity) × (1 To numberOfColumnsInPortfolioTable)
    ReDim arrOut(1 To capacity, 1 To loPort.ListColumns.Count)
    Set colMap = BuildIndex(loPort)   ' maps PortfolioTable column names → their column index in arrOut

    '––– STEP 9: Copy all rows from the TRIGGER table into arrOut (flag = "Trigger") –––
    ptr = FillArrayFast( _
            loSrc:=loTrig, _
            flag:="Trigger", _
            skipCol:="", skipVal:="", _
            dictAll:=dictAll, dictData:=dictData, _
            arrOut:=arrOut, startPtr:=0, colMap:=colMap _
          )

    '––– STEP 10: Copy all rows from the NON-TRIGGER table into arrOut (flag = "Non-Trigger"), skipping Business Unit="FI-ASIA" –––
    ptr = FillArrayFast( _
            loSrc:=loNon, _
            flag:="Non-Trigger", _
            skipCol:="Business Unit", skipVal:="FI-ASIA", _
            dictAll:=dictAll, dictData:=dictData, _
            arrOut:=arrOut, startPtr:=ptr, colMap:=colMap _
          )

    '––– STEP 11: Write arrOut back into PortfolioTable in one bulk operation –––
    If ptr > 0 Then
        loPort.HeaderRowRange.Offset(1, 0).Resize(ptr, UBound(arrOut, 2)).Value = arrOut
        loPort.Resize loPort.HeaderRowRange.Resize(ptr + 1)
    End If

    '––– STEP 12: Remap Region codes in‐place: “US” → “AMRS”, “ASIA” → “APAC” –––
    With loPort.ListColumns("Region").DataBodyRange
        .Replace What:="US",   Replacement:="AMRS", xlWhole
        .Replace What:="ASIA", Replacement:="APAC", xlWhole
    End With

Finalize:
    '––– STEP 13: Re-enable updates/events/calculation –––
    With Application
        .Calculation    = xlCalculationAutomatic
        .EnableEvents   = True
        .ScreenUpdating = True
    End With
End Sub

'===========================
'  FILEPICKER HELPER
'===========================
' Returns the full path of the selected file, or "" if the user cancels.
Private Function PickFile(promptText As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = promptText
        .AllowMultiSelect = False
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

'===========================
'  ENSURE A SHEET IS A TABLE
'===========================
'  ws: the Worksheet you want to turn into a ListObject  
'  deleteRow1 = True → delete the top row before converting  
'  filterApproved = True → immediately filter out any row whose "Review Status" <> "Approved", then delete those filtered rows  
Private Function EnsureTable(ws As Worksheet, deleteRow1 As Boolean, filterApproved As Boolean) As ListObject
    If deleteRow1 Then
        ws.Rows(1).Delete
    End If

    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0

    If ur Is Nothing Then
        ' If the sheet is entirely blank after deleting row 1, create a 1×1 dummy table
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1"), , xlYes)
    Else
        ' Convert the entire UsedRange into a table (first row = header)
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ur, , xlYes)
    End If

    ' If we need to keep only "Approved" in Review Status, delete everything else
    If filterApproved Then
        If ColumnExists(EnsureTable, "Review Status") Then
            Dim colReview As Long
            colReview = EnsureTable.ListColumns("Review Status").Index

            With EnsureTable.Range
                ' Filter to show only rows where Review Status <> "Approved"
                .AutoFilter Field:=colReview, Criteria1:="<>Approved"
                On Error Resume Next
                ' Delete the visible (non-Approved) rows
                EnsureTable.DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete
                On Error GoTo 0
                .AutoFilter  ' clear the filter
            End With
        End If
    End If
End Function

'===========================
'  BUILD A LOOKUP DICTIONARY
'===========================
'  lo        = a ListObject  
'  keyCol    = the column to use as dictionary key (e.g. "Fund GCI" or "Fund Manager GCI")  
'  valCols   = an array of column names whose values you want in the dictionary array  
'  filterCol = optional column name; if provided, we only include rows where that column = filterVal  
'  filterVal = the required value in filterCol to include that row  
'  
' → returns a Dictionary where  
'    Key = each row’s keyCol value  
'    Item(Key) = Variant array of that row’s valCols values  
Private Function BuildDictFromTable( _
    lo As ListObject, _
    keyCol As String, _
    valCols As Variant, _
    filterCol As String, _
    filterVal As String _
) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If lo.DataBodyRange Is Nothing Then
        Set BuildDictFromTable = dict: Exit Function
    End If

    ' Read the entire DataBodyRange into a 2D Variant array
    Dim vData As Variant
    vData = lo.DataBodyRange.Value

    ' Build a header→index map for this table
    Dim idxSrc As Object: Set idxSrc = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        idxSrc(lo.ListColumns(c).Name) = c
    Next c

    ' Find the key column index
    If Not idxSrc.Exists(keyCol) Then
        Set BuildDictFromTable = dict: Exit Function
    End If
    Dim keyIdx As Long: keyIdx = idxSrc(keyCol)

    ' If we have a filter column, get its index (otherwise 0)
    Dim filterIdx As Long: filterIdx = 0
    If filterCol <> "" Then
        If idxSrc.Exists(filterCol) Then filterIdx = idxSrc(filterCol)
    End If

    ' Prepare an array of indices for valCols
    Dim valIdx() As Long, i As Long
    ReDim valIdx(0 To UBound(valCols))
    For i = 0 To UBound(valCols)
        If idxSrc.Exists(valCols(i)) Then
            valIdx(i) = idxSrc(valCols(i))
        Else
            valIdx(i) = 0
        End If
    Next i

    ' Loop through every row of vData
    Dim r As Long, k As Variant, vArr() As Variant
    For r = 1 To UBound(vData, 1)
        ' If filter is specified, skip rows where filterCol <> filterVal
        If filterIdx > 0 Then
            If CStr(vData(r, filterIdx)) <> filterVal Then GoTo NextRowDict
        End If

        k = vData(r, keyIdx)
        If Len(Trim(CStr(k))) > 0 Then
            ReDim vArr(0 To UBound(valCols))
            For i = 0 To UBound(valCols)
                If valIdx(i) > 0 Then
                    vArr(i) = vData(r, valIdx(i))
                Else
                    vArr(i) = vbNullString
                End If
            Next i
            dict(k) = vArr
        End If
NextRowDict:
    Next r

    Set BuildDictFromTable = dict
End Function

'===========================
'  BUILD A COLUMN→INDEX MAP
'===========================
' Given a ListObject, returns a Dictionary where  
'   Key = column name, Value = its position (1-based) in that ListObject  
Private Function BuildIndex(lo As ListObject) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        m(lo.ListColumns(c).Name) = c
    Next c
    Set BuildIndex = m
End Function

'===========================
'  FILL THE OUTPUT ARRAY
'===========================
' loSrc        = the source ListObject (Trigger or Non-Trigger)  
' flag         = either "Trigger" or "Non-Trigger"  
' skipCol      = name of a column to test for skipping (e.g. "Business Unit")  
' skipVal      = if skipCol=skipVal, that row is omitted  
' dictAll      = dictionary built from All-Funds (key=Fund GCI → IA GCI/LEI/Code)  
' dictData     = dictionary built from DatasetTable (key=Fund Mgr GCI → Family/ECA)  
' arrOut       = the pre-dim’d 2D array (1 To capacity, 1 To columnCount)  
' startPtr     = how many rows have already been filled in arrOut (0-based count).  
' colMap       = dictionary mapping PortfolioTable column names → arrOut column index  
'  
'→ returns the new pointer (row count) in arrOut after filling all valid rows  
Private Function FillArrayFast( _
    loSrc As ListObject, _
    flag As String, _
    skipCol As String, _
    skipVal As String, _
    dictAll As Object, _
    dictData As Object, _
    arrOut As Variant, _
    startPtr As Long, _
    colMap As Object _
) As Long

    ' If the source has no data, skip immediately
    If loSrc.DataBodyRange Is Nothing Then
        FillArrayFast = startPtr
        Exit Function
    End If

    ' Read entire source data into a Variant array
    Dim vSrc As Variant
    vSrc = loSrc.DataBodyRange.Value

    ' Build a header→index map for the source table
    Dim idxSrc As Object: Set idxSrc = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To loSrc.ListColumns.Count
        idxSrc(loSrc.ListColumns(c).Name) = c
    Next c

    ' Find indices of columns we actually need (skipCol, "Fund GCI", etc.)
    Dim skipIdx As Long: skipIdx = 0
    If skipCol <> "" Then
        If idxSrc.Exists(skipCol) Then skipIdx = idxSrc(skipCol)
    End If

    Dim idx_FundGCI As Long:   idx_FundGCI   = IIf(idxSrc.Exists("Fund GCI"),           idxSrc("Fund GCI"),           0)
    Dim idx_FundMgr As Long:   idx_FundMgr   = IIf(idxSrc.Exists("Fund Manager"),       idxSrc("Fund Manager"),       0)
    Dim idx_FundName As Long:  idx_FundName  = IIf(idxSrc.Exists("Fund Name"),          idxSrc("Fund Name"),          0)
    Dim idx_CreditOff As Long: idx_CreditOff = IIf(idxSrc.Exists("Credit Officer"),     idxSrc("Credit Officer"),     0)
    Dim idx_WCA As Long:       idx_WCA       = IIf(idxSrc.Exists("WCA"),                idxSrc("WCA"),                0)
    Dim idx_Region As Long:    idx_Region    = IIf(idxSrc.Exists("Region"),             idxSrc("Region"),             0)
    Dim idx_WksMissing As Long:    idx_WksMissing    = IIf(idxSrc.Exists("Wks Missing"),     idxSrc("Wks Missing"),     0)
    Dim idx_WeeksMissing As Long:  idx_WeeksMissing  = IIf(idxSrc.Exists("Weeks Missing"),   idxSrc("Weeks Missing"),   0)
    Dim idx_LatestNAV As Long: idx_LatestNAV = IIf(idxSrc.Exists("Latest NAV Date"),   idxSrc("Latest NAV Date"),    0)
    Dim idx_ReqNAV As Long
    idx_ReqNAV = 0
    If idxSrc.Exists("Req NAV Date") Then
        idx_ReqNAV = idxSrc("Req NAV Date")
    ElseIf idxSrc.Exists("Required NAV Date") Then
        idx_ReqNAV = idxSrc("Required NAV Date")
    End If

    ' Main loop: for each row in vSrc(r, *) ...
    Dim r As Long, outRow As Long
    Dim fGCI As Variant, fMgrGCI As Variant

    For r = 1 To UBound(vSrc, 1)
        ' 1) If skipCol is defined and that cell = skipVal, skip this row
        If skipIdx > 0 Then
            If CStr(vSrc(r, skipIdx)) = skipVal Then GoTo NextRow
        End If

        ' 2) Advance our pointer in arrOut (fill row #)
        startPtr = startPtr + 1
        outRow = startPtr

        ' 3) Copy “Fund GCI” and other core fields
        fGCI = IIf(idx_FundGCI > 0, vSrc(r, idx_FundGCI), vbNullString)
        arrOut(outRow, colMap("Fund GCI")) = fGCI
        If idx_FundMgr > 0 Then  arrOut(outRow, colMap("Fund Manager"))    = vSrc(r, idx_FundMgr)
        If idx_FundName > 0 Then arrOut(outRow, colMap("Fund Name"))       = vSrc(r, idx_FundName)
        If idx_CreditOff > 0 Then arrOut(outRow, colMap("Credit Officer"))  = vSrc(r, idx_CreditOff)
        If idx_WCA > 0 Then       arrOut(outRow, colMap("WCA"))             = vSrc(r, idx_WCA)
        If idx_Region > 0 Then    arrOut(outRow, colMap("Region"))          = vSrc(r, idx_Region)

        ' “Wks Missing” vs “Weeks Missing” alias
        If idx_WksMissing > 0 Then
            arrOut(outRow, colMap("Wks Missing")) = vSrc(r, idx_WksMissing)
        ElseIf idx_WeeksMissing > 0 Then
            arrOut(outRow, colMap("Wks Missing")) = vSrc(r, idx_WeeksMissing)
        End If

        If idx_LatestNAV > 0 Then arrOut(outRow, colMap("Latest NAV Date")) = vSrc(r, idx_LatestNAV)
        If idx_ReqNAV > 0 Then    arrOut(outRow, colMap("Req NAV Date"))   = vSrc(r, idx_ReqNAV)

        arrOut(outRow, colMap("Trigger/Non-Trigger")) = flag

        ' 4) Lookup in dictAll → fill “Fund Manager GCI”, “Fund LEI”, “Fund Code”
        fMgrGCI = vbNullString
        If Not IsEmpty(fGCI) Then
            If dictAll.Exists(fGCI) Then
                fMgrGCI = dictAll(fGCI)(0)
                arrOut(outRow, colMap("Fund Manager GCI")) = dictAll(fGCI)(0)
                arrOut(outRow, colMap("Fund LEI"))         = dictAll(fGCI)(1)
                arrOut(outRow, colMap("Fund Code"))        = dictAll(fGCI)(2)
            End If
        End If

        ' 5) Lookup in dictData (via fMgrGCI) → fill “Family” & “ECA India Analyst”
        If Len(Trim(CStr(fMgrGCI))) > 0 Then
            If dictData.Exists(fMgrGCI) Then
                arrOut(outRow, colMap("Family"))            = dictData(fMgrGCI)(0)
                arrOut(outRow, colMap("ECA India Analyst")) = dictData(fMgrGCI)(1)
            End If
        End If

NextRow:
    Next r

    FillArrayFast = startPtr
End Function

'===========================
'  COLUMN‐EXISTS CHECK
'===========================
Private Function ColumnExists(lo As ListObject, colName As String) As Boolean
    On Error Resume Next
    ColumnExists = (lo.ListColumns(colName).Index > 0)
    Err.Clear: On Error GoTo 0
End Function