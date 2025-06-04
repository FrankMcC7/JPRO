Option Explicit

'===========================
'  MAIN ENTRY POINT
'===========================
Sub Refresh_PortfolioTable()
    On Error GoTo ErrHandler
    
    Dim fTrig As String, fNon As String, fAll As String
    Dim wbTrig As Workbook, wbNon As Workbook, wbAll As Workbook
    Dim loTrig As ListObject, loNon As ListObject, loAll As ListObject
    Dim loPort As ListObject, loData As ListObject
    Dim dictAll As Object, dictData As Object
    Dim arrOut As Variant, ptr As Long, capacity As Long
    Dim colMap As Object
    Dim startTime As Single, endTime As Single, elapsedSec As Single
    Dim trigCount As Long, nonCount As Long
    Dim wbMain As Workbook
    
    Set wbMain = ThisWorkbook
    startTime = Timer
    
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
    Set loTrig = EnsureTable(wbTrig.Worksheets(1), False, False)
    Set loNon  = EnsureTable(wbNon.Worksheets(1),  False, False)
    Set loAll  = EnsureTable(wbAll.Worksheets(1),  True,  True)   ' delete row 1; filter out non-"Approved"
    
    '––– STEP 4: Grab the two target tables in THIS workbook –––
    Set loPort = wbMain.Worksheets("Portfolio").ListObjects("PortfolioTable")
    Set loData = wbMain.Worksheets("Dataset").ListObjects("DatasetTable")
    
    '––– STEP 5: Build dictAll in chunks to limit memory usage –––
    Set dictAll = CreateObject("Scripting.Dictionary")
    If Not loAll.DataBodyRange Is Nothing Then
        Dim totalRows As Long, blockSize As Long, startRow As Long, endRow As Long
        Dim arrBlock As Variant, r As Long
        Dim idxAll As Object: Set idxAll = CreateObject("Scripting.Dictionary")
        Dim c As Long, keyColIdx As Long, iA_GCI As Long, fundLEI As Long, fundCode As Long, reviewIdx As Long
        ' Build header→index map for loAll
        For c = 1 To loAll.ListColumns.Count
            idxAll(loAll.ListColumns(c).Name) = c
        Next c
        ' Determine relevant column indices
        keyColIdx = IIf(idxAll.Exists("Fund GCI"), idxAll("Fund GCI"), 0)
        iA_GCI    = IIf(idxAll.Exists("IA GCI"), idxAll("IA GCI"), 0)
        fundLEI   = IIf(idxAll.Exists("Fund LEI"), idxAll("Fund LEI"), 0)
        fundCode  = IIf(idxAll.Exists("Fund Code"), idxAll("Fund Code"), 0)
        reviewIdx = IIf(idxAll.Exists("Review Status"), idxAll("Review Status"), 0)
        
        totalRows = loAll.DataBodyRange.Rows.Count
        blockSize = 50000
        For startRow = 1 To totalRows Step blockSize
            endRow = Application.Min(startRow + blockSize - 1, totalRows)
            arrBlock = loAll.DataBodyRange.Rows(startRow & ":" & endRow).Value
            For r = 1 To UBound(arrBlock, 1)
                ' Progress feedback in status bar
                If (r Mod 5000) = 0 Then _
                    Application.StatusBar = "Building All-Funds dictionary: row " & (startRow + r - 1) & " of " & totalRows
                ' Skip if Review Status <> "Approved"
                If reviewIdx > 0 Then
                    If CStr(arrBlock(r, reviewIdx)) <> "Approved" Then GoTo NextAllRow
                End If
                ' Capture key and values
                Dim k As Variant
                k = arrBlock(r, keyColIdx)
                If Len(Trim(CStr(k))) > 0 Then
                    Dim vArr(0 To 2) As Variant
                    If iA_GCI > 0 Then   vArr(0) = arrBlock(r, iA_GCI)
                    If fundLEI > 0 Then  vArr(1) = arrBlock(r, fundLEI)
                    If fundCode > 0 Then vArr(2) = arrBlock(r, fundCode)
                    dictAll(k) = vArr
                End If
NextAllRow:
            Next r
        Next startRow
        Application.StatusBar = False
    End If
    
    '––– STEP 6: Build dictData from DatasetTable (one pass) –––
    Set dictData = BuildDictFromTable( _
                        lo:=loData, _
                        keyCol:="Fund Manager GCI", _
                        valCols:=Array("Family", "ECA India Analyst"), _
                        filterCol:="", _
                        filterVal:="" _
                   )
    
    '––– STEP 7: Clear existing data/filters from PortfolioTable –––
    On Error Resume Next
    loPort.Range.AutoFilter.ShowAllData
    On Error GoTo 0
    If Not loPort.DataBodyRange Is Nothing Then loPort.DataBodyRange.Delete
    
    '––– STEP 8: Performance switches –––
    With Application
        .ScreenUpdating = False
        .EnableEvents   = False
        .Calculation    = xlCalculationManual
    End With
    
    '––– STEP 9: Pre-allocate output array size –––
    If Not loTrig.DataBodyRange Is Nothing Then
        trigCount = loTrig.DataBodyRange.Rows.Count
    Else
        trigCount = 0
    End If
    If Not loNon.DataBodyRange Is Nothing Then
        nonCount = loNon.DataBodyRange.Rows.Count
    Else
        nonCount = 0
    End If
    capacity = trigCount + nonCount
    
    If capacity = 0 Then
        GoTo CleanAndAlert
    End If
    
    ReDim arrOut(1 To capacity, 1 To loPort.ListColumns.Count)
    Set colMap = BuildIndex(loPort)
    
    '––– STEP 10: Fill Trigger rows into arrOut –––
    ptr = FillArrayFast( _
            loSrc:=loTrig, _
            flag:="Trigger", _
            skipCol:="", skipVal:="", _
            dictAll:=dictAll, dictData:=dictData, _
            arrOut:=arrOut, startPtr:=0, colMap:=colMap _
          )
    
    '––– STEP 11: Fill Non-Trigger rows into arrOut (skip FI-ASIA) –––
    ptr = FillArrayFast( _
            loSrc:=loNon, _
            flag:="Non-Trigger", _
            skipCol:="Business Unit", skipVal:="FI-ASIA", _
            dictAll:=dictAll, dictData:=dictData, _
            arrOut:=arrOut, startPtr:=ptr, colMap:=colMap _
          )
    
    '––– STEP 12: Write arrOut back to PortfolioTable –––
    If ptr > 0 Then
        loPort.HeaderRowRange.Offset(1, 0).Resize(ptr, UBound(arrOut, 2)).Value = arrOut
        loPort.Resize loPort.HeaderRowRange.Resize(ptr + 1)
    End If
    
    '––– STEP 13: Remap Region codes –––
    With loPort
        With .ListColumns("Region").DataBodyRange
            .Replace What:="US",   Replacement:="AMRS", LookAt:=xlWhole
            .Replace What:="ASIA", Replacement:="APAC", LookAt:=xlWhole
        End With
    End With

CleanAndAlert:
    '––– STEP 14: Close the three helper workbooks –––
    On Error Resume Next
    wbTrig.Close SaveChanges:=False
    wbNon.Close  SaveChanges:=False
    wbAll.Close  SaveChanges:=False
    On Error GoTo 0
    
    '––– STEP 15: Restore application settings –––
    With Application
        .Calculation    = xlCalculationAutomatic
        .EnableEvents   = True
        .ScreenUpdating = True
        .StatusBar      = False
    End With
    
    '––– STEP 16: Compute elapsed time and show stats –––
    endTime = Timer
    elapsedSec = Round(endTime - startTime, 2)
    
    Dim msg As String
    msg = "PortfolioTable refresh complete!" & vbCrLf & _
          "Trigger rows processed:    " & trigCount & vbCrLf & _
          "Non-Trigger rows processed: " & nonCount & vbCrLf & _
          "Total rows loaded:         " & ptr & vbCrLf & _
          "Time taken (seconds):      " & elapsedSec
    MsgBox msg, vbInformation + vbOKOnly, "Refresh Complete"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Refresh Failed"
    On Error Resume Next
    wbTrig.Close SaveChanges:=False
    wbNon.Close  SaveChanges:=False
    wbAll.Close  SaveChanges:=False
    With Application
        .Calculation    = xlCalculationAutomatic
        .EnableEvents   = True
        .ScreenUpdating = True
        .StatusBar      = False
    End With
End Sub

'===========================
'  FILEPICKER HELPER
'===========================
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
Private Function EnsureTable(ws As Worksheet, deleteRow1 As Boolean, filterApproved As Boolean) As ListObject
    If deleteRow1 Then
        ws.Rows(1).Delete
    End If

    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0

    If ur Is Nothing Then
        ' Create a blank 1×1 table if nothing remains
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1"), , xlYes)
    Else
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ur, , xlYes)
    End If

    If filterApproved Then
        If ColumnExists(EnsureTable, "Review Status") Then
            With EnsureTable
                Dim idxReview As Long
                idxReview = .ListColumns("Review Status").Index
                With .Range
                    .AutoFilter Field:=idxReview, Criteria1:="<>Approved"
                    On Error Resume Next
                    .SpecialCells(xlCellTypeVisible).EntireRow.Delete
                    On Error GoTo 0
                    .AutoFilter
                End With
            End With
        End If
    End If
End Function

'===========================
'  BUILD LOOKUP DICTIONARY
'===========================
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

    Dim vData As Variant
    vData = lo.DataBodyRange.Value

    Dim idxSrc As Object: Set idxSrc = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        idxSrc(lo.ListColumns(c).Name) = c
    Next c

    If Not idxSrc.Exists(keyCol) Then
        Set BuildDictFromTable = dict: Exit Function
    End If
    Dim keyIdx As Long: keyIdx = idxSrc(keyCol)

    Dim filterIdx As Long: filterIdx = 0
    If filterCol <> "" Then
        If idxSrc.Exists(filterCol) Then filterIdx = idxSrc(filterCol)
    End If

    Dim valIdx() As Long, i As Long
    ReDim valIdx(0 To UBound(valCols))
    For i = 0 To UBound(valCols)
        If idxSrc.Exists(valCols(i)) Then
            valIdx(i) = idxSrc(valCols(i))
        Else
            valIdx(i) = 0
        End If
    Next i

    Dim r As Long, k As Variant, vArr() As Variant
    For r = 1 To UBound(vData, 1)
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
'  BUILD COLUMN → INDEX MAP
'===========================
Private Function BuildIndex(lo As ListObject) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        m(lo.ListColumns(c).Name) = c
    Next c
    Set BuildIndex = m
End Function

'===========================
'  FILL OUTPUT ARRAY FROM SOURCE
'===========================
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

    If loSrc.DataBodyRange Is Nothing Then
        FillArrayFast = startPtr
        Exit Function
    End If

    Dim vSrc As Variant
    vSrc = loSrc.DataBodyRange.Value

    Dim idxSrc As Object: Set idxSrc = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To loSrc.ListColumns.Count
        idxSrc(loSrc.ListColumns(c).Name) = c
    Next c

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

    Dim r As Long, outRow As Long
    Dim fGCI As Variant, fMgrGCI As Variant

    For r = 1 To UBound(vSrc, 1)
        If skipIdx > 0 Then
            If CStr(vSrc(r, skipIdx)) = skipVal Then GoTo NextRow
        End If

        startPtr = startPtr + 1
        outRow = startPtr

        fGCI = IIf(idx_FundGCI > 0, vSrc(r, idx_FundGCI), vbNullString)
        arrOut(outRow, colMap("Fund GCI")) = fGCI
        If idx_FundMgr > 0 Then  arrOut(outRow, colMap("Fund Manager"))    = vSrc(r, idx_FundMgr)
        If idx_FundName > 0 Then arrOut(outRow, colMap("Fund Name"))       = vSrc(r, idx_FundName)
        If idx_CreditOff > 0 Then arrOut(outRow, colMap("Credit Officer"))  = vSrc(r, idx_CreditOff)
        If idx_WCA > 0 Then       arrOut(outRow, colMap("WCA"))             = vSrc(r, idx_WCA)
        If idx_Region > 0 Then    arrOut(outRow, colMap("Region"))          = vSrc(r, idx_Region)

        If idx_WksMissing > 0 Then
            arrOut(outRow, colMap("Wks Missing")) = vSrc(r, idx_WksMissing)
        ElseIf idx_WeeksMissing > 0 Then
            arrOut(outRow, colMap("Wks Missing")) = vSrc(r, idx_WeeksMissing)
        End If

        If idx_LatestNAV > 0 Then arrOut(outRow, colMap("Latest NAV Date")) = vSrc(r, idx_LatestNAV)
        If idx_ReqNAV > 0 Then    arrOut(outRow, colMap("Req NAV Date"))   = vSrc(r, idx_ReqNAV)

        arrOut(outRow, colMap("Trigger/Non-Trigger")) = flag

        fMgrGCI = vbNullString
        If Not IsEmpty(fGCI) Then
            If dictAll.Exists(fGCI) Then
                fMgrGCI = dictAll(fGCI)(0)
                arrOut(outRow, colMap("Fund Manager GCI")) = dictAll(fGCI)(0)
                arrOut(outRow, colMap("Fund LEI"))         = dictAll(fGCI)(1)
                arrOut(outRow, colMap("Fund Code"))        = dictAll(fGCI)(2)
            End If
        End If

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