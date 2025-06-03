Option Explicit
'==================== MAIN PROCEDURE ====================
Sub Refresh_PortfolioTable()

    '––– let the user pick the three source files –––
    Dim fTrig As String, fNon As String, fAll As String
    fTrig = PickFile("Select TRIGGER file"):        If fTrig = "" Then Exit Sub
    fNon  = PickFile("Select NON-TRIGGER file"):    If fNon = "" Then Exit Sub
    fAll  = PickFile("Select ALL-FUNDS file"):      If fAll = "" Then Exit Sub

    '––– open sources (read-only) –––
    Dim wbTrig As Workbook: Set wbTrig = Workbooks.Open(fTrig, ReadOnly:=True)
    Dim wbNon  As Workbook: Set wbNon  = Workbooks.Open(fNon,  ReadOnly:=True)
    Dim wbAll  As Workbook: Set wbAll  = Workbooks.Open(fAll,  ReadOnly:=True)

    '––– convert each sheet to a table –––
    Dim loTrig As ListObject: Set loTrig = EnsureTable(wbTrig.Worksheets(1), False, False)
    Dim loNon  As ListObject: Set loNon  = EnsureTable(wbNon.Worksheets(1),  False, False)
    Dim loAll  As ListObject: Set loAll  = EnsureTable(wbAll.Worksheets(1),  True,  True)   'delete row-1, keep only Approved

    '––– tables in this workbook –––
    Dim loPort As ListObject: Set loPort = ThisWorkbook.Worksheets("Portfolio").ListObjects("PortfolioTable")
    Dim loData As ListObject: Set loData = ThisWorkbook.Worksheets("Dataset").ListObjects("DatasetTable")

    '––– dictionaries for lightning-fast look-ups –––
    Dim dictAll  As Object: Set dictAll  = BuildDict(loAll,  "Fund GCI",         Array("IA GCI", "Fund LEI", "Fund Code"))
    Dim dictData As Object: Set dictData = BuildDict(loData,"Fund Manager GCI", Array("Family", "ECA India Analyst"))

    '––– clear PortfolioTable (unfilter first) –––
    On Error Resume Next: loPort.Range.AutoFilter.ShowAllData: On Error GoTo 0
    If Not loPort.DataBodyRange Is Nothing Then loPort.DataBodyRange.Delete

    '––– speed switches –––
    With Application
        .ScreenUpdating = False: .EnableEvents = False: .Calculation = xlCalculationManual
    End With

    '––– build the output array –––
    Dim colMap As Object: Set colMap = BuildIndex(loPort)
    Dim capacity As Long: capacity = loTrig.DataBodyRange.Rows.Count + loNon.DataBodyRange.Rows.Count
    Dim arr(): ReDim arr(1 To capacity, 1 To loPort.ListColumns.Count)
    Dim ptr As Long

    ptr = FillArray(loTrig, "Trigger",      "",        "",       dictAll, dictData, arr, ptr, colMap)
    ptr = FillArray(loNon,  "Non-Trigger", "Business Unit", "FI-ASIA", dictAll, dictData, arr, ptr, colMap)

    '––– push the data to the worksheet –––
    If ptr > 0 Then
        loPort.DataBodyRange.Parent.Range(loPort.HeaderRowRange(1, 1)).Offset(1).Resize(ptr, UBound(arr, 2)).Value = arr
        loPort.Resize loPort.Range.Resize(ptr + 1)
    End If

    '––– region remap –––
    With loPort.ListColumns("Region").DataBodyRange
        .Replace "US",   "AMRS", xlWhole
        .Replace "ASIA", "APAC", xlWhole
    End With

CleanUp:
    With Application
        .Calculation = xlCalculationAutomatic: .EnableEvents = True: .ScreenUpdating = True
    End With
End Sub

'==================== HELPER FUNCTIONS ====================
Private Function PickFile(msg As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = msg: .AllowMultiSelect = False
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

'Create / convert a sheet into a proper table
'deleteRow1 – drop first row before conversion
'filterApproved – keep only “Approved” in Review Status
Private Function EnsureTable(ws As Worksheet, deleteRow1 As Boolean, filterApproved As Boolean) As ListObject
    If deleteRow1 Then ws.Rows(1).Delete

    Dim lastR As Long, lastC As Long
    With ws
        lastR = .Cells.Find("*", SearchOrder:=xlByRows,    SearchDirection:=xlPrevious).Row
        lastC = .Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End With

    Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)), , xlYes)

    'optional speed filter
    If filterApproved And ColumnExists(EnsureTable, "Review Status") Then
        With EnsureTable.Range
            .AutoFilter Field:=EnsureTable.ListColumns("Review Status").Index, Criteria1:="<>Approved"
            If EnsureTable.DataBodyRange.SpecialCells(xlCellTypeVisible).CountLarge > 1 Then _
                EnsureTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete xlShiftUp
            .AutoFilter             'clear filter
        End With
    End If
End Function

'Return dictionary keyed on keyCol, values = array(valCols)
Private Function BuildDict(lo As ListObject, keyCol As String, valCols As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k, i As Long, v()
    For r = 1 To lo.DataBodyRange.Rows.Count
        k = lo.DataBodyRange(r, lo.ListColumns(keyCol).Index).Value
        If Len(k) > 0 Then
            ReDim v(0 To UBound(valCols))
            For i = 0 To UBound(valCols)
                v(i) = lo.DataBodyRange(r, lo.ListColumns(valCols(i)).Index).Value
            Next i
            d(k) = v
        End If
    Next r
    Set BuildDict = d
End Function

'Column-name → index map for PortfolioTable
Private Function BuildIndex(lo As ListObject) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim c As Long: For c = 1 To lo.ListColumns.Count: m(lo.ListColumns(c).Name) = c: Next c
    Set BuildIndex = m
End Function

'Copy rows from loSrc into array, apply look-ups and flags
Private Function FillArray(loSrc As ListObject, flag As String, _
                            skipCol As String, skipVal As String, _
                            dictAll As Object, dictData As Object, _
                            arr, ptr As Long, idx As Object) As Long

    Dim r As Long, fGCI As String, fMgrGCI As String
    For r = 1 To loSrc.DataBodyRange.Rows.Count

        If skipCol <> "" Then _
            If loSrc.DataBodyRange(r, loSrc.ListColumns(skipCol).Index).Value = skipVal Then GoTo NextRow

        ptr = ptr + 1: fMgrGCI = ""

        fGCI = loSrc.DataBodyRange(r, loSrc.ListColumns("Fund GCI").Index).Value
        arr(ptr, idx("Fund GCI"))        = fGCI
        arr(ptr, idx("Fund Manager"))    = loSrc.DataBodyRange(r, loSrc.ListColumns("Fund Manager").Index).Value
        arr(ptr, idx("Fund Name"))       = loSrc.DataBodyRange(r, loSrc.ListColumns("Fund Name").Index).Value
        arr(ptr, idx("Credit Officer"))  = loSrc.DataBodyRange(r, loSrc.ListColumns("Credit Officer").Index).Value
        arr(ptr, idx("WCA"))             = loSrc.DataBodyRange(r, loSrc.ListColumns("WCA").Index).Value
        arr(ptr, idx("Region"))          = loSrc.DataBodyRange(r, loSrc.ListColumns("Region").Index).Value
        arr(ptr, idx("Wks Missing"))     = loSrc.DataBodyRange(r, loSrc.ListColumns(OptionalAlias(loSrc,"Wks Missing","Weeks Missing")).Index).Value
        arr(ptr, idx("Latest NAV Date")) = loSrc.DataBodyRange(r, loSrc.ListColumns("Latest NAV Date").Index).Value
        arr(ptr, idx("Req NAV Date"))    = loSrc.DataBodyRange(r, loSrc.ListColumns(OptionalAlias(loSrc,"Req NAV Date","Required NAV Date")).Index).Value
        arr(ptr, idx("Trigger/Non-Trigger")) = flag

        'All-Funds look-ups
        If dictAll.exists(fGCI) Then
            arr(ptr, idx("Fund Manager GCI")) = dictAll(fGCI)(0)
            arr(ptr, idx("Fund LEI"))         = dictAll(fGCI)(1)
            arr(ptr, idx("Fund Code"))        = dictAll(fGCI)(2)
            fMgrGCI = dictAll(fGCI)(0)
        End If

        'Dataset look-ups
        If Len(fMgrGCI) > 0 And dictData.exists(fMgrGCI) Then
            arr(ptr, idx("Family"))            = dictData(fMgrGCI)(0)
            arr(ptr, idx("ECA India Analyst")) = dictData(fMgrGCI)(1)
        End If
NextRow:
    Next r
    FillArray = ptr
End Function

'Return nm1 if column present, else nm2
Private Function OptionalAlias(lo As ListObject, nm1 As String, nm2 As String) As String
    OptionalAlias = IIf(ColumnExists(lo, nm1), nm1, nm2)
End Function

'True = column exists in table
Private Function ColumnExists(lo As ListObject, colName As String) As Boolean
    On Error Resume Next
    ColumnExists = (lo.ListColumns(colName).Index > 0)
    Err.Clear: On Error GoTo 0
End Function
'==========================================================