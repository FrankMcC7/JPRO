Option Explicit
'=========================== MASTER ROUTINE ===========================
Sub Refresh_PortfolioTable()

    '--- pick files
    Dim fTrig As String, fNon As String, fAll As String
    fTrig = PickFile("Select TRIGGER file"): If fTrig = "" Then Exit Sub
    fNon  = PickFile("Select NON-TRIGGER file"): If fNon = "" Then Exit Sub
    fAll  = PickFile("Select ALL-FUNDS file"):  If fAll = "" Then Exit Sub

    '--- open sources
    Dim wbTrig As Workbook: Set wbTrig = Workbooks.Open(fTrig, ReadOnly:=True)
    Dim wbNon  As Workbook: Set wbNon  = Workbooks.Open(fNon,  ReadOnly:=True)
    Dim wbAll  As Workbook: Set wbAll  = Workbooks.Open(fAll,  ReadOnly:=True)

    '--- convert to tables
    Dim loTrig As ListObject: Set loTrig = EnsureTable(wbTrig.Worksheets(1), False)
    Dim loNon  As ListObject: Set loNon  = EnsureTable(wbNon.Worksheets(1),  False)
    Dim loAll  As ListObject: Set loAll  = EnsureTable(wbAll.Worksheets(1),  True)   'delete row 1

    '--- target & dataset tables in this workbook
    Dim wsPort As Worksheet: Set wsPort = ThisWorkbook.Worksheets("Portfolio")
    Dim loPort As ListObject: Set loPort = wsPort.ListObjects("PortfolioTable")
    Dim loData As ListObject: Set loData = ThisWorkbook.Worksheets("Dataset").ListObjects("DatasetTable")

    '--- dictionaries for rapid look-ups
    Dim dictAll  As Object: Set dictAll  = BuildDict(loAll,  "Fund GCI",         Array("IA GCI","Fund LEI","Fund Code"))
    Dim dictData As Object: Set dictData = BuildDict(loData,"Fund Manager GCI", Array("Family","ECA India Analyst"))

    '--- clear PortfolioTable (keep headers, remove filters)
    On Error Resume Next: loPort.Range.AutoFilter.ShowAllData: On Error GoTo 0
    If Not loPort.DataBodyRange Is Nothing Then loPort.DataBodyRange.Delete

    '--- performance switches
    With Application: .ScreenUpdating = False: .EnableEvents = False: .Calculation = xlCalculationManual: End With

    '--- build output array (rows × columns)
    Dim colMap As Object: Set colMap = BuildIndex(loPort)
    Dim rowMax As Long: rowMax = loTrig.DataBodyRange.Rows.Count + loNon.DataBodyRange.Rows.Count
    Dim arr(): ReDim arr(1 To rowMax, 1 To loPort.ListColumns.Count)
    Dim ptr As Long

    ptr = FillArray(loTrig,"Trigger", "",        "",       dictAll,dictData,arr,ptr,colMap)
    ptr = FillArray(loNon, "Non-Trigger","Business Unit","FI-ASIA",dictAll,dictData,arr,ptr,colMap)

    '--- write data then resize table
    If ptr > 0 Then
        loPort.HeaderRowRange.Offset(1).Resize(ptr, UBound(arr, 2)).Value = arr
        loPort.Resize loPort.Range.Resize(ptr + 1)
    End If

    '--- region mapping
    With loPort.ListColumns("Region").DataBodyRange
        .Replace "US",   "AMRS", xlWhole
        .Replace "ASIA", "APAC", xlWhole
    End With

CleanUp:
    With Application: .Calculation = xlCalculationAutomatic: .EnableEvents = True: .ScreenUpdating = True: End With
End Sub
'=========================== SUPPORT ===========================
Private Function PickFile(msg As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = msg: .AllowMultiSelect = False
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Function EnsureTable(ws As Worksheet, deleteRow1 As Boolean) As ListObject
    If deleteRow1 Then ws.Rows(1).Delete
    If ws.ListObjects.Count = 0 Then
        Dim lr As Long, lc As Long
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)), , xlYes)
    Else
        Set EnsureTable = ws.ListObjects(1)
    End If
End Function

Private Function BuildDict(lo As ListObject, keyCol As String, valCols As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k
    For r = 1 To lo.DataBodyRange.Rows.Count
        k = lo.DataBodyRange(r, lo.ListColumns(keyCol).Index).Value
        If Len(k) > 0 Then
            Dim v(): ReDim v(0 To UBound(valCols))
            Dim i As Long: For i = 0 To UBound(valCols)
                v(i) = lo.DataBodyRange(r, lo.ListColumns(valCols(i)).Index).Value
            Next i
            d(k) = v
        End If
    Next r
    Set BuildDict = d
End Function

Private Function BuildIndex(lo As ListObject) As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    Dim c As Long: For c = 1 To lo.ListColumns.Count: m(lo.ListColumns(c).Name) = c: Next c
    Set BuildIndex = m
End Function

Private Function FillArray(loSrc As ListObject, flag As String, _
                            skipCol As String, skipVal As String, _
                            dictAll As Object, dictData As Object, _
                            arr, ptr As Long, idx As Object) As Long

    If loSrc.DataBodyRange Is Nothing Then FillArray = ptr: Exit Function  'no data2

    Dim r As Long, fGCI As String, fMgrGCI As String
    For r = 1 To loSrc.DataBodyRange.Rows.Count
        If skipCol <> "" Then
            If loSrc.DataBodyRange(r, loSrc.ListColumns(skipCol).Index).Value = skipVal Then GoTo NextRow
        End If

        ptr = ptr + 1
        fGCI = loSrc.DataBodyRange(r, loSrc.ListColumns("Fund GCI").Index).Value

        arr(ptr, idx("Fund GCI"))    = fGCI
        arr(ptr, idx("Fund Manager"))= loSrc.DataBodyRange(r, loSrc.ListColumns("Fund Manager").Index).Value
        arr(ptr, idx("Fund Name"))   = loSrc.DataBodyRange(r, loSrc.ListColumns("Fund Name").Index).Value
        arr(ptr, idx("Credit Officer")) = loSrc.DataBodyRange(r, loSrc.ListColumns("Credit Officer").Index).Value
        arr(ptr, idx("WCA"))         = loSrc.DataBodyRange(r, loSrc.ListColumns("WCA").Index).Value
        arr(ptr, idx("Region"))      = loSrc.DataBodyRange(r, loSrc.ListColumns("Region").Index).Value
        arr(ptr, idx("Wks Missing")) = loSrc.DataBodyRange(r, loSrc.ListColumns(OptionalAlias(loSrc,"Wks Missing","Weeks Missing")).Index).Value
        arr(ptr, idx("Latest NAV Date")) = loSrc.DataBodyRange(r, loSrc.ListColumns("Latest NAV Date").Index).Value
        arr(ptr, idx("Req NAV Date"))= loSrc.DataBodyRange(r, loSrc.ListColumns(OptionalAlias(loSrc,"Req NAV Date","Required NAV Date")).Index).Value
        arr(ptr, idx("Trigger/Non-Trigger")) = flag

        If dictAll.exists(fGCI) Then
            arr(ptr, idx("Fund Manager GCI")) = dictAll(fGCI)(0)
            arr(ptr, idx("Fund LEI"))         = dictAll(fGCI)(1)
            arr(ptr, idx("Fund Code"))        = dictAll(fGCI)(2)
            fMgrGCI = dictAll(fGCI)(0)
        End If

        If Len(fMgrGCI) > 0 And dictData.exists(fMgrGCI) Then
            arr(ptr, idx("Family"))            = dictData(fMgrGCI)(0)
            arr(ptr, idx("ECA India Analyst")) = dictData(fMgrGCI)(1)
        End If
NextRow:
    Next r
    FillArray = ptr
End Function

Private Function OptionalAlias(lo As ListObject, nm1 As String, nm2 As String) As String
    OptionalAlias = IIf(lo.ListColumns.Contains(nm1), nm1, nm2)
End Function
'======================================================================