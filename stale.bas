Option Explicit

Sub Refresh_PortfolioTable()
    Dim fTrig As String, fNon As String, fAll As String
    Dim wbTrig As Workbook, wbNon As Workbook, wbAll As Workbook
    Dim loPort As ListObject, loTrig As ListObject, loNon As ListObject, loAll As ListObject
    Dim wsPort As Worksheet, wsData As Worksheet
    Dim dictAll As Object, arrOut(), r As Long, c As Long, n As Long
    Dim hdrs As Variant, idx As Object
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '––- pick files ––-
    fTrig = GetPath("Select TRIGGER file")          : If fTrig = "" Then GoTo tidy
    fNon  = GetPath("Select NON-TRIGGER file")      : If fNon = "" Then GoTo tidy
    fAll  = GetPath("Select ALL-FUNDS file")        : If fAll = "" Then GoTo tidy
    
    Set wbTrig = Workbooks.Open(fTrig, ReadOnly:=True)
    Set wbNon  = Workbooks.Open(fNon,  ReadOnly:=True)
    Set wbAll  = Workbooks.Open(fAll,  ReadOnly:=True)
    
    Set wsPort = ThisWorkbook.Worksheets("Portfolio")
    Set loPort = wsPort.ListObjects("PortfolioTable")
    
    'clear filters & old data
    On Error Resume Next
    loPort.Range.AutoFilter.ShowAllData
    On Error GoTo 0
    If Not loPort.DataBodyRange Is Nothing Then loPort.DataBodyRange.Delete
    
    '--- build All-Funds dictionary ---
    Set loAll = ConvertToTable(wbAll.Worksheets(1), True) 'deletes first row
    Set dictAll = CreateObject("scripting.dictionary")
    For r = 1 To loAll.DataBodyRange.Rows.Count
        dictAll(loAll.DataBodyRange(r, loAll.ListColumns("Fund GCI").Index).Value) = _
            Array( _
                loAll.DataBodyRange(r, loAll.ListColumns("IA GCI").Index).Value, _
                loAll.DataBodyRange(r, loAll.ListColumns("Fund LEI").Index).Value, _
                loAll.DataBodyRange(r, loAll.ListColumns("Fund Code").Index).Value)
    Next r
    
    '--- headings we need, in order ---
    hdrs = Array("Fund GCI", "Fund Manager", "Fund Name", "Credit Officer", "WCA", _
                 "Region", "Wks Missing", "Latest NAV Date", "Req NAV Date", "Trigger/Non-Trigger", _
                 "Fund Manager GCI", "Fund LEI", "Fund Code")
    
    Set idx = CreateObject("scripting.dictionary")
    For c = 1 To loPort.ListColumns.Count: idx(loPort.ListColumns(c).Name) = c: Next
    
    '--- append Trigger rows ---
    n = LoadRows(loTrig, "Trigger", loPort, hdrs, idx, dictAll)
    '--- append Non-Trigger rows (skip FI-ASIA) ---
    n = n + LoadRows(loNon, "Non-Trigger", loPort, hdrs, idx, dictAll, "Business Unit", "FI-ASIA")
    
tidy:
    'region mapping
    With loPort.ListColumns(idx("Region")).DataBodyRange
        .Replace "US", "AMRS", xlWhole
        .Replace "ASIA", "APAC", xlWhole
    End With
    
    'restore
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function GetPath(msg As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = msg: .AllowMultiSelect = False
        If .Show = -1 Then GetPath = .SelectedItems(1)
    End With
End Function

Private Function ConvertToTable(ws As Worksheet, Optional deleteRow1 As Boolean = False) As ListObject
    If deleteRow1 Then ws.Rows(1).Delete
    Dim lastR As Long, lastC As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set ConvertToTable = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)), , xlYes)
End Function

Private Function LoadRows(loSrc As ListObject, typ As String, _
                           loDst As ListObject, hdrs, idx, dictAll, _
                           Optional critCol As String = "", Optional critVal As String = "") As Long
    Dim arr(), arrOut(), r As Long, c As Long, p As Long, v
    arr = loSrc.DataBodyRange.Value
    ReDim arrOut(1 To UBound(arr), 1 To UBound(hdrs) + 1)
    
    For r = 1 To UBound(arr)
        If critCol = "" Or loSrc.DataBodyRange(r, loSrc.ListColumns(critCol).Index).Value <> critVal Then
            For p = 0 To UBound(hdrs) - 4 'first 9 direct columns
                v = loSrc.DataBodyRange(r, loSrc.ListColumns(hdrs(p)).Index).Value
                arrOut(r, idx(hdrs(p))) = IIf(hdrs(p) = "Region", MapRegion(v), v)
            Next p
            'trigger/non-trigger flag
            arrOut(r, idx("Trigger/Non-Trigger")) = typ
            'dictionary look-ups
            If dictAll.exists(arrOut(r, idx("Fund GCI"))) Then
                arrOut(r, idx("Fund Manager GCI")) = dictAll(arrOut(r, idx("Fund GCI")))(0)
                arrOut(r, idx("Fund LEI")) = dictAll(arrOut(r, idx("Fund GCI")))(1)
                arrOut(r, idx("Fund Code")) = dictAll(arrOut(r, idx("Fund GCI")))(2)
            End If
            LoadRows = LoadRows + 1
        End If
    Next r
    'dump to table
    loDst.ListRows.Add AlwaysInsert:=True
    loDst.DataBodyRange.Rows(loDst.DataBodyRange.Rows.Count).Resize(LoadRows).Value = arrOut
End Function

Private Function MapRegion(rg As Variant) As String
    Select Case UCase(rg)
        Case "US": MapRegion = "AMRS"
        Case "ASIA": MapRegion = "APAC"
        Case Else: MapRegion = rg
    End Select
End Function