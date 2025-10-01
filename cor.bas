Option Explicit

' ===== Entry Point =====
Public Sub Run_ApprovedFunds_CreditStudio_Workflow()
    On Error GoTo Fail

    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    Dim wsDate As Worksheet, wsRecali As Worksheet
    Dim approvedPath As String, creditPath As String
    Dim wbApproved As Workbook, wbCredit As Workbook
    Dim loApproved As ListObject, loCredit As ListObject
    Dim joinedCoper As String

    ' 1) Ask user to provide approved funds CSV
    approvedPath = PickFile("Select APPROVED FUNDS CSV", "CSV Files (*.csv),*.csv")
    If Len(approvedPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: Exit Sub

    ' 2) Open CSV and delete first row (headers are on row 2)
    Set wbApproved = Workbooks.Open(Filename:=approvedPath, Local:=True)
    With wbApproved.Sheets(1)
        .Rows(1).Delete
    End With

    ' 3) Convert data to table with headers
    Set loApproved = EnsureTable(wbApproved.Sheets(1), "ApprovedTbl")

    ' 4) Keep only Business Unit in {FI-GMC-ASIA, FI-US, FI-EMEA}
    FilterKeepOnlyBusinessUnits loApproved, Array("FI-GMC-ASIA", "FI-US", "FI-EMEA")

    ' 5) Join Fund CoPER values, copy to clipboard, show copy-again / move-on loop
    joinedCoper = JoinColumnValues(loApproved, "Fund CoPER", ",")
    If Len(joinedCoper) = 0 Then
        MsgBox "No 'Fund CoPER' values found after filtering. Cannot proceed.", vbCritical
        GoTo Cleanup
    End If

    CopyToClipboard joinedCoper
    Do
        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
            "Fund CoPER values have been copied for Credit Studio." & vbCrLf & vbCrLf & _
            "• Click YES to copy the same again (if you got distracted)." & vbCrLf & _
            "• Click NO to MOVE ON (continue the process).", _
            vbYesNo + vbInformation, "Copied for Credit Studio")
        If resp = vbYes Then
            CopyToClipboard joinedCoper
        Else
            Exit Do    ' Move On
        End If
    Loop

    ' 5) Ask user to upload the Credit Studio xlsx file
    creditPath = PickFile("Select CREDIT STUDIO XLSX", "Excel Files (*.xlsx),*.xlsx")
    If Len(creditPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: GoTo Cleanup
    Set wbCredit = Workbooks.Open(Filename:=creditPath, ReadOnly:=True)

    ' 6) Create a new sheet in MAIN file named as today's date
    Set wsDate = CreateDatedSheet(wbMain)

    ' 7) From credit studio, convert to table and copy "Coper ID" & "Country of Risk" to new sheet "CoR Recali"
    Set loCredit = EnsureTable(wbCredit.Sheets(1), "CreditTbl")
    Set wsRecali = EnsureSheet(wbMain, "CoR Recali", True) ' recreate fresh
    CopyColumnsByName loCredit, wsRecali, Array("Coper ID", "Country of Risk")

    ' 8) Lookup Approved CoR by Coper (Approved: Fund CoPER -> Country of Risk)
    Dim approvedMap As Object: Set approvedMap = BuildCoperToCoRMap(loApproved, "Fund CoPER", "Country of Risk")
    AppendApprovedCoR wsRecali, approvedMap, "Coper ID", "Approved CoR"

    ' 9) Convert "CoR Recali" to table, compare CoR vs Approved CoR,
    '    and create mismatch summary: unique mismatched CoR (from Credit Studio) with joined Copers
    Dim loRecali As ListObject
    Set loRecali = EnsureTable(wsRecali, "CoRRecaliTbl")

    CreateMismatchSummary wbMain, loRecali, _
        creditCoRColName:="Country of Risk", _
        approvedCoRColName:="Approved CoR", _
        coperColName:="Coper ID", _
        summarySheetName:="CoR Mismatch Summary"

    MsgBox "Done. Sheets created: '" & wsDate.Name & "', 'CoR Recali', and (if mismatches) 'CoR Mismatch Summary'.", vbInformation
    GoTo Cleanup

Fail:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical

Cleanup:
    On Error Resume Next
    If Not wbCredit Is Nothing Then wbCredit.Close SaveChanges:=False
    If Not wbApproved Is Nothing Then wbApproved.Close SaveChanges:=True
End Sub

' ===== Helpers =====

Private Function PickFile(promptTitle As String, filterStr As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = promptTitle
        .Filters.Clear
        .Filters.Add Split(filterStr, "|")(0), Split(filterStr, "|")(0)
        ' Excel expects "Description,Pattern" format:
        ' But FileDialog uses .Filters.Add "Excel Files","*.xlsx"
        ' So we parse "Excel Files (*.xlsx),*.xlsx" if passed
        Dim parts() As String: parts = Split(filterStr, ",")
        If UBound(parts) >= 1 Then
            .Filters.Clear
            .Filters.Add Trim(Replace(parts(0), "|", "")), Trim(parts(1))
        End If
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFile = .SelectedItems(1)
        Else
            PickFile = ""
        End If
    End With
End Function

Private Function EnsureTable(ws As Worksheet, tableName As String) As ListObject
    ' Create a ListObject from used range; if one already exists, return it.
    Dim lo As ListObject
    If ws.ListObjects.Count > 0 Then
        Set EnsureTable = ws.ListObjects(1)
        Exit Function
    End If

    Dim rng As Range
    Set rng = ws.UsedRange
    If rng Is Nothing Then Err.Raise 1001, , "No data found on sheet '" & ws.Name & "'."

    ' Clean: remove fully blank first/last rows/cols often seen after CSVs
    Set rng = TrimUsedRange(ws)

    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    On Error Resume Next
    lo.Name = tableName
    On Error GoTo 0
    Set EnsureTable = lo
End Function

Private Function TrimUsedRange(ws As Worksheet) As Range
    Dim ur As Range: Set ur = ws.UsedRange
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    r1 = ur.Row: r2 = ur.Rows(ur.Rows.Count).Row
    c1 = ur.Column: c2 = ur.Columns(ur.Columns.Count).Column

    ' Move r1 down while entire row is blank
    Do While r1 <= r2 And Application.CountA(ws.Rows(r1)) = 0: r1 = r1 + 1: Loop
    ' Move r2 up while entire row is blank
    Do While r2 >= r1 And Application.CountA(ws.Rows(r2)) = 0: r2 = r2 - 1: Loop
    ' Move c1 right while entire column is blank
    Do While c1 <= c2 And Application.CountA(ws.Columns(c1)) = 0: c1 = c1 + 1: Loop
    ' Move c2 left while entire column is blank
    Do While c2 >= c1 And Application.CountA(ws.Columns(c2)) = 0: c2 = c2 - 1: Loop

    If r2 < r1 Or c2 < c1 Then
        Set TrimUsedRange = ws.Range("A1")
    Else
        Set TrimUsedRange = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    End If
End Function

Private Sub FilterKeepOnlyBusinessUnits(lo As ListObject, keepArr As Variant)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim buCol As Long: buCol = GetColumnIndex(lo, "Business Unit")
    If buCol = 0 Then Err.Raise 1002, , "Column 'Business Unit' not found."

    ' Build a criteria array to show only the kept values
    lo.Range.AutoFilter Field:=buCol, Criteria1:=keepArr, Operator:=xlFilterValues

    ' Delete visible rows NOT matching keep list:
    ' Simpler: show Disallowed and delete; but AutoFilter can’t do NOT IN easily.
    ' So: Copy visible (kept) to a new temp range, then replace table.
    Dim tmp As Worksheet: Set tmp = lo.Parent.Parent.Worksheets.Add(After:=ws)
    lo.Range.SpecialCells(xlCellTypeVisible).Copy tmp.Range("A1")

    ' Replace original sheet content with tmp content
    ws.Cells.Clear
    tmp.UsedRange.Copy ws.Range("A1")
    Application.DisplayAlerts = False
    tmp.Delete
    Application.DisplayAlerts = True

    ' Recreate table
    ws.Cells.RemoveDuplicates Columns:=Evaluate("ROW(1:" & ws.UsedRange.Columns.Count & ")"), Header:=xlYes
    lo.Parent.ListObjects.Delete ' in case old pointer lingers
    Set lo = Nothing
    Set lo = EnsureTable(ws, "ApprovedTbl")
End Sub

Private Function JoinColumnValues(lo As ListObject, headerName As String, delim As String) As String
    Dim idx As Long: idx = GetColumnIndex(lo, headerName)
    If idx = 0 Then Err.Raise 1003, , "Column '" & headerName & "' not found."

    Dim arr, i As Long, s As String
    arr = lo.DataBodyRange.Columns(idx).Value
    For i = 1 To UBound(arr, 1)
        If Len(Trim$(arr(i, 1))) > 0 Then
            If Len(s) > 0 Then s = s & delim
            s = s & Trim$(CStr(arr(i, 1)))
        End If
    Next i
    JoinColumnValues = s
End Function

Private Sub CopyToClipboard(textVal As String)
    On Error GoTo fallback
    Dim o As Object
    Set o = CreateObject("MSForms.DataObject")
    o.SetText textVal
    o.PutInClipboard
    Exit Sub
fallback:
    ' Fallback using a hidden Temp sheet + Copy (works even without MSForms)
    Dim ws As Worksheet, prevScr As Boolean
    prevScr = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Visible = xlSheetVeryHidden
    ws.Range("A1").Value = textVal
    ws.Range("A1").Copy
    Application.CutCopyMode = False
    ws.Delete
    Application.ScreenUpdating = prevScr
End Sub

Private Function CreateDatedSheet(wb As Workbook) As Worksheet
    Dim baseName As String: baseName = Format(Date, "yyyy-mm-dd")
    Dim nameCandidate As String: nameCandidate = baseName
    Dim n As Long: n = 1
    Do While SheetExists(wb, nameCandidate)
        n = n + 1
        nameCandidate = baseName & " (" & n & ")"
    Loop
    Set CreateDatedSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    CreateDatedSheet.Name = nameCandidate
End Function

Private Function EnsureSheet(wb As Workbook, name As String, Optional recreate As Boolean = False) As Worksheet
    If recreate And SheetExists(wb, name) Then
        Application.DisplayAlerts = False
        wb.Worksheets(name).Delete
        Application.DisplayAlerts = True
    End If
    If SheetExists(wb, name) Then
        Set EnsureSheet = wb.Worksheets(name)
    Else
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function SheetExists(wb As Workbook, name As String) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

Private Sub CopyColumnsByName(lo As ListObject, wsDest As Worksheet, fieldNames As Variant)
    Dim i As Long, idx As Long, nextCol As Long: nextCol = 1
    wsDest.Cells.Clear
    ' headers
    For i = LBound(fieldNames) To UBound(fieldNames)
        idx = GetColumnIndex(lo, CStr(fieldNames(i)))
        If idx = 0 Then Err.Raise 1004, , "Column '" & CStr(fieldNames(i)) & "' not found in Credit Studio."
        wsDest.Cells(1, nextCol).Value = CStr(fieldNames(i))
        lo.DataBodyRange.Columns(idx).Copy wsDest.Cells(2, nextCol)
        nextCol = nextCol + 1
    Next i
End Sub

Private Function GetColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim$(LCase$(lo.HeaderRowRange.Cells(1, i).Value)) = Trim$(LCase$(headerName)) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Private Function BuildCoperToCoRMap(lo As ListObject, coperCol As String, corCol As String) As Object
    Dim idxC As Long, idxR As Long
    idxC = GetColumnIndex(lo, coperCol)
    idxR = GetColumnIndex(lo, corCol)
    If idxC = 0 Then Err.Raise 1005, , "Column '" & coperCol & "' not found in Approved."
    If idxR = 0 Then Err.Raise 1006, , "Column '" & corCol & "' not found in Approved."

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, vCoper As String, vCoR As String
    For r = 1 To lo.DataBodyRange.Rows.Count
        vCoper = Trim$(CStr(lo.DataBodyRange.Cells(r, idxC).Value))
        vCoR = Trim$(CStr(lo.DataBodyRange.Cells(r, idxR).Value))
        If Len(vCoper) > 0 Then
            If Not dict.Exists(vCoper) Then dict.Add vCoper, vCoR Else dict(vCoper) = vCoR
        End If
    Next r
    Set BuildCoperToCoRMap = dict
End Function

Private Sub AppendApprovedCoR(ws As Worksheet, approvedMap As Object, coperColName As String, outColName As String)
    Dim lastCol As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim coperCol As Long: coperCol = FindHeader(ws, coperColName)
    If coperCol = 0 Then Err.Raise 1007, , "Column '" & coperColName & "' not found on " & ws.Name

    ws.Cells(1, lastCol + 1).Value = outColName
    Dim r As Long, key As String
    For r = 2 To lastRow
        key = Trim$(CStr(ws.Cells(r, coperCol).Value))
        If Len(key) > 0 And approvedMap.Exists(key) Then
            ws.Cells(r, lastCol + 1).Value = approvedMap(key)
        Else
            ws.Cells(r, lastCol + 1).Value = "" ' not found
        End If
    Next r
End Sub

Private Function FindHeader(ws As Worksheet, headerName As String) As Long
    Dim c As Range
    For Each c In ws.Rows(1).Cells
        If Len(CStr(c.Value)) = 0 Then Exit For
        If Trim$(LCase$(c.Value)) = Trim$(LCase$(headerName)) Then
            FindHeader = c.Column
            Exit Function
        End If
    Next c
    FindHeader = 0
End Function

Private Sub CreateMismatchSummary( _
    wb As Workbook, _
    loRecali As ListObject, _
    ByVal creditCoRColName As String, _
    ByVal approvedCoRColName As String, _
    ByVal coperColName As String, _
    ByVal summarySheetName As String)

    Dim idxCredit As Long, idxApproved As Long, idxCoper As Long
    idxCredit = GetColumnIndex(loRecali, creditCoRColName)
    idxApproved = GetColumnIndex(loRecali, approvedCoRColName)
    idxCoper = GetColumnIndex(loRecali, coperColName)

    If idxCredit * idxApproved * idxCoper = 0 Then
        Err.Raise 1008, , "One or more required columns missing in CoR Recali."
    End If

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, valCredit As String, valApproved As String, coper As String

    For r = 1 To loRecali.DataBodyRange.Rows.Count
        valCredit = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxCredit).Value))
        valApproved = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxApproved).Value))
        coper = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxCoper).Value))

        If Len(valCredit) > 0 And Len(coper) > 0 Then
            If StrComp(valCredit, valApproved, vbTextCompare) <> 0 Then
                ' mismatch: group by CREDIT STUDIO CoR
                If Not dict.Exists(valCredit) Then
                    dict.Add valCredit, New Collection
                End If
                On Error Resume Next
                dict(valCredit).Add coper ' allow duplicates prevention below
                On Error GoTo 0
            End If
        End If
    Next r

    ' If no mismatches, exit gracefully
    If dict.Count = 0 Then
        If SheetExists(wb, summarySheetName) Then
            Application.DisplayAlerts = False
            wb.Worksheets(summarySheetName).Delete
            Application.DisplayAlerts = True
        End If
        Exit Sub
    End If

    Dim wsSum As Worksheet
    Set wsSum = EnsureSheet(wb, summarySheetName, True)
    wsSum.Range("A1").Value = "Country of Risk (Credit Studio)"
    wsSum.Range("B1").Value = "Coper IDs (mismatched; comma-joined)"

    Dim k As Variant, rowOut As Long: rowOut = 2
    For Each k In dict.Keys
        Dim joined As String
        joined = JoinUniqueFromCollection(dict(k), ",")
        wsSum.Cells(rowOut, 1).Value = k
        wsSum.Cells(rowOut, 2).Value = joined
        rowOut = rowOut + 1
    Next k

    ' Format as table
    Dim lo As ListObject
    Set lo = wsSum.ListObjects.Add(xlSrcRange, wsSum.Range("A1").CurrentRegion, , xlYes)
    On Error Resume Next
    lo.Name = "CoRMismatchSummaryTbl"
    On Error GoTo 0

    wsSum.Columns.AutoFit
End Sub

Private Function JoinUniqueFromCollection(col As Collection, delim As String) As String
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, val As String, s As String
    For i = 1 To col.Count
        val = Trim$(CStr(col(i)))
        If Len(val) > 0 Then
            If Not seen.Exists(val) Then
                seen.Add val, True
                If Len(s) > 0 Then s = s & delim
                s = s & val
            End If
        End If
    Next i
    JoinUniqueFromCollection = s
End Function