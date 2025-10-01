Option Explicit

' =======================
' Entry Point
' =======================
Public Sub Run_ApprovedFunds_CreditStudio_Workflow()
    On Error GoTo Fail

    Dim wbMain As Workbook: Set wbMain = ThisWorkbook

    ' NEW: refuse to run if workbook structure is protected (cannot add/delete sheets)
    If IsStructureProtected(wbMain) Then
        MsgBox "This workbook's structure is protected. Unprotect it (Review → Protect Workbook) and run again.", vbCritical
        Exit Sub
    End If

    Dim wsDate As Worksheet, wsRecali As Worksheet
    Dim approvedPath As String, creditPath As String
    Dim wbApproved As Workbook, wbCredit As Workbook
    Dim loApproved As ListObject, loCredit As ListObject
    Dim joinedCoper As String
    Dim resp As VbMsgBoxResult

    ' 1) Ask user to provide approved funds CSV
    approvedPath = PickFile("Select APPROVED FUNDS CSV", "CSV Files (*.csv)", "*.csv")
    If Len(approvedPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: Exit Sub

    ' 2) Open CSV and delete first row (headers are on row 2)
    Set wbApproved = Workbooks.Open(Filename:=approvedPath, Local:=True)
    With wbApproved.Worksheets(1)
        .Rows(1).Delete
    End With

    ' 3) Convert data to table with headers
    Set loApproved = EnsureTable(wbApproved.Worksheets(1), "ApprovedTbl")

    ' 4) Keep only Business Unit in {FI-GMC-ASIA, FI-US, FI-EMEA}
    FilterKeepOnlyBusinessUnits loApproved, Array("FI-GMC-ASIA", "FI-US", "FI-EMEA")
    ' loApproved refreshed inside the sub

    ' 5) Join Fund CoPER values, copy to clipboard, show copy-again / move-on loop
    joinedCoper = JoinColumnValues(loApproved, "Fund CoPER", ",")
    If Len(joinedCoper) = 0 Then
        MsgBox "No 'Fund CoPER' values found after filtering. Cannot proceed.", vbCritical
        GoTo Cleanup
    End If

    CopyToClipboard joinedCoper
    Do
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

    ' 5b) Ask user to upload the Credit Studio xlsx file
    creditPath = PickFile("Select CREDIT STUDIO XLSX", "Excel Files (*.xlsx)", "*.xlsx")
    If Len(creditPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: GoTo Cleanup

    Set wbCredit = Workbooks.Open(Filename:=creditPath, ReadOnly:=True)

    ' 6) Create a new sheet in MAIN file named as today's date
    Set wsDate = CreateDatedSheet(wbMain)

    ' 7) From credit studio, convert to table and copy "Coper ID" & "Country of Risk" to new sheet "CoR Recali"
    Set loCredit = EnsureTable(wbCredit.Worksheets(1), "CreditTbl")
    Set wsRecali = EnsureSheet(wbMain, "CoR Recali", True) ' recreate fresh
    CopyColumnsByName loCredit, wsRecali, Array("Coper ID", "Country of Risk")

    ' 8) Lookup Approved CoR by Coper (Approved: Fund CoPER -> Country of Risk)
    Dim approvedMap As Object
    Set approvedMap = BuildCoperToCoRMap(loApproved, "Fund CoPER", "Country of Risk")
    AppendApprovedCoR wsRecali, approvedMap, "Coper ID", "Approved CoR"

    ' 9) Make "CoR Recali" a table; compare CoR vs Approved CoR; create mismatch summary
    Dim loRecali As ListObject
    Set loRecali = EnsureTable(wsRecali, "CoRRecaliTbl")

    CreateMismatchSummary ThisWorkbook, loRecali, _
                          creditCoRColName:="Country of Risk", _
                          approvedCoRColName:="Approved CoR", _
                          coperColName:="Coper ID", _
                          summarySheetName:="CoR Mismatch Summary"

    MsgBox "Done. Sheets created: '" & wsDate.Name & "', 'CoR Recali', and (if mismatches) 'CoR Mismatch Summary'.", vbInformation
    GoTo Cleanup

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical

Cleanup:
    On Error Resume Next
    If Not wbCredit Is Nothing Then wbCredit.Close SaveChanges:=False
    If Not wbApproved Is Nothing Then wbApproved.Close SaveChanges:=True
End Sub

' =======================
' Helpers
' =======================

Private Function PickFile(ByVal promptTitle As String, _
                          ByVal filterDesc As String, _
                          ByVal filterPattern As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = promptTitle
        .Filters.Clear
        .Filters.Add filterDesc, filterPattern
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFile = .SelectedItems(1)
        Else
            PickFile = ""
        End If
    End With
    Set fd = Nothing
End Function

Private Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject
    Dim rng As Range

    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        On Error Resume Next
        lo.Name = tableName
        On Error GoTo 0
        Set EnsureTable = lo
        Exit Function
    End If

    Set rng = TrimUsedRange(ws)
    If rng Is Nothing Then Err.Raise 1001, , "No data found on sheet '" & ws.Name & "'."

    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes)
    On Error Resume Next
    lo.Name = tableName
    On Error GoTo 0

    Set EnsureTable = lo
End Function

Private Function TrimUsedRange(ByVal ws As Worksheet) As Range
    Dim ur As Range
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long

    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    r1 = ur.Row: r2 = ur.Rows(ur.Rows.Count).Row
    c1 = ur.Column: c2 = ur.Columns(ur.Columns.Count).Column

    Do While r1 <= r2 And Application.CountA(ws.Rows(r1)) = 0: r1 = r1 + 1: Loop
    Do While r2 >= r1 And Application.CountA(ws.Rows(r2)) = 0: r2 = r2 - 1: Loop
    Do While c1 <= c2 And Application.CountA(ws.Columns(c1)) = 0: c1 = c1 + 1: Loop
    Do While c2 >= c1 And Application.CountA(ws.Columns(c2)) = 0: c2 = c2 - 1: Loop

    If r2 < r1 Or c2 < c1 Then
        Set TrimUsedRange = Nothing
    Else
        Set TrimUsedRange = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    End If
End Function

Private Sub FilterKeepOnlyBusinessUnits(ByRef lo As ListObject, ByVal keepArr As Variant)
    Dim ws As Worksheet
    Dim buCol As Long
    Dim rngVisible As Range
    Dim tmp As Worksheet
    Dim loIt As ListObject

    Set ws = lo.Parent
    buCol = GetColumnIndex(lo, "Business Unit")
    If buCol = 0 Then Err.Raise 1002, , "Column 'Business Unit' not found."

    On Error Resume Next
    lo.Range.AutoFilter Field:=buCol, Criteria1:=keepArr, Operator:=xlFilterValues
    On Error GoTo 0

    ' If nothing visible except header, SpecialCells throws 1004. Handle it.
    On Error Resume Next
    Set rngVisible = lo.Range.SpecialCells(xlCellTypeVisible)
    If Err.Number <> 0 Then
        Err.Clear
        ' Rebuild sheet with headers only (no matching rows)
        ws.Cells.Clear
        ws.Range("A1").Resize(1, lo.HeaderRowRange.Columns.Count).Value = lo.HeaderRowRange.Value
    Else
        ' Copy kept rows to a temporary sheet then replace original
        Set tmp = ws.Parent.Worksheets.Add(After:=ws)
        rngVisible.Copy tmp.Range("A1")

        ws.Cells.Clear
        tmp.UsedRange.Copy ws.Range("A1")

        SafeDeleteSheet tmp ' <<< SAFE DELETE
    End If
    On Error GoTo 0

    ' Remove any existing tables and recreate a clean one
    If ws.ListObjects.Count > 0 Then
        For Each loIt In ws.ListObjects
            loIt.Delete
        Next loIt
    End If

    Set lo = EnsureTable(ws, "ApprovedTbl")
End Sub

Private Function JoinColumnValues(ByVal lo As ListObject, ByVal headerName As String, ByVal delim As String) As String
    Dim idx As Long
    Dim arr As Variant
    Dim i As Long
    Dim s As String
    Dim valStr As String

    idx = GetColumnIndex(lo, headerName)
    If idx = 0 Then Err.Raise 1003, , "Column '" & headerName & "' not found."

    If lo.DataBodyRange Is Nothing Then
        JoinColumnValues = ""
        Exit Function
    End If

    arr = lo.DataBodyRange.Columns(idx).Value
    For i = 1 To UBound(arr, 1)
        valStr = Trim$(CStr(arr(i, 1)))
        If Len(valStr) > 0 Then
            If Len(s) > 0 Then s = s & delim
            s = s & valStr
        End If
    Next i
    JoinColumnValues = s
End Function

Private Sub CopyToClipboard(ByVal textVal As String)
    Dim prevScr As Boolean
    Dim wsTmp As Worksheet
    Dim o As Object

    On Error Resume Next
    Set o = CreateObject("MSForms.DataObject")
    If Err.Number = 0 Then
        o.SetText textVal
        o.PutInClipboard
        Exit Sub
    End If
    On Error GoTo 0

    ' Fallback using hidden sheet (safe add/delete)
    prevScr = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set wsTmp = ThisWorkbook.Worksheets.Add
    wsTmp.Visible = xlSheetVeryHidden
    wsTmp.Range("A1").Value = textVal
    wsTmp.Range("A1").Copy
    Application.CutCopyMode = False

    SafeDeleteSheet wsTmp ' <<< SAFE DELETE

    Application.ScreenUpdating = prevScr
End Sub

Private Function CreateDatedSheet(ByVal wb As Workbook) As Worksheet
    Dim baseName As String
    Dim nameCandidate As String
    Dim n As Long

    baseName = Format(Date, "yyyy-mm-dd")
    nameCandidate = baseName
    n = 1

    Do While SheetExists(wb, nameCandidate)
        n = n + 1
        nameCandidate = baseName & " (" & n & ")"
    Loop

    Set CreateDatedSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    CreateDatedSheet.Name = nameCandidate
End Function

Private Function EnsureSheet(ByVal wb As Workbook, ByVal name As String, Optional ByVal recreate As Boolean = False) As Worksheet
    If recreate And SheetExists(wb, name) Then
        SafeDeleteSheet wb.Worksheets(name) ' <<< SAFE DELETE
    End If

    If SheetExists(wb, name) Then
        Set EnsureSheet = wb.Worksheets(name)
    Else
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    On Error Resume Next
    SheetExists = Not wb.Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

Private Sub CopyColumnsByName(ByVal lo As ListObject, ByVal wsDest As Worksheet, ByVal fieldNames As Variant)
    Dim i As Long
    Dim idx As Long
    Dim nextCol As Long
    nextCol = 1

    wsDest.Cells.Clear

    ' headers
    For i = LBound(fieldNames) To UBound(fieldNames)
        idx = GetColumnIndex(lo, CStr(fieldNames(i)))
        If idx = 0 Then Err.Raise 1004, , "Column '" & CStr(fieldNames(i)) & "' not found in Credit Studio."
        wsDest.Cells(1, nextCol).Value = CStr(fieldNames(i))

        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.Columns(idx).Copy wsDest.Cells(2, nextCol)
        End If
        nextCol = nextCol + 1
    Next i
End Sub

Private Function GetColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim$(LCase$(CStr(lo.HeaderRowRange.Cells(1, i).Value))) = Trim$(LCase$(headerName)) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Private Function BuildCoperToCoRMap(ByVal lo As ListObject, ByVal coperCol As String, ByVal corCol As String) As Object
    Dim idxC As Long, idxR As Long
    Dim dict As Object
    Dim r As Long
    Dim vCoper As String, vCoR As String

    idxC = GetColumnIndex(lo, coperCol)
    idxR = GetColumnIndex(lo, corCol)
    If idxC = 0 Then Err.Raise 1005, , "Column '" & coperCol & "' not found in Approved."
    If idxR = 0 Then Err.Raise 1006, , "Column '" & corCol & "' not found in Approved."

    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            vCoper = Trim$(CStr(lo.DataBodyRange.Cells(r, idxC).Value))
            vCoR = Trim$(CStr(lo.DataBodyRange.Cells(r, idxR).Value))
            If Len(vCoper) > 0 Then
                dict(vCoper) = vCoR
            End If
        Next r
    End If
    Set BuildCoperToCoRMap = dict
End Function

Private Sub AppendApprovedCoR(ByVal ws As Worksheet, ByVal approvedMap As Object, _
                              ByVal coperColName As String, ByVal outColName As String)
    Dim lastRow As Long, lastCol As Long
    Dim coperCol As Long
    Dim r As Long
    Dim key As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    coperCol = FindHeader(ws, coperColName)
    If coperCol = 0 Then Err.Raise 1007, , "Column '" & coperColName & "' not found on " & ws.Name

    ws.Cells(1, lastCol + 1).Value = outColName

    For r = 2 To lastRow
        key = Trim$(CStr(ws.Cells(r, coperCol).Value))
        If Len(key) > 0 Then
            If approvedMap.Exists(key) Then
                ws.Cells(r, lastCol + 1).Value = approvedMap(key)
            Else
                ws.Cells(r, lastCol + 1).Value = ""
            End If
        End If
    Next r
End Sub

Private Function FindHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(LCase$(CStr(ws.Cells(1, c).Value))) = Trim$(LCase$(headerName)) Then
            FindHeader = c
            Exit Function
        End If
    Next c
    FindHeader = 0
End Function

Private Sub CreateMismatchSummary(ByVal wb As Workbook, ByVal loRecali As ListObject, _
                                  ByVal creditCoRColName As String, _
                                  ByVal approvedCoRColName As String, _
                                  ByVal coperColName As String, _
                                  ByVal summarySheetName As String)

    Dim idxCredit As Long, idxApproved As Long, idxCoper As Long
    Dim dict As Object
    Dim r As Long
    Dim valCredit As String, valApproved As String, coper As String
    Dim wsSum As Worksheet
    Dim lo As ListObject
    Dim k As Variant
    Dim rowOut As Long
    Dim joined As String

    idxCredit = GetColumnIndex(loRecali, creditCoRColName)
    idxApproved = GetColumnIndex(loRecali, approvedCoRColName)
    idxCoper = GetColumnIndex(loRecali, coperColName)

    If idxCredit * idxApproved * idxCoper = 0 Then
        Err.Raise 1008, , "One or more required columns missing in CoR Recali."
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    If Not loRecali.DataBodyRange Is Nothing Then
        For r = 1 To loRecali.DataBodyRange.Rows.Count
            valCredit = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxCredit).Value))
            valApproved = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxApproved).Value))
            coper = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxCoper).Value))

            If Len(valCredit) > 0 And Len(coper) > 0 Then
                If StrComp(valCredit, valApproved, vbTextCompare) <> 0 Then
                    If Not dict.Exists(valCredit) Then dict.Add valCredit, New Collection
                    dict(valCredit).Add coper
                End If
            End If
        Next r
    End If

    ' If no mismatches, remove any old summary and exit
    If dict.Count = 0 Then
        If SheetExists(wb, summarySheetName) Then
            SafeDeleteSheet wb.Worksheets(summarySheetName) ' <<< SAFE DELETE
        End If
        Exit Sub
    End If

    Set wsSum = EnsureSheet(wb, summarySheetName, True)
    wsSum.Range("A1").Value = "Country of Risk (Credit Studio)"
    wsSum.Range("B1").Value = "Coper IDs (mismatched; comma-joined)"

    rowOut = 2
    For Each k In dict.Keys
        joined = JoinUniqueFromCollection(dict(k), ",")
        wsSum.Cells(rowOut, 1).Value = CStr(k)
        wsSum.Cells(rowOut, 2).Value = joined
        rowOut = rowOut + 1
    Next k

    Set lo = wsSum.ListObjects.Add(xlSrcRange, wsSum.Range("A1").CurrentRegion, , xlYes)
    On Error Resume Next
    lo.Name = "CoRMismatchSummaryTbl"
    On Error GoTo 0

    wsSum.Columns.AutoFit
End Sub

Private Function JoinUniqueFromCollection(ByVal col As Collection, ByVal delim As String) As String
    Dim seen As Object
    Dim i As Long
    Dim valStr As String
    Dim s As String

    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To col.Count
        valStr = Trim$(CStr(col(i)))
        If Len(valStr) > 0 Then
            If Not seen.Exists(valStr) Then
                seen.Add valStr, True
                If Len(s) > 0 Then s = s & delim
                s = s & valStr
            End If
        End If
    Next i
    JoinUniqueFromCollection = s
End Function

' =======================
' Safe utilities (fix 1004 delete failures)
' =======================

Private Function IsStructureProtected(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    IsStructureProtected = wb.ProtectStructure
    On Error GoTo 0
End Function

Private Sub SafeDeleteSheet(ByVal ws As Worksheet)
    Dim wb As Workbook
    Set wb = ws.Parent

    ' If structure is protected or only one sheet remains, do NOT delete
    If IsStructureProtected(wb) Then Exit Sub
    If wb.Worksheets.Count <= 1 Then Exit Sub

    ' Make sure sheet is visible before delete (Excel can be fussy)
    On Error Resume Next
    ws.Visible = xlSheetVisible
    On Error GoTo 0

    ' Delete with alerts suppressed
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ws.Delete
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
End Sub