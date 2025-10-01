Option Explicit

'========================
' Clipboard API (bitness-safe, fallback)
'========================
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
    Private Const GMEM_MOVEABLE As Long = &H2
    Private Const CF_UNICODETEXT As Long = 13&
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
    Private Const GMEM_MOVEABLE As Long = &H2
    Private Const CF_UNICODETEXT As Long = 13&
#End If

'========================
' Globals for the modeless controller
'========================
Public G_Batches() As String          ' each element is a comma-joined payload
Public G_BatchIndex As Long           ' 1-based
Public G_BatchCount As Long
Public G_WorkflowReadyToContinue As Boolean

'========================
' Entry point
'========================
Public Sub Run_ApprovedFunds_CreditStudio_Workflow()
    On Error GoTo Fail

    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    If IsStructureProtected(wbMain) Then
        MsgBox "Workbook structure is protected. Unprotect (Review → Protect Workbook) and run again.", vbCritical
        Exit Sub
    End If

    Dim approvedPath As String
    Dim wbApproved As Workbook
    Dim loApproved As ListObject

    ' 1) Pick Approved Funds CSV
    approvedPath = PickFile("Select APPROVED FUNDS CSV", "CSV Files (*.csv)", "*.csv")
    If Len(approvedPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: Exit Sub

    ' 2) Open & delete first row (second row has headers)
    Set wbApproved = Workbooks.Open(Filename:=approvedPath, Local:=True)
    wbApproved.Worksheets(1).Rows(1).Delete

    ' 3) Table-ize
    Set loApproved = EnsureTable(wbApproved.Worksheets(1), "ApprovedTbl")

    ' 4) Keep only target Business Units
    FilterKeepOnlyBusinessUnits loApproved, Array("FI-GMC-ASIA", "FI-US", "FI-EMEA")
    ' loApproved is refreshed inside the helper

    ' 5) Build Fund CoPER batches of 600 and show modeless controller
    BuildCoperBatches loApproved, "Fund CoPER", 600
    G_WorkflowReadyToContinue = False
    frmCoperBatches.Show vbModeless
    Exit Sub

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'========================
' Phase 2 (called by the form after last batch)
'========================
Public Sub Continue_After_Batches()
    On Error GoTo Fail

    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    Dim creditFiles As Collection, p As Variant
    Dim wbCredit As Workbook
    Dim wsDate As Worksheet, wsRecali As Worksheet
    Dim loCredit As ListObject, loRecali As ListObject
    Dim approvedMap As Object
    Dim wbApproved As Workbook, loApproved As ListObject

    ' Reuse the open Approved workbook/table
    Set wbApproved = GetApprovedWorkbook()
    If wbApproved Is Nothing Then
        MsgBox "Approved Funds workbook not found. Re-run the workflow.", vbCritical
        Exit Sub
    End If
    Set loApproved = EnsureTable(wbApproved.Worksheets(1), "ApprovedTbl")

    ' Pick MULTIPLE Credit Studio files
    Set creditFiles = PickFilesMulti("Select one or more CREDIT STUDIO XLSX files", _
                                     "Excel Files (*.xlsx)", "*.xlsx")
    If creditFiles Is Nothing Or creditFiles.Count = 0 Then
        MsgBox "No Credit Studio files selected. Stopping.", vbInformation
        Exit Sub
    End If

    ' 6) Dated sheet
    Set wsDate = CreateDatedSheet(wbMain)

    ' 7) Prepare CoR Recali (recreate fresh; we will append)
    Set wsRecali = EnsureSheet(wbMain, "CoR Recali", True)
    EnsureRecaliHeaders wsRecali

    ' Append all selected Credit Studio files
    For Each p In creditFiles
        Set wbCredit = Workbooks.Open(Filename:=CStr(p), ReadOnly:=True)
        Set loCredit = EnsureTable(wbCredit.Worksheets(1), "CreditTbl") ' adjust if needed
        AppendColumnsByName loCredit, wsRecali, "Coper ID", "Country of Risk", VBA.Dir(CStr(p))
        wbCredit.Close SaveChanges:=False
    Next p

    ' (Optional) De-duplicate by Coper ID (keep last occurrence). Uncomment to use.
    'DedupRecaliByCoperID wsRecali

    ' 8) Build Approved map and append Approved CoR
    Set approvedMap = BuildCoperToCoRMap(loApproved, "Fund CoPER", "Country of Risk")
    AppendApprovedCoR wsRecali, approvedMap, "Coper ID", "Approved CoR"

    ' 9) Convert to table and build mismatch summary
    Set loRecali = EnsureTable(wsRecali, "CoRRecaliTbl")
    CreateMismatchSummary wbMain, loRecali, _
        creditCoRColName:="Country of Risk", _
        approvedCoRColName:="Approved CoR", _
        coperColName:="Coper ID", _
        summarySheetName:="CoR Mismatch Summary"

    MsgBox "Done. Consolidated " & creditFiles.Count & " Credit Studio file(s) into 'CoR Recali' and built the summary.", vbInformation
    Exit Sub

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function GetApprovedWorkbook() As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If SheetHasTableNamed(wb, "ApprovedTbl") Then
            Set GetApprovedWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetApprovedWorkbook = Nothing
End Function

Private Function SheetHasTableNamed(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If LCase(lo.Name) = LCase(tableName) Then
                SheetHasTableNamed = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

'========================
' Batching utilities
'========================
Private Sub BuildCoperBatches(ByVal lo As ListObject, ByVal headerName As String, ByVal batchSize As Long)
    Dim vals() As String
    vals = GetColumnValues(lo, headerName)
    If Not IsArrayAllocated(vals) Then
        Err.Raise vbObjectError + 700, , "No '" & headerName & "' values to batch."
    End If

    Dim total As Long: total = UBound(vals) - LBound(vals) + 1
    Dim totalBatches As Long
    totalBatches = (total \ batchSize) + IIf(total Mod batchSize = 0, 0, 1)

    ReDim G_Batches(1 To totalBatches)
    Dim b As Long, sIdx As Long, eIdx As Long
    For b = 1 To totalBatches
        sIdx = (b - 1) * batchSize + 1
        eIdx = WorksheetFunction.Min(b * batchSize, total)
        G_Batches(b) = JoinRange(vals, sIdx, eIdx, ",")
    Next b

    G_BatchIndex = 1
    G_BatchCount = totalBatches
End Sub

Private Function GetColumnValues(ByVal lo As ListObject, ByVal headerName As String) As String()
    Dim idx As Long: idx = GetColumnIndex(lo, headerName)
    Dim out() As String
    Dim r As Long, n As Long, v As String

    If idx = 0 Or lo.DataBodyRange Is Nothing Then Exit Function

    ReDim out(1 To lo.DataBodyRange.Rows.Count)
    For r = 1 To lo.DataBodyRange.Rows.Count
        v = Trim$(CStr(lo.DataBodyRange.Cells(r, idx).Value))
        If Len(v) > 0 Then
            n = n + 1
            out(n) = v
        End If
    Next r

    If n = 0 Then
        Erase out
    ElseIf n < UBound(out) Then
        ReDim Preserve out(1 To n)
    End If
    GetColumnValues = out
End Function

Private Function JoinRange(ByRef arr() As String, ByVal startIdx As Long, ByVal endIdx As Long, ByVal delim As String) As String
    Dim i As Long, s As String
    For i = startIdx To endIdx
        If Len(s) > 0 Then s = s & delim
        s = s & arr(i)
    Next i
    JoinRange = s
End Function

Private Function IsArrayAllocated(ByRef arr() As String) As Boolean
    On Error GoTo EH
    Dim lb As Long, ub As Long
    lb = LBound(arr): ub = UBound(arr)
    IsArrayAllocated = (ub >= lb)
    Exit Function
EH:
    IsArrayAllocated = False
End Function

'========================
' File pickers (single / multi)
'========================
Private Function PickFile(ByVal promptTitle As String, ByVal filterDesc As String, ByVal filterPattern As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = promptTitle
        .Filters.Clear
        .Filters.Add filterDesc, filterPattern
        .AllowMultiSelect = False
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Function PickFilesMulti(ByVal promptTitle As String, _
                                ByVal filterDesc As String, _
                                ByVal filterPattern As String) As Collection
    Dim fd As FileDialog, i As Long
    Set PickFilesMulti = New Collection
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = promptTitle
        .Filters.Clear
        .Filters.Add filterDesc, filterPattern
        .AllowMultiSelect = True
        If .Show = -1 Then
            For i = 1 To .SelectedItems.Count
                PickFilesMulti.Add .SelectedItems(i)
            Next i
        End If
    End With
End Function

'========================
' Sheet utilities
'========================
Private Function CreateDatedSheet(ByVal wb As Workbook) As Worksheet
    Dim baseName As String, nameCandidate As String, n As Long
    baseName = Format(Date, "yyyy-mm-dd")
    nameCandidate = baseName: n = 1
    Do While SheetExists(wb, nameCandidate)
        n = n + 1
        nameCandidate = baseName & " (" & n & ")"
    Loop
    Set CreateDatedSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    CreateDatedSheet.Name = nameCandidate
End Function

Private Function EnsureSheet(ByVal wb As Workbook, ByVal name As String, Optional ByVal recreate As Boolean = False) As Worksheet
    If recreate And SheetExists(wb, name) Then SafeDeleteSheet wb.Worksheets(name)
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

Private Function IsStructureProtected(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    IsStructureProtected = wb.ProtectStructure
    On Error GoTo 0
End Function

Private Sub SafeDeleteSheet(ByVal ws As Worksheet)
    Dim wb As Workbook: Set wb = ws.Parent
    If IsStructureProtected(wb) Then Exit Sub
    If wb.Worksheets.Count <= 1 Then Exit Sub
    On Error Resume Next: ws.Visible = xlSheetVisible: On Error GoTo 0
    Dim prevAlerts As Boolean: prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next: ws.Delete: On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
End Sub

'========================
' Table utilities
'========================
Private Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject, rng As Range
    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        On Error Resume Next: lo.Name = tableName: On Error GoTo 0
        Set EnsureTable = lo: Exit Function
    End If
    Set rng = TrimUsedRange(ws)
    If rng Is Nothing Then Err.Raise 1001, , "No data found on sheet '" & ws.Name & "'."
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    On Error Resume Next: lo.Name = tableName: On Error GoTo 0
    Set EnsureTable = lo
End Function

Private Function TrimUsedRange(ByVal ws As Worksheet) As Range
    Dim ur As Range, r1 As Long, r2 As Long, c1 As Long, c2 As Long
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

Private Function GetColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim$(LCase$(CStr(lo.HeaderRowRange.Cells(1, i).Value))) = Trim$(LCase$(headerName)) Then
            GetColumnIndex = i: Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

'========================
' Filtering + transforms
'========================
Private Sub FilterKeepOnlyBusinessUnits(ByRef lo As ListObject, ByVal keepArr As Variant)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim buCol As Long: buCol = GetColumnIndex(lo, "Business Unit")
    Dim rngVisible As Range
    Dim tmp As Worksheet
    Dim loIt As ListObject

    If buCol = 0 Then Err.Raise 1002, , "Column 'Business Unit' not found."

    On Error Resume Next
    lo.Range.AutoFilter Field:=buCol, Criteria1:=keepArr, Operator:=xlFilterValues
    On Error GoTo 0

    ' Copy visible rows; if only headers visible, rebuild headers only
    On Error Resume Next
    Set rngVisible = lo.Range.SpecialCells(xlCellTypeVisible)
    If Err.Number <> 0 Then
        Err.Clear
        ws.Cells.Clear
        ws.Range("A1").Resize(1, lo.HeaderRowRange.Columns.Count).Value = lo.HeaderRowRange.Value
    Else
        Set tmp = ws.Parent.Worksheets.Add(After:=ws)
        rngVisible.Copy tmp.Range("A1")
        ws.Cells.Clear
        tmp.UsedRange.Copy ws.Range("A1")
        SafeDeleteSheet tmp
    End If
    On Error GoTo 0

    ' Rebuild table cleanly
    If ws.ListObjects.Count > 0 Then
        For Each loIt In ws.ListObjects
            loIt.Delete
        Next loIt
    End If
    Set lo = EnsureTable(ws, "ApprovedTbl")
End Sub

'========================
' Append & dedupe (for multi Credit Studio files)
'========================
' Ensure CoR Recali headers (with Source File)
Private Sub EnsureRecaliHeaders(ByVal ws As Worksheet)
    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Cells(1, 1).Value = "Coper ID"
        ws.Cells(1, 2).Value = "Country of Risk"
        ws.Cells(1, 3).Value = "Source File"
    End If
End Sub

' Next empty row in column A
Private Function NextWriteRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < 1 Then NextWriteRow = 1 Else NextWriteRow = r + 1
End Function

' Append two named columns from a table to wsDest, plus Source File tag
Private Sub AppendColumnsByName(ByVal lo As ListObject, _
                                ByVal wsDest As Worksheet, _
                                ByVal field1 As String, _
                                ByVal field2 As String, _
                                ByVal sourceTag As String)
    Dim idx1 As Long, idx2 As Long, r As Long, outRow As Long
    idx1 = GetColumnIndex(lo, field1)
    idx2 = GetColumnIndex(lo, field2)
    If idx1 = 0 Then Err.Raise 2001, , "Column '" & field1 & "' not found."
    If idx2 = 0 Then Err.Raise 2002, , "Column '" & field2' not found."
    If lo.DataBodyRange Is Nothing Then Exit Sub

    EnsureRecaliHeaders wsDest
    outRow = NextWriteRow(wsDest)

    For r = 1 To lo.DataBodyRange.Rows.Count
        wsDest.Cells(outRow, 1).Value = lo.DataBodyRange.Cells(r, idx1).Value
        wsDest.Cells(outRow, 2).Value = lo.DataBodyRange.Cells(r, idx2).Value
        wsDest.Cells(outRow, 3).Value = sourceTag
        outRow = outRow + 1
    Next r
End Sub

' Optional: de-duplicate by Coper ID (keep last occurrence)
Private Sub DedupRecaliByCoperID(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), Order:=xlDescending
        .SetRange ws.Range("A1:C" & lastRow)
        .Header = xlYes
        .Apply
    End With
    ws.Range("A1:C" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
End Sub

'========================
' Copy columns + matching
'========================
Private Function BuildCoperToCoRMap(ByVal lo As ListObject, ByVal coperCol As String, ByVal corCol As String) As Object
    Dim idxC As Long, idxR As Long, dict As Object, r As Long
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
            If Len(vCoper) > 0 Then dict(vCoper) = vCoR
        Next r
    End If
    Set BuildCoperToCoRMap = dict
End Function

Private Sub AppendApprovedCoR(ByVal ws As Worksheet, ByVal approvedMap As Object, _
                              ByVal coperColName As String, ByVal outColName As String)
    Dim lastRow As Long, lastCol As Long, coperCol As Long, r As Long, key As String
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
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(LCase$(CStr(ws.Cells(1, c).Value))) = Trim$(LCase$(headerName)) Then
            FindHeader = c: Exit Function
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
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, valCredit As String, valApproved As String, coper As String
    Dim wsSum As Worksheet, lo As ListObject, k As Variant, rowOut As Long, joined As String

    idxCredit = GetColumnIndex(loRecali, creditCoRColName)
    idxApproved = GetColumnIndex(loRecali, approvedCoRColName)
    idxCoper = GetColumnIndex(loRecali, coperColName)
    If idxCredit * idxApproved * idxCoper = 0 Then Err.Raise 1008, , "Required columns missing in CoR Recali."

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

    If dict.Count = 0 Then
        If SheetExists(wb, summarySheetName) Then SafeDeleteSheet wb.Worksheets(summarySheetName)
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
    On Error Resume Next: lo.Name = "CoRMismatchSummaryTbl": On Error GoTo 0
    wsSum.Columns.AutoFit
End Sub

Private Function JoinUniqueFromCollection(ByVal col As Collection, ByVal delim As String) As String
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, valStr As String, s As String
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

'========================
' Clipboard (robust)
'========================
Public Sub CopyTextToClipboard(ByVal textVal As String)
    ' Try MSForms first; fallback to API.
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("MSForms.DataObject")
    If Err.Number = 0 Then
        o.SetText textVal
        o.PutInClipboard
        Exit Sub
    End If
    On Error GoTo 0
    ClipboardSetTextAPI textVal
End Sub

Private Sub ClipboardSetTextAPI(ByVal textVal As String)
    #If VBA7 Then
        Dim bytesNeeded As LongPtr, hGlobal As LongPtr, pGlobal As LongPtr, copyRes As LongPtr
        Dim ok As Long
    #Else
        Dim bytesNeeded As Long, hGlobal As Long, pGlobal As Long, copyRes As Long, ok As Long
    #End If

    bytesNeeded = (Len(textVal) * 2) + 2 ' UTF-16 + null
    hGlobal = GlobalAlloc(GMEM_MOVEABLE, bytesNeeded)
    If hGlobal = 0 Then Err.Raise vbObjectError + 600, , "Clipboard alloc failed."

    pGlobal = GlobalLock(hGlobal)
    If pGlobal = 0 Then Err.Raise vbObjectError + 601, , "Clipboard lock failed."

    copyRes = lstrcpyW(pGlobal, StrPtr(textVal))
    GlobalUnlock hGlobal

    ok = OpenClipboard(0)
    If ok = 0 Then Err.Raise vbObjectError + 602, , "OpenClipboard failed."
    EmptyClipboard
    SetClipboardData CF_UNICODETEXT, hGlobal
    CloseClipboard
End Sub