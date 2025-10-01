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
' Entry point
'========================
Public Sub Run_ApprovedFunds_CreditStudio_Workflow()
    On Error GoTo Fail

    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    If IsStructureProtected(wbMain) Then
        MsgBox "Workbook structure is protected. Unprotect (Review → Protect Workbook) and run again.", vbCritical
        Exit Sub
    End If

    Dim wsDate As Worksheet, wsRecali As Worksheet
    Dim approvedPath As String, creditPath As String
    Dim wbApproved As Workbook, wbCredit As Workbook
    Dim loApproved As ListObject, loCredit As ListObject
    Dim approvedMap As Object
    Dim loRecali As ListObject

    ' 1) Pick Approved Funds CSV
    approvedPath = PickFile("Select APPROVED FUNDS CSV", "CSV Files (*.csv)", "*.csv")
    If Len(approvedPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: Exit Sub

    ' 2) Open & delete first row
    Set wbApproved = Workbooks.Open(Filename:=approvedPath, Local:=True)
    With wbApproved.Worksheets(1)
        .Rows(1).Delete
    End With

    ' 3) Table-ize
    Set loApproved = EnsureTable(wbApproved.Worksheets(1), "ApprovedTbl")

    ' 4) Keep only target Business Units
    FilterKeepOnlyBusinessUnits loApproved, Array("FI-GMC-ASIA", "FI-US", "FI-EMEA")
    ' loApproved refreshed inside

    ' 5) Copy Fund CoPER in batches of 600 (user pastes each batch in Credit Studio)
    CopyCoperBatches loApproved, "Fund CoPER", 600

    ' 5b) Pick Credit Studio XLSX
    creditPath = PickFile("Select CREDIT STUDIO XLSX", "Excel Files (*.xlsx)", "*.xlsx")
    If Len(creditPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: GoTo Cleanup
    Set wbCredit = Workbooks.Open(Filename:=creditPath, ReadOnly:=True)

    ' 6) Create dated sheet in MAIN
    Set wsDate = CreateDatedSheet(wbMain)

    ' 7) Credit Studio → table; copy "Coper ID" & "Country of Risk" to "CoR Recali"
    Set loCredit = EnsureTable(wbCredit.Worksheets(1), "CreditTbl")
    Set wsRecali = EnsureSheet(wbMain, "CoR Recali", True)
    CopyColumnsByName loCredit, wsRecali, Array("Coper ID", "Country of Risk")

    ' 8) Build Approved map and append Approved CoR
    Set approvedMap = BuildCoperToCoRMap(loApproved, "Fund CoPER", "Country of Risk")
    AppendApprovedCoR wsRecali, approvedMap, "Coper ID", "Approved CoR"

    ' 9) Table-ize & mismatch summary
    Set loRecali = EnsureTable(wsRecali, "CoRRecaliTbl")
    CreateMismatchSummary wbMain, loRecali, _
        creditCoRColName:="Country of Risk", _
        approvedCoRColName:="Approved CoR", _
        coperColName:="Coper ID", _
        summarySheetName:="CoR Mismatch Summary"

    MsgBox "Done. Sheets created: '" & wsDate.Name & "', 'CoR Recali', and (if needed) 'CoR Mismatch Summary'.", vbInformation
    GoTo Cleanup

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical

Cleanup:
    On Error Resume Next
    If Not wbCredit Is Nothing Then wbCredit.Close SaveChanges:=False
    If Not wbApproved Is Nothing Then wbApproved.Close SaveChanges:=True
End Sub

'========================
' Batching: copy Fund CoPER in chunks
'========================
Private Sub CopyCoperBatches(ByVal lo As ListObject, ByVal headerName As String, ByVal batchSize As Long)
    Dim vals() As String
    vals = GetColumnValues(lo, headerName)
    If Not IsArrayAllocated(vals) Then
        MsgBox "No '" & headerName & "' values found.", vbCritical
        Exit Sub
    End If

    Dim total As Long: total = UBound(vals) - LBound(vals) + 1
    If total = 0 Then
        MsgBox "No '" & headerName & "' values found.", vbCritical
        Exit Sub
    End If

    Dim totalBatches As Long
    totalBatches = (total \ batchSize) + IIf(total Mod batchSize = 0, 0, 1)

    Dim b As Long, startIdx As Long, endIdx As Long, sizeThis As Long
    For b = 1 To totalBatches
        startIdx = (b - 1) * batchSize + 1
        endIdx = WorksheetFunction.Min(b * batchSize, total)
        sizeThis = endIdx - startIdx + 1

        Dim payload As String
        payload = JoinRange(vals, startIdx, endIdx, ",")

        CopyToClipboard payload

        If b < totalBatches Then
            If MsgBox( _
                "Batch " & b & " of " & totalBatches & " copied (" & sizeThis & " IDs)." & vbCrLf & _
                "Paste into Credit Studio (or wherever needed), then click OK for the next batch." & vbCrLf & _
                "Click Cancel to abort.", _
                vbOKCancel + vbInformation, "Copy CoPER Batches") = vbCancel Then
                Err.Clear
                Exit Sub
            End If
        Else
            MsgBox "Final batch " & b & " of " & totalBatches & " copied (" & sizeThis & " IDs). Paste it now. Moving on next.", vbInformation, "Copy CoPER Batches"
        End If
    Next b
End Sub

Private Function GetColumnValues(ByVal lo As ListObject, ByVal headerName As String) As String()
    Dim idx As Long: idx = GetColumnIndex(lo, headerName)
    Dim out() As String
    Dim r As Long, n As Long

    If idx = 0 Or lo.DataBodyRange Is Nothing Then
        ' return unallocated
        Exit Function
    End If

    ReDim out(1 To lo.DataBodyRange.Rows.Count)
    n = 0
    For r = 1 To lo.DataBodyRange.Rows.Count
        Dim v As String
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
    ' Works for uninitialized dynamic arrays without relying on Is Nothing
    On Error GoTo EH
    Dim lb As Long, ub As Long
    lb = LBound(arr)
    ub = UBound(arr)
    IsArrayAllocated = (ub >= lb)
    Exit Function
EH:
    IsArrayAllocated = False
End Function

'========================
' File pickers / sheets
'========================
Private Function PickFile(ByVal promptTitle As String, ByVal filterDesc As String, ByVal filterPattern As String) As String
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

'========================
' Table utilities
'========================
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

    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
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

    ' If only headers visible, SpecialCells throws.
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

    ' Drop any left-over tables, recreate clean one
    If ws.ListObjects.Count > 0 Then
        For Each loIt In ws.ListObjects
            loIt.Delete
        Next loIt
    End If
    Set lo = EnsureTable(ws, "ApprovedTbl")
End Sub

'========================
' Copy columns + matching
'========================
Private Sub CopyColumnsByName(ByVal lo As ListObject, ByVal wsDest As Worksheet, ByVal fieldNames As Variant)
    Dim i As Long, idx As Long, nextCol As Long
    wsDest.Cells.Clear
    nextCol = 1
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

Private Function BuildCoperToCoRMap(ByVal lo As ListObject, ByVal coperCol As String, ByVal corCol As String) As Object
    Dim idxC As Long, idxR As Long
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    Dim vCoper As String, vCoR As String

    idxC = GetColumnIndex(lo, coperCol)
    idxR = GetColumnIndex(lo, corCol)
    If idxC = 0 Then Err.Raise 1005, , "Column '" & coperCol & "' not found in Approved."
    If idxR = 0 Then Err.Raise 1006, , "Column '" & corCol & "' not found in Approved."

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
    Dim lastRow As Long, lastCol As Long, coperCol As Long
    Dim r As Long, key As String

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
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
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
    On Error Resume Next
    lo.Name = "CoRMismatchSummaryTbl"
    On Error GoTo 0
    wsSum.Columns.AutoFit
End Sub

Private Function JoinUniqueFromCollection(ByVal col As Collection, ByVal delim As String) As String
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim valStr As String
    Dim s As String
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
Private Sub CopyToClipboard(ByVal textVal As String)
    ' Try MSForms first (if available), then API fallback.
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
        Dim bytesNeeded As LongPtr
        Dim hGlobal As LongPtr
        Dim pGlobal As LongPtr
        Dim copyRes As LongPtr
        Dim ok As Long
    #Else
        Dim bytesNeeded As Long
        Dim hGlobal As Long
        Dim pGlobal As Long
        Dim copyRes As Long
        Dim ok As Long
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
    ' Do not free hGlobal after SetClipboardData; ownership transferred to system.
End Sub

'========================
' Safe sheet ops
'========================
Private Function IsStructureProtected(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    IsStructureProtected = wb.ProtectStructure
    On Error GoTo 0
End Function

Private Sub SafeDeleteSheet(ByVal ws As Worksheet)
    Dim wb As Workbook: Set wb = ws.Parent
    If IsStructureProtected(wb) Then Exit Sub
    If wb.Worksheets.Count <= 1 Then Exit Sub

    On Error Resume Next
    ws.Visible = xlSheetVisible
    On Error GoTo 0

    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    ws.Delete
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
End Sub