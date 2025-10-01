Option Explicit

' ====== USER CONFIG ======
Private Const ITERATION_DIR As String = "C:\Your\Folder\Here\"   ' <-- SET THIS

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
' Entry point (renamed for AllFund)
'========================
Public Sub Run_AllFund_CreditStudio_Workflow()
    On Error GoTo Fail

    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    If IsStructureProtected(wbMain) Then
        MsgBox "Workbook structure is protected. Unprotect (Review → Protect Workbook) and run again.", vbCritical
        Exit Sub
    End If

    Dim allFundPath As String
    Dim wbAllFund As Workbook
    Dim loAllFund As ListObject

    ' 1) Pick AllFund CSV
    allFundPath = PickFile("Select ALLFUND CSV", "CSV Files (*.csv)", "*.csv")
    If Len(allFundPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: Exit Sub

    ' 2) Open & delete first row (second row has headers)
    Set wbAllFund = Workbooks.Open(Filename:=allFundPath, Local:=True)
    wbAllFund.Worksheets(1).Rows(1).Delete

    ' 3) Table-ize
    Set loAllFund = EnsureTable(wbAllFund.Worksheets(1), "AllFundTbl")

    ' 4) Filter: BU ∈ {FI-US, FI-EMEA, FI-GMC- Asia / FI-GMC-ASIA} AND Review Status ∈ {Approved, Submitted}
    FilterAllFundCriteria loAllFund, _
        keepBUs:=Array("FI-US", "FI-EMEA", "FI-GMC- Asia", "FI-GMC-ASIA"), _
        keepReviewStatus:=Array("Approved", "Submitted")
    ' loAllFund is refreshed inside

    ' 5) Build Fund CoPER batches of 600 and show modeless controller
    BuildCoperBatches loAllFund, "Fund CoPER", 600
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
    Dim allFundCoRMap As Object, allFundRegionMap As Object
    Dim wbAllFund As Workbook, loAllFund As ListObject

    ' Reuse the open AllFund workbook/table
    Set wbAllFund = GetAllFundWorkbook()
    If wbAllFund Is Nothing Then
        MsgBox "AllFund workbook not found. Re-run the workflow.", vbCritical
        Exit Sub
    End If
    Set loAllFund = EnsureTable(wbAllFund.Worksheets(1), "AllFundTbl")

    ' Pick MULTIPLE Credit Studio files
    Set creditFiles = PickFilesMulti("Select one or more CREDIT STUDIO XLSX files", _
                                     "Excel Files (*.xlsx)", "*.xlsx")
    If creditFiles Is Nothing Or creditFiles.Count = 0 Then
        MsgBox "No Credit Studio files selected. Stopping.", vbInformation
        Exit Sub
    End If

    ' 6) Dated sheet (in Main, but we will move to Iteration file)
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

    ' 8) Build AllFund maps and append Approved CoR (truth)
    Set allFundCoRMap = BuildCoperToCoRMap(loAllFund, "Fund CoPER", "Country of Risk")
    Set allFundRegionMap = BuildCoperToRegionMap(loAllFund, "Fund CoPER", "Business Unit")
    AppendApprovedCoR wsRecali, allFundCoRMap, "Coper ID", "Approved CoR"

    ' 9) Convert to table and build mismatch summary (group by Approved CoR, IDs sanitized)
    Set loRecali = EnsureTable(wsRecali, "CoRRecaliTbl")
    CreateMismatchSummary ThisWorkbook, loRecali, _
        creditCoRColName:="Country of Risk", _
        approvedCoRColName:="Approved CoR", _
        coperColName:="Coper ID", _
        summarySheetName:="CoR Mismatch Summary"

    ' 10) Create Iteration file (previous month name), move sheets + build Stats
    Dim iterWb As Workbook
    Set iterWb = CreateIterationWorkbook()
    CopySheetValues wbMain.Worksheets(wsDate.Name), iterWb, wsDate.Name
    CopySheetValues wbMain.Worksheets("CoR Recali"), iterWb, "CoR Recali"
    If SheetExists(wbMain, "CoR Mismatch Summary") Then
        CopySheetValues wbMain.Worksheets("CoR Mismatch Summary"), iterWb, "CoR Mismatch Summary"
    End If

    ' Build Stats in iteration file (wording adapted to AllFund)
    BuildStatsSheet iterWb, iterWb.Worksheets("CoR Recali"), allFundCoRMap, allFundRegionMap

    ' Save iteration file
    SaveIterationWorkbook iterWb

    MsgBox "Done. Iteration file created with CoR Recali, Mismatch Summary (if any) and Stats.", vbInformation
    Exit Sub

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'========================
' Locate open AllFund workbook
'========================
Private Function GetAllFundWorkbook() As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If SheetHasTableNamed(wb, "AllFundTbl") Then
            Set GetAllFundWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetAllFundWorkbook = Nothing
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
' ALLFUND dual-filter (BU + Review Status)
'========================
Private Sub FilterAllFundCriteria(ByRef lo As ListObject, ByVal keepBUs As Variant, ByVal keepReviewStatus As Variant)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim buCol As Long: buCol = GetColumnIndex(lo, "Business Unit")
    Dim revCol As Long: revCol = GetColumnIndex(lo, "Review Status")
    Dim rngVisible As Range
    Dim tmp As Worksheet
    Dim loIt As ListObject

    If buCol = 0 Then Err.Raise vbObjectError + 1101, , "Column 'Business Unit' not found."
    If revCol = 0 Then Err.Raise vbObjectError + 1102, , "Column 'Review Status' not found."

    ' Apply both filters
    On Error Resume Next
    lo.Range.AutoFilter Field:=buCol, Criteria1:=keepBUs, Operator:=xlFilterValues
    lo.Range.AutoFilter Field:=revCol, Criteria1:=keepReviewStatus, Operator:=xlFilterValues
    On Error GoTo 0

    ' Copy visible rows; handle header-only case
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
    Set lo = EnsureTable(ws, "AllFundTbl")
End Sub

'========================
' Iteration file helpers
'========================
Private Function CreateIterationWorkbook() As Workbook
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add(xlWBATWorksheet) ' single sheet
    wb.Worksheets(1).Name = "Stats"
    Set CreateIterationWorkbook = wb
End Function

Private Sub CopySheetValues(ByVal srcWs As Worksheet, ByVal destWb As Workbook, ByVal newName As String)
    Dim ws As Worksheet
    Set ws = destWb.Worksheets.Add(After:=destWb.Worksheets(destWb.Worksheets.Count))
    ws.Name = newName
    srcWs.UsedRange.Copy
    ws.Range("A1").PasteSpecial xlPasteValues
    ws.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Function PrevMonthFileName() As String
    Dim d As Date, s As String
    d = DateSerial(Year(Date), Month(Date), 1) - 1   ' last day of prev month
    s = Format(d, "mmmm-yyyy") & ".xlsx"             ' e.g., September-2025.xlsx
    PrevMonthFileName = s
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Sub SaveIterationWorkbook(ByVal wb As Workbook)
    Dim path As String
    EnsureFolder ITERATION_DIR
    path = ITERATION_DIR
    If Right$(path, 1) <> "\" And Right$(path, 1) <> "/" Then path = path & "\"
    path = path & PrevMonthFileName()
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook ' .xlsx
    Application.DisplayAlerts = True
End Sub

'========================
' Stats sheet (global + regional)
'========================
Private Sub BuildStatsSheet(ByVal iterWb As Workbook, _
                            ByVal recaliWs As Worksheet, _
                            ByVal allFundCoRMap As Object, _
                            ByVal allFundRegionMap As Object)

    Dim ws As Worksheet: Set ws = iterWb.Worksheets("Stats")
    ws.Cells.Clear

    ' Build credit dict (unique Copers -> Credit CoR)
    Dim creditDict As Object: Set creditDict = CreateObject("Scripting.Dictionary")
    Dim cLastRow As Long, r As Long
    Dim colCoper As Long, colCreditCoR As Long

    colCoper = FindHeader(recaliWs, "Coper ID")
    colCreditCoR = FindHeader(recaliWs, "Country of Risk")
    If colCoper = 0 Or colCreditCoR = 0 Then Err.Raise vbObjectError + 800, , "CoR Recali missing columns."

    cLastRow = recaliWs.Cells(recaliWs.Rows.Count, 1).End(xlUp).Row
    For r = 2 To cLastRow
        Dim cid As String, cCoR As String
        cid = SanitizeCoperID(recaliWs.Cells(r, colCoper).Value)
        If Len(cid) > 0 Then
            cCoR = Trim$(CStr(recaliWs.Cells(r, colCreditCoR).Value))
            creditDict(cid) = cCoR   ' last occurrence wins
        End If
    Next r

    ' Regions based on AllFund "Business Unit"
    Dim REGIONS As Variant
    REGIONS = Array("AMRS", "EMEA", "APAC")

    ' Counters (global)
    Dim gAllFundTotal As Long, gCreditTotal As Long
    Dim gAllFundInCredit As Long, gAllFundNotInCredit As Long
    Dim gCorrect As Long, gMismatch As Long

    ' Per-region dictionaries of counters
    Dim regAllFundTotal As Object, regAllFundInCredit As Object
    Dim regAllFundNotInCredit As Object, regCorrect As Object, regMismatch As Object
    Set regAllFundTotal = CreateObject("Scripting.Dictionary")
    Set regAllFundInCredit = CreateObject("Scripting.Dictionary")
    Set regAllFundNotInCredit = CreateObject("Scripting.Dictionary")
    Set regCorrect = CreateObject("Scripting.Dictionary")
    Set regMismatch = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(REGIONS) To UBound(REGIONS)
        regAllFundTotal(REGIONS(i)) = 0
        regAllFundInCredit(REGIONS(i)) = 0
        regAllFundNotInCredit(REGIONS(i)) = 0
        regCorrect(REGIONS(i)) = 0
        regMismatch(REGIONS(i)) = 0
    Next i

    ' Iterate AllFund set (keys of map)
    Dim k As Variant
    For Each k In allFundCoRMap.Keys
        gAllFundTotal = gAllFundTotal + 1

        Dim reg As String
        reg = RegionFromBU(allFundRegionMap(k))
        If Len(reg) = 0 Then reg = "Unmapped"
        If Not regAllFundTotal.Exists(reg) Then
            regAllFundTotal(reg) = 0
            regAllFundInCredit(reg) = 0
            regAllFundNotInCredit(reg) = 0
            regCorrect(reg) = 0
            regMismatch(reg) = 0
        End If
        regAllFundTotal(reg) = CLng(regAllFundTotal(reg)) + 1

        If creditDict.Exists(k) Then
            gAllFundInCredit = gAllFundInCredit + 1
            regAllFundInCredit(reg) = CLng(regAllFundInCredit(reg)) + 1

            If StrComp(creditDict(k), allFundCoRMap(k), vbTextCompare) = 0 Then
                gCorrect = gCorrect + 1
                regCorrect(reg) = CLng(regCorrect(reg)) + 1
            Else
                gMismatch = gMismatch + 1
                regMismatch(reg) = CLng(regMismatch(reg)) + 1
            End If
        Else
            gAllFundNotInCredit = gAllFundNotInCredit + 1
            regAllFundNotInCredit(reg) = CLng(regAllFundNotInCredit(reg)) + 1
        End If
    Next k

    ' Global credit total (all unique Copers found in Credit Studio files)
    gCreditTotal = creditDict.Count

    ' --- Write Global table (AllFund wording) ---
    Dim row As Long: row = 1
    ws.Cells(row, 1).Value = "Global Stats (AllFund)":
    row = row + 1
    ws.Cells(row, 1).Value = "Metric": ws.Cells(row, 2).Value = "Count"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 2)).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "AllFund Total": ws.Cells(row, 2).Value = gAllFundTotal: row = row + 1
    ws.Cells(row, 1).Value = "Credit Total (unique Copers)": ws.Cells(row, 2).Value = gCreditTotal: row = row + 1
    ws.Cells(row, 1).Value = "AllFund Present in Credit": ws.Cells(row, 2).Value = gAllFundInCredit: row = row + 1
    ws.Cells(row, 1).Value = "AllFund NOT in Credit": ws.Cells(row, 2).Value = gAllFundNotInCredit: row = row + 1
    ws.Cells(row, 1).Value = "CoR Correct in Credit": ws.Cells(row, 2).Value = gCorrect: row = row + 1
    ws.Cells(row, 1).Value = "CoR Mismatched (to rectify)": ws.Cells(row, 2).Value = gMismatch: row = row + 2

    ' --- Regional table ---
    ws.Cells(row, 1).Value = "Regional Stats (based on AllFund 'Business Unit')"
    row = row + 1
    ws.Cells(row, 1).Value = "Region"
    ws.Cells(row, 2).Value = "AllFund Total"
    ws.Cells(row, 3).Value = "AllFund Present in Credit"
    ws.Cells(row, 4).Value = "AllFund NOT in Credit"
    ws.Cells(row, 5).Value = "CoR Correct in Credit"
    ws.Cells(row, 6).Value = "CoR Mismatched (to rectify)"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Font.Bold = True

    Dim startDataRow As Long: startDataRow = row + 1
    row = startDataRow

    Dim regKey As Variant
    Dim allRegions As Object: Set allRegions = CreateObject("Scripting.Dictionary")
    For i = LBound(REGIONS) To UBound(REGIONS): allRegions(REGIONS(i)) = True: Next i
    For Each regKey In regAllFundTotal.Keys: allRegions(regKey) = True: Next regKey

    For Each regKey In allRegions.Keys
        ws.Cells(row, 1).Value = CStr(regKey)
        ws.Cells(row, 2).Value = NzLng(regAllFundTotal, regKey)
        ws.Cells(row, 3).Value = NzLng(regAllFundInCredit, regKey)
        ws.Cells(row, 4).Value = NzLng(regAllFundNotInCredit, regKey)
        ws.Cells(row, 5).Value = NzLng(regCorrect, regKey)
        ws.Cells(row, 6).Value = NzLng(regMismatch, regKey)
        row = row + 1
    Next regKey

    ' Format as tables
    Dim globalLo As ListObject, regionalLo As ListObject
    Set globalLo = ws.ListObjects.Add(xlSrcRange, ws.Range("A2:B8"), , xlYes)
    On Error Resume Next: globalLo.Name = "GlobalStatsTbl": On Error GoTo 0

    Dim lastRegRow As Long: lastRegRow = row - 1
    If lastRegRow >= startDataRow Then
        Set regionalLo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startDataRow - 1, 1), ws.Cells(lastRegRow, 6)), , xlYes)
        On Error Resume Next: regionalLo.Name = "RegionalStatsTbl": On Error GoTo 0
    End If

    ws.Columns.AutoFit
End Sub

Private Function NzLng(ByVal dict As Object, ByVal key As Variant) As Long
    If dict.Exists(key) Then NzLng = CLng(dict(key)) Else NzLng = 0
End Function

Private Function RegionFromBU(ByVal bu As String) As String
    ' Normalize common variants (spaces/hyphens/case)
    Dim s As String: s = UCase$(Replace(Trim$(bu), "  ", " "))
    s = Replace(s, "- ", "-")
    Select Case s
        Case "FI-US": RegionFromBU = "AMRS"
        Case "FI-EMEA": RegionFromBU = "EMEA"
        Case "FI-GMC-ASIA", "FI-GMC- ASIA", "FI-GMC- ASIA ":
            RegionFromBU = "APAC"
        Case Else
            RegionFromBU = ""
    End Select
End Function

Private Function BuildCoperToRegionMap(ByVal lo As ListObject, ByVal coperCol As String, ByVal buCol As String) As Object
    Dim idxC As Long, idxBU As Long, dict As Object, r As Long
    idxC = GetColumnIndex(lo, coperCol)
    idxBU = GetColumnIndex(lo, buCol)
    If idxC = 0 Then Err.Raise vbObjectError + 902, , "Column '" & coperCol & "' not found in AllFund."
    If idxBU = 0 Then Err.Raise vbObjectError + 903, , "Column '" & buCol & "' not found in AllFund."
    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            dict(SanitizeCoperID(lo.DataBodyRange.Cells(r, idxC).Value)) = Trim$(CStr(lo.DataBodyRange.Cells(r, idxBU).Value))
        Next r
    End If
    Set BuildCoperToRegionMap = dict
End Function

'========================
' Sanitizers
'========================
Private Function SanitizeCoperID(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    s = Replace(s, Chr$(160), "")
    SanitizeCoperID = s
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
    Dim fd As FileDialog, i As Long, c As New Collection
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = promptTitle
        .Filters.Clear
        .Filters.Add filterDesc, filterPattern
        .AllowMultiSelect = True
        If .Show = -1 Then
            For i = 1 To .SelectedItems.Count
                c.Add .SelectedItems(i)
            Next i
        End If
    End With
    Set PickFilesMulti = c
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
' Append (multi Credit Studio) with sanitization
'========================
Private Sub EnsureRecaliHeaders(ByVal ws As Worksheet)
    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Cells(1, 1).Value = "Coper ID"
        ws.Cells(1, 2).Value = "Country of Risk"
        ws.Cells(1, 3).Value = "Source File"
    End If
End Sub

Private Function NextWriteRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < 1 Then NextWriteRow = 1 Else NextWriteRow = r + 1
End Function

Private Sub AppendColumnsByName(ByVal lo As ListObject, _
                                ByVal wsDest As Worksheet, _
                                ByVal field1 As String, _
                                ByVal field2 As String, _
                                ByVal sourceTag As String)
    Dim idx1 As Long, idx2 As Long, r As Long, outRow As Long
    idx1 = GetColumnIndex(lo, field1)
    idx2 = GetColumnIndex(lo, field2)
    If idx1 = 0 Then Err.Raise 2001, , "Column '" & field1 & "' not found."
    If idx2 = 0 Then Err.Raise 2002, , "Column '" & field2 & "' not found."
    If lo.DataBodyRange Is Nothing Then Exit Sub

    EnsureRecaliHeaders wsDest
    outRow = NextWriteRow(wsDest)

    For r = 1 To lo.DataBodyRange.Rows.Count
        wsDest.Cells(outRow, 1).NumberFormat = "@"
        wsDest.Cells(outRow, 1).Value = SanitizeCoperID(lo.DataBodyRange.Cells(r, idx1).Value)

        wsDest.Cells(outRow, 2).NumberFormat = "@"
        wsDest.Cells(outRow, 2).Value = Trim$(CStr(lo.DataBodyRange.Cells(r, idx2).Value))

        wsDest.Cells(outRow, 3).NumberFormat = "@"
        wsDest.Cells(outRow, 3).Value = sourceTag

        outRow = outRow + 1
    Next r
End Sub

'========================
' Matching + mismatch summary (AllFund is source of truth)
'========================
Private Function BuildCoperToCoRMap(ByVal lo As ListObject, ByVal coperCol As String, ByVal corCol As String) As Object
    Dim idxC As Long, idxR As Long
    Dim dict As Object, r As Long
    Dim vCoper As String, vCoR As String

    idxC = GetColumnIndex(lo, coperCol)
    idxR = GetColumnIndex(lo, corCol)
    If idxC = 0 Then Err.Raise 1005, , "Column '" & coperCol & "' not found in AllFund."
    If idxR = 0 Then Err.Raise 1006, , "Column '" & corCol & "' not found in AllFund."

    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            vCoper = SanitizeCoperID(lo.DataBodyRange.Cells(r, idxC).Value)
            vCoR = Trim$(CStr(lo.DataBodyRange.Cells(r, idxR).Value))
            If Len(vCoper) > 0 Then dict(vCoper) = vCoR
        Next r
    End If
    Set BuildCoperToCoRMap = dict
End Function

Private Function BuildCoperToRegionMap(ByVal lo As ListObject, ByVal coperCol As String, ByVal buCol As String) As Object
    Dim idxC As Long, idxBU As Long, dict As Object, r As Long
    idxC = GetColumnIndex(lo, coperCol)
    idxBU = GetColumnIndex(lo, buCol)
    If idxC = 0 Then Err.Raise vbObjectError + 902, , "Column '" & coperCol & "' not found in AllFund."
    If idxBU = 0 Then Err.Raise vbObjectError + 903, , "Column '" & buCol & "' not found in AllFund."
    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            dict(SanitizeCoperID(lo.DataBodyRange.Cells(r, idxC).Value)) = Trim$(CStr(lo.DataBodyRange.Cells(r, idxBU).Value))
        Next r
    End If
    Set BuildCoperToRegionMap = dict
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

    ws.Cells(1, lastCol + 1).NumberFormat = "@"
    ws.Cells(1, lastCol + 1).Value = outColName

    For r = 2 To lastRow
        key = SanitizeCoperID(ws.Cells(r, coperCol).Value)
        If Len(key) > 0 And approvedMap.Exists(key) Then
            ws.Cells(r, lastCol + 1).NumberFormat = "@"
            ws.Cells(r, lastCol + 1).Value = approvedMap(key)
        Else
            ws.Cells(r, lastCol + 1).ClearContents
        End If
    Next r
End Sub

Private Sub CreateMismatchSummary(ByVal wb As Workbook, ByVal loRecali As ListObject, _
                                  ByVal creditCoRColName As String, _
                                  ByVal approvedCoRColName As String, _
                                  ByVal coperColName As String, _
                                  ByVal summarySheetName As String)

    Dim idxCredit As Long, idxApproved As Long, idxCoper As Long
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    Dim valCredit As String, valApproved As String, coper As String
    Dim wsSum As Worksheet, lo As ListObject
    Dim k As Variant, rowOut As Long, joined As String

    idxCredit = GetColumnIndex(loRecali, creditCoRColName)
    idxApproved = GetColumnIndex(loRecali, approvedCoRColName)
    idxCoper = GetColumnIndex(loRecali, coperColName)
    If idxCredit * idxApproved * idxCoper = 0 Then
        Err.Raise 1008, , "Required columns missing in CoR Recali."
    End If

    If Not loRecali.DataBodyRange Is Nothing Then
        For r = 1 To loRecali.DataBodyRange.Rows.Count
            valCredit = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxCredit).Value))
            valApproved = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxApproved).Value))
            coper = SanitizeCoperID(loRecali.DataBodyRange.Cells(r, idxCoper).Value)

            If Len(coper) > 0 Then
                If StrComp(valCredit, valApproved, vbTextCompare) <> 0 Then
                    If Not dict.Exists(valApproved) Then dict.Add valApproved, New Collection
                    dict(valApproved).Add coper
                End If
            End If
        Next r
    End If

    If dict.Count = 0 Then
        If SheetExists(wb, summarySheetName) Then SafeDeleteSheet wb.Worksheets(summarySheetName)
        Exit Sub
    End If

    Set wsSum = EnsureSheet(wb, summarySheetName, True)
    wsSum.Range("A1").NumberFormat = "@"
    wsSum.Range("B1").NumberFormat = "@"
    wsSum.Range("A1").Value = "Country of Risk (AllFund)"
    wsSum.Range("B1").Value = "Coper IDs with wrong CoR in Credit Studio (comma-joined)"

    rowOut = 2
    For Each k In dict.Keys
        joined = JoinUniqueFromCollection(dict(k), ",")
        wsSum.Cells(rowOut, 1).NumberFormat = "@"
        wsSum.Cells(rowOut, 2).NumberFormat = "@"
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
    Dim i As Long, valStr As String, s As String
    For i = 1 To col.Count
        valStr = SanitizeCoperID(col(i))
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