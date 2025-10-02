Option Explicit

' ====== USER CONFIG ======
Private Const ITERATION_DIR As String = "C:\Your\Folder\Here"   ' <-- SET THIS (no trailing slash)

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
' Entry point (ALLFUND)
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
    Dim pendingAppendPath As String

    ' 1) Pick AllFund CSV
    allFundPath = PickFile("Select ALLFUND CSV", "CSV Files (*.csv)", "*.csv")
    If Len(allFundPath) = 0 Then MsgBox "Operation cancelled.", vbInformation: Exit Sub

    ' 1a) Ask about appending previous Pending
    If MsgBox("Append previous month's PENDING file to AllFund before proceeding?", vbQuestion + vbYesNo, "Append Pending?") = vbYes Then
        pendingAppendPath = PickFile("Select PENDING_MMM-YYYY.xlsx to append", "Excel Files (*.xlsx)", "*.xlsx")
    End If

    ' 2) Open & delete first row (second row has headers)
    Set wbAllFund = Workbooks.Open(Filename:=allFundPath, Local:=True)
    wbAllFund.Worksheets(1).Rows(1).Delete

    ' 3) Table-ize
    Set loAllFund = EnsureTable(wbAllFund.Worksheets(1), "AllFundTbl")

    ' (Optional) Append Pending into AllFund table now
    If Len(pendingAppendPath) > 0 Then
        AppendPendingIntoAllFund loAllFund, pendingAppendPath
        Set loAllFund = EnsureTable(wbAllFund.Worksheets(1), "AllFundTbl") ' refresh pointer
    End If

    ' 4) Filter: BU in (FI-US, FI-EMEA, FI-GMC- Asia/ASIA) AND Review Status in (Approved, Submitted)
    FilterAllFundCriteria loAllFund, _
        keepBUs:=Array("FI-US", "FI-EMEA", "FI-GMC- Asia", "FI-GMC-ASIA"), _
        keepReviewStatus:=Array("Approved", "Submitted")

    ' 4a) Remove blank CoR rows & export region files (all columns); store counters
    Call RemoveAndExportBlankCoR(loAllFund, OutputFolderPrevMonth(), wbMain)

    ' 5) Build Fund CoPER batches of 600 and show modeless controller
    BuildCoperBatches loAllFund, "Fund CoPER", 600
    G_WorkflowReadyToContinue = False
    frmCoperBatches.Show vbModeless

    Exit Sub

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    On Error Resume Next
    If Not wbAllFund Is Nothing Then wbAllFund.Close SaveChanges:=False ' keep CSV untouched
    On Error GoTo 0
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
    Dim allFundCoRMap As Object, allFundRegionMap As Object, allFundStatusMap As Object
    Dim wbAllFund As Workbook, loAllFund As ListObject
    Dim blanksCounters As Object
    Dim keysMap As Object

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
        GoTo Cleanup
    End If

    ' 6) Dated sheet (in Main; moved later to iteration file)
    Set wsDate = CreateDatedSheet(wbMain)

    ' 7) Prepare CoR Recali (recreate fresh; we will append)
    Set wsRecali = EnsureSheet(wbMain, "CoR Recali", True)
    EnsureRecaliHeaders wsRecali

    ' Append all selected Credit Studio files
    For Each p In creditFiles
        Set wbCredit = Workbooks.Open(Filename:=CStr(p), ReadOnly:=True)
        Set loCredit = EnsureTable(wbCredit.Worksheets(1), "CreditTbl")
        AppendColumnsByName loCredit, wsRecali, "Coper ID", "Country of Risk", VBA.Dir(CStr(p))
        wbCredit.Close SaveChanges:=False
    Next p

    ' 8) Build AllFund maps (truth) + append Approved CoR
    Set allFundCoRMap = BuildCoperToCoRMap(loAllFund, "Fund CoPER", "Country of Risk")
    Set allFundRegionMap = BuildCoperToRegionMap(loAllFund, "Fund CoPER", "Business Unit")
    Set allFundStatusMap = BuildCoperToStatusMap(loAllFund, "Fund CoPER", "Review Status")
    AppendApprovedCoR wsRecali, allFundCoRMap, "Coper ID", "Approved CoR"

    ' Keys map for CoR normalization
    Set keysMap = LoadKeysMap(wbMain) ' empty dict if not present

    ' 9) Convert to table and build mismatch summary (Keys normalization)
    Set loRecali = EnsureTable(wsRecali, "CoRRecaliTbl")
    CreateMismatchSummary wbMain, loRecali, _
        creditCoRColName:="Country of Risk", _
        approvedCoRColName:="Approved CoR", _
        coperColName:="Coper ID", _
        summarySheetName:="CoR Mismatch Summary", _
        keysMap:=keysMap

    ' 10) Create Iteration file; move sheets + build Stats (includes blanks + review status)
    Dim iterWb As Workbook
    Set iterWb = CreateIterationWorkbook()
    CopySheetValues wbMain.Worksheets(wsDate.Name), iterWb, wsDate.Name
    CopySheetValues wbMain.Worksheets("CoR Recali"), iterWb, "CoR Recali"
    If SheetExists(wbMain, "CoR Mismatch Summary") Then
        CopySheetValues wbMain.Worksheets("CoR Mismatch Summary"), iterWb, "CoR Mismatch Summary"
    End If

    ' Rehydrate blanks counters saved during phase 1
    Set blanksCounters = GetBlanksCountersFromHiddenName(wbMain)

    ' Build Stats
    BuildStatsSheet iterWb, iterWb.Worksheets("CoR Recali"), allFundCoRMap, allFundRegionMap, allFundStatusMap, blanksCounters

    ' 11) Save Pending (AllFund copers NOT found in Credit Studio)
    ExportPendingNotInCredit loAllFund, iterWb, OutputFolderPrevMonth()

    ' Save iteration file
    SaveIterationWorkbook iterWb, OutputFolderPrevMonth()

Cleanup:
    ' 12) Close AllFund WITHOUT saving (leave CSV unchanged)
    On Error Resume Next
    If Not wbAllFund Is Nothing Then wbAllFund.Close SaveChanges:=False
    On Error GoTo 0

    MsgBox "Done. Outputs saved under: " & OutputFolderPrevMonth(), vbInformation
    Exit Sub

Fail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
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
' Append Pending into AllFund
'========================
Private Sub AppendPendingIntoAllFund(ByRef loAllFund As ListObject, ByVal pendingPath As String)
    Dim wbP As Workbook, loP As ListObject
    Set wbP = Workbooks.Open(Filename:=pendingPath, ReadOnly:=True)
    Set loP = EnsureTable(wbP.Worksheets(1), "PendingTbl")

    ' Map matching headers
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim i As Long, idx As Long
    For i = 1 To loP.HeaderRowRange.Columns.Count
        idx = FindHeaderIndexInTable(loAllFund, CStr(loP.HeaderRowRange.Cells(1, i).Value))
        If idx > 0 Then map.Add i, idx
    Next i

    ' Append rows
    If Not loP.DataBodyRange Is Nothing Then
        Dim r As Long, newRow As ListRow, cSrc As Variant
        For r = 1 To loP.DataBodyRange.Rows.Count
            Set newRow = loAllFund.ListRows.Add
            For Each cSrc In map.Keys
                loAllFund.DataBodyRange.Cells(newRow.Index, map(cSrc)).Value = loP.DataBodyRange.Cells(r, cSrc).Value
            Next cSrc
        Next r
    End If

    wbP.Close SaveChanges:=False
End Sub

Private Function FindHeaderIndexInTable(ByVal lo As ListObject, ByVal header As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If LCase$(Trim$(CStr(lo.HeaderRowRange.Cells(1, i).Value))) = LCase$(Trim$(header)) Then
            FindHeaderIndexInTable = i: Exit Function
        End If
    Next
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

    On Error Resume Next
    lo.Range.AutoFilter Field:=buCol, Criteria1:=keepBUs, Operator:=xlFilterValues
    lo.Range.AutoFilter Field:=revCol, Criteria1:=keepReviewStatus, Operator:=xlFilterValues
    On Error GoTo 0

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
' Remove blank CoR rows & export region files (ALL columns)
'========================
Private Sub RemoveAndExportBlankCoR(ByRef loAllFund As ListObject, ByVal outFolder As String, ByVal wbMain As Workbook)
    Dim idxCoR As Long, idxBU As Long, idxStatus As Long
    idxCoR = GetColumnIndex(loAllFund, "Country of Risk")
    idxBU = GetColumnIndex(loAllFund, "Business Unit")
    idxStatus = GetColumnIndex(loAllFund, "Review Status")
    If idxCoR * idxBU * idxStatus = 0 Then Err.Raise vbObjectError + 1200, , "AllFund: required columns missing (Country of Risk, Business Unit, Review Status)."

    Dim headers As Variant: headers = loAllFund.HeaderRowRange.Value

    ' Collect blank rows
    Dim amrs As Collection, emea As Collection, apac As Collection
    Set amrs = New Collection: Set emea = New Collection: Set apac = New Collection

    Dim r As Long, bu As String, corv As String, rowVals As Variant
    If Not loAllFund.DataBodyRange Is Nothing Then
        For r = 1 To loAllFund.DataBodyRange.Rows.Count
            corv = Trim$(CStr(loAllFund.DataBodyRange.Cells(r, idxCoR).Value))
            If Len(corv) = 0 Then
                bu = UCase$(Trim$(CStr(loAllFund.DataBodyRange.Cells(r, idxBU).Value)))
                rowVals = loAllFund.DataBodyRange.Rows(r).Value
                Select Case RegionFromBU(bu)
                    Case "AMRS": amrs.Add rowVals
                    Case "EMEA": emea.Add rowVals
                    Case "APAC": apac.Add rowVals
                    Case Else: emea.Add rowVals
                End Select
            End If
        Next r
    End If

    ' Export region files if any blank rows
    If amrs.Count + emea.Count + apac.Count > 0 Then
        ExportBlankRegion outFolder, "AMRS", headers, amrs, loAllFund
        ExportBlankRegion outFolder, "EMEA", headers, emea, loAllFund
        ExportBlankRegion outFolder, "APAC", headers, apac, loAllFund

        ' Remove blanks from AllFund table (iterate bottom-up)
        Dim last As Long: last = loAllFund.DataBodyRange.Rows.Count
        For r = last To 1 Step -1
            If Len(Trim$(CStr(loAllFund.DataBodyRange.Cells(r, idxCoR).Value))) = 0 Then
                loAllFund.ListRows(r).Delete
            End If
        Next r
    End If

    ' Count blanks global/region/status using captured rows
    Dim counters As Object: Set counters = CreateObject("Scripting.Dictionary")
    counters("AMRS_Approved") = 0: counters("AMRS_Submitted") = 0
    counters("EMEA_Approved") = 0: counters("EMEA_Submitted") = 0
    counters("APAC_Approved") = 0: counters("APAC_Submitted") = 0

    CountBlankCollection amrs, headers, idxBU, idxStatus, counters
    CountBlankCollection emea, headers, idxBU, idxStatus, counters
    CountBlankCollection apac, headers, idxBU, idxStatus, counters

    ' Persist counters (overwrite if already present)
    SaveBlanksCountersToHiddenName wbMain, counters
End Sub

Private Sub CountBlankCollection(ByVal col As Collection, ByVal headers As Variant, ByVal idxBU As Long, ByVal idxStatus As Long, ByRef counters As Object)
    Dim i As Long, bu As String, st As String, reg As String, key As String
    For i = 1 To col.Count
        bu = UCase$(Trim$(CStr(col(i)(1, idxBU))))
        st = ProperStatus(CStr(col(i)(1, idxStatus)))
        reg = RegionFromBU(bu)
        key = reg & "_" & st
        If Not counters.Exists(key) Then counters(key) = 0
        counters(key) = CLng(counters(key)) + 1
    Next i
End Sub

Private Sub ExportBlankRegion(ByVal outFolder As String, ByVal region As String, ByVal headers As Variant, ByVal rowsCol As Collection, ByVal templateLO As ListObject)
    If rowsCol.Count = 0 Then Exit Sub
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    On Error Resume Next: ws.Name = region & "_Blank_CoR": On Error GoTo 0

    ' Write headers
    ws.Range("A1").Resize(1, templateLO.HeaderRowRange.Columns.Count).Value = headers

    ' Write rows
    Dim i As Long
    For i = 1 To rowsCol.Count
        ws.Range("A" & (i + 1)).Resize(1, UBound(rowsCol(i), 2)).Value = rowsCol(i)
    Next i

    Set lo = EnsureTable(ws, region & "_BlankTbl")
    SaveWorkbook wb, outFolder & "\" & region & "_" & PrevMonthFileName()
    wb.Close SaveChanges:=False
End Sub

'========================
' Keys Table / Normalization
'========================
Private Function LoadKeysMap(ByVal wbMain As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    On Error GoTo Done
    Dim ws As Worksheet, lo As ListObject
    Set ws = wbMain.Worksheets("Keys")
    Set lo = ws.ListObjects("Keys")

    Dim idxAF As Long, idxCS As Long, r As Long
    idxAF = GetListColumnIndex(lo, "All Funds")
    idxCS = GetListColumnIndex(lo, "Credit Studio")
    If idxAF = 0 Or idxCS = 0 Then GoTo Done

    Dim af As String, cs As String, canon As String
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            af = LCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, idxAF).Value)))
            cs = LCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, idxCS).Value)))
            If Len(af) > 0 And Len(cs) > 0 Then
                canon = af ' use AllFund spelling as canonical
                dict(af) = canon
                dict(cs) = canon
            End If
        Next r
    End If
Done:
    Set LoadKeysMap = dict
End Function

Private Function NormalizeCoR(ByVal v As String, ByVal keysMap As Object) As String
    Dim k As String: k = LCase$(Trim$(v))
    If Len(k) = 0 Then NormalizeCoR = k: Exit Function
    If Not keysMap Is Nothing Then
        If keysMap.Exists(k) Then NormalizeCoR = keysMap(k): Exit Function
    End If
    NormalizeCoR = k
End Function

Private Function GetListColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If LCase$(Trim$(CStr(lo.HeaderRowRange.Cells(1, i).Value))) = LCase$(Trim$(headerName)) Then
            GetListColumnIndex = i: Exit Function
        End If
    Next i
End Function

'========================
' Iteration file helpers & output folder (prev month)
'========================
Private Function OutputFolderPrevMonth() As String
    Dim subf As String: subf = PrevMonthFolderName()
    Dim full As String: full = ITERATION_DIR
    If Right$(full, 1) = "\" Or Right$(full, 1) = "/" Then full = Left$(full, Len(full) - 1)
    full = full & "\" & subf
    EnsureFolder full
    OutputFolderPrevMonth = full
End Function

Private Function PrevMonthFileName() As String
    Dim d As Date, s As String
    d = DateSerial(Year(Date), Month(Date), 1) - 1   ' last day of prev month
    s = Format(d, "mmmm-yyyy") & ".xlsx"             ' e.g., September-2025.xlsx
    PrevMonthFileName = s
End Function

Private Function PrevMonthFolderName() As String
    Dim d As Date
    d = DateSerial(Year(Date), Month(Date), 1) - 1
    PrevMonthFolderName = Format(d, "mmmm-yyyy")     ' e.g., September-2025
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Sub SaveWorkbook(ByVal wb As Workbook, ByVal path As String)
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook ' .xlsx
    Application.DisplayAlerts = True
End Sub

Private Function CreateIterationWorkbook() As Workbook
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add(xlWBATWorksheet) ' single sheet
    wb.Worksheets(1).Name = "Stats"
    Set CreateIterationWorkbook = wb
End Function

Private Sub CopySheetValues(ByVal srcWs As Worksheet, ByVal destWb As Workbook, ByVal newName As String)
    Dim ws As Worksheet
    Set ws = destWb.Worksheets.Add(After:=destWb.Worksheets(destWb.Worksheets.Count))
    On Error Resume Next: ws.Name = newName: On Error GoTo 0
    srcWs.UsedRange.Copy
    ws.Range("A1").PasteSpecial xlPasteValues
    ws.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Sub SaveIterationWorkbook(ByVal wb As Workbook, ByVal outFolder As String)
    SaveWorkbook wb, outFolder & "\Iteration_" & PrevMonthFileName()
End Sub

' ===============================
' ENHANCED STATS (Formatting + %)
' ===============================
Private Sub BuildStatsSheet(ByVal iterWb As Workbook, _
                            ByVal recaliWs As Worksheet, _
                            ByVal allFundCoRMap As Object, _
                            ByVal allFundRegionMap As Object, _
                            ByVal allFundStatusMap As Object, _
                            ByVal blanksCounters As Object)

    Dim ws As Worksheet: Set ws = iterWb.Worksheets("Stats")
    ws.Cells.Clear

    ' ---- Fonts & layout defaults ----
    With ws.Parent
        ' keep defaults; we’ll format section headers and tables explicitly
    End With

    Dim row As Long: row = 1
    Dim colCoper As Long, colCreditCoR As Long
    Dim r As Long

    ' Build credit dict (unique Copers -> Credit CoR)
    Dim creditDict As Object: Set creditDict = CreateObject("Scripting.Dictionary")
    colCoper = FindHeader(recaliWs, "Coper ID")
    colCreditCoR = FindHeader(recaliWs, "Country of Risk")
    If colCoper = 0 Or colCreditCoR = 0 Then Err.Raise vbObjectError + 800, , "CoR Recali missing columns."

    Dim cLastRow As Long: cLastRow = recaliWs.Cells(recaliWs.Rows.Count, 1).End(xlUp).Row
    For r = 2 To cLastRow
        Dim cid As String, cCoR As String
        cid = SanitizeCoperID(recaliWs.Cells(r, colCoper).Value)
        If Len(cid) > 0 Then
            cCoR = Trim$(CStr(recaliWs.Cells(r, colCreditCoR).Value))
            creditDict(cid) = cCoR   ' last occurrence wins
        End If
    Next r

    ' --- counters ---
    Dim STATUSES As Variant: STATUSES = Array("Approved", "Submitted")

    Dim gTotal As Long, gCreditTotal As Long, gIn As Long, gOut As Long, gCor As Long, gMis As Long
    Dim gTotalS As Object, gInS As Object, gOutS As Object, gCorS As Object, gMisS As Object
    Set gTotalS = CreateObject("Scripting.Dictionary")
    Set gInS = CreateObject("Scripting.Dictionary")
    Set gOutS = CreateObject("Scripting.Dictionary")
    Set gCorS = CreateObject("Scripting.Dictionary")
    Set gMisS = CreateObject("Scripting.Dictionary")

    Dim s As Variant
    For Each s In STATUSES
        gTotalS(s) = 0: gInS(s) = 0: gOutS(s) = 0: gCorS(s) = 0: gMisS(s) = 0
    Next s

    ' Regional by status: dict(reg) -> dict(status)->count
    Dim regTotal As Object, regIn As Object, regOut As Object, regCor As Object, regMis As Object
    Set regTotal = CreateObject("Scripting.Dictionary")
    Set regIn = CreateObject("Scripting.Dictionary")
    Set regOut = CreateObject("Scripting.Dictionary")
    Set regCor = CreateObject("Scripting.Dictionary")
    Set regMis = CreateObject("Scripting.Dictionary")

    Dim key As Variant, reg As String, st As String

    ' Iterate AllFund
    For Each key In allFundCoRMap.Keys
        reg = RegionFromBU(allFundRegionMap(key)): If Len(reg) = 0 Then reg = "Unmapped"
        st = ProperStatus(allFundStatusMap(key))

        gTotal = gTotal + 1
        gTotalS(st) = CLng(gTotalS(st)) + 1
        InitRegStat regTotal, reg: regTotal(reg)(st) = CLng(regTotal(reg)(st)) + 1

        If creditDict.Exists(key) Then
            gIn = gIn + 1
            gInS(st) = CLng(gInS(st)) + 1
            InitRegStat regIn, reg: regIn(reg)(st) = CLng(regIn(reg)(st)) + 1

            If StrComp(creditDict(key), allFundCoRMap(key), vbTextCompare) = 0 Then
                gCor = gCor + 1
                gCorS(st) = CLng(gCorS(st)) + 1
                InitRegStat regCor, reg: regCor(reg)(st) = CLng(regCor(reg)(st)) + 1
            Else
                gMis = gMis + 1
                gMisS(st) = CLng(gMisS(st)) + 1
                InitRegStat regMis, reg: regMis(reg)(st) = CLng(regMis(reg)(st)) + 1
            End If
        Else
            gOut = gOut + 1
            gOutS(st) = CLng(gOutS(st)) + 1
            InitRegStat regOut, reg: regOut(reg)(st) = CLng(regOut(reg)(st)) + 1
        End If
    Next key

    gCreditTotal = creditDict.Count

    ' =========================
    ' SECTION 1: GLOBAL SUMMARY
    ' =========================
    Call WriteSectionHeader(ws, row, "Global Summary")
    row = row + 1

    ' Headers
    ws.Cells(row, 1).Value = "Metric"
    ws.Cells(row, 2).Value = "Count"
    ws.Cells(row, 3).Value = "% of AllFund"
    ws.Cells(row, 4).Value = "Notes"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Font.Bold = True
    row = row + 1

    Dim startGlobal As Long: startGlobal = row
    ws.Cells(row, 1).Value = "AllFund Total": ws.Cells(row, 2).Value = gTotal: ws.Cells(row, 3).Value = 1: ws.Cells(row, 4).Value = "Denominator for %": row = row + 1
    ws.Cells(row, 1).Value = "Credit Total (unique Copers)":          ws.Cells(row, 2).Value = gCreditTotal: ws.Cells(row, 4).Value = "De-duplicated from Credit files": row = row + 1
    ws.Cells(row, 1).Value = "AllFund Present in Credit":             ws.Cells(row, 2).Value = gIn:        ws.Cells(row, 3).Value = SafePercent(gIn, gTotal):         ws.Cells(row, 4).Value = "Coverage":                         row = row + 1
    ws.Cells(row, 1).Value = "AllFund NOT in Credit":                 ws.Cells(row, 2).Value = gOut:       ws.Cells(row, 3).Value = SafePercent(gOut, gTotal):        ws.Cells(row, 4).Value = "Gap":                              row = row + 1
    ws.Cells(row, 1).Value = "CoR Correct in Credit":                 ws.Cells(row, 2).Value = gCor:       ws.Cells(row, 3).Value = SafePercent(gCor, gTotal):        ws.Cells(row, 4).Value = "Correct vs AllFund":               row = row + 1
    ws.Cells(row, 1).Value = "CoR Mismatched (to rectify)":           ws.Cells(row, 2).Value = gMis:       ws.Cells(row, 3).Value = SafePercent(gMis, gTotal):        ws.Cells(row, 4).Value = "Mismatch vs AllFund":              row = row + 1

    Dim loGlobal As ListObject
    Set loGlobal = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startGlobal - 1, 1), ws.Cells(row - 1, 4)), , xlYes)
    loGlobal.Name = "GlobalStatsTbl": ApplyModernTableStyle loGlobal
    ws.Range(ws.Cells(startGlobal, 3), ws.Cells(row - 1, 3)).NumberFormat = "0.0%"

    row = row + 2

    ' =======================================
    ' SECTION 2: GLOBAL BY REVIEW STATUS (S)
    ' =======================================
    Call WriteSectionHeader(ws, row, "Global by Review Status")
    row = row + 1

    ws.Cells(row, 1).Value = "Status"
    ws.Cells(row, 2).Value = "Total"
    ws.Cells(row, 3).Value = "Present in Credit"
    ws.Cells(row, 4).Value = "NOT in Credit"
    ws.Cells(row, 5).Value = "CoR Correct"
    ws.Cells(row, 6).Value = "CoR Mismatch"
    ws.Cells(row, 7).Value = "Present % of Status"
    ws.Cells(row, 8).Value = "Mismatch % of Status"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Font.Bold = True
    row = row + 1

    Dim startGStatus As Long: startGStatus = row
    For Each s In STATUSES
        ws.Cells(row, 1).Value = s
        ws.Cells(row, 2).Value = NzLng2(gTotalS, s)
        ws.Cells(row, 3).Value = NzLng2(gInS, s)
        ws.Cells(row, 4).Value = NzLng2(gOutS, s)
        ws.Cells(row, 5).Value = NzLng2(gCorS, s)
        ws.Cells(row, 6).Value = NzLng2(gMisS, s)
        ws.Cells(row, 7).Value = SafePercent(NzLng2(gInS, s), NzLng2(gTotalS, s))
        ws.Cells(row, 8).Value = SafePercent(NzLng2(gMisS, s), NzLng2(gTotalS, s))
        row = row + 1
    Next s

    Dim loGStatus As ListObject
    Set loGStatus = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startGStatus - 1, 1), ws.Cells(row - 1, 8)), , xlYes)
    loGStatus.Name = "GlobalByStatusTbl": ApplyModernTableStyle loGStatus
    ws.Range(ws.Cells(startGStatus, 7), ws.Cells(row - 1, 8)).NumberFormat = "0.0%"

    row = row + 2

    ' ===========================================
    ' SECTION 3: REGIONAL BY REVIEW STATUS (R,S)
    ' ===========================================
    Call WriteSectionHeader(ws, row, "Regional by Review Status")
    row = row + 1

    ws.Cells(row, 1).Value = "Region"
    ws.Cells(row, 2).Value = "Status"
    ws.Cells(row, 3).Value = "Total"
    ws.Cells(row, 4).Value = "Present in Credit"
    ws.Cells(row, 5).Value = "NOT in Credit"
    ws.Cells(row, 6).Value = "CoR Correct"
    ws.Cells(row, 7).Value = "CoR Mismatch"
    ws.Cells(row, 8).Value = "Present % of Status"
    ws.Cells(row, 9).Value = "Mismatch % of Status"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 9)).Font.Bold = True
    row = row + 1

    Dim startReg As Long: startReg = row
    Dim regKey As Variant, statusKey As Variant
    For Each regKey In UnionKeys(regTotal, regIn, regOut, regCor, regMis).Keys
        For Each statusKey In STATUSES
            ws.Cells(row, 1).Value = CStr(regKey)
            ws.Cells(row, 2).Value = statusKey
            ws.Cells(row, 3).Value = NzLngNested(regTotal, regKey, statusKey)
            ws.Cells(row, 4).Value = NzLngNested(regIn, regKey, statusKey)
            ws.Cells(row, 5).Value = NzLngNested(regOut, regKey, statusKey)
            ws.Cells(row, 6).Value = NzLngNested(regCor, regKey, statusKey)
            ws.Cells(row, 7).Value = NzLngNested(regMis, regKey, statusKey)
            ws.Cells(row, 8).Value = SafePercent(NzLngNested(regIn, regKey, statusKey), NzLngNested(regTotal, regKey, statusKey))
            ws.Cells(row, 9).Value = SafePercent(NzLngNested(regMis, regKey, statusKey), NzLngNested(regTotal, regKey, statusKey))
            row = row + 1
        Next statusKey
    Next regKey

    If row > startReg Then
        Dim loReg As ListObject
        Set loReg = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startReg - 1, 1), ws.Cells(row - 1, 9)), , xlYes)
        loReg.Name = "RegionalByStatusTbl": ApplyModernTableStyle loReg
        ws.Range(ws.Cells(startReg, 8), ws.Cells(row - 1, 9)).NumberFormat = "0.0%"
    End If

    row = row + 2

    ' ===========================
    ' SECTION 4: BLANK CoR SPLIT
    ' ===========================
    Call WriteSectionHeader(ws, row, "Blank CoR in AllFund (exported before processing)")
    row = row + 1

    ws.Cells(row, 1).Value = "Region"
    ws.Cells(row, 2).Value = "Approved"
    ws.Cells(row, 3).Value = "Submitted"
    ws.Cells(row, 4).Value = "Total Blanks"
    ws.Cells(row, 5).Value = "% of Regional AllFund"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Font.Bold = True
    row = row + 1

    Dim regs As Variant: regs = Array("AMRS", "EMEA", "APAC")
    Dim startBlank As Long: startBlank = row
    Dim rr As Long, aVal As Long, sVal As Long, totBlank As Long, regTotalAllFund As Long

    For rr = LBound(regs) To UBound(regs)
        aVal = GetDictLong(blanksCounters, regs(rr) & "_Approved")
        sVal = GetDictLong(blanksCounters, regs(rr) & "_Submitted")
        totBlank = aVal + sVal

        ' Regional allfund = sum of Approved+Submitted in regTotal for that region
        regTotalAllFund = 0
        If regTotal.Exists(regs(rr)) Then
            regTotalAllFund = NzLngNested(regTotal, regs(rr), "Approved") + NzLngNested(regTotal, regs(rr), "Submitted")
        End If

        ws.Cells(row, 1).Value = regs(rr)
        ws.Cells(row, 2).Value = aVal
        ws.Cells(row, 3).Value = sVal
        ws.Cells(row, 4).Value = totBlank
        ws.Cells(row, 5).Value = SafePercent(totBlank, regTotalAllFund)
        row = row + 1
    Next rr

    Dim loBlank As ListObject
    Set loBlank = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startBlank - 1, 1), ws.Cells(row - 1, 5)), , xlYes)
    loBlank.Name = "BlankCoRTbl": ApplyModernTableStyle loBlank
    ws.Range(ws.Cells(startBlank, 5), ws.Cells(row - 1, 5)).NumberFormat = "0.0%"

    ' --- polish ---
    ws.Columns.AutoFit
End Sub

' ---------- helpers used by BuildStatsSheet ----------
Private Sub WriteSectionHeader(ByVal ws As Worksheet, ByVal atRow As Long, ByVal title As String)
    With ws.Cells(atRow, 1)
        .Value = title
        .Font.Bold = True
        .Font.Size = 14
    End With
End Sub

Private Sub ApplyModernTableStyle(ByRef lo As ListObject)
    On Error Resume Next
    lo.TableStyle = "TableStyleMedium9"
    lo.ShowTableStyleRowStripes = True
    lo.ShowTableStyleColumnStripes = False
    On Error GoTo 0
End Sub

Private Function GetDictLong(ByVal dict As Object, ByVal key As Variant) As Long
    If dict Is Nothing Then
        GetDictLong = 0
    ElseIf dict.Exists(key) Then
        GetDictLong = CLng(dict(key))
    Else
        GetDictLong = 0
    End If
End Function

Private Function SafePercent(ByVal num As Long, ByVal den As Long) As Double
    If den <= 0 Then
        SafePercent = 0
    Else
        SafePercent = num / den
    End If
End Function


'========================
' Save Pending (AllFund copers NOT in Credit Studio)
'========================
Private Sub ExportPendingNotInCredit(ByVal loAllFund As ListObject, ByVal iterWb As Workbook, ByVal outFolder As String)
    Dim idxCoper As Long: idxCoper = GetColumnIndex(loAllFund, "Fund CoPER")
    If idxCoper = 0 Then Err.Raise vbObjectError + 1400, , "AllFund: 'Fund CoPER' not found."

    ' Build set of Copers present in Credit (from CoR Recali sheet inside iterWb)
    Dim wsRec As Worksheet: Set wsRec = iterWb.Worksheets("CoR Recali")
    Dim colRecCoper As Long: colRecCoper = FindHeader(wsRec, "Coper ID")
    If colRecCoper = 0 Then Exit Sub
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim lastR As Long, r As Long
    lastR = wsRec.Cells(wsRec.Rows.Count, colRecCoper).End(xlUp).Row
    For r = 2 To lastR
        seen(SanitizeCoperID(wsRec.Cells(r, colRecCoper).Value)) = True
    Next r

    ' Build workbook with rows not in Credit
    Dim wb As Workbook, ws As Worksheet
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1): On Error Resume Next: ws.Name = "Pending": On Error GoTo 0

    ' Headers
    ws.Range("A1").Resize(1, loAllFund.HeaderRowRange.Columns.Count).Value = loAllFund.HeaderRowRange.Value

    ' Rows
    Dim wr As Long: wr = 2
    If Not loAllFund.DataBodyRange Is Nothing Then
        For r = 1 To loAllFund.DataBodyRange.Rows.Count
            Dim afID As String
            afID = SanitizeCoperID(loAllFund.DataBodyRange.Cells(r, idxCoper).Value)
            If Len(afID) > 0 Then
                If Not seen.Exists(afID) Then
                    ws.Range("A" & wr).Resize(1, loAllFund.HeaderRowRange.Columns.Count).Value = loAllFund.DataBodyRange.Rows(r).Value
                    wr = wr + 1
                End If
            End If
        Next r
    End If

    If wr > 2 Then
        EnsureTable ws, "PendingTbl"
        SaveWorkbook wb, outFolder & "\Pending_" & PrevMonthFileName()
    End If
    wb.Close SaveChanges:=False
End Sub

'========================
' Persist/Retrieve blank counters between phases
'========================
Private Sub SaveBlanksCountersToHiddenName(ByVal wb As Workbook, ByVal counters As Object)
    Dim json As String: json = DictToSimpleJSON(counters)
    On Error Resume Next
    wb.Names("__BLANK_COR_COUNTERS__").Delete
    On Error GoTo 0
    wb.Names.Add Name:="__BLANK_COR_COUNTERS__", RefersTo:="=""" & json & """"
End Sub

Private Function GetBlanksCountersFromHiddenName(ByVal wb As Workbook) As Object
    Dim n As Name, json As String
    On Error Resume Next
    Set n = wb.Names("__BLANK_COR_COUNTERS__")
    On Error GoTo 0
    If n Is Nothing Then Set GetBlanksCountersFromHiddenName = CreateObject("Scripting.Dictionary"): Exit Function
    json = n.RefersTo
    json = Replace(json, "=", "")
    json = Replace(json, """", "")
    Set GetBlanksCountersFromHiddenName = SimpleJSONToDict(json)
End Function

Private Function DictToSimpleJSON(ByVal dict As Object) As String
    Dim k As Variant, s As String
    s = "{"
    For Each k In dict.Keys
        If Len(s) > 1 Then s = s & ","
        s = s & k & ":" & dict(k)
    Next k
    s = s & "}"
    DictToSimpleJSON = s
End Function

Private Function SimpleJSONToDict(ByVal s As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    s = Trim$(s)
    If Left$(s, 1) = "{" Then s = Mid$(s, 2)
    If Right$(s, 1) = "}" Then s = Left$(s, Len(s) - 1)
    If Len(s) = 0 Then Set SimpleJSONToDict = d: Exit Function

    Dim parts() As String, i As Long, kv() As String
    parts = Split(s, ",")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), ":")
        If UBound(kv) = 1 Then
            d(Trim$(kv(0))) = CLng(Trim$(kv(1)))
        End If
    Next i
    Set SimpleJSONToDict = d
End Function

'========================
' Matching + mismatch summary (AllFund is truth; Keys normalization)
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

Private Function BuildCoperToStatusMap(ByVal lo As ListObject, ByVal coperCol As String, ByVal stCol As String) As Object
    Dim idxC As Long, idxS As Long, dict As Object, r As Long
    idxC = GetColumnIndex(lo, coperCol)
    idxS = GetColumnIndex(lo, stCol)
    If idxC = 0 Then Err.Raise vbObjectError + 904, , "Column '" & coperCol & "' not found in AllFund."
    If idxS = 0 Then Err.Raise vbObjectError + 905, , "Column '" & stCol & "' not found in AllFund."
    Set dict = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            dict(SanitizeCoperID(lo.DataBodyRange.Cells(r, idxC).Value)) = ProperStatus(CStr(lo.DataBodyRange.Cells(r, idxS).Value))
        Next r
    End If
    Set BuildCoperToStatusMap = dict
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
                                  ByVal summarySheetName As String, _
                                  ByVal keysMap As Object)

    Dim idxCredit As Long, idxApproved As Long, idxCoper As Long
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    Dim valCredit As String, valApproved As String, coper As String
    Dim valCreditN As String, valApprovedN As String
    Dim wsSum As Worksheet, lo As ListObject
    Dim k As Variant, rowOut As Long, joined As String

    idxCredit = GetColumnIndex(loRecali, creditCoRColName)
    idxApproved = GetColumnIndex(loRecali, approvedCoRColName)
    idxCoper = GetColumnIndex(loRecali, coperColName)
    If idxCredit * idxApproved * idxCoper = 0 Then Err.Raise 1008, , "Required columns missing in CoR Recali."

    If Not loRecali.DataBodyRange Is Nothing Then
        For r = 1 To loRecali.DataBodyRange.Rows.Count
            valCredit = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxCredit).Value))
            valApproved = Trim$(CStr(loRecali.DataBodyRange.Cells(r, idxApproved).Value))
            valCreditN = NormalizeCoR(valCredit, keysMap)
            valApprovedN = NormalizeCoR(valApproved, keysMap)

            coper = SanitizeCoperID(loRecali.DataBodyRange.Cells(r, idxCoper).Value)

            If Len(coper) > 0 Then
                If StrComp(valCreditN, valApprovedN, vbTextCompare) <> 0 Then
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
    wsSum.Cells(1, 1).Value = "Country of Risk (AllFund, post-Keys normalization)"
    wsSum.Cells(1, 2).Value = "Coper IDs with wrong CoR in Credit Studio (comma-joined)"

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
    On Error Resume Next: lo.Name = "CoRMismatchSummaryTbl": On Error GoTo 0

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
    If Not IsArrayAllocated(vals) Then Err.Raise vbObjectError + 700, , "No '" & headerName & "' values to batch."

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
' Sheet & table utilities + FindHeader
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

' Find the column index (1-based) of a header on row 1; 0 if not found.
Private Function FindHeader(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long, target As String, val As String
    If ws Is Nothing Then Exit Function
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function

    target = LCase$(Trim$(headerName))
    For c = 1 To lastCol
        val = LCase$(Trim$(CStr(ws.Cells(1, c).Value)))
        If val = target Then
            FindHeader = c
            Exit Function
        End If
    Next c
    FindHeader = 0
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


' ====== Added minimal helpers to resolve compile references (no change to existing workflow logic) ======
Public Sub EnsureRecaliHeaders(ByVal wsRecali As Worksheet)
    'Idempotently ensure headers exist on "CoR Recali" sheet.
    'Will not alter any downstream logic; only guarantees header row.
    Dim headers As Variant
    headers = Array("Coper ID", "Country of Risk", "Source File")
    With wsRecali
        If Application.WorksheetFunction.CountA(.Rows(1)) = 0 Then
            .Cells.Clear
            .Range("A1").Resize(1, UBound(headers) + 1).Value = headers
        End If
    End With
End Sub

Public Sub AppendColumnsByName(ByVal lo As ListObject, _
                               ByVal wsDest As Worksheet, _
                               ByVal colKey1 As String, _
                               ByVal colKey2 As String, _
                               ByVal sourceName As String)
    'Appends two named columns from source ListObject to destination sheet,
    'and writes the sourceName in the 3rd column. This mirrors the intended call sites.
    On Error GoTo CleanExit
    Dim idx1 As Long, idx2 As Long
    idx1 = lo.ListColumns(colKey1).Index
    idx2 = lo.ListColumns(colKey2).Index

    Dim data As Variant
    If lo.DataBodyRange Is Nothing Then GoTo CleanExit
    data = lo.DataBodyRange.Value

    Dim r As Long, destRow As Long
    destRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    If destRow < 1 Then destRow = 1
    ' If headers are in row 1, next write row is +1
    If destRow = 1 And Application.WorksheetFunction.CountA(wsDest.Rows(1)) > 0 Then
        destRow = 1
    End If
    destRow = destRow + 1

    Dim nRows As Long
    If IsArray(data) Then
        nRows = UBound(data, 1)
        For r = 1 To nRows
            wsDest.Cells(destRow, 1).Value = data(r, idx1)
            wsDest.Cells(destRow, 2).Value = data(r, idx2)
            wsDest.Cells(destRow, 3).Value = sourceName
            destRow = destRow + 1
        Next r
    End If

CleanExit:
End Sub


' ====== Added minimal helper to resolve missing reference (keeps existing flow intact) ======
' Map Business Unit to Region (normalizes spacing/hyphens/case)
Private Function RegionFromBU(ByVal bu As String) As String
    Dim s As String
    s = UCase$(Trim$(bu))
    s = Replace(s, "  ", " ")
    s = Replace(s, "- ", "-")    ' normalize "GMC- Asia" -> "GMC-Asia"
    s = Replace(s, " -", "-")
    s = Replace(s, "  ", " ")

    Select Case s
        Case "FI-US"
            RegionFromBU = "AMRS"
        Case "FI-EMEA"
            RegionFromBU = "EMEA"
        Case "FI-GMC-ASIA", "FI-GMC- ASIA", "FI-GMC-ASIA ", "FI-GMC- ASIA "
            RegionFromBU = "APAC"
        Case Else
            RegionFromBU = ""     ' unmapped/unknown
    End Select
End Function


' Returns Long value from a dictionary by key; 0 (or defaultVal) if missing/non-numeric/null.
Public Function NzLng3(ByVal dict As Object, ByVal key As String, Optional ByVal defaultVal As Long = 0) As Long
    On Error GoTo SafeExit
    If dict Is Nothing Then
        NzLng3 = defaultVal
        Exit Function
    End If

    If dict.Exists(key) Then
        If IsNumeric(dict(key)) Then
            NzLng3 = CLng(dict(key))
        Else
            NzLng3 = defaultVal
        End If
    Else
        NzLng3 = defaultVal
    End If
    Exit Function

SafeExit:
    NzLng3 = defaultVal
End Function

' Normalize Review Status to consistent casing used across the stats
Private Function ProperStatus(ByVal s As Variant) As String
    Dim t As String
    t = UCase$(Trim$(CStr(s)))
    Select Case t
        Case "APPROVED":  ProperStatus = "Approved"
        Case "SUBMITTED": ProperStatus = "Submitted"
        Case Else:        ProperStatus = Trim$(CStr(s))   ' pass through as-is
    End Select
End Function


' Ensure a nested region→status counter dict exists with zeroed keys
Private Sub InitRegStat(ByRef dict As Object, ByVal reg As String)
    If dict Is Nothing Then Set dict = CreateObject("Scripting.Dictionary")
    If Not dict.Exists(reg) Then
        Dim s As Object
        Set s = CreateObject("Scripting.Dictionary")
        s("Approved") = 0
        s("Submitted") = 0
        dict.Add reg, s
    End If
End Sub

' Safe Long from a 1-level dictionary; 0 if missing
Private Function NzLng2(ByVal dict As Object, ByVal key As Variant) As Long
    If (Not dict Is Nothing) And dict.Exists(key) Then
        If IsNumeric(dict(key)) Then NzLng2 = CLng(dict(key)) Else NzLng2 = 0
    Else
        NzLng2 = 0
    End If
End Function

' Safe Long from a 2-level dictionary: dict(reg)(status); 0 if missing
Private Function NzLngNested(ByVal dict As Object, ByVal regKey As Variant, ByVal statusKey As Variant) As Long
    If dict Is Nothing Then Exit Function
    If Not dict.Exists(regKey) Then Exit Function
    Dim inner As Object
    Set inner = dict(regKey)
    If inner Is Nothing Then Exit Function
    If inner.Exists(statusKey) Then
        If IsNumeric(inner(statusKey)) Then NzLngNested = CLng(inner(statusKey))
    End If
End Function

' Union of top-level keys across up to 5 dictionaries; returns a dictionary of unique keys
' Union of top-level keys across up to 5 dictionaries; returns a dictionary of unique keys
Private Function UnionKeys(ByVal d1 As Object, _
                           Optional ByVal d2 As Object, _
                           Optional ByVal d3 As Object, _
                           Optional ByVal d4 As Object, _
                           Optional ByVal d5 As Object) As Object
    Dim u As Object: Set u = CreateObject("Scripting.Dictionary")
    Dim k As Variant

    If Not d1 Is Nothing Then
        For Each k In d1.Keys
            If Not u.Exists(k) Then u.Add k, True
        Next k
    End If

    If Not d2 Is Nothing Then
        For Each k In d2.Keys
            If Not u.Exists(k) Then u.Add k, True
        Next k
    End If

    If Not d3 Is Nothing Then
        For Each k In d3.Keys
            If Not u.Exists(k) Then u.Add k, True
        Next k
    End If

    If Not d4 Is Nothing Then
        For Each k In d4.Keys
            If Not u.Exists(k) Then u.Add k, True
        Next k
    End If

    If Not d5 Is Nothing Then
        For Each k In d5.Keys
            If Not u.Exists(k) Then u.Add k, True
        Next k
    End If

    Set UnionKeys = u
End Function

' --- CLEANUP TEMP SHEETS IN MAIN FILE ---
Dim tmpWs As Worksheet
For Each tmpWs In wbMain.Worksheets
    Select Case tmpWs.Name
        Case Format(Date, "ddmmyy"), "CoR Recali", "CoR Mismatch Summary"
            Application.DisplayAlerts = False
            tmpWs.Delete
            Application.DisplayAlerts = True
    End Select
Next tmpWs
