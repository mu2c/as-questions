Attribute VB_Name = "ConsolidateTables"
Option Explicit

' =============================================================================
' ConsolidateTables.bas
'
' Scans a folder for .pptx files, extracts every table whose first row is
' exactly "#" | "Question" | "Answer" (3 columns, case-sensitive), and writes
' all rows into a single consolidated.xlsx in that folder. One sheet per source
' file, plus a "_log" sheet with diagnostics.
'
' Entry point: ConsolidateTables (Alt+F8 -> Run).
' Late binding for PowerPoint - no library reference needed.
' =============================================================================

' --- Configuration constants (edit these to change behaviour) ---------------

Private Const DRY_RUN As Boolean = False        ' True = log only, no data sheets written
Private Const VERBOSE As Boolean = False        ' True = log every slide examined

Private Const HEADER_COL1 As String = "#"
Private Const HEADER_COL2 As String = "Question"
Private Const HEADER_COL3 As String = "Answer"

Private Const OUTPUT_FILENAME As String = "consolidated.xlsx"
Private Const LOG_SHEET_NAME As String = "_log"

' --- Office enum values (declared because of late binding) ------------------

Private Const MSO_GROUP As Long = 6             ' msoGroup
Private Const MSO_TABLE As Long = 19            ' msoTable
Private Const MSO_PLACEHOLDER As Long = 14      ' msoPlaceholder

Private Const XL_OPEN_XML_WORKBOOK As Long = 51 ' xlOpenXMLWorkbook (.xlsx)
Private Const XL_CALC_MANUAL As Long = -4135    ' xlCalculationManual
Private Const XL_CALC_AUTO As Long = -4105      ' xlCalculationAutomatic

' --- Run-stats container (passed by ref through the run) --------------------

Private Type RunStats
    FilesProcessed As Long
    SheetsCreated As Long
    RowsExtracted As Long
    Warnings As Long
    Errors As Long
End Type

' ===========================================================================
' Sub ConsolidateTables -- entry point and orchestration
' ===========================================================================
Public Sub ConsolidateTables()
    Dim folderPath As String
    Dim files As Collection
    Dim outputPath As String
    Dim wbOut As Workbook
    Dim logSheet As Worksheet
    Dim pptApp As Object
    Dim startedPpt As Boolean
    Dim stats As RunStats
    Dim i As Long
    Dim savedScreen As Boolean, savedEvents As Boolean
    Dim savedCalc As Long
    Dim alertsOff As Boolean

    ' --- Save Excel app state so Cleanup can restore it -------------------
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedCalc = Application.Calculation
    alertsOff = False

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = XL_CALC_MANUAL

    On Error GoTo Cleanup

    ' --- 1. Pick folder ---------------------------------------------------
    folderPath = PickFolder()
    If Len(folderPath) = 0 Then GoTo Cleanup    ' user cancelled

    ' --- 2. Enumerate .pptx files (excluding ~$* and hidden) --------------
    Set files = EnumeratePptFiles(folderPath)
    If files.Count = 0 Then
        MsgBox "No .pptx files found in:" & vbCrLf & folderPath, vbInformation
        GoTo Cleanup
    End If

    ' --- 3. Output collision check ---------------------------------------
    outputPath = JoinPath(folderPath, OUTPUT_FILENAME)
    If Not HandleOutputCollision(outputPath) Then GoTo Cleanup

    ' --- 4. Create the output workbook in memory --------------------------
    Application.DisplayAlerts = False
    alertsOff = True
    Set wbOut = Workbooks.Add
    ' Drop all but one starter sheet; we'll rename it as the log.
    Do While wbOut.Worksheets.Count > 1
        wbOut.Worksheets(wbOut.Worksheets.Count).Delete
    Loop
    Set logSheet = wbOut.Worksheets(1)
    logSheet.Name = LOG_SHEET_NAME
    InitLogSheet logSheet

    ' --- 5. Acquire PowerPoint -------------------------------------------
    Set pptApp = GetOrStartPowerPoint(startedPpt)
    If pptApp Is Nothing Then
        MsgBox "Could not start PowerPoint. Aborting.", vbCritical
        GoTo Cleanup
    End If

    ' --- 6. Process each file --------------------------------------------
    For i = 1 To files.Count
        Application.StatusBar = "Processing " & i & " of " & files.Count & ": " & _
                                FileNameOnly(CStr(files(i)))
        ProcessFile pptApp, CStr(files(i)), wbOut, logSheet, stats
        stats.FilesProcessed = stats.FilesProcessed + 1
    Next i

    ' --- 7. Save output ---------------------------------------------------
    If Not DRY_RUN Then
        On Error Resume Next
        wbOut.SaveAs Filename:=outputPath, FileFormat:=XL_OPEN_XML_WORKBOOK
        If Err.Number <> 0 Then
            MsgBox "Could not save " & outputPath & vbCrLf & _
                   "Error: " & Err.Description & vbCrLf & vbCrLf & _
                   "The workbook is still open in Excel; use File > Save As to keep your data.", _
                   vbExclamation
            Err.Clear
        End If
        On Error GoTo Cleanup
    End If

    ' --- 8. Activate _log if any warnings/errors --------------------------
    If stats.Warnings > 0 Or stats.Errors > 0 Then
        On Error Resume Next
        logSheet.Activate
        On Error GoTo Cleanup
    End If

    ' --- 9. Final summary -------------------------------------------------
    Application.StatusBar = False
    MsgBox SummaryText(stats, IIf(DRY_RUN, "(DRY RUN - no data sheets written)", outputPath)), _
           vbInformation, "ConsolidateTables - Done"

Cleanup:
    Dim cleanupErr As Long, cleanupMsg As String
    cleanupErr = Err.Number
    cleanupMsg = Err.Description

    On Error Resume Next
    If Not pptApp Is Nothing Then
        If startedPpt Then pptApp.Quit
    End If
    Set pptApp = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = savedScreen
    Application.EnableEvents = savedEvents
    Application.Calculation = savedCalc
    If alertsOff Then Application.DisplayAlerts = True

    If cleanupErr <> 0 Then
        MsgBox "Unexpected error: " & cleanupErr & " - " & cleanupMsg, vbCritical
    End If
End Sub

' ===========================================================================
' Function PickFolder -- returns chosen folder path, or "" if cancelled
' ===========================================================================
Private Function PickFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)          ' msoFileDialogFolderPicker = 4
    fd.Title = "Choose folder containing .pptx files"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = ""
    End If
End Function

' ===========================================================================
' Function EnumeratePptFiles -- returns Collection of full paths
' Excludes Office lock files (~$*) and hidden files.
' ===========================================================================
Private Function EnumeratePptFiles(ByVal folderPath As String) As Collection
    Dim col As New Collection
    Dim fso As Object, folder As Object, f As Object
    Dim nm As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    For Each f In folder.files
        nm = f.Name
        If LCase$(fso.GetExtensionName(nm)) = "pptx" Then
            If Left$(nm, 2) <> "~$" Then
                ' Skip hidden files (Hidden bit = 2)
                If (f.Attributes And 2) = 0 Then
                    col.Add f.Path
                End If
            End If
        End If
    Next f
    Set EnumeratePptFiles = col
End Function

' ===========================================================================
' Function HandleOutputCollision -- prompts user; returns True if OK to proceed
' ===========================================================================
Private Function HandleOutputCollision(ByVal outputPath As String) As Boolean
    Dim resp As VbMsgBoxResult
    If Len(Dir$(outputPath)) = 0 Then
        HandleOutputCollision = True
        Exit Function
    End If

    resp = MsgBox(OUTPUT_FILENAME & " already exists in this folder." & vbCrLf & _
                  "Overwrite it?", vbYesNo + vbQuestion, "File exists")
    If resp <> vbYes Then
        HandleOutputCollision = False
        Exit Function
    End If

    On Error Resume Next
    Kill outputPath
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "Cannot delete " & outputPath & vbCrLf & _
               "It may be open in another Excel window. Close it and try again.", _
               vbExclamation
        HandleOutputCollision = False
        Exit Function
    End If
    On Error GoTo 0
    HandleOutputCollision = True
End Function

' ===========================================================================
' Function GetOrStartPowerPoint -- returns app object; sets startedByUs
' ===========================================================================
Private Function GetOrStartPowerPoint(ByRef startedByUs As Boolean) As Object
    Dim app As Object
    startedByUs = False
    On Error Resume Next
    Set app = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then
        On Error Resume Next
        Set app = CreateObject("PowerPoint.Application")
        On Error GoTo 0
        If Not app Is Nothing Then startedByUs = True
    End If
    Set GetOrStartPowerPoint = app
End Function

' ===========================================================================
' Sub ProcessFile -- opens one .pptx, extracts matching tables, writes sheet
' ===========================================================================
Private Sub ProcessFile(ByVal pptApp As Object, ByVal filePath As String, _
                        ByVal wbOut As Workbook, ByVal logSheet As Worksheet, _
                        ByRef stats As RunStats)
    Dim pres As Object
    Dim slide As Object
    Dim tables As Collection
    Dim tbl As Object
    Dim slideIdx As Long
    Dim tableIdx As Long
    Dim allRows As Collection      ' of 2-D one-row Variants
    Dim rows2D As Variant
    Dim sheetName As String
    Dim fileBase As String
    Dim presOpened As Boolean

    fileBase = FileNameOnly(filePath)
    Set allRows = New Collection

    On Error GoTo FileError

    Set pres = pptApp.Presentations.Open(FileName:=filePath, _
                                         ReadOnly:=True, _
                                         Untitled:=False, _
                                         WithWindow:=False)
    presOpened = True

    For slideIdx = 1 To pres.Slides.Count
        Set slide = pres.Slides(slideIdx)
        Set tables = New Collection
        CollectTablesFromShapes slide.Shapes, tables

        Select Case tables.Count
            Case 0
                If VERBOSE Then LogWrite logSheet, "INFO", fileBase, slideIdx, 0, 0, "No tables on slide"
                stats = stats   ' no-op to keep ByRef happy
            Case 1
                Set tbl = tables(1)
                If IsMatchingHeader(tbl) Then
                    Dim rowsThisTable As Long
                    rowsThisTable = AppendExtractedRows(tbl, allRows)
                    LogWrite logSheet, "INFO", fileBase, slideIdx, 1, rowsThisTable, _
                             "Extracted matching table"
                Else
                    LogWrite logSheet, "WARN", fileBase, slideIdx, 1, 0, _
                             "Table header does not match #/Question/Answer (3 cols, case-sensitive)"
                    stats.Warnings = stats.Warnings + 1
                End If
            Case Else
                LogWrite logSheet, "WARN", fileBase, slideIdx, tables.Count, 0, _
                         "Slide has " & tables.Count & " tables; skipped per multi-table rule"
                stats.Warnings = stats.Warnings + 1
        End Select
    Next slideIdx

    If allRows.Count = 0 Then
        LogWrite logSheet, "WARN", fileBase, 0, 0, 0, _
                 "No matching tables found in file; no sheet created"
        stats.Warnings = stats.Warnings + 1
    Else
        rows2D = CollectionTo2D(allRows)
        If Not DRY_RUN Then
            sheetName = SafeSheetName(StripExtension(fileBase), wbOut)
            WriteSheet wbOut, sheetName, rows2D, logSheet, fileBase
            stats.SheetsCreated = stats.SheetsCreated + 1
        End If
        stats.RowsExtracted = stats.RowsExtracted + UBound(rows2D, 1)
    End If

    GoTo FileCleanup

FileError:
    LogWrite logSheet, "ERROR", fileBase, 0, 0, 0, _
             "Failed to process file: " & Err.Description
    stats.Errors = stats.Errors + 1
    Err.Clear

FileCleanup:
    If presOpened Then
        On Error Resume Next
        pres.Close
        On Error GoTo 0
    End If
    Set pres = Nothing
End Sub

' ===========================================================================
' Sub CollectTablesFromShapes -- recurses into groups, fills tables collection
' ===========================================================================
Private Sub CollectTablesFromShapes(ByVal shapes As Object, ByRef tables As Collection)
    Dim shp As Object
    Dim sub_ As Object
    For Each shp In shapes
        On Error Resume Next
        Dim isGroup As Boolean, hasTable As Boolean
        isGroup = (shp.Type = MSO_GROUP)
        hasTable = shp.HasTable
        On Error GoTo 0

        If isGroup Then
            ' Recurse into the group's child shapes.
            On Error Resume Next
            CollectTablesFromShapes shp.GroupItems, tables
            On Error GoTo 0
        ElseIf hasTable Then
            tables.Add shp.Table
        End If
    Next shp
End Sub

' ===========================================================================
' Function IsMatchingHeader -- exact, case-sensitive, 3 cols
' Strips trailing vbCr (PowerPoint always appends one paragraph terminator).
' ===========================================================================
Private Function IsMatchingHeader(ByVal tbl As Object) As Boolean
    IsMatchingHeader = False
    On Error GoTo HdrErr
    If tbl.Columns.Count <> 3 Then Exit Function
    If tbl.Rows.Count < 1 Then Exit Function

    Dim c1 As String, c2 As String, c3 As String
    c1 = NormalizeHeaderCell(tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text)
    c2 = NormalizeHeaderCell(tbl.Cell(1, 2).Shape.TextFrame.TextRange.Text)
    c3 = NormalizeHeaderCell(tbl.Cell(1, 3).Shape.TextFrame.TextRange.Text)

    ' StrComp with vbBinaryCompare = case-sensitive comparison.
    If StrComp(c1, HEADER_COL1, vbBinaryCompare) = 0 _
       And StrComp(c2, HEADER_COL2, vbBinaryCompare) = 0 _
       And StrComp(c3, HEADER_COL3, vbBinaryCompare) = 0 Then
        IsMatchingHeader = True
    End If
    Exit Function
HdrErr:
    IsMatchingHeader = False
End Function

' Header normalizer: Trim only. Strip trailing vbCr/vbLf because PowerPoint
' always appends a paragraph terminator to cell text.
Private Function NormalizeHeaderCell(ByVal s As String) As String
    Dim t As String
    t = s
    Do While Len(t) > 0
        Dim ch As String
        ch = Right$(t, 1)
        If ch = vbCr Or ch = vbLf Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    NormalizeHeaderCell = Trim$(t)
End Function

' ===========================================================================
' Function AppendExtractedRows -- pulls data rows into allRows; returns count
' Each row is added as a 1x3 Variant array.
' ===========================================================================
Private Function AppendExtractedRows(ByVal tbl As Object, _
                                     ByRef allRows As Collection) As Long
    Dim r As Long, added As Long
    Dim a As String, b As String, c As String
    Dim row(1 To 3) As Variant

    For r = 2 To tbl.Rows.Count
        a = NormalizeBodyCell(tbl.Cell(r, 1).Shape.TextFrame.TextRange.Text)
        b = NormalizeBodyCell(tbl.Cell(r, 2).Shape.TextFrame.TextRange.Text)
        c = NormalizeBodyCell(tbl.Cell(r, 3).Shape.TextFrame.TextRange.Text)
        If Len(a) = 0 And Len(b) = 0 And Len(c) = 0 Then
            ' Skip fully empty row
        Else
            row(1) = a
            row(2) = b
            row(3) = c
            allRows.Add Array(row(1), row(2), row(3))
            added = added + 1
        End If
    Next r
    AppendExtractedRows = added
End Function

' Body normalizer: convert vbCr to vbLf so Excel renders multi-line answers
' as proper line breaks (with wrap text). Trim outer whitespace and strip a
' single trailing vbLf left over from PowerPoint's paragraph terminator.
Private Function NormalizeBodyCell(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCr, vbLf)
    Do While Len(t) > 0
        If Right$(t, 1) = vbLf Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    NormalizeBodyCell = Trim$(t)
End Function

' ===========================================================================
' Function CollectionTo2D -- converts a Collection of 1x3 arrays into one
' 2-D Variant sized (1..N, 1..3) suitable for one-shot Range assignment.
' ===========================================================================
Private Function CollectionTo2D(ByVal rows As Collection) As Variant
    Dim n As Long, i As Long
    Dim out() As Variant
    Dim oneRow As Variant
    n = rows.Count
    ReDim out(1 To n, 1 To 3)
    For i = 1 To n
        oneRow = rows(i)
        out(i, 1) = oneRow(0)
        out(i, 2) = oneRow(1)
        out(i, 3) = oneRow(2)
    Next i
    CollectionTo2D = out
End Function

' ===========================================================================
' Sub WriteSheet -- creates sheet, writes header + 2-D body in one assignment
' ===========================================================================
Private Sub WriteSheet(ByVal wbOut As Workbook, ByVal sheetName As String, _
                       ByVal rows2D As Variant, ByVal logSheet As Worksheet, _
                       ByVal sourceFile As String)
    Dim ws As Worksheet
    Dim n As Long
    n = UBound(rows2D, 1)

    ' Insert at end (just before _log if _log is last; either order is fine).
    Set ws = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
    ws.Name = sheetName

    ' Format target columns as Text BEFORE writing so values like "=foo",
    ' "+1", "01/02/2024", or long numeric strings are preserved verbatim.
    ws.Columns("A:C").NumberFormat = "@"

    ws.Range("A1").Value = HEADER_COL1
    ws.Range("B1").Value = HEADER_COL2
    ws.Range("C1").Value = HEADER_COL3
    ws.Range("A1:C1").Font.Bold = True

    If n > 0 Then
        ws.Range("A2").Resize(n, 3).Value = rows2D
        ws.Range("A2").Resize(n, 3).WrapText = True
    End If

    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B:C").ColumnWidth = 60
    ws.Rows("1:1").Font.Bold = True
    ws.Activate
    ws.Range("A1").Select
End Sub

' ===========================================================================
' Function SafeSheetName -- 31-char max, sanitize, dedupe, fallback
' ===========================================================================
Private Function SafeSheetName(ByVal stem As String, ByVal wbOut As Workbook) As String
    Dim s As String
    Dim base As String
    Dim suffix As Long
    Dim candidate As String

    s = stem
    ' Strip invalid characters: : \ / ? * [ ]
    Dim badChars As Variant, i As Long
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "_")
    Next i
    ' Strip leading/trailing apostrophes (Excel rejects them)
    Do While Len(s) > 0 And Left$(s, 1) = "'"
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0 And Right$(s, 1) = "'"
        s = Left$(s, Len(s) - 1)
    Loop
    s = Trim$(s)

    If Len(s) = 0 Then s = "Sheet"
    If StrComp(s, "History", vbTextCompare) = 0 Then s = "History_"
    If StrComp(s, LOG_SHEET_NAME, vbTextCompare) = 0 Then s = s & "_src"

    ' Truncate to 28 chars to leave room for "_NN" dedupe suffix.
    If Len(s) > 28 Then s = Left$(s, 28)
    base = s

    candidate = base
    suffix = 1
    Do While SheetExists(wbOut, candidate)
        suffix = suffix + 1
        candidate = base & "_" & suffix
        If Len(candidate) > 31 Then
            ' Trim base further to keep within 31.
            candidate = Left$(base, 31 - Len("_" & suffix)) & "_" & suffix
        End If
    Loop
    SafeSheetName = candidate
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ===========================================================================
' Log-sheet helpers
' ===========================================================================
Private Sub InitLogSheet(ByVal logSheet As Worksheet)
    logSheet.Range("A1:G1").Value = Array("Timestamp", "Level", "File", _
                                          "Slide", "TableIdx", "RowsExtracted", "Message")
    logSheet.Range("A1:G1").Font.Bold = True
    logSheet.Columns("A").ColumnWidth = 20
    logSheet.Columns("B").ColumnWidth = 8
    logSheet.Columns("C").ColumnWidth = 35
    logSheet.Columns("D:F").ColumnWidth = 8
    logSheet.Columns("G").ColumnWidth = 70
    logSheet.Rows("1:1").Font.Bold = True
End Sub

Private Sub LogWrite(ByVal logSheet As Worksheet, ByVal level As String, _
                     ByVal file As String, ByVal slide As Long, _
                     ByVal tableIdx As Long, ByVal rowsExtracted As Long, _
                     ByVal msg As String)
    Dim r As Long
    r = logSheet.Cells(logSheet.Rows.Count, "A").End(-4162).Row + 1   ' xlUp = -4162
    If r < 2 Then r = 2
    logSheet.Cells(r, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    logSheet.Cells(r, 2).Value = level
    logSheet.Cells(r, 3).Value = file
    logSheet.Cells(r, 4).Value = IIf(slide > 0, slide, "")
    logSheet.Cells(r, 5).Value = IIf(tableIdx > 0, tableIdx, "")
    logSheet.Cells(r, 6).Value = IIf(rowsExtracted > 0, rowsExtracted, "")
    logSheet.Cells(r, 7).Value = msg
End Sub

' ===========================================================================
' Misc helpers
' ===========================================================================
Private Function JoinPath(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Or Right$(a, 1) = "/" Then
        JoinPath = a & b
    Else
        JoinPath = a & "\" & b
    End If
End Function

Private Function FileNameOnly(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then
        FileNameOnly = Mid$(p, i + 1)
    Else
        FileNameOnly = p
    End If
End Function

Private Function StripExtension(ByVal nm As String) As String
    Dim i As Long
    i = InStrRev(nm, ".")
    If i > 1 Then
        StripExtension = Left$(nm, i - 1)
    Else
        StripExtension = nm
    End If
End Function

Private Function SummaryText(ByRef stats As RunStats, ByVal where As String) As String
    SummaryText = "Done." & vbCrLf & vbCrLf & _
                  "Files processed:  " & stats.FilesProcessed & vbCrLf & _
                  "Sheets created:   " & stats.SheetsCreated & vbCrLf & _
                  "Rows extracted:   " & stats.RowsExtracted & vbCrLf & _
                  "Warnings:         " & stats.Warnings & vbCrLf & _
                  "Errors:           " & stats.Errors & vbCrLf & vbCrLf & _
                  "Output: " & where & vbCrLf & vbCrLf & _
                  "See the _log sheet for per-slide details."
End Function
