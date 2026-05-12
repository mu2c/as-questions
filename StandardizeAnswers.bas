Attribute VB_Name = "StandardizeAnswers"
Option Explicit

' =============================================================================
' StandardizeAnswers.bas
'
' Post-processing macro for the workbook produced by ConsolidateTables.bas.
' Walks every sheet whose C1 cell is "Answer" and rewrites column C so that
' multi-item answers appear one item per line.
'
' Run: open consolidated.xlsx (or any similarly-shaped workbook), then
' Alt+F8 -> StandardizeAnswers -> Run.
'
' Detected marker styles (first set yielding 2+ sequential markers wins):
'   a) b) c) ...   A) B) C) ...   a. b. c. ...   A. B. C. ...
'   1) 2) 3) ...   1. 2. 3. ...
'
' A marker is only accepted if the character immediately after it is a
' space or end-of-string (so "3.7" / "B.O." don't get treated as markers).
' Markers may be preceded by any character, so "GLB." correctly resolves to
' a "B." marker stuck onto the previous word.
'
' Cells already containing line breaks ARE re-processed (whitespace is
' flattened first). Cells where no marker pattern is detected are left
' unchanged.
' =============================================================================

' --- Config -----------------------------------------------------------------
Private Const ANSWER_HEADER As String = "Answer"   ' must match C1 to process
Private Const LOG_SHEET_NAME As String = "_log"    ' skipped by name

' --- Excel constants (avoid late-binding surprises) -------------------------
Private Const XL_CALC_MANUAL As Long = -4135
Private Const XL_UP As Long = -4162

' ===========================================================================
' Sub StandardizeAnswers -- entry point
' ===========================================================================
Public Sub StandardizeAnswers()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook. Open consolidated.xlsx first.", vbExclamation
        Exit Sub
    End If

    Dim savedScreen As Boolean, savedEvents As Boolean
    Dim savedCalc As Long
    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = XL_CALC_MANUAL

    On Error GoTo Cleanup

    Dim ws As Worksheet
    Dim sheetsScanned As Long, cellsProcessed As Long, cellsModified As Long
    Dim r As Long, lastRow As Long
    Dim original As String, standardized As String

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, LOG_SHEET_NAME, vbTextCompare) = 0 Then GoTo NextSheet

        ' Only process sheets whose C1 header is exactly "Answer".
        If StrComp(CStr(ws.Range("C1").Value), ANSWER_HEADER, vbBinaryCompare) <> 0 Then
            GoTo NextSheet
        End If

        sheetsScanned = sheetsScanned + 1
        lastRow = ws.Cells(ws.Rows.Count, "C").End(XL_UP).Row
        If lastRow < 2 Then GoTo NextSheet

        For r = 2 To lastRow
            original = CStr(ws.Cells(r, "C").Value)
            If Len(Trim$(original)) > 0 Then
                cellsProcessed = cellsProcessed + 1
                standardized = StandardizeAnswerText(original)
                If StrComp(standardized, original, vbBinaryCompare) <> 0 Then
                    ws.Cells(r, "C").Value = standardized
                    ws.Cells(r, "C").WrapText = True
                    cellsModified = cellsModified + 1
                End If
            End If
        Next r

NextSheet:
    Next ws

    MsgBox "Done." & vbCrLf & vbCrLf & _
           "Sheets scanned:   " & sheetsScanned & vbCrLf & _
           "Cells processed:  " & cellsProcessed & vbCrLf & _
           "Cells modified:   " & cellsModified, _
           vbInformation, "StandardizeAnswers"

Cleanup:
    Dim eN As Long, eD As String
    eN = Err.Number
    eD = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = savedScreen
    Application.EnableEvents = savedEvents
    Application.Calculation = savedCalc
    If eN <> 0 Then
        MsgBox "Error: " & eN & " - " & eD, vbCritical
    End If
End Sub

' ===========================================================================
' Function StandardizeAnswerText -- core splitter
' Returns one item per line, or the input unchanged if no markers detected.
' ===========================================================================
Private Function StandardizeAnswerText(ByVal s As String) As String
    ' Flatten all whitespace (including existing line breaks) to single
    ' spaces so the splitter sees one continuous string. This is what lets
    ' the macro re-process previously line-broken cells.
    Dim norm As String
    norm = Replace(s, vbCrLf, " ")
    norm = Replace(norm, vbCr, " ")
    norm = Replace(norm, vbLf, " ")
    norm = Replace(norm, vbTab, " ")
    Do While InStr(norm, "  ") > 0
        norm = Replace(norm, "  ", " ")
    Loop
    norm = Trim$(norm)
    If Len(norm) = 0 Then
        StandardizeAnswerText = s
        Exit Function
    End If

    ' Try marker sets in priority order. First set yielding >=2 sequential
    ' markers wins.
    Dim markerSets As Variant
    markerSets = Array( _
        Array("a)", "b)", "c)", "d)", "e)", "f)", "g)", "h)"), _
        Array("A)", "B)", "C)", "D)", "E)", "F)", "G)", "H)"), _
        Array("a.", "b.", "c.", "d.", "e.", "f.", "g.", "h."), _
        Array("A.", "B.", "C.", "D.", "E.", "F.", "G.", "H."), _
        Array("1)", "2)", "3)", "4)", "5)", "6)", "7)", "8)"), _
        Array("1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.") _
    )

    Dim setIdx As Long
    For setIdx = LBound(markerSets) To UBound(markerSets)
        Dim markers As Variant
        markers = markerSets(setIdx)

        Dim positions As Collection
        Set positions = FindSequentialMarkers(norm, markers)

        If positions.Count >= 2 Then
            StandardizeAnswerText = SplitAtMarkers(norm, positions, markers)
            Exit Function
        End If
    Next setIdx

    ' No marker pattern detected -- return input unchanged.
    StandardizeAnswerText = s
End Function

' ===========================================================================
' Function FindSequentialMarkers -- 1-based positions of markers appearing
' in order. Each accepted marker must be followed by a space or end-of-
' string so things like "3.7" / "B.O." don't get mistaken for markers.
' ===========================================================================
Private Function FindSequentialMarkers(ByVal s As String, ByVal markers As Variant) As Collection
    Dim positions As New Collection
    Dim searchFrom As Long
    searchFrom = 1

    Dim i As Long
    For i = LBound(markers) To UBound(markers)
        Dim marker As String
        marker = CStr(markers(i))
        Dim mLen As Long
        mLen = Len(marker)

        Dim pos As Long
        Dim accepted As Boolean
        accepted = False
        pos = searchFrom

        Do
            pos = InStr(pos, s, marker, vbBinaryCompare)
            If pos = 0 Then Exit Do

            Dim afterPos As Long
            afterPos = pos + mLen
            Dim ch As String
            If afterPos > Len(s) Then
                ch = ""
            Else
                ch = Mid$(s, afterPos, 1)
            End If

            If ch = "" Or ch = " " Or ch = vbTab Then
                accepted = True
                Exit Do
            End If

            pos = pos + 1                  ' keep scanning for a valid occurrence
        Loop

        If Not accepted Then Exit For
        positions.Add pos
        searchFrom = pos + mLen
    Next i

    Set FindSequentialMarkers = positions
End Function

' ===========================================================================
' Function SplitAtMarkers -- builds the line-broken output. Any text before
' the first marker is preserved as a leading line.
' ===========================================================================
Private Function SplitAtMarkers(ByVal s As String, ByVal positions As Collection, _
                                ByVal markers As Variant) As String
    Dim result As String
    Dim n As Long
    n = positions.Count

    Dim firstPos As Long
    firstPos = positions(1)
    If firstPos > 1 Then
        Dim preamble As String
        preamble = Trim$(Mid$(s, 1, firstPos - 1))
        If Len(preamble) > 0 Then result = preamble
    End If

    Dim i As Long
    For i = 1 To n
        Dim markerLen As Long
        markerLen = Len(CStr(markers(LBound(markers) + i - 1)))

        Dim startPos As Long, endPos As Long
        startPos = positions(i) + markerLen
        If i < n Then
            endPos = positions(i + 1) - 1
        Else
            endPos = Len(s)
        End If

        Dim piece As String
        If endPos >= startPos Then
            piece = Trim$(Mid$(s, startPos, endPos - startPos + 1))
        Else
            piece = ""
        End If

        If Len(piece) > 0 Then
            If Len(result) > 0 Then result = result & vbLf
            result = result & piece
        End If
    Next i

    SplitAtMarkers = result
End Function
