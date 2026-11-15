Option Explicit
Option Compare Text

' ================================================================
'  Constraint-Based Data Processor — VBA Toolkit
'  Five data-transformation modules for row-level processing:
'
'    Module 1 — Clean & Normalize
'         Standardize entity codes (strip suffixes), repair postal codes
'         (pad leading zeros), backfill empty codes from a lookup table.
'
'    Module 2 — Split & Redistribute
'         Split rows containing multiple slash-separated entity IDs
'         into one row each. Redistribute weight/volume proportionally
'         by piece count. Re-aggregate to absorb rounding drift.
'
'    Module 3 — Validate & Guard
'         Cross-check count consistency across multiple identifier
'         columns; flag mismatches so operator can intervene before split.
'
'    Module 4 — Constraint Convergence (the core algorithm)
'         Adjust dimensions uniformly across all rows until aggregate
'         volume lands within the target range [TARGET_MIN, TARGET_MAX].
'
'    Module 5 — Route & Classify
'         Pattern-match tracking numbers to carrier codes via regex
'         rules; auto-populate a routing column.
'
'  Author: ivy92888-glitch  |  License: MIT
'  Disclaimer: Unaffiliated with any spreadsheet vendor.
'              Generic clean-room implementation. Any resemblance to
'              specific business systems is coincidental.
' ================================================================

' --- SHEET CONSTANTS ---
Private Const SHEET_INPUT     As String = "Input"
Private Const SHEET_LOOKUP   As String = "Lookup"
Private Const SHEET_OUTPUT    As String = "Output"

' --- COLUMN LETTERS (adjust to your layout) ---
Private Const COL_CODE        As String = "E"   ' Entity code / warehouse ID (needs suffix strip)
Private Const COL_PART_NO     As String = "C"   ' Primary identifier column (slash-separated)
Private Const COL_PO          As String = "D"   ' Secondary identifier column (slash-separated)
Private Const COL_WEIGHT      As String = "F"   ' Weight per piece
Private Const COL_PIECES      As String = "G"   ' Piece count
Private Const COL_VOLUME      As String = "H"   ' Total volume per row
Private Const COL_LENGTH      As String = "I"   ' Box length
Private Const COL_WIDTH       As String = "J"   ' Box width
Private Const COL_HEIGHT      As String = "K"   ' Box height
Private Const COL_POSTAL      As String = "L"   ' Postal code (needs leading zero pad)
Private Const COL_COURIER     As String = "I"   ' Courier carrier column
Private Const COL_TRACKING    As String = "J"   ' Tracking number for routing

' --- CONSTRAINT TARGET ---
Private Const VOLUME_MIN      As Double = 68.0  ' lower bound of target aggregate volume
Private Const VOLUME_MAX      As Double = 68.1  ' upper bound
Private Const MIN_DIMENSION   As Double = 0.1   ' no dimension can go below this

' --- ROUTING RULES (regex patterns → carrier codes) ---
' Populated at runtime — add your own in InitGovernance
Private routingRules As Variant

' --- STANDARD ENTITY CODE SUFFIXES (strip from warehouse codes) ---
Private entitySuffixes As Variant

' --- POSTAL CODE LENGTH (pad with leading zeros to this) ---
Private Const POSTAL_LENGTH   As Integer = 5

' ================================================================
'  MAIN ENTRY
' ================================================================
Sub RunFullPipeline()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call InitPipeline

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_INPUT)

    Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long
    n1 = RunModule1_CleanNormalize(ThisWorkbook)
    n3 = RunModule3_Validate(ThisWorkbook)   ' validate BEFORE split
    n2 = RunModule2_SplitRedistribute(ThisWorkbook)
    n4 = RunModule4_ConstraintConverge(ThisWorkbook)
    n5 = RunModule5_RouteClassify(ThisWorkbook)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "=== Constraint-Based Data Processor ===" & vbCrLf & vbCrLf & _
           "Module 1 — Clean+Normalize:   " & n1 & " cells normalized" & vbCrLf & _
           "Module 2 — Split+Redistribute: " & n2 & " new rows created" & vbCrLf & _
           "Module 3 — Validate+Guard:     " & n3 & " validation issues" & vbCrLf & _
           "Module 4 — ConstraintConverge: volume adjusted into [" & VOLUME_MIN & ", " & VOLUME_MAX & "]" & vbCrLf & _
           "Module 5 — Route+Classify:     " & n5 & " rows routed", _
           vbInformation, "Pipeline Complete"
End Sub

Sub InitPipeline()
    ' Routing rules: (regex pattern → carrier code)
    ' Patterns checked in order; first match wins.
    routingRules = Array( _
        Array("^1Z[A-Z0-9]{16}$", "UPS"), _
        Array("^\d{12,14}$",      "DHL"), _
        Array("^[A-Z]{2}\d{9}$",   "DPD"), _
        Array("^[A-Z0-9]{8,}$",    "OTHER") _
    )
    ' Entity code suffixes to strip (keep only the code portion before the suffix)
    entitySuffixes = Array("-30001", "-30002", "-30000", "-backup", "-temporary")
End Sub

' ================================================================
'  MODULE 1 — Clean & Normalize
'  Strip suffixes from entity codes, pad postal codes to fixed length,
'  backfill missing codes from the Lookup sheet.
' ================================================================
Function RunModule1_CleanNormalize(ByVal wb As Workbook) As Long
    Dim ws As Worksheet, lookup As Worksheet
    Set ws = wb.Sheets(SHEET_INPUT)

    ' Build lookup dict for code backfill
    Dim lookupDict As Object
    Set lookupDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set lookup = wb.Sheets(SHEET_LOOKUP)
    On Error GoTo 0
    If Not lookup Is Nothing Then
        Dim lastRowL As Long: lastRowL = lookup.Cells(lookup.Rows.Count, 1).End(xlUp).Row
        Dim i As Long
        For i = 2 To lastRowL
            Dim key As String: key = Trim(lookup.Cells(i, 1).Value)
            Dim val As String: val = Trim(lookup.Cells(i, 2).Value)
            If Len(key) > 0 And Len(val) > 0 And Not lookupDict.Exists(key) Then
                lookupDict.Add key, val
            End If
        Next i
    End If

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_CODE).End(xlUp).Row
    Dim changed As Long: changed = 0

    For i = 2 To lastRow
        ' 1a — Strip suffixes from entity codes
        Dim code As String: code = Trim(ws.Cells(i, COL_CODE).Value)
        If Len(code) > 0 Then
            Dim suf As Variant
            For Each suf In entitySuffixes
                If InStr(code, suf) > 0 Then
                    code = Replace(code, suf, "")
                    ws.Cells(i, COL_CODE).Value = Trim(code)
                    changed = changed + 1
                    Exit For
                End If
            Next suf
        End If

        ' 1b — Backfill missing codes from lookup
        If Len(code) = 0 And lookupDict.Count > 0 Then
            Dim key2 As String: key2 = Trim(ws.Cells(i, "B").Value)   ' adjust: your ID column
            If lookupDict.Exists(key2) Then
                ws.Cells(i, COL_CODE).Value = lookupDict(key2)
                changed = changed + 1
            End If
        End If

        ' 1c — Pad postal codes to fixed length with leading zeros
        Dim postal As String: postal = Trim(ws.Cells(i, COL_POSTAL).Value)
        If Len(postal) > 0 And IsNumeric(postal) And Len(postal) < POSTAL_LENGTH Then
            postal = String(POSTAL_LENGTH - Len(postal), "0") & postal
            ws.Cells(i, COL_POSTAL).NumberFormat = "@"   ' force text
            ws.Cells(i, COL_POSTAL).Value = postal
            changed = changed + 1
        End If
    Next i

    RunModule1_CleanNormalize = changed
End Function

' ================================================================
'  MODULE 2 — Split & Redistribute
'  Split rows where COL_PART_NO contains "/" into one row each.
'  Redistribute weight/volume proportionally by piece count.
'  Re-aggregate after proportional split to absorb rounding drift.
' ================================================================
Function RunModule2_SplitRedistribute(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Set ws = wb.Sheets(SHEET_INPUT)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_PART_NO).End(xlUp).Row

    ' Collect rows that need splitting (PartNo contains "/")
    Dim splitList As Object
    Set splitList = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = lastRow To 2 Step -1   ' collect bottom-up so row numbers stay valid
        Dim partNo As String: partNo = Trim(ws.Cells(i, COL_PART_NO).Value)
        If InStr(partNo, "/") > 0 Then
            splitList.Add i, True
        End If
    Next i

    If splitList.Count = 0 Then
        RunModule2_SplitRedistribute = 0: Exit Function
    End If

    Dim rowsCreated As Long: rowsCreated = 0

    ' Process each split row (from bottom up so Insert shifts don't invalidate row nums)
    Dim rowNum As Variant
    For Each rowNum In splitList.Keys
        Dim srcRow As Long: srcRow = CLng(rowNum)

        ' Split the identifiers
        Dim parts() As String
        parts = Split(Trim(ws.Cells(srcRow, COL_PART_NO).Value), "/")
        Dim pos() As String
        pos = Split(Trim(ws.Cells(srcRow, COL_PO).Value), "/")

        Dim totalPieces As Double: totalPieces = CDbl(Val(Trim(ws.Cells(srcRow, COL_PIECES).Value)))
        If totalPieces = 0 Then totalPieces = 1
        Dim totalWeight As Double: totalWeight = CDbl(Val(Trim(ws.Cells(srcRow, COL_WEIGHT).Value))) * totalPieces
        Dim totalVol As Double:  totalVol = CDbl(Val(Trim(ws.Cells(srcRow, COL_VOLUME).Value)))

        ' Insert N-1 new rows below the source row
        Dim numNew As Long: numNew = UBound(parts)   ' total = UBound+1 parts; insert UBound rows
        ws.Rows(srcRow + 1 & ":" & srcRow + numNew).Insert Shift:=xlDown
        rowsCreated = rowsCreated + numNew

        ' Redistribute proportionally
        Dim p As Long
        For p = 0 To UBound(parts)
            Dim targetRow As Long: targetRow = srcRow + p
            Dim ratio As Double: ratio = 1 / (UBound(parts) + 1)   ' equal split; adjust if pieces differ per part

            ' Copy all columns from source row, then override split fields
            ws.Rows(srcRow).Copy Destination:=ws.Rows(targetRow)

            ws.Cells(targetRow, COL_PART_NO).Value = Trim(parts(p))
            If p <= UBound(pos) Then
                ws.Cells(targetRow, COL_PO).Value = Trim(pos(p))
            Else
                ws.Cells(targetRow, COL_PO).Value = Trim(pos(0))   ' fallback
            End If

            ' Redistribute weight and volume
            ws.Cells(targetRow, COL_WEIGHT).Value = Round((totalWeight * ratio) / totalPieces, 3)
            ws.Cells(targetRow, COL_VOLUME).Value = Round(totalVol * ratio, 4)
        Next p

        ' Re-aggregate and absorb rounding drift (on the newly split block)
        Call AbsorbRoundingDrift(ws, srcRow, srcRow + UBound(parts), totalWeight, totalVol, totalPieces)
    Next rowNum

    RunModule2_SplitRedistribute = rowsCreated
End Function

Private Sub AbsorbRoundingDrift(ByVal ws As Worksheet, ByVal firstRow As Long, _
                                 ByVal lastRow As Long, ByVal origWeight As Double, _
                                 ByVal origVol As Double, ByVal pieces As Double)
    Dim currentWeight As Double, currentVol As Double
    Dim r As Long
    For r = firstRow To lastRow
        currentWeight = currentWeight + CDbl(Val(Trim(ws.Cells(r, COL_WEIGHT).Value))) * pieces
        currentVol = currentVol + CDbl(Val(Trim(ws.Cells(r, COL_VOLUME).Value)))
    Next r

    Dim driftWeight As Double: driftWeight = origWeight - currentWeight
    Dim driftVol As Double:    driftVol = origVol - currentVol

    ' Spread drift across all rows (first row absorbs any remainder)
    Dim driftPerRowW As Double: driftPerRowW = driftWeight / (lastRow - firstRow + 1)
    Dim driftPerRowV As Double: driftPerRowV = driftVol / (lastRow - firstRow + 1)

    For r = firstRow To lastRow
        Dim adjW As Double: adjW = CDbl(Val(Trim(ws.Cells(r, COL_WEIGHT).Value))) + driftPerRowW / pieces
        Dim adjV As Double: adjV = CDbl(Val(Trim(ws.Cells(r, COL_VOLUME).Value))) + driftPerRowV
        ws.Cells(r, COL_WEIGHT).Value = Round(adjW, 3)
        ws.Cells(r, COL_VOLUME).Value = Round(adjV, 4)
    Next r
End Sub

' ================================================================
'  MODULE 3 — Validate & Guard
'  Before splitting: cross-check identifier column count consistency.
'  If COL_PART_NO and COL_PO have different number of "/" separators,
'  flag the row so operator intervenes.
' ================================================================
Function RunModule3_Validate(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Set ws = wb.Sheets(SHEET_INPUT)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_PART_NO).End(xlUp).Row
    Dim issues As Long: issues = 0
    Dim i As Long

    For i = 2 To lastRow
        Dim partNo As String: partNo = Trim(ws.Cells(i, COL_PART_NO).Value)
        Dim po As String:     po     = Trim(ws.Cells(i, COL_PO).Value)

        If InStr(partNo, "/") > 0 Or InStr(po, "/") > 0 Then
            Dim partCount As Long: partCount = UBound(Split(partNo, "/")) + 1
            Dim poCount As Long:   poCount   = UBound(Split(po, "/")) + 1

            If partCount <> poCount Then
                ' Flag: red font on both columns, add warning comment in column Z
                ws.Cells(i, COL_PART_NO).Font.Color = vbRed
                ws.Cells(i, COL_PO).Font.Color = vbRed
                issues = issues + 1
            End If
        End If
    Next i

    RunModule3_Validate = issues
End Function

' ================================================================
'  MODULE 4 — Constraint Convergence (the core algorithm)
'  Compute current aggregate volume → determine direction →
'  uniformly adjust length/width/height of every row until aggregate
'  volume lands within [VOLUME_MIN, VOLUME_MAX].
'
'  Volume = Σ(L × W × H) over all rows
'  Adjustment: each dimension in every row ± Δ (same delta per row)
'  Iterate until aggregate lands in range; respect MIN_DIMENSION floor.
' ================================================================
Function RunModule4_ConstraintConverge(ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    Set ws = wb.Sheets(SHEET_INPUT)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_LENGTH).End(xlUp).Row

    Dim currentVol As Double: currentVol = ComputeTotalVolume(ws, lastRow)

    If currentVol >= VOLUME_MIN And currentVol <= VOLUME_MAX Then
        RunModule4_ConstraintConverge = True: Exit Function   ' already in range
    End If

    ' Determine direction
    Dim delta As Double, direction As Integer
    If currentVol > VOLUME_MAX Then
        direction = -1   ' shrink
        delta = -0.01    ' shrink by 1 cm per iteration
    Else
        direction = 1    ' grow
        delta = 0.01
    End If

    ' Iterate (max 200 iterations to guard against non-convergence)
    Dim iter As Long
    For iter = 1 To 200
        ' Adjust dimensions uniformly across all rows
        Dim i As Long
        For i = 2 To lastRow
            Dim L As Double, W As Double, H As Double
            L = CDbl(Val(Trim(ws.Cells(i, COL_LENGTH).Value)))
            W = CDbl(Val(Trim(ws.Cells(i, COL_WIDTH).Value)))
            H = CDbl(Val(Trim(ws.Cells(i, COL_HEIGHT).Value)))

            L = Round(L + delta, 2)
            W = Round(W + delta, 2)
            H = Round(H + delta, 2)

            ' Respect minimum dimension constraint
            If L < MIN_DIMENSION Then L = MIN_DIMENSION
            If W < MIN_DIMENSION Then W = MIN_DIMENSION
            If H < MIN_DIMENSION Then H = MIN_DIMENSION

            ws.Cells(i, COL_LENGTH).Value = L
            ws.Cells(i, COL_WIDTH).Value  = W
            ws.Cells(i, COL_HEIGHT).Value = H
        Next i

        currentVol = ComputeTotalVolume(ws, lastRow)

        ' Update volume column for each row (L×W×H)
        For i = 2 To lastRow
            L = CDbl(Val(Trim(ws.Cells(i, COL_LENGTH).Value)))
            W = CDbl(Val(Trim(ws.Cells(i, COL_WIDTH).Value)))
            H = CDbl(Val(Trim(ws.Cells(i, COL_HEIGHT).Value)))
            ws.Cells(i, COL_VOLUME).Value = Round(L * W * H, 4)
        Next i

        ' Check convergence
        If direction = -1 And currentVol <= VOLUME_MAX Then
            RunModule4_ConstraintConverge = True: Exit Function
        End If
        If direction = 1 And currentVol >= VOLUME_MIN Then
            RunModule4_ConstraintConverge = True: Exit Function
        End If
    Next iter

    ' Did not converge within 200 iterations
    RunModule4_ConstraintConverge = False
End Function

Function ComputeTotalVolume(ByVal ws As Worksheet, ByVal lastRow As Long) As Double
    Dim total As Double: total = 0
    Dim i As Long
    For i = 2 To lastRow
        Dim v As Double
        v = CDbl(Val(Trim(ws.Cells(i, COL_VOLUME).Value)))
        total = total + v
    Next i
    ComputeTotalVolume = total
End Function

' ================================================================
'  MODULE 5 — Route & Classify
'  Match tracking numbers against regex rules to auto-populate
'  carrier / routing column. First matching rule wins.
' ================================================================
Function RunModule5_RouteClassify(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Set ws = wb.Sheets(SHEET_INPUT)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_TRACKING).End(xlUp).Row
    Dim routed As Long: routed = 0
    Dim i As Long

    For i = 2 To lastRow
        Dim tracking As String: tracking = Trim(ws.Cells(i, COL_TRACKING).Value)
        If Len(tracking) = 0 Then GoTo R5_Next

        Dim carrier As String: carrier = ""
        Dim rule As Variant
        For Each rule In routingRules
            Dim pat As String: pat = CStr(rule(0))
            Dim code As String: code = CStr(rule(1))
            If RegexMatch(tracking, pat) Then
                carrier = code: Exit For
            End If
        Next rule

        If Len(carrier) > 0 Then
            ws.Cells(i, COL_COURIER).Value = carrier
            routed = routed + 1
        End If
R5_Next:
    Next i

    RunModule5_RouteClassify = routed
End Function

Function RegexMatch(ByVal text As String, ByVal pattern As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = pattern
    RegexMatch = re.Test(text)
End Function
