REM  =====================================================================
REM  MovingAverages.bas
REM
REM  LibreOffice Calc Basic module providing one user-defined function:
REM
REM    MOVAVG(data, period, [type])  - Moving Average
REM        type "S" (or omitted) -> Simple Moving Average
REM        type "E"              -> Exponential Moving Average
REM
REM  It is an ordinary (scalar) cell function: it returns the moving
REM  average value at the END of the supplied range. Anchor the start of
REM  the range and fill down, e.g. put  =MOVAVG(A$2:A2, 5)  in B2 and drag
REM  it down the column. Rows that don't yet have <period> values come back
REM  blank. Anchoring the start ($) matters for the EMA, which accumulates
REM  from the first value.
REM
REM  Examples:
REM    =MOVAVG(A$2:A2, 5)        ' simple (default)
REM    =MOVAVG(A$2:A2, 5, "E")   ' exponential
REM    =MOVAVG(A$2:A2, 5, "S")   ' simple, explicit
REM  =====================================================================

Option Explicit


REM  ---------------------------------------------------------------------
REM  MOVAVG - Moving Average at the end of <data>.
REM    period  number of values in the window (whole number >= 1)
REM    type    "S"/omitted = simple trailing average of the last <period>
REM            values; "E" = exponential (alpha = 2 / (period + 1), seeded
REM            with the simple average of the first <period> values and
REM            rolled forward to the end of the range).
REM  ---------------------------------------------------------------------
Function MOVAVG(ByVal data As Variant, ByVal period As Long, Optional ByVal kind As Variant) As Variant
    Dim k As String
    Dim values As Variant
    Dim n As Long, i As Long
    Dim total As Double, alpha As Double, seed As Double, prev As Double

    If period < 1 Then
        MOVAVG = "Error: period must be >= 1"
        Exit Function
    End If

    If IsMissing(kind) Then
        k = "S"
    Else
        k = UCase(Trim(CStr(kind)))
    End If
    If k = "" Then k = "S"

    values = FlattenToDoubles(data)
    n = UBound(values)
    If n < period Then
        MOVAVG = ""        ' not enough data yet
        Exit Function
    End If

    If Left(k, 1) = "E" Then
        ' --- Exponential moving average ---
        seed = 0
        For i = 1 To period
            seed = seed + values(i)
        Next i
        prev = seed / period

        alpha = 2 / (period + 1)
        For i = period + 1 To n
            prev = alpha * values(i) + (1 - alpha) * prev
        Next i
        MOVAVG = prev
    Else
        ' --- Simple moving average ---
        total = 0
        For i = n - period + 1 To n
            total = total + values(i)
        Next i
        MOVAVG = total / period
    End If
End Function


REM  =====================================================================
REM  Helpers.
REM  =====================================================================

REM  FlattenToDoubles - normalize a range / array / scalar argument into a
REM  1-based array of Doubles (returned inside a Variant). Cell ranges
REM  arrive as 2-D arrays; non-numeric / blank cells are treated as 0.
Private Function FlattenToDoubles(ByVal data As Variant) As Variant
    Dim out() As Double
    Dim r As Long, c As Long, idx As Long
    Dim nRows As Long, nCols As Long

    If Not IsArray(data) Then
        ReDim out(1 To 1)
        out(1) = ToDouble(data)
        FlattenToDoubles = out
        Exit Function
    End If

    On Error GoTo OneDim
    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1

    ReDim out(1 To nRows * nCols)
    idx = 0
    For r = LBound(data, 1) To UBound(data, 1)
        For c = LBound(data, 2) To UBound(data, 2)
            idx = idx + 1
            out(idx) = ToDouble(data(r, c))
        Next c
    Next r
    FlattenToDoubles = out
    Exit Function

OneDim:
    On Error GoTo 0
    ReDim out(1 To UBound(data) - LBound(data) + 1)
    idx = 0
    For r = LBound(data) To UBound(data)
        idx = idx + 1
        out(idx) = ToDouble(data(r))
    Next r
    FlattenToDoubles = out
End Function


REM  ToDouble - safe numeric conversion; blanks / text become 0.
Private Function ToDouble(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        ToDouble = CDbl(v)
    Else
        ToDouble = 0
    End If
End Function
