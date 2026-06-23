REM  =====================================================================
REM  MovingAverages.bas
REM
REM  LibreOffice Calc Basic module providing two user-defined functions:
REM
REM    SMA(data, period)  - Simple Moving Average
REM    EMA(data, period)  - Exponential Moving Average
REM
REM  Both are array functions: select an output range the same height as
REM  the input, type the formula, and confirm with Ctrl+Shift+Enter.
REM  =====================================================================

Option Explicit


REM  ---------------------------------------------------------------------
REM  SMA - Simple Moving Average (cell function / array formula)
REM  ---------------------------------------------------------------------
Function SMA(ByVal data As Variant, ByVal period As Long) As Variant
    If period < 1 Then
        SMA = "Error: period must be >= 1"
        Exit Function
    End If

    Dim values() As Double
    Dim res() As Double
    Dim validFrom As Long
    values = FlattenToDoubles(data)
    res = ComputeSMA(values, period, validFrom)
    SMA = ToColumn(res, validFrom)
End Function


REM  ---------------------------------------------------------------------
REM  EMA - Exponential Moving Average (cell function / array formula)
REM  ---------------------------------------------------------------------
Function EMA(ByVal data As Variant, ByVal period As Long) As Variant
    If period < 1 Then
        EMA = "Error: period must be >= 1"
        Exit Function
    End If

    Dim values() As Double
    Dim res() As Double
    Dim validFrom As Long
    values = FlattenToDoubles(data)
    res = ComputeEMA(values, period, validFrom)
    EMA = ToColumn(res, validFrom)
End Function


REM  =====================================================================
REM  Shared math helpers.
REM
REM  Each takes a 1-based array of Doubles and returns a 1-based array of
REM  Doubles. validFrom is set (ByRef) to the first index that holds a real
REM  result; everything before it should be treated as "not enough data".
REM  =====================================================================

REM  ComputeSMA - trailing average of the last <period> values.
Private Function ComputeSMA(values() As Double, ByVal period As Long, ByRef validFrom As Long) As Double()
    Dim n As Long, i As Long
    Dim total As Double
    Dim res() As Double

    n = UBound(values)
    ReDim res(1 To n)
    If period < 1 Then period = 1
    If n >= period Then validFrom = period Else validFrom = n + 1

    total = 0
    For i = 1 To n
        total = total + values(i)
        If i > period Then total = total - values(i - period)
        If i >= period Then res(i) = total / period
    Next i

    ComputeSMA = res
End Function


REM  ComputeEMA - smoothing factor alpha = 2 / (period + 1), seeded with
REM  the simple average of the first <period> values.
Private Function ComputeEMA(values() As Double, ByVal period As Long, ByRef validFrom As Long) As Double()
    Dim n As Long, i As Long
    Dim alpha As Double, seed As Double, prev As Double
    Dim res() As Double

    n = UBound(values)
    ReDim res(1 To n)
    If period < 1 Then period = 1
    alpha = 2 / (period + 1)

    If n < period Then
        validFrom = n + 1
        ComputeEMA = res
        Exit Function
    End If

    validFrom = period
    seed = 0
    For i = 1 To period
        seed = seed + values(i)
    Next i
    prev = seed / period
    res(period) = prev

    For i = period + 1 To n
        prev = alpha * values(i) + (1 - alpha) * prev
        res(i) = prev
    Next i

    ComputeEMA = res
End Function


REM  =====================================================================
REM  Conversion utilities.
REM  =====================================================================

REM  ToColumn - turn a 1-based Double array into a 2-D Variant column for a
REM  cell array formula. Indexes before validFrom become empty strings.
Private Function ToColumn(res() As Double, ByVal validFrom As Long) As Variant
    Dim n As Long, i As Long
    Dim out() As Variant
    n = UBound(res)
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        If i < validFrom Then
            out(i, 1) = ""
        Else
            out(i, 1) = res(i)
        End If
    Next i
    ToColumn = out
End Function


REM  FlattenToDoubles - normalize a range/array argument into a 1-based
REM  1-D array of Doubles. Non-numeric / blank cells are treated as 0.
Private Function FlattenToDoubles(ByVal data As Variant) As Double()
    Dim out() As Double
    Dim r As Long, c As Long
    Dim idx As Long
    Dim rows As Long, cols As Long

    If Not IsArray(data) Then
        ReDim out(1 To 1)
        out(1) = ToDouble(data)
        FlattenToDoubles = out
        Exit Function
    End If

    ' Cell ranges arrive as 2-D arrays (rows, cols).
    On Error GoTo OneDim
    rows = UBound(data, 1) - LBound(data, 1) + 1
    cols = UBound(data, 2) - LBound(data, 2) + 1

    ReDim out(1 To rows * cols)
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


REM  ToDouble - safe numeric conversion; blanks/text become 0.
Private Function ToDouble(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        ToDouble = CDbl(v)
    Else
        ToDouble = 0
    End If
End Function
