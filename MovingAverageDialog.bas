REM  =====================================================================
REM  MovingAverageDialog.bas
REM
REM  A dialog box for the moving-average tool. It lets you:
REM    - select the input data range,
REM    - choose the starting cell for the output,
REM    - enter one or more window sizes (comma-separated),
REM    - tick which averages to produce (SMA, EMA, or both).
REM
REM  On OK it computes each requested series (reusing ComputeSMA /
REM  ComputeEMA from MovingAverages.bas) and writes one column per
REM  measure/window, aligned row-for-row with the input data, with a
REM  header label in the row above the output start cell.
REM
REM  Run it from Tools > Macros, or bind it to a toolbar button.
REM  The dialog is built entirely in code, so nothing needs to be drawn
REM  in the Dialog Editor.
REM  =====================================================================

Option Explicit


Sub ShowMovingAverageDialog()
    Dim oModel As Object, oDialog As Object
    Dim txtInput As Object, txtOutput As Object, txtWindows As Object
    Dim chkSMA As Object, chkEMA As Object
    Dim iResult As Integer

    ' --- Build the dialog model -------------------------------------------
    oModel = createUnoService("com.sun.star.awt.UnoControlDialogModel")
    oModel.Width = 210
    oModel.Height = 134
    oModel.Title = "Moving Averages"

    AddControl(oModel, "FixedText", "lblInput",   6,   6, 198, 10).Label = "Input data range (e.g. A2:A100):"
    txtInput   = AddControl(oModel, "Edit", "txtInput",   6,  17, 198, 12)

    AddControl(oModel, "FixedText", "lblOutput",  6,  34, 198, 10).Label = "Output start cell (e.g. D2):"
    txtOutput  = AddControl(oModel, "Edit", "txtOutput",  6,  45, 198, 12)

    AddControl(oModel, "FixedText", "lblWindows", 6,  62, 198, 10).Label = "Window sizes, comma-separated (e.g. 5,10,20):"
    txtWindows = AddControl(oModel, "Edit", "txtWindows", 6,  73, 198, 12)

    chkSMA = AddControl(oModel, "CheckBox", "chkSMA",   6,  92,  95, 10)
    chkSMA.Label = "Simple MA (SMA)"
    chkSMA.State = 1
    chkEMA = AddControl(oModel, "CheckBox", "chkEMA", 105,  92,  99, 10)
    chkEMA.Label = "Exponential MA (EMA)"
    chkEMA.State = 1

    Dim oOK As Object, oCancel As Object
    oOK = AddControl(oModel, "Button", "btnOK", 96, 112, 50, 14)
    oOK.Label = "OK"
    oOK.PushButtonType = com.sun.star.awt.PushButtonType.OK
    oOK.DefaultButton = True
    oCancel = AddControl(oModel, "Button", "btnCancel", 152, 112, 50, 14)
    oCancel.Label = "Cancel"
    oCancel.PushButtonType = com.sun.star.awt.PushButtonType.CANCEL

    ' --- Pre-fill the input range from the current selection --------------
    txtWindows.Text = "10"
    PrefillFromSelection(txtInput)

    ' --- Show it ----------------------------------------------------------
    oDialog = createUnoService("com.sun.star.awt.UnoControlDialog")
    oDialog.setModel(oModel)
    oDialog.setVisible(False)
    oDialog.createPeer(createUnoService("com.sun.star.awt.Toolkit"), Null)

    iResult = oDialog.execute()
    If iResult <> com.sun.star.awt.PushButtonType.OK Then
        oDialog.dispose()
        Exit Sub
    End If

    Dim sInput As String, sOutput As String, sWindows As String
    Dim bSMA As Boolean, bEMA As Boolean
    sInput   = Trim(txtInput.Text)
    sOutput  = Trim(txtOutput.Text)
    sWindows = Trim(txtWindows.Text)
    bSMA     = (chkSMA.State = 1)
    bEMA     = (chkEMA.State = 1)
    oDialog.dispose()

    RunMovingAverages(sInput, sOutput, sWindows, bSMA, bEMA)
End Sub


REM  ---------------------------------------------------------------------
REM  Validate inputs, compute the requested series, and write them out.
REM  ---------------------------------------------------------------------
Private Sub RunMovingAverages(ByVal sInput As String, ByVal sOutput As String, _
                              ByVal sWindows As String, ByVal bSMA As Boolean, ByVal bEMA As Boolean)
    Dim oSheet As Object
    Dim oInputRange As Object, oOutputCell As Object
    Dim windows() As Long
    Dim nWindows As Long

    If Not (bSMA Or bEMA) Then
        MsgBox "Select at least one of SMA or EMA.", 48, "Moving Averages"
        Exit Sub
    End If

    windows = ParseWindows(sWindows, nWindows)
    If nWindows = 0 Then
        MsgBox "Enter at least one valid (whole, positive) window size.", 48, "Moving Averages"
        Exit Sub
    End If

    oSheet = ThisComponent.CurrentController.ActiveSheet

    On Error GoTo RangeError
    oInputRange = ResolveRange(sInput)
    oOutputCell = ResolveRange(sOutput)
    On Error GoTo 0

    ' --- Read the input column --------------------------------------------
    Dim values() As Double
    Dim n As Long
    values = FlattenToDoubles(oInputRange.getDataArray())
    n = UBound(values)
    If n < 1 Then
        MsgBox "The input range is empty.", 48, "Moving Averages"
        Exit Sub
    End If

    ' --- Output anchor ----------------------------------------------------
    Dim oOutSheet As Object, oAddr As Object
    oOutSheet = oOutputCell.getSpreadsheet()
    oAddr = oOutputCell.getRangeAddress()
    Dim startCol As Long, startRow As Long
    startCol = oAddr.StartColumn
    startRow = oAddr.StartRow

    ' --- Compute and write one column per measure/window ------------------
    Dim colOffset As Long
    Dim w As Long
    colOffset = 0

    For w = 0 To nWindows - 1
        If bSMA Then
            WriteSeries(oOutSheet, startCol + colOffset, startRow, _
                        "SMA" & windows(w), ComputeOne(values, windows(w), True))
            colOffset = colOffset + 1
        End If
        If bEMA Then
            WriteSeries(oOutSheet, startCol + colOffset, startRow, _
                        "EMA" & windows(w), ComputeOne(values, windows(w), False))
            colOffset = colOffset + 1
        End If
    Next w

    Exit Sub

RangeError:
    MsgBox "Could not read the range. Check the input range and output cell." & Chr(10) & _
           "Input: '" & sInput & "'   Output: '" & sOutput & "'", 16, "Moving Averages"
End Sub


REM  ---------------------------------------------------------------------
REM  ComputeOne - dispatch to ComputeSMA or ComputeEMA and package the
REM  result with its validFrom index so WriteSeries knows where data starts.
REM  Returns a Variant array: element 0 = Double() values, 1 = validFrom.
REM  ---------------------------------------------------------------------
Private Function ComputeOne(values() As Double, ByVal period As Long, ByVal useSMA As Boolean) As Variant
    Dim res() As Double
    Dim validFrom As Long
    If useSMA Then
        res = ComputeSMA(values, period, validFrom)
    Else
        res = ComputeEMA(values, period, validFrom)
    End If
    Dim packed(1) As Variant
    packed(0) = res
    packed(1) = validFrom
    ComputeOne = packed
End Function


REM  ---------------------------------------------------------------------
REM  WriteSeries - write a header (row above the anchor) plus the values,
REM  aligned row-for-row with the input. Cells before validFrom stay blank.
REM  ---------------------------------------------------------------------
Private Sub WriteSeries(oSheet As Object, ByVal col As Long, ByVal startRow As Long, _
                        ByVal header As String, ByVal packed As Variant)
    Dim res() As Double
    Dim validFrom As Long
    Dim n As Long, i As Long
    res = packed(0)
    validFrom = packed(1)
    n = UBound(res)

    If startRow > 0 Then
        oSheet.getCellByPosition(col, startRow - 1).setString(header)
    End If

    For i = 1 To n
        If i >= validFrom Then
            oSheet.getCellByPosition(col, startRow + (i - 1)).setValue(res(i))
        End If
    Next i
End Sub


REM  =====================================================================
REM  Small helpers.
REM  =====================================================================

REM  AddControl - create a control model, position it, register it, and
REM  return it so the caller can set type-specific properties.
Private Function AddControl(oModel As Object, ByVal sType As String, ByVal sName As String, _
                            ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Object
    Dim oCtrl As Object
    oCtrl = oModel.createInstance("com.sun.star.awt.UnoControl" & sType & "Model")
    oCtrl.PositionX = x
    oCtrl.PositionY = y
    oCtrl.Width = w
    oCtrl.Height = h
    oCtrl.Name = sName
    oModel.insertByName(sName, oCtrl)
    AddControl = oCtrl
End Function


REM  PrefillFromSelection - if the current selection is a cell range, drop
REM  its address into the input field.
Private Sub PrefillFromSelection(txtInput As Object)
    On Error Resume Next
    Dim oSel As Object
    oSel = ThisComponent.CurrentSelection
    If IsNull(oSel) Then Exit Sub
    If oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then
        txtInput.Text = oSel.AbsoluteName
    End If
End Sub


REM  ParseWindows - split a comma-separated string into a Long() of valid,
REM  positive whole window sizes. nCount returns how many were found.
Private Function ParseWindows(ByVal s As String, ByRef nCount As Long) As Variant
    Dim parts() As String
    Dim out() As Long
    Dim i As Long, v As Long
    Dim piece As String

    parts = Split(s, ",")
    ReDim out(0 To UBound(parts))
    nCount = 0
    For i = 0 To UBound(parts)
        piece = Trim(parts(i))
        If piece <> "" And IsNumeric(piece) Then
            v = CLng(Val(piece))
            If v >= 1 And CDbl(v) = Val(piece) Then
                out(nCount) = v
                nCount = nCount + 1
            End If
        End If
    Next i
    If nCount > 0 Then ReDim Preserve out(0 To nCount - 1)
    ParseWindows = out
End Function


REM  ResolveRange - turn a user-typed range/cell string into a cell-range
REM  object. Accepts an optional sheet prefix ("Sheet1.A2:A100") and $
REM  anchors; otherwise resolves against the active sheet.
Private Function ResolveRange(ByVal sName As String) As Object
    Dim oSheet As Object
    Dim sLocal As String
    Dim dotPos As Long

    sName = Trim(sName)
    sName = Replace(sName, "$", "")

    dotPos = InStr(sName, ".")
    If dotPos > 0 Then
        Dim sSheet As String
        sSheet = Left(sName, dotPos - 1)
        sLocal = Mid(sName, dotPos + 1)
        oSheet = ThisComponent.Sheets.getByName(sSheet)
    Else
        oSheet = ThisComponent.CurrentController.ActiveSheet
        sLocal = sName
    End If

    ResolveRange = oSheet.getCellRangeByName(sLocal)
End Function
