Attribute VB_Name = "modUIHelpers"
Option Explicit

' Format a single textbox as numeric with N decimals (uses system locale)
Public Sub FormatNumericTextBox(tb As MSForms.TextBox, Optional decimals As Long = 2)
    On Error Resume Next
    If tb Is Nothing Then Exit Sub
    Dim v As String
    v = Trim(tb.text)
    If v = "" Then
        tb.text = ""
        Exit Sub
    End If
    If IsNumeric(v) Then
        tb.text = FormatNumber(val(v), decimals)
    Else
        ' keep text if not numeric (do not overwrite)
    End If
    On Error GoTo 0
End Sub

' ---------- small helper: numeric key filter (use in KeyPress events) ----------

Public Function AllowNumericKey(ByVal KeyAscii As MSForms.ReturnInteger) As Long
    On Error GoTo ErrHandler
    Dim decSep As String
    decSep = Application.International(xlDecimalSeparator) ' locale-aware decimal

    ' Allow Backspace
    If KeyAscii = 8 Then
        AllowNumericKey = KeyAscii
        Exit Function
    End If

    ' Allow digits 0-9
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        AllowNumericKey = KeyAscii
        Exit Function
    End If

    ' Allow decimal characters
    If KeyAscii = Asc(decSep) Or KeyAscii = 46 Or KeyAscii = 44 Then
        AllowNumericKey = KeyAscii
        Exit Function
    End If

    ' Suppress everything else
    AllowNumericKey = 0
    Exit Function

ErrHandler:
    ' Just in case something weird happens
    AllowNumericKey = 0
End Function

' ---------- date validator ----------
Public Function IsValidDateNotFuture(val As Variant) As Boolean
    On Error GoTo Bad
    If Trim(CStr(val)) = "" Then
        IsValidDateNotFuture = False
        Exit Function
    End If
    If Not IsDate(val) Then GoTo Bad
    If CDate(val) > Date Then GoTo Bad
    IsValidDateNotFuture = True
    Exit Function
Bad:
    IsValidDateNotFuture = False
End Function


' Convert a textbox value to Double respecting system decimal/thousand separators.
Public Function ToDbl(ByVal v As Variant) As Double
    Dim s As String, decSep As String, thSep As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then ToDbl = 0#: Exit Function

    decSep = Application.International(xlDecimalSeparator)
    thSep = Application.International(xlThousandsSeparator)

    ' Remove spaces / non-breaking spaces
    s = Replace$(s, Chr$(160), " ")
    s = Replace$(s, " ", "")

    ' Strip thousand sep (if present)
    If thSep <> "" Then s = Replace$(s, thSep, "")

    ' Normalize alternate decimal char
    If decSep = "," Then
        s = Replace$(s, ".", ",")
    Else
        s = Replace$(s, ",", ".")
    End If

    On Error GoTo Bad
    ToDbl = CDbl(s)
    Exit Function
Bad:
    ToDbl = 0#
End Function

' Format a Double to a 2-dec string for display only.
Public Function Fmt2(ByVal d As Double) As String
    Fmt2 = Format(d, "0.00")
End Function

Public Function ParseNum(valIn As Variant) As Double
    On Error GoTo Fallback
    Dim tmp As String
    tmp = CStr(valIn)
    If Len(Trim(tmp)) = 0 Then
        ParseNum = 0
        Exit Function
    End If
    ' Try to call existing ParseNumericFromText if available
    On Error Resume Next
    ParseNum = Application.Run("ParseNumericFromText", tmp)
    If Err.Number = 0 Then Exit Function
    Err.Clear
Fallback:
    ' fallback: remove common currency symbols and thousand separators
    tmp = Replace(tmp, " ", "")
    tmp = Replace(tmp, "€", "")
    tmp = Replace(tmp, "$", "")
    tmp = Replace(tmp, ",", ".")
    If IsNumeric(tmp) Then ParseNum = CDbl(tmp) Else ParseNum = val(tmp)
End Function

