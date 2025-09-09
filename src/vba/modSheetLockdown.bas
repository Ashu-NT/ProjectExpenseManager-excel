Attribute VB_Name = "modSheetLockdown"
Option Explicit
' ---- modSheetLockdown ----

Public Function GetSheetPassword() As String
    Dim pwd As String
    pwd = GetSetting("SheetProtectPassword", "")
    If Trim(pwd) = "" Then pwd = "sheetlock"
    GetSheetPassword = pwd
End Function

Public Function AllowedSheetsArray() As Variant
    Dim csv As String, arr As Variant, i As Long
    csv = GetSetting("AllowedVisibleSheets", "UI")
    csv = Trim(csv)
    If csv = "" Then csv = "UI"
    arr = Split(csv, ",")
    For i = LBound(arr) To UBound(arr)
        arr(i) = Trim$(arr(i))
    Next i
    AllowedSheetsArray = arr
End Function

Public Function IsSheetAllowed(sheetName As String) As Boolean
    Dim arr As Variant, s As Variant
    arr = AllowedSheetsArray()
    For Each s In arr
        If StrComp(Trim(s), Trim(sheetName), vbTextCompare) = 0 Then
            IsSheetAllowed = True: Exit Function
        End If
    Next s
End Function

' -------- Apply lockdown  ----------
Public Sub ApplySheetLockdown(Optional userVisibleOnly As Boolean = True)
Dim ws As Worksheet, pwd As String
    Dim arrAllowed As Variant
    Dim s As Variant

    pwd = GetSheetPassword() ' from settings (fallback sheetlock)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ' 1) Ensure workbook is unprotected so we can change sheet visibility
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=pwd
    On Error GoTo CleanFail

    ' 2) Normalize allowed sheet names (uses your existing AllowedSheetsArray)
    arrAllowed = AllowedSheetsArray()

    ' 3) Set visibility for each sheet (VISIBLE for allowed, VERYHIDDEN for others)
    For Each ws In ThisWorkbook.Worksheets
        If userVisibleOnly Then
            If IsSheetAllowed(ws.name) Then
                ws.Visible = xlSheetVisible
            Else
                ws.Visible = xlSheetVeryHidden
            End If
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws

    ' 4) Unprotect ALL allowed sheets and ensure they remain unprotected
    'For Each s In arrAllowed
       ' On Error Resume Next
        'Set ws = Nothing
        'Set ws = ThisWorkbook.Worksheets(CStr(s))
       ' If Not ws Is Nothing Then
            'ws.Unprotect Password:=pwd        ' try sheet password
            'ws.Unprotect                      ' try without password (in case none)
            ' ensure it's visible and unprotected
            'ws.Visible = xlSheetVisible
        'End If
        'On Error GoTo CleanFail
    'Next s

    ' 5) Protect ONLY the disallowed sheets with macro-friendly flags
    'For Each ws In ThisWorkbook.Worksheets
        'If Not IsSheetAllowed(ws.name) Then
           ' On Error Resume Next
            'ws.Unprotect Password:=pwd
            'ws.Protect Password:=pwd, UserInterfaceOnly:=True, _
                       'AllowFiltering:=True, AllowSorting:=True, _
                       AllowInsertingRows:=True, AllowDeletingRows:=True, _
                       AllowFormattingRows:=True, AllowFormattingColumns:=True
            'On Error GoTo CleanFail
        'End If
    'Next ws

    ' 6) Finally protect workbook structure
    ThisWorkbook.Protect Structure:=True, Windows:=False, Password:=pwd

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "ApplySheetLockdown error: " & Err.Number & " - " & Err.Description
    MsgBox "Error applying sheet lockdown: " & Err.Description, vbExclamation
End Sub

' -------- Remove lockdown  ----------
Public Sub RemoveSheetLockdown()
    Dim ws As Worksheet, pwd As String
    pwd = GetSheetPassword()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Clean2

    ' 1) Unprotect workbook so we can change sheet visibility/unprotect sheets
    ThisWorkbook.Unprotect Password:=pwd

    ' 2) Unprotect each sheet and make visible
   
    For Each ws In ThisWorkbook.Worksheets
        ' Try to unprotect with likely passwords
        On Error Resume Next
        If Not TryUnprotectSheet(ws) Then
            ' If still protected, attempt to unprotect with the sheet password explicitly
            ws.Unprotect Password:=GetSheetPassword
            Err.Clear
        End If
        ' Make sure it's visible
        ws.Visible = xlSheetVisible
        ' Finally, ensure it's unprotected
        On Error Resume Next
        ws.Unprotect Password:=GetSheetPassword
        ws.Unprotect Password:="settingslock"
        ws.Unprotect Password:=DeobfuscateString(GetSetting("AdminPassword_Obf", ""))
        Err.Clear
    Next ws


    ' 3) Leave workbook unprotected (admin may want to re-protect later)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Clean2:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Debug.Print "RemoveSheetLockdown error: " & Err.Number & " - " & Err.Description
    MsgBox "Error removing sheet lockdown: " & Err.Description, vbExclamation
End Sub

' -------- Admin unlock (uses admin pwd or sheet pwd per setting) ----------
Public Sub AdminUnlockAllSheets()
    Dim useAdmin As String, promptPwd As String, adminOk As Boolean
    useAdmin = UCase(GetSetting("UseAdminPasswordForUnprotect", "TRUE"))
    adminOk = False

    If useAdmin = "TRUE" Then
        promptPwd = InputBox("Enter Admin password to unlock all sheets:", "Admin unlock")
        If StrPtr(promptPwd) = 0 Then Exit Sub
        If VerifyAdminPassword(promptPwd) Then adminOk = True
    Else
        promptPwd = InputBox("Enter Sheet protection password to unlock all sheets:", "Unlock sheets")
        If StrPtr(promptPwd) = 0 Then Exit Sub
        If promptPwd = GetSheetPassword() Then adminOk = True
    End If

    If Not adminOk Then
        MsgBox "Password incorrect. Unlock aborted.", vbExclamation
        Exit Sub
    End If

    RemoveSheetLockdown
    MsgBox "All sheets unhidden and unprotected. Remember to re-apply lockdown when done.", vbInformation
End Sub

' Ensures ApplySheetLockdown is called at open
Public Sub EnforceLockdownOnOpen()
    On Error Resume Next
    ApplySheetLockdown True
    ' also reapply UserInterfaceOnly flags (in case Excel reset them)
    Dim ws As Worksheet, pwd As String
    pwd = GetSheetPassword()
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Next ws
End Sub

' ---- small validator to notify about invalid sheet names ----
Public Sub ValidateAllowedSheets(arr As Variant)
    Dim s As Variant, found As Boolean, ws As Worksheet, missing As Collection, nm As String
    Set missing = New Collection
    For Each s In arr
        nm = Trim$(s)
        found = False
        For Each ws In ThisWorkbook.Worksheets
            If StrComp(ws.name, nm, vbTextCompare) = 0 Then found = True: Exit For
        Next ws
        If Not found Then missing.Add nm
    Next s
    If missing.Count > 0 Then
        Debug.Print "AllowedVisibleSheets contains names not found in workbook: "
        For Each s In missing
            Debug.Print " - " & s
        Next s
        ' show once to the user (non-modal)
        MsgBox "Warning: AllowedVisibleSheets contains sheet names that don't exist: " & vbCrLf & Join(Application.Transpose(missingToArray(missing)), ", "), vbExclamation, "AllowedVisibleSheets warning"
    End If
End Sub

Private Function missingToArray(col As Collection) As Variant
    Dim a() As String, i As Long
    ReDim a(0 To col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = col(i)
    Next i
    missingToArray = a
End Function


' Try to unprotect a worksheet using a small list of likely passwords.
Public Function TryUnprotectSheet(ws As Worksheet) As Boolean
    Dim pwdCandidates As Variant
    Dim i As Long, p As String

    If ws Is Nothing Then Exit Function

    ' Candidate list: current sheet pwd, legacy "settingslock", admin password (deobf), empty (blank)
    pwdCandidates = Array(GetSheetPassword(), "settingslock", DeobfuscateString(GetSetting("AdminPassword_Obf", "")), "")

    On Error Resume Next
    For i = LBound(pwdCandidates) To UBound(pwdCandidates)
        p = CStr(pwdCandidates(i))
        ws.Unprotect Password:=p
        If Err.Number = 0 Then
            TryUnprotectSheet = True
            Exit For
        Else
            Err.Clear
        End If
    Next i
    On Error GoTo 0
End Function

