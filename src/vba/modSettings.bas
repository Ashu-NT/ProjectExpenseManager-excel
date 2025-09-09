Attribute VB_Name = "modSettings"
' ---------- modSettings ----------
Option Explicit

' A short, simple obfuscation key. This is NOT cryptographically secure.
' Change this if you want a different obfuscation pattern.
Private Const SETTINGS_OBF_KEY As String = "weldS3tt1ngKey"

' ---------- Low-level helpers ----------
' Returns the ListObject for tblSettings or Nothing
Public Function GetSettingsTable() As ListObject
    Set GetSettingsTable = GetTable("tblSettings")
End Function

' Get a setting value as text. If missing, returns defaultValue (if provided) or "".
Public Function GetSetting(key As String, Optional defaultValue As Variant) As Variant
    Dim lo As ListObject, r As Range
    Set lo = GetSettingsTable
    If lo Is Nothing Then
        GetSetting = IIf(IsMissing(defaultValue), vbNullString, defaultValue)
        Exit Function
    End If

    If lo.DataBodyRange Is Nothing Then
        GetSetting = IIf(IsMissing(defaultValue), vbNullString, defaultValue)
        Exit Function
    End If

    For Each r In lo.ListColumns("Key").DataBodyRange.rows
        If Trim(CStr(r.value)) = key Then
            GetSetting = CStr(r.Offset(0, 1).value)
            Exit Function
        End If
    Next r

    GetSetting = IIf(IsMissing(defaultValue), vbNullString, defaultValue)
End Function

' Set (or create) a setting key => value (value stored as string)
Public Sub SetSetting(key As String, value As Variant)
    Dim lo As ListObject, r As Range, added As Boolean
    Set lo = GetSettingsTable
    If lo Is Nothing Then
        MsgBox "Settings table 'tblSettings' not found.", vbCritical
        Exit Sub
    End If

    added = False
    If Not lo.DataBodyRange Is Nothing Then
        For Each r In lo.ListColumns("Key").DataBodyRange.rows
            If Trim(CStr(r.value)) = key Then
                r.Offset(0, 1).value = CStr(value)
                added = True
                Exit For
            End If
        Next r
    End If

    If Not added Then
        Dim lr As ListRow
        Set lr = lo.ListRows.Add
        lr.Range(lo.ListColumns("Key").Index).value = key
        lr.Range(lo.ListColumns("Value").Index).value = CStr(value)
    End If
End Sub

' Initialize defaults if missing (safe to call on workbook open)
Public Sub InitDefaultSettings()
    ' Only write defaults if key missing
    EnsureSettingExists "AutoCommitOnSave", "FALSE"
    EnsureSettingExists "RequireClearBeforeCommit", "TRUE"
    EnsureSettingExists "EnableCommitConfirmation", "TRUE"
    EnsureSettingExists "MaxRowsPerCommit", "500"
    EnsureSettingExists "DateDisplayFormat", "yyyy-mm-dd"
    EnsureSettingExists "AdminPassword_Obf", ""      ' store obfuscated password via SetAdminPassword
    EnsureSettingExists "AllowSettingsEditUsernames", ""   ' comma-separated
    EnsureSettingExists "EnableAuditOnCommit", "TRUE"
    EnsureSettingExists "BackupBeforeCommit", "TRUE"
    EnsureSettingExists "TrustedSheetsLockdown", "TRUE"
End Sub

' Ensure a key exists (adds with defaultValue if missing)
Public Sub EnsureSettingExists(key As String, defaultValue As String)
    Dim lo As ListObject, r As Range, missing As Boolean
    Set lo = GetSettingsTable
    If lo Is Nothing Then Exit Sub
    missing = True
    If Not lo.DataBodyRange Is Nothing Then
        For Each r In lo.ListColumns("Key").DataBodyRange.rows
            If Trim(CStr(r.value)) = key Then missing = False: Exit For
        Next r
    End If
    If missing Then
        Dim lr As ListRow
        Set lr = lo.ListRows.Add
        lr.Range(lo.ListColumns("Key").Index).value = key
        lr.Range(lo.ListColumns("Value").Index).value = defaultValue
    End If
End Sub

' ---------- Simple obfuscation for Admin password ----------
' NOT secure for secrets — acceptable for basic UI protection only.
Public Function ObfuscateString(s As String) As String
    Dim i As Long, out As String, k As String
    k = SETTINGS_OBF_KEY
    If Len(s) = 0 Then ObfuscateString = "": Exit Function
    For i = 1 To Len(s)
        out = out & Chr(Asc(mID(s, i, 1)) Xor Asc(mID(k, ((i - 1) Mod Len(k)) + 1, 1)))
    Next i
    ObfuscateString = EncodeBase64(out)
End Function

Public Function DeobfuscateString(s As String) As String
    Dim raw As String, i As Long, out As String, k As String
    k = SETTINGS_OBF_KEY
    If Len(s) = 0 Then DeobfuscateString = "": Exit Function
    raw = DecodeBase64(s)
    For i = 1 To Len(raw)
        out = out & Chr(Asc(mID(raw, i, 1)) Xor Asc(mID(k, ((i - 1) Mod Len(k)) + 1, 1)))
    Next i
    DeobfuscateString = out
End Function

' --- Proper Base64 encode ---
Public Function EncodeBase64(text As String) As String
    Dim arr() As Byte
    arr = StrConv(text, vbFromUnicode)
    
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = arr
        EncodeBase64 = .text
    End With
End Function

' --- Proper Base64 decode ---
Public Function DecodeBase64(text As String) As String
    Dim arr() As Byte
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
        .DataType = "bin.base64"
        .text = text
        arr = .nodeTypedValue
    End With
    DecodeBase64 = StrConv(arr, vbUnicode)
End Function


' ---------- Admin password helpers ----------
' Call to set admin password (stores obfuscated form in AdminPassword_Obf)
Public Sub SetAdminPassword(plainPassword As String)
    If Len(Trim(plainPassword)) = 0 Then
        MsgBox "Password cannot be empty.", vbExclamation: Exit Sub
    End If
    SetSetting "AdminPassword_Obf", ObfuscateString(plainPassword)
    MsgBox "Admin password set. Please remember it; the stored value is obfuscated.", vbInformation
End Sub

' Verify admin password; returns True if match
Public Function VerifyAdminPassword(plainPassword As String) As Boolean
    Dim storedObf As String
    storedObf = GetSetting("AdminPassword_Obf", "")
    If Trim(storedObf) = "" Then
        ' No password set — treat as open (or force admin to set)
        VerifyAdminPassword = False
        Exit Function
    End If
    VerifyAdminPassword = (DeobfuscateString(storedObf) = plainPassword)
End Function

' ---------- Validation example ----------
' Validates specific settings types before saving; returns True if OK or shows MsgBox
Public Function ValidateSettingValue(key As String, value As String) As Boolean
    Select Case key
        Case "AutoCommitOnSave", "RequireClearBeforeCommit", "EnableCommitConfirmation", "EnableAuditOnCommit", "BackupBeforeCommit", "TrustedSheetsLockdown"
            If UCase(value) <> "TRUE" And UCase(value) <> "FALSE" Then
                MsgBox key & " must be TRUE or FALSE.", vbExclamation: Exit Function
            End If
        Case "MaxRowsPerCommit"
            If Not IsNumeric(value) Or val(value) < 1 Then MsgBox "MaxRowsPerCommit must be a positive number.", vbExclamation: Exit Function
        Case "DateDisplayFormat"
            ' basic sanity: not empty
            If Trim(value) = "" Then MsgBox "DateDisplayFormat cannot be empty.", vbExclamation: Exit Function
        Case Else
            ' no strict validation for unknown keys
    End Select
    ValidateSettingValue = True
End Function
' ---------- end modSettings ----------

' Protects the Settings sheet (locks structure and cells). Use workbook-level password optional
' ----- Replace existing ProtectSettingsSheet / UnprotectSettingsSheet with these -----
Public Sub ProtectSettingsSheet(Optional pwd As String = "")
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    If ws Is Nothing Then Exit Sub
    If Trim(pwd) = "" Then pwd = GetSheetPassword()   ' unify on sheet password
    ws.Unprotect Password:=pwd
    ws.Cells.Locked = True
    ws.rows(1).Locked = True
    ws.Protect Password:=pwd, UserInterfaceOnly:=True, AllowFormattingColumns:=False
End Sub

Public Sub UnprotectSettingsSheet(Optional pwd As String = "")
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    If ws Is Nothing Then Exit Sub
    If Trim(pwd) = "" Then pwd = GetSheetPassword()   ' unify on sheet password
    ws.Unprotect Password:=pwd
End Sub


' ---------------- Form-only password helpers ----------------
' Requires: ObfuscateString(s), DeobfuscateString(s), GetSetting, SetSetting, VerifyAdminPassword

' Set the form-only password (admin-protected). Call this macro to change the form password.
' Usage: run SetFormPassword_Admin from Immediate or assign to an admin-only button
Public Sub SetFormPassword_Admin(plainPassword As String)
    If Len(Trim(plainPassword)) = 0 Then
            MsgBox "Password cannot be empty.", vbExclamation: Exit Sub
    End If
    SetSetting "FormAccessPassword_Obf", ObfuscateString(plainPassword)
    MsgBox "User password set. Please remember it; the stored value is obfuscated.", vbInformation
End Sub

' Verify the form-only password (returns True if match)
Public Function VerifyFormPassword(plainPassword As String) As Boolean
    Dim storedObf As String
    storedObf = GetSetting("FormAccessPassword_Obf", "")
    If Trim(storedObf) = "" Then
        ' No form password set -> return False (caller should fall back to AllowedFormUsernames or other policy)
        VerifyFormPassword = False
        Exit Function
    End If
    VerifyFormPassword = (DeobfuscateString(storedObf) = plainPassword)
End Function



