Attribute VB_Name = "modAuth"
Option Explicit

' Tracks if current session user has already verified as Admin
Public gIsAdminVerified As Boolean
Public gCurrentAdminUser As String

' ----------------- modAuth -----------------
' Use these helpers to gate who can open the main form and who can edit settings.

Public Function CurrentUserName() As String
    CurrentUserName = Environ$("USERNAME")
End Function

Public Function IsUserInCsv(userList As String, userName As String) As Boolean
    Dim arr As Variant, u As Variant
    If Trim(userList) = "" Then Exit Function
    arr = Split(userList, ",")
    For Each u In arr
        If LCase(Trim(u)) = LCase(Trim(userName)) Then IsUserInCsv = True: Exit Function
    Next u
End Function


' Wrapper to show the main form but allow form-password (NOT admin)
' Behavior:
'  - If user is in AllowedFormUsernames -> open form (no prompt)
'  - Else, if FormAccessPassword_Obf exists -> prompt for form password and verify only using VerifyFormPassword
'  - Else, fallback to previous behavior (which may prompt Admin password if ShowMainForm does)
Public Sub ShowFormWithFormPassword()
    Dim curUser As String
    Dim isAdminAnswer As VbMsgBoxResult
    Dim inputPwd As String
    Dim allowed As String, formPwdSet As String

    curUser = Environ$("USERNAME")
    allowed = GetSetting("AllowedFormUsernames", "")
    formPwdSet = GetSetting("FormAccessPassword_Obf", "")

    ' First check if current user is explicitly allowed as admin
    isAdminAnswer = MsgBox("Are you an Admin?", vbYesNo + vbQuestion, "Form Access")

    If isAdminAnswer = vbYes Then
        ' Prompt for Admin password
        inputPwd = InputBox("Enter Admin password to open the form:", "Admin Access")
        If StrPtr(inputPwd) = 0 Then Exit Sub ' cancelled
        If VerifyAdminPassword(inputPwd) Then
        
            ' Set global flags
            gIsAdminVerified = True
            gCurrentAdminUser = curUser
            
            frm_UI.Show vbModeless
            Exit Sub
        Else
            MsgBox "Incorrect Admin password. Cannot open the form.", vbExclamation
            Exit Sub
        End If
    Else
        ' Non-admin user
        ' If user is in allowed list, open directly
        If Trim(allowed) <> "" Then
            If IsUserInCsv(allowed, curUser) Then
                frm_UI.Show vbModeless
                Exit Sub
            End If
        End If

        ' If form password is set, prompt for it
        If Trim(formPwdSet) <> "" Then
            inputPwd = InputBox("Enter Form Access password to open the UI:", "User Access")
            If StrPtr(inputPwd) = 0 Then Exit Sub ' cancelled
            If VerifyFormPassword(inputPwd) Then
                gIsAdminVerified = False
                frm_UI.Show vbModeless
                Exit Sub
            Else
                MsgBox "Incorrect Form Access password. Cannot open the form.", vbExclamation
                Exit Sub
            End If
        End If

        ' fallback: no allowed list and no form password set
        frm_UI.Show vbModeless
    End If
End Sub




' Secure settings UI wrapper
Public Sub ShowSettingsForm_Secure()
    ' Securely open the settings form.
    ' If AllowSettingsEditUsernames includes the current user, open without password.
    ' Otherwise prompt for admin password and open if correct.
    Dim curUser As String: curUser = CurrentUserName()
    Dim allowed As String: allowed = GetSetting("AllowSettingsEditUsernames", "")
    
    ' If allowed list present and current user is in it -> allow
    If Trim(allowed) <> "" Then
        If IsUserInCsv(allowed, curUser) Then
            UnprotectSettingsSheet
            frm_Settings.Show vbModal
            Exit Sub
        End If
    End If

    ' Otherwise require admin password (but also allow a blank-admin policy if you want)
    Dim storedObf As String
    storedObf = GetSetting("AdminPassword_Obf", "")
    If Trim(storedObf) = "" Then
        ' No admin password set — decide policy: here we require setting one before editing
        If MsgBox("No admin password is set. Would you like to set one now? (recommended)", vbYesNo + vbQuestion) = vbYes Then
            Dim newPwd As String
            newPwd = InputBox("Enter new Admin password:", "Set Admin password")
            If StrPtr(newPwd) <> 0 And Trim(newPwd) <> "" Then
                SetAdminPassword newPwd
                UnprotectSettingsSheet
                frm_Settings.Show vbModal
                Exit Sub
            Else
                MsgBox "Admin password was not set. Aborting.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Settings editing canceled.", vbInformation
            Exit Sub
        End If
    Else
        ' Ask for admin password
        Dim pwd As String
        pwd = InputBox("Enter Admin password to edit settings:", "Admin login")
        If StrPtr(pwd) = 0 Then Exit Sub
        If VerifyAdminPassword(pwd) Then
            UnprotectSettingsSheet
            frm_Settings.Show vbModal
        Else
            MsgBox "Incorrect password — cannot open Settings.", vbExclamation
        End If
    End If
End Sub


Public Function IsUserAllowedToEditSettings() As Boolean
    Dim allowed As String, curUser As String
    curUser = CurrentUserName()

    ' Check if already verified admin
    If gIsAdminVerified And gCurrentAdminUser = curUser Then
        IsUserAllowedToEditSettings = True
        Exit Function
    End If

    ' Check if user is in allowed list
    allowed = GetSetting("AllowSettingsEditUsernames", "")
    If Trim(allowed) <> "" Then
        If IsUserInCsv(allowed, curUser) Then
            IsUserAllowedToEditSettings = True
            Exit Function
        End If
    End If

    ' Otherwise require admin password
    Dim pwd As String
    pwd = InputBox("Enter Admin password to edit settings:", "Admin login")
    If StrPtr(pwd) = 0 Then Exit Function
    If VerifyAdminPassword(pwd) Then
        gIsAdminVerified = True
        gCurrentAdminUser = curUser
        IsUserAllowedToEditSettings = True
    End If
End Function

Public Sub ShowAdminReportForm()
    Dim curUser As String
    curUser = Environ$("USERNAME")
    ' If admin already verified in session, allow
    If gIsAdminVerified And gCurrentAdminUser = curUser Then
        frm_AdminReport.Show vbModeless
        Exit Sub
    End If

    ' Ask if admin; if yes require Admin password, else reject (only admin allowed here)
    If MsgBox("Admin access required to open Reports. Are you an Admin?", vbYesNo + vbQuestion, "Reports - Admin") <> vbYes Then Exit Sub
    Dim pwd As String
    pwd = InputBox("Enter Admin password:", "Admin authentication")
    If StrPtr(pwd) = 0 Then Exit Sub
    If VerifyAdminPassword(pwd) Then
        gIsAdminVerified = True
        gCurrentAdminUser = curUser
        frm_AdminReport.Show vbModeless
    Else
        MsgBox "Incorrect Admin password.", vbExclamation
    End If
End Sub

