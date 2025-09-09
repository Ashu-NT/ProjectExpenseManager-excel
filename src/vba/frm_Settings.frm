VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Settings 
   Caption         =   "Application Settings (Admin)"
   ClientHeight    =   10236
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11952
   OleObjectBlob   =   "frm_Settings.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



' Admin-only settings editor
Private Sub UserForm_Initialize()
    ' Ensure only admins open this form - double-check
    'If Not IsUserAllowedToEditSettings() Then
        'MsgBox "You are not authorized to edit settings.", vbExclamation
        'Unload Me
        'Exit Sub
    'End If

    Me.lstSettings.Clear
    LoadSettingsToList
    Me.txtKey.value = ""
    Me.txtValue.value = ""
    Me.txtNotes.value = ""
    Me.lblStatus.Caption = "Loaded settings (" & Me.lstSettings.ListCount & ")"
End Sub

' Load settings table into the ListBox (3 columns)
Public Sub LoadSettingsToList()
    Dim lo As ListObject, r As Range
    Dim i As Long
    Set lo = GetSettingsTable
    Me.lstSettings.Clear
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For Each r In lo.ListColumns("Key").DataBodyRange.rows
        Me.lstSettings.AddItem
        Me.lstSettings.List(Me.lstSettings.ListCount - 1, 0) = r.value
        Me.lstSettings.List(Me.lstSettings.ListCount - 1, 1) = r.Offset(0, 1).value ' Value
        On Error Resume Next
        Me.lstSettings.List(Me.lstSettings.ListCount - 1, 2) = r.Offset(0, ColIndex(lo, "Notes") - lo.ListColumns("Key").Index).value
        On Error GoTo 0
    Next r
End Sub

' Add new setting (does not allow duplicates)
Private Sub btnAdd_Click()
    Dim k As String, v As String, n As String
    k = Trim(Me.txtKey.value)
    v = Trim(Me.txtValue.value)
    n = Trim(Me.txtNotes.value)

    If k = "" Then MsgBox "Key required", vbExclamation: Exit Sub

    ' Validate: no duplicate keys
    Dim keys As Variant, kk As Variant
    keys = AllSettingKeys()
    For Each kk In keys
        If StrComp(kk, k, vbTextCompare) = 0 Then
            MsgBox "A setting with this key already exists. Use Edit to modify it.", vbExclamation
            Exit Sub
        End If
    Next kk

    ' Create
    UpsertSetting k, v, n
    LoadSettingsToList
    Me.lblStatus.Caption = "Added " & k
    ' select it in list
    Dim idx As Long: idx = FindIndexInListBoxByKey(Me.lstSettings, k)
    If idx >= 0 Then Me.lstSettings.ListIndex = idx
End Sub

' Load selected row into edit boxes
Private Sub btnEdit_Click()
    Dim idx As Long
    idx = Me.lstSettings.ListIndex
    If idx < 0 Then MsgBox "Select a setting to load", vbExclamation: Exit Sub
    Dim keySel As String
    keySel = Me.lstSettings.List(idx, 0)
    Me.txtKey.value = keySel
    Me.txtValue.value = Me.lstSettings.List(idx, 1)
    Me.txtNotes.value = Me.lstSettings.List(idx, 2)
    Me.lblStatus.Caption = "Loaded " & keySel & " for editing"
    ' Prevent key edit on protected keys
    If IsProtectedSetting(keySel) Then
        Me.txtKey.Enabled = False
        Me.btnDelete.Enabled = False
    Else
        Me.txtKey.Enabled = True
        Me.btnDelete.Enabled = True
    End If
End Sub

' Save edits to selected (or treat as upsert if new)
Private Sub btnSave_Click()
    Dim oldKey As String, newKey As String, v As String, n As String
    oldKey = ""
    If Me.lstSettings.ListIndex >= 0 Then oldKey = Me.lstSettings.List(Me.lstSettings.ListIndex, 0)
    newKey = Trim(Me.txtKey.value)
    v = Trim(Me.txtValue.value)
    n = Trim(Me.txtNotes.value)

    If newKey = "" Then MsgBox "Key required", vbExclamation: Exit Sub

    ' If editing an existing row and key changed, ensure no other key exists
    If oldKey <> "" And StrComp(oldKey, newKey, vbTextCompare) <> 0 Then
        ' check duplicate
        Dim kk As Variant
        For Each kk In AllSettingKeys()
            If StrComp(kk, newKey, vbTextCompare) = 0 Then
                MsgBox "A different setting already uses that key. Choose a unique key.", vbExclamation
                Exit Sub
            End If
        Next kk
    End If

    ' If setting is protected and key changed, prevent
    If oldKey <> "" And IsProtectedSetting(oldKey) And StrComp(oldKey, newKey, vbBinaryCompare) <> 0 Then
        MsgBox "This setting is protected and its key cannot be changed.", vbExclamation
        Exit Sub
    End If

    ' Upsert (will create new if oldKey empty)
    UpsertSetting newKey, v, n

    ' If oldKey existed and changed, remove oldKey
    If oldKey <> "" And StrComp(oldKey, newKey, vbTextCompare) <> 0 Then
        DeleteSettingByKey oldKey
    End If

    LoadSettingsToList
    Me.lblStatus.Caption = "Saved " & newKey
End Sub

' Delete selected key (double confirmation, protected keys blocked)
Private Sub btnDelete_Click()
    Dim idx As Long, keySel As String
    idx = Me.lstSettings.ListIndex
    If idx < 0 Then MsgBox "Select a setting to delete", vbExclamation: Exit Sub
    keySel = Me.lstSettings.List(idx, 0)
    If IsProtectedSetting(keySel) Then
        MsgBox "This setting is protected and cannot be deleted.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Delete setting '" & keySel & "' (this cannot be undone)?", vbYesNo + vbCritical, "Confirm delete") <> vbYes Then Exit Sub
    If DeleteSettingByKey(keySel) Then
        LoadSettingsToList
        Me.txtKey.value = ""
        Me.txtValue.value = ""
        Me.txtNotes.value = ""
        Me.lblStatus.Caption = "Deleted " & keySel
    Else
        MsgBox "Unable to delete setting.", vbExclamation
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' Utility: find index in listbox by key (returns -1 if not found)
Private Function FindIndexInListBoxByKey(lb As MSForms.ListBox, key As String) As Long
    Dim i As Long
    FindIndexInListBoxByKey = -1
    If lb.ListCount = 0 Then Exit Function
    For i = 0 To lb.ListCount - 1
        If StrComp(lb.List(i, 0), key, vbTextCompare) = 0 Then
            FindIndexInListBoxByKey = i
            Exit Function
        End If
    Next i
End Function


Private Sub btnSetAdminPassword_Click()
    Dim pwd As String
    pwd = InputBox("Enter new Admin password:", "Set Admin Password", "")
    If StrPtr(pwd) = 0 Or Trim(pwd) = "" Then MsgBox "Password not set.", vbInformation: Exit Sub
    SetAdminPassword pwd
    MsgBox "Admin password set.", vbInformation
    ' reload
    Me.lstSettings.Clear
    LoadSettingsToList
End Sub


Private Sub btnSetUserPassword_Click()
    Dim pwd As String
    pwd = InputBox("Enter new User password:", "Set User Password", "")
    If StrPtr(pwd) = 0 Or Trim(pwd) = "" Then MsgBox "Password not set.", vbInformation: Exit Sub
    SetFormPassword_Admin pwd
    MsgBox "User password set.", vbInformation
    ' reload
    Me.lstSettings.Clear
    LoadSettingsToList

End Sub
