VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Nome del Progetto: PrinterListFolder
' ****************************************************************************************************
' Copyright © 2008 - 2009 Nome del programmatore - Società: Nome della società
' Tutti i diritti riservati, Indirizzo Internet
' ****************************************************************************************************
' Attenzione: Questo programma per computer è protetto dalle vigenti leggi sul copyright
' e sul diritto d'autore. Le riproduzioni non autorizzate di questo codice, la sua distribuzione
' la distribuzione anche parziale è considerata una violazione delle leggi, e sarà pertanto
' perseguita con l'estensione massima prevista dalla legge in vigore.
' ****************************************************************************************************

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

Private var_INIFile As String
Private var_DefaultRetVal As String

Private Sub Class_Initialize()
    ResetINIFilePath
    var_DefaultRetVal = ""
End Sub

Private Sub primeINIFile()
    On Error Resume Next
    Open var_INIFile For Append As #1
    Close #1
End Sub

Public Property Let INIFile(ByVal INIFileIn As String)
    var_INIFile = INIFileIn
End Property

Public Property Let DefaultReturnValue(ByVal RetVal As String)
    var_DefaultRetVal = RetVal
End Property


Public Property Get INIFile() As String
    INIFile = var_INIFile
End Property

Public Property Get DefaultReturnValue() As String
    DefaultReturnValue = var_DefaultRetVal
End Property

Public Sub ResetINIFilePath()
    If Right$(App.Path, 1) = "\" Then
        var_INIFile = App.Path & App.EXEName & "_setting.ini"
    Else
        var_INIFile = App.Path & "\" & App.EXEName & "_setting.ini"
    End If
End Sub


Public Sub CreateKeyValue(ByVal section As String, ByVal key As String, ByVal value As String, Optional ByVal INIFileLoc As String)
Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    WritePrivateProfileString section, key, value, UseFile
End Sub

Public Sub CreateKey(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    WritePrivateProfileString section, key, "", UseFile
End Sub

Public Function GetKeyValue(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    primeINIFile
    RetVal = String$(255, 0)
    KeyLen = GetPrivateProfileString(section, key, "", RetVal, Len(RetVal), UseFile)
    If KeyLen = 0 Then
        GetKeyValue = ""
    Else
        GetKeyValue = Left$(RetVal, KeyLen)
    End If
End Function

Public Function GetKeysInSection(ByVal section As String, Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    primeINIFile
    RetVal = String$(1024, 0)
    KeyLen = GetPrivateProfileString(section, vbNullString, var_DefaultRetVal, RetVal, Len(RetVal), UseFile)
    If KeyLen = 0 Then
        GetKeysInSection = ""
    Else
        GetKeysInSection = Left$(RetVal, KeyLen - 1)
    End If
End Function

Public Function GetSections(Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    primeINIFile
    RetVal = String$(1024, 0)
    KeyLen = GetPrivateProfileString(vbNullString, vbNullString, var_DefaultRetVal, RetVal, Len(RetVal), UseFile)
    If KeyLen = 0 Then
        GetSections = ""
    Else
        GetSections = Left$(RetVal, KeyLen - 1)
    End If
End Function

Public Sub DeleteKeyValue(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    WritePrivateProfileString section, key, "", UseFile
End Sub

Public Sub DeleteKey(ByVal section As String, ByVal key As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    WritePrivateProfileString section, key, vbNullString, UseFile
End Sub

Public Sub DeleteSection(ByVal section As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    WritePrivateProfileString section, vbNullString, vbNullString, UseFile
End Sub

Public Sub RenameKey(ByVal section As String, ByVal CurrentKey As String, ByVal NewKey As String, Optional ByVal INIFileLoc As String)
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    primeINIFile
    CreateKeyValue section, NewKey, GetKeyValue(section, CurrentKey, UseFile), UseFile
    DeleteKey section, CurrentKey, UseFile
End Sub

Public Sub RenameSection(ByVal CurrentSection As String, ByVal NewSection As String, Optional ByVal INIFileLoc As String)
    Dim KeyList() As String
    Dim KeyCount As Integer
    Dim index As Integer
    Dim UseFile As String
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        UseFile = var_INIFile
    Else
        UseFile = INIFileLoc
    End If
    primeINIFile
    KeyList = Split(GetKeysInSection(CurrentSection, UseFile), Chr$(0))
    KeyCount = UBound(KeyList)
    For index = 0 To KeyCount
        CreateKeyValue NewSection, KeyList(index), GetKeyValue(CurrentSection, KeyList(index), UseFile), UseFile
    Next index
    DeleteSection CurrentSection, UseFile
End Sub
