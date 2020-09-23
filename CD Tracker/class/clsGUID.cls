VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUIDGenerator"
Attribute VB_GlobalNameSpace = False
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

Private Type GUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Function CreateGUID(Optional strRemoveChars As String = "{}-") As String
Dim udtGUID As GUID
Dim strGUID As String
Dim bytGUID() As Byte
Dim lngLen As Long
Dim lngRetVal As Long
Dim lngPos As Long
On Local Error GoTo ErrorHandler
lngLen = 40
bytGUID = String(lngLen, 0)
CoCreateGuid udtGUID
lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
strGUID = bytGUID
If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then lngRetVal = lngRetVal - 1
strGUID = Left$(strGUID, lngRetVal)
For lngPos = 1 To Len(strRemoveChars)
    strGUID = Replace(strGUID, Mid(strRemoveChars, lngPos, 1), "")
Next
CreateGUID = strGUID
Exit Function
ErrorHandler:
        CreateGUID = "000"
    Err.Clear
End Function


