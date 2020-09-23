Attribute VB_Name = "modHook"
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

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Public lpPrevWndProc As Long, gHW As Long, OtherInstanceHwnd As Long, Hooked As Boolean

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Public Sub Hook(sForm As Form)
    gHW = sForm.hwnd
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
    Hooked = True
End Sub

Public Sub MySub(lParam As Long)
    Dim cds As COPYDATASTRUCT
    Dim buf(1 To 255) As Byte, a As String
    Call CopyMemory(cds, ByVal lParam, Len(cds))
    Select Case cds.dwData
        Case 1
            Debug.Print "got a 1"
        Case 2
            Debug.Print "got a 2"
        Case 3
            Call CopyMemory(buf(1), ByVal cds.lpData, cds.cbData)
            a = StrConv(buf, vbUnicode)
            a = Left(a, InStr(1, a, Chr(0)) - 1)
            If StrComp((Right$(a, 3)), "scl", vbTextCompare) = 0 _
            Or StrComp((Right$(a, 3)), "dcd", vbTextCompare) = 0 Then
                If ParseCommand(a) = False Then: MsgBox "Error to Parse Command$ " & vbCrLf & Command$, vbExclamation, "CD Tracker modHook:MySub"
            End If
    End Select
End Sub

Public Sub Unhook()
    Dim temp As Long
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_COPYDATA Then
        Call MySub(lParam)
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function
