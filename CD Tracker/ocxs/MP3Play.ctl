VERSION 5.00
Begin VB.UserControl MP3Play 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   675
   Picture         =   "MP3Play.ctx":0000
   PropertyPages   =   "MP3Play.ctx":0ECA
   ScaleHeight     =   690
   ScaleWidth      =   675
   ToolboxBitmap   =   "MP3Play.ctx":0EDB
End
Attribute VB_Name = "MP3Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
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

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Default Property Values:
Const m_def_FileName = ""

'Property Variables:
Dim m_FileName As String

Private Sub UserControl_Resize()
    UserControl.Height = 690
    UserControl.Width = 675
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileName() As String
Attribute FileName.VB_Description = "The file that will be played"
Attribute FileName.VB_ProcData.VB_Invoke_Property = "ppFileName"
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FileName = m_def_FileName
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
End Sub

Private Sub UserControl_Terminate()
    mmStop
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
End Sub

Public Function IsPlaying() As Boolean
    Static s As String * 30
    mciSendString "status MP3Play mode", s, Len(s), 0
    If Mid$(s, 1, 7) = "playing" Then IsPlaying = True Else IsPlaying = False
End Function

Public Function mmPlay()
    Dim cmdToDo As String * 255
    Dim dwReturn As Long
    Dim ret As String * 128

    Dim tmp As String * 255
    Dim lenShort As Long
    Dim ShortPathAndFie As String
    
    If Dir$(FileName) = "" Then
        mmOpen = "Error with input file"
        Exit Function
    End If
    lenShort = GetShortPathName(FileName, tmp, 255)
    ShortPathAndFie = Left$(tmp, lenShort)
    glo_hWnd = hwnd
    cmdToDo = "open " & ShortPathAndFie & " type MPEGVideo Alias MP3Play"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)

    If dwReturn <> 0 Then
        mciGetErrorString dwReturn, ret, 128
        mmOpen = ret
        MsgBox ret, vbCritical, "MP3 Player"
        Exit Function
    End If
    
    mmOpen = "Success"
    mciSendString "play MP3Play", 0, 0, 0
End Function

Public Function mmPause()
    If IsPlaying = True Then
        mciSendString "pause MP3Play", 0, 0, 0
    Else
        mciSendString "play MP3Play", 0, 0, 0
    End If
End Function

Public Function mmStop() As String
    mciSendString "stop MP3Play", 0, 0, 0
    mciSendString "close MP3Play", 0, 0, 0
End Function

Public Function PositionInSec()
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", s, Len(s), 0
    PositionInSec = Round(Mid$(s, 1, Len(s)) / 1000)
End Function

Public Function Position()
    Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play position", s, Len(s), 0
    sec = Round(Mid$(s, 1, Len(s)) / 1000)
    If sec < 60 Then Position = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        Position = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function LengthInSec()
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    LengthInSec = Round(Val(Mid$(s, 1, Len(s))) / 1000) 'Round(CInt(Mid$(s, 1, Len(s))) / 1000)
End Function

Public Function length()
Static s As String * 30
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    mciSendString "status MP3Play length", s, Len(s), 0
    sec = Round(Val(Mid$(s, 1, Len(s))) / 1000)
    If sec < 60 Then length = "0:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        length = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Public Function About()
Attribute About.VB_UserMemId = -552
    MsgBox "MP3 Player v1.0.1 by Salvo cortesiano." _
    & vbCrLf & "All Right Reserved!", vbInformation, "MP3 Player"
End Function

Public Function SeekTo(Second)
    Dim ret As Long
    'On Local Error GoTo ErrorHandler
    If IsPlaying Then
        ret = mciSendString("seek MP3Play from " + Second, vbNullString, 0&, 0&)
    Else
        ret = mciSendString("seek MP3Play to " + Second, vbNullString, 0&, 0&)
    End If
Exit Function
'ErrorHandler:
        'MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, "MP3 Player"
    'Err.Clear
End Function
Public Function OpenCloseCD(drvLetter As String, Optional strCloseorOpen As Boolean = True)
    Dim mssg As String * 255: Dim ReturnValue As Long
    On Local Error Resume Next
    drvLetter = drvLetter & ":\"
    ReturnValue = mciSendString("open " & drvLetter & " Type cdaudio Alias cd", mssg, 255, 0)
    If strCloseorOpen Then
        ReturnValue = mciSendString("set cd door open", vbNullString, 0, 0)
    Else
        ReturnValue = mciSendString("set cd door close", vbNullString, 0, 0)
    End If
    ReturnValue = mciSendString("close cd", 0&, 0, 0)
End Function
