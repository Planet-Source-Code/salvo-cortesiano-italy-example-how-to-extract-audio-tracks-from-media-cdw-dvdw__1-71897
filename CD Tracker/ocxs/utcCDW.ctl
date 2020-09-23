VERSION 5.00
Begin VB.UserControl utcCDW 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   InvisibleAtRuntime=   -1  'True
   Picture         =   "utcCDW.ctx":0000
   ScaleHeight     =   525
   ScaleWidth      =   600
   ToolboxBitmap   =   "utcCDW.ctx":08CA
   Begin VB.ListBox lstdrv 
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   825
      Width           =   2055
   End
End
Attribute VB_Name = "utcCDW"
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

Public Enum DisplayInfo
  [Track_Min_Sec_Mil] = 0
  [Track_Min_Sec] = 1
  [Min_Sec_Mil] = 2
  [Min_Sec] = 3
End Enum

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private mciOpen As Boolean

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function getdrivetype Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const Drive_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DriverVersion = 0

Private Declare Function NetShareGetInfo Lib "NETAPI32" (ByRef ServerName As Byte, ByRef NetName As Byte, ByVal Level As Long, ByRef Buffer As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (bufptr As Any) As Long

Private SharedDrv As Boolean
Private VName As String
Private DriveSet As String
Private drvCaption As String

Private Declare Function DoesFileExist Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1

Private AutorunPath As String
Private Sub UserControl_Initialize()
    If GetCDAudioList = False Then:
End Sub




Private Sub UserControl_Resize()
    UserControl.Height = 480
    UserControl.Width = 480
End Sub

Public Function CloseCD() As Boolean
    mciOpen = False
    CloseCD = Not CBool(mciSendString("close cdr", vbNullString, 0, 0))
End Function

Public Function EjectCD() As Boolean
EjectCD = False
If mciOpen Then
    mciSendString "set cdr door open", vbNullString, 0, 0
    EjectCD = True
End If
End Function

Public Function GetCDID() As String
Dim ReturnString As String, i As Integer
ReturnString = Space$(64)
If mciOpen Then
    mciSendString "info cdr identity", ReturnString, 64, 0
    GetCDID = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
    Else
    GetCDID = "device_not_open"
End If
End Function

Public Function MediaPresent() As Boolean
Dim ReturnString As String * 30
MediaPresent = False
If mciOpen Then
    mciSendString "status cdr media present", ReturnString, Len(ReturnString), 0
    MediaPresent = CBool(ReturnString)
End If
End Function

Public Function OpenCD(CD_Drive As String) As Boolean
    Dim ReturnString As String * 30
    CloseCD
    CD_Drive = Mid$(CD_Drive, 1, 1) + ":\"
    OpenCD = Not CBool(mciSendString("open " + CD_Drive + " Type cdaudio alias cdr wait shareable", ReturnString, Len(ReturnString), 0))
    If OpenCD Then mciOpen = True Else mciOpen = False
    mciSendString "set cdr time format tmsf wait", vbNullString, 0, 0
End Function

Public Function GetCDLength() As String
Dim ReturnString As String * 30
If mciOpen Then
    If MediaPresent Then
        mciSendString "status cdr length wait", ReturnString, Len(ReturnString), 0
        GetCDLength = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
        Else
        GetCDLength = "no_disc_present"
    End If
    Else
    GetCDLength = "device_not_open"
End If
End Function

Public Function GetNumberOfTracks() As Integer
Dim ReturnString As String * 30
If MediaPresent Then
    mciSendString "status cdr number of tracks wait", ReturnString, Len(ReturnString), 0
    GetNumberOfTracks = CInt(Mid$(ReturnString, 1, 2))
    Else
    GetNumberOfTracks = 0
End If
End Function

Public Function GetCDStatus() As String
Dim ReturnString As String * 30
If mciOpen Then
    mciSendString "status cdr mode", ReturnString, Len(ReturnString), 0
    GetCDStatus = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
    Else
    GetCDStatus = "device_not_open"
End If
End Function

Public Function GetCurrentPosition(strSetPosition As DisplayInfo) As String
    Dim ReturnString As String * 30
    Static s As String * 30
    Dim sec, min, mil, track
    Dim Status As String
    On Local Error GoTo ErrorHandler
    GetCurrentPosition = "device_not_open"
    Select Case strSetPosition
        Case Track_Min_Sec_Mil
    If mciOpen Then
        If MediaPresent Then
            mciSendString "status cdr position", ReturnString, Len(ReturnString), 0
            GetCurrentPosition = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
        Else
            GetCurrentPosition = "no_disc_present"
        End If
    End If
        Case Min_Sec, Track_Min_Sec, Min_Sec_Mil
            If mciOpen Then
                If MediaPresent Then
                    mciSendString "status cdr position", s, Len(s), 0
                    track = CInt(Mid$(s, 1, 2))
                    min = CInt(Mid$(s, 4, 2))
                    sec = CInt(Mid$(s, 7, 2))
                    mil = CInt(Mid$(s, 10, 2))
                If strSetPosition = Min_Sec Then
                    GetCurrentPosition = Format(min, "00") & ":" & Format(sec, "00")
                ElseIf strSetPosition = Track_Min_Sec Then
                    GetCurrentPosition = track & "- " & Format(min, "00") & ":" & Format(sec, "00")
                ElseIf strSetPosition = Min_Sec_Mil Then
                    GetCurrentPosition = Format(min, "00") & ":" & Format(sec, "00") & ":" & Format(mil, "00")
                End If
            Else
                GetCurrentPosition = "no_disc_present"
            End If
        End If
    End Select
Exit Function
ErrorHandler:
        GetCurrentPosition = "Error!"
    Err.Clear
End Function
Public Function GetTrackLength(track As Long, Optional getFormattedTime As Boolean = True) As String
Dim ReturnString As String * 30
If mciOpen Then
    If GetNumberOfTracks > 0& Then
        If getFormattedTime = False Then
            mciSendString "status cdr length track " & track, ReturnString, Len(ReturnString), 0
            GetTrackLength = ReturnString
        Else
            mciSendString "status cdr length track " & Val(track), ReturnString, Len(ReturnString), 0
            GetTrackLength = Mid$(ReturnString, 1, 8)
        End If
    Else
        GetTrackLength = "no_tracks"
    End If
    
    Else
        GetTrackLength = "device_not_open"
End If
End Function

Public Function SetCurrentTime(TimePoint As String) As Boolean
If Len(TimePoint) <> 11 Then SetCurrentTime = False
If IsPlaying Then SetCurrentTime = False: Exit Function
If Not CBool(mciSendString("seek cdr to " + TimePoint, vbNullString, 0, 0)) Then
    SetCurrentTime = True
    Else
    SetCurrentTime = False
End If
End Function

Public Function SetCurrentTrack(TrackNumber As Long) As Boolean
If (TrackNumber <= CLng(GetNumberOfTracks)) And (TrackNumber > 0&) Then
    If Not CBool(mciSendString("seek cdr to " & TrackNumber, vbNullString, 0, 0)) Then
        SetCurrentTrack = True
        Else
        SetCurrentTrack = False
    End If
    Else
    SetCurrentTrack = False
End If
End Function

Public Function IsPlaying() As Boolean
    IsPlaying = False
    If mciOpen Then IsPlaying = CBool(InStr(1, GetCDStatus, "playing"))
End Function

Public Function IsStopped() As Boolean
    IsStopped = False
    If mciOpen Then IsStopped = CBool(InStr(1, GetCDStatus, "stopped"))
End Function

Public Function ShutCD() As Boolean
    ShutCD = False
    If mciOpen Then
        mciSendString "set cdr door closed", vbNullString, 0, 0
        ShutCD = True
    End If
End Function

Public Function PauseCD() As Boolean
    If IsPlaying Then
        If PauseCD = False Then
            mciSendString "pause cdr", vbNullString, 0, 0
            PauseCD = True
    ElseIf PauseCD Then
        PauseCD = False
        mciSendString "play cdr", vbNullString, 0, 0
    End If
    End If
End Function

Public Function StopCD() As Boolean
    StopCD = Not CBool(mciSendString("stop cdr wait", vbNullString, 0, 0))
    SetCurrentTrack (1)
End Function

Public Function PlayCD() As Boolean
    PlayCD = Not CBool(mciSendString("play cdr", vbNullString, 0, 0))
End Function

Public Function CDinDrive(strDriveSet As String) As Long
    VName = String$(255, Chr$(0))
    CDinDrive = GetVolumeInformation(strDriveSet, VName, 255, 0, 0, 0, 0, 255)
    VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
End Function

Public Function CDType(strDriveSet As String) As String
    If DoesFileExist(strDriveSet & "\Track01.cda") = 1 Then
        CDType = "CD Audio"
    ElseIf DoesFileExist(strDriveSet & "\VCD\INFO.VCD") = 1 Then
        CDType = "VCD"
    ElseIf DoesFileExist(strDriveSet & "\VIDEO_TS\VIDEO_TS.ifo") = 1 Then
        CDType = "DVD Video"
    ElseIf DoesFileExist(strDriveSet & "\AUDIO_TS\AUDIO_TS.ifo") = 1 And DoesFileExist(DriveSet & "\VIDEO_TS\VIDEO_TS.ifo") = 1 Then
        CDType = "DVD Audio/Video"
    Else
        CDType = "Unknow or No CD/DVD"
    End If
End Function

Public Function ChangeExt(strFileName As String, strNewExt As String) As Boolean
    On Local Error GoTo ErrorHandler
    ChangeExt = Left$(strFileName, InStrRev(strFileName, ".")) & strNewExt
    ChangeExt = False
Exit Function
ErrorHandler:
        ChangeExt = True
    Err.Clear
End Function

Public Function FindAutorun(InputDrive As String, Optional RunAutoRun As Boolean = False) As Boolean
    Dim InputText As String, ProgName As String
    On Local Error GoTo ErrorHandler
    InputDrive = Mid$(InputDrive, 1, 1) & ":"
    If DoesFileExist(InputDrive & "\autorun.inf") = 1 Then
        Open InputDrive & "\autorun.inf" For Input As 1
            Do Until EOF(1)
                Line Input #1, InputText
                If Mid$(LCase(InputText), 1, 5) = "open=" Then
                        AutorunPath = InputText
                    Exit Do
                End If
            Loop
        Close #1
        ProgName = Mid$(AutorunPath, 6, (Len(AutorunPath)))
        If RunAutoRun Then Shell (InputDrive & "\" & ProgName), vbNormalFocus
        ' .... Or
        'Shell (InputDrive + "\aocsetup.exe /autorun"), vbNormalFocus
    End If
    FindAutorun = True
Exit Function
ErrorHandler:
        FindAutorun = False
    Err.Clear
End Function

Public Function isDriveShared(strDriveSet As String) As Boolean
    Dim bsServer() As Byte, bsShare() As Byte
    Dim Result As Long, buf As Long
    On Local Error GoTo ErrorHandler
    bsServer = "\\" & Environ$("COMPUTERNAME") & Chr(0)
    bsShare = Mid$(strDriveSet, 1, 1) & Chr(0)
    Result = NetShareGetInfo(bsServer(0), bsShare(0), 2, buf)
    If Result = 0 Then
        Result = NetAPIBufferFree(buf)
        isDriveShared = True
    Else
        isDriveShared = False
    End If
    Exit Function
ErrorHandler:
isDriveShared = False
End Function
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
End Sub
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property
Private Function GetCDAudioList() As Boolean
    Dim cddriver As Integer
    On Local Error GoTo ErrorHandler
    lstdrv.Clear
    GetCDAudioList = False
    For cddriver = 65 To 90
        If getdrivetype(Chr$(cddriver) & ":\") = Drive_CDROM Then
            lstdrv.AddItem CStr(Chr$(cddriver))
            GetCDAudioList = True
    End If
Next cddriver
    If lstdrv.ListCount > 0 Then
        GetCDAudioList = True
    ElseIf lstdrv.ListCount < 1 Then
        GetCDAudioList = False
    End If
Exit Function
ErrorHandler:
    GetCDAudioList = False
    Err.Clear
End Function

Public Function FastForward(Speed)
    Dim s As String * 40
    On Local Error Resume Next
    mciSendString "set cdr time format milliseconds", 0, 0, 0
    mciSendString "status cdr position wait", s, Len(s), 0
    If IsPlaying Then
        mciSendString "play cdr from " & CStr(CLng(s) + Speed), 0, 0, 0
    Else
        mciSendString "seek cdr to " & CStr(CLng(s) + Speed), 0, 0, 0
    End If
    mciSendString "set cdr time format tmsf wait", 0, 0, 0
End Function

Public Function FastRewind(Speed)
    Dim s As String * 40
    On Local Error Resume Next
    mciSendString "set cdr time format milliseconds", 0, 0, 0
    mciSendString "status cdr position wait", s, Len(s), 0
    If IsPlaying Then
        mciSendString "play cdr from " & CStr(CLng(s) - Speed), 0, 0, 0
    Else
        mciSendString "seek cdr to " & CStr(CLng(s) - Speed), 0, 0, 0
    End If
    mciSendString "set cdr time format tmsf wait", 0, 0, 0
End Function
