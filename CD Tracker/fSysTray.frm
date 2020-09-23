VERSION 5.00
Begin VB.Form frmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   4365
   ClientLeft      =   5565
   ClientTop       =   2925
   ClientWidth     =   2640
   ControlBox      =   0   'False
   Icon            =   "fSysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optIcon 
      Caption         =   "&ToDo"
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   3675
      Width           =   1515
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&CD List"
      Height          =   375
      Index           =   4
      Left            =   855
      TabIndex        =   4
      Top             =   2925
      Width           =   1515
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&Audio"
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   2250
      Width           =   1515
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&Program"
      Height          =   375
      Index           =   0
      Left            =   990
      TabIndex        =   2
      Top             =   165
      Value           =   -1  'True
      Width           =   990
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&Note"
      Height          =   375
      Index           =   1
      Left            =   990
      TabIndex        =   1
      Top             =   765
      Width           =   1485
   End
   Begin VB.OptionButton optIcon 
      Caption         =   "&Log"
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   0
      Top             =   1365
      Width           =   1515
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Index           =   5
      Left            =   90
      Picture         =   "fSysTray.frx":1CFA
      Top             =   3570
      Width           =   720
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Index           =   4
      Left            =   75
      Picture         =   "fSysTray.frx":39F4
      Top             =   2745
      Width           =   720
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   180
      Picture         =   "fSysTray.frx":5DC6
      Top             =   2190
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Index           =   0
      Left            =   105
      Picture         =   "fSysTray.frx":72E0
      Top             =   30
      Width           =   720
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Index           =   1
      Left            =   105
      Picture         =   "fSysTray.frx":96B2
      Top             =   645
      Width           =   720
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Index           =   2
      Left            =   105
      Picture         =   "fSysTray.frx":B3AC
      Top             =   1350
      Width           =   720
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Begin VB.Menu mnuSysTray 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAA) As Long
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3

Private Type NOTIFYICONDATAA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeOutOrVersion As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
   guidItem As Long
End Type
Private Type NOTIFYICONDATAW
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip(0 To 255) As Byte
   dwState As Long
   dwStateMask As Long
   szInfo(0 To 511) As Byte
   uTimeOutOrVersion As Long
   szInfoTitle(0 To 127) As Byte
   dwInfoFlags As Long
   guidItem As Long
End Type


Private nfIconDataA As NOTIFYICONDATAA
Private nfIconDataW As NOTIFYICONDATAW

Private Const NOTIFYICONDATAA_V1_SIZE_A = 88
Private Const NOTIFYICONDATAA_V1_SIZE_U = 152
Private Const NOTIFYICONDATAA_V2_SIZE_A = 488
Private Const NOTIFYICONDATAA_V2_SIZE_U = 936

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const WM_USER = &H400

Private Const NIN_SELECT = WM_USER
Private Const NINF_KEY = &H1
Private Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Private Declare Function GetVersion Lib "kernel32" () As Long

Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Public Event BalloonShow()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()

Public Enum EBalloonIconTypes
   NIIF_NONE = 0
   NIIF_INFO = 1
   NIIF_WARNING = 2
   NIIF_ERROR = 3
   NIIF_NOSOUND = &H10
End Enum

Private m_bAddedMenuItem As Boolean
Private m_iDefaultIndex As Long

Private m_bUseUnicode As Boolean
Private m_bSupportsNewVersion As Boolean

Public Sub ShowBalloonTip( _
      ByVal sMessage As String, _
      Optional ByVal sTitle As String, _
      Optional ByVal eIcon As EBalloonIconTypes, _
      Optional ByVal lTimeOutMs = 30000 _
   )
Dim lR As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         stringToArray sMessage, nfIconDataW.szInfo, 512
         stringToArray sTitle, nfIconDataW.szInfoTitle, 128
         nfIconDataW.uTimeOutOrVersion = lTimeOutMs
         nfIconDataW.dwInfoFlags = eIcon
         nfIconDataW.uFlags = NIF_INFO
         lR = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
      Else
         nfIconDataA.szInfo = sMessage
         nfIconDataA.szInfoTitle = sTitle
         nfIconDataA.uTimeOutOrVersion = lTimeOutMs
         nfIconDataA.dwInfoFlags = eIcon
         nfIconDataA.uFlags = NIF_INFO
         lR = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
      End If
   Else
      ' can't do it, fail silently.
   End If
End Sub

Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long
    sTip = nfIconDataA.szTip
    iPos = InStr(sTip, Chr$(0))
    If (iPos <> 0) Then
        sTip = Left$(sTip, iPos - 1)
    End If
    ToolTip = sTip
End Property
Public Property Let ToolTip(ByVal sTip As String)
   If (m_bUseUnicode) Then
      stringToArray sTip, nfIconDataW.szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
      nfIconDataW.uFlags = NIF_TIP
      Shell_NotifyIconW NIM_MODIFY, nfIconDataW
   Else
      If (sTip & Chr$(0) <> nfIconDataA.szTip) Then
         nfIconDataA.szTip = sTip & Chr$(0)
         nfIconDataA.uFlags = NIF_TIP
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Public Property Get IconHandle() As Long
    IconHandle = nfIconDataA.hIcon
End Property
Public Property Let IconHandle(ByVal hIcon As Long)
   If (m_bUseUnicode) Then
      If (hIcon <> nfIconDataW.hIcon) Then
         nfIconDataW.hIcon = hIcon
         nfIconDataW.uFlags = NIF_ICON
         Shell_NotifyIconW NIM_MODIFY, nfIconDataW
      End If
   Else
      If (hIcon <> nfIconDataA.hIcon) Then
         nfIconDataA.hIcon = hIcon
         nfIconDataA.uFlags = NIF_ICON
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Public Function AddMenuItem(ByVal sCaption As String, Optional ByVal sKey As String = "", Optional ByVal bDefault As Boolean = False) As Long
Dim iIndex As Long
    If Not (m_bAddedMenuItem) Then
        iIndex = 0
        m_bAddedMenuItem = True
    Else
        iIndex = mnuSysTray.UBound + 1
        Load mnuSysTray(iIndex)
    End If
    mnuSysTray(iIndex).Visible = True
    mnuSysTray(iIndex).Tag = sKey
    mnuSysTray(iIndex).Caption = sCaption
    If (bDefault) Then
        m_iDefaultIndex = iIndex
    End If
    AddMenuItem = iIndex
End Function

Private Function ValidIndex(ByVal lIndex As Long) As Boolean
    ValidIndex = (lIndex >= mnuSysTray.LBound And lIndex <= mnuSysTray.UBound)
End Function

Public Sub EnableMenuItem(ByVal lIndex As Long, ByVal bState As Boolean)
    If (ValidIndex(lIndex)) Then
        mnuSysTray(lIndex).Enabled = bState
    End If
End Sub

Public Function RemoveMenuItem(ByVal iIndex As Long) As Long
Dim i As Long
   If ValidIndex(iIndex) Then
      If (iIndex = 0) Then
         mnuSysTray(0).Caption = ""
      Else
         ' remove the item:
         For i = iIndex + 1 To mnuSysTray.UBound
            mnuSysTray(iIndex - 1).Caption = mnuSysTray(iIndex).Caption
            mnuSysTray(iIndex - 1).Tag = mnuSysTray(iIndex).Tag
         Next i
         Unload mnuSysTray(mnuSysTray.UBound)
      End If
   End If
End Function

Public Property Get DefaultMenuIndex() As Long
   DefaultMenuIndex = m_iDefaultIndex
End Property

Public Property Let DefaultMenuIndex(ByVal lIndex As Long)
   If (ValidIndex(lIndex)) Then
      m_iDefaultIndex = lIndex
   Else
      m_iDefaultIndex = 0
   End If
End Property

Public Function ShowMenu()
   SetForegroundWindow Me.hwnd
   If (m_iDefaultIndex > -1) Then
      Me.PopupMenu mnuPopup, 0, , , mnuSysTray(m_iDefaultIndex)
   Else
      Me.PopupMenu mnuPopup, 0
   End If
End Function

Private Sub Form_Load()
   Dim lMajor As Long
   Dim lMinor As Long
   Dim bIsNt As Long
   On Local Error GoTo ErrorHandler
   GetWindowsVersion lMajor, lMinor, , , bIsNt
   If (bIsNt) Then
      m_bUseUnicode = True
      If (lMajor >= 5) Then
         m_bSupportsNewVersion = True
      End If
   ElseIf (lMajor = 4) And (lMinor = 90) Then
      m_bSupportsNewVersion = True
   End If
   Dim lR As Long
   If (m_bUseUnicode) Then
      With nfIconDataW
         .hwnd = Me.hwnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.Handle
         stringToArray App.FileDescription, .szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconW(NIM_ADD, nfIconDataW)
      If (m_bSupportsNewVersion) Then
         Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
      End If
   Else
      With nfIconDataA
         .hwnd = Me.hwnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.Handle
         .szTip = App.FileDescription & Chr$(0)
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconA(NIM_ADD, nfIconDataA)
      If (m_bSupportsNewVersion) Then
         lR = Shell_NotifyIconA(NIM_SETVERSION, nfIconDataA)
      End If
   End If
Exit Sub
ErrorHandler:
        Call WriteErrorLogs(Err.Number, Err.Description, "frmSysTray {Sub: Form_Load}", True, True)
    Err.Clear
End Sub

Private Sub stringToArray( _
      ByVal sString As String, _
      bArray() As Byte, _
      ByVal lMaxSize As Long _
   )
Dim b() As Byte
Dim i As Long
Dim j As Long
   If Len(sString) > 0 Then
      b = sString
      For i = LBound(b) To UBound(b)
         bArray(i) = b(i)
         If (i = (lMaxSize - 2)) Then
            Exit For
         End If
      Next i
      For j = i To lMaxSize - 1
         bArray(j) = 0
      Next j
   End If
End Sub
Private Function unicodeSize(ByVal lSize As Long) As Long
   If (m_bUseUnicode) Then
      unicodeSize = lSize * 2
   Else
      unicodeSize = lSize
   End If
End Function

Private Property Get nfStructureSize() As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_A
      End If
   Else
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_A
      End If
   End If
End Property

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lX As Long
   lX = ScaleX(x, Me.ScaleMode, vbPixels)
   Select Case lX
   Case WM_MOUSEMOVE
      RaiseEvent SysTrayMouseMove
   Case WM_LBUTTONUP
      RaiseEvent SysTrayMouseDown(vbLeftButton)
   Case WM_LBUTTONUP
      RaiseEvent SysTrayMouseUp(vbLeftButton)
   Case WM_LBUTTONDBLCLK
      RaiseEvent SysTrayDoubleClick(vbLeftButton)
   Case WM_RBUTTONDOWN
      RaiseEvent SysTrayMouseDown(vbRightButton)
   Case WM_RBUTTONUP
      RaiseEvent SysTrayMouseUp(vbRightButton)
   Case WM_RBUTTONDBLCLK
      RaiseEvent SysTrayDoubleClick(vbRightButton)
   Case NIN_BALLOONSHOW
      RaiseEvent BalloonShow
   Case NIN_BALLOONHIDE
      RaiseEvent BalloonHide
   Case NIN_BALLOONTIMEOUT
      RaiseEvent BalloonTimeOut
   Case NIN_BALLOONUSERCLICK
      RaiseEvent BalloonClicked
   End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (m_bUseUnicode) Then Shell_NotifyIconW NIM_DELETE, nfIconDataW _
   Else Shell_NotifyIconA NIM_DELETE, nfIconDataA
End Sub

Private Sub mnuPopup_Click()

End Sub

Private Sub mnuSysTray_Click(Index As Integer)
   RaiseEvent MenuClick(Index, mnuSysTray(Index).Tag)
End Sub

Private Sub GetWindowsVersion( _
      Optional ByRef lMajor = 0, _
      Optional ByRef lMinor = 0, _
      Optional ByRef lRevision = 0, _
      Optional ByRef lBuildNumber = 0, _
      Optional ByRef bIsNt = False _
   )
Dim lR As Long
   lR = GetVersion()
   lBuildNumber = (lR And &H7F000000) \ &H1000000
   If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
   lRevision = (lR And &HFF0000) \ &H10000
   lMinor = (lR And &HFF00&) \ &H100
   lMajor = (lR And &HFF)
   bIsNt = ((lR And &H80000000) = 0)
End Sub


