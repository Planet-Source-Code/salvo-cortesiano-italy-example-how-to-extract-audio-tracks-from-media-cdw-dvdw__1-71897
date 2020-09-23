Attribute VB_Name = "modFunction"
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

' .... ProgressBar
Public sIntCount As Long

' .... Disabled Closed
Public readyToClose As Boolean
Public readyToCloseII As Boolean

' .... Init the Classes
Public objGUIDE As New clsGUIDGenerator

' .... Class INI
Public INI As New clsINI

' .... Oter Instance
Private Const SW_RESTORE = 9
Public OpenError As Boolean
Public secInstance As Boolean
Private Declare Function ShowWindowAsync& Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

' ... Init control's XP or Vista
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public m_hMod As Long

' Di conseguenza possiamo risolvere questo problema semplicemente ignorandolo.
' L'unico problema in questo modo è che l'applicazione continua a inviare messaggi al sistema e danno origine
' alla nota finestra che invita a trasmettere le informazioni del Microsoft sul problema:
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOOPENFILEERRORBOX = &H8000&

' ... Exception Handler (Call the Stack)
Public Const MySEH_ERROR = 12345&

Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_EXECUTE_HANDLER = 1

Public Declare Sub DebugBreak Lib "kernel32" ()
Private m_bInIDE As Boolean

' .... Make the change into Registry (=Refresh=)
Private Declare Sub SHChangeNotify Lib "Shell32" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

' .... Convert LongFileName to ShortFileName
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
' .... Convert ShortFileName to LongFileName
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

' .... Constant for SendMessage
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

' .... Open File/Document
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

' .... Play Sound Resource
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_ALIAS = &H10000
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ALIAS_START = 0
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_VALID = &H1F
Private Const SND_NOWAIT = &H2000
Private Const SND_VALIDFLAGS = &H17201F
Private Const SND_RESERVED = &HFF000000
Private Const SND_TYPE_MASK = &H170007

Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)

Private m_snd() As Byte

' .... BrowserForFolders
Private Type BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_BROWSEINCLUDEURLS = 128
Private Const BIF_EDITBOX = 16
Private Const BIF_NEWDIALOGSTYLE = 64
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_STATUSTEXT = 4
Private Const BIF_VALIDATE = 32
Public Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILEDA = 3
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Dim m_StartFolder As String
Dim bValidateFailed As Boolean
Public Property Get InIDE() As Boolean
   Debug.Assert (pIsInIDE)
   InIDE = m_bInIDE
End Property

Private Sub InitControlsCtx()
 On Local Error GoTo ErrorHandler
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Public Function MyExceptionHandler(lpEP As Long) As Long
   Dim lRes As VbMsgBoxResult
   lRes = MsgBox("Exception Handler!" & vbCrLf & "Ignore, Close, or Call the Debugger?", _
   vbAbortRetryIgnore Or vbCritical, App.Title & "Exception Handler")
   Select Case lRes
      Case vbIgnore
         If InIDE Then
            Stop
            MyExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
            On Error GoTo 0
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
         Else
            MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
         End If
       Case vbAbort
         MyExceptionHandler = EXCEPTION_EXECUTE_HANDLER
         Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
       Case vbRetry
         MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
         Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION ???"
   End Select
End Function

Public Property Get pIsInIDE() As Boolean
   m_bInIDE = True
   pIsInIDE = True
End Property

Public Sub Main()
On Local Error GoTo ErrorHandler
    ' .... FLAG_FALSE to Instance
    secInstance = False
    ' .... Verify the Instance
    If App.PrevInstance Then
        OtherInstanceHwnd = fActivateWindowClass("ThunderRT6FormDC", "CD Tracker v1.0.3b")
        If OtherInstanceHwnd = 0 Then
                Call WriteErrorLogs("0", "Problem activating the other instance!", "modFunction {Sub: Main}", True, True)
            secInstance = True
        Else
            secInstance = True
            If Command$ <> "" And StrComp((Right$(Command$, 3)), "scl", vbTextCompare) = 0 _
            Or StrComp((Right$(Command$, 3)), "dcd", vbTextCompare) = 0 Then
                Dim cds As COPYDATASTRUCT, ThWnd As Long, buf(1 To 255) As Byte, a As String
                ThWnd = OtherInstanceHwnd
                a = Command$
                CopyMemory buf(1), ByVal a, Len(a)
                cds.dwData = 3
                cds.cbData = Len(a) + 1
                cds.lpData = VarPtr(buf(1))
                SendMessage OtherInstanceHwnd, WM_COPYDATA, frmMain.hwnd, cds
            End If
        End If
            OpenError = True
        End
    End If
    ' .... Subclass the SO
    SetErrorMode SEM_NOGPFAULTERRORBOX
    ' .... Load the Library
    m_hMod = LoadLibrary("shell32.dll")
    ' .... Init the Controls
    InitControlsCtx
    ' .... Show the Form
    Load frmMain
    frmMain.Show
Exit Sub
ErrorHandler:
    Call WriteErrorLogs(Err.Number, Err.Description, "ModMain {Sub: Main}", True, True)
        Err.Clear
    End
End Sub

Public Sub WriteErrorLogs(strErrNumber As String, strErrDescription As String, Optional strErrSource As String = "Unknow", _
                        Optional visError As Boolean = True, Optional errAppend As Boolean = True)
    Dim FileNum As Integer
    On Error GoTo ErrorHandler
    FileNum = FreeFile
    If Dir$(App.Path & "\_errs.scl") = "" Then
        Open App.Path & "\_errs.scl" For Output As FileNum
        Print #FileNum, Tab(5); "Log Error Generate from [" & App.EXEName & "]..."
        Print #FileNum, Tab(5); Format(Now, "Long Date") & "/" & Time
        Print #FileNum, Tab(5); "----------------------------------------------------------------------------"
        Print #FileNum, Tab(5); ""
        Print #FileNum, Tab(5); ""
        Print #FileNum, Tab(5); "*/___ LOG STARTED..."
        Print #FileNum, Tab(5); ""
        Close FileNum
    End If
    If errAppend Then Open App.Path & "\_errs.scl" For Append As FileNum Else _
                            Open App.Path & "\_errs.scl" For Output As FileNum
        Print #FileNum, Tab(5); Format(Now, "Long Date") & "/" & Time
        Print #FileNum, Tab(5); "Error #" & CStr(strErrNumber)
        Print #FileNum, Tab(5); "Description: " & CStr(strErrDescription)
        Print #FileNum, Tab(5); "Source: " & CStr(strErrSource)
        Print #FileNum, Tab(5); "GUI: " & objGUIDE.CreateGUID("")
        Print #FileNum, Tab(5); ""
        Print #FileNum, Tab(5); ""
        Close FileNum
        If visError Then
            MsgBox "Error #" & CStr(strErrNumber) & "." & vbCrLf & "Description: " & CStr(strErrDescription) _
            & vbCrLf & "Source: " & CStr(strErrSource) & vbCrLf & vbCrLf & "For more info, see the Log file!", vbCritical, App.Title
        End If
    Exit Sub
ErrorHandler:
    MsgBox "Unexpected Error #" & Err.Number & "!" & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Public Function fActivateWindowClass(psClassname As String, App As String) As Long
    Dim hwnd As Long
    hwnd = FindWindow(psClassname, App)
    If hwnd > 0 Then
        ShowWindowAsync hwnd, SW_RESTORE
        SetForegroundWindow hwnd
    End If
    fActivateWindowClass = hwnd
End Function

Public Function ParseCommand(strCommand As String, Optional sInstanza As String = "Unknow") As Boolean
Dim sCommand As String
Dim aFileName As String
On Local Error GoTo ErrorHandler
If Len(strCommand) > 0 Then
    sCommand = UCase(Left$(strCommand, 2))
    aFileName = Right$(strCommand, Len(strCommand) - 3)
        If Len(sCommand) > 0 Then
            Select Case sCommand
            Case "/L" ' .... *.scl Log File
                If OpenFile(GetLongFilename(aFileName)) Then
                    frmMain.lblPath.Caption = GetLongFilename(aFileName)
                    frmMain.picFrame(2).Visible = True
                    frmMain.picFrame(1).Visible = False
                    frmMain.picFrame(0).Visible = False
                    frmMain.TBS.Tabs(3).Selected = True
                End If
            Case "/F" ' .... *.dcd CD List
                If OpenFile(GetLongFilename(aFileName)) Then
                    frmMain.lblPath.Caption = GetLongFilename(aFileName)
                    frmMain.picFrame(2).Visible = True
                    frmMain.picFrame(1).Visible = False
                    frmMain.picFrame(0).Visible = False
                    frmMain.TBS.Tabs(3).Selected = True
                End If
            End Select
        End If
    End If
    ParseCommand = True
Exit Function
ErrorHandler:
    Call WriteErrorLogs(Err.Number, Err.Description, "ModMain {Function: ParseCommand}", True, True)
        ParseCommand = False
    Err.Clear
End Function

Public Function ShellDocument(sDocName As String, Optional Action As String = "Open", Optional _
Parameters As String = vbNullString, Optional Directory As String = vbNullString, Optional WindowState As StartWindowState) As Boolean
    Dim Response
    Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
    Select Case Response
        Case Is < 33
            ShellDocument = False
        Case Else
            ShellDocument = True
    End Select
End Function

Public Function GetLongFilename(ByVal sShortFilename As String) As String
    Dim lRet As Long
    Dim sLongFilename As String
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + 1, " ")
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    If lRet > 0 Then GetLongFilename = Left$(sLongFilename, lRet) Else GetLongFilename = ""
End Function

Public Function GetShortFileName(ByVal FileName As String) As String
    'USAGE:
    'Dim Lblshort, LongName as string
    'Lblshort = GetShortFileName(LongName)
    Dim rc As Long
    Dim ShortPath As String
    Const PATH_LEN& = 164
    ShortPath = String$(PATH_LEN + 1, 0)
    rc = GetShortPathName(FileName, ShortPath, PATH_LEN)
    GetShortFileName = Left$(ShortPath, rc)
End Function

Public Function RemuveExtension(ByVal fExtensionType As String) As Boolean
    On Error GoTo ErrorHeadler
    Call DeleteKey(HKEY_CLASSES_ROOT, "." & fExtensionType)
    Call DeleteKey(HKEY_CLASSES_ROOT, fExtensionType & "file\DefaultIcon")
    Call DeleteKey(HKEY_CLASSES_ROOT, fExtensionType & "file\Shell\Open\Command")
    Call DeleteKey(HKEY_CLASSES_ROOT, fExtensionType & "file\Shell\Open")
    Call DeleteKey(HKEY_CLASSES_ROOT, fExtensionType & "file\Shell\Open")
    Call DeleteKey(HKEY_CLASSES_ROOT, fExtensionType & "file\Shell")
    Call DeleteKey(HKEY_CLASSES_ROOT, fExtensionType & "file")
    ' .... Change and Refresh the Registry
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    RemuveExtension = True
Exit Function
ErrorHeadler:
    RemuveExtension = False
        Call WriteErrorLogs(Err.Number, Err.Description, "ModMain {Function: RemuveExtension}", True, True)
    Err.Clear
End Function

Public Function AssociateExtension(ByVal fExtensionType As String, ByVal TypeDescriptionFile _
                As String, ByVal ContentType As String, ByVal iconFile As Integer, Optional ByVal _
                comVar As String, Optional useNotepadToEdit As Boolean = False) As Boolean
    Dim lRetVal As Long
    Dim hKey As Long
    
    On Error GoTo ErrorHeadler
        
    Call SaveString(HKEY_CLASSES_ROOT, "." & fExtensionType, "", fExtensionType & "file")
    Call SaveString(HKEY_CLASSES_ROOT, "." & fExtensionType, "Content Type", ContentType)
    Call SaveString(HKEY_CLASSES_ROOT, fExtensionType & "file", "", TypeDescriptionFile)
    Call SaveDword(HKEY_CLASSES_ROOT, fExtensionType & "file", "EditFlags", "0000")
    If Dir$(App.Path & "\sc.dll") <> "" Then Call SaveString(HKEY_CLASSES_ROOT, fExtensionType & "file\DefaultIcon", "", _
    App.Path & "\sc.dll," & iconFile) Else Call SaveString(HKEY_CLASSES_ROOT, fExtensionType & "file\DefaultIcon", "", _
    App.Path & "\" & App.EXEName & ".exe,2")
    Call SaveString(HKEY_CLASSES_ROOT, fExtensionType & "file\Shell", "", "")
    Call SaveString(HKEY_CLASSES_ROOT, fExtensionType & "file\Shell\Open", "", "&Open with " & App.EXEName)
    Call SaveString(HKEY_CLASSES_ROOT, fExtensionType & "file\Shell\Open\Command", "", App.Path & "\" & App.EXEName & ".exe" & comVar & " %1")
    If useNotepadToEdit = True Then
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.EXEName & "\shell\edit\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, "notepad.exe %1")
        RegCloseKey (hKey)
    End If
    ' .... Change and Refresh the Registry
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    AssociateExtension = True
Exit Function
ErrorHeadler:
    AssociateExtension = False
        Call WriteErrorLogs(Err.Number, Err.Description, "ModMain {Function: AssociateExtension}", True, True)
    Err.Clear
End Function

Public Sub EndPlaySound()
    On Error Resume Next
    sndPlaySound ByVal vbNullString, 0&
End Sub

Public Function PlaySoundResource(ByVal SndID As Long) As Long
   Const flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
   'USAGE: Call PlaySoundResource(101)
   On Error GoTo ErrorHandler
   DoEvents
   m_snd = LoadResData(SndID, "WAVE")
   PlaySoundResource = PlaySoundData(m_snd(0), 0, flags)
Exit Function
ErrorHandler:
    Err.Clear
End Function

Public Sub DelayTime(ByVal Second As Long, Optional ByVal Refresh As Boolean = True)
    On Error Resume Next
    Dim Start As Date
    Start = Now
    Do
    If Refresh Then DoEvents
    Loop Until DateDiff("s", Start, Now) >= Second
End Sub

Public Function OpenFile(strFileName As String, Optional strAdd As Boolean = False) As Boolean
    Dim FF As Variant: Dim TempText As String: Dim strText As String: Dim iret As Long
    On Error GoTo ErrorHandler
    FF = FreeFile
    Open strFileName For Input As #FF
    TempText = StrConv(InputB(LOF(FF), FF), vbUnicode)
    DoEvents
    If strAdd Then strText = frmMain.txtTextLog.Text
    iret = SendMessage(frmMain.txtTextLog.hwnd, WM_SETTEXT, 0&, ByVal TempText)
    iret = SendMessage(frmMain.txtTextLog.hwnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
    Close #FF
     If strAdd Then
        strText = strText & vbCrLf & frmMain.txtTextLog.Text
        frmMain.txtTextLog = strText
    End If
    TempText = Empty: strText = Empty
    OpenFile = True
    Exit Function
ErrorHandler:
    OpenFile = False
        Call WriteErrorLogs(Err.Number, Err.Description, "ModMain {Function: OpenFile}", True, True)
    Err.Clear
End Function

Public Function MakeDirectory(szDirectory As String) As Boolean
Dim strFolder As String
Dim szRslt As String
On Error GoTo IllegalFolderName
If Right$(szDirectory, 1) <> "\" Then szDirectory = szDirectory & "\"
strFolder = szDirectory
szRslt = Dir(strFolder, 63)
While szRslt = ""
    DoEvents
    szRslt = Dir(strFolder, 63)
    strFolder = Left$(strFolder, Len(strFolder) - 1)
    If strFolder = "" Then GoTo IllegalFolderName
Wend
If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
While strFolder <> szDirectory
    strFolder = Left$(szDirectory, Len(strFolder) + 1)
    If Right$(strFolder, 1) = "\" Then MkDir strFolder
Wend
MakeDirectory = True
Exit Function
IllegalFolderName:
        MakeDirectory = False
    Err.Clear
End Function

Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    On Error Resume Next
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
    Select Case uMsg
        Case BFFM_INITIALIZED
            SendMessageA hwnd, BFFM_SETSELECTION, 1, m_StartFolder
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                SendMessageA hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer
            End If
        Case BFFM_VALIDATEFAILEDA
            bValidateFailed = True
    End Select
    BrowseCallbackProc = 0
End Function

Public Function BrowseForFolder(ByVal hwndOwner As Long, ByVal Prompt As String, Optional ByVal StartFolder) As String
    Dim lNull As Long
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    On Local Error Resume Next
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = Prompt
        .ulFlags = BIF_BROWSEINCLUDEURLS Or BIF_NEWDIALOGSTYLE Or BIF_EDITBOX Or BIF_VALIDATE Or BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
        If Not IsMissing(StartFolder) Then
            m_StartFolder = StartFolder
            If Right$(m_StartFolder, 1) <> Chr$(0) Then m_StartFolder = m_StartFolder & Chr$(0)
            .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
        End If
    End With
    bValidateFailed = False
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList And Not bValidateFailed Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        CoTaskMemFree lpIDList
        lNull = InStr(sPath, vbNullChar)
        If lNull Then
            sPath = Left$(sPath, lNull - 1)
        End If
    End If
    BrowseForFolder = sPath
End Function

Private Function GetAddressofFunction(Add As Long) As Long
    GetAddressofFunction = Add
End Function
