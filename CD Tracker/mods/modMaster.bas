Attribute VB_Name = "modMaster"
Option Explicit
Option Compare Text
Option Base 0

Public Enum Tipo_CmdWrite
    cw_None = 0
    cw_CreaISO = 1
    cw_MasterizzaISO = 2
    cw_CreaISO_MasterizzaISO = 3
    cw_MasterizzaSenzaISO = 4
    cw_CalcolaDimensioneItems = 5
    cw_ImpostaMasterizzatore = 6
    cw_ImpostaNomeISO = 7
    cw_InizializzaSetMaster = 8
    cw_AddItem = 9
    cw_AddNoItem = 10
    cw_RemoveItem = 11
    cw_RemoveNoItem = 12
    cw_CalcolaDimensioneISO = 13
    cw_CancellaCDRiscrivibile = 14
    cw_ScansioneCartellaFlat = 15
    cw_ScansioneCartellaDepth = 16
    cw_GetAbsolutePath = 17
    cw_CalcolaDimensioneCartella = 18
    cw_GetRelativePath = 19
    cw_UnitaConPiuSpazio = 20
    cw_SpazioLiberoSuUnita = 21
    cw_AddAutoRun_Inf = 22
    cw_ReportISOFileList = 23
    cw_ReportItemsFileList = 24
    cw_VerificaCD = 25
    cw_GetRelativeFullPath = 26
    cw_CancellaImmagineISO = 27
    cw_AzzeraItems = 28
    cw_Open_CD_Door = 29
    cw_Close_CD_Door = 30
    cw_GetIDDevices = 31
    cw_FormattaDimensione = 32
End Enum

Public Enum Tipo_CD_Error
    cderr_NessunErrore = 0
    cderr_FallitaCreazioneISO = 1
    cderr_FallitaMasterizzazioneISO = 2
    cderr_FallitaMasterizzazioneSenzaISO = 3
    cderr_MasterizzatoreNonIdentificato = 4
    cderr_InterrottoDaUtente = 5
    cderr_ErroreScansionePercorso = 6
    cderr_ItemNonValido = 7
    cderr_ItemNonPresente = 8
    cderr_FallitaCancellazioneCD = 9
    cderr_ArgomentoNonValido = 10
    cderr_NonTrovato_CdRecord_exe = 11
    cderr_NonTrovato_mkisofs_exe = 12
    cderr_ElementoGiaPresente = 13
    cderr_NonTrovato_IsoInfo_exe = 14
    cderr_FallitoReportList = 15
    cderr_Warning = 16
    cderr_ErroreCreandoSottoProcesso = 17
    cderr_NonTrovatoGo_PIF = 18
End Enum

Public Type Tipo_ISOImage
    Id_Application As String * 256
    Id_Publisher As String * 128
    Id_Preparer As String * 128
    Id_VolumeIdentifier As String * 32
    Id_VolumeSetName As String * 278
    Id_Abstract As String * 37
    Id_Bibliographic As String * 37
    Id_Copyright As String * 37
    TotItems As Long
    VetItems() As String
    TotNoItems As Long
    VetNoItems() As String
    PathIsoImage As String
    SizeIsoImage As Double
    SizeItems As Double
    SizeNoItems As Double
    AddRelativePath As Boolean
    ExtraOptCmdLine As String
End Type

Public Type Tipo_MasterInfo
    LetteraUnita As String
    LetteraUnitaInput As String
    SCSI_Bus As Long
    SCSI_Id As Long
    SCSI_Lun As Long
    Speed As Long
    LastSpeed As Long
    DontWrite As Boolean
    EspelliCD As Boolean
    ExtraOptCmdLine As String
    IDDevice As String
    OverBurn As Boolean
    BurnProof As Boolean
End Type

Public Type Tipo_RefList
    Nome As String
    Dimensione As Long
    Data As Date
End Type

Public Type Tipo_Globali
    VetFiles() As String
    TotFiles As Long
    VetFolders() As String
    TotFolders As Long
    VetRefList() As Tipo_RefList
    TotRefList As Long
End Type

Public Type Tipo_Inferfaccia
    LabelInfo As Label
    ProgrBar As ProgressBar
    ControllaSeStop As Boolean
    TestStop As Boolean
    TestNascondi As Boolean
End Type

Public Type Tipo_OutPut
    LastGeneralError As String
    LastSpecificError As String
    LastScreenLog As String
    LastCommand As Tipo_CmdWrite
    ErrorCode As Tipo_CD_Error
    ResultValue As Variant
End Type

Public Type Tipo_SetMaster
    Interfaccia As Tipo_Inferfaccia
    ISOImage As Tipo_ISOImage
    Response As Tipo_OutPut
    CdWriter As Tipo_MasterInfo
    GlobArray As Tipo_Globali
End Type

Private Enum Tipo_Redirect
    tr_StdOut = 1
    tr_StdErr = 2
End Enum

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type Tipo_InfoUtility
    GeneralError As String
    MexOperazioneInCorso As String
    CodiceGenerale As Tipo_CD_Error
    TestLabel As Boolean
    TestProgr As Boolean
    LineaShell As String
    UltimoComando As Tipo_CmdWrite
    TestCdRecord As Boolean
    TestReport As Boolean
    LogOut As String
    BufLinea As String
    TotMB As Double
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Enum Tipo_GenereUnita
    tu_Sconosciuta = 0
    tu_Removibile = 1
    tu_Fissa = 2
    tu_Rete = 3
    tu_CD_ROM = 4
    tu_Disco_Ram = 5
End Enum

Private Type Tipo_Drive
    Lettera As String * 1
    Tipo As Tipo_GenereUnita
    Pronta As Boolean
    Seriale As Long
    Nome As String * 64
    SpazioTotale As Double
    SpazioLibero As Double
    NomeCondiviso As String
End Type

Private Type OVERLAPPED

        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type Tipo_InfoUnita
    Lettera As String * 1
    Tipo As Long
    Pronta As Boolean
    Seriale As Long
    Nome As String * 64
    SpazioTotale As Double
    SpazioLibero As Double
End Type

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const WAIT_FAILED = -1
Private Const WAIT_ABANDONED = &H80
Private Const WAIT_TIMEOUT = &H102
Private Const WAIT_STATUS_0 = 0

Private Const SW_SHOW = 5
Private Const SW_HIDE = 0
Private Const ERROR_SUCCESS = 0&
Private Const HKEY_LOCAL_MACHINE = &H80000002
Const KEY_QUERY_VALUE = &H1
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Const REALTIME_PRIORITY_CLASS = &H100
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef pSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const INVALID_HANDLE_VALUE = -1
Private Const STD_OUTPUT_HANDLE = -11&
Private Const PROCESS_TERMINATE = 1

Private Const MCI_CLOSE = &H804
Private Const MCI_OPEN = &H803
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_OPEN_SHAREABLE = &H100&
Private Const MCI_OPEN_TYPE = &H2000&
Private Const MCI_SET = &H80D
Private Const MCI_SET_DOOR_OPEN = &H100&
Private Const MCI_SET_DOOR_CLOSED = &H200&
Private Const MMSYSERR_NOERROR = 0

Private Type MCI_OPEN_PARMS
    dwCallback As Long
    wDeviceID As Long
    lpstrDeviceType As String
    lpstrElementName As String
    lpstrAlias As String
End Type

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Any, ByVal lpThreadAttributes As Any, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long

Private Const HIGH_PRIORITY_CLASS = &H80

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (ByRef phReadPipe As Long, ByRef phWritePipe As Long, ByRef lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal NSize As Long) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const STD_ERROR_HANDLE = -12&
Private Const STD_INPUT_HANDLE = -10&
Private Const DUPLICATE_SAME_ACCESS = &H2

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long

Private Const STARTF_USESHOWWINDOW = &H1

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private TotFiles As Long
Private TotFolders As Long
Private VetFiles() As String
Private VetFolders() As String
Public SetMaster As Tipo_SetMaster
Private Function EseguiRedirect( _
                NomeExe As String, _
                LineaComandi As String, _
                Optional TotElementi As Long = 0) As Boolean

Dim NomeOut As String, TestLabel As Boolean, TestProgr As Boolean
Dim NomeFine As String
Dim NF As Long, Testo As String
Dim proc As PROCESS_INFORMATION
Dim StartInf As STARTUPINFO
Dim saAttr As SECURITY_ATTRIBUTES
Dim hSaveStdOutput As Long, hSaveStdError As Long
Dim hChildStdOutputRd As Long, hChildStdErrorRd As Long
Dim hChildStdOutputWr As Long, hChildStdErrorWr As Long
Dim HandleProcess As Long
Dim hInputFile As Long
Dim dwRead As Long, dwWritten As Long
Dim chBuf As String * 256
Dim RetVal As Variant
Dim BufLinea As String
Dim VetParti() As String
Dim NParti  As Long, TestIgnora As Boolean
Dim ValRedirect As Long, GeneralError As String
Dim VetBytes(256) As Byte, j As Long
Dim NBytes As Long, i As Long
Dim ValCar As Byte, OldLinea As String
Dim TestCdRecord As Boolean, StrNumero As String
Dim Parziale As Double, Totale As Double, OldIgnora As Boolean
Dim LastPerc As Long, CodiceGenerale As Tipo_CD_Error, n As Long
Dim TestReport As Boolean, Nome As String
Dim NomeEscl As String, LineaShell As String
Dim InfoU As Tipo_InfoUtility, LineaFinale As String
Dim TestPrimo As Boolean, OldNBytes As Long
Dim NomeDebug As String

Const BUFSIZE = 256

On Error Resume Next

With SetMaster.Response
    .LastGeneralError = ""
    .ErrorCode = cderr_NessunErrore
    .LastSpecificError = ""
    .LastScreenLog = ""
End With
TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False
TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False
SetMaster.Interfaccia.TestStop = False
Totale = 0
TestReport = False
If (SoloNome(NomeExe) = "cdrecord.exe") Then
    TestCdRecord = True
    GeneralError = "Fallita Masterizzazione CD"
    CodiceGenerale = cderr_FallitaMasterizzazioneISO
Else
    TestCdRecord = False
    If (SoloNome(NomeExe) = "mkisofs.exe") Then
        GeneralError = "Fallita creazione immagine ISO"
        CodiceGenerale = cderr_FallitaCreazioneISO
        If (TestLabel = True) Then
            With SetMaster.Interfaccia
                .LabelInfo.Caption = "Creazione immagine ISO in corso ..."
                .LabelInfo.Refresh
            End With
        End If
    Else
        GeneralError = "Fallita creazione report list"
        CodiceGenerale = cderr_FallitoReportList
        Totale = TotElementi
        TestReport = True
        If (TestLabel = True) Then
            With SetMaster.Interfaccia
                .LabelInfo.Caption = "Creazione report list file in corso ..."
                .LabelInfo.Refresh
            End With
        End If
    End If
End If
If (EsisteFile(NomeExe) = False) Then
    With SetMaster.Response
        If (TestCdRecord = True) Then
            .ErrorCode = cderr_NonTrovato_CdRecord_exe
            .LastGeneralError = "Manca l'utility 'CdRecord.exe' nella directory dell'applicazione"
            .LastSpecificError = "Manca l'utility 'CdRecord.exe' nella directory dell'applicazione"
        Else
            .ErrorCode = cderr_NonTrovato_mkisofs_exe
            .LastGeneralError = "Manca l'utility 'mkisofs.exe' nella directory dell'applicazione"
            .LastSpecificError = "Manca l'utility 'mkisofs.exe' nella directory dell'applicazione"
        End If
    End With
    EseguiRedirect = False
    Exit Function
End If
With InfoU
    .LineaShell = TrovaNomeCorto(NomeExe) & " " & LineaComandi
    .TestCdRecord = TestCdRecord
    .TestLabel = TestLabel
    .TestProgr = TestProgr
    .TestReport = TestReport
    .UltimoComando = SetMaster.Response.LastCommand
    .LogOut = ""
    .TotMB = Totale
    .GeneralError = GeneralError
    .CodiceGenerale = CodiceGenerale
End With

Call ImpostaOperazioneInCorso(InfoU)

If (IsWindowsXP() = True) Then
    LineaShell = TrovaNomeCorto(NomeExe) & " " & LineaComandi
    EseguiRedirect = EseguiCon_BAT(LineaShell, InfoU)
    Exit Function
End If

With saAttr
    .nLength = Len(saAttr)
    .bInheritHandle = True
    .lpSecurityDescriptor = 0
End With

hSaveStdOutput = GetStdHandle(STD_OUTPUT_HANDLE)
hSaveStdError = GetStdHandle(STD_ERROR_HANDLE)
If (CreatePipe(hChildStdOutputRd, hChildStdOutputWr, saAttr, 0) = 0) Then
    With SetMaster.Response
        .ErrorCode = CodiceGenerale
        .LastGeneralError = GeneralError
        .LastSpecificError = "Errore in funzione CreatePipe()"
    End With
    EseguiRedirect = False
    Exit Function
End If
If (SetStdHandle(STD_OUTPUT_HANDLE, hChildStdOutputWr) = 0) Then
    With SetMaster.Response
        .ErrorCode = CodiceGenerale
        .LastGeneralError = GeneralError
        .LastSpecificError = "Errore redirezionando Stadard output con SetStdHandle"
    End With
    EseguiRedirect = False
    Exit Function
End If
If (SetStdHandle(STD_ERROR_HANDLE, hChildStdOutputWr) = 0) Then
    With SetMaster.Response
        .ErrorCode = CodiceGenerale
        .LastGeneralError = GeneralError
        .LastSpecificError = "Errore redirezionando Stadard output con SetStdHandle"
    End With
    EseguiRedirect = False
    Exit Function
End If

With StartInf
    .cb = Len(StartInf)
    If (SetMaster.Interfaccia.TestNascondi = True) Then
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
    End If
End With

If (InfoU.TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = InfoU.MexOperazioneInCorso
        .LabelInfo.Refresh
    End With
End If

RetVal = CreateProcess(NomeExe, _
            NomeExe & " " & LineaComandi, _
            vbNullString, _
            vbNullString, _
            True, _
            0, _
            vbNullString, _
            SoloDir(NomeExe), _
            StartInf, _
            proc)

If (RetVal = 0) Then
    With SetMaster.Response
        .LastGeneralError = GeneralError
        .ErrorCode = CodiceGenerale
        .LastSpecificError = "Errore creando sottoprocesso: " & _
            Chr(34) & NomeExe & Chr(34) & vbCrLf & _
            "Con linea comandi: " & Chr(34) & LineaComandi & Chr(34)
    End With
    EseguiRedirect = False
    Exit Function
End If
HandleProcess = OpenProcess( _
            PROCESS_TERMINATE, _
            False, proc.dwProcessID)

Call SetStdHandle(STD_OUTPUT_HANDLE, hSaveStdOutput)
Call SetStdHandle(STD_ERROR_HANDLE, hSaveStdError)
Call CloseHandle(hChildStdOutputWr)

NBytes = 0
Testo = ""
OldLinea = ""
OldIgnora = False
LastPerc = 0
Parziale = 0
TestPrimo = True

Do
    TestIgnora = False
    If (ReadFile(hChildStdOutputRd, _
                    ByVal chBuf, 1, _
                    dwRead, ByVal 0&) = 0) Then Exit Do
    ValCar = Asc(chBuf)
    LineaFinale = ""
    Select Case ValCar
        Case 13, 10
            LineaFinale = ""
            For i = 0 To NBytes - 1
                LineaFinale = LineaFinale & Chr(VetBytes(i))
            Next i
            NBytes = 0
            If (TestReport = True) Then
                Parziale = Parziale + 1
            End If
            
        Case 8
            NBytes = NBytes - 1
            If (NBytes < 0) Then NBytes = 0
        Case 9
            For i = 0 To 3
                VetBytes(NBytes + i) = 32
            Next i
            NBytes = NBytes + 4
        Case Else
            VetBytes(NBytes) = ValCar
            NBytes = NBytes + 1
    End Select
    If (LineaFinale <> "" Or TestPrimo = True) Then
        InfoU.BufLinea = LineaFinale
    
        Call ScansioneLinea(InfoU)
    End If
    TestPrimo = False
 
    If (SetMaster.Interfaccia.ControllaSeStop = True) Then
        DoEvents
        If (SetMaster.Interfaccia.TestStop = True) Then
            Call TerminateProcess(HandleProcess, -1)
            Call AzzeraInterfaccia
            With SetMaster.Response
                .ErrorCode = cderr_InterrottoDaUtente
                .LastScreenLog = InfoU.LogOut
                .LastGeneralError = GeneralError
                .LastSpecificError = "Operazione interrotta dall'utente"
            End With
            EseguiRedirect = False
            Exit Function
        End If
    End If
Loop While (dwRead <> 0)

SetMaster.Response.LastScreenLog = InfoU.LogOut
Call AzzeraInterfaccia

EseguiRedirect = True

End Function

Private Function InitSetMaster(Optional IDDevice As String = "", _
                    Optional Speed As Long = -1) As Boolean

With SetMaster.ISOImage
    .Id_Application = ""
    .Id_Preparer = "VB Master"
    .Id_Publisher = ""
    .Id_VolumeIdentifier = "Nuovo"
    .Id_Abstract = ""
    .Id_Bibliographic = ""
    .Id_Copyright = ""
    .Id_VolumeSetName = ""
    .PathIsoImage = ""
    .SizeIsoImage = 0
    .TotItems = 0
    .TotNoItems = 0
    .AddRelativePath = True
    .ExtraOptCmdLine = ""
End With

With SetMaster.Interfaccia
    .ControllaSeStop = False
    Set .LabelInfo = Nothing
    Set .ProgrBar = Nothing
    .TestNascondi = True
    .TestStop = False
End With

With SetMaster.Response
    .LastGeneralError = ""
    .ErrorCode = cderr_NessunErrore
    .LastSpecificError = ""
    .LastScreenLog = ""
    .ErrorCode = cderr_NessunErrore
    .LastCommand = cw_None
    .ResultValue = 0
End With

With SetMaster.GlobArray
    .TotFiles = 0
    .TotFolders = 0
    .TotRefList = 0
End With

With SetMaster.CdWriter
    .DontWrite = False
    If (Speed <> -1) Then
        .Speed = Speed
    Else
    
        If (.Speed = 0) Then
            If (.LastSpeed <> 0) Then
                .Speed = .LastSpeed
            Else
                .Speed = 2
            End If
        End If
    End If
    
    .EspelliCD = False
    .ExtraOptCmdLine = ""
    InitSetMaster = TrovaMasterizzatore(IDDevice)
End With

End Function

Private Function TrovaMasterizzatore(IDDevice As String) As Boolean

Dim i As Long, LetteraCd As String
Dim VetParti() As String, j As Long, z As Long
Dim StrNumero1 As String, StrNumero2 As String
Dim Testo As String, Car As String

If (IDDevice = "") Then IDDevice = SetMaster.CdWriter.IDDevice

If (IDDevice = "") Then
    With SetMaster.Response
        .ErrorCode = cderr_MasterizzatoreNonIdentificato
        .LastGeneralError = "Unita' corrispondente al masterizzatore non identificabile"
        .LastSpecificError = "Non e' stata fornita la stringa per identificare il dispositivo"
        .LastScreenLog = ""
        .ResultValue = 0
    End With
    TrovaMasterizzatore = False
    Exit Function
End If

VetParti = Split(IDDevice, "'")
If (UBound(VetParti) < 7) Then GoTo ErroreSintassi

For i = 0 To 2
    If (VetParti(i) = "" Or _
            IsNumeric(VetParti(i)) = False) Then GoTo ErroreSintassi
Next i

With SetMaster.CdWriter
    .LetteraUnita = RisolveLetteraCd(VetParti(3))
    .SCSI_Lun = Val(VetParti(2))
    .SCSI_Id = Val(VetParti(1))
    .SCSI_Bus = Val(VetParti(0))
    
End With
TrovaMasterizzatore = True
Exit Function

ErroreSintassi:

With SetMaster.Response
    .ErrorCode = cderr_ArgomentoNonValido
    .LastGeneralError = "Argomento fornito non valido"
    .LastSpecificError = "Sintassi di nome dispositivo non valida"
    .LastScreenLog = ""
    .ResultValue = 0
End With
TrovaMasterizzatore = False

End Function


Private Function PathIsoDefault() As String
Dim Spazio As Double

PathIsoDefault = TrovaUnitaPiuSpazio(Spazio) & "\imagecd.iso"

End Function
Private Function CalcolaDimensioneItems(TestAncheIso As Boolean) As Boolean

Dim TotItems As Long, i As Long
Dim PathItem As String, z As Long
Dim Nome As String, NomeDir As String
Dim VetTuttiItem() As String
Dim VetIsItemNo() As Boolean
Dim j As Long, SizeNow As Double
Dim TestLabel As Boolean, TestProgr As Boolean
Dim DimFileSystem As Long
Dim VetNomi() As String, NNomi As Long

On Error Resume Next

TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False

TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False
Call AzzeraInterfaccia

With SetMaster.ISOImage
    .SizeItems = 0
    .SizeNoItems = 0
    .SizeIsoImage = 0
    
    TotItems = .TotItems + .TotNoItems
    ReDim VetTuttiItem(TotItems)
    ReDim VetIsItemNo(TotItems)
    j = 0
    For i = 0 To .TotItems - 1
        VetIsItemNo(j) = False
        VetTuttiItem(j) = SoloPathAssoluto(.VetItems(i))
        j = j + 1
    Next i
    For i = 0 To .TotNoItems - 1
        Nome = SoloPathAssoluto(.VetNoItems(i))
        VetIsItemNo(j) = True
        VetTuttiItem(j) = Nome
        j = j + 1
    Next i
    TotItems = j
End With

With SetMaster.Interfaccia
    If (TestLabel = True) Then
        .LabelInfo.Caption = "Calcolo dimensione dei file in corso ..."
        .LabelInfo.Refresh
    End If
End With

For i = 0 To TotItems - 1
    With SetMaster.Interfaccia
        If (TestProgr = True) Then
            Call MostraPercentuale(.ProgrBar, _
                                TotItems, i + 1)
        End If
    End With
    With SetMaster.Interfaccia
        If (.ControllaSeStop = True) Then
            DoEvents
            If (.TestStop = True) Then
                With SetMaster.Response
                    .ErrorCode = cderr_InterrottoDaUtente
                    .LastGeneralError = "Operazione interrotta dall'utente"
                End With
                Call AzzeraInterfaccia
                CalcolaDimensioneItems = False
                Exit Function
            End If
        End If
    End With
    PathItem = VetTuttiItem(i)

    With SetMaster.ISOImage
        If (EsisteDir(PathItem) = True) Then
            SizeNow = TrovaDimFolder(PathItem)
            If (VetIsItemNo(i) = True) Then
                .SizeNoItems = .SizeNoItems + SizeNow
            Else
                .SizeItems = .SizeItems + SizeNow
            End If

        Else
            If (EsisteFile(PathItem) = False) Then
                With SetMaster.Response
                    .ErrorCode = cderr_ErroreScansionePercorso
                    .LastGeneralError = "Errore durante calcolo dimensione dei file da inserire/escludere nell'immagine ISO"
                    .LastSpecificError = _
                        "File: " & Chr(34) & PathItem & Chr(34) & " non esiste"
                        
                End With
                Call AzzeraInterfaccia
    
                CalcolaDimensioneItems = False
                Exit Function
            End If
            
            SizeNow = FileLen(PathItem)
            If (VetIsItemNo(i) = True) Then
                .SizeNoItems = .SizeNoItems + SizeNow
            Else
                .SizeItems = .SizeItems + SizeNow
            End If
                
        End If
    End With
    
Next i
Call AzzeraInterfaccia
If (TestAncheIso = True) Then
    Call TrovaDimFileSystem(DimFileSystem)
Else
    DimFileSystem = 0
End If

With SetMaster.ISOImage
    .SizeIsoImage = .SizeItems - .SizeNoItems
    .SizeIsoImage = .SizeIsoImage + DimFileSystem
End With



CalcolaDimensioneItems = True
End Function
Public Function VBMaster(Comando As Tipo_CmdWrite, _
            Optional Argomento1 As String = "", _
            Optional Argomento2 As String = "") As Boolean

Dim i As Long
Dim SizeIso As Double, NomeIso As String, j As Long
Dim NomeFileItem As String, NomeFileNoItem As String, NomeDir
Dim DirCorta As String, InfoDrive As Tipo_Drive, Nome As String
Dim TestFile As Boolean, VetParti() As String, NParti As Long
Dim InfoUnita As Tipo_Drive, NF As Long, PathRel As String
Dim Spazio As Double, Testo As String, MexMask As String, NomeFile As String

With SetMaster.Response
    .ErrorCode = cderr_NessunErrore
    .LastGeneralError = ""
    .LastSpecificError = ""
    .LastScreenLog = ""
    .LastCommand = Comando
    .ResultValue = 0
End With


Select Case Comando
    Case cw_GetIDDevices
                
        VBMaster = TrovaIdDispositivi()
        Exit Function
        
    Case Tipo_CmdWrite.cw_CalcolaDimensioneItems
        
        If (CalcolaDimensioneItems(False) = False) Then
            VBMaster = False
            SetMaster.ISOImage.SizeItems = 0
            SetMaster.ISOImage.SizeNoItems = 0
        Else
            SetMaster.Response.ResultValue = _
                    SetMaster.ISOImage.SizeIsoImage
            VBMaster = True
        End If
        Exit Function
    
    Case Tipo_CmdWrite.cw_CalcolaDimensioneISO
    
        If (CalcolaDimensioneItems(True) = False) Then
            VBMaster = False
            SetMaster.ISOImage.SizeIsoImage = 0
            SetMaster.Response.ResultValue = 0
        Else
            SetMaster.Response.ResultValue = _
                    SetMaster.ISOImage.SizeIsoImage
            
            VBMaster = True
        
        End If
        Exit Function
    
    Case Tipo_CmdWrite.cw_ImpostaMasterizzatore
        
        VBMaster = TrovaMasterizzatore( _
                        SetMaster.CdWriter.IDDevice)
        If (VBMaster = True) Then
            SetMaster.Response.ResultValue = _
                    SetMaster.CdWriter.IDDevice
        End If
        Exit Function
    
    Case Tipo_CmdWrite.cw_ImpostaNomeISO
        
        SetMaster.ISOImage.PathIsoImage = PathIsoDefault()
        SetMaster.Response.ResultValue = _
                    SetMaster.ISOImage.PathIsoImage
        VBMaster = True
        Exit Function
    
    Case Tipo_CmdWrite.cw_InizializzaSetMaster
        i = -1
        If (Argomento2 = "AUTO") Then
            i = 0
        End If
        If (IsNumeric(Argomento2)) Then
            i = Val(Argomento2)
        End If
        
        VBMaster = InitSetMaster(Argomento1, i)
        Exit Function

    Case Tipo_CmdWrite.cw_AddItem
        
        Nome = SoloPathAssoluto(Argomento1)
        PathRel = SoloPathRelativo(Argomento1)
        
        If (Right(Nome, 1) = "\") Then
            With SetMaster.Response
                .ErrorCode = cderr_ErroreScansionePercorso
                .LastGeneralError = "Percorso non valido"
                .LastSpecificError = "Non e' ammesso terminare un percorso col carattere '\' o '/'"
            End With
            VBMaster = False
            Exit Function
        End If
        If (IsMask(Nome) = True) Then
            With SetMaster.Response
                .ErrorCode = cderr_ErroreScansionePercorso
                .LastGeneralError = "Percorso non valido"
                .LastSpecificError = "Non e' ammesso utilizzare caratteri jolly"
            End With
            VBMaster = False
            Exit Function
        End If
        If (InStr(1, Nome, ":", vbBinaryCompare) = 0) Then
            With SetMaster.Response
                .ErrorCode = cderr_ItemNonValido
                .LastGeneralError = "Impossibile inserire elemento in VetItems()"
                .LastSpecificError = "In VetItems() non si possono inserire percorsi privi di lettera di unita'"
            End With
            VBMaster = False
            Exit Function
        End If
        If (EsisteFile(Nome) = False And _
            EsisteDir(Nome) = False) Then
            With SetMaster.Response
            
                .ErrorCode = cderr_ItemNonValido
                .LastGeneralError = _
                    "Impossibile inserire elemento in VetItems()"
                .LastSpecificError = _
                    "Il percorso " & Chr(34) & Nome & Chr(34) & _
                    " non esiste"
            End With
            VBMaster = False
            Exit Function
        End If

        If (EsisteDir(Nome) = True And _
            InStr(1, Argomento1, "=", vbBinaryCompare) > 0 And _
            Right(PathRel, 1) <> "\") Then
            With SetMaster.Response
            
                .ErrorCode = cderr_ItemNonValido
                .LastGeneralError = _
                    "Impossibile inserire elemento in VetItems()"
                .LastSpecificError = _
                    "Il percorso relativo per una cartella deve terminare col carattere '\' o '/'"
            End With
            VBMaster = False
            Exit Function
        End If

        If (SetMaster.ISOImage.AddRelativePath = True And _
            InStr(1, Argomento1, "=", vbBinaryCompare) = 0) Then
            
            TestFile = True
            If (IsMask(Nome) = False) Then
                If (EsisteDir(Nome) = True) Then _
                                        TestFile = False
            End If
            VetParti = Split(Nome, "\", -1, vbBinaryCompare)
            NParti = UBound(VetParti) + 1
            
            If (TestFile = True) Then
                NomeDir = VetParti(NParti - 2)
            Else
                NomeDir = VetParti(NParti - 1)
            End If
            If (InStr(1, NomeDir, ":", vbBinaryCompare) > 0) Then
                NomeDir = "\"
            End If
            If (Right(NomeDir, 1) <> "\") Then
                NomeDir = NomeDir & "\"
            End If
            Argomento1 = NomeDir & "=" & Argomento1
        End If

        With SetMaster.ISOImage
            Nome = UsaBarreUnix(Argomento1)
            For i = 0 To .TotItems - 1
                If (Nome = SoloPathAssoluto( _
                            .VetItems(i))) Then Exit For
            Next i
            
            If (i = .TotItems) Then
                ReDim Preserve .VetItems(.TotItems)
                .VetItems(.TotItems) = Nome
                .TotItems = .TotItems + 1
            End If
            
        End With
        VBMaster = True
        Exit Function

    Case Tipo_CmdWrite.cw_AddNoItem
        If (InStr(1, Argomento1, "=", vbBinaryCompare) > 0) Then
            With SetMaster.Response
                .ErrorCode = cderr_ArgomentoNonValido
                .LastGeneralError = "Percorso non valido"
                .LastSpecificError = "Non e' ammesso inserire un percorso relativo '=' nel vettore VetNoItems()"
            End With
            VBMaster = False
            Exit Function
        End If
        Nome = UsaBarreMsDos(Argomento1)

        If (IsMask(Nome) = True) Then
            With SetMaster.Response
                .ErrorCode = cderr_ErroreScansionePercorso
                .LastGeneralError = "Percorso non valido"
                .LastSpecificError = "Non e' ammesso utilizzare caratteri jolly"
            End With
            VBMaster = False
            Exit Function
        End If

        If (EsisteFile(Nome) = False And _
            EsisteDir(Nome) = False) Then
            With SetMaster.Response
                .ErrorCode = cderr_ItemNonValido
            
                .LastGeneralError = _
                    "Impossibile inserire elemento in VetNoItems()"
                .LastSpecificError = _
                    "Il percorso " & Chr(34) & Nome & Chr(34) & _
                    " non esiste"
            End With
            VBMaster = False
            Exit Function
        End If
        With SetMaster.ISOImage
            For i = 0 To .TotNoItems - 1
                If (Nome = UsaBarreMsDos(.VetNoItems(i))) Then
                    Exit For
                End If
            Next i
            If (i < .TotNoItems) Then
                With SetMaster.Response
                    .ErrorCode = cderr_ElementoGiaPresente
                    .LastGeneralError = "Elemento gia' presente"
                    .LastSpecificError = "Elemento NoItem " & Chr(34) & Nome & Chr(34) & " e' gia' presente in VetNoItems()"
                End With
                VBMaster = False
                Exit Function
            End If
            ReDim Preserve .VetNoItems(.TotNoItems)
            .VetNoItems(.TotNoItems) = Argomento1
            .TotNoItems = .TotNoItems + 1
        End With
        VBMaster = True
        Exit Function

    Case Tipo_CmdWrite.cw_RemoveItem
    
        With SetMaster.ISOImage
            Nome = SoloPathAssoluto(Argomento1)
            For i = 0 To .TotItems - 1
                If (Nome = SoloPathAssoluto(.VetItems(i))) Then Exit For
            Next i
            If (i = .TotItems) Then
                With SetMaster.Response
                    .ErrorCode = cderr_ItemNonPresente
                    .LastGeneralError = "Argomento1 non presente"
                End With
                VBMaster = False
                Exit Function
            End If
            
            j = .TotItems - 1
            .VetItems(i) = .VetItems(j)
            .TotItems = .TotItems - 1
            VBMaster = True
        End With
        Exit Function

    Case Tipo_CmdWrite.cw_RemoveNoItem
        
        With SetMaster.ISOImage
            Nome = SoloPathAssoluto(Argomento1)
            
            For i = 0 To .TotNoItems - 1
                If (Nome = SoloPathAssoluto(.VetNoItems(i))) Then Exit For
            Next i
            If (i = .TotNoItems) Then
                With SetMaster.Response
                    .ErrorCode = cderr_ItemNonPresente
                    .LastGeneralError = "Argomento1 non presente"
                End With
                VBMaster = False
                Exit Function
            End If
            
            j = .TotNoItems - 1
            .VetNoItems(i) = .VetNoItems(j)
            .TotNoItems = .TotNoItems - 1
            VBMaster = True
        End With
        Exit Function
    Case Tipo_CmdWrite.cw_CancellaCDRiscrivibile
        If (Argomento1 <> "all" And _
            Argomento1 <> "fast") Then
                With SetMaster.Response
                    .ErrorCode = cderr_ArgomentoNonValido
                    .LastGeneralError = "Argomento non valido per comando corrente"
                    .LastSpecificError = "Il comando CancellaCDRiscrivibile ammette come operando solo i valori 'ALL' e 'FAST'"
                End With
                VBMaster = False
                Exit Function
        End If
        VBMaster = CancellaCD(Argomento1)
        Exit Function
    Case Tipo_CmdWrite.cw_ScansioneCartellaDepth
        TotFiles = 0
        TotFolders = 0
        If (EsisteDir(Argomento1) = False) Then
            With SetMaster.Response
                .ErrorCode = cderr_ErroreScansionePercorso
                .LastGeneralError = "Percorso non esiste"
                .LastSpecificError = "L'argomento fornito non corrisponde ad una cartella valida"
            End With
            VBMaster = False
            Exit Function
        End If
        Call ScansioneCartelle(Argomento1)
        With SetMaster.GlobArray
            ReDim .VetFiles(TotFiles)
            ReDim .VetFolders(TotFolders)
            For i = 0 To TotFiles - 1
                .VetFiles(i) = VetFiles(i)
            Next i
            .TotFiles = TotFiles
            For i = 0 To TotFolders - 1
                .VetFolders(i) = VetFolders(i)
            Next i
            .TotFolders = TotFolders
        End With
        
        VBMaster = True
        Exit Function
    Case Tipo_CmdWrite.cw_ScansioneCartellaFlat
        If (EsisteDir(Argomento1) = False) Then
            With SetMaster.Response
                .ErrorCode = cderr_ErroreScansionePercorso
                .LastGeneralError = "Percorso non esiste"
                .LastSpecificError = "L'argomento fornito non corrisponde ad una cartella valida"
            End With
            VBMaster = False
            Exit Function
        End If
        If (Argomento2 = "") Then
            MexMask = "*"
        Else
            MexMask = Argomento2
        End If
            
        TotFolders = 0
        VetFiles = TrovaFiles(Argomento1, MexMask, TotFiles)
        With SetMaster.GlobArray
            ReDim .VetFiles(TotFiles)
            For i = 0 To TotFiles - 1
                .VetFiles(i) = VetFiles(i)
            Next i
            .TotFiles = TotFiles
        End With
        VBMaster = True
        Exit Function
    Case Tipo_CmdWrite.cw_CalcolaDimensioneCartella
        If (EsisteDir(Argomento1) = False Or _
                    Trim(Argomento1) = "") Then
            With SetMaster.Response
                .ErrorCode = cderr_ErroreScansionePercorso
                .LastGeneralError = "Percorso non esiste"
                .LastSpecificError = "L'argomento fornito non corrisponde ad una cartella valida"
            End With
            VBMaster = False
            Exit Function
        End If
        
        SetMaster.Response.ResultValue = _
                        TrovaDimFolder(Argomento1, "*")
        VBMaster = True
        Exit Function
        
    Case Tipo_CmdWrite.cw_GetAbsolutePath
        If (Trim(Argomento1) = "") Then
            With SetMaster.Response
                .ErrorCode = cderr_ArgomentoNonValido
                .LastGeneralError = "Argomento non valido per comando GetAbsolutePath"
                .LastSpecificError = "Argomento nullo. E' necessario inserire l'elemento da convertire"
            End With
            VBMaster = False
            Exit Function
        End If
        
        SetMaster.Response.ResultValue = _
                        SoloPathAssoluto(Argomento1)
        VBMaster = True
        Exit Function
        
    Case Tipo_CmdWrite.cw_GetRelativePath
        If (Trim(Argomento1) = "") Then
            With SetMaster.Response
                .ErrorCode = cderr_ArgomentoNonValido
                .LastGeneralError = "Argomento non valido per comando GetRelativePath"
                .LastSpecificError = "Argomento nullo. E' necessario inserire l'elemento da convertire"
            End With
            VBMaster = False
            Exit Function
        End If
        Nome = SoloPathAssoluto(Argomento1)
        
        With SetMaster.ISOImage
            For i = 0 To .TotItems - 1
                If (Nome = SoloPathAssoluto(.VetItems(i))) Then Exit For
            Next i
            If (i = .TotItems) Then
                With SetMaster.Response
                    .ErrorCode = cderr_ItemNonPresente
                    .LastGeneralError = "Elemento non trovato"
                    .LastSpecificError = "L'elememento " & Chr(34) & Nome & Chr(34) & " non esiste in VetItems()"
                End With
                VBMaster = False
                Exit Function
            
            End If
            SetMaster.Response.ResultValue = _
                        SoloPathRelativo(.VetItems(i))
            

        End With
        VBMaster = True
        Exit Function
    Case Tipo_CmdWrite.cw_GetRelativeFullPath
        If (Trim(Argomento1) = "") Then
            With SetMaster.Response
                .ErrorCode = cderr_ArgomentoNonValido
                .LastGeneralError = "Argomento non valido per comando GetRelativePath"
                .LastSpecificError = "Argomento nullo. E' necessario inserire l'elemento da convertire"
            End With
            VBMaster = False
            Exit Function
        End If
        If (InStr(1, Argomento1, "=", vbBinaryCompare) = 0) Then
            Nome = SoloPathAssoluto(Argomento1)
            
            With SetMaster.ISOImage
                For i = 0 To .TotItems - 1
                    If (Nome = SoloPathAssoluto(.VetItems(i))) Then Exit For
                Next i
                If (i = .TotItems) Then
                    With SetMaster.Response
                        .ErrorCode = cderr_ItemNonPresente
                        .LastGeneralError = "Elemento non trovato"
                        .LastSpecificError = "L'elememento " & Chr(34) & Nome & Chr(34) & " non esiste in VetItems()"
                    End With
                    VBMaster = False
                    Exit Function
                
                End If
                Nome = .VetItems(i)
            End With
        Else
            Nome = UsaBarreMsDos(Argomento1)
        End If
        
        SetMaster.Response.ResultValue = _
                    PathRelativoFull(Nome, _
                            SoloPathAssoluto(Nome))

        VBMaster = True
        Exit Function

    Case Tipo_CmdWrite.cw_SpazioLiberoSuUnita
        If (Trim(Argomento1) = "") Then
            With SetMaster.Response
                .ErrorCode = cderr_ArgomentoNonValido
                .LastGeneralError = "Argomento non valido per comando SpazioLiberoSuUnita"
                .LastSpecificError = "Argomento nullo. E' necessario inserire la lettera di unita' da esaminare"
            End With
            VBMaster = False
            Exit Function
        End If

        SetMaster.Response.ResultValue = _
                SpazioLiberoSulDisco(Argomento1)
        VBMaster = True
        Exit Function
    Case Tipo_CmdWrite.cw_UnitaConPiuSpazio
        SetMaster.Response.ResultValue = _
                TrovaUnitaPiuSpazio(Spazio)
        VBMaster = True
        Exit Function
    Case Tipo_CmdWrite.cw_AddAutoRun_Inf
        Nome = SoloPathAssoluto(Argomento1)
        
        If (Argomento1 = "" Or _
            EsisteFile(Nome) = False) Then
            
            With SetMaster.Response
                .ErrorCode = cderr_ArgomentoNonValido
                .LastGeneralError = "Argomento fornito per comando 'AddAutoRun_Inf' non e' valido"
                .LastSpecificError = "Devi inserire il nome del file sorgente completo, dove si trova attualmente"
            End With
            VBMaster = False
            Exit Function
        End If
        NomeFile = PathRelativoFull(Argomento1, Nome)
        If (Left(NomeFile, 1) = "\") Then
            NomeFile = Right(NomeFile, Len(NomeFile) - 1)
        End If
        Testo = "[autorun]" & vbCrLf & "OPEN=" & NomeFile & vbCrLf
        
        Nome = App.Path & "\autorun.inf"
        NF = FreeFile
        Open Nome For Output As #NF
        Print #NF, Testo;
        Close #NF

        VBMaster = VBMaster(cw_AddItem, "/=" & Nome)
        If (VBMaster = False Or Argomento2 = "NotAddFile") Then Exit Function
        VBMaster = VBMaster(cw_AddItem, Argomento1)
        
        Exit Function
        
    Case Tipo_CmdWrite.cw_ReportISOFileList
        VBMaster = CreaListFile(Argomento1, Argomento2)
        Exit Function
        
    Case Tipo_CmdWrite.cw_ReportItemsFileList
        VBMaster = CreaListFileItem(Argomento1)
        Exit Function
    Case Tipo_CmdWrite.cw_VerificaCD
        VBMaster = VerificaCD(Argomento1)
        Exit Function
    Case Tipo_CmdWrite.cw_AzzeraItems
        With SetMaster.ISOImage
            .TotItems = 0
            .TotNoItems = 0
        End With
        VBMaster = True
        Exit Function
    Case Tipo_CmdWrite.cw_CancellaImmagineISO
        Nome = SetMaster.ISOImage.PathIsoImage
        If (EsisteFile(Nome) = True) Then
            Call DeleteFile(Nome)
        End If
        VBMaster = True
        Exit Function
        
    Case Tipo_CmdWrite.cw_Close_CD_Door, _
         Tipo_CmdWrite.cw_Open_CD_Door
         
        If (Argomento1 = "" Or Left(Argomento1, 1) = "X") Then
            With SetMaster.CdWriter
                If (.LetteraUnita <> "") Then
                    Argomento1 = RisolveLetteraCd(Left(.LetteraUnita, 1) & ":")
                End If
            End With
            
        End If
        If (Len(Argomento1) < 2) Then
            With SetMaster.Response
                .LastGeneralError = "Argomento non valido per comando Open/Close_Cd_Door"
                .LastSpecificError = "Il comando Open/Close_CD_Door richiede la lettera di unita' corrispondente all'unita' CD da aprire/chiudere, nel formato a due caratteri come 'D:'"
                .ErrorCode = cderr_ArgomentoNonValido
            End With
            VBMaster = False
            Exit Function
        End If
        
        If (Comando = cw_Close_CD_Door) Then
            VBMaster = CassettoCD(Argomento1, False)
        Else
            VBMaster = CassettoCD(Argomento1, True)
        End If
        Exit Function
    Case Tipo_CmdWrite.cw_FormattaDimensione
        If (Argomento1 = "" Or IsNumeric(Argomento1) = False) Then
            With SetMaster.Response
                .ErrorCode = cderr_ArgomentoNonValido
                .LastGeneralError = "Valore di argomento non valido"
                .LastSpecificError = "Il comando cw_FormattaDimensione richiede un argomento numerico"
            End With
            VBMaster = False
            Exit Function
        End If
        Spazio = Val(Argomento1)
        SetMaster.Response.ResultValue = _
                        FormattaDimensione(Spazio)
        VBMaster = True
        Exit Function
End Select

If (Comando = cw_CreaISO Or _
    Comando = cw_CreaISO_MasterizzaISO) Then

    If (CreaImmagineIso() = False) Then
        VBMaster = False
        Exit Function
    End If
    
End If

If (Comando = cw_CreaISO_MasterizzaISO Or _
    Comando = cw_MasterizzaISO) Then
    
    VBMaster = MasterizzaIso()
    Exit Function
End If

If (Comando = cw_MasterizzaSenzaISO) Then
    VBMaster = MasterizzaSenzaIso()
    Exit Function
End If

VBMaster = True
End Function
Private Sub AzzeraInterfaccia()

On Error Resume Next

With SetMaster.Interfaccia
    .LabelInfo.Caption = ""
    .LabelInfo.Refresh
    .TestStop = False
    .ProgrBar.Value = 0
    .ProgrBar.Max = 100
End With

End Sub

Private Function CreaImmagineIso() As Boolean
Dim NomeExe As String, LineaComandi As String, i As Long
Dim TestLabel As String, TestProgr As Boolean
Dim DirCorta As String, NomeIso As String, Testo As String
Dim BufErrore As String, j As Long, Car As String
On Error Resume Next

TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False

TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False

If (PreparaCreaIso(LineaComandi) = False) Then
    CreaImmagineIso = False
    Exit Function
End If
DirCorta = TrovaNomeCorto(App.Path)

NomeExe = DirCorta & "\mkisofs.exe"
NomeIso = SetMaster.ISOImage.PathIsoImage

Call DeleteFile(NomeIso)
If (TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = "Creazione immagine ISO in corso ..."
        .LabelInfo.Refresh
    End With
End If

If (EseguiRedirect(NomeExe, LineaComandi) = False) Then
    Call DeleteFile(NomeIso)
    CreaImmagineIso = False
    Exit Function
End If

If (EsisteFile(NomeIso) = True) Then
    SetMaster.ISOImage.SizeIsoImage = FileLen(NomeIso)
    Call ControllaWarning("mkisofs:")
    CreaImmagineIso = True
    Exit Function
End If
SetMaster.Response.LastGeneralError = "Fallita creazione immagine iso"
BufErrore = TrovaErroremkisofs()
SetMaster.Response.LastSpecificError = BufErrore

CreaImmagineIso = False

End Function

Private Function TrovaErroremkisofs() As String
Dim i As Long, j As Long
Dim BufErrore As String
Dim Car As String, Testo As String
Dim Start As Long, BufLinea As String

Testo = SetMaster.Response.LastScreenLog
Start = 1
BufErrore = ""

Do
    BufLinea = ""
    i = InStr(Start, Testo, "error:", vbTextCompare)
    If (i > 0) Then
        BufLinea = PrendiLinea(Testo, i)
        Start = i
    End If
    
    If (BufLinea = "") Then
        i = InStr(Start, Testo, "warning:", vbTextCompare)
        If (i > 0) Then
            BufLinea = PrendiLinea(Testo, i)
            Start = i
        End If
    End If
    If (BufLinea = "") Then
        i = InStr(Start, Testo, "mkisofs:", vbTextCompare)
        If (i > 0) Then
            BufLinea = PrendiLinea(Testo, i)
            Start = i
        End If
    End If
    If (BufLinea <> "") Then
        BufErrore = BufErrore & BufLinea & vbCrLf
    End If
Loop While (BufLinea <> "" And Start < Len(Testo))

TrovaErroremkisofs = BufErrore

End Function
Private Function MasterizzaIso() As Boolean
Dim NomeIso As String
Dim LineaComandi As String, BufLinea As String
Dim NomeExe As String, i As Long, n As Long, StrNumero As String
Dim DirCorta As String, Testo As String
Dim TestOk As Boolean, BufErrore As String

With SetMaster.Response
    .LastGeneralError = ""
    .LastScreenLog = ""
    .LastSpecificError = ""
End With

NomeIso = SetMaster.ISOImage.PathIsoImage

If (NomeIso = "" Or EsisteFile(NomeIso) = False) Then
    With SetMaster.Response
        .LastGeneralError = "Fallita Masterizzazione CD"
        .LastSpecificError = "Non esiste alcuna immagine ISO da masterizzare"
    End With
    MasterizzaIso = False
    Exit Function
End If

DirCorta = TrovaNomeCorto(App.Path)
NomeExe = DirCorta & "\cdrecord.exe"
LineaComandi = LineaComandiPerRecord()

If (EseguiRedirect(NomeExe, LineaComandi) = False) Then
    MasterizzaIso = False
    Exit Function
End If
MasterizzaIso = VerificaMasterizzazione()

End Function

Private Function PreparaCreaIso(ByRef LineaComandi As String) As Boolean
Dim SizeIso As Double, NomeIso As String
Dim InfoDrive As Tipo_Drive, i As Long
Dim Testo As String, NF As Long
Dim NomeItems As String, NomeNoItems As String
Dim DirCorta As String
Dim NomeConfig As String
Dim PrimoPath As String, j As Long, Car As String
Dim NomeLog As String, BufErrore As String, TestLabel As Boolean
Dim TestProgr As Boolean

On Error Resume Next

NomeIso = SetMaster.ISOImage.PathIsoImage
If (NomeIso = "") Then
    NomeIso = PathIsoDefault()
    SetMaster.ISOImage.PathIsoImage = NomeIso
End If


SetMaster.ISOImage.PathIsoImage = NomeIso

Err.Clear
NF = FreeFile
DirCorta = TrovaNomeCorto(App.Path)

NomeItems = DirCorta & "\list_yes.txt"
NomeNoItems = DirCorta & "\list_no.txt"
Call DeleteFile(NomeItems)
Call DeleteFile(NomeNoItems)

Testo = ""
With SetMaster.ISOImage
    For i = 1 To .TotItems - 1
        Testo = Testo & UsaBarreUnix(.VetItems(i)) & Chr(10)
    Next i
End With
Err.Clear
Open NomeItems For Output As #NF
If (Err.Number <> 0) Then
    With SetMaster.Response
        .LastGeneralError = "Fallita creazione immagine iso"
        .LastSpecificError = "Errore in preparazione immagine iso. Errore creando file in : " & NomeItems
        .ErrorCode = cderr_FallitaCreazioneISO
    End With
    PreparaCreaIso = False
    Exit Function
End If
Print #NF, Testo;
Close #NF

Testo = ""
With SetMaster.ISOImage
    For i = 0 To .TotNoItems - 1
        Testo = Testo & UsaBarreUnix( _
                                .VetNoItems(i)) & Chr(10)
    Next i
End With
NF = FreeFile
Err.Clear

Open NomeNoItems For Output As #NF
If (Err.Number <> 0) Then
    With SetMaster.Response
        .LastGeneralError = "Fallita creazione immagine iso"
        .ErrorCode = cderr_FallitaCreazioneISO
        .LastSpecificError = "Errore in preparazione immagine iso. Errore creando file in : " & NomeNoItems
    End With
    PreparaCreaIso = False
    Exit Function
End If
Print #NF, Testo;
Close #NF


PrimoPath = SetMaster.ISOImage.VetItems(0)


LineaComandi = _
            "-gui -J -R -exclude-list " & NomeNoItems & _
            " -path-list " & NomeItems & _
            " -m '..' " & _
            " -o " & Chr(34) & NomeIso & Chr(34) & _
            " -graft-points " & _
            " '" & PrimoPath & "'"

With SetMaster.ISOImage
    If (.ExtraOptCmdLine <> "") Then
        LineaComandi = .ExtraOptCmdLine & " " & LineaComandi
    End If
End With

NomeConfig = DirCorta & "\mkiso.src"

NF = FreeFile
Open NomeConfig For Output As #NF

With SetMaster.ISOImage
    Testo = ""
    If (Trim(.Id_Application) <> "") Then
        Testo = "APPI=" & Trim(.Id_Application) & Chr(10)
    End If
    If (Trim(.Id_Preparer) <> "") Then
        Testo = Testo & "PREP=" & Trim(.Id_Preparer) & Chr(10)
    End If
    
    If (Trim(.Id_Publisher) <> "") Then
        Testo = Testo & "PUBL=" & Trim(.Id_Publisher) & Chr(10)
    End If
    
    If (Trim(.Id_VolumeIdentifier) <> "") Then
        Testo = Testo & "VOLI=" & Trim(.Id_VolumeIdentifier) & Chr(10)
    End If
    If (Trim(.Id_VolumeSetName) <> "") Then
        Testo = Testo & "VOLS=" & Trim(.Id_VolumeSetName) & Chr(10)
    End If
    If (Trim(.Id_Abstract) <> "") Then
        Testo = Testo & "ABST=" & Trim(.Id_Abstract) & Chr(10)
    End If
    If (Trim(.Id_Bibliographic) <> "") Then
        Testo = Testo & "BIBL=" & Trim(.Id_Bibliographic) & Chr(10)
    End If
    If (Trim(.Id_Copyright) <> "") Then
        Testo = Testo & "COPY=" & Trim(.Id_Copyright) & Chr(10)
    End If
    
End With
Print #NF, Testo;
Close #NF
PreparaCreaIso = True

End Function

Private Function TrovaDimFileSystem(ByRef DimFileSystem As Long) As Boolean

Dim NomeExe As String, LineaComandi As String, i As Long
Dim TestLabel As String, TestProgr As Boolean
Dim DirCorta As String, NomeIso As String, Testo As String
Dim BufErrore As String, j As Long, Car As String
Dim StrNumero As String
DimFileSystem = 0

TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False

TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False

If (PreparaCreaIso(LineaComandi) = False) Then
    TrovaDimFileSystem = False
    Exit Function
End If

LineaComandi = "-print-size " & LineaComandi

DirCorta = TrovaNomeCorto(App.Path)

NomeExe = DirCorta & "\mkisofs.exe"
NomeIso = SetMaster.ISOImage.PathIsoImage

If (TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = "Calcolo dimensione file system in corso ..."
        .LabelInfo.Refresh
    End With
End If

If (EseguiRedirect(NomeExe, LineaComandi) = False) Then
    TrovaDimFileSystem = False
    Exit Function
End If
Testo = SetMaster.Response.LastScreenLog

i = InStr(1, Testo, "Total extents", vbTextCompare)
If (i = 0) Then
    TrovaDimFileSystem = False
    Exit Function
End If

i = InStr(1, Testo, "=", vbBinaryCompare)
If (i = 0) Then
    TrovaDimFileSystem = False
    Exit Function
End If
i = i + 1
StrNumero = PrendiCifre(Testo, i)
DimFileSystem = Val(StrNumero) * 11.3
TrovaDimFileSystem = True


End Function
Private Function SoloPathAssoluto(NomeFile As String) As String

Dim i As Long
Dim Nome As String

i = InStr(1, NomeFile, "=", vbBinaryCompare)
If (i = 0) Then
    Nome = NomeFile
Else
    Nome = Trim(Mid(NomeFile, i + 1, Len(NomeFile)))
End If
SoloPathAssoluto = UsaBarreMsDos(Nome)
    
End Function
Private Function SoloPathRelativo(NomeFile As String) As String

Dim i As Long

i = InStr(1, NomeFile, "=", vbBinaryCompare)
If (i = 0) Then
    SoloPathRelativo = ""
Else
    SoloPathRelativo = UsaBarreMsDos(Left(NomeFile, i - 1))
End If

End Function

Private Function MasterizzaSenzaIso() As Boolean
Dim LineaComandi1 As String
Dim LineaComandi2 As String
Dim Nome As String, NomeIso As String

With SetMaster.Response
    .LastGeneralError = ""
    .LastScreenLog = ""
    .LastSpecificError = ""
End With



If (PreparaCreaIso(LineaComandi1) = False) Then
    MasterizzaSenzaIso = False
    Exit Function
End If
NomeIso = SetMaster.ISOImage.PathIsoImage
Nome = "-o " & Chr(34) & NomeIso & Chr(34)
LineaComandi1 = Replace(LineaComandi1, Nome, "", 1, 1, vbTextCompare)
LineaComandi2 = LineaComandiPerRecord()
Nome = Chr(34) & NomeIso & Chr(34)
LineaComandi2 = Replace(LineaComandi2, Nome, "-", 1, 1, vbTextCompare)

If (EseguiPipe(LineaComandi1, LineaComandi2) = False) Then
    MasterizzaSenzaIso = False
    Exit Function
End If
If (VerificaMasterizzazione() = False) Then
    With SetMaster.Response
        .LastGeneralError = "Fallita Masterizzazione CD senza immagine ISO"
        .ErrorCode = cderr_FallitaMasterizzazioneSenzaISO
    End With
    MasterizzaSenzaIso = False
Else
    MasterizzaSenzaIso = True
End If

End Function
Private Function LineaComandiPerRecord() As String
Dim LineaComandi As String
Dim NomeIso As String

NomeIso = SetMaster.ISOImage.PathIsoImage

With SetMaster.CdWriter
    LineaComandi = "-v speed=" & CStr(.Speed) & " " & _
                "-dev=" & CStr(.SCSI_Bus) & "," & _
                CStr(.SCSI_Id) & "," & _
                CStr(.SCSI_Lun) & " " & _
                Chr(34) & NomeIso & Chr(34)
    If (.DontWrite = True) Then
        LineaComandi = "-dummy " & LineaComandi
    End If
    If (.EspelliCD = True) Then
        LineaComandi = "-eject " & LineaComandi
    End If
    If (.ExtraOptCmdLine <> "") Then
        LineaComandi = .ExtraOptCmdLine & " " & LineaComandi
    End If
    If (.OverBurn = True) Then
        LineaComandi = "-overburn " & LineaComandi
    End If
    If (.BurnProof = True) Then
        LineaComandi = "driveropts=burnproof " & _
                                    LineaComandi
    End If
End With
LineaComandiPerRecord = LineaComandi
End Function
Private Function VerificaMasterizzazione() As Boolean
Dim TestOk As Boolean, Testo As String
Dim i As Long, BufErrore As String
Dim n As Long, StrNumero As String, BufLinea As String

TestOk = True
Testo = SetMaster.Response.LastScreenLog
i = InStr(1, Testo, "Writing  time:", vbTextCompare)
If (i > 0) Then
    i = i + Len("Writing  time:")
    StrNumero = PrendiCifre(Testo, i)
    n = Val(StrNumero)
    If (n < 2) Then TestOk = False
    
Else
    TestOk = False
End If

If (TestOk = True) Then
    i = InStr(1, Testo, "Fixating time:", vbTextCompare)
    If (i > 0) Then
        i = i + Len("Fixating time:")
        StrNumero = PrendiCifre(Testo, i)
        n = Val(StrNumero)
        If (n < 2 And _
            SetMaster.CdWriter.DontWrite = _
                    False) Then TestOk = False
    Else
        TestOk = False
    End If
End If

If (TestOk = True) Then
    i = InStr(1, Testo, "at speed", vbTextCompare)
    If (i > 0) Then
        i = i + Len("at speed")
        StrNumero = PrendiCifre(Testo, i)
        SetMaster.CdWriter.LastSpeed = Val(StrNumero)
    End If
    VerificaMasterizzazione = True
    Exit Function
End If
i = 1
BufErrore = TrovaErroreCdRecord()
If (BufErrore = "") Then
    BufErrore = TrovaErroremkisofs()
End If

With SetMaster.Response
    .ErrorCode = cderr_FallitaMasterizzazioneISO
    .LastGeneralError = "Fallita Masterizzazione CD"
    .LastSpecificError = BufErrore
End With

VerificaMasterizzazione = False
End Function
Private Function CancellaCD(TipoCanc As String) As Boolean
Dim LineaComandi As String, i As Long
Dim NomeExe As String
Dim BufErrore As String
Dim TestOk As Boolean, Testo As String
Dim MexBlank As String, StrNumero As Long

With SetMaster.CdWriter
    LineaComandi = "-v speed=" & CStr(.Speed) & " " & _
                "-dev=" & CStr(.SCSI_Bus) & "," & _
                CStr(.SCSI_Id) & "," & _
                CStr(.SCSI_Lun) & " blank=" & TipoCanc
                

End With
NomeExe = TrovaNomeCorto(App.Path) & "\cdrecord.exe"
If (EseguiRedirect(NomeExe, LineaComandi) = False) Then
    With SetMaster.Response
        .LastGeneralError = "Fallita cancellazione cd riscrivibile"
        .ErrorCode = cderr_FallitaCancellazioneCD
    End With
    CancellaCD = False
    Exit Function
End If
TestOk = True
MexBlank = "Blanking time:"
Testo = SetMaster.Response.LastScreenLog
i = InStr(1, Testo, MexBlank, vbTextCompare)
If (i = 0) Then
    TestOk = False
Else
    i = i + Len(MexBlank)
    StrNumero = PrendiCifre(Testo, i)
    i = Val(StrNumero)
    If (i < 4) Then
        TestOk = False
    End If
End If

If (TestOk = False) Then
    With SetMaster.Response
        .ErrorCode = cderr_FallitaCancellazioneCD
        .LastGeneralError = "Fallita cancellazione CD"
        .LastSpecificError = TrovaErroreCdRecord()
    End With
Else
    i = InStr(1, Testo, "at speed", vbTextCompare)
    If (i > 0) Then
        i = i + Len("at speed")
        StrNumero = PrendiCifre(Testo, i)
        SetMaster.CdWriter.LastSpeed = Val(StrNumero)
    End If
End If

CancellaCD = TestOk

End Function
Private Function TrovaErroreCdRecord() As String

Dim BufErrore As String, BufLinea As String
Dim Testo As String, i As Long

Testo = SetMaster.Response.LastScreenLog
i = 1
BufErrore = ""
Do
    i = InStr(i, Testo, "cdrecord:", vbTextCompare)
    If (i = 0) Then Exit Do
    i = i + Len("cdrecord:")
    BufLinea = PrendiLinea(Testo, i)
    If (InStr(1, BufLinea, "Warning: using inofficial", vbTextCompare) > 0) Then
        BufLinea = ""
    Else
        BufErrore = BufErrore & BufLinea & vbCrLf
    End If
Loop While (i < Len(Testo))

If (BufErrore = "") Then
    i = InStr(1, Testo, "Sense Code:", vbTextCompare)
    If (i > 0) Then
        i = i + Len("Sense Code:")
        
        BufErrore = PrendiLinea(Testo, i)
    End If
End If
TrovaErroreCdRecord = BufErrore

End Function
Private Function LeggiDaFile(NomeFile As String) As String
Dim Testo As String
Dim NF As Long
Dim BufLinea As String

On Error Resume Next

Err.Clear
NF = FreeFile
Open NomeFile For Input As #NF
If (Err.Number <> 0) Then
    LeggiDaFile = ""
    Exit Function
End If

While (EOF(NF) = False)
    Line Input #NF, BufLinea
    Testo = Testo & BufLinea & vbCrLf
Wend
Close #NF
LeggiDaFile = Testo

End Function


Private Function EseguiPipe( _
        LineaComandi1 As String, _
        LineaComandi2 As String) As Boolean
Dim NomeExe1 As String, StartTime As Long, i As Long
Dim NomeExe2 As String, DirCorta As String
Dim LineaFull As String, NomeOut As String
Dim ValShow As Long, NomeFine As String, StrNumero As String
Dim NomeBat As String, NF As Long, TestScrittura As Boolean
Dim MexInfo As String, Testo As String, Speed As Double
Dim TempoTotale As Double, TempValue As Double
Dim TempoNow As Double, ProcessId As Long
Dim TestLabel As Boolean, TestProgr As Boolean
Dim ContaPassaggi As Double
Dim hMkhyRead As Long, hMkhyWrite As Long
Dim hSalvaStdIn As Long, hSalvaStdOut As Long
Dim hSalvaStdErr As Long, GeneralError As String
Dim CodiceGenerale As Tipo_CD_Error
Dim proc As PROCESS_INFORMATION
Dim StartInf As STARTUPINFO, ValCar As Long
Dim saAttr As SECURITY_ATTRIBUTES
Dim HandleProcess1 As Long, HandleProcess2 As Long
Dim hCdRecordRead As Long, hCdRecordWrite As Long
Dim RetVal As Long, NBytes As Long
Dim OldLinea As String, OldIgnora As Boolean
Dim LastPerc As Double, TestIgnora As Boolean
Dim dwRead As Long, dwWritten As Long
Dim chBuf As String * 256, n As Long
Dim VetBytes(256) As Byte, BufLinea As String
Dim Totale As Double, Parziale As Double
Dim LineaShell As String, InfoU As Tipo_InfoUtility
Dim TestPrimo As Boolean, LineaFinale As String
Const UN_MB = 1048576

CodiceGenerale = cderr_FallitaMasterizzazioneSenzaISO
GeneralError = "Fallita masterizzazione senza immagine ISO"
With SetMaster.Response
    .LastGeneralError = ""
    .ErrorCode = cderr_NessunErrore
    .LastSpecificError = ""
    .LastScreenLog = ""
End With
On Error Resume Next
Call AzzeraInterfaccia

If (CalcolaDimensioneItems(True) = False) Then
    EseguiPipe = False
    Exit Function
End If

If (SetMaster.ISOImage.SizeItems = 0) Then
    With SetMaster.Response
        .ErrorCode = CodiceGenerale
        .LastGeneralError = GeneralError
        .LastSpecificError = "Non ci sono file da inserire nell'immagine"
        .LastScreenLog = ""
    End With
    EseguiPipe = False
    Exit Function
End If

TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False

TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False

SetMaster.Interfaccia.TestStop = False
SetMaster.ISOImage.PathIsoImage = ""
DirCorta = TrovaNomeCorto(App.Path)

NomeExe1 = DirCorta & "\mkisofs.exe"
NomeExe2 = DirCorta & "\cdrecord.exe"

If (EsisteFile(NomeExe1) = False) Then
    With SetMaster.Response
        .ErrorCode = cderr_NonTrovato_mkisofs_exe
        .LastGeneralError = "Manca l'utility 'mkisofs.exe' nella directory dell'applicazione"
        .LastSpecificError = "Manca l'utility 'mkisofs.exe' nella directory dell'applicazione"
    End With
    EseguiPipe = False
    Exit Function
End If
If (EsisteFile(NomeExe2) = False) Then
    With SetMaster.Response
        .ErrorCode = cderr_NonTrovato_CdRecord_exe
        .LastGeneralError = "Manca l'utility 'CdRecord.exe' nella directory dell'applicazione"
        .LastSpecificError = "Manca l'utility 'CdRecord.exe' nella directory dell'applicazione"
    End With
    EseguiPipe = False
    Exit Function
End If

With SetMaster.ISOImage
    Totale = .SizeIsoImage
    Totale = Totale / UN_MB
End With

With InfoU
    .BufLinea = ""
    .LineaShell = NomeExe1 & " " & LineaComandi1 & " | " & _
                  NomeExe2 & " " & LineaComandi2
    .LogOut = ""
    .TestCdRecord = True
    .TestLabel = TestLabel
    .TestProgr = TestProgr
    .TestReport = False
    .TotMB = Totale
    .UltimoComando = SetMaster.Response.LastCommand
    .CodiceGenerale = CodiceGenerale
    .GeneralError = GeneralError
    .MexOperazioneInCorso = "Masterizzazione CD senza immagine ISO in corso ..."
End With

If (IsWindowsXP() = True) Then
    LineaShell = NomeExe1 & " " & _
                LineaComandi1 & " | " & _
                NomeExe2 & " " & _
                LineaComandi2
    EseguiPipe = EseguiCon_BAT(LineaShell, InfoU)
    
    Exit Function
End If

hSalvaStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
hSalvaStdErr = GetStdHandle(STD_ERROR_HANDLE)
hSalvaStdIn = GetStdHandle(STD_INPUT_HANDLE)
With saAttr
    .nLength = Len(saAttr)
    .bInheritHandle = True
    .lpSecurityDescriptor = vbNull
End With

If (CreatePipe(hMkhyRead, hMkhyWrite, saAttr, 0) = 0) Then
    EseguiPipe = False
    Exit Function
End If

Call SetStdHandle(STD_OUTPUT_HANDLE, hMkhyWrite)

With StartInf

    .cb = Len(StartInf)
    
    If (SetMaster.Interfaccia.TestNascondi = True) Then
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
    End If
End With

If (InfoU.TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = InfoU.MexOperazioneInCorso
        .LabelInfo.Refresh
    End With
End If
RetVal = CreateProcess(NomeExe1, _
            NomeExe1 & " " & LineaComandi1, _
            vbNullString, _
            vbNullString, _
            True, _
            0, _
            vbNullString, _
            SoloDir(NomeExe1), _
            StartInf, _
            proc)
            
            
If (RetVal = 0) Then
    With SetMaster.Response
        .LastGeneralError = GeneralError
        .ErrorCode = CodiceGenerale
        .LastSpecificError = "Errore creando sottoprocesso: " & _
            Chr(34) & NomeExe1 & Chr(34) & vbCrLf & _
            "Con linea comandi: " & Chr(34) & LineaComandi1 & Chr(34)
    End With
    EseguiPipe = False
    Exit Function
End If
HandleProcess1 = OpenProcess( _
            PROCESS_TERMINATE, _
            False, proc.dwProcessID)

Call SetStdHandle(STD_OUTPUT_HANDLE, hSalvaStdOut)
Call CloseHandle(hMkhyWrite)
Call SetStdHandle(STD_INPUT_HANDLE, hMkhyRead)

If (CreatePipe(hCdRecordRead, hCdRecordWrite, saAttr, 0) = 0) Then
    With SetMaster.Response
        .ErrorCode = CodiceGenerale
        .LastGeneralError = GeneralError
        .LastSpecificError = "Creazione pipe di input fallita"
    End With
    EseguiPipe = False
    Exit Function
End If

Call SetStdHandle(STD_OUTPUT_HANDLE, hCdRecordWrite)
Call SetStdHandle(STD_ERROR_HANDLE, hCdRecordWrite)

RetVal = CreateProcess(NomeExe2, _
            NomeExe2 & " " & LineaComandi2, _
            vbNullString, _
            vbNullString, _
            True, _
            0, _
            vbNullString, _
            SoloDir(NomeExe2), _
            StartInf, _
            proc)
            
            
If (RetVal = 0) Then
    With SetMaster.Response
        .LastGeneralError = GeneralError
        .ErrorCode = CodiceGenerale
        .LastSpecificError = "Errore creando sottoprocesso: " & _
            Chr(34) & NomeExe2 & Chr(34) & vbCrLf & _
            "Con linea comandi: " & Chr(34) & LineaComandi2 & Chr(34)
    End With
    EseguiPipe = False
    Exit Function
End If
HandleProcess2 = OpenProcess( _
            PROCESS_TERMINATE, _
            False, proc.dwProcessID)

Call SetStdHandle(STD_OUTPUT_HANDLE, hSalvaStdOut)
Call SetStdHandle(STD_INPUT_HANDLE, hSalvaStdIn)
Call SetStdHandle(STD_ERROR_HANDLE, hSalvaStdErr)
Call CloseHandle(hCdRecordWrite)

NBytes = 0
Testo = ""
OldLinea = ""
OldIgnora = False
LastPerc = 0

TestPrimo = True


Do
    TestIgnora = False
    If (ReadFile(hCdRecordRead, _
                    ByVal chBuf, 1, _
                    dwRead, ByVal 0&) = 0) Then Exit Do
    ValCar = Asc(chBuf)
    LineaFinale = ""
    Select Case ValCar
        Case 13, 10
            For i = 0 To NBytes - 1
                LineaFinale = LineaFinale & Chr(VetBytes(i))
            Next i
            NBytes = 0
            
        Case 8
            NBytes = NBytes - 1
            If (NBytes < 0) Then NBytes = 0
        Case 9
            For i = 0 To 3
                VetBytes(NBytes + i) = 32
            Next i
            NBytes = NBytes + 4
        Case Else
            VetBytes(NBytes) = ValCar
            NBytes = NBytes + 1
    End Select
    If (LineaFinale <> "" Or TestPrimo = True) Then
        InfoU.BufLinea = LineaFinale
    
        Call ScansioneLinea(InfoU)
    End If
    TestPrimo = False

    If (SetMaster.Interfaccia.ControllaSeStop = True) Then
        DoEvents
        If (SetMaster.Interfaccia.TestStop = True) Then
            Call TerminateProcess(HandleProcess1, -1)
            Call TerminateProcess(HandleProcess2, -1)
            Call AzzeraInterfaccia
            With SetMaster.Response
                .ErrorCode = cderr_InterrottoDaUtente
                .LastScreenLog = Testo
                .LastGeneralError = GeneralError
                .LastSpecificError = "Operazione interrotta dall'utente"
            End With
            EseguiPipe = False
            Exit Function
        End If
    End If
Loop While (dwRead <> 0)

Call TerminateProcess(HandleProcess1, -1)
Call TerminateProcess(HandleProcess2, -1)

SetMaster.Response.LastScreenLog = InfoU.LogOut
Call AzzeraInterfaccia

EseguiPipe = True
End Function
Private Function CreaListFile( _
        TipoList As String, _
        NewNomeIso As String) As Boolean

Dim NomeIso As String, Testo As String
Dim NomeExe As String, i As Long, j As Long
Dim LineaComandi As String, TestEmpty As Boolean
Dim VetRighe() As String, NRighe As Long, z As Long
Dim VetParti() As String, NParti As Long
Dim VetOldParti() As String, NOldParti As Long
Dim BufLinea As String, Pezzo As String
Dim TestLabel As Boolean, NomeDir As String
Dim TestDir As Boolean, TestFile As Boolean
Dim NomeFile As String, DataFile As Date
Dim SizeFile As Long, NMese As Long
Dim VetMesiEng As Variant, NGiorno As Long
Dim MexData As String, LogRef As String
Dim TestProgr As Boolean, TestIgnora As Boolean
Dim NomeOut As String, DirCorta As String
Dim Totale As Long, NF As Long

VetMesiEng = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

On Error Resume Next

TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False

TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False

If (NewNomeIso <> "") Then
    NomeIso = NewNomeIso
Else
    NomeIso = SetMaster.ISOImage.PathIsoImage
End If

If (NomeIso = "" Or EsisteFile(NomeIso) = False) Then
    With SetMaster.Response
        .ErrorCode = cderr_FallitoReportList
        .LastGeneralError = "Impossibile creare un report list dei file di immagine ISO"
        .LastSpecificError = "Non esiste nessuna immagine ISO da analizzare"
    End With
    CreaListFile = False
    Exit Function
End If
DirCorta = TrovaNomeCorto(App.Path)

NomeExe = DirCorta & "\isoinfo.exe"

If (EsisteFile(NomeExe) = False) Then
    With SetMaster.Response
        .ErrorCode = cderr_NonTrovato_IsoInfo_exe
        .LastGeneralError = "In percorso applicazione corrente manca utility 'isoinfo.exe'"
    End With
    CreaListFile = False
    Exit Function
End If

Select Case TipoList
    Case "FlatList", "TreeList"
        LineaComandi = " -J -f"
    Case "ReferenceList"
        LineaComandi = " -J -l"
    Case Else
        With SetMaster.Response
            .ErrorCode = cderr_ArgomentoNonValido
            .LastGeneralError = "Opzione non valida"
            .LastSpecificError = "Per creare un Report Iso List File, devi fornire il tipo di list desiderato, scegliendo tra: 'FlatList' , 'TreeList' o 'ReferenceList'"
        End With
        CreaListFile = False
        Exit Function
End Select

If (TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = "Scansione file in corso..."
        .LabelInfo.Refresh
    End With
End If
Totale = NumeroFileIso(NomeIso)

LineaComandi = LineaComandi & " -i " & Chr(34) & NomeIso & Chr(34)


If (EseguiRedirect(NomeExe, LineaComandi, Totale) = False) Then
    With SetMaster.Response
        .ErrorCode = cderr_FallitoReportList
        .LastGeneralError = "Fallita creazione report list da immagine ISO"
    End With
    CreaListFile = False
    Exit Function
End If

Testo = SetMaster.Response.LastScreenLog
i = InStr(1, Testo, "isoinfo:", vbTextCompare)

If (i > 0) Then
    i = i + Len("isoinfo:")
    With SetMaster.Response
        .ErrorCode = cderr_FallitoReportList
        .LastGeneralError = "Fallita creazione report list"
        .LastSpecificError = PrendiLinea(Testo, i)
    End With
    CreaListFile = False
    Exit Function
End If

If (TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = "Formattazione report list in corso..."
        .LabelInfo.Refresh
    End With
End If

Select Case TipoList
    Case "FlatList"
        SetMaster.Response.ResultValue = Testo
    Case "TreeList"
        VetRighe = Split(Testo, vbCrLf)
        NRighe = UBound(VetRighe) + 1
        Testo = ""
        NOldParti = 0
        For i = 0 To NRighe - 1
            If (TestProgr = True) Then
                Call MostraPercentuale(SetMaster.Interfaccia.ProgrBar, _
                        NRighe, i)
            End If
            VetParti = Split(VetRighe(i), "/")
            NParti = UBound(VetParti) + 1

            For j = 0 To NParti - 1
                If (j >= NOldParti) Then Exit For
                If (VetParti(j) <> VetOldParti(j)) Then Exit For
            Next j
            BufLinea = ""
            For z = 0 To j - 1
                If (BufLinea <> "") Then
                    BufLinea = BufLinea & " "
                End If
                Pezzo = String(Len(VetParti(z)), " ")
                
                BufLinea = BufLinea & Pezzo
            Next z
            
            For z = j To NParti - 1
                If (BufLinea <> "") Then
                    BufLinea = BufLinea & "/"
                End If
                BufLinea = BufLinea & VetParti(z)
            Next z
            ReDim VetOldParti(NParti)
            For j = 0 To NParti - 1
                VetOldParti(j) = VetParti(j)
            Next j
            NOldParti = NParti
            Testo = Testo & BufLinea & vbCrLf
        Next i
        SetMaster.Response.ResultValue = Testo
        
    Case "ReferenceList"
        With SetMaster.GlobArray
            .TotRefList = 0
        End With
        
        VetRighe = Split(Testo, vbCrLf)
        NRighe = UBound(VetRighe) + 1
        NomeDir = ""
        LogRef = ""
        For i = 0 To NRighe - 1
            If (TestProgr = True) Then
                Call MostraPercentuale(SetMaster.Interfaccia.ProgrBar, _
                        NRighe, i)
            End If
            BufLinea = VetRighe(i)
            TestDir = False
            TestFile = False
            Select Case Left(BufLinea, 2)
                Case "Di"
                    NomeDir = Mid(BufLinea, 22, Len(BufLinea))
                Case "d-"
                    TestDir = True
                Case "--"
                    TestFile = True
            End Select
            
            If (TestDir = True Or TestFile = True) Then
                NomeFile = Mid(BufLinea, 65, Len(BufLinea))
                NomeFile = Trim(NomeFile)
                If (TestFile = True) Then
                    Testo = Mid(BufLinea, 42, 3)
                    For j = 0 To 11
                        If (Testo = VetMesiEng(j)) Then Exit For
                    Next j
                    If (j = 12) Then
                        Call MsgBox("Errore scandendo mesi")
                        j = 11
                    End If
                    NMese = j + 1
                    Testo = Mid(BufLinea, 46, 2)
                    NGiorno = Val(Testo)
                    MexData = CStr(NGiorno) & "/" & CStr(NMese) & "/" & Mid(BufLinea, 49, 4)
                    DataFile = CDate(MexData)
                    j = 27
                    Testo = PrendiCifre(BufLinea, j)
                    SizeFile = Val(Testo)
                Else
                    NomeFile = NomeFile & "\"
                    SizeFile = 0
                    DataFile = CDate("0")
                    
                End If
                If (NomeFile <> ".\" And NomeFile <> "..\") Then
                    NomeFile = UsaBarreMsDos(NomeDir & "\" & NomeFile)
                    NomeFile = Replace(NomeFile, "\\", "\", 1, -1, vbBinaryCompare)
                    
                    With SetMaster.GlobArray
                        TestIgnora = False
                        If (.TotRefList > 0) Then
                            If (.VetRefList(.TotRefList - 1).Nome = _
                                NomeFile) Then TestIgnora = True
                        End If
                        If (TestIgnora = False) Then
                            ReDim Preserve .VetRefList(.TotRefList)
                                                
                            With .VetRefList(.TotRefList)
                                .Data = DataFile
                                .Dimensione = SizeFile
                                .Nome = NomeFile
                            End With
                            .TotRefList = .TotRefList + 1
                        End If
                    End With
                    If (TestIgnora = False) Then
                        LogRef = LogRef & NomeFile & vbTab & CStr(SizeFile) & vbTab & Format(DataFile, "dd/mm/yyyy") & vbCrLf
                    End If
                End If
                         
            End If
        Next i
        SetMaster.Response.ResultValue = LogRef
                
End Select
Call AzzeraInterfaccia
CreaListFile = True
 

End Function



Private Function NumeroFileIso(NomeIso As String) As Long

Dim NomeBat As String, i As Long
Dim NomeExe As String, NomeOut As String
Dim DirCorta As String, NF As Long
Dim Testo As String, Totale As Long
Dim InfoU As Tipo_InfoUtility
Dim TestLabel As Boolean

On Error Resume Next

With SetMaster.Interfaccia
    Err.Clear
    .LabelInfo.Refresh
    If (Err.Number = 0) Then
        TestLabel = True
    Else
        TestLabel = False
    End If
End With

NomeExe = TrovaNomeCorto(App.Path) & "\isoinfo.exe"

Testo = NomeExe & " -i " & Chr(34) & NomeIso & Chr(34) & _
        " -f -J "
With InfoU
    .BufLinea = ""
    .CodiceGenerale = cderr_FallitoReportList
    .GeneralError = "Errore in scansione immagine ISO"
    .LineaShell = Testo
    .LogOut = ""
    .TestCdRecord = False
    .TestLabel = TestLabel
    .TestProgr = False
    .TestReport = False
    .TotMB = 0
    .UltimoComando = SetMaster.Response.LastCommand
End With
Call ImpostaOperazioneInCorso(InfoU)

If (EseguiCon_BAT(Testo, InfoU) = False) Then
    NumeroFileIso = 0
    Exit Function
End If

Totale = 0
Testo = SetMaster.Response.LastScreenLog
i = 1
While (i <> 0)
    i = InStr(i, Testo, Chr(10), vbBinaryCompare)
    If (i > 0) Then
        Totale = Totale + 1
        i = i + 1
    End If
Wend

Close #NF
NumeroFileIso = Totale
End Function
Private Function CreaListFileItem(TipoList As String) As Boolean

Dim i As Long, j As Long, TestLabel As Boolean
Dim TestSource As Boolean, TestProgr As Boolean
Dim VetElementi() As String, TotElementi As Long
Dim Nome As String, VetIndiceRelativo() As Long
Dim ItemNo As String, LogRef As String
Dim NomeFile As String, DataFile As Date
Dim SizeFile As Long, z As Long
Dim PathRel As String
Dim PathAbs As String

On Error Resume Next
Call AzzeraInterfaccia

With SetMaster
    .Response.ResultValue = ""
    .GlobArray.TotRefList = 0
End With

Select Case TipoList
    Case "Source"
        TestSource = True
    Case "Target"
        TestSource = False
    Case Else
        With SetMaster.Response
            .ErrorCode = cderr_ArgomentoNonValido
            .LastGeneralError = "Argomento non valido per comando cw_ReportItemsFileList"
            .LastSpecificError = "Gli unici argomenti validi sono " & Chr(34) & "Target" & Chr(34) & " oppure " & Chr(34) & "Source" & Chr(34)
        End With
        CreaListFileItem = False
        Exit Function
End Select

TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False

TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False

If (TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = "Creazione report file list in corso..."
        .LabelInfo.Refresh
    End With
End If
TotElementi = 0

With SetMaster.ISOImage
    For i = 0 To .TotItems - 1
        If (TestProgr = True) Then
            Call MostraPercentuale(SetMaster.Interfaccia.ProgrBar, _
                    .TotItems, i)
        End If
        With SetMaster.Interfaccia
            If (.ControllaSeStop = True) Then
                DoEvents
                If (.TestStop = True) Then
                    With SetMaster.Response
                        .ErrorCode = cderr_InterrottoDaUtente
                        .LastGeneralError = "Operazione interrotta dall'utente"
                    End With
                    Call AzzeraInterfaccia
                    CreaListFileItem = False
                    Exit Function
                End If
            End If
        End With
        
        Nome = SoloPathAssoluto(.VetItems(i))
        
        If (EsisteDir(Nome) = True) Then
            TotFiles = 0
            TotFolders = 0
            Call ScansioneCartelle(Nome)
            ReDim Preserve VetElementi(TotElementi + TotFiles + TotFolders)
            ReDim Preserve VetIndiceRelativo(TotElementi + TotFiles + TotFolders)
            
            For j = 0 To TotFolders - 1
                VetElementi(TotElementi) = VetFolders(j) & "\"
                VetIndiceRelativo(TotElementi) = i
                TotElementi = TotElementi + 1
            Next j
            
            For j = 0 To TotFiles - 1
                VetElementi(TotElementi) = VetFiles(j)
                VetIndiceRelativo(TotElementi) = i
                TotElementi = TotElementi + 1
            Next j
        Else
            ReDim Preserve VetElementi(TotElementi)
            ReDim Preserve VetIndiceRelativo(TotElementi)
            VetElementi(TotElementi) = Nome
            VetIndiceRelativo(TotElementi) = i
            TotElementi = TotElementi + 1
        End If
    Next i
    LogRef = ""
    If (TestLabel = True) Then
        With SetMaster.Interfaccia
            .LabelInfo.Caption = "Formattazione report file list in corso..."
            .LabelInfo.Refresh
        End With
    End If
    For i = 0 To TotElementi - 1
        If (TestProgr = True) Then
            Call MostraPercentuale(SetMaster.Interfaccia.ProgrBar, _
                TotElementi, i)
        End If
        With SetMaster.Interfaccia
            If (.ControllaSeStop = True) Then
                DoEvents
                If (.TestStop = True) Then
                    With SetMaster.Response
                        .ErrorCode = cderr_InterrottoDaUtente
                        .LastGeneralError = "Operazione interrotta dall'utente"
                    End With
                    Call AzzeraInterfaccia
                    CreaListFileItem = False
                    Exit Function
                End If
            End If
        End With
        Nome = VetElementi(i)

        For j = 0 To .TotNoItems - 1
            ItemNo = UsaBarreMsDos(.VetNoItems(j))
            
            If (ItemNo = Left(Nome, Len(ItemNo))) Then
                z = Len(ItemNo) + 1
                If (z > Len(Nome)) Then Exit For
                If (Mid(Nome, z, 1) = "\") Then Exit For
            End If
        Next j
        
        If (j = .TotNoItems) Then
            If (EsisteDir(Nome) = True) Then
                NomeFile = Nome
                SizeFile = 0
                DataFile = CDate("0")
            Else
                NomeFile = Nome
                SizeFile = FileLen(Nome)
                DataFile = FileDateTime(Nome)
            End If
            
            If (TestSource = False) Then
                z = VetIndiceRelativo(i)
                NomeFile = PathRelativoFull(.VetItems(z), Nome)
             End If
            LogRef = LogRef & _
                    NomeFile & vbTab & _
                    CStr(SizeFile) & vbTab & _
                    Format(DataFile, "dd/mm/yyyy") & vbCrLf
            With SetMaster.GlobArray
                ReDim Preserve .VetRefList(.TotRefList)
                With .VetRefList(.TotRefList)
                    .Data = DataFile
                    .Dimensione = SizeFile
                    .Nome = NomeFile
                End With
                .TotRefList = .TotRefList + 1
            End With
        End If
    Next i
End With
Call AzzeraInterfaccia
SetMaster.Response.ResultValue = LogRef
CreaListFileItem = True

End Function

Private Sub ControllaWarning(NomeExe As String)
Dim i As Long
Dim Testo As String

Testo = SetMaster.Response.LastScreenLog

i = InStr(1, Testo, NomeExe, vbTextCompare)
If (i = 0) Then Exit Sub
i = i + Len(NomeExe)

With SetMaster.Response
    .ErrorCode = cderr_Warning
    .LastGeneralError = "Operazione conclusa con successo ma c'e' un messaggio di avviso (warning)"
    .LastSpecificError = PrendiLinea(Testo, i)
End With

End Sub
Private Function PathRelativoFull( _
        StringaItem As String, NomeAbs As String) As String

Dim PathAbs As String, PathRel As String
Dim NomeFile As String

PathAbs = SoloPathAssoluto(StringaItem)
PathRel = SoloPathRelativo(StringaItem)

NomeFile = PathRel & Replace(NomeAbs, PathAbs, "", 1, 1, vbTextCompare)

If (EsisteDir(PathAbs) = False And _
    Right(PathRel, 1) = "\") Then
    NomeFile = NomeFile & SoloNome(PathAbs)
End If

NomeFile = Replace(NomeFile, "\\", "\", 1, -1, vbBinaryCompare)

If (Left(NomeFile, 1) <> "\") Then
    NomeFile = "\" & NomeFile
End If
PathRelativoFull = NomeFile

End Function
Private Function VerificaCD(TipoVerifica As String) As Boolean

Dim InfoDrive As Tipo_InfoUnita

Dim ErrGenerale As String
Dim Esito As Boolean
Dim Lettera As String, TestLabel As Boolean
Dim TestProgr As Boolean, TempoNow As Long
Dim TestRapido As Boolean, i As Long
Dim Nome As String, NomeDir As String
Dim SpecErrore As String, StartTime
Dim TestErrore As Boolean

On Error Resume Next

Call AzzeraInterfaccia
If (SetMaster.CdWriter.DontWrite = True) Then
    VerificaCD = True
    Exit Function
End If

TestLabel = True
Err.Clear
SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number <> 0) Then TestLabel = False

TestProgr = True
Err.Clear
SetMaster.Interfaccia.ProgrBar.Max = 100
If (Err.Number <> 0) Then TestProgr = False

Select Case TipoVerifica
    Case "Fast"
        TestRapido = True
    Case "All"
        TestRapido = False
    Case Else
        With SetMaster.Response
            .ErrorCode = cderr_ArgomentoNonValido
            .LastGeneralError = "Argomento non valido per comando cw_VerificaCD"
            .LastSpecificError = "Gli argomenti validi per comando cw_VerificaCD sono o 'FAST' oppure 'ALL'"
        End With
        VerificaCD = False
        Exit Function
End Select

ErrGenerale = "Il CD non e' stato masterizzato in modo corretto"
Lettera = SetMaster.CdWriter.LetteraUnita

Call CassettoCD(Lettera, True)
Call AttendiD(2)
Call CassettoCD(Lettera, False)

StartTime = GetTickCount
Do
    Call GetInfoUnita(Lettera, InfoDrive)
    If (InfoDrive.Pronta = True) Then Exit Do
    
    Call Attendi(5)
    Call GetInfoUnita(Lettera, InfoDrive)
    If (InfoDrive.Pronta = True) Then Exit Do
    
    If (MsgBox("L'unita' " & Left(Lettera, 1) & _
        ": non e' pronta" & vbCrLf & _
        "Rimango ancora in attesa che diventi pronta?", vbYesNo) = vbNo) Then Exit Do

Loop While (InfoDrive.Pronta = False)

If (InfoDrive.Pronta = False) Then
    With SetMaster.Response
        .ErrorCode = cderr_FallitaMasterizzazioneISO
        .LastGeneralError = ErrGenerale
        .LastSpecificError = "Unita' non pronta:  cd illeggibile (?)"
    End With
    VerificaCD = False
    Exit Function
End If

NomeDir = Left(Lettera, 1) & ":"

If (TestRapido = True) Then
    Nome = Dir(NomeDir & "\*.*")
    
    While (Nome <> "")
        If (Nome <> "." And Nome <> "..") Then
            VerificaCD = True
            Exit Function
        End If
        Nome = Dir
    Wend
    
    With SetMaster.Response
        .ErrorCode = cderr_FallitaMasterizzazioneISO
        .LastGeneralError = ErrGenerale
        .LastSpecificError = "Il CD risulta vuoto"
    End With
    Call AzzeraInterfaccia
    VerificaCD = False
    Exit Function
End If

If (CreaListFileItem("Target") = False) Then
    VerificaCD = False
    Exit Function
End If
Call AzzeraInterfaccia

If (TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = "Verifica del cd: controllo di tutti i files in corso..."
        .LabelInfo.Refresh
    End With
End If

With SetMaster.GlobArray
    For i = 0 To .TotRefList - 1
        If (TestProgr = True) Then
            With SetMaster
                Call MostraPercentuale(.Interfaccia.ProgrBar, .GlobArray.TotRefList, i)
            End With
        End If
        
        If (SetMaster.Interfaccia.ControllaSeStop = True) Then
            DoEvents
            If (SetMaster.Interfaccia.TestStop = True) Then
                With SetMaster.Response
                    .ErrorCode = cderr_InterrottoDaUtente
                    .LastGeneralError = "Verifica CD interrotta dall'utente"
                End With
                Call AzzeraInterfaccia
                VerificaCD = False
                Exit Function
            End If
        End If
        
        Nome = NomeDir & .VetRefList(i).Nome
        SpecErrore = ""
        If (Right(Nome, 1) = "\") Then
            If (EsisteDir(Nome) = False) Then
                SpecErrore = "Cartella " & Chr(34) & Nome & Chr(34) & " non esiste sul CD"
            End If
        Else
            If (EsisteFile(Nome) = False) Then
                SpecErrore = "File " & Chr(34) & Nome & Chr(34) & " non esiste sul CD"
            Else
                If (FileLen(Nome) <> .VetRefList(i).Dimensione) Then
                    SpecErrore = "File " & Chr(34) & Nome & Chr(34) & " ha sul cd una dimensione diversa dall'originale"
                End If
            End If
        End If
        If (SpecErrore <> "") Then
            With SetMaster.Response
                .ErrorCode = cderr_FallitaMasterizzazioneISO
                .LastGeneralError = ErrGenerale
                .LastSpecificError = SpecErrore
            End With
            Call AzzeraInterfaccia
            VerificaCD = False
            Exit Function
        End If
    Next i
End With
Call AzzeraInterfaccia
VerificaCD = True

End Function
Private Function SoloNome(ByVal NomeFile As String) As String

Dim i As Integer

For i = Len(NomeFile) To 1 Step -1
    If (Mid(NomeFile, i, 1) = "\" Or Mid(NomeFile, i, 1) = ":") Then
        SoloNome = Mid(NomeFile, i + 1, Len(NomeFile) + 1 - i)
        Exit Function
    End If
Next i

SoloNome = NomeFile


End Function
Private Function EsisteFile(NomeFile As String) As Boolean

On Error Resume Next
Err.Clear
If (Trim(NomeFile) = "") Then
    EsisteFile = False
    Exit Function
End If

If (Dir(NomeFile) = "") Then EsisteFile = False Else EsisteFile = True
If (Err.Number <> 0) Then
    EsisteFile = False
End If

End Function
Private Function SoloDir(NomePath As String) As String
Dim i As Integer

For i = Len(NomePath) To 1 Step -1
    If (Mid(NomePath, i, 1) = "\") Then
        SoloDir = Left(NomePath, i - 1)
        Exit Function
    End If
Next i
SoloDir = ""

End Function
Private Function PrendiCifre( _
        BufLinea As String, _
        ByRef Indice As Long) As String

Dim i As Long, j As Long
Dim Car As String
Dim StrCifre As String

For i = Indice To Len(BufLinea)
    If (IsNumeric(Mid(BufLinea, i, 1)) = True) Then Exit For
Next i

If (i > Len(BufLinea)) Then
    PrendiCifre = ""
    Indice = i
    Exit Function
End If

StrCifre = ""
For j = i To Len(BufLinea)
    Car = Mid(BufLinea, j, 1)
    If (IsNumeric(Car) = False) Then Exit For
    StrCifre = StrCifre & Car
Next j
PrendiCifre = StrCifre
Indice = j

End Function


Private Sub MostraPercentuale(ProgBar As Control, _
                            ByVal Totale As Double, _
                            ByVal Parziale As Double)

Dim Valore As Double
Dim Temp As Long

On Error Resume Next

If (Totale = 0) Then Totale = 1
If (Parziale > Totale) Then
    Parziale = Totale
End If

Valore = Parziale / Totale
Valore = Int(Valore * 100)
Temp = Valore
If (Temp = ProgBar.Value) Then Exit Sub

ProgBar.Value = Valore
ProgBar.Refresh

End Sub



Private Function EsisteDir(ByVal NomeDir As String) As Boolean

On Error Resume Next
Err.Clear
If (Dir(NomeDir, vbDirectory) <> "") Then
    If (GetAttr(NomeDir) And vbDirectory) Then
        EsisteDir = True
    Else
        EsisteDir = False
    End If
Else
    EsisteDir = False
End If
If (Err.Number <> 0) Then EsisteDir = False
End Function
Private Function TrovaDimFolder( _
        BaseDir As String, _
        Optional MaskSearch As String = "*") As Double

Dim i As Long
Dim j As Long, Totale As Double
Dim NomeDir As String
Dim VetNomi() As String, NNomi As Long

TotFiles = 0
TotFolders = 0
Call ScansioneCartelle(BaseDir)

Totale = 0

If (MaskSearch <> "*") Then
    For i = 0 To TotFolders - 1
        NomeDir = VetFolders(i)
        VetNomi = TrovaFiles(NomeDir, MaskSearch, NNomi)
        For j = 0 To NNomi - 1
            Totale = Totale + FileLen(VetNomi(j))
        Next j
    Next i
Else
    For i = 0 To TotFiles - 1
        Totale = Totale + FileLen(VetFiles(i))
    Next i

End If

TrovaDimFolder = Totale

End Function

Private Function IsMask(NomeFile As String) As Boolean

If (InStr(1, NomeFile, "*", vbBinaryCompare) > 0 Or _
    InStr(1, NomeFile, "?", vbBinaryCompare) > 0) Then
    IsMask = True
Else
    IsMask = False
End If

End Function

Private Function UsaBarreUnix(NomeFile As String) As String
    UsaBarreUnix = Replace(NomeFile, "\", "/", 1, -1, vbBinaryCompare)
End Function

Private Function UsaBarreMsDos(NomeFile As String) As String
    UsaBarreMsDos = Replace(NomeFile, "/", "\", 1, -1, vbBinaryCompare)
End Function

Private Function ScansioneCartelle(ByVal ActPath$) As Boolean

    Dim AttributiFile  As Integer
    Dim ElencoCartelle$
    Dim file$

DoEvents

ReDim Preserve VetFolders(TotFolders)
VetFolders(TotFolders) = ActPath$
TotFolders = TotFolders + 1

ElencoCartelle$ = ""
If Right$(ActPath$, 1) <> "\" Then ActPath$ = ActPath$ & "\"

On Error Resume Next
file$ = Dir$(ActPath$ & "*", vbDirectory)
Do While Len(file$)
    If file$ <> "." And file$ <> ".." Then
        Err = 0
        AttributiFile = GetAttr(ActPath$ & file$)
        If Err = 0 Then
            If (AttributiFile And vbDirectory) Then
                ElencoCartelle$ = ElencoCartelle$ & file$ & vbCrLf
            Else
                If (ElaboraFile(ActPath$ & file$) = False) Then
                    ScansioneCartelle = False
                    Exit Function
                End If
            End If
        End If
    End If
    file$ = Dir$
Loop
On Error GoTo 0

Dim Pos%
Do While Len(ElencoCartelle$)
    
    Pos% = InStr(ElencoCartelle$, vbCrLf)
    If Pos% Then
        file$ = ActPath$ & Left$(ElencoCartelle$, Pos% - 1)
        ElencoCartelle$ = Mid$(ElencoCartelle$, Pos% + Len(vbCrLf))
    Else
        file$ = ActPath$ & ElencoCartelle$
        ElencoCartelle$ = ""
    End If
    
    If (ScansioneCartelle(file$) = False) Then
        ScansioneCartelle = False
        Exit Function
    End If
Loop

ScansioneCartelle = True

End Function
Private Function TrovaFiles(ByVal DirSorgente As String, _
        ByVal MascheraFile As String, _
        ByRef NFiles As Long) As String()
Static VetNomi() As String
Dim Nome As String
Dim i As Long

i = 0
Nome = Dir(DirSorgente & "\" & MascheraFile)
While (Nome <> "")
    ReDim Preserve VetNomi(i)
    VetNomi(i) = DirSorgente & "\" & Nome
    i = i + 1
    Nome = Dir
Wend
NFiles = i
TrovaFiles = VetNomi

End Function


Private Function TrovaInfoUnita(Lettera As String, ByRef InfoData As Tipo_Drive) As Boolean
    
Dim fs, Unita
Dim SoloLettera As String
Dim MexUnita As String

MexUnita = Left(Lettera, 1) & ":\"

On Error Resume Next
Err.Clear
Set fs = CreateObject( _
        "Scripting.FileSystemObject")
Set Unita = fs.GetDrive(MexUnita)
If (Err.Number = 68) Then
    With InfoData
        .Lettera = ""
        .Nome = ""
        .Pronta = False
        .Seriale = 0
        .SpazioLibero = 0
        .SpazioTotale = 0
        .Tipo = 0
        .NomeCondiviso = ""
    End With
    TrovaInfoUnita = False
    Exit Function
End If

With InfoData
    .Lettera = Unita.driveletter
    .Pronta = Unita.isready
    .NomeCondiviso = Unita.ShareName
    
    If (.Pronta = True) Then
        .Nome = Unita.volumename
        .Seriale = Unita.serialnumber
        .SpazioLibero = Unita.freespace
        .SpazioTotale = Unita.TotalSize
    Else
        .Nome = ""
        .Seriale = 0
        .SpazioLibero = 0
        .SpazioTotale = 0
    End If
 
    .Tipo = Unita.drivetype
End With

If (Trim(InfoData.Lettera) = "") Then
    TrovaInfoUnita = False
Else
    TrovaInfoUnita = True
End If
End Function

Private Function TrovaNomeCorto(NomeLungo As String) As String

Dim NomeCorto As String

NomeCorto = String(256, " ")


Call GetShortPathName(NomeLungo, NomeCorto, 255)

TrovaNomeCorto = Trim(NomeCorto)
If (TrovaNomeCorto = "") Then
    TrovaNomeCorto = NomeLungo
Else

    TrovaNomeCorto = Left(TrovaNomeCorto, Len(TrovaNomeCorto) - 1)
End If

End Function

Private Function DeleteFile(NomeFile As String) As Boolean
On Error GoTo 0
On Error Resume Next
If (EsisteFile(NomeFile) = False) Then
    DeleteFile = True
    Exit Function
End If
Err.Clear
Call SetAttr(NomeFile, vbNormal Or vbArchive)
If (Err.Number <> 0) Then
    DeleteFile = False
    Exit Function
End If

Call Kill(NomeFile)
If (Err.Number <> 0) Then
    DeleteFile = False
Else
    DeleteFile = True
End If

End Function
Private Function PrendiLinea( _
        BufLinea As String, _
        ByRef Indice As Long) As String

Dim i As Long, j As Long
Dim Car As String
Dim Riga As String

Riga = ""

For j = Indice To Len(BufLinea)
    Car = Mid(BufLinea, j, 1)
    If (Car = Chr(10) Or Car = Chr(13)) Then Exit For
    Riga = Riga & Car
Next j
PrendiLinea = Riga
Indice = j

End Function




Private Function CaricaFile(NomeFile As String) As String

Dim MiaStringa As String
Dim Ret As Integer
Dim TotSize As Long
Dim NFile As Integer

On Error Resume Next
Err.Number = 0

If (EsisteFile(NomeFile) = False) Then
    CaricaFile = ""
    Exit Function
End If

TotSize = FileLen(NomeFile)
MiaStringa = String(TotSize, Chr(0))
NFile = FreeFile
Open NomeFile For Binary Access Read As NFile

Get NFile, 1, MiaStringa
If (Err.Number > 0) Then
    CaricaFile = ""
    Close NFile
    Exit Function
End If
Close NFile

CaricaFile = MiaStringa
End Function

Private Function GetInfoUnita(Lettera As String, _
    ByRef InfoData As Tipo_InfoUnita) As Boolean

Dim fs, Unita

On Error Resume Next
Err.Clear
Set fs = CreateObject( _
        "Scripting.FileSystemObject")
Set Unita = fs.GetDrive(Lettera)
If (Err.Number = 68) Then
    With InfoData
        .Lettera = ""
        .Nome = ""
        .Pronta = False
        .Seriale = 0
        .SpazioLibero = 0
        .SpazioTotale = 0
        .Tipo = 0
    End With
    GetInfoUnita = False
    Exit Function
End If

With InfoData
    .Lettera = Unita.driveletter
    .Pronta = Unita.isready
    If (.Pronta = True) Then
        .Nome = Unita.volumename
        .Seriale = Unita.serialnumber
        .SpazioLibero = Unita.freespace
        .SpazioTotale = Unita.TotalSize
    Else
        .Nome = ""
        .Seriale = 0
        .SpazioLibero = 0
        .SpazioTotale = 0
    End If
 
    .Tipo = Unita.drivetype
End With

If (Trim(InfoData.Lettera) = "") Then
    GetInfoUnita = False
Else
    GetInfoUnita = True
End If



End Function

Private Function TrovaUnitaHd() As String()

Static VetLettere() As String
Dim NLettere As Long
Dim i As Long
Dim InfoDati As Tipo_Drive

NLettere = 0
For i = Asc("C") To Asc("Z")
    If (TrovaInfoUnita(Chr(i), InfoDati) = True) Then
        If (InfoDati.Tipo = tu_Fissa) Then
            ReDim Preserve VetLettere(NLettere)
            VetLettere(NLettere) = Chr(i) & ":"
            NLettere = NLettere + 1
        End If
    End If
Next i
TrovaUnitaHd = VetLettere

End Function
Public Function TrovaUnitaCD() As String()
Static VetLettere() As String
Dim NLettere As Long
Dim i As Long, j As Long
Dim InfoDati As Tipo_Drive
Dim Scambio As String

NLettere = 0
For i = Asc("C") To Asc("Z")
    If (TrovaInfoUnita(Chr(i), InfoDati) = True) Then
        If (InfoDati.Tipo = tu_CD_ROM) Then
            ReDim Preserve VetLettere(NLettere)
            VetLettere(NLettere) = Chr(i) & ":"
            NLettere = NLettere + 1
        End If
    End If
Next i


TrovaUnitaCD = VetLettere

End Function
Private Function ElaboraFile(NomeFile As String) As Boolean



ReDim Preserve VetFiles(TotFiles)
VetFiles(TotFiles) = NomeFile
TotFiles = TotFiles + 1

ElaboraFile = True

End Function

Private Function EseguiCon_BAT(LineaShell As String, _
        InfoU As Tipo_InfoUtility) As Boolean

Dim RetVal As Long, TempoNow As Long, BufLinea As String
Dim TipoShow As Variant, NF As Long
Dim HandleProc As Long, StartTime As Long
Dim Result As Long
Dim NomeBat As String, DirCorta As String
Dim NomeOut As String, NomePIF As String
Dim Testo As String, MexRedirect As String
Dim i As Long, j As Long
Dim Flag As Long, OldLineaOut As String, OldLineaErr As String
Dim Linea As String, NomeErr As String
Dim LineaOut As String, LineaErr As String

On Error Resume Next

Err.Clear

DirCorta = TrovaNomeCorto(App.Path)
NomeBat = DirCorta & "\go.bat"
NomeOut = DirCorta & "\temp_out.txt"
NomeErr = DirCorta & "\temp_err.txt"
NomePIF = DirCorta & "\go.pif"

If (IsWindowsXP() = True) Then
    Testo = "@echo off" & vbCrLf & LineaShell & ">" & NomeOut & "  " & "2>" & NomeErr
Else
    Testo = "@echo off" & vbCrLf & LineaShell & ">" & NomeOut
    If (EsisteFile(NomePIF) = False) Then
        With SetMaster.Response
            .ErrorCode = cderr_NonTrovatoGo_PIF
            .LastGeneralError = "Nella cartella dell'applicazione manca il file 'go.pif' necessario per la chiusura dell'applicazione 'go.bat'"
            .LastSpecificError = ""
            .LastScreenLog = ""
        End With
        EseguiCon_BAT = False
        Exit Function
    End If
End If

NF = FreeFile
Open NomeBat For Output As #NF
Print #NF, Testo
Close #NF
Call DeleteFile(NomeOut)
Call DeleteFile(NomeErr)

On Error Resume Next
SetMaster.Interfaccia.TestStop = False

If (SetMaster.Interfaccia.TestNascondi = True) Then
    TipoShow = vbHide
Else
    TipoShow = vbNormalFocus
End If

Err.Clear

If (InfoU.TestLabel = True) Then
    With SetMaster.Interfaccia
        .LabelInfo.Caption = InfoU.MexOperazioneInCorso
        .LabelInfo.Refresh
    End With
End If

RetVal = Shell(NomeBat, TipoShow)

If (RetVal = 0 Or Err.Number <> 0) Then
    Call AzzeraInterfaccia
    With SetMaster.Response
        .ErrorCode = InfoU.CodiceGenerale
        .LastGeneralError = InfoU.GeneralError
        .LastSpecificError = "Errore cercando di eseguire la seguente linea di shell: " & vbCrLf & _
                LineaShell
    End With
    EseguiCon_BAT = False
    Exit Function
End If
If (IsWindowsXP = True) Then
    Flag = SYNCHRONIZE
Else
    Flag = PROCESS_TERMINATE
End If
OldLineaOut = ""
OldLineaErr = ""

HandleProc = OpenProcess(Flag, False, RetVal)
Do
    Result = WaitForSingleObject(HandleProc, 1000)
    Select Case Result
        Case WAIT_FAILED
            Call AzzeraInterfaccia
            With SetMaster.Response
                .ErrorCode = InfoU.CodiceGenerale
                .LastGeneralError = InfoU.GeneralError
                .LastSpecificError = "Errore eseguendo sotto-processo 'GO.BAT'"
                .LastScreenLog = ""
            End With
            EseguiCon_BAT = False
            Exit Function
        Case WAIT_ABANDONED
            Exit Do
        Case WAIT_TIMEOUT
            With SetMaster.Interfaccia
                LineaOut = PrendiUltimaLinea(NomeOut)
                LineaErr = PrendiUltimaLinea(NomeErr)
                
                BufLinea = ""
                
                If (LineaOut <> OldLineaOut) Then
                    BufLinea = LineaOut
                End If
                
                If (LineaErr <> OldLineaErr) Then
                    BufLinea = LineaErr
                End If
                
                If (BufLinea <> "") Then
                    InfoU.BufLinea = BufLinea
                    Call ScansioneLinea(InfoU)
                End If
                
                OldLineaOut = LineaOut
                OldLineaErr = LineaErr
                
                If (.ControllaSeStop = True) Then
                    DoEvents
                    If (.TestStop = True) Then
                        Call TerminateProcess(HandleProc, -1)
                        Call AzzeraInterfaccia
                        With SetMaster.Response
                            .LastGeneralError = InfoU.GeneralError
                            .ErrorCode = cderr_InterrottoDaUtente
                            .LastSpecificError = "Operazione interrotta dall'utente"
                        End With
                        EseguiCon_BAT = False
                        Exit Function
                    End If
                End If
                
            End With
    End Select
Loop While (Result <> WAIT_STATUS_0)
Call CloseHandle(HandleProc)

Call AzzeraInterfaccia
Testo = CaricaFile(NomeOut)
SetMaster.Response.LastScreenLog = Testo & vbCrLf & _
                    CaricaFile(NomeErr)

EseguiCon_BAT = True

End Function
Private Function PrendiUltimaLinea(NomeFile As String) As String

Dim BufLinea As String, Inizio As Long, Fine As Long
Dim Testo As String

Testo = CaricaFile(NomeFile)

If (Testo = "") Then
    PrendiUltimaLinea = ""
    Exit Function
End If

Fine = InStrRev(Testo, Chr(13), -1, vbBinaryCompare)

If (Fine < 0) Then
    PrendiUltimaLinea = ""
    Exit Function
End If

Inizio = InStrRev(Testo, Chr(13), Fine - 1, vbBinaryCompare)

If (Inizio < 0) Then
    Inizio = 1
Else
    Inizio = Inizio + 1
End If

BufLinea = Mid(Testo, Inizio, Fine - Inizio)
BufLinea = Replace(BufLinea, Chr(10), "", 1, -1, vbBinaryCompare)

PrendiUltimaLinea = BufLinea


End Function
Private Sub AttendiD(ByVal DecimiSecondo As Long)

Sleep (DecimiSecondo * 100)

End Sub


Private Function OpenCdDevice(Drive As String) As Long

On Error GoTo ErrHandler

Dim lCode As Long
Dim mciOpenParms As MCI_OPEN_PARMS
Dim sBuffer As String * 128

mciOpenParms.lpstrDeviceType = "cdaudio"
mciOpenParms.lpstrElementName = Left(Drive, 2)
lCode = mciSendCommand(0, MCI_OPEN, _
            (MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT Or MCI_OPEN_SHAREABLE), _
            mciOpenParms)
If lCode <> MMSYSERR_NOERROR Then GoTo ErrHandler

OpenCdDevice = mciOpenParms.wDeviceID
Exit Function

ErrHandler:
    OpenCdDevice = 0
End Function

Private Function CassettoCD(LetteraUnita As String, TestApri As Boolean) As Boolean

Dim Handle As Long
Dim Comando As Long

Handle = OpenCdDevice(LetteraUnita)
If (Handle = 0) Then
    With SetMaster.Response
        .ErrorCode = cderr_ArgomentoNonValido
        .LastGeneralError = "Unit letter not correct"
        .LastSpecificError = "Error opening device with unit letter: " & LetteraUnita
    End With
    CassettoCD = False
    Exit Function
End If

If (TestApri = True) Then
    Comando = MCI_SET_DOOR_OPEN
Else
    Comando = MCI_SET_DOOR_CLOSED
End If

If (mciSendCommand(Handle, MCI_SET, Comando, 0) <> _
            MMSYSERR_NOERROR) Then
    With SetMaster.Response
        .ErrorCode = cderr_ArgomentoNonValido
        .LastGeneralError = "Unit letter not correct"
        .LastSpecificError = "Error sending command to device with unit letter: " & LetteraUnita
    End With
    CassettoCD = False
    Exit Function
End If

mciSendCommand Handle, MCI_CLOSE, 0, 0
CassettoCD = True

End Function

Private Function TrovaIdDispositivi() As Boolean

Dim VetNomi() As String
Dim TotNomi As Long
Dim i As Long, j As Long, LetteraCd As String
Dim NomeExe As String, Log As String
Dim LineaShell As String
Dim VetLinee() As String
Dim TotLinee As Long, Linea As String
Dim NomeID As String, StrNumero1 As String
Dim z As Long, StrNumero2 As String
Dim Car As String, LineaComandi As String
Dim VetCD() As String, TotCD As Long
Dim NumeroBus As String, InfoU As Tipo_InfoUtility
Dim VetParti() As String, TestLabel As Boolean

On Error Resume Next
Err.Clear

SetMaster.Interfaccia.LabelInfo.Refresh
If (Err.Number = 0) Then
    TestLabel = True
Else
    TestLabel = False
End If

VetCD = TrovaUnitaCD()

TotCD = UBound(VetCD) + 1

NomeExe = App.Path & "\cdrecord.exe"
NomeExe = TrovaNomeCorto(NomeExe)
LineaComandi = NomeExe & " -scanbus "
With SetMaster.GlobArray
    .TotFiles = 0
End With
SetMaster.Response.ResultValue = 0
With InfoU
    .BufLinea = ""
    .CodiceGenerale = cderr_MasterizzatoreNonIdentificato
    .GeneralError = "Impossibile identificare masterizzatore"
    .LineaShell = LineaComandi
    .LogOut = ""
    .TestCdRecord = False
    .TestLabel = TestLabel
    .TestProgr = False
    .TestReport = False
    .TotMB = 0
    .UltimoComando = SetMaster.Response.LastCommand
End With

Call ImpostaOperazioneInCorso(InfoU)


If (EseguiCon_BAT(LineaComandi, InfoU) = False) Then
    TrovaIdDispositivi = False
    Exit Function
End If

Log = SetMaster.Response.LastScreenLog

If (Len(Log) < 10) Then
    With SetMaster.Response
        .ErrorCode = cderr_MasterizzatoreNonIdentificato
        .LastGeneralError = "Non e' stato possibile identificare l'unita' di masterizzazione"
        .LastScreenLog = ""
        .LastSpecificError = "Fallita operazione '-scanbus' con cdrecord"
    End With
    TrovaIdDispositivi = False
    Exit Function
End If

'l'output e' come il seguente:
'    0,3,0     3) *
'    0,4,0     4) *
'    0,5,0     5) *
'    0,6,0     6) *
'    0,7,0     7) HOST ADAPTOR
'scsibus1:
'    1,0,0   100) 'LG      ' 'CD-ROM CRD-8520B' '1.00' Removable CD-ROM

' Considerare solo linee che iniziano col
'carattere di tabulazione
Log = Replace(Log, Chr(13), "", 1, -1, vbBinaryCompare)

VetLinee = Split(Log, vbLf)
TotLinee = UBound(VetLinee) + 1
TotNomi = 0
For i = 0 To TotLinee - 1
    Linea = VetLinee(i)
    
    If (Left(Linea, 1) = vbTab) Then
        j = InStr(1, Linea, ")", vbBinaryCompare)
        If (j > 0) Then
            j = j + 2
            If (Mid(Linea, j, 1) <> "*" And _
                Mid(Linea, j, Len("HOST ADAPTOR")) <> _
                    "HOST ADAPTOR" And _
                    InStr(1, Linea, "NON CCS Disk", vbTextCompare) = 0) Then
                NomeID = Mid(Linea, j, Len(Linea))
                j = TotNomi
                If (j < TotCD) Then
                    LetteraCd = VetCD(j)
                Else
                    LetteraCd = "X:"
                End If
                NomeID = Replace(NomeID, "Removable CD-ROM", "")
                VetParti = Split(Linea, vbTab)
                Linea = VetParti(1)
                VetParti = Split(Linea, ",")
                Linea = VetParti(0) & "'" & VetParti(1) & "'" & _
                            VetParti(2) & "'" & LetteraCd & "'" & _
                            NomeID
                ReDim Preserve VetNomi(TotNomi)
                VetNomi(TotNomi) = Linea
                TotNomi = TotNomi + 1
            End If
        End If

    End If
Next i

With SetMaster
    .Response.ResultValue = TotNomi
    With .GlobArray
        .TotFiles = TotNomi
        ReDim .VetFiles(TotNomi)
        For i = 0 To TotNomi - 1
            .VetFiles(i) = VetNomi(i)
        Next i
    End With
End With
TrovaIdDispositivi = True
End Function
Private Function TrovaVersioneWindows() As Long

Dim Ret As Long
Dim OSVI As OSVERSIONINFO
Dim Versione As Long
Dim WinVer As Long

OSVI.dwOSVersionInfoSize = Len(OSVI)
Ret = GetVersionEx(OSVI)

If Ret <> 0 Then

   WinVer = OSVI.dwPlatformId

   Select Case WinVer
      Case 0
         Versione = 0

      Case 1
         If OSVI.dwMinorVersion < 10 Then
            Versione = 1
         Else
            Versione = 4
         End If

      Case 2
         If OSVI.dwMajorVersion = 5 Then
            Versione = 5
         Else
            Versione = 6
         End If

    End Select

End If
TrovaVersioneWindows = Versione
End Function
Private Function IsWindowsXP() As Boolean
Dim i As Long

i = TrovaVersioneWindows()
If (i > 4) Then
    IsWindowsXP = True
Else
    IsWindowsXP = False
End If
End Function
Private Function TrovaUnitaPiuSpazio(ByRef SpazioLibero As Double) As String
' ... Scan All Driver for more Space Free ;)
Dim VetUnita() As String, NUnita As Long
Dim VetSpazio() As Double, Lettera As String
Dim i As Long, MaxValore As Double
Dim j As Long, InfoUnita As Tipo_Drive

VetUnita = TrovaUnitaHd
NUnita = UBound(VetUnita) + 1
ReDim VetSpazio(NUnita)

For i = 0 To NUnita - 1
    Lettera = VetUnita(i)
    VetSpazio(i) = SpazioLiberoSulDisco(Lettera)
Next i

MaxValore = 0
j = 0
For i = 0 To NUnita - 1
    If (VetSpazio(i) > MaxValore) Then
        j = i
        MaxValore = VetSpazio(i)
    End If
Next i
SpazioLibero = MaxValore
TrovaUnitaPiuSpazio = VetUnita(j)

End Function

Private Function SpazioLiberoSulDisco(ByVal Drive As String) As Double
' .... This Function Return the Free Space > more to 2GB, in any case Call
' .... Call the --> Function TrovaUnitaPiuSpazio()
Dim lpSectorPerCluster As Long
Dim lpBytePerSector As Long
Dim lpNumberOfFreeClusters As Long
Dim lpTotalNumberOfClusters  As Long
Dim cBytesFreeToCaller As Currency
Dim cCapacitaDisco As Currency
Dim MexUnita As String
Dim ValoreSpazio As Currency

MexUnita = Left(Drive, 1) & ":\"

    On Local Error Resume Next
    If (GetDiskFreeSpaceEx(MexUnita, cBytesFreeToCaller, cCapacitaDisco, _
                ValoreSpazio) = 0) Then
                
        SpazioLiberoSulDisco = SpazioLiberoUnita(MexUnita)
        Exit Function
    End If
    
    SpazioLiberoSulDisco = ValoreSpazio
    If Err = 0 Then
        On Local Error GoTo 0
        SpazioLiberoSulDisco = SpazioLiberoSulDisco * 10000
    Else
        On Local Error GoTo 0
        SpazioLiberoSulDisco = SpazioLiberoUnita(MexUnita)
        
    End If
End Function

Private Function SpazioLiberoUnita(ByVal NomeUnita As String) As Double
Dim InfoData As Tipo_Drive
Dim Spazio As Double

If (TrovaInfoUnita(NomeUnita, InfoData) = False) Then
    SpazioLiberoUnita = 0
Else
    SpazioLiberoUnita = InfoData.SpazioLibero
End If


End Function


Private Sub Attendi(ByVal Secondi As Long)

Sleep (Secondi * 1000)

End Sub
Private Sub ImpostaOperazioneInCorso(ByRef InfoU As Tipo_InfoUtility)

Dim Testo As String

Testo = ""

With InfoU
    Select Case .UltimoComando
        Case Tipo_CmdWrite.cw_CalcolaDimensioneISO
            Testo = "Calcolo dimensione file nell'immagine ISO in corso ..."
        Case Tipo_CmdWrite.cw_CancellaCDRiscrivibile
            Testo = "Cancellazione CD riscrivibile in corso ..."
        Case Tipo_CmdWrite.cw_CreaISO
            Testo = "Creazione immagine ISO in corso ..."
        Case Tipo_CmdWrite.cw_GetIDDevices
            Testo = "Ricerca dispositivi sul bus SCSI in corso ..."
        Case Tipo_CmdWrite.cw_MasterizzaISO
            Testo = "Masterizzazione CD in corso ..."
        Case Tipo_CmdWrite.cw_MasterizzaSenzaISO
            Testo = "Masterizzazione CD senza immagine ISO in corso ..."
        Case Tipo_CmdWrite.cw_ReportISOFileList
            Testo = "Scansione file immagine ISO in corso ..."
        Case Tipo_CmdWrite.cw_VerificaCD
            Testo = "Verifica CD in corso ..."
    End Select
    
    .MexOperazioneInCorso = Testo
End With

End Sub

Private Function FormattaDimensione(ByVal Valore As Double) As String
Dim Testo As String
If (Valore > 1073741824) Then
    Valore = Valore / 1073741824
    Testo = Format(Valore, "###.0") & " Gb"
Else
    If (Valore > 1048576) Then
        Valore = Valore / 1048576
        Testo = Format(Valore, "###.0") & " Mb"
    Else
        If (Valore > 1024) Then
            Valore = Valore \ 1024
            Testo = CStr(Valore) & " Kb"
        Else
            Testo = CStr(Valore) & " bytes"
        End If
    End If
End If

FormattaDimensione = Testo

End Function
Private Sub ScansioneLinea(InfoUtility As Tipo_InfoUtility)

Dim TestIgnora As Boolean, n As Long
Dim Totale As Double, Parziale As Double
Dim i As Long, StrNumero As String
Dim LastPerc As Double, TempStr As String
Dim MessaggioOut As String
Dim DirCorta As String

DirCorta = TrovaNomeCorto(App.Path) & "\"

DirCorta = Right(DirCorta, Len(DirCorta) - 3)
DirCorta = UsaBarreUnix(DirCorta)

With InfoUtility
    .BufLinea = Replace(.BufLinea, DirCorta, "", 1, -1, vbTextCompare)
End With

MessaggioOut = ""
TestIgnora = False
With InfoUtility
    If (.TestCdRecord = True) Then
        n = Len("Last chance")
        If (Len(.BufLinea) < n) Then
            n = Len(.BufLinea)
        End If
        If (n > 0) Then
            If (Left("last chance", n) = _
                Left(.BufLinea, n)) Then TestIgnora = True
        End If
        If (InStr(1, .BufLinea, _
            "Warning: using inofficial", vbTextCompare) > 0) Then TestIgnora = True
    Else
        If (Left(.BufLinea, Len("Using")) = _
            "Using") Then TestIgnora = True
    End If

    If (.BufLinea <> "") Then
        If (.TestLabel = True And _
            TestIgnora = False) Then
            With SetMaster.Interfaccia.LabelInfo
                TempStr = InfoUtility.BufLinea
                If (InStr(1, TempStr, "Total bytes read/written", vbTextCompare) > 0) Then
                    TempStr = "Fixating, lead-in and lead-out ..."
                    InfoUtility.MexOperazioneInCorso = TempStr
                End If
                If (.Caption <> TempStr And TempStr <> "") Then
                    .Caption = TempStr
                    .Refresh
                End If
                MessaggioOut = TempStr
            End With
        End If
    End If
    If (.BufLinea <> "" And TestIgnora = False) Then
        .LogOut = .LogOut & .BufLinea & vbCrLf
    End If
    
    If (.TestProgr = True) Then
        If (.TestCdRecord = True) Then
            If (.UltimoComando = cw_CreaISO_MasterizzaISO Or _
                .UltimoComando = cw_MasterizzaISO Or .UltimoComando = cw_MasterizzaSenzaISO) Then
                If (Left(.BufLinea, 5) = "TRACK" And _
                    InStr(1, .BufLinea, "MB written", vbTextCompare) > 0) Then
                    i = Len("TRACK 01: ")
                    StrNumero = PrendiCifre(.BufLinea, i)
                    Parziale = Val(StrNumero)
                    If (.TotMB = 0) Then
                        i = i + 2
                        StrNumero = PrendiCifre(.BufLinea, i)
                        Totale = Val(StrNumero)
                    Else
                        Totale = .TotMB
                    End If
                    
                    With SetMaster.Interfaccia
                        Call MostraPercentuale(.ProgrBar, Totale, Parziale)
                    End With
                End If
            End If

        Else
            If (.TestReport = True) Then
                If (Totale <> 0) Then
                    If (Parziale > Totale) Then
                        Parziale = Totale
                    End If
                    Call MostraPercentuale(SetMaster.Interfaccia.ProgrBar, Totale, Parziale)
                End If
            Else
                If (InStr(1, .BufLinea, "done, estimate finish", vbTextCompare) > 0) Then
                    i = 1
                    StrNumero = PrendiCifre(.BufLinea, i)
                    Parziale = Val(StrNumero)
                    If (Parziale < LastPerc) Then
                        Parziale = 100
                    End If
                    Totale = 100
                    With SetMaster.Interfaccia
                        Call MostraPercentuale(.ProgrBar, Totale, Parziale)
                    End With
                    LastPerc = Parziale
                End If
            End If
        End If
    End If
    If (Trim(MessaggioOut) = "" And _
                    .TestLabel = True) Then
        With SetMaster.Interfaccia
            .LabelInfo.Caption = InfoUtility.MexOperazioneInCorso
            .LabelInfo.Refresh
        End With
    End If
    
End With

End Sub

Private Function RisolveLetteraCd(ByVal LetteraAttuale As String) As String

Dim Testo As String
With SetMaster.CdWriter
    If (.LetteraUnitaInput <> "") Then
        RisolveLetteraCd = .LetteraUnitaInput
        Exit Function
    End If
End With

Select Case Len(LetteraAttuale)
    Case 0:
        LetteraAttuale = "X:"
    Case 1:
        LetteraAttuale = LetteraAttuale & ":"
End Select


If (Left(LetteraAttuale, 1) <> "X") Then
    RisolveLetteraCd = LetteraAttuale
    Exit Function
End If

With SetMaster.CdWriter
    If (.LetteraUnitaInput <> "") Then
        RisolveLetteraCd = Left(.LetteraUnitaInput, 1) & _
                Right(LetteraAttuale, Len(LetteraAttuale) - 1)
        Exit Function
    End If
    
    Do
        Testo = InputBox("Non e' stato possibile identificare la lettera di unita' che corrisponde al masterizzatore." & vbCrLf & _
                "Inserisci la lettera di unita' del tuo masterizzatore (C, D, E ecc.)", _
                "Richiesta lettera di unita' del masterizzatore")
    Loop While (Testo = "")
    .LetteraUnitaInput = Left(Testo, 1) & ":"
    RisolveLetteraCd = Left(.LetteraUnitaInput, 1) & Right(LetteraAttuale, Len(LetteraAttuale) - 1)
End With

End Function
