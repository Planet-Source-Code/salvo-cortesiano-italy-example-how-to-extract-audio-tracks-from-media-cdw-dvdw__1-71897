VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ID3 tag info @ ID3.org

Public Enum GenreConstants
    Cover = -2
    Remix
    Blues
    ClassicRock
    Country
    Dance
    Disco
    Funk
    Grunge
    HipHop
    Jazz
    Metal
    NewAge
    Oldies
    OtherGenre
    Pop
    RNB
    Rap
    Reggae
    Rock
    Techno
    Industrial
    Alternative
    Ska
    DeathMetal
    Pranks
    Soundtrack
    EuroTechno
    Ambient
    TripHop
    Vocal
    JazzAndFunk
    Fusion
    Trance
    Classical
    Instrumental
    Acid
    House
    Game
    SoundClip
    Gospel
    Noise
    AlternRock
    Bass
    Soul
    Punk
    SpaceGenre
    Meditative
    InstrumentalPop
    InstrumentalRock
    Ethnic
    Gothic
    Darkwave
    TechnoIndustrial
    Electronic
    PopFolk
    Eurodance
    Dream
    SouthernRock
    Comedy
    Cult
    Gangsta
    Top40
    ChristianRap
    PopFunk
    Jungle
    NativeAmerican
    Cabaret
    NewWave
    Psychedelic
    Rave
    Showtunes
    Trailer
    LoFi
    Tribal
    AcidPunk
    AcidJazz
    Polka
    Retro
    Musical
    RockNRoll
    HardRock
    Folk
    FolkRock
    NationalFolk
    Swing
    FastFusion
    Bebob
    Latin
    Revival
    Celtic
    Bluegrass
    Avantgarde
    GothicRock
    ProgressiveRock
    PsychedelicRock
    SymphonicRock
    SlowRock
    BigBand
    Chorus
    EasyListening
    Acoustic
    Humor
    Speech
    Chanson
    Opera
    ChamberMusic
    Sonata
    Symphony
    BootyBass
    Primus
    PornGroove
    Satire
    SlowJam
    Club
    Tango
    Samba
    Folklore
    Ballad
    PowerBallad
    RhythmicSoul
    Freestyle
    Duet
    PunkRock
    DrumSolo
    ACappella
    EuroHouse
    DanceHall
    Goa
    DrumNBass
    ClubHouse
    Hardcore
    Terror
    Indie
    BritPop
    Negerpunk
    PolskPunk
    Beat
    ChristianGangstaRap
    HeavyMetal
    BlackMetal
    Crossover
    ContemporaryChristian
    ChristianRock
    Merengue
    Salsa
    ThrashMetal
    Anime
    JPop
    Synthpop
    Unknown = 255
End Enum

Public Enum MPEGVersionEnum
    MPEG25
    MPEGVersionReserved
    MPEG1
    MPEG2
End Enum

Public Enum ChannelModeEnum
    Stereo
    JointStereo
    DualChannel
    SingleChannel
End Enum

Public Enum EncodingEnum
    CBR
    VBR
End Enum

Public Enum LayerEnum
    LayerReserved
    LayerIII
    LayerII
    LayerI
End Enum

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

' Lyrics3 constant declarations
Private Const LYRICSBEGIN As String = "LYRICSBEGIN"
Private Const LYRICSEND As String = "LYRICSEND"
Private Const LYRICS200 As String = "LYRICS200"
Private Const IND0000200 As String = "IND0000200"

' FRAME VARIABLE DECLARATION 'Frame ID of ID3v2.3+ (Frame ID of ID3v2.0 and ID3v2.2, if available)
Dim m_Comments As String 'COMM (COM)
Dim m_Album As String 'TALB (TAL)
Dim m_BeatsPerMinute As String 'TBPM (TBP)
Dim m_Composer As String 'TCOM (TCM)
Dim m_GenreID As GenreConstants 'TCON (TCO)
Dim m_Genre As String 'TCON (TCO)
Dim m_Copyright As String 'TCOP (TCR)
Dim m_EncodedBy As String 'TENC (TEN)
Dim m_Lyricist As String 'TEXT (TXT)
Dim m_Title As String 'TIT2 (TT2)
Dim m_InitialKey As String 'TKEY (TKE)
Dim m_Languages As String 'TLAN (TLA)
Dim m_OriginalAlbum As String 'TOAL (TOT)
Dim m_OriginalFileName As String 'TOFN (TOF)
Dim m_OriginalLyricist As String 'TOLY (TOL)
Dim m_OriginalArtist As String 'TOPE (TOA)
Dim m_OriginalReleaseYear As String 'TORY (TOR) - TDOR in ID3v2.4+
Dim m_FileOwner As String 'TOWN
Dim m_Artist As String 'TPE1 (TP1)
Dim m_Band As String 'TPE2 (TP2)
Dim m_Conductor As String 'TPE3 (TP3)
Dim m_InterpretedBy As String 'TPE4 (TP4)
Dim m_DiscNumber As String 'TPOS (TPA)
Dim m_DiscsTotal As String 'TPOS (TPA)
Dim m_Publisher As String 'TPUB (TPB)
Dim m_TrackNumber As String 'TRCK (TRK)
Dim m_TracksTotal As String 'TRCK (TRK)
Dim m_InternetRadioStationName As String 'TRSN
Dim m_InternetRadioStationOwner As String 'TRSO
Dim m_ISRC As String 'TSRC (TRC)
Dim m_Year As String 'TYER (TYE) - TDRC in ID3v2.4+
Dim m_Lyrics As String 'USLT (ULT)
Dim m_CopyrightInfo As String 'WCOP (WCP)
Dim m_AudioURL As String 'WOAF (WAF)
Dim m_AudioSourceURL As String 'WOAS (WAS)
Dim m_InternetRadioURL As String 'WORS
Dim m_PaymentURL As String 'WPAY
Dim m_PublisherURL As String 'WPUB (WPB)

Dim m_FileName As String

Dim m_ID3v1Exists As Boolean
Dim m_ID3v2Exists As Boolean
Dim m_Lyrics3v1Exists As Boolean
Dim m_Lyrics3v2Exists As Boolean

Dim m_ID3Revision As Byte
Dim LyricsTimestampsExist As Boolean

Dim IncrementA As Long
Dim IncrementB As Long
Dim IncrementC As Long

Dim OldFrameNames(32) As String
Dim FrameNames(37) As String
Dim Lyrics3FrameNames(4) As String

Dim bytBytes(2) As Byte
Dim m_HeaderPos As Long
Dim m_AudioSize As Long
Dim m_VBR As Boolean

Dim m_Xing_FrameCount As Long
Dim m_Xing_FileLength As Long
Dim m_Xing_TOC() As Byte
Dim m_Xing_Quality As Long

Dim m_VBRI_NumEntries As Long
Dim m_VBRI_SizePerEntry As Long

Dim m_TagUnsynchronized As Boolean ' ID3v2 only

Public Function GenreName(ByVal GenreID As GenreConstants) As String
    Dim x As String
    
    Select Case GenreID
        Case Cover: x = "Cover"
        Case Remix: x = "Remix"
        Case Blues: x = "Blues"
        Case ClassicRock: x = "Classic Rock"
        Case Country: x = "Country"
        Case Dance: x = "Dance"
        Case Disco: x = "Disco"
        Case Funk: x = "Funk"
        Case Grunge: x = "Grunge"
        Case HipHop: x = "Hip-Hop"
        Case Jazz: x = "Jazz"
        Case Metal: x = "Metal"
        Case NewAge: x = "New Age"
        Case Oldies: x = "Oldies"
        Case OtherGenre: x = "Other"
        Case Pop: x = "Pop"
        Case RNB: x = "R&B"
        Case Rap: x = "Rap"
        Case Reggae: x = "Reggae"
        Case Rock: x = "Rock"
        Case Techno: x = "Techno"
        Case Industrial: x = "Industrial"
        Case Alternative: x = "Alternative"
        Case Ska: x = "Ska"
        Case DeathMetal: x = "Death Metal"
        Case Pranks: x = "Pranks"
        Case Soundtrack: x = "Soundtrack"
        Case EuroTechno: x = "Euro-Techno"
        Case Ambient: x = "Ambient"
        Case TripHop: x = "Trip-Hop"
        Case Vocal: x = "Vocal"
        Case JazzAndFunk: x = "Jazz+Funk"
        Case Fusion: x = "Fusion"
        Case Trance: x = "Trance"
        Case Classical: x = "Classical"
        Case Instrumental: x = "Instrumental"
        Case Acid: x = "Acid"
        Case House: x = "House"
        Case Game: x = "Game"
        Case SoundClip: x = "Sound Clip"
        Case Gospel: x = "Gospel"
        Case Noise: x = "Noise"
        Case AlternRock: x = "AlternRock"
        Case Bass: x = "Bass"
        Case Soul: x = "Soul"
        Case Punk: x = "Punk"
        Case SpaceGenre: x = "Space"
        Case Meditative: x = "Meditative"
        Case InstrumentalPop: x = "Instrumental Pop"
        Case InstrumentalRock: x = "Instrumental Rock"
        Case Ethnic: x = "Ethnic"
        Case Gothic: x = "Gothic"
        Case Darkwave: x = "Darkwave"
        Case TechnoIndustrial: x = "Techno-Industrial"
        Case Electronic: x = "Electronic"
        Case PopFolk: x = "Pop-Folk"
        Case Eurodance: x = "Eurodance"
        Case Dream: x = "Dream"
        Case SouthernRock: x = "Southern Rock"
        Case Comedy: x = "Comedy"
        Case Cult: x = "Cult"
        Case Gangsta: x = "Gangsta"
        Case Top40: x = "Top 40"
        Case ChristianRap: x = "Christian Rap"
        Case PopFunk: x = "Pop/Funk"
        Case Jungle: x = "Jungle"
        Case NativeAmerican: x = "Native American"
        Case Cabaret: x = "Cabaret"
        Case NewWave: x = "New Wave"
        Case Psychedelic: x = "Psychedelic"
        Case Rave: x = "Rave"
        Case Showtunes: x = "Showtunes"
        Case Trailer: x = "Trailer"
        Case LoFi: x = "Lo-Fi"
        Case Tribal: x = "Tribal"
        Case AcidPunk: x = "Acid Punk"
        Case AcidJazz: x = "Acid Jazz"
        Case Polka: x = "Polka"
        Case Retro: x = "Retro"
        Case Musical: x = "Musical"
        Case RockNRoll: x = "Rock & Roll"
        Case HardRock: x = "Hard Rock"
        Case Folk: x = "Folk"
        Case FolkRock: x = "Folk-Rock"
        Case NationalFolk: x = "National Folk"
        Case Swing: x = "Swing"
        Case FastFusion: x = "Fast Fusion"
        Case Bebob: x = "Bebob"
        Case Latin: x = "Latin"
        Case Revival: x = "Revival"
        Case Celtic: x = "Celtic"
        Case Bluegrass: x = "Bluegrass"
        Case Avantgarde: x = "Avantgarde"
        Case GothicRock: x = "Gothic Rock"
        Case ProgressiveRock: x = "Progressive Rock"
        Case PsychedelicRock: x = "Psychedelic Rock"
        Case SymphonicRock: x = "Symphonic Rock"
        Case SlowRock: x = "Slow Rock"
        Case BigBand: x = "Big Band"
        Case Chorus: x = "Chorus"
        Case EasyListening: x = "Easy Listening"
        Case Acoustic: x = "Acoustic"
        Case Humor: x = "Humor"
        Case Speech: x = "Speech"
        Case Chanson: x = "Chanson"
        Case Opera: x = "Opera"
        Case ChamberMusic: x = "Chamber Music"
        Case Sonata: x = "Sonata"
        Case Symphony: x = "Symphony"
        Case BootyBass: x = "Booty Bass"
        Case Primus: x = "Primus"
        Case PornGroove: x = "Porn Groove"
        Case Satire: x = "Satire"
        Case SlowJam: x = "Slow Jam"
        Case Club: x = "Club"
        Case Tango: x = "Tango"
        Case Samba: x = "Samba"
        Case Folklore: x = "Folklore"
        Case Ballad: x = "Ballad"
        Case PowerBallad: x = "Power Ballad"
        Case RhythmicSoul: x = "Rhythmic Soul"
        Case Freestyle: x = "Freestyle"
        Case Duet: x = "Duet"
        Case PunkRock: x = "Punk Rock"
        Case DrumSolo: x = "Drum Solo"
        Case ACappella: x = "A Cappella"
        Case EuroHouse: x = "Euro-House"
        Case DanceHall: x = "Dance Hall"
        Case Goa: x = "Goa"
        Case DrumNBass: x = "Drum & Bass"
        Case ClubHouse: x = "Club-House"
        Case Hardcore: x = "Hardcore"
        Case Terror: x = "Terror"
        Case Indie: x = "Indie"
        Case BritPop: x = "BritPop"
        Case Negerpunk: x = "Negerpunk"
        Case PolskPunk: x = "Polsk Punk"
        Case Beat: x = "Beat"
        Case ChristianGangstaRap: x = "Christian Gangsta Rap"
        Case HeavyMetal: x = "Heavy Metal"
        Case BlackMetal: x = "Black Metal"
        Case Crossover: x = "Crossover"
        Case ContemporaryChristian: x = "Contemporary Christian"
        Case ChristianRock: x = "Christian Rock"
        Case Merengue: x = "Merengue"
        Case Salsa: x = "Salsa"
        Case ThrashMetal: x = "Thrash Metal"
        Case Anime: x = "Anime"
        Case JPop: x = "JPop"
        Case Synthpop: x = "Synthpop"
        Case Unknown: x = "Unknown"
    End Select
    
    GenreName = x
End Function

Private Function ConvertString(ByVal str As String) As String
    ConvertString = LCase$(Replace(Replace(Replace(Replace(Replace(Replace(str, "/", ""), " and ", ""), "&", ""), "-", ""), "+", ""), " ", ""))
End Function

Public Function ToGenreID(ByVal Genre As String) As GenreConstants
    Dim i As Long
    
    Dim bSuccess As Boolean
    Dim x As String
    Dim Y As String
    
    If Trim$(Genre) = "" Then
        ToGenreID = Unknown
    Else
        bSuccess = False
        For i = -2 To 255
            x = ConvertString(Genre)
            Y = ConvertString(GenreName(i))
            
            If x = Y Then
                ToGenreID = i
                bSuccess = True
                Exit For
            End If
            
            If i >= 147 And i <= 253 Then i = 254
        Next
        
        If Not bSuccess Then ToGenreID = OtherGenre
    End If
End Function

Public Property Get Layer() As LayerEnum
    Select Case CLng(bytBytes(0)) \ 2 Mod 4 ' xxxxx00x xxxxxxxx xxxxxxxx
        Case 1 ' 01
            Layer = LayerIII
        Case 2 ' 10
            Layer = LayerII
        Case 3 ' 11
            Layer = LayerI
        Case 0 ' 00
            Layer = LayerReserved
    End Select
End Property

Public Property Get Padding() As Long
    If CLng(bytBytes(1)) \ 2 Mod 2 = 1 Then ' xxxxxxxx xxxxxx0x xxxxxxxx
        If Layer = LayerI Then
            Padding = 4
        Else
            Padding = 1
        End If
    Else
        Padding = 0
    End If
End Property

Public Property Get MPEGVersion() As MPEGVersionEnum
    Select Case CLng(bytBytes(0)) \ 8 Mod 4 ' xxx00xxx xxxxxxxx xxxxxxxx
        Case 3 ' 11
            MPEGVersion = MPEG1
        Case 2 ' 10
            MPEGVersion = MPEG2
        Case 0 ' 00
            MPEGVersion = MPEG25
        Case 1 ' 01
            MPEGVersion = MPEGVersionReserved
    End Select
End Property

Public Property Get ChannelMode() As ChannelModeEnum
    Select Case CLng(bytBytes(2)) \ 64 ' xxxxxxxx xxxxxxxx 00xxxxxx
        Case 0 ' 00
            ChannelMode = Stereo
        Case 1 ' 01
            ChannelMode = JointStereo
        Case 2 ' 10
            ChannelMode = DualChannel
        Case 3 ' 11
            ChannelMode = SingleChannel
    End Select
End Property

Public Property Get Encoding() As EncodingEnum
    Encoding = IIf(m_VBR, VBR, CBR)
End Property

Public Property Get SamplingRateFrequency() As Long
    Select Case CLng(bytBytes(1)) \ 4 Mod 4 ' xxxxxxxx xxxx00xx xxxxxxxx
        Case 0 ' 00
            Select Case MPEGVersion
                Case MPEG1
                    SamplingRateFrequency = 44100
                Case MPEG2
                    SamplingRateFrequency = 22050
                Case MPEG25
                    SamplingRateFrequency = 11025
                Case MPEGVersionReserved
                    SamplingRateFrequency = -1
            End Select
        Case 1 ' 01
            Select Case MPEGVersion
                Case MPEG1
                    SamplingRateFrequency = 48000
                Case MPEG2
                    SamplingRateFrequency = 24000
                Case MPEG25
                    SamplingRateFrequency = 12000
                Case MPEGVersionReserved
                    SamplingRateFrequency = -1
            End Select
        Case 2 ' 10
            Select Case MPEGVersion
                Case MPEG1
                    SamplingRateFrequency = 32000
                Case MPEG2
                    SamplingRateFrequency = 16000
                Case MPEG25
                    SamplingRateFrequency = 8000
                Case MPEGVersionReserved
                    SamplingRateFrequency = -1
            End Select
        Case 3 ' 11
            SamplingRateFrequency = -1
    End Select
End Property

Public Property Get BitRate() As Double
    If Encoding = CBR Then
        Dim x As Long: x = CLng(bytBytes(1)) \ 16 ' xxxxxxxx 0000xxxx xxxxxxxx (the resulting value ranges from 0 (2^0 - 1) to 15 (2^4 - 1))
        If MPEGVersion = MPEG1 And Layer = LayerI Then
            BitRate = Choose(x + 1, 0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448, -1) * 1000
        End If
        If MPEGVersion = MPEG1 And Layer = LayerII Then
            BitRate = Choose(x + 1, 0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384, -1) * 1000
        End If
        If MPEGVersion = MPEG1 And Layer = LayerIII Then
            BitRate = Choose(x + 1, 0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, -1) * 1000
        End If
        If (MPEGVersion = MPEG2 Or MPEGVersion = MPEG25) And Layer = LayerI Then
            BitRate = Choose(x + 1, 0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256, -1) * 1000
        End If
        If (MPEGVersion = MPEG2 Or MPEGVersion = MPEG25) And (Layer = LayerII Or Layer = LayerIII) Then
            BitRate = Choose(x + 1, 0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160, -1) * 1000
        End If
    Else
        On Error Resume Next
        If FrameSize = 0 And length > 0 Then
            BitRate = CDbl(m_AudioSize) / length * 8
        Else
            BitRate = (FrameSize * CDbl(SamplingRateFrequency)) / 144
        End If
    End If
End Property

Public Property Get FrameSize() As Double
    If Encoding = CBR Then
        FrameSize = ((CDbl(FrameSamples) / 8 * BitRate) / CDbl(SamplingRateFrequency)) + Padding
    Else
        FrameSize = CDbl(m_Xing_FileLength) / CDbl(m_Xing_FrameCount)
    End If
End Property

Public Property Get FrameSamples() As Long
    Select Case Layer
        Case LayerI
            FrameSamples = 384
        Case LayerII
            FrameSamples = 1152
        Case LayerIII
            If MPEGVersion = MPEG1 Then
                FrameSamples = 1152
            Else
                FrameSamples = 576
            End If
    End Select
End Property

Public Property Get length() As Double
    On Error Resume Next
    If Encoding = CBR Then
        length = CDbl(m_AudioSize) / BitRate * 8
    Else
        length = CDbl(m_Xing_FrameCount) * CDbl(FrameSamples) / CDbl(SamplingRateFrequency)
    End If
End Property

Public Property Get Comments() As String
    Comments = m_Comments
End Property

Public Property Let Comments(ByVal Value As String)
    m_Comments = Value
End Property

Public Property Get Album() As String
    Album = m_Album
End Property

Public Property Let Album(ByVal Value As String)
    m_Album = Value
End Property

Public Property Get BeatsPerMinute() As String
    BeatsPerMinute = m_BeatsPerMinute
End Property

Public Property Let BeatsPerMinute(ByVal Value As String)
    m_BeatsPerMinute = Value
End Property

Public Property Get Composer() As String
    Composer = m_Composer
End Property

Public Property Let Composer(ByVal Value As String)
    m_Composer = Value
End Property

Public Property Get GenreID() As GenreConstants
    GenreID = m_GenreID
End Property

Public Property Let GenreID(ByVal Value As GenreConstants)
    m_GenreID = Value
End Property

Public Property Get Genre() As String
    Genre = m_Genre
End Property

Public Property Let Genre(ByVal Value As String)
    m_Genre = Value
End Property

Public Property Get Copyright() As String
    Copyright = m_Copyright
End Property

Public Property Let Copyright(ByVal Value As String)
    m_Copyright = Value
End Property

Public Property Get EncodedBy() As String
    EncodedBy = m_EncodedBy
End Property

Public Property Let EncodedBy(ByVal Value As String)
    m_EncodedBy = Value
End Property

Public Property Get Lyricist() As String
    Lyricist = m_Lyricist
End Property

Public Property Let Lyricist(ByVal Value As String)
    m_Lyricist = Value
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal Value As String)
    m_Title = Value
End Property

Public Property Get InitialKey() As String
    InitialKey = m_InitialKey
End Property

Public Property Let InitialKey(ByVal Value As String)
    m_InitialKey = Value
End Property

Public Property Get Languages() As String
    Languages = m_Languages
End Property

Public Property Let Languages(ByVal Value As String)
    m_Languages = Value
End Property

Public Property Get OriginalAlbum() As String
    OriginalAlbum = m_OriginalAlbum
End Property

Public Property Let OriginalAlbum(ByVal Value As String)
    m_OriginalAlbum = Value
End Property

Public Property Get OriginalFileName() As String
    OriginalFileName = m_OriginalFileName
End Property

Public Property Let OriginalFileName(ByVal Value As String)
    m_OriginalFileName = Value
End Property

Public Property Get OriginalLyricist() As String
    OriginalLyricist = m_OriginalLyricist
End Property

Public Property Let OriginalLyricist(ByVal Value As String)
    m_OriginalLyricist = Value
End Property

Public Property Get OriginalArtist() As String
    OriginalArtist = m_OriginalArtist
End Property

Public Property Let OriginalArtist(ByVal Value As String)
    m_OriginalArtist = Value
End Property

Public Property Get OriginalReleaseYear() As String
    OriginalReleaseYear = m_OriginalReleaseYear
End Property

Public Property Let OriginalReleaseYear(ByVal Value As String)
    m_OriginalReleaseYear = Value
End Property

Public Property Get FileOwner() As String
    FileOwner = m_FileOwner
End Property

Public Property Let FileOwner(ByVal Value As String)
    m_FileOwner = Value
End Property

Public Property Get Artist() As String
    Artist = m_Artist
End Property

Public Property Let Artist(ByVal Value As String)
    m_Artist = Value
End Property

Public Property Get Band() As String
    Band = m_Band
End Property

Public Property Let Band(ByVal Value As String)
    m_Band = Value
End Property

Public Property Get Conductor() As String
    Conductor = m_Conductor
End Property

Public Property Let Conductor(ByVal Value As String)
    m_Conductor = Value
End Property

Public Property Get InterpretedBy() As String
    InterpretedBy = m_InterpretedBy
End Property

Public Property Let InterpretedBy(ByVal Value As String)
    m_InterpretedBy = Value
End Property

Public Property Get DiscNumber() As String
    DiscNumber = m_DiscNumber
End Property

Public Property Let DiscNumber(ByVal Value As String)
    m_DiscNumber = Value
End Property

Public Property Get DiscsTotal() As String
    DiscsTotal = m_DiscsTotal
End Property

Public Property Let DiscsTotal(ByVal Value As String)
    m_DiscsTotal = Value
End Property

Public Property Get Publisher() As String
    Publisher = m_Publisher
End Property

Public Property Let Publisher(ByVal Value As String)
    m_Publisher = Value
End Property

Public Property Get TrackNumber() As String
    TrackNumber = m_TrackNumber
End Property

Public Property Let TrackNumber(ByVal Value As String)
    m_TrackNumber = Value
End Property

Public Property Get TracksTotal() As String
    TracksTotal = m_TracksTotal
End Property

Public Property Let TracksTotal(ByVal Value As String)
    m_TracksTotal = Value
End Property

Public Property Get InternetRadioStationName() As String
    InternetRadioStationName = m_InternetRadioStationName
End Property

Public Property Let InternetRadioStationName(ByVal Value As String)
    m_InternetRadioStationName = Value
End Property

Public Property Get InternetRadioStationOwner() As String
    InternetRadioStationOwner = m_InternetRadioStationOwner
End Property

Public Property Let InternetRadioStationOwner(ByVal Value As String)
    m_InternetRadioStationOwner = Value
End Property

Public Property Get ISRC() As String
    ISRC = m_ISRC
End Property

Public Property Let ISRC(ByVal Value As String)
    m_ISRC = Value
End Property

Public Property Get Year() As String
    Year = m_Year
End Property

Public Property Let Year(ByVal Value As String)
    m_Year = Value
End Property

Public Property Get Lyrics() As String
    Lyrics = m_Lyrics
End Property

Public Property Let Lyrics(ByVal Value As String)
    m_Lyrics = Value
End Property
Public Property Get CopyrightInfo() As String
    CopyrightInfo = m_CopyrightInfo
End Property

Public Property Let CopyrightInfo(ByVal Value As String)
    m_CopyrightInfo = Value
End Property

Public Property Get AudioURL() As String
    AudioURL = m_AudioURL
End Property

Public Property Let AudioURL(ByVal Value As String)
    m_AudioURL = Value
End Property

Public Property Get AudioSourceURL() As String
    AudioSourceURL = m_AudioSourceURL
End Property

Public Property Let AudioSourceURL(ByVal Value As String)
    m_AudioSourceURL = Value
End Property

Public Property Get InternetRadioURL() As String
    InternetRadioURL = m_InternetRadioURL
End Property

Public Property Let InternetRadioURL(ByVal Value As String)
    m_InternetRadioURL = Value
End Property

Public Property Get PaymentURL() As String
    PaymentURL = m_PaymentURL
End Property

Public Property Let PaymentURL(ByVal Value As String)
    m_PaymentURL = Value
End Property

Public Property Get PublisherURL() As String
    PublisherURL = m_PublisherURL
End Property

Public Property Let PublisherURL(ByVal Value As String)
    m_PublisherURL = Value
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal Value As String)
    m_FileName = Value
    GetID3Tags
End Property

Public Property Get ID3v1Exists() As Boolean
    ID3v1Exists = m_ID3v1Exists
End Property

Public Property Get ID3v2Exists() As Boolean
    ID3v2Exists = m_ID3v2Exists
End Property

Public Property Get Lyrics3v1Exists() As Boolean
    Lyrics3v1Exists = m_Lyrics3v1Exists
End Property

Public Property Get Lyrics3v2Exists() As Boolean
    Lyrics3v2Exists = m_Lyrics3v2Exists
End Property

Public Property Get ID3RevisionV2() As Byte
    ID3RevisionV2 = m_ID3Revision
End Property

Public Sub RefreshID3Tags()
    GetID3Tags
End Sub

Private Sub GetID3Tags()
    On Error Resume Next
    
    Dim IDTest As String
    Dim bCompression As Boolean
    Dim bExtendedHeader As Boolean
    Dim bFooterPresent As Boolean
    Dim TagSize As Double
    Dim Ridiculous As Boolean
    Dim TempData As String
    Dim TempData2 As String
    Dim f As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim s As String
    Dim PStart As Long
    Dim PEnd As Long
    Dim sGenreID As String
    Dim tGenre As String
    Dim tPartOfASet() As String
    Dim tTrack() As String
    Dim temp As String
    Dim arrTemp() As String
    
    Dim lLOF As Long
    Dim ReadPos As Long
    Dim ReadByte As Byte
    Dim CurrPos As Long
    Dim PrevPos As Long
    Dim bytXingBytes(3) As Byte
    Dim bytXingFlags(3) As Byte
    Dim bytTemp(3) As Byte
    Dim bytTemp2(1) As Byte
    Dim lngIndex As Long
    Dim lngOffset As Long
    Dim blnReturn As Boolean
    Dim lngXingOffset As Long
    Dim bytFlags(3) As Byte
    Dim bytFrameCount(3) As Byte
    Dim bytFileLength(3) As Byte
    Dim bytQuality(3) As Byte
    Dim bytTOC(99) As Byte
    Dim bytXingHeader(116) As Byte
    Dim bExit As Boolean
    Dim lXingSize As Long
    Dim bIntensityStereo As Boolean
    Dim Lyrics3Length As Long
    
    m_Comments = ""
    m_Album = ""
    m_BeatsPerMinute = ""
    m_Composer = ""
    m_GenreID = Unknown
    m_Genre = ""
    m_Copyright = ""
    m_EncodedBy = ""
    m_Lyricist = ""
    m_Title = ""
    m_InitialKey = ""
    m_Languages = ""
    m_OriginalAlbum = ""
    m_OriginalFileName = ""
    m_OriginalLyricist = ""
    m_OriginalArtist = ""
    m_OriginalReleaseYear = ""
    m_FileOwner = ""
    m_Artist = ""
    m_Band = ""
    m_Conductor = ""
    m_InterpretedBy = ""
    m_DiscNumber = ""
    m_DiscsTotal = ""
    m_Publisher = ""
    m_TrackNumber = ""
    m_TracksTotal = ""
    m_InternetRadioStationName = ""
    m_InternetRadioStationOwner = ""
    m_ISRC = ""
    m_Year = ""
    m_Lyrics = ""
    m_CopyrightInfo = ""
    m_AudioURL = ""
    m_AudioSourceURL = ""
    m_InternetRadioURL = ""
    m_PaymentURL = ""
    m_PublisherURL = ""
    
    m_ID3Revision = 0

    m_ID3v1Exists = False
    m_ID3v2Exists = False
    m_Lyrics3v1Exists = False
    m_Lyrics3v2Exists = False
    LyricsTimestampsExist = False
    
    bytBytes(0) = 0
    bytBytes(1) = 0
    bytBytes(2) = 0
    m_HeaderPos = 0
    m_AudioSize = 0
    m_VBR = False
    bIntensityStereo = False
    
    m_Xing_FileLength = 0
    m_Xing_FrameCount = 0
    m_Xing_Quality = 0
    Erase m_Xing_TOC
    
    m_VBRI_NumEntries = 0
    m_VBRI_SizePerEntry = 0
    
    m_TagUnsynchronized = False
    bCompression = False
    bExtendedHeader = False
    bFooterPresent = False
    
    f = FreeFile
    
    Open m_FileName For Binary Access Read Shared As #f
    If Err Then m_ID3Revision = 3: Exit Sub
    
        ' ID3v2
                
        IDTest = Space$(10) ' Give 10 blank bytes to the variable
        Get #f, 1, IDTest
        
        If Left$(IDTest, 3) = "ID3" And _
          Asc(Mid$(IDTest, 4, 1)) < &HFF& And _
          Asc(Mid$(IDTest, 5, 1)) < &HFF& And _
          Asc(Mid$(IDTest, 7, 1)) < &H80& And _
          Asc(Mid$(IDTest, 8, 1)) < &H80& And _
          Asc(Mid$(IDTest, 9, 1)) < &H80& And _
          Asc(Mid$(IDTest, 10, 1)) < &H80& Then ' ID3 header found
            
            m_ID3v2Exists = True
            
            m_ID3Revision = Asc(Mid$(IDTest, 4, 1))
            GetTagHeaderFlags IDTest, m_TagUnsynchronized, bCompression, bExtendedHeader, bFooterPresent
            If Not bCompression Then ' Do not read the frames if they are all compressed (ID3v2.0 and ID3v2.2 only)
                TagSize = StringToSize(Right$(IDTest, 4))
                TempData = Space$(CLng(TagSize)) ' Prepare string for the ID3 tag
                Get #f, 11, TempData
                
                If m_TagUnsynchronized Then ' Undo unsynchronization if necessary
                    If m_ID3Revision < 4 Then TempData = DeUnsynchronize(TempData)
                End If
                
                If bExtendedHeader Then ' Omit the extended header if detected
                    TempData = Mid$(TempData, StringToSize(Left$(TempData, 4), (m_ID3Revision >= 4)) + 5)
                End If
                
                i = 1
                Do
                    If i > Len(TempData) Then Exit Do ' No more data; therefore, exit
                    temp = Mid$(TempData, i, FrameNameSizeID3v2)
                    If (m_ID3Revision <= 2 And temp = String$(3, 0)) Or (m_ID3Revision > 2 And temp = String$(4, 0)) Or IsPaddingCorruptedByMP3ext(TempData, i) Then Exit Do
                    ReadTag temp, TempData, i, tGenre, tPartOfASet, tTrack
                    i = i + FinalOffsetID3v2(TempData, i, Ridiculous)
                    If Ridiculous Then Exit Do
                Loop
                
                PStart = InStr(tGenre, "(")
                If PStart = 1 Then
                    PEnd = InStr(PStart + 1, tGenre, ")")
                    If PEnd > 0 Then
                        sGenreID = Mid$(tGenre, PStart + 1, PEnd - PStart - 1)
                        m_Genre = Mid$(tGenre, PEnd + 1)
                        
                        If IsNumeric(sGenreID) Then
                            m_GenreID = CLng(Mid$(tGenre, PStart + 1, PEnd - PStart - 1))
                        Else
                            m_GenreID = Unknown
                            Select Case sGenreID
                                Case "CR": m_GenreID = Cover
                                Case "RX": m_GenreID = Remix
                            End Select
                        End If
                    Else
                        m_Genre = tGenre
                    End If
                Else
                    m_Genre = tGenre
                End If
                
                If IsValidArray(tTrack) Then
                    m_TrackNumber = ValidateNum(tTrack(0))
                    If UBound(tTrack) = 1 Then _
                        m_TracksTotal = ValidateNum(tTrack(1))
                End If
                
                If IsValidArray(tPartOfASet) Then
                    m_DiscNumber = ValidateNum(tPartOfASet(0))
                    If UBound(tPartOfASet) = 1 Then _
                        m_DiscsTotal = ValidateNum(tPartOfASet(1))
                End If
            End If
        Else
            m_ID3Revision = 3
        End If
        
        ' ID3v1
        
        TempData2 = Space$(128) ' Give 128 blank bytes to the variable
        Get #f, LOF(f) - 127, TempData2
        
        lLOF = LOF(f)
        If Left$(TempData2, 3) = "TAG" Then
            m_ID3v1Exists = True
            lLOF = lLOF - 128
        End If
        
        ' Lyrics3
        
        TempData = Space$(9)
        Get #f, lLOF - 8, TempData
        
        If TempData = LYRICS200 Then
            m_Lyrics3v2Exists = True
            
            TempData = Space$(6)
            Get #f, lLOF - 14, TempData
            Lyrics3Length = CLng(TempData)
            
            TempData = Space$(Lyrics3Length)
            Get #f, lLOF - 14 - Lyrics3Length, TempData
            
            If Left$(TempData, 11) = LYRICSBEGIN Then
                lLOF = lLOF - 15 - Lyrics3Length
                
                i = 12
                Do
                    If i > Len(TempData) Then Exit Do ' No more data; therefore, exit
                    temp = Mid$(TempData, i, 3)
                    ReadLyrics3Tag temp, TempData, i
                    s = Mid$(TempData, i + 3, 5)
                    If Not ProperSize(s) Then Exit Do ' The Lyrics3 tag may be corrupt
                    i = i + 8 + CLng(s)
                Loop
            End If
        ElseIf TempData = LYRICSEND Then
            m_Lyrics3v1Exists = True
            
            If lLOF - 20 > 5100 Then
                TempData = Space$(5100)
                Get #f, lLOF - 5119, TempData
            Else
                TempData = Space$(lLOF - 20)
                Get #f, 1, TempData
            End If
            
            i = InStrRev(TempData, LYRICSBEGIN)
            If i > 0 Then
                If m_Lyrics = "" Then m_Lyrics = Mid$(TempData, i + 11)
            End If
        End If
        
        ' ID3v1
        
        If m_ID3v1Exists Then
            StoreBytes TempData2, 4, 30, m_Title
            StoreBytes TempData2, 34, 30, m_Artist
            StoreBytes TempData2, 64, 30, m_Album
            StoreBytes TempData2, 94, 4, m_Year
            
            i = CLng(Asc(Mid$(TempData2, 126, 1))) * 256 + _
                CLng(Asc(Mid$(TempData2, 127, 1)))
            
            If i > 0 And i <= 255 Then
                TStoreByte TempData2, m_TrackNumber
                StoreBytes TempData2, 98, 28, m_Comments
            Else
                StoreBytes TempData2, 98, 30, m_Comments
            End If
            
            GStoreByte TempData2, m_GenreID
        End If
        
        ' Bit rate and song length
        
        ' To save time and increase accuracy...
        If m_ID3v2Exists Then
            ReadPos = 11 + TagSize
            If bFooterPresent Then ReadPos = ReadPos + 10
        Else
            ReadPos = 1
        End If
        
        blnReturn = False
        
        Dim counter As Long
        counter = ReadPos
        For i = counter To lLOF Step 20480 ' Analyze the MP3 files in buffers of 20480 bytes (20 KB)
            If lLOF - i + 1 > 20480 Then
                TempData = Space$(20480)
            Else
                TempData = Space$(lLOF - i + 1)
            End If
            
            Get #f, i, TempData
            
            j = 0
            Do
                j = InStr(j + 1, TempData, Chr$(255))
                If j > 0 Then ' 11111111 detected
                    
                    k = 1: ReadPos = i + j
                    
                    If lLOF - ReadPos + 1 > 156 Then
                        TempData2 = Space$(156)
                    Else
                        TempData2 = Space$(lLOF - ReadPos + 1)
                    End If
                    
                    Get #f, ReadPos, TempData2
                    StringToBytes TempData2, bytBytes, k
                    
                    If bytBytes(0) \ 32 = 7 Then ' 111xxxxx detected
                        ReadPos = ReadPos + 3: j = j + 3: k = k + 3
                        CurrPos = i + j - 1
                        
                        ' Check to see if intensity stereo is used
                        If ChannelMode = JointStereo Then
                            If Layer = LayerI Or Layer = LayerII Then
                                bIntensityStereo = True
                            Else
                                bIntensityStereo = CBool(bytBytes(2) \ 16 Mod 2)
                            End If
                        Else
                            bIntensityStereo = False
                        End If
                        
                        ' Filter out headers with reserved and unacceptable values and those that are potentially invalid. Free bitrate frames are RARE and therefore not supported.
                        If MPEGVersion <> MPEGVersionReserved And _
                           Layer <> LayerReserved And _
                           SamplingRateFrequency <> -1 And _
                           BitRate <> 0 And BitRate <> -1000 And _
                           bytBytes(2) Mod 4 <> 2 And _
                           (Not (((BitRate >= 32 And BitRate <= 56) Or BitRate = 80) And ((ChannelMode = Stereo) Or bIntensityStereo Or (ChannelMode = DualChannel)))) And _
                           (Not ((BitRate >= 224 And BitRate <= 384) And (ChannelMode = SingleChannel))) Then
                            m_HeaderPos = ReadPos
                            blnReturn = True
                            Exit Do
                        End If
                        
                        PrevPos = CurrPos
                    End If
                Else
                    Exit Do
                End If
            Loop
            
            If blnReturn Then Exit For
        Next
        
        m_AudioSize = lLOF - m_HeaderPos
        
        If MPEGVersion = MPEG1 Then
            If ChannelMode = SingleChannel Then
                lngXingOffset = 17
            Else
                lngXingOffset = 32
            End If
        Else
            If ChannelMode = SingleChannel Then
                lngXingOffset = 9
            Else
                lngXingOffset = 17
            End If
        End If
        ReadPos = m_HeaderPos + lngXingOffset: k = 4 + lngXingOffset
        
        StringToBytes TempData2, bytXingBytes, k
        ReadPos = ReadPos + 4: k = k + 4
        
        If LCase$(StrConv(bytXingBytes, vbUnicode)) = "xing" Or LCase$(StrConv(bytXingBytes, vbUnicode)) = "info" Then
            
            StringToBytes TempData2, bytXingFlags, k
            ReadPos = ReadPos + 4: k = k + 4
            lXingSize = 8
            
            If bytXingFlags(3) Mod 2 = 1 Then ' xxxxxxxx xxxxxxxx xxxxxxxx xxxxxxx0
                StringToBytes TempData2, bytTemp, k
                m_Xing_FrameCount = ToLong(bytTemp)
                ReadPos = ReadPos + 4: k = k + 4
                lXingSize = lXingSize + 4
            End If
            
            If bytXingFlags(3) \ 2 Mod 2 = 1 Then ' xxxxxxxx xxxxxxxx xxxxxxxx xxxxxx0x
                StringToBytes TempData2, bytTemp, k
                m_Xing_FileLength = ToLong(bytTemp)
                ReadPos = ReadPos + 4: k = k + 4
                lXingSize = lXingSize + 4
            End If
            
            If bytXingFlags(3) \ 4 Mod 2 = 1 Then ' xxxxxxxx xxxxxxxx xxxxxxxx xxxxx0xx
                StringToBytes TempData2, bytTOC, k
                m_Xing_TOC = bytTOC
                ReadPos = ReadPos + 100: k = k + 100
                lXingSize = lXingSize + 100
            End If
            
            If bytXingFlags(3) \ 8 Mod 2 = 1 Then ' xxxxxxxx xxxxxxxx xxxxxxxx xxxx0xxx
                StringToBytes TempData2, bytTemp, k
                m_Xing_Quality = ToLong(bytTemp)
                ReadPos = ReadPos + 4: k = k + 4
                lXingSize = lXingSize + 4
            End If
            
            If LCase$(StrConv(bytXingBytes, vbUnicode)) = "xing" Then
                m_VBR = True
            End If
            
            m_AudioSize = m_AudioSize - lXingSize
        End If
        
        ReadPos = m_HeaderPos + 32: k = 36
        
        StringToBytes TempData2, bytXingBytes, k
        ReadPos = ReadPos + 4
        k = k + 4
        
        If LCase$(StrConv(bytXingBytes, vbUnicode)) = "vbri" Then
                    
            m_VBR = True
            ReadPos = ReadPos + 4: k = k + 4
            
            StringToBytes TempData2, bytTemp2, k
            m_Xing_Quality = ToLong(bytTemp2)
            ReadPos = ReadPos + 2: k = k + 2
            
            StringToBytes TempData2, bytTemp, k
            m_Xing_FileLength = ToLong(bytTemp)
            ReadPos = ReadPos + 4: k = k + 4
            
            StringToBytes TempData2, bytTemp, k
            m_Xing_FrameCount = ToLong(bytTemp)
            ReadPos = ReadPos + 4: k = k + 4
            
            StringToBytes TempData2, bytTemp2, k
            m_VBRI_NumEntries = ToLong(bytTemp2)
            ReadPos = ReadPos + 4: k = k + 4
            
            StringToBytes TempData2, bytTemp2, k
            m_VBRI_SizePerEntry = ToLong(bytTemp2)
            m_AudioSize = m_AudioSize - 26 - m_VBRI_NumEntries * m_VBRI_SizePerEntry
        End If
    Close #f
End Sub

Public Sub UpdateID3Tags()
    On Error Resume Next
    
    Dim TempMP3File As String
    Dim IDTest As String
    Dim IDTest2 As String
    Dim ResetUnsyncFlags As Boolean
    Dim bCompression As Boolean
    Dim bExtendedHeader As Boolean
    Dim bFooterPresent As Boolean
    Dim TagSize As Double
    Dim Ridiculous As Boolean
    Dim TempData As String
    Dim PrevFinalOffset As Long
    Dim FinalOffset As Long
    Dim TempData2 As String
    Dim temp As String
    Dim temp2 As String
    Dim tFrameData As String
    Dim bHasDesc As Boolean
    Dim bExist As Boolean
    Dim f As Integer
    Dim f2 As Integer
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim PaddingSize As Long
    Dim NewStart As Long
    Dim OrigStart As Long
    Dim tGenre As String
    Dim tGenreID As GenreConstants
    Dim FLen As Long
    Dim colFrameNames As New Collection
    Dim Lyrics3Length As Long
    Dim ReplaceLyrics3Tag As Boolean
    
    Dim hFile As Long
    Dim lpCreationTime As FILETIME
    Dim lpLastAccessTime As FILETIME
    Dim lpLastWriteTime As FILETIME
    
    hFile = CreateFile(m_FileName, _
                       GENERIC_READ, _
                       FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                       0&, OPEN_EXISTING, 0&, 0&)
    GetFileTime hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime
    CloseHandle hFile
    
    ReplaceLyrics3Tag = False
    f = FreeFile
    
    ' In many cases, we need to insert new bytes into the MP3 file. To do this, we need to create a temporary MP3 file.
    TempMP3File = Environ$("TEMP") & "\mp3" & format$(VBA.Year(Now), "0000") & format$(Month(Now), "00") & format$(Day(Now), "00") & format$(Hour(Now), "00") & format$(Minute(Now), "00") & format$(Second(Now), "00") & ".mp3"
    
    Open m_FileName For Binary Access Read Shared As #f
    
        ' ID3v2
        
        IDTest = Space$(10) ' Give 10 blank bytes to the variable
        Get #f, 1, IDTest
        
        If Left$(IDTest, 3) = "ID3" And _
          Asc(Mid$(IDTest, 4, 1)) < &HFF& And _
          Asc(Mid$(IDTest, 5, 1)) < &HFF& And _
          Asc(Mid$(IDTest, 7, 1)) < &H80& And _
          Asc(Mid$(IDTest, 8, 1)) < &H80& And _
          Asc(Mid$(IDTest, 9, 1)) < &H80& And _
          Asc(Mid$(IDTest, 10, 1)) < &H80& Then ' ID3 header found
            
            m_ID3Revision = Asc(Mid$(IDTest, 4, 1))
            GetTagHeaderFlags IDTest, m_TagUnsynchronized, bCompression, bExtendedHeader, bFooterPresent
            TagSize = StringToSize(Right$(IDTest, 4))
            OrigStart = 11 + TagSize
            If bCompression Then
                ' Start a whole new tag; since all the frames are compressed, they will be of no use (ID3v2.0 and ID3v2.2 only)
                AllTagHeaderFlagsFalse m_TagUnsynchronized, bCompression, bExtendedHeader, bFooterPresent
                TempData = ""
            Else
                TempData = Space$(CLng(TagSize)) ' Prepare string for the ID3 tag
                Get #f, 11, TempData
                If m_TagUnsynchronized Then ' Undo unsynchronization if necessary and leave the tag that way
                    m_TagUnsynchronized = False
                    If m_ID3Revision < 4 Then
                        TempData = DeUnsynchronize(TempData) ' As unsynchronization in ID3v2.0, ID3v2.2, and ID3v2.3 is done on tag level, undo unsynchronization for the entire tag
                    End If
                End If
                If m_ID3Revision >= 4 Then ResetUnsyncFlags = True ' As unsynchronization in ID3v2.4 and newer is done on frame level, undo unsychronization frame by frame
                If bExtendedHeader Then ' Drop the extended header if detected
                    bExtendedHeader = False
                    TempData = Mid$(TempData, StringToSize(Left$(TempData, 4), (m_ID3Revision >= 4)) + 5)
                End If
                If bFooterPresent Then ' Since the tag is located at the very beginning, there is no need to have a footer
                    bFooterPresent = False
                    OrigStart = OrigStart + 10
                End If
            End If
        Else ' ID3 tag not found; we shall make one
            ' The version we shall input is ID3 v2.3
            m_ID3Revision = 3
            AllTagHeaderFlagsFalse m_TagUnsynchronized, bCompression, bExtendedHeader, bFooterPresent
            OrigStart = 1
            TempData = ""
        End If
        
        If m_ID3Revision > 2 Then  ' The newer standard
            For i = LBound(FrameNames) To UBound(FrameNames)
                If m_ID3Revision >= 4 Then ' ID3v2.4+
                    If FrameNames(i) = "TORY" Then
                        colFrameNames.Add "TDOR", "TDOR"
                    ElseIf FrameNames(i) = "TYER" Then
                        colFrameNames.Add "TDRC", "TDRC"
                    Else
                        GoTo AddNormal
                    End If
                Else
AddNormal:
                    colFrameNames.Add FrameNames(i), FrameNames(i)
                End If
            Next
        Else ' The older standard
            For i = LBound(OldFrameNames) To UBound(OldFrameNames)
                colFrameNames.Add OldFrameNames(i), OldFrameNames(i)
            Next
        End If
        
        i = 1
        IncrementA = 1
        IncrementB = 1
        IncrementC = 1
        Do
            If i > Len(TempData) Then Exit Do ' No more data; therefore, exit
            
            FinalOffset = FinalOffsetID3v2(TempData, i, Ridiculous)
            If Ridiculous Then Exit Do
            
            temp = Mid$(TempData, i, FrameNameSizeID3v2)
            temp2 = ""
            temp2 = colFrameNames(temp)
            
            If temp = "COM" Or temp = "COMM" Or temp = "ULT" Or temp = "USLT" Then  ' Comment or lyrics frame detected
                tFrameData = GetFrame(TempData, i, bHasDesc, True)  ' Check to see if a description exists
                If bHasDesc Then  ' The frame has some description
                    AdvancePosition TempData, i, i, ResetUnsyncFlags
                Else  ' The frame has no description
                    ' Allow only one comment or lyrics frame with no description -- delete any with ambiguous descriptions
                    If temp = temp2 Then
                        WriteTag temp, TempData, i, i, , ResetUnsyncFlags
                        If Not Ridiculous Then colFrameNames.Remove temp
                    Else
                        ' Ambiguous description detected; ergo, delete the frame
                        TempData = Left$(TempData, i - 1) & Mid$(TempData, i + FinalOffset)
                    End If
                End If
            Else  ' Some other frame detected
                If (m_ID3Revision <= 2 And temp = String$(3, 0)) Or (m_ID3Revision > 2 And temp = String$(4, 0)) Or IsPaddingCorruptedByMP3ext(TempData, i) Then
                    Exit Do ' Stop editing when we approach the padding or the corruption in the padding thanks to MP3ext (as you can tell, I'm not very pleased with compliance issues like this one)
                Else
                    WriteTag temp2, TempData, i, i, , ResetUnsyncFlags
                    If Not Ridiculous And temp = temp2 Then colFrameNames.Remove temp
                End If
            End If
            
            If Ridiculous Then Exit Do
        Loop
        PaddingSize = Len(TempData) - i + 1
        TempData = Left$(TempData, i - 1)
        
        
        For j = 1 To colFrameNames.Count
            temp = colFrameNames.Item(j)
            If temp = "WCM" Or temp = "WCOM" Then
            ElseIf temp = "WAR" Or temp = "WOAR" Then
            ElseIf temp = "PIC" Or temp = "APIC" Then
            Else
                WriteTag temp, TempData, i, i, True
            End If
        Next
        
        If PaddingSize = 0 Then PaddingSize = 2048 ' Set the padding size to 2048 bytes (2 KB) if there is none
        TempData = TempData & String$(PaddingSize, 0) ' Restore the padding
        
        IDTest = "ID3" & Chr$(m_ID3Revision) & Chr$(0) & GetTagHeaderFlagByte(m_TagUnsynchronized, bCompression, bExtendedHeader, bFooterPresent) & SizeToString(Len(TempData))
        NewStart = 11 + Len(TempData)
        
        ' ID3v1
        
        IDTest2 = Space$(3) ' Give 3 blank bytes to the variable
        Get #f, LOF(f) - 127, IDTest2
        FLen = LOF(f) - IIf(IDTest2 = "TAG", 128, 0)
        
        ' Lyrics3
        
        TempData2 = Space$(9)
        Get #f, FLen - 8, TempData2
        
        If TempData2 = LYRICS200 Then
            TempData2 = Space$(6)
            Get #f, FLen - 14, TempData2
            Lyrics3Length = CLng(TempData2)
            
            TempData2 = Space$(Lyrics3Length)
            Get #f, FLen - 14 - Lyrics3Length, TempData2
            
            If Left$(TempData2, 11) = LYRICSBEGIN Then
                FLen = FLen - 15 - Lyrics3Length
            End If
        ElseIf TempData2 = LYRICSEND Then  ' This version is extremely obsolete; it must be replaced
            ReplaceLyrics3Tag = True
            
            If FLen - 20 > 5100 Then
                j = 5100
                TempData2 = Space$(j)
                Get #f, FLen - 5119, TempData2
            Else
                j = FLen - 20
                TempData2 = Space$(j)
                Get #f, 1, TempData2
            End If
            
            i = InStrRev(TempData2, LYRICSBEGIN)
            If i > 0 Then
                FLen = FLen - 16 - j + i
            End If
        Else
            ReplaceLyrics3Tag = True
        End If
        
        If Len(m_Title) > 30 Or Len(m_Artist) > 30 Or Len(m_Album) > 30 Then
            TempData2 = LYRICSBEGIN & IND0000200
            If Not ReplaceLyrics3Tag Then TempData = TempData & Mid$(TempData2, 12)
            
            Set colFrameNames = Nothing
            Set colFrameNames = New Collection
            
            For i = LBound(Lyrics3FrameNames) To UBound(Lyrics3FrameNames)
                colFrameNames.Add Lyrics3FrameNames(i), Lyrics3FrameNames(i)
            Next
            
            i = 12
            Do
                If i > Len(TempData2) Then Exit Do ' No more data; therefore, exit
                temp = Mid$(TempData2, i, 3)
                temp2 = ""
                temp2 = colFrameNames(temp)
                
                s = Mid$(TempData2, i + 3, 5)
                If Not ProperSize(s) Then Exit Do ' The Lyrics3 tag may be corrupt
                If temp = "IND" Then
                    j = i + 8 + CLng(s)
                    If i > 12 Then TempData2 = Left$(TempData2, i - 1) & Mid$(TempData2, j)
                    i = j
                Else
                    If temp = temp2 Then colFrameNames.Remove temp
                    WriteLyrics3Tag temp2, TempData2, i, i
                End If
            Loop
            TempData2 = Left$(TempData2, i - 1) ' Truncate the data in case of corruption
            
            For j = 1 To colFrameNames.Count
                temp = colFrameNames.Item(j)
                WriteLyrics3Tag temp, TempData2, i, i, True
            Next
            
            TempData2 = TempData2 & format$(Len(TempData2), "000000") & LYRICS200
        Else
            TempData2 = ""
        End If
        
        ' ID3v1
        
        If (m_GenreID = OtherGenre And LCase$(m_Genre) <> "other") Or (m_GenreID = Unknown And LCase$(m_Genre) <> "unknown") Then
            tGenreID = ToGenreID(m_Genre)
        Else
            tGenreID = m_GenreID
        End If
        If tGenreID = Cover Or tGenreID = Remix Then tGenreID = OtherGenre
        
        If m_Title <> "" Or m_Artist <> "" Or m_Album <> "" Or m_Year <> "" Or m_Comments <> "" Or m_TrackNumber <> "" Or (tGenreID <> OtherGenre And tGenreID <> Unknown) Then
            temp = ""
            
            If IsNumeric(m_TrackNumber) Then
                If CDbl(m_TrackNumber) >= 1 And CDbl(m_TrackNumber) <= 255 Then
                    temp = AppendNullChars(m_Comments, 28) & Chr$(0) & Chr$(CLng(m_TrackNumber))
                End If
            End If
            
            If temp = "" Then _
               temp = AppendNullChars(m_Comments, 30)
            
            TempData2 = TempData2 & _
               "TAG" & AppendNullChars(m_Title, 30) & _
                       AppendNullChars(m_Artist, 30) & _
                       AppendNullChars(m_Album, 30) & _
                       AppendNullChars(m_Year, 4) & _
                       temp & Chr$(tGenreID)
        End If
    
        f2 = FreeFile
        Open TempMP3File For Binary Access Write Shared As #f2
            Put #f2, 1, IDTest
            Put #f2, 11, TempData
            
            For i = OrigStart To FLen Step 20480 ' Write to the temporary file in buffers of 20480 bytes (20 KB)
                If FLen - i + 1 > 20480 Then
                    TempData = Space$(20480)
                Else
                    TempData = Space$(FLen - i + 1)
                End If
                
                Get #f, i, TempData
                Put #f2, i - OrigStart + NewStart, TempData
            Next
            
            Put #f2, LOF(f2) + 1, TempData2
        Close #f2
    Close #f
    
    f = FreeFile
    Open m_FileName For Output Shared As #f ' Empty the target MP3 file
    Close #f
    
    f = FreeFile
    Open TempMP3File For Binary Access Read Shared As #f
    
        f2 = FreeFile
        Open m_FileName For Binary Access Write Shared As #f2
        
            For i = 1 To LOF(f) Step 20480 ' Write to the target MP3 file in buffers of 20480 bytes (20 KB)
                If LOF(f) - i + 1 > 20480 Then
                    TempData = Space$(20480)
                Else
                    TempData = Space$(LOF(f) - i + 1)
                End If
                
                Get #f, i, TempData
                Put #f2, i, TempData
            Next
        Close #f2
    Close #f
    
    hFile = CreateFile(m_FileName, _
                       GENERIC_WRITE, _
                       FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                       0&, OPEN_EXISTING, 0&, 0&)
    SetFileTime hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime
    CloseHandle hFile
    
    Kill TempMP3File ' Delete the temporary file
End Sub

Public Sub DeleteID3Tags()
    On Error Resume Next
    
    Dim IDTest As String
    Dim bCompression As Boolean
    Dim bExtendedHeader As Boolean
    Dim bFooterPresent As Boolean
    Dim TagSize As Double
    Dim TempData As String
    Dim f As Integer
    Dim f2 As Integer
    Dim FLen As Long
    Dim i As Long
    Dim j As Long
    Dim NewStart As Long
    Dim TempMP3File As String
    Dim Lyrics3Length As Long
    
    Dim hFile As Long
    Dim lpCreationTime As FILETIME
    Dim lpLastAccessTime As FILETIME
    Dim lpLastWriteTime As FILETIME
    
    m_Comments = ""
    m_Album = ""
    m_BeatsPerMinute = ""
    m_Composer = ""
    m_GenreID = Unknown
    m_Genre = ""
    m_Copyright = ""
    m_EncodedBy = ""
    m_Lyricist = ""
    m_Title = ""
    m_InitialKey = ""
    m_Languages = ""
    m_OriginalAlbum = ""
    m_OriginalFileName = ""
    m_OriginalLyricist = ""
    m_OriginalArtist = ""
    m_OriginalReleaseYear = ""
    m_FileOwner = ""
    m_Artist = ""
    m_Band = ""
    m_Conductor = ""
    m_InterpretedBy = ""
    m_DiscNumber = ""
    m_DiscsTotal = ""
    m_Publisher = ""
    m_TrackNumber = ""
    m_TracksTotal = ""
    m_InternetRadioStationName = ""
    m_InternetRadioStationOwner = ""
    m_ISRC = ""
    m_Year = ""
    m_Lyrics = ""
    m_CopyrightInfo = ""
    m_AudioURL = ""
    m_AudioSourceURL = ""
    m_InternetRadioURL = ""
    m_PaymentURL = ""
    m_PublisherURL = ""
    
    m_ID3Revision = 0
    
    m_ID3v1Exists = False
    m_ID3v2Exists = False
    m_Lyrics3v1Exists = False
    m_Lyrics3v2Exists = False
    LyricsTimestampsExist = False
    
    m_TagUnsynchronized = False
    
    hFile = CreateFile(m_FileName, _
                       GENERIC_READ, _
                       FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                       0&, OPEN_EXISTING, 0&, 0&)
    GetFileTime hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime
    CloseHandle hFile
    
    f = FreeFile
    
    ' We MUST create a temporary MP3 file in order to remove bytes from the target MP3 file. Without doing so will make it IMPOSSIBLE!
    TempMP3File = Environ$("TEMP") & "\mp3" & format$(VBA.Year(Now), "0000") & format$(Month(Now), "00") & format$(Day(Now), "00") & format$(Hour(Now), "00") & format$(Minute(Now), "00") & format$(Second(Now), "00") & ".mp3"
    
    Open m_FileName For Binary Access Read Shared As #f
    
        IDTest = Space$(10) ' Give 10 blank bytes to the variable
        Get #f, 1, IDTest
        
        If Left$(IDTest, 3) = "ID3" And _
          Asc(Mid$(IDTest, 4, 1)) < &HFF& And _
          Asc(Mid$(IDTest, 5, 1)) < &HFF& And _
          Asc(Mid$(IDTest, 7, 1)) < &H80& And _
          Asc(Mid$(IDTest, 8, 1)) < &H80& And _
          Asc(Mid$(IDTest, 9, 1)) < &H80& And _
          Asc(Mid$(IDTest, 10, 1)) < &H80& Then ' ID3 header found
            
            m_ID3Revision = Asc(Mid$(IDTest, 4, 1))
            GetTagHeaderFlags IDTest, m_TagUnsynchronized, bCompression, bExtendedHeader, bFooterPresent
            TagSize = StringToSize(Right$(IDTest, 4))
            NewStart = 11 + TagSize
            If bFooterPresent Then NewStart = NewStart + 10
        Else
            NewStart = 1
        End If
        
        IDTest = Space$(3) ' Give 3 blank bytes to the variable
        Get #f, LOF(f) - 127, IDTest
        FLen = LOF(f) - IIf(IDTest = "TAG", 128, 0)
                
        TempData = Space$(9)
        Get #f, FLen - 8, TempData
        
        If TempData = LYRICS200 Then
            TempData = Space$(6)
            Get #f, FLen - 14, TempData
            Lyrics3Length = CLng(TempData)
            
            TempData = Space$(Lyrics3Length)
            Get #f, FLen - 14 - Lyrics3Length, TempData
            
            If Left$(TempData, 11) = LYRICSBEGIN Then
                FLen = FLen - 15 - Lyrics3Length
            End If
        ElseIf TempData = LYRICSEND Then
            If FLen - 20 > 5100 Then
                j = 5100
                TempData = Space$(j)
                Get #f, FLen - 5119, TempData
            Else
                j = FLen - 20
                TempData = Space$(j)
                Get #f, 1, TempData
            End If
            
            i = InStrRev(TempData, LYRICSBEGIN)
            If i > 0 Then
                FLen = FLen - 16 - j + i
            End If
        End If
                
        f2 = FreeFile
        Open TempMP3File For Binary Access Write Shared As #f2
            For i = NewStart To FLen Step 20480 ' Write to the temporary file in buffers of 20480 bytes (20 KB)
                If FLen - i + 1 > 20480 Then
                    TempData = Space$(20480)
                Else
                    TempData = Space$(FLen - i + 1)
                End If
                
                Get #f, i, TempData
                Put #f2, i - NewStart + 1, TempData
            Next
        Close #f2
        
    Close #f
    
    f = FreeFile
    Open m_FileName For Output Shared As #f ' Empty the target MP3 file
    Close #f
    
    f = FreeFile
    Open TempMP3File For Binary Access Read Shared As #f
    
        f2 = FreeFile
        Open m_FileName For Binary Access Write Shared As #f2
        
            For i = 1 To LOF(f) Step 20480 ' Write to the target MP3 file in buffers of 20480 bytes (20 KB)
                If LOF(f) - i + 1 > 20480 Then
                    TempData = Space$(20480)
                Else
                    TempData = Space$(LOF(f) - i + 1)
                End If
                
                Get #f, i, TempData
                Put #f2, i, TempData
            Next
        Close #f2
    Close #f
    
    hFile = CreateFile(m_FileName, _
                       GENERIC_WRITE, _
                       FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                       0&, OPEN_EXISTING, 0&, 0&)
    SetFileTime hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime
    CloseHandle hFile
    
    Kill TempMP3File ' Delete the temporary file
End Sub

' ID3v1
Private Sub StoreBytes(ByVal Data As String, ByVal Start As Long, ByVal length As Long, Variable As String)
    Dim tData As String
    
    tData = RTrim$(StripNulls(Mid$(Data, Start, length)))
    If tData <> "" Then
        If Variable = "" Then Variable = tData
    End If
End Sub

' ID3v1
Private Sub TStoreByte(ByVal Data As String, Variable As String)
    Dim tData As String
    
    tData = Mid$(Data, 127, 1)
    If tData <> "" Then
        If Variable = "" And Asc(tData) > 0 Then Variable = CStr(Asc(tData))
    End If
End Sub

' ID3v1
Private Sub GStoreByte(ByVal Data As String, Variable As GenreConstants)
    Dim tData As String
    
    tData = Right$(Data, 1)
    If tData <> "" Then
        If Variable = OtherGenre Or Variable = Unknown Then Variable = CLng(Asc(tData))
    End If
End Sub

' ID3v2.4+
Private Function FormatYear(ByVal TimeStamp As String) As String
    Dim i As Long
    
    i = InStr(TimeStamp, "-")
    If i > 0 Then
        FormatYear = Left$(TimeStamp, i - 1)
    Else
        FormatYear = TimeStamp
    End If
End Function

' ID3v2
Private Function DeUnsynchronize(ByVal str As String) As String
    ' The purpose of unsynchronization is for certain applications not to read data within ID3 tags as MPEG headers.
    Dim i As Long
    Dim sRet As String
    
    sRet = str
    Do
        i = InStr(i + 1, sRet, "ÿ") ' 11111111
        If i > 0 And i < Len(sRet) Then
            If Mid$(sRet, i + 1, 1) = Chr$(0) Then ' 00000000 (ANY null character found immediately after ASCII character 255 needs to be removed in undoing unsynchronization)
                sRet = Left$(sRet, i) & Mid$(sRet, i + 2) ' ... 11111111 ...
            End If
        Else
            Exit Do
        End If
    Loop
    
    DeUnsynchronize = sRet
End Function

' ID3v2
Private Sub GetTagHeaderFlags(ByVal Tag As String, Unsynchronization As Boolean, Compression As Boolean, ExtendedHeader As Boolean, FooterPresent As Boolean)
    ' In uc000000 (ID3v2.0 and ID3v2.2), uex00000 (ID3v2.3), and uexf0000 (ID3v2.4+):
    ' -------------------------------------------------------------------------------
    ' u = Unsynchronization flag           x = Experimental indicator flag
    ' c = Compression flag                 f = Footer present flag
    ' e = Extended header flag
    
    Dim FlagByte As Byte
    FlagByte = Asc(Mid$(Tag, 6, 1))
    Unsynchronization = CBool(FlagByte \ 128) ' The unsynchronization flag is ALWAYS the very first flag on ALL versions of ID3v2 tags
    If m_ID3Revision > 2 Then
        Compression = False
        ExtendedHeader = CBool(FlagByte \ 64 Mod 2)
        If m_ID3Revision >= 4 Then ' uexf0000
            FooterPresent = CBool(FlagByte \ 16 Mod 2)
        Else ' uex00000
            FooterPresent = False
        End If
    Else ' uc000000
        Compression = CBool(FlagByte \ 64 Mod 2)
        ExtendedHeader = False
        FooterPresent = False
    End If
End Sub

' ID3v2
Private Sub AllTagHeaderFlagsFalse(Unsynchronization As Boolean, Compression As Boolean, ExtendedHeader As Boolean, FooterPresent As Boolean)
    Unsynchronization = False
    Compression = False
    ExtendedHeader = False
    FooterPresent = False
End Sub

' ID3v2
Private Function GetTagHeaderFlagByte(ByVal Unsynchronization As Boolean, ByVal Compression As Boolean, ByVal ExtendedHeader As Boolean, ByVal FooterPresent As Boolean) As String
    GetTagHeaderFlagByte = Chr$(Abs(Unsynchronization) * 2 ^ 7 + Abs((m_ID3Revision <= 2 And Compression) Or (m_ID3Revision > 2 And ExtendedHeader)) * 2 ^ 6 + Abs(m_ID3Revision >= 4 And FooterPresent) * 2 ^ 4)
End Function

' ID3v2.4+
Private Sub ResetFrameUnsyncFlag(Tag As String, ByVal Start As Long, lEnd As Long)
    On Error Resume Next
    
    Dim size As Long
    Dim Val As String
    Dim LVal As Long
    Dim FirstFlag As Byte
    Dim LastFlag As Byte
    
    If m_ID3Revision >= 4 Then ' Safeguard the procedure just in case the tag version is not ID3v2.4 or higher
        size = CLng(StringToSize(Mid$(Tag, Start + 4, 4)))
        Val = Mid$(Tag, Start + 10, size)
        FirstFlag = Asc(Mid$(Tag, Start + 8, 1))
        LastFlag = Asc(Mid$(Tag, Start + 9, 1))
        If LastFlag \ 2 Mod 2 = 1 Then ' Unsynchronization detected
            LastFlag = (LastFlag \ 4) * 4 + LastFlag Mod 2 ' leave out bit 2 (the unsynchronization bit)
            Val = DeUnsynchronize(Val)
            LVal = Len(Val)
            Tag = Left$(Tag, Start + 4) & SizeToString(LVal) & Chr$(FirstFlag) & Chr$(LastFlag) & Val & Mid$(Tag, Start + 10 + size)
            lEnd = Start + 10 + LVal
        Else ' Unsynchronization not detected; therefore, do nothing
            lEnd = Start + 10 + size
        End If
    End If
End Sub

' Lyrics3
Private Function ProperSize(ByVal Expression As String) As Boolean
    Dim i As Long
    Dim a As Byte
    Dim bRet As Boolean
    
    bRet = True
    For i = 1 To Len(Expression)
        a = Asc(Mid$(Expression, i, 1))
        If a < 48 Or a > 57 Then bRet = False: Exit For
    Next
    
    ProperSize = bRet
End Function

' ID3v2
Private Function FrameNameSizeID3v2() As Long
    FrameNameSizeID3v2 = 3 + Abs(m_ID3Revision > 2) ' Older versions: 3; Newer versions: 4
End Function

' ID3v2
Private Function FinalOffsetID3v2(ByVal Data As String, ByVal Start As Long, Ridiculous As Boolean) As Long
    ' When it comes to getting the best equation for the data offset, let's think algebraically:
    ' --------------------------------------------------------------------------------------------
    ' <frame name size>*2 + (<frame name size> - 3)*2  =  2x + 2(x - 3)  =  2x + 2x - 6  =  4x - 6
    '
    ' Frame name size    Header size (data offset)
    ' ---------------    -------------------------
    ' 3                  3*4 - 6  =  12 - 6  =  6    <- That's the offset we're looking for
    ' 4                  4*4 - 6  =  16 - 6  =  10   <- That's the other offset we expected, so we're all good
    
    ' The "Ridiculous" argument is a lifesaver. It saves us from the catastrophic endless loop in applicable areas
    
    On Error GoTo Err
    Dim FrameNameSize As Long: FrameNameSize = FrameNameSizeID3v2
    FinalOffsetID3v2 = FrameNameSize * 4 - 6 + CLng(StringToSize(Mid$(Data, Start + FrameNameSize, FrameNameSize), (m_ID3Revision >= 4)))
    On Error GoTo 0
    Ridiculous = False
    Exit Function
Err:
    Ridiculous = True
End Function

' ID3v2
Private Sub AdvancePosition(Data As String, ByVal Start As Long, lEnd As Long, ByVal ResetUnsyncFlag As Boolean)
    Dim Ridiculous As Boolean
    If ResetUnsyncFlag Then
        ResetFrameUnsyncFlag Data, Start, lEnd
    Else
        lEnd = Start + FinalOffsetID3v2(Data, Start, Ridiculous)
    End If
End Sub

' ID3v2
Private Function IsPaddingCorruptedByMP3ext(ByVal Data As String, ByVal Start As Long) As Boolean
    Dim mStart As Long
    Dim mStr As String
    Dim bRet As Boolean
    Const S_UNICODE As String = "(unicode)"
    
    mStr = String$(3, 0) & "MP3ext V"
    mStart = Len(mStr)
    Do
        If Mid$(Data, Start, mStart) = mStr Then
            Exit Do
        Else
            mStart = mStart - 1
            mStr = Right$(mStr, mStart)
            If mStart = 0 Then Exit Do
        End If
    Loop
    
    mStart = Start + mStart
    Do
        mStr = Mid$(Data, mStart, 1)
        If IsNumeric(mStr) Or mStr = "." Then
            mStart = mStart + 1
        Else
            Exit Do
        End If
    Loop
    bRet = (Mid$(Data, mStart, 9) = S_UNICODE)
    
    If Not bRet Then
        mStr = Right$(S_UNICODE, Len(S_UNICODE) - 1)
        mStart = Len(mStr)
        Do
            If Mid$(Data, Start, mStart) = mStr Then
                bRet = True
                Exit Do
            Else
                mStart = mStart - 1
                mStr = Right$(mStr, mStart)
                If mStart = 0 Then Exit Do
            End If
        Loop
    End If
    
    IsPaddingCorruptedByMP3ext = bRet
End Function

' ID3v2
Private Function GetFrame(ByVal Tag As String, ByVal Start As Long, HasDescription As Boolean, Optional ByVal LanguagePresent As Boolean = False, Optional ByVal TerminateNulls As Boolean = True) As String
    Dim LastFlag As Byte
    Dim bUnsynched As Boolean
    Dim bCompressed As Boolean
    Dim TextEncoding As Byte
    Dim Val As String
    Dim bUnicode As Boolean
    Dim i As Long
    
    ' We will need to use the second byte of the flags field
    LastFlag = Asc(Mid$(Tag, Start + 9, 1))
    
    ' Check for compression
    If m_ID3Revision > 2 Then
        If m_ID3Revision >= 4 Then
            bCompressed = CBool(LastFlag \ 8 Mod 2) ' xxxx0xxx
        Else
            bCompressed = CBool(LastFlag \ 128) ' 0xxxxxxx
        End If
    End If
    
    ' Skip the frame if compressed
    If Not bCompressed Then
        ' Since unsynchronization in ID3v2.4 is done on frame level, we must check the unsynchronization flag of the frame
        If m_ID3Revision >= 4 Then bUnsynched = CBool(LastFlag \ 2 Mod 2) ' xxxxxx0x
        
        If m_ID3Revision > 2 Then  ' The newer standard
            Val = Mid$(Tag, Start + 10, StringToSize(Mid$(Tag, Start + 4, 4), (m_ID3Revision >= 4)))
        Else ' The older standard
            Val = Mid$(Tag, Start + 6, StringToSize(Mid$(Tag, Start + 3, 3), False))
        End If
        If bUnsynched Then Val = DeUnsynchronize(Val) ' Undo unsynchronization if necessary
        
        ' Check the text encoding
        TextEncoding = Asc(Left$(Val, 1))
        
        ' Handle the text encoding; weed out an invalid byte
        If TextEncoding = 0 Or TextEncoding = 1 Or (m_ID3Revision >= 4 And (TextEncoding = 2 Or TextEncoding = 3)) Then
            Val = Mid$(Val, 2)
            Select Case TextEncoding
                Case 1: Val = UnicodeToANSI(Val)
                Case 2, 3: If m_ID3Revision >= 4 Then Val = UnicodeToANSINoBOM(Val, TextEncoding)
            End Select
        End If
        
        ' Take appropriate action when we need to handle the three-letter language identity
        If LanguagePresent Then Val = Mid$(Val, 4)
        
        ' Check to see if the description doesn't exist
        If Left$(Val, 1) = Chr$(0) Then
            HasDescription = False
            Val = Mid$(Val, 2)
        Else
            HasDescription = True
        End If
        
        ' Restrict the string to anything preceding the first null character if specified
        If TerminateNulls Then Val = StripNulls(Val)
        GetFrame = Val
    End If
End Function

' Lyrics3
Private Function GetLyrics3Frame(ByVal Data As String, ByVal Start As Long) As String
    Dim length As Long
    Dim Val As String
    
    length = CLng(Mid$(Data, Start + 3, 5))
    Val = Mid$(Data, Start + 8, length)
    
    ' Replace [CR] and/or [LF] with carriages and/or line feeds, respectively
    GetLyrics3Frame = Replace(Replace(Replace(Val, "[CR][LF]", vbCrLf), "[CR]", vbCr), "[LF]", vbLf)
End Function

' ID3v2
Private Sub AddOrChangeFrame(Tag As String, ByVal FrameID As String, ByVal Start As Long, lEnd As Long, ByVal NewValue As String, Optional ByVal AddFrame As Boolean = False, Optional ByVal LanguagePresent As Boolean = False)
    On Error Resume Next
    
    Dim StartAdd As Long
    Dim StartAdd2 As Long
    Dim size As Double
    Dim sValue As String
    Dim NewSize As String
    Dim LValue As Long
    Dim tStr As String
    Dim sFlags As String
    
    ' Generate the data appropriate for the frame
    If LanguagePresent Then
        sValue = Chr$(0) & "eng" & Chr$(0) & NewValue
    Else
        sValue = Chr$(0) & NewValue
    End If
    
    ' Calculate the size of the new frame data
    LValue = Len(sValue)
    If m_ID3Revision > 2 Then  ' The newer standard
        NewSize = SizeToString(LValue, 4, (m_ID3Revision >= 4))
        StartAdd = 10
        StartAdd2 = 4
        sFlags = String$(2, 0) ' 00000000 00000000
    Else ' The older standard
        NewSize = SizeToString(LValue, 3, False)
        StartAdd = 6
        StartAdd2 = 3
        sFlags = ""
    End If
    
    If AddFrame Then   ' Add the frame
        If NewValue = "" Then
            lEnd = Start
        Else
            Tag = Tag & FrameID & NewSize & sFlags & sValue
            lEnd = Start + StartAdd + LValue
        End If
    Else  ' Modify the frame
        size = StringToSize(Mid$(Tag, Start + StartAdd2, StartAdd2), (m_ID3Revision >= 4))
        If NewValue = "" Then
            Tag = Left$(Tag, Start - 1) & Mid$(Tag, Start + StartAdd + size)
            lEnd = Start
        Else
            Tag = Left$(Tag, Start + StartAdd2 - 1) & NewSize & sFlags & sValue & Mid$(Tag, Start + StartAdd + size)
            lEnd = Start + StartAdd + LValue
        End If
    End If
End Sub

' Lyrics3
Private Sub AddOrChangeLyrics3Frame(Data As String, ByVal FrameID As String, ByVal Start As Long, lEnd As Long, ByVal NewValue As String, Optional ByVal AddFrame As Boolean = False)
    On Error Resume Next
    
    Dim size As Long
    Dim sValue As String
    Dim NewSize As String
    Dim LValue As Long
    Dim tStr As String
    
    ' Replace carriages and/or line feeds with [CR] and/or [LF], respectively
    sValue = Replace(Replace(Replace(NewValue, vbCrLf, "[CR][LF]"), vbCr, "[CR]"), vbLf, "[LF]")
    
    ' Calculate the size of the new frame data
    LValue = Len(sValue)
    NewSize = format$(LValue, "00000")
    
    If AddFrame Then  ' Add the frame
        If NewValue = "" Then
            lEnd = Start
        Else
            Data = Data & FrameID & NewSize & sValue
            lEnd = Start + 8 + LValue
        End If
    Else  ' Modify the frame
        size = CLng(Mid$(Data, Start + 3, 5))
        If NewValue = "" Then
            Data = Left$(Data, Start - 1) & Mid$(Data, Start + 8 + size)
            lEnd = Start
        Else
            Data = Left$(Data, Start + 2) & NewSize & sValue & Mid$(Data, Start + 8 + size)
            lEnd = Start + 8 + LValue
        End If
    End If
End Sub

' ID3v2
Private Function IsValidArray(str() As String) As Boolean
    On Error Resume Next
    
    ' Is the variable set?
    Dim tmp1 As Long
    Dim tmp2 As Long
    
    tmp1 = LBound(str)
    tmp2 = UBound(str)
    
    IsValidArray = (Err.Number = 0)
End Function

' ID3v2
Private Function ValidateNum(ByVal str As String) As String
    Dim tmp As String
    
    tmp = str
    If tmp <> "" Then
        If Not IsNumeric(tmp) Then tmp = "0"
    End If
    
    ValidateNum = tmp
End Function

' ID3v1 and ID3v2
Private Function StripNulls(ByVal str As String) As String
    Dim i As Long
    
    i = InStr(str, Chr$(0)) ' Find the very first null character of the string
    If i > 0 Then
        StripNulls = Left$(str, i - 1) ' Remove the null character and any data after it
    Else
        StripNulls = str ' No null character found; keep the string as it is
    End If
End Function

' ID3v2
Private Function AllNulls(ByVal str As String) As Boolean
    Dim i As Long
    Dim bAll As Boolean
    
    bAll = True
    For i = 1 To Len(str)
        If Mid$(str, i, 1) <> Chr$(0) Then ' Not a null character
            bAll = False
            Exit For
        End If
    Next
    
    AllNulls = bAll
End Function

' ID3v2
Private Function StringToSize(ByVal str As String, Optional ByVal SynchSafe As Boolean = True) As Double
    Dim i As Long
    Dim num As Double
    Dim Base As Double
    
    num = 0
    Base = 256 * (2 ^ CDbl(SynchSafe))
    
    For i = Len(str) To 1 Step -1
        num = num + CDbl(Asc(Mid$(str, i, 1))) * Base ^ CDbl(Len(str) - i)
    Next
    
    StringToSize = num
End Function

' ID3v2
Private Function SizeToString(ByVal num As Double, Optional ByVal size As Long = 4, Optional ByVal SynchSafe As Boolean = True) As String
    Dim i As Long
    Dim mTmp As Double
    Dim dTmp As Double
    Dim sTmp As String
    Dim Base As Double
    
    dTmp = num
    sTmp = ""
    Base = 256 * (2 ^ CDbl(SynchSafe))
    
    Do
        mTmp = Fix(dTmp / Base)
        sTmp = Chr$(CLng(dTmp - Base * mTmp)) & sTmp
        If mTmp = 0 Then Exit Do
        dTmp = mTmp
    Loop
    
    If Len(sTmp) > size Then
        SizeToString = Right$(sTmp, size)
    Else
        SizeToString = String$(size - Len(sTmp), 0) & sTmp
    End If
End Function

' ID3v1
Private Function AppendNullChars(ByVal str As String, ByVal NewLength As Long, Optional ByVal RestrictLength As Boolean = True) As String
    Dim sStr As String
    
    If RestrictLength Then
        If Len(str) > NewLength Then
            sStr = Left$(str, NewLength)
        Else
            sStr = str
        End If
    Else
        sStr = str
    End If
    
    If NewLength > Len(sStr) Then
        AppendNullChars = sStr & String$(NewLength - Len(sStr), 0)
    Else
        AppendNullChars = sStr
    End If
End Function

' ID3v2
Private Function RemoveString(ByVal StringCheck As String, ByVal StringMatch As String) As String
    Dim i As Long
    Dim s As String
    
    s = StringCheck
    i = 1
    
    Do
        i = InStr(i, s, StringMatch)
        If i > 0 Then
            s = Left$(s, i - 1) & Mid$(s, i + Len(StringMatch))
        Else
            Exit Do
        End If
    Loop
    
    RemoveString = s
End Function

' Lyrics3
Private Function RemoveTimestamps(ByVal StringCheck As String) As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim s As String
    Dim t As String
    Dim mn As String
    Dim sc As String
    Dim cMn As Double
    Dim cSc As Double
    
    s = StringCheck
    i = 1
    
    Do
        i = InStr(i, s, "[")
        If i > 0 Then
            j = InStr(i + 1, s, "]")
            If j > 0 Then
                t = Mid$(s, i + 1, j - i - 1)
                k = InStr(t, ":")
                If k > 0 Then
                    mn = Left$(t, k - 1)
                    sc = Mid$(t, k + 1)
                    If IsNumeric(mn) And IsNumeric(sc) Then ' Potential timestamp found
                        cMn = CDbl(mn)
                        cSc = CDbl(sc)
                        
                        If cMn = Fix(cMn) And cSc = Fix(cSc) And cMn >= 0 And cSc >= 0 Then ' Genuine timestamp found
                            s = Left$(s, i + 1) & Mid$(s, j + 1)
                        End If
                    End If
                End If
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    RemoveTimestamps = s
End Function

' ID3v2
Private Function SwapBytes(ByVal str As String) As String
    ' little endian -> big endian   OR   big endian -> little endian
    
    Dim i As Long
    Dim x As String
    
    For i = 1 To Len(str) Step 2
        x = x & Mid$(str, i + 1, 1) & Mid$(str, i, 1)
    Next
    
    SwapBytes = x
End Function

' ID3v2
Private Function UnicodeToANSI(ByVal UnicodeStr As String) As String
    Dim i0 As Long
    Dim i As Long
    Dim i2 As Long
    Dim lastI As Long
    
    Dim sRet As String
    Dim lUnicode As Long ' 0 = No unicode, 1 = Little endian, 2 = Big endian
    
    lastI = 1
    Do
        i = InStr(i0 + 1, UnicodeStr, Chr$(&HFF&) & Chr$(&HFE&)) ' Little-endian BOM position
        i2 = InStr(i0 + 1, UnicodeStr, Chr$(&HFE&) & Chr$(&HFF&)) ' Big-endian BOM position
        
        If (i > 0 And i2 > 0 And i < i2) Or (i > 0 And i2 = 0) Then
            Select Case lUnicode
                Case 0: sRet = sRet & Mid$(UnicodeStr, lastI, i - lastI) ' No unicode previously detected
                Case 1: sRet = sRet & StrConv(Mid$(UnicodeStr, lastI + 2, i - lastI - 2), vbFromUnicode) ' Little-endian unicode previously detected
                Case 2: sRet = sRet & StrConv(SwapBytes(Mid$(UnicodeStr, lastI + 2, i - lastI - 2)), vbFromUnicode) ' Big-endian unicode previously detected
            End Select
            lastI = i
            lUnicode = 1 ' Little endian
            i0 = i + 1
        ElseIf (i > 0 And i2 > 0 And i2 < i) Or (i = 0 And i2 > 0) Then
            Select Case lUnicode
                Case 0: sRet = sRet & Mid$(UnicodeStr, lastI, i2 - lastI) ' No unicode previously detected
                Case 1: sRet = sRet & StrConv(Mid$(UnicodeStr, lastI + 2, i2 - lastI - 2), vbFromUnicode) ' Little-endian unicode previously detected
                Case 2: sRet = sRet & StrConv(SwapBytes(Mid$(UnicodeStr, lastI + 2, i2 - lastI - 2)), vbFromUnicode) ' Big-endian unicode previously detected
            End Select
            lastI = i2
            lUnicode = 2 ' Big endian
            i0 = i2 + 1
        Else
            Select Case lUnicode
                Case 0: sRet = sRet & Mid$(UnicodeStr, lastI) ' No unicode previously detected
                Case 1: sRet = sRet & StrConv(Mid$(UnicodeStr, lastI + 2), vbFromUnicode) ' Little-endian unicode previously detected
                Case 2: sRet = sRet & StrConv(SwapBytes(Mid$(UnicodeStr, lastI + 2)), vbFromUnicode) ' Big-endian unicode previously detected
            End Select
            Exit Do
        End If
    Loop
    
    UnicodeToANSI = sRet
End Function

' ID3v2.4+
Private Function UnicodeToANSINoBOM(ByVal UnicodeStr As String, ByVal TextEncoding As Byte) As String
    Dim sRet As String
    Select Case TextEncoding
        Case 2: sRet = StrConv(SwapBytes(UnicodeStr), vbFromUnicode)
        Case 3: sRet = UTF8toANSI(UnicodeStr)
    End Select
    UnicodeToANSINoBOM = sRet
End Function

' ID3v2
Private Sub ReadTag(ByVal DataID As String, ByVal Data As String, ByVal Start As Long, strGenre As String, strPartOfASet() As String, strTrack() As String)
    Dim sValue As String
    Dim bHasDesc As Boolean
    
    sValue = GetFrame(Data, Start, bHasDesc, _
        (DataID = "COM" Or DataID = "COMM" Or DataID = "ULT" Or DataID = "USLT"), _
        (DataID <> "PIC" And DataID <> "APIC"))
    
    Select Case DataID
        Case "COM", "COMM": If Not bHasDesc Then m_Comments = sValue
        Case "TAL", "TALB": m_Album = sValue
        Case "TBP", "TBPM": m_BeatsPerMinute = sValue
        Case "TCM", "TCOM": m_Composer = sValue
        Case "TCO", "TCON": strGenre = sValue
        Case "TCR", "TCOP": m_Copyright = sValue
        Case "TEN", "TENC": m_EncodedBy = sValue
        Case "TXT", "TEXT": m_Lyricist = sValue
        Case "TT2", "TIT2": m_Title = sValue
        Case "TKE", "TKEY": m_InitialKey = sValue
        Case "TLA", "TLAN": m_Languages = sValue
        Case "TOT", "TOAL": m_OriginalAlbum = sValue
        Case "TOF", "TOFN": m_OriginalFileName = sValue
        Case "TOL", "TOLY": m_OriginalLyricist = sValue
        Case "TOA", "TOPE": m_OriginalArtist = sValue
        Case "TOR", "TORY": m_OriginalReleaseYear = sValue
        Case "TDOR": m_OriginalReleaseYear = FormatYear(sValue)
        Case "TOWN": m_FileOwner = sValue
        Case "TP1", "TPE1": m_Artist = sValue
        Case "TP2", "TPE2": m_Band = sValue
        Case "TP3", "TPE3": m_Conductor = sValue
        Case "TP4", "TPE4": m_InterpretedBy = sValue
        Case "TPA", "TPOS": strPartOfASet = Split(sValue, "/")
        Case "TPB", "TPUB": m_Publisher = sValue
        Case "TRK", "TRCK": strTrack = Split(sValue, "/")
        Case "TRSN": m_InternetRadioStationName = sValue
        Case "TRSO": m_InternetRadioStationOwner = sValue
        Case "TRC", "TSRC": m_ISRC = sValue
        Case "TYE", "TYER": m_Year = sValue
        Case "TDRC": m_Year = FormatYear(sValue)
        Case "ULT", "USLT": If Not bHasDesc Then m_Lyrics = sValue
        Case "WCP", "WCOP": m_CopyrightInfo = sValue
        Case "WAF", "WOAF": m_AudioURL = sValue
        Case "WAS", "WOAS": m_AudioSourceURL = sValue
        Case "WORS": m_InternetRadioURL = sValue
        Case "WPAY": m_PaymentURL = sValue
        Case "WPB", "WPUB": m_PublisherURL = sValue
    End Select
End Sub

' Lyrics3
Private Sub ReadLyrics3Tag(ByVal DataID As String, ByVal Data As String, ByVal Start As Long)
    Dim sValue As String
    
    sValue = GetLyrics3Frame(Data, Start)
    
    Select Case DataID
        Case "IND"
            LyricsTimestampsExist = (Mid$(sValue, Start + 5, 1) = "1")
        Case "LYR"
            If m_Lyrics = "" Then
                If LyricsTimestampsExist Then
                    m_Lyrics = RemoveTimestamps(sValue)
                Else
                    m_Lyrics = sValue
                End If
            End If
        Case "AUT"
            If m_Lyricist = "" Then m_Lyricist = sValue
        Case "EAL"
            If m_Album = "" Then m_Album = sValue
        Case "EAR"
            If m_Artist = "" Then m_Artist = sValue
        Case "ETT"
            If m_Title = "" Then m_Title = sValue
    End Select
End Sub

' ID3v2
Private Sub WriteTag(ByVal DataID As String, Data As String, ByVal Start As Long, lEnd As Long, Optional ByVal AddFrame As Boolean = False, Optional ByVal ResetUnsyncFlag As Boolean = False)
    Dim sValue As String
    Dim sGenre As String
    Dim bLang As Boolean
    Dim bWrite As Boolean
    Dim tGenreID As String
    
    bWrite = True
    Select Case DataID
        Case "COM", "COMM": sValue = m_Comments: bLang = True
        Case "TAL", "TALB": sValue = m_Album
        Case "TBP", "TBPM": sValue = m_BeatsPerMinute
        Case "TCM", "TCOM": sValue = m_Composer
        Case "TCO", "TCON"
            sGenre = Trim$(LCase$(m_Genre))
            
            If m_GenreID = Cover Then
                tGenreID = "CR"
            ElseIf m_GenreID = Remix Then
                tGenreID = "RX"
            Else
                tGenreID = CStr(m_GenreID)
            End If
            
            If (m_GenreID = OtherGenre Or m_GenreID = Unknown) And sGenre <> "" Then
                sValue = m_Genre
            ElseIf (m_GenreID <> OtherGenre Or m_GenreID <> Unknown) And (sGenre = "" Or sGenre = "other" Or sGenre = "unknown") Then
                sValue = "(" & tGenreID & ")"
            Else
                sValue = "(" & tGenreID & ")" & m_Genre
            End If
        Case "TCR", "TCOP": sValue = m_Copyright
        Case "TEN", "TENC": sValue = m_EncodedBy
        Case "TXT", "TEXT": sValue = m_Lyricist
        Case "TT2", "TIT2": sValue = m_Title
        Case "TKE", "TKEY": sValue = m_InitialKey
        Case "TLA", "TLAN": sValue = m_Languages
        Case "TOT", "TOAL": sValue = m_OriginalAlbum
        Case "TOF", "TOFN": sValue = m_OriginalFileName
        Case "TOL", "TOLY": sValue = m_OriginalLyricist
        Case "TOA", "TOPE": sValue = m_OriginalArtist
        Case "TOR", "TORY", "TDOR": sValue = m_OriginalReleaseYear
        Case "TOWN": sValue = m_FileOwner
        Case "TP1", "TPE1": sValue = m_Artist
        Case "TP2", "TPE2": sValue = m_Band
        Case "TP3", "TPE3": sValue = m_Conductor
        Case "TP4", "TPE4": sValue = m_InterpretedBy
        Case "TPA", "TPOS"
            If m_DiscNumber <> "" And m_DiscsTotal = "" Then
                sValue = m_DiscNumber
            ElseIf m_DiscNumber <> "" And m_DiscsTotal <> "" Then
                sValue = m_DiscNumber & "/" & m_DiscsTotal
            Else
                sValue = ""
            End If
        Case "TPB", "TPUB": sValue = m_Publisher
        Case "TRK", "TRCK"
            If m_TrackNumber <> "" And m_TracksTotal = "" Then
                sValue = m_TrackNumber
            ElseIf m_TrackNumber <> "" And m_TracksTotal <> "" Then
                sValue = m_TrackNumber & "/" & m_TracksTotal
            Else
                sValue = ""
            End If
        Case "TRSN": sValue = m_InternetRadioStationName
        Case "TRSO": sValue = m_InternetRadioStationOwner
        Case "TRC", "TSRC": sValue = m_ISRC
        Case "TYE", "TYER", "TDRC": sValue = m_Year
        Case "ULT", "USLT": sValue = m_Lyrics: bLang = True
        Case "WCP", "WCOP": sValue = m_CopyrightInfo
        Case "WAF", "WOAF": sValue = m_AudioURL
        Case "WAS", "WOAS": sValue = m_AudioSourceURL
        Case "WORS": sValue = m_InternetRadioURL
        Case "WPAY": sValue = m_PaymentURL
        Case "WPB", "WPUB": sValue = m_PublisherURL
        Case Else
            bWrite = False
            If AddFrame Then lEnd = Start Else AdvancePosition Data, Start, lEnd, ResetUnsyncFlag
    End Select
    
    If bWrite Then AddOrChangeFrame Data, DataID, Start, lEnd, sValue, AddFrame, bLang
End Sub

' Lyrics3
Private Sub WriteLyrics3Tag(ByVal DataID As String, Data As String, ByVal Start As Long, lEnd As Long, Optional ByVal AddFrame As Boolean = False)
    Dim sValue As String
    Dim bLang As Boolean
    Dim bWrite As Boolean
    
    sValue = ""
    bWrite = True
    
    Select Case DataID
        Case "EAL"
            If Len(m_Album) > 30 Then sValue = m_Album
        Case "LYR", "AUT" ' Delete
        Case "EAR"
            If Len(m_Artist) > 30 Then sValue = m_Artist
        Case "ETT"
            If Len(m_Title) > 30 Then sValue = m_Title
        Case Else
            bWrite = False
            lEnd = Start
            If Not AddFrame Then lEnd = lEnd + 8 + CLng(Mid$(Data, Start + 3, 5))
    End Select
    
    If Len(sValue) > 250 Then sValue = Left$(sValue, 250)
    If bWrite Then AddOrChangeLyrics3Frame Data, DataID, Start, lEnd, sValue, AddFrame
End Sub

Private Sub StringToBytes(ByVal str As String, Arr() As Byte, ByVal Start As Long)
    Dim arrTemp() As Byte
    Dim i As Long
    
    arrTemp = StrConv(Mid$(str, Start, UBound(Arr) - LBound(Arr) + 1), vbFromUnicode)
    
    For i = LBound(arrTemp) To UBound(arrTemp)
        Arr(i) = arrTemp(i)
    Next
End Sub

Private Function ToLong(Value() As Byte) As Long
    Dim i As Long
    Dim Count As Long
    Dim tmp As Long
    Dim LVal As Long
    
    Dim lowerBound As Long
    Dim upperBound As Long
    
    lowerBound = LBound(Value)
    upperBound = UBound(Value)
    
    Count = upperBound - lowerBound + 1
    If Count > 4 Then Err.Raise 6
    
    For i = lowerBound To upperBound
        tmp = CLng(Value(i))
        tmp = tmp - ((upperBound - i + 1) \ 2 And 2) * (tmp And 128)
        LVal = LVal + tmp * 256 ^ (upperBound - i)
    Next
    
    ToLong = LVal
End Function

Private Function Choose(ByVal Index As Double, ParamArray Choice())
    Choose = Choice(Index - 1 + LBound(Choice))
End Function

Private Sub Class_Initialize()
    m_ID3Revision = 3
    
    ' Load all the frame names this class uses
    FrameNames(0) = "APIC"
    FrameNames(1) = "COMM"
    FrameNames(2) = "TALB"
    FrameNames(3) = "TBPM"
    FrameNames(4) = "TCOM"
    FrameNames(5) = "TCON"
    FrameNames(6) = "TCOP"
    FrameNames(7) = "TENC"
    FrameNames(8) = "TEXT"
    FrameNames(9) = "TIT2"
    FrameNames(10) = "TKEY"
    FrameNames(11) = "TLAN"
    FrameNames(12) = "TOAL"
    FrameNames(13) = "TOFN"
    FrameNames(14) = "TOLY"
    FrameNames(15) = "TOPE"
    FrameNames(16) = "TORY" ' TDOR in ID3v2.4+
    FrameNames(17) = "TOWN"
    FrameNames(18) = "TPE1"
    FrameNames(19) = "TPE2"
    FrameNames(20) = "TPE3"
    FrameNames(21) = "TPE4"
    FrameNames(22) = "TPOS"
    FrameNames(23) = "TPUB"
    FrameNames(24) = "TRCK"
    FrameNames(25) = "TRSN"
    FrameNames(26) = "TRSO"
    FrameNames(27) = "TSRC"
    FrameNames(28) = "TYER" ' TDRC in ID3v2.4+
    FrameNames(29) = "USLT"
    FrameNames(30) = "WCOM"
    FrameNames(31) = "WCOP"
    FrameNames(32) = "WOAF"
    FrameNames(33) = "WOAR"
    FrameNames(34) = "WOAS"
    FrameNames(35) = "WORS"
    FrameNames(36) = "WPAY"
    FrameNames(37) = "WPUB"

    ' Load all the obsolete frame names this class uses
    OldFrameNames(0) = "COM"
    OldFrameNames(1) = "PIC"
    OldFrameNames(2) = "TAL"
    OldFrameNames(3) = "TBP"
    OldFrameNames(4) = "TCM"
    OldFrameNames(5) = "TCO"
    OldFrameNames(6) = "TCR"
    OldFrameNames(7) = "TEN"
    OldFrameNames(8) = "TXT"
    OldFrameNames(9) = "TT2"
    OldFrameNames(10) = "TKE"
    OldFrameNames(11) = "TLA"
    OldFrameNames(12) = "TOT"
    OldFrameNames(13) = "TOF"
    OldFrameNames(14) = "TOL"
    OldFrameNames(15) = "TOA"
    OldFrameNames(16) = "TOR"
    OldFrameNames(17) = "TP1"
    OldFrameNames(18) = "TP2"
    OldFrameNames(19) = "TP3"
    OldFrameNames(20) = "TP4"
    OldFrameNames(21) = "TPA"
    OldFrameNames(22) = "TPB"
    OldFrameNames(23) = "TRK"
    OldFrameNames(24) = "TRC"
    OldFrameNames(25) = "TYE"
    OldFrameNames(26) = "ULT"
    OldFrameNames(27) = "WCM"
    OldFrameNames(28) = "WCP"
    OldFrameNames(29) = "WAF"
    OldFrameNames(30) = "WAR"
    OldFrameNames(31) = "WAS"
    OldFrameNames(32) = "WPB"
    
    ' Load all the Lyrics3 frame names this class uses (with the exception of IND)
    Lyrics3FrameNames(0) = "LYR"
    Lyrics3FrameNames(1) = "AUT"
    Lyrics3FrameNames(2) = "EAL"
    Lyrics3FrameNames(3) = "EAR"
    Lyrics3FrameNames(4) = "ETT"
End Sub

Private Function UTF8toANSI(ByVal str As String) As String
    Dim i As Long
    Dim CCode As Integer
    Dim CCode2 As Integer
    Dim CCode3 As Integer
    Dim s As String
    
    For i = 1 To Len(str)
        CCode = Asc(Mid$(str, i, 1))
        
        Select Case CCode
            Case &H0 To &HBF
                s = s & ChrW$(CCode)
            Case &HC0 To &HDF
                If i > Len(str) - 1 Then Exit For
                i = i + 1
                CCode2 = Asc(Mid$(str, i, 1))
                s = s & ChrW$((CCode - &HC0) * &H40 + CCode2 - &H80)
            Case Else
                If i > Len(str) - 2 Then Exit For
                i = i + 1
                CCode2 = Asc(Mid$(str, i, 1)): i = i + 1
                CCode3 = Asc(Mid$(str, i, 1))
                s = s & ChrW$((CCode - &HE0) * &H1000 + (CCode2 - &H80) * &H40 + CCode2 - &H80)
        End Select
    Next
    
    UTF8toANSI = s
End Function
