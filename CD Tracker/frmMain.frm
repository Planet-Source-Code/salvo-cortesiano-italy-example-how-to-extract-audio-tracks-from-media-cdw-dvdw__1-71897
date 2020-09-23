VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Object = "{3C391B72-C020-4837-9B6B-5BB0AACFAA24}#76.0#0"; "utcCover.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD Tracker v1.0.3b"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13980
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13980
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   12615
      TabIndex        =   110
      Top             =   7410
      Width           =   1245
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4635
      ScaleHeight     =   300
      ScaleWidth      =   2370
      TabIndex        =   102
      Top             =   6360
      Width           =   2370
      Begin VB.CheckBox CheckReplace 
         Caption         =   "Replace Exist File"
         Height          =   255
         Left            =   30
         TabIndex        =   103
         Top             =   30
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.PictureBox PicFrames 
      BorderStyle     =   0  'None
      Height          =   1185
      Index           =   2
      Left            =   45
      ScaleHeight     =   1185
      ScaleWidth      =   6930
      TabIndex        =   89
      Top             =   6690
      Visible         =   0   'False
      Width           =   6930
      Begin VB.Frame Frame6 
         Caption         =   "Extra Tags"
         Height          =   1110
         Left            =   60
         TabIndex        =   90
         Top             =   30
         Width           =   6825
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   870
            Left            =   30
            ScaleHeight     =   870
            ScaleWidth      =   6765
            TabIndex        =   91
            Top             =   210
            Width           =   6765
            Begin VB.CheckBox CheckIncludeCover 
               Caption         =   "Include Cover"
               Height          =   255
               Left            =   5010
               TabIndex        =   99
               Top             =   465
               Width           =   1710
            End
            Begin VB.CheckBox CheckIncludeExtraTags 
               Caption         =   "Include Extra"
               Height          =   255
               Left            =   3255
               TabIndex        =   98
               Top             =   465
               Value           =   1  'Checked
               Width           =   1770
            End
            Begin VB.TextBox txtLanguage 
               Height          =   270
               Left            =   1305
               TabIndex        =   97
               Text            =   "Italian"
               Top             =   435
               Width           =   1875
            End
            Begin VB.TextBox txtCopyrightInfo 
               Height          =   270
               Left            =   4905
               TabIndex        =   95
               Text            =   "http://www.netshadows.it"
               Top             =   105
               Width           =   1815
            End
            Begin VB.TextBox txtEncodedBy 
               Height          =   270
               Left            =   1305
               TabIndex        =   92
               Text            =   "Salvo Cortesiano"
               Top             =   105
               Width           =   1875
            End
            Begin VB.Label Label20 
               Caption         =   "Language:"
               Height          =   255
               Left            =   270
               TabIndex        =   96
               Top             =   465
               Width           =   1005
            End
            Begin VB.Label Label19 
               Caption         =   "Copyright Info:"
               Height          =   255
               Left            =   3270
               TabIndex        =   94
               Top             =   105
               Width           =   1725
            End
            Begin VB.Label Label13 
               Caption         =   "Encoded by:"
               Height          =   255
               Left            =   60
               TabIndex        =   93
               Top             =   105
               Width           =   1245
            End
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   4965
      ScaleHeight     =   345
      ScaleWidth      =   1575
      TabIndex        =   81
      Top             =   1050
      Width           =   1575
      Begin VB.Label Label1 
         Caption         =   "Default Path:"
         Height          =   210
         Left            =   105
         TabIndex        =   82
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.ListBox lstMedia 
      Height          =   270
      Left            =   7335
      TabIndex        =   80
      Top             =   75
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   1
      Left            =   60
      ScaleHeight     =   4920
      ScaleWidth      =   13830
      TabIndex        =   64
      Top             =   1410
      Visible         =   0   'False
      Width           =   13830
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   8385
         ScaleHeight     =   390
         ScaleWidth      =   5385
         TabIndex        =   119
         Top             =   2490
         Width           =   5385
         Begin VB.CommandButton cmdBrowseMedia 
            Caption         =   "..."
            Height          =   240
            Left            =   4875
            TabIndex        =   121
            ToolTipText     =   "Browse Media Files"
            Top             =   75
            Width           =   495
         End
         Begin VB.TextBox txtMediaPath 
            Height          =   255
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   120
            ToolTipText     =   "This is the default media Path"
            Top             =   75
            Width           =   4770
         End
      End
      Begin VB.CommandButton cmdpPlay 
         Caption         =   "4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9075
         TabIndex        =   116
         ToolTipText     =   "Play"
         Top             =   2925
         Width           =   300
      End
      Begin VB.CommandButton cmdpStop 
         Caption         =   "<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9735
         TabIndex        =   115
         ToolTipText     =   "Stop"
         Top             =   2925
         Width           =   300
      End
      Begin VB.CommandButton cmdpPause 
         Caption         =   ";"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9405
         TabIndex        =   114
         ToolTipText     =   "Pause"
         Top             =   2925
         Width           =   300
      End
      Begin VB.CommandButton cmdnpPrev 
         Caption         =   "9"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   7.5
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   8400
         TabIndex        =   113
         ToolTipText     =   "Prev"
         Top             =   2925
         Width           =   300
      End
      Begin VB.CommandButton cmdLoop 
         Caption         =   "q"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   7.5
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   8745
         TabIndex        =   112
         ToolTipText     =   "Loop No"
         Top             =   2925
         Width           =   300
      End
      Begin VB.CommandButton cmdnNext 
         Caption         =   ":"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   7.5
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   10080
         TabIndex        =   111
         ToolTipText     =   "Next"
         Top             =   2925
         Width           =   300
      End
      Begin VB.Frame Frame8 
         Caption         =   "Commands and Settings"
         Height          =   2415
         Left            =   8370
         TabIndex        =   108
         Top             =   60
         Width           =   5415
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   60
            ScaleHeight     =   2175
            ScaleWidth      =   5310
            TabIndex        =   109
            Top             =   195
            Width           =   5310
            Begin VB.TextBox Text1 
               Height          =   1995
               Left            =   75
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   125
               Text            =   "frmMain.frx":23D2
               Top             =   105
               Width           =   5115
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "List MP3/WAV"
         Height          =   3180
         Left            =   60
         TabIndex        =   104
         Top             =   60
         Width           =   8250
         Begin VB.PictureBox Picture9 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   5805
            ScaleHeight     =   255
            ScaleWidth      =   2415
            TabIndex        =   123
            Top             =   2880
            Width           =   2415
            Begin VB.CheckBox CheckPlayAll 
               Caption         =   "Play All Files"
               Height          =   255
               Left            =   435
               TabIndex        =   124
               Top             =   0
               Width           =   1935
            End
         End
         Begin VB.Timer tTimer 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   7590
            Top             =   435
         End
         Begin CDTracker.MP3Play MP3 
            Height          =   690
            Left            =   7410
            TabIndex        =   107
            Top             =   345
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   1217
         End
         Begin VB.ListBox lstmp3wavPath 
            Height          =   270
            Left            =   6390
            TabIndex        =   106
            Top             =   345
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.ListBox lstmp3wav 
            Height          =   2580
            Left            =   75
            TabIndex        =   105
            Top             =   285
            Width           =   8085
         End
         Begin VB.Label lblInfoScan 
            Caption         =   "n/a"
            Height          =   210
            Left            =   75
            TabIndex        =   122
            Top             =   2895
            Width           =   6105
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Burning messages"
         Height          =   1665
         Left            =   60
         TabIndex        =   65
         Top             =   3240
         Width           =   13725
         Begin VB.ListBox lst_Messages 
            Height          =   1320
            Left            =   75
            TabIndex        =   66
            Top             =   255
            Width           =   13575
         End
      End
      Begin VB.Label lblDuration 
         Caption         =   "Total Time 00:00"
         Height          =   225
         Left            =   11925
         TabIndex        =   118
         Top             =   2955
         Width           =   1845
      End
      Begin VB.Label lblPosition 
         Caption         =   "Time: 00:00"
         Height          =   210
         Left            =   10650
         TabIndex        =   117
         Top             =   2955
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdSelectPath 
      Caption         =   "..."
      Height          =   240
      Left            =   13425
      TabIndex        =   63
      ToolTipText     =   "Select the Default path of Extracted Tracks"
      Top             =   1095
      Width           =   495
   End
   Begin VB.TextBox txtDestPath 
      Height          =   255
      Left            =   6555
      Locked          =   -1  'True
      TabIndex        =   62
      ToolTipText     =   "This is the default Path of Extracted Tracks"
      Top             =   1095
      Width           =   6840
   End
   Begin VB.PictureBox PicFrames 
      BorderStyle     =   0  'None
      Height          =   1170
      Index           =   1
      Left            =   45
      ScaleHeight     =   1170
      ScaleWidth      =   6930
      TabIndex        =   61
      Top             =   6690
      Visible         =   0   'False
      Width           =   6930
      Begin VB.CheckBox CheckVBR 
         Caption         =   "VBR"
         Height          =   270
         Left            =   5865
         TabIndex        =   88
         Top             =   405
         Width           =   1035
      End
      Begin VB.CheckBox CheckPrivate 
         Caption         =   "Private"
         Height          =   270
         Left            =   3360
         TabIndex        =   87
         Top             =   405
         Width           =   1155
      End
      Begin VB.CheckBox CheckOriginal 
         Caption         =   "Original"
         Height          =   270
         Left            =   4515
         TabIndex        =   86
         Top             =   405
         Width           =   1245
      End
      Begin VB.CheckBox CheckCopyright 
         Caption         =   "Copyright"
         Height          =   270
         Left            =   4095
         TabIndex        =   85
         Top             =   855
         Width           =   1290
      End
      Begin VB.ComboBox cmbBitRate 
         Height          =   330
         ItemData        =   "frmMain.frx":2541
         Left            =   2850
         List            =   "frmMain.frx":2563
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   810
         Width           =   1170
      End
      Begin VB.CheckBox CheckWriteTag 
         Caption         =   "Write Tag of extract Track"
         Height          =   240
         Left            =   75
         TabIndex        =   79
         Top             =   390
         Width           =   3135
      End
      Begin VB.CheckBox CheckManually 
         Caption         =   "Mod track  Manually"
         Height          =   240
         Left            =   75
         TabIndex        =   74
         Top             =   105
         Width           =   2295
      End
      Begin VB.ComboBox cmbMode 
         Height          =   330
         ItemData        =   "frmMain.frx":25BA
         Left            =   2400
         List            =   "frmMain.frx":25D0
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   60
         Width           =   4485
      End
      Begin VB.CommandButton cmdEncode 
         Caption         =   "&Encoding"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5460
         TabIndex        =   72
         ToolTipText     =   "Start the Encoding"
         Top             =   735
         Width           =   1335
      End
      Begin VB.OptionButton OptEncode 
         Caption         =   "Encode to MP3"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   70
         Top             =   645
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton OptEncode 
         Caption         =   "Encode to WAV"
         Height          =   225
         Index           =   1
         Left            =   75
         TabIndex        =   69
         Top             =   870
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "BitRate:"
         Height          =   240
         Left            =   1875
         TabIndex        =   84
         Top             =   855
         Width           =   975
      End
   End
   Begin VB.PictureBox PicFrames 
      BorderStyle     =   0  'None
      Height          =   1170
      Index           =   0
      Left            =   45
      ScaleHeight     =   1170
      ScaleWidth      =   6930
      TabIndex        =   40
      Top             =   6690
      Width           =   6930
      Begin VB.OptionButton OptionTime 
         Caption         =   "MS"
         Height          =   225
         Index           =   3
         Left            =   4635
         TabIndex        =   59
         ToolTipText     =   "Show only mm:ss"
         Top             =   870
         Width           =   615
      End
      Begin VB.OptionButton OptionTime 
         Caption         =   "MSM"
         Height          =   225
         Index           =   2
         Left            =   3900
         TabIndex        =   58
         ToolTipText     =   "Show only mm:ss:mm"
         Top             =   870
         Width           =   735
      End
      Begin VB.OptionButton OptionTime 
         Caption         =   "T- MS"
         Height          =   225
         Index           =   1
         Left            =   3015
         TabIndex        =   57
         ToolTipText     =   "Show Track- mm:ss"
         Top             =   870
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton OptionTime 
         Caption         =   "T:MSM"
         Height          =   225
         Index           =   0
         Left            =   2100
         TabIndex        =   56
         ToolTipText     =   "Show Track- mm:ss:mm"
         Top             =   870
         Width           =   870
      End
      Begin VB.CommandButton cmdOpenClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   7.5
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   75
         TabIndex        =   55
         ToolTipText     =   "Open/Close Device"
         Top             =   555
         Width           =   300
      End
      Begin CDTracker.utcWait utcWait 
         Height          =   360
         Left            =   1650
         TabIndex        =   54
         Top             =   270
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
      End
      Begin VB.Timer TCD 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   5400
         Top             =   -315
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   735
         TabIndex        =   52
         ToolTipText     =   "Play"
         Top             =   840
         Width           =   300
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   51
         ToolTipText     =   "Stop"
         Top             =   840
         Width           =   300
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   405
         TabIndex        =   50
         ToolTipText     =   "Pause"
         Top             =   840
         Width           =   300
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   49
         ToolTipText     =   "Prev Track"
         Top             =   840
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   48
         ToolTipText     =   "Next Track"
         Top             =   840
         Width           =   345
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   810
         Left            =   2085
         ScaleHeight     =   810
         ScaleWidth      =   4800
         TabIndex        =   41
         Top             =   60
         Width           =   4800
         Begin CDTracker.utcCDW CDW 
            Left            =   2850
            Top             =   -375
            _ExtentX        =   847
            _ExtentY        =   847
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "LcdD"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   225
            Left            =   3075
            TabIndex        =   47
            ToolTipText     =   "Time Track"
            Top             =   15
            Width           =   1125
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 "
            BeginProperty Font 
               Name            =   "LcdD"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   360
            Left            =   4125
            TabIndex        =   46
            ToolTipText     =   "Total Tracks"
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Lbrano 
            BackStyle       =   0  'Transparent
            Caption         =   "track 0"
            BeginProperty Font 
               Name            =   "LcdD"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1770
            TabIndex        =   45
            ToolTipText     =   "Selected Tracks"
            Top             =   15
            Width           =   1170
         End
         Begin VB.Label statuslabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Standby"
            BeginProperty Font 
               Name            =   "LcdD"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   45
            TabIndex        =   44
            ToolTipText     =   "Status CD"
            Top             =   15
            Width           =   1650
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "1- 00:00 "
            BeginProperty Font 
               Name            =   "LcdD"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   405
            Left            =   1920
            TabIndex        =   43
            ToolTipText     =   "Time"
            Top             =   285
            Width           =   2760
         End
         Begin VB.Label lblTAG 
            Caption         =   "TAGS-ARTIST"
            BeginProperty Font 
               Name            =   "LcdD"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   60
            TabIndex        =   42
            ToolTipText     =   "Track Artist"
            Top             =   345
            Width           =   2175
         End
      End
   End
   Begin ComctlLib.TabStrip TBS2 
      Height          =   1530
      Left            =   30
      TabIndex        =   39
      Top             =   6360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2699
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CDs Player"
            Key             =   "CDPlayer"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Internal CD-W/DVDW Player"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Encode && Opiton"
            Key             =   "EncodeOption"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Option Encoded CD-W/DVDW"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced Tags"
            Key             =   "AdvancedTags"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Advanced ID3-MP3 Tags"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTemp 
      Height          =   270
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Top             =   45
      Visible         =   0   'False
      Width           =   690
   End
   Begin InetCtlsObjects.Inet inetConnexion 
      Left            =   210
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   4890
      Index           =   2
      Left            =   60
      ScaleHeight     =   4890
      ScaleWidth      =   13830
      TabIndex        =   10
      Top             =   1410
      Visible         =   0   'False
      Width           =   13830
      Begin VB.TextBox txtTextLog 
         Height          =   4455
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   405
         Width           =   13665
      End
      Begin VB.Label lblPath 
         Caption         =   "n/a"
         Height          =   270
         Left            =   75
         TabIndex        =   12
         Top             =   135
         Width           =   13650
      End
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   4875
      Index           =   0
      Left            =   60
      ScaleHeight     =   4875
      ScaleWidth      =   13830
      TabIndex        =   6
      Tag             =   "http://freedb.freedb.org/~cddb/cddb.cgi"
      Top             =   1425
      Width           =   13830
      Begin VB.Frame Frame4 
         Caption         =   "Data CD"
         Height          =   3090
         Left            =   6930
         TabIndex        =   21
         Top             =   1770
         Width           =   6855
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   3840
            ScaleHeight     =   300
            ScaleWidth      =   1095
            TabIndex        =   100
            Top             =   2745
            Width           =   1095
            Begin VB.CommandButton cmdSaveCover 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   240
               Left            =   450
               TabIndex        =   101
               Top             =   30
               Width           =   540
            End
         End
         Begin VB.CheckBox CheckAlbum 
            Caption         =   "Download Album Art"
            Height          =   240
            Left            =   930
            TabIndex        =   75
            Top             =   2775
            Width           =   3375
         End
         Begin utcCoverDownload.utcCover utcCover 
            Height          =   1815
            Left            =   4980
            TabIndex        =   60
            ToolTipText     =   "Cover Album/Artist"
            Top             =   1260
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   3201
            PicType         =   3
            MouseLeave      =   0   'False
         End
         Begin VB.TextBox txtComment 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   945
            TabIndex        =   36
            Top             =   2085
            Width           =   4005
         End
         Begin VB.TextBox txtTrack 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6060
            TabIndex        =   35
            Top             =   885
            Width           =   690
         End
         Begin VB.TextBox txtBand 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   945
            TabIndex        =   32
            Top             =   1710
            Width           =   4005
         End
         Begin VB.TextBox txtTitle 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   945
            TabIndex        =   31
            Top             =   1350
            Width           =   4005
         End
         Begin VB.TextBox txtGenre 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2700
            TabIndex        =   28
            Top             =   885
            Width           =   2235
         End
         Begin VB.TextBox txtYear 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   945
            TabIndex        =   26
            Top             =   885
            Width           =   975
         End
         Begin VB.TextBox txtAlbum 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   945
            TabIndex        =   24
            Top             =   525
            Width           =   5805
         End
         Begin VB.TextBox txtArtist 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   945
            TabIndex        =   22
            Top             =   165
            Width           =   5805
         End
         Begin VB.Label Label11 
            Caption         =   "Comment:"
            Height          =   255
            Left            =   105
            TabIndex        =   37
            Top             =   2100
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Track:"
            Height          =   240
            Left            =   5265
            TabIndex        =   34
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label9 
            Caption         =   "Band:"
            Height          =   255
            Left            =   105
            TabIndex        =   33
            Top             =   1755
            Width           =   705
         End
         Begin VB.Label Label8 
            Caption         =   "Title:"
            Height          =   255
            Left            =   105
            TabIndex        =   30
            Top             =   1410
            Width           =   705
         End
         Begin VB.Label Label7 
            Caption         =   "Genre:"
            Height          =   255
            Left            =   2025
            TabIndex        =   29
            Top             =   930
            Width           =   705
         End
         Begin VB.Label Label6 
            Caption         =   "Year:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   930
            Width           =   705
         End
         Begin VB.Label Label5 
            Caption         =   "Album:"
            Height          =   255
            Left            =   105
            TabIndex        =   25
            Top             =   555
            Width           =   705
         End
         Begin VB.Label Label3 
            Caption         =   "Artist:"
            Height          =   255
            Left            =   105
            TabIndex        =   23
            Top             =   210
            Width           =   705
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Response"
         Height          =   1650
         Left            =   6930
         TabIndex        =   17
         Top             =   90
         Width           =   6855
         Begin VB.ListBox lstResponse 
            Height          =   1320
            Left            =   135
            TabIndex        =   18
            Top             =   255
            Width           =   6660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4110
         Left            =   75
         TabIndex        =   13
         Top             =   750
         Width           =   6810
         Begin VB.ListBox lsTracks 
            Height          =   3660
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   71
            Top             =   390
            Visible         =   0   'False
            Width           =   6675
         End
         Begin MSComDlg.CommonDialog cDialog 
            Left            =   5595
            Top             =   -345
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ComctlLib.ListView lstTracks 
            Height          =   3540
            Left            =   75
            TabIndex        =   14
            Top             =   480
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   6244
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Title of Track"
               Object.Width           =   7832
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Track"
               Object.Width           =   2647
            EndProperty
         End
         Begin VB.Label lblTop 
            BackColor       =   &H8000000C&
            Caption         =   " Play Title                                  Track"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   60
            TabIndex        =   78
            Top             =   135
            Width           =   6675
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Driver"
         Height          =   660
         Left            =   75
         TabIndex        =   7
         Top             =   90
         Width           =   6810
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   4935
            ScaleHeight     =   480
            ScaleWidth      =   1830
            TabIndex        =   19
            Top             =   135
            Width           =   1830
            Begin VB.CommandButton cmdReload 
               Caption         =   "CDs"
               Height          =   375
               Left            =   60
               TabIndex        =   53
               ToolTipText     =   "Re-Query CDs List"
               Top             =   60
               Width           =   810
            End
            Begin VB.CommandButton cmdQuery 
               Caption         =   "Query"
               Height          =   375
               Left            =   930
               TabIndex        =   20
               ToolTipText     =   "Get Query CD"
               Top             =   60
               Width           =   855
            End
         End
         Begin VB.ComboBox Combo_Dispositivo 
            Height          =   330
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   3960
         End
         Begin VB.ComboBox Combo_Lettera 
            Height          =   330
            Left            =   4065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   840
         End
      End
   End
   Begin ComctlLib.TabStrip TBS 
      Height          =   5280
      Left            =   30
      TabIndex        =   5
      Top             =   1080
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   9313
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CD Extract"
            Key             =   "CDExtract"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Extract Audio Track's from CD"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Burn to CD/DVD"
            Key             =   "Burn"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Burn MP3 Files or oter Files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Log query CD List"
            Key             =   "LogCDList"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   195
      Left            =   8880
      TabIndex        =   3
      Top             =   7950
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label21 
      Caption         =   $"frmMain.frx":2697
      Height          =   1350
      Left            =   7140
      TabIndex        =   126
      Top             =   6465
      Width           =   5325
   End
   Begin VB.Label lblEncoding 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11820
      TabIndex        =   77
      Top             =   7950
      Width           =   2085
   End
   Begin VB.Label lblTagManually 
      Height          =   210
      Left            =   5010
      TabIndex        =   76
      Top             =   150
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CD Tracker Extract"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   1080
      TabIndex        =   68
      Top             =   45
      Width           =   3210
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CD Tracker Extract"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   480
      Left            =   1050
      TabIndex        =   67
      Top             =   75
      Width           =   3210
   End
   Begin VB.Label Label_Info 
      BackColor       =   &H00FFFFFF&
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   210
      Left            =   60
      TabIndex        =   16
      Top             =   7950
      Width           =   7320
   End
   Begin VB.Label lblCDID 
      BackColor       =   &H00FFFFFF&
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   210
      Left            =   7425
      TabIndex        =   15
      Top             =   7950
      Width           =   1350
   End
   Begin VB.Label labelBotton 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   285
      Left            =   -15
      TabIndex        =   4
      Top             =   7890
      Width           =   13980
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "n/a"
      Height          =   210
      Left            =   1005
      TabIndex        =   2
      Top             =   750
      Width           =   7395
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.3b"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   300
      Left            =   4035
      TabIndex        =   1
      Top             =   195
      Width           =   1005
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.3b"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4050
      TabIndex        =   0
      Top             =   165
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Picture         =   "frmMain.frx":279F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   945
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   -330
      Picture         =   "frmMain.frx":2EB5
      Top             =   -195
      Width           =   14325
   End
   Begin VB.Shape shapBack 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   1005
      Left            =   0
      Top             =   -15
      Width           =   13995
   End
End
Attribute VB_Name = "frmMain"
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

Dim strMediaFile As String

Dim pPos As Integer
Dim strLoop As Boolean

Dim CoverOK As Boolean
Dim CoverDir As String

' variables Nero
Dim Source_Dir As String
Dim FSO As New FileSystemObject
Dim DateFolder As NeroFolder
Dim sFile As NeroFile
Dim rootfolder As NeroFolder
Dim ISOTrack As NeroISOTrack
Dim CDStamp As NeroCDStamp

' load NERO references
Public WithEvents Nero As Nero
Attribute Nero.VB_VarHelpID = -1
Public WithEvents Drive As NeroDrive
Attribute Drive.VB_VarHelpID = -1

Dim Cnt As Integer
Dim IsDriveWriteable As Boolean
Dim DriveMediaType As String
Dim NumExistingTracks As Integer
Dim DriveFinished As Boolean
Dim Drives As NeroDrives
Dim CancelPressed As Boolean
Dim BurnError As Boolean

Dim ABORT_ENCODING As Boolean

Dim modeTrack As Boolean

Private ssLeft As Long
Private ssTop As Long

Private tbsKey As String
Private sParse As Boolean
Private VetDispositivi() As String
Private TotDispositivi As Long

Private II_ndex As Integer
Private OpenDevice As Boolean
Private asWorking As Boolean

Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

Private Enum sSetIcon
    ICON_PROGRAM = 0
    ICON_NOTE = 1
    ICON_LOG = 2
    ICON_AUDIO = 3
    ICON_CDLIST = 4
    ICON_TODO = 5
    ICON_DOG = 6
End Enum

Private MyString As String

Private Enum Extract
  [Only_Extension] = 0
  [Only_FileName_and_Extension] = 1
  [Only_FileName_no_Extension] = 2
  [Only_Path] = 3
End Enum

Private strURL As String
Private strHello As String
Private strCategorie As String
Private strCDiD As String
Private strQuery As String

Private Type sTracksTypes
    Title   As String
    Autor  As String
End Type

Private TTracks() As sTracksTypes

Private Const SW_SHOWNORMAL As Long = 1

' .... For the ListView
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim itmX As ListItem

Private Const LVM_FIRST = &H1000
Private Const LVIF_STATE = &H8

Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_TRACKSELECT = &H8
Private Const LVS_EX_ONECLICKACTIVATE = &H40
Private Const LVS_EX_TWOCLICKACTIVATE = &H80
Private Const LVS_EX_SUBITEMIMAGES = &H2
 
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVIS_STATEIMAGEMASK As Long = &HF000


' .... Variables Encode/Decode
Private Const IOCTL_CDROM_READ_TOC      As Long = &H24000
Private Const IOCTL_CDROM_RAW_READ      As Long = &H2403E
Private Const Drive_CDROM               As Long = 5

Private Const RAW_SECTOR_SIZE           As Long = 2352
Private Const LARGEST_SECTORS_PER_READ  As Long = 27
Private Const SAMPLES_PER_SECTOR        As Long = RAW_SECTOR_SIZE / 4

Private Const GENERIC_READ              As Long = &H80000000
Private Const FILE_SHARE_READ           As Long = &H1
Private Const OPEN_EXISTING             As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80

Private Enum TRACK_MODE_TYPE
    YellowMode2
    XAForm2
    CDDA
End Enum

Private Type WAVCHUNKHEADER
    ChunkID As String * 4
    ChunkSize As Long
End Type

Private Type WAVCHUNKFORMAT
    wFormatTag As Integer
    wChannels As Integer
    dwSamplesPerSec As Long
    dwAvgBytesPerSec As Long
    wBlockAlign As Integer
    wBitsPerSample   As Integer
End Type

Private Type TRACK_DATA
    Reserved As Byte
    Adr As Byte
    TrackNumber As Byte
    Reserved1 As Byte
    Address(3) As Byte
End Type

Private Type CDROM_TOC
    length(1) As Byte
    FirstTrack As Byte
    LastTrack As Byte
    TrackData(99) As TRACK_DATA
End Type

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type RAW_READ_INFO
    DiskOffset  As LARGE_INTEGER
    SectorCount As Long
    TrackMode   As TRACK_MODE_TYPE
End Type

Private mfile As Long
Private mSize As Long
Private hwo As Long
Private sndofs As Long

Private Declare Sub MemoryCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long

' .... Special Folders
' USAGE= Dir$(GetSpecialFolderLocation(CSIDL_SYSTEM) & "\")
' .... Converts an Item identifier list to a file System Path.
Private Const MAX_PATH As Long = 260
Private Const S_OK = 0

Private Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_INTERNET = &H1
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const CSIDL_COMMON_STARTMENU = &H16
Private Const CSIDL_COMMON_PROGRAMS = &H17
Private Const CSIDL_COMMON_STARTUP = &H18
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Private Const CSIDL_APPDATA = &H1A
Private Const CSIDL_PRINTHOOD = &H1B
Private Const CSIDL_LOCAL_APPDATA = &H1C
Private Const CSIDL_ALTSTARTUP = &H1D
Private Const CSIDL_COMMON_ALTSTARTUP = &H1E
Private Const CSIDL_COMMON_FAVORITES = &H1F
Private Const CSIDL_INTERNET_CACHE = &H20
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_HISTORY = &H22
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const CSIDL_WINDOWS = &H24
Private Const CSIDL_SYSTEM = &H25
Private Const CSIDL_PROGRAM_FILES = &H26
Private Const CSIDL_MYPICTURES = &H27
Private Const CSIDL_PROFILE = &H28
Private Const CSIDL_SYSTEMX86 = &H29
Private Const CSIDL_PROGRAM_FILESX86 = &H2A
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B
Private Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
Private Const CSIDL_COMMON_TEMPLATES = &H2D
Private Const CSIDL_COMMON_DOCUMENTS = &H2E
Private Const CSIDL_COMMON_ADMINTOOLS = &H2F
Private Const CSIDL_ADMINTOOLS = &H30
Private Const CSIDL_FLAG_CREATE = &H8000&
Private Const CSIDL_FLAG_DONT_VERIFY = &H4000
Private Const CSIDL_FLAG_MASK = &HFF00
Private Const SHGFP_TYPE_CURRENT = &H0
Private Const SHGFP_TYPE_DEFAULT = &H1

Private Sub CheckManually_Click()
    On Local Error Resume Next
    If CheckManually.Value = 1 Then cmbMode.ListIndex = 5 Else cmbMode.ListIndex = 2
End Sub

Private Sub cmbMode_Click()
    On Local Error Resume Next
    If cmbMode.ListIndex = 4 Then
            cmbMode.ListIndex = 2
        Exit Sub
    End If
    If cmbMode.ListIndex = 5 Then CheckManually.Value = 1 Else CheckManually.Value = 0
    ' .... Save track mode
    INI.DeleteKey "SETTING", "SAVE_TRACK_MODE"
    INI.CreateKeyValue "SETTING", "SAVE_TRACK_MODE", cmbMode.ListIndex
End Sub


Private Sub cmdBrowseMedia_Click()
    Dim strFolder As String
    If txtMediaPath.Text <> "" Then
        If MsgBox("Scan the Default media Path?", vbYesNo + vbInformation + _
        vbDefaultButton1, App.Title) = vbYes Then
        cmdBrowseMedia.Enabled = False
        Call AllFileInFolder(txtMediaPath.Text, True)
    End If
    Else
    strFolder = BrowseFolder("Select the folder that contains the media files:", App.Path)
    If strFolder <> "" And strFolder <> "Error!" Then
        txtMediaPath.Text = strFolder + "\"
    ' .... Save default Media path
    INI.DeleteKey "SETTING", "DEFAULT_MEDIA_PATH"
    INI.CreateKeyValue "SETTING", "DEFAULT_MEDIA_PATH", txtMediaPath.Text
    cmdBrowseMedia.Enabled = False
        Call AllFileInFolder(strFolder, True)
    End If
    End If
    If lstmp3wav.ListCount > 0 Then lstmp3wav.Selected(0) = True
    cmdBrowseMedia.Enabled = True
End Sub

Private Sub cmdEncode_Click()
    Dim i As Integer
    Dim XX As Integer
    Dim KK As Integer
    Dim strFolder As String
    Dim Msg As String
    Dim sArtist As String
    Dim sAlbum As String
    Dim sYear As String
    Dim sInput As String
    Dim sTitleOfTrack As String
    Dim sSkip As Integer
    
    If (Combo_Lettera.ListIndex < 0) Then Exit Sub
    If (lsTracks.ListCount < 0) Or (lsTracks.Visible = False) Then Exit Sub
    
    On Local Error GoTo ErrorHandler
    
    For i = 0 To lsTracks.ListCount - 1
        If lsTracks.Selected(i) = True And lsTracks.List(i) <> "" Then XX = XX + 1
    Next i
    
    ' .... I no tracks selected then warning and exit
    If XX = 0 Then
            MsgBox "No tracks selected! Please select one track before.", vbExclamation, App.Title
        Exit Sub
    End If
    
    i = 0
    XX = 0
    
    ' .... Stop the Encoding?
    If cmdEncode.Caption = "&Abort" Then
            If MsgBox("Are you sure to STOP the Encoding?", vbYesNo + vbExclamation + _
        vbDefaultButton2, "Stop Encoding") = vbYes Then
            ABORT_ENCODING = True
                cmdEncode.Caption = "&Encoding"
            cmdEncode.ToolTipText = "Start the Encoding"
        Else
            Exit Sub
        End If
    Else
    
    ' .... Init Button and Flag
    ABORT_ENCODING = False
    cmdEncode.Caption = "&Abort"
    cmdEncode.ToolTipText = "Abort current Encoding"
    
    ' .... Add the encoded tracks to this List...
    ' .... this List only to remuve the incomplete Files ;)
    lstMedia.Clear
    
    ' .... If default Path is nothing display Browser for Folder
    If txtDestPath.Text = "" Then
        strFolder = BrowseFolder("Extract tracks to:", App.Path)
        If strFolder <> "" And strFolder <> "Error!" Then
            txtDestPath.Text = strFolder + "\"
        Else
            ' .... the Default Folder is necessary
                cmdEncode.Caption = "&Encoding"
                cmdEncode.ToolTipText = "Start the Encoding"
            Exit Sub
        End If
    Else
        strFolder = txtDestPath.Text
    End If
    
    If Not FSO.FolderExists(strFolder) Then
        ' .... Create folder CD Query
        If MakeDirectory(strFolder + "CD Query") = False Then:
        ' .... Until display the Error now, because if the Folder exist return a Error ;)
    End If
    
    ' .... Create the SubFolder of the Tracks
    sArtist = txtArtist.Text: sAlbum = txtAlbum.Text: sYear = txtYear.Text
    If sYear = Empty Then sYear = Format(Now, "mm-yyyy")
    
    If sArtist = Empty Or sAlbum = Empty Then
        sInput = InputBox("Enter the name of Folder to save the extracted Tracks:", App.Title, App.Path)
        If sInput = Empty Then
                MsgBox "The Folder to save the extracted tracks is necessary!", vbExclamation, App.Title
                    cmdEncode.Caption = "&Encoding"
                cmdEncode.ToolTipText = "Start the Encoding"
            Exit Sub
        End If
        strFolder = strFolder + "CD Query\" + sInput
    Else
        strFolder = strFolder + "CD Query\" + txtArtist.Text + "\" + txtArtist.Text + "-" + txtAlbum.Text + "-" + txtYear.Text
    End If
    
    ' .... Create the SubFolder
    If MakeDirectory(strFolder) = False Then: ' .... Until display the Error!
    
    ' .... if the Folder not exist Exit
    If Not FSO.FolderExists(strFolder) Then
            MsgBox "Error - Source folder does not Exist!", vbExclamation, App.Title
                cmdEncode.Caption = "&Encoding"
            cmdEncode.ToolTipText = "Start the Encoding"
        Exit Sub
    End If
    
    ' .... Put manually the title of Track?
    ' .... I leave the User to decide the mode of the title of extracted tracks (sorry for my english) :)
    If cmbMode.ListIndex = 5 Or CheckManually.Value = 1 Then
        If lblTagManually.Caption = Empty Then sInput = "year|track|title|album|artist" Else sInput = lblTagManually.Caption
        sInput = InputBox("Enter the String Mode to save the extracted tracks separated by comma: (|)" & vbCr & vbCr _
            & "Example: year|track|title|album|artist", App.Title, sInput)
        If sInput = Empty Then
                MsgBox "The mode to save the extracted tracks is necessary!", vbExclamation, App.Title
                    cmdEncode.Caption = "&Encoding"
                cmdEncode.ToolTipText = "Start the Encoding"
            Exit Sub
        End If
        
        ' .... Verify the immission
        If Mid$(sInput, 1, 1) = "." Or Mid$(sInput, 1, 1) = "-" Then
                MsgBox "The title of the Track is incorrect!" & vbCr & "Verify your typed please." & vbCr & vbCr _
                & "Incorrect char (" & Mid$(sInput, 1, 1) & ").", vbExclamation, App.Title
                    cmdEncode.Caption = "&Encoding"
                cmdEncode.ToolTipText = "Start the Encoding"
            Exit Sub
        End If
        
        ' .... Split the string and put the title in the right order
        sTitleOfTrack = SplitTrack(sInput, "|")
        
        ' .... If error then
        If sTitleOfTrack = "Error!" Or sTitleOfTrack = Empty Then
                MsgBox "Error to parse your string!" & vbCr & "Make sure you typed the correct parameters! Try again please.", vbExclamation, App.Title
                    cmdEncode.Caption = "&Encoding"
                cmdEncode.ToolTipText = "Start the Encoding"
            Exit Sub
        End If
        
        ' .... Checked the Box manually
        CheckManually.Value = 1
        
        ' .... Save the Manually Tags
        If lblTagManually.Caption <> Empty Then
            INI.DeleteKey "SETTING", "SAVE_TRACK_MODE_MANUALLY_TAGs"
            INI.CreateKeyValue "SETTING", "SAVE_TRACK_MODE_MANUALLY_TAGs", lblTagManually.Caption
        End If
    End If
    
    ' .... Enum the Tracks to Encode
    For i = 0 To lsTracks.ListCount - 1
        If lsTracks.Selected(i) = True And lsTracks.List(i) <> "" Then XX = XX + 1
    Next i
    
    ' .... Reset var
    KK = 0
    i = 0

    lstmp3wavPath.Clear
    lstmp3wav.Clear
    
    asWorking = True
    cmdExit.Enabled = False
    TBS.Enabled = False
    TBS2.Enabled = False
    
    ' Start Encoding
    For i = 0 To lsTracks.ListCount - 1
        If lsTracks.Selected(i) = True And lsTracks.List(i) <> "" Then
            ' .... display the line of List
            lsTracks.ListIndex = i
            KK = KK + 1
            lblEncoding.Caption = "Track " & KK & " to " & XX
            DoEvents
            If OptEncode(0).Value Then
                ' .... Encode to MP3
                
                If modeTrack = False Then ' .... Info Artist Found, parse the Title of Track
                
                txtTitle.Text = StripLeft(lsTracks.List(i), ":", True)
                txtTrack.Text = Mid$(StripLeft(lsTracks.List(i), ":", False), 7, Len(StripLeft(lsTracks.List(i), ":", False)))
                
                If CheckManually.Value = 1 Then
                    sTitleOfTrack = GetTitleTrack(lblTagManually.Caption)
                Else
                ' .... Mode save Track
                Select Case cmbMode.ListIndex
                    Case 0  '.... <Track>.<Artist>-<Album>-<Title>-<Year>
                        sTitleOfTrack = txtTrack.Text + "." + txtArtist.Text + "-" + txtAlbum.Text + "-" _
                        + StripLeft(lsTracks.List(i), ":", True) + "-" + txtYear.Text
                    Case 1  ' .... <Track>.<Title>-<Album>-<Artist>-<Year>
                        sTitleOfTrack = txtTrack.Text + "." + StripLeft(lsTracks.List(i), ":", True) + "-" + txtAlbum.Text + _
                        "-" + txtArtist.Text + "-" + txtYear.Text
                    Case 2  ' .... <Artist>-<Album>-<Track>.<Title>-<Year>
                        sTitleOfTrack = txtArtist.Text + "-" + txtAlbum.Text + "-" + txtTrack.Text + "." _
                        + StripLeft(lsTracks.List(i), ":", True) + "-" + txtYear.Text
                    Case 3  ' .... <Album>-<Artist>-<Track>.<Title>-<Year>
                        sTitleOfTrack = txtAlbum.Text + "-" + txtArtist.Text + "." + txtTrack.Text + "." + _
                        StripLeft(lsTracks.List(i), ":", True) + "-" + txtYear.Text
                End Select
                End If
                
                Else ' .... No info Artist Found, parse only Track number
                    sTitleOfTrack = StripLeft(lsTracks.List(i), ":", True)
                End If
                
                lstMedia.AddItem strFolder + sTitleOfTrack + ".mp3"
                lstmp3wavPath.AddItem strFolder
                lstmp3wav.AddItem sTitleOfTrack + ".mp3"
                
                
                ' .... Start Encoding to MP3
                If CDWDVDWToMp3(Combo_Lettera.List(Combo_Lettera.ListIndex), i, strFolder + sTitleOfTrack + ".mp3") = _
                                False Then sSkip = sSkip + 1
            ' .... Encode to WAV
            ElseIf OptEncode(1).Value Then
                
                If modeTrack = False Then ' .... Info Artist Found, parse the Title of Track
                
                txtTitle.Text = StripLeft(lsTracks.List(i), ":", True)
                txtTrack.Text = Mid$(StripLeft(lsTracks.List(i), ":", False), 7, Len(StripLeft(lsTracks.List(i), ":", False)))
                
                If CheckManually.Value = 1 Then
                    sTitleOfTrack = GetTitleTrack(lblTagManually.Caption)
                Else
                ' .... Mode save Track
                Select Case cmbMode.ListIndex
                    Case 0  '.... <Track>.<Artist>-<Album>-<Title>-<Year>
                        sTitleOfTrack = txtTrack.Text + "." + txtArtist.Text + "-" + txtAlbum.Text + "-" _
                        + StripLeft(lsTracks.List(i), ":", True) + "-" + txtYear.Text
                    Case 1  ' .... <Track>.<Title>-<Album>-<Artist>-<Year>
                        sTitleOfTrack = txtTrack.Text + "." + StripLeft(lsTracks.List(i), ":", True) + "-" + txtAlbum.Text + _
                        "-" + txtArtist.Text + "-" + txtYear.Text
                    Case 2  ' .... <Artist>-<Album>-<Track>.<Title>-<Year>
                        sTitleOfTrack = txtArtist.Text + "-" + txtAlbum.Text + "-" + txtTrack.Text + "." _
                        + StripLeft(lsTracks.List(i), ":", True) + "-" + txtYear.Text
                    Case 3  ' .... <Album>-<Artist>-<Track>.<Title>-<Year>
                        sTitleOfTrack = txtAlbum.Text + "-" + txtArtist.Text + "." + txtTrack.Text + "." + _
                        StripLeft(lsTracks.List(i), ":", True) + "-" + txtYear.Text
                End Select
                End If
                
                 Else ' .... No info Artist Found, parse only Track number
                    sTitleOfTrack = StripLeft(lsTracks.List(i), ":", True)
                End If
                
                
                lstMedia.AddItem strFolder + sTitleOfTrack + ".wav"
                lstmp3wavPath.AddItem strFolder
                lstmp3wav.AddItem sTitleOfTrack + ".mp3"
                
                ' .... Start Encoding
                If CDWDVDWToWav(Combo_Lettera.List(Combo_Lettera.ListIndex), i, strFolder + sTitleOfTrack + ".wav") = _
                            False Then sSkip = sSkip + 1
            End If
        End If
        lsTracks.Selected(i) = False
        DoEvents
        If ABORT_ENCODING = True Then Exit For
    Next i
    
    cmdEncode.Caption = "&Encoding"
    cmdEncode.ToolTipText = "Start the Encoding"
    
    If ABORT_ENCODING = True Then
        If lstMedia.ListCount > 0 And Dir$(lstMedia.List(lstMedia.ListCount - 1)) <> "" Then
            If MsgBox("The Encoding was not completed because stopped by User!" & vbCr _
            & "You want to Delete the incomplete File:" & vbCr & GetFilePath(lstMedia.List(lstMedia.ListCount - 1), Only_FileName_and_Extension), vbYesNo + vbExclamation + _
                vbDefaultButton2, "Encoding Stopped!") = vbYes Then
                    Call Kill(lstMedia.List(lstMedia.ListCount - 1))
                        lstMedia.RemoveItem (lstMedia.ListCount - 1)
                        lstmp3wavPath.RemoveItem (lstmp3wavPath.ListCount - 1)
                        lstmp3wav.RemoveItem (lstmp3wav.ListCount - 1)
                        MsgBox "The incomplete file was successfully removed!", vbInformation, App.Title
                    Exit Sub
                Else
                    Exit Sub
            End If
        End If
    End If
End If
    
    PB.Value = 0
    
    ' .... Save the Cover to Root Folder
    If utcCover.CoverFound Then
        utcCover.SaveCoverAs "Cover." + txtAlbum.Text + "-" + txtArtist.Text, strFolder, Qlt_80, True
        If Dir$(strFolder + "Cover." + txtAlbum.Text + "-" + txtArtist.Text + ".jpg") <> "" Then
            CoverOK = True
            CoverDir = strFolder + "Cover." + txtAlbum.Text + "-" + txtArtist.Text + ".jpg"""
        Else
            CoverOK = False
        End If
    End If
    
    Msg = "Encoding finish!" & vbCr & vbCr & "Total Encoding: " & XX & " to: " & KK & " Skip: " & sSkip
    
    ' .... Display the message ;)
    lblEncoding.Caption = "Encodin finish!"
    MsgBox Msg, vbInformation, App.Title
    cmdEncode.Enabled = True
    
    asWorking = False
    cmdExit.Enabled = True
    TBS.Enabled = True
    TBS2.Enabled = True
    
    Exit Sub

ErrorHandler:
        Call WriteErrorLogs(Err.Number, Err.Description, "{cmdEncode}", True, True)
    Err.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLoop_Click()
    If strLoop = False Then
        strLoop = True
        cmdLoop.ToolTipText = "Loop yes"
        CheckPlayAll.Value = 0
        CheckPlayAll.Enabled = False
    Else
        strLoop = False
        cmdLoop.ToolTipText = "Loop no"
        CheckPlayAll.Enabled = True
    End If
End Sub

Private Sub cmdNext_Click()
    Dim p As Integer
    On Error Resume Next
    If lstTracks.ListItems.Count = 0 Then
            MsgBox "Get Query CD first!", vbExclamation, App.Title
        Exit Sub
    End If
    If OpenDevice = False Then cmdOpenClose = True
    If lstTracks.SelectedItem.Index = lstTracks.ListItems.Count Then
        p = 1
    Else
        p = lstTracks.SelectedItem.Index + 1
    End If
    lstTracks.ListItems(p).Selected = True
    cmdStop = True
    cmdPlay = True
    lstTracks.SetFocus
End Sub

Private Sub cmdnNext_Click()
    Dim p As Integer
    On Local Error Resume Next
    If lstmp3wav.ListCount = 0 Then Exit Sub
    
    If lstmp3wav.ListIndex + 1 >= lstmp3wav.ListCount Then
        p = 0
    ElseIf lstmp3wav.ListIndex < lstmp3wav.ListCount Then
        p = lstmp3wav.ListIndex + 1
    End If
    
    lstmp3wav.Selected(p) = True
    
    cmdpStop = True
    cmdpPlay = True
End Sub

Private Sub cmdnpPrev_Click()
    On Local Error Resume Next
    Dim p As Integer
    
    If lstmp3wav.ListCount = 0 Then Exit Sub
    
    If lstmp3wav.ListIndex + 1 = lstmp3wav.ListCount Then
        p = lstmp3wav.ListIndex - 1
    Else
        If lstmp3wav.ListIndex - 1 = "-1" Then
            p = lstmp3wav.ListCount - 1
        Else
            p = lstmp3wav.ListIndex - 1
        End If
    End If

    lstmp3wav.Selected(p) = True
    
    cmdpStop = True
    cmdpPlay = True
End Sub


Private Sub cmdOpenClose_Click()
If lstTracks.ListItems.Count > 0 Then
    If OpenDevice = False Then
        InitCDAUDIO True
    ElseIf OpenDevice Then
        InitCDAUDIO False
        utcWait.Start = False
    End If
Else
    MsgBox "Get Query CD first!", vbExclamation, App.Title
End If
End Sub

Private Sub cmdPause_Click()
On Local Error Resume Next
    If lstTracks.ListItems.Count = 0 Then
            MsgBox "Get Query CD first!", vbExclamation, App.Title
        Exit Sub
    End If
    If OpenDevice = False Then cmdOpenClose = True
If CDW.IsPlaying Then
    If CDW.PauseCD Then
        statuslabel.Caption = "paused"
        TCD.Enabled = False
        utcWait.Start = False
    Else
        statuslabel.Caption = "error!"
    End If
    Else
        CDW.PlayCD
        statuslabel.Caption = "resumed"
        TCD.Enabled = True
        utcWait.Start = True
    End If
End Sub

Private Sub cmdPlay_Click()
    On Local Error Resume Next
    If lstTracks.ListItems.Count = 0 Then
            MsgBox "Get Query CD first!", vbExclamation, App.Title
        Exit Sub
    End If
    If OpenDevice = False Then cmdOpenClose = True
    CDW.SetCurrentTrack (CLng(str(lstTracks.SelectedItem.Index)))
    Lbrano.Caption = "track " & CLng(str(lstTracks.SelectedItem.Index))
    Label18.Caption = CDW.GetTrackLength(CLng(lstTracks.SelectedItem.Index))
    Label12.Caption = CLng(str(lstTracks.SelectedItem.Index)) & " - 00:00 "
    If CDW.PlayCD Then
        statuslabel.Caption = "playing"
        TCD.Enabled = True
        utcWait.Start = True
    Else
        statuslabel.Caption = "error!"
    End If
End Sub

Private Sub cmdpPause_Click()
    On Local Error Resume Next
    If MP3.IsPlaying = True Then
        cmdpPlay.Enabled = False
        cmdpStop.Enabled = True
        MP3.mmPause
        tTimer.Enabled = False
    Else
        MP3.mmPause
        tTimer.Enabled = True
    End If
    If lstmp3wav.ListCount > 1 Then
        cmdnNext.Enabled = True
        cmdnpPrev.Enabled = True
    ElseIf lstmp3wav.ListCount = 1 Then
        cmdnNext.Enabled = False
        cmdnpPrev.Enabled = False
    End If
End Sub

Private Sub cmdpPlay_Click()
    On Local Error Resume Next
    TBS.Enabled = False
    TBS2.Enabled = False
    tTimer.Enabled = False
    MP3.mmStop
    MP3.FileName = strMediaFile
    MP3.mmPlay
    lblDuration = "Total Time: " & MP3.length
    tTimer.Enabled = True
    cmdpStop.Enabled = True
    cmdpPause.Enabled = True
    cmdLoop.Enabled = True
    cmdpPlay.Enabled = False
End Sub

Private Sub cmdPrev_Click()
    Dim p As Integer
    On Error Resume Next
    If lstTracks.ListItems.Count = 0 Then
            MsgBox "Get Query CD first!", vbExclamation, App.Title
        Exit Sub
    End If
    If OpenDevice = False Then cmdOpenClose = True
    If lstTracks.SelectedItem.Index = 1 Then
        p = lstTracks.ListItems.Count
    Else
        p = lstTracks.SelectedItem.Index - 1
    End If
    lstTracks.ListItems(p).Selected = True
    cmdStop = True
    cmdPlay = True
    lstTracks.SetFocus
End Sub

Private Sub cmdpStop_Click()
    TBS.Enabled = True
    TBS2.Enabled = True
    On Local Error Resume Next
    MP3.mmStop
    tTimer.Enabled = False
    cmdpPlay.Enabled = True
    cmdpStop.Enabled = False
    cmdpPause.Enabled = False
    cmdLoop.Enabled = False
    lblDuration = "Total Time: 00:00"
    lblPosition = "Time: 00:00"
    If lstmp3wav.ListCount > 1 Then
        cmdnNext.Enabled = True
        cmdnpPrev.Enabled = True
    ElseIf lstmp3wav.ListCount = 1 Then
        cmdnNext.Enabled = False
        cmdnpPrev.Enabled = False
    End If
End Sub

Private Sub cmdQuery_Click()
    Dim i As Integer
    cmdQuery.Enabled = False
    asWorking = True
    cmdEncode.Enabled = False
    Call GetQueryCD
    If lstTracks.ListItems.Count > 0 Then
        lstTracks.ListItems(1).Selected = True
        lstTracks.SetFocus
        txtTitle.Text = lstTracks.ListItems(1).Text
        utcCover.Artist = txtArtist.Text
        utcCover.Album = txtAlbum.Text
        utcCover.Title = lstTracks.ListItems(1).Text
        modeTrack = False
        If CheckAlbum.Value = 1 Then
            utcCover.SearchCover
        ' .... Display Info Tag
        If utcCover.CoverFound = True Then
            cmdSaveCover.Enabled = True
            txtTitle.Text = utcCover.Title
            txtArtist.Text = utcCover.Artist
            txtAlbum.Text = utcCover.Album
            txtBand.Text = utcCover.Band
            'txtYear.Text = utcCover.Year
            txtGenre.Text = utcCover.Genre
            txtTrack.Text = utcCover.track
            txtComment.Text = utcCover.Comment
        Else
            cmdSaveCover.Enabled = False
        End If
        End If
        cmdEncode.Enabled = True
    ElseIf lstTracks.ListItems.Count < 1 Then
        If MsgBox("The response of the Server (Freedb.org) returned null! Meybe because the Server is busy or not reachable at this time! " & vbCr _
            & "You want to upload the tracks however?", vbYesNo + vbInformation + _
            vbDefaultButton1, "Freedb.org Query") = vbYes Then
            If CDW.OpenCD(Combo_Lettera.List(Combo_Lettera.ListIndex)) = True Then
                If CDW.GetNumberOfTracks > 0 Then
                    lsTracks.Clear
                    For i = 1 To CDW.GetNumberOfTracks
                        Set itmX = lstTracks.ListItems.Add(, , "Track " + str(i))
                        itmX.SubItems(1) = CDW.GetTrackLength(CLng(i))
                        lsTracks.AddItem "Track " + str(i) & ": Time " & CDW.GetTrackLength(CLng(i))
                    Next i
                    modeTrack = True
                    cmdEncode.Enabled = True
                Else
                    MsgBox "Could not get the track lengths. Is there a valid CD-W/DVDW in the drive?", vbExclamation, App.Title
                End If
                Else
                    MsgBox "Could not open the CD-W/DVDW Player. Are you trying to open an invalid drive type?", vbExclamation, App.Title
            End If
        End If
    End If
    ' .... Close CD if it is Opened
    If OpenDevice Then InitCDAUDIO False
    cmdQuery.Enabled = True
    asWorking = False
    'If lsTracks.ListCount > 0 Then
    '    lsTracks.Visible = True
    '    TBS2.Tabs(2).Selected = True
    'End If
End Sub

Private Sub cmdReload_Click()
    ' .... Get the list of DRVs
    Call GetAdapters
End Sub

Private Sub cmdSaveCover_Click()
If utcCover.CoverFound = False Then
            MsgBox "Nothing to Save!", vbExclamation, App.Title
        Exit Sub
    End If
    utcCover.SaveCoverAs txtAlbum.Text, , Qlt_80
End Sub

Private Sub cmdSelectPath_Click()
    Dim strFolder As String
    strFolder = BrowseFolder("Extract tracks to:", App.Path)
    If strFolder <> "" And strFolder <> "Error!" Then
        txtDestPath.Text = strFolder + "\"
    ' .... Save default path
    INI.DeleteKey "SETTING", "DEFAULT_PATH"
    INI.CreateKeyValue "SETTING", "DEFAULT_PATH", txtDestPath.Text
    End If
End Sub

Private Sub cmdStop_Click()
    On Local Error Resume Next
    If lstTracks.ListItems.Count = 0 Then
            MsgBox "Get Query CD first!", vbExclamation, App.Title
        Exit Sub
    End If
    If OpenDevice = False Then cmdOpenClose = True
    If CDW.StopCD Then
        statuslabel.Caption = "stopped"
        lblTAG.Caption = "TAGS-ARTIST"
        TCD.Enabled = False
        utcWait.Start = False
    Else
        statuslabel.Caption = "error!"
        lblTAG.Caption = "TAGS-ARTIST"
        utcWait.Start = False
    End If
End Sub

Private Sub Combo_Dispositivo_Click()
    On Error Resume Next
    If LCase$(Combo_Dispositivo.List(Combo_Dispositivo.ListIndex)) <> "image recorder" Then
        Combo_Lettera.ListIndex = Combo_Dispositivo.ListIndex
        Label_Info.Caption = Combo_Lettera.List(Combo_Dispositivo.ListIndex) & "\" & _
        Combo_Dispositivo.List(Combo_Dispositivo.ListIndex)
    End If
End Sub


Private Sub Combo_Lettera_Click()
    On Error Resume Next
    Combo_Dispositivo.ListIndex = Combo_Lettera.ListIndex
    Label_Info.Caption = Combo_Lettera.List(Combo_Lettera.ListIndex) & "\" & Combo_Dispositivo.List(Combo_Lettera.ListIndex)
    lblCDID.Caption = InitCDInfo()
    If lblCDID.Caption = "Not Ready" Then
        cmdQuery.Enabled = False
    Else
        cmdQuery.Enabled = True
    End If
End Sub

Private Sub Drive_OnAborted(Abort As Boolean)
    Abort = False
End Sub

Private Sub Drive_OnAddLogLine(TextType As NEROLib.NERO_TEXT_TYPE, Text As String)
    AddMessage Text
End Sub


Private Sub Drive_OnDoneBAOWriteToFile(ByVal StatusCode As NEROLib.NERO_BURN_ERROR, ByVal lNumberOfBytesWritten As Long)
    AddMessage "WriteToFile " & StatusCode
    AddMessage "Number of bytes Written " & lNumberOfBytesWritten
End Sub

Private Sub Drive_OnDoneBurn(StatusCode As NEROLib.NERO_BURN_ERROR)
    AddMessage Nero.ErrorLog
    AddMessage Nero.LastError
    If StatusCode <> NEROLib.NERO_BURN_OK Then
        AddMessage "Burn NOT finished successfully! (" & StatusCode & ")"
        ' .... Play Boooo!
        Call PlaySoundResource(102)
    Else
        AddMessage "Burn finished successfully! (" & StatusCode & ")"
        ' .... Play Ok
        Call PlaySoundResource(101)
    End If
    'btnAbort.Enabled = False
    'Browse.Enabled = True
    'Burn.Enabled = True
    PB.Value = 0
End Sub


Private Sub Drive_OnDoneCDInfo(ByVal pCDInfo As NEROLib.INeroCDInfo)
    'set number of existing sessions
    On Local Error GoTo NoTracks:
    NumExistingTracks = pCDInfo.Tracks.Count
    IsDriveWriteable = pCDInfo.IsWriteable
    DriveMediaType = pCDInfo.MediaType

    'set done flag
    DriveFinished = True
    Exit Sub
NoTracks:
    NumExistingTracks = 0
    DriveFinished = True
End Sub

Private Sub Drive_OnDoneErase(Ok As Boolean)
    On Local Error GoTo backError
    If Ok Then
        AddMessage "Disc Erase Successful!"
    Else
        AddMessage "Disc Erase Failed!"
    End If
    DriveFinished = True
Exit Sub
backError:
        'set done flag
        DriveFinished = True
    Err.Clear
End Sub

Private Sub Drive_OnDoneImport2(ByVal bOk As Boolean, ByVal pFolder As NEROLib.INeroFolder, ByVal pCDStamp As NEROLib.INeroCDStamp, ByVal pImportInfo As NEROLib.INeroImportDataTrackInfo, ByVal importResult As NEROLib.NERO_IMPORT_DATA_TRACK_RESULT)
    Dim i As Integer
        If bOk Then
            Set rootfolder = pFolder
        Else
            MsgBox "Error Reading In Data!", vbCritical, App.Title
            AddMessage "Error Reading In Data!"
        End If
    ' set done flag
    DriveFinished = True
End Sub


Private Sub Drive_OnDoneWaitForMedia(Success As Boolean)
    AddMessage "Done waiting for media... (" & Success & ")"
End Sub


Private Sub Drive_OnDriveStatusChanged(ByVal driveStatus As NEROLib.NERO_DRIVESTATUS_RESULT)
    AddMessage "Drive Changed... (" & driveStatus & ")"
End Sub


Private Sub Drive_OnMajorPhase(phase As NEROLib.NERO_MAJOR_PHASE)
    SplitText phase
End Sub


Private Sub Drive_OnProgress(ProgressInPercent As Long, Abort As Boolean)
    Abort = False
    PB.Value = ProgressInPercent
End Sub

Private Sub Drive_OnRoboPrintLabel(pbSuccess As Boolean)
    If pbSuccess Then
        AddMessage "Print Label success..."
    Else
        AddMessage "Error to Print Label!"
    End If
    'set done flag
    'DriveFinished = True
End Sub

Private Sub Drive_OnSetPhase(Text As String)
    SplitText Text
End Sub


Private Sub Drive_OnWriteDAE(ignore As Long, Data As Variant)
    AddMessage "Write DAE (" & Data & ")"
End Sub


Private Sub Form_Initialize()
    Call GetAdapters
End Sub

Private Sub Form_Load()
    Dim strFolder As String
    On Local Error GoTo ErrorHandler
    ' .... Exception Handler' = (Call the stack)
    SetUnhandledExceptionFilter AddressOf MyExceptionHandler
    
    ' .... Display the Copyright
    lblInfo.Caption = "© 2008/" & Format(Now, "yyyy") & " by Salvo Cortesiano. All Right Reserved!"
    
    ' .... Associate File to this Application
    ' .... Until show the Error :)
    ' .... ico=0 Program ico
    ' .... ico=1 Note
    ' .... ico=2 Log
    ' .... ico=3 Audio
    ' .... ico=4 CD List
    ' .... ico=5 ToDo
    
    'If RemuveExtension("dcl") = False Then
    '    MsgBox "Error to remuve the association extension {dcl}!", vbExclamation, App.Title
    'End If
    
    If Len(GetString(HKEY_CLASSES_ROOT, ".scl", "")) < 1 Then
        If AssociateExtension("scl", "Log file of " & App.EXEName, "text/plain", 2, " /L", True) = False Then:
    End If
    
    If Len(GetString(HKEY_CLASSES_ROOT, ".dcd", "")) < 1 Then
        If AssociateExtension("dcd", "CD List file of " & App.EXEName, "text/plain", 4, " /F", True) = False Then:
    End If
    
    ' .... Fix the Image Top
    Image2.Top = -195
    
    ' .... Subclassed Application for multyple Instance
    DoEvents
    If Not Hooked Then Hook Me
    
    ' .... Parse Command$
    If StrComp((Right$(Command$, 3)), "scl", vbTextCompare) = 0 _
            Or StrComp((Right$(Command$, 3)), "dcd", vbTextCompare) = 0 Then
        If ParseCommand(Command$) = False Then: MsgBox "Error to Parse Command$ " & vbCrLf & Command$, vbExclamation, "CD Tracker Main Form"
    End If
    
    ' .... FLAG_TAB => Select one
    tbsKey = "CDExtract"
    
    ' .... UTL CDDB FreeDB.org
    strURL = picFrame(0).Tag
    
    ' .... HELLO for FreeDB Server
    strHello = "&hello=me+at.home.com+xmcd+v1.0PL0&proto=5"
    
    ' .... Init ListView lstTracks
    If InitListView(lstTracks, True, True, True, True, False, True) Then:
    
    ' .... Display the Default Cover
    GetDefCover
    
    ' .... Reset INI File Path
    INI.ResetINIFilePath
    
    ' .... Read the File INI
        If INI.GetKeyValue("SETTING", "S_TOP") <> "" Then ssTop = INI.GetKeyValue("SETTING", "S_TOP")
        If INI.GetKeyValue("SETTING", "S_LEFT") <> "" Then ssLeft = INI.GetKeyValue("SETTING", "S_LEFT")
        
        ' .... Position MainForm
        If Len(ssLeft) = 0 Then
        
        ' .... Center Form
            ssTop = (Screen.Height - frmMain.Height) \ 2
            ssLeft = (Screen.Width - frmMain.Width) \ 2
            frmMain.Move ssLeft, ssTop
        Else
            frmMain.Move ssLeft, ssTop
        End If
        
        ' .... Display Time Track Mode
        If INI.GetKeyValue("SETTING", "DISPLAY_TIME") = "" Then
            OptionTime(1).Value = True
            II_ndex = 1
        Else
            OptionTime(INI.GetKeyValue("SETTING", "DISPLAY_TIME")).Value = True
            II_ndex = INI.GetKeyValue("SETTING", "DISPLAY_TIME")
        End If
        
        ' .... Encode To
        If INI.GetKeyValue("SETTING", "ENCODE_TO") = "" Then
            OptEncode(0).Value = True
            CheckWriteTag.Enabled = True
        Else
            OptEncode(INI.GetKeyValue("SETTING", "ENCODE_TO")).Value = True
            If INI.GetKeyValue("SETTING", "ENCODE_TO") = 0 Then CheckWriteTag.Enabled = True Else CheckWriteTag.Enabled = False
        End If
        
        ' .... Default Path
        If INI.GetKeyValue("SETTING", "DEFAULT_PATH") = "" Then
            strFolder = BrowseFolder("Select the Default folder of extracted Tracks:", App.Path)
                If strFolder <> "" And strFolder <> "Error!" Then
                    txtDestPath.Text = strFolder + "\"
                    ' .... Save default path
                    INI.DeleteKey "SETTING", "DEFAULT_PATH"
                    INI.CreateKeyValue "SETTING", "DEFAULT_PATH", txtDestPath.Text
                    MsgBox "The default path of extracted Tracks is:" & vbCr & txtDestPath.Text, vbInformation, App.Title
                Else
                    txtDestPath.Text = App.Path + "\"
                    INI.DeleteKey "SETTING", "DEFAULT_PATH"
                    INI.CreateKeyValue "SETTING", "DEFAULT_PATH", txtDestPath.Text
                    MsgBox "The default path of extracted Tracks is:" & vbCr & txtDestPath.Text, vbInformation, App.Title
                End If
        Else
            txtDestPath.Text = INI.GetKeyValue("SETTING", "DEFAULT_PATH")
        End If
        
        
        ' .... Track mode
        If INI.GetKeyValue("SETTING", "SAVE_TRACK_MODE") = "" Then cmbMode.ListIndex = 0 Else _
        cmbMode.ListIndex = INI.GetKeyValue("SETTING", "SAVE_TRACK_MODE")
        
        ' .... Manually Tags
        If INI.GetKeyValue("SETTING", "SAVE_TRACK_MODE_MANUALLY_TAGs") = "" Then _
        lblTagManually.Caption = "Title|Track|Album|Artist|Year" Else _
        lblTagManually.Caption = INI.GetKeyValue("SETTING", "SAVE_TRACK_MODE_MANUALLY_TAGs")
        
        ' .... Download Cover?
        If INI.GetKeyValue("SETTING", "DOWNLOAD_COVER") = "" Then CheckAlbum.Value = 0 Else _
        CheckAlbum.Value = INI.GetKeyValue("SETTING", "DOWNLOAD_COVER")
        
        ' .... Write Tags
        If INI.GetKeyValue("SETTING", "WRITE_TAGS") = "" Then CheckWriteTag.Value = 0 Else _
        CheckWriteTag.Value = INI.GetKeyValue("SETTING", "WRITE_TAGS")
        
        ' .... BitRate
        If INI.GetKeyValue("SETTING", "BIT_RATE") = "" Then cmbBitRate.ListIndex = 10 Else _
        cmbBitRate.ListIndex = INI.GetKeyValue("SETTING", "BIT_RATE")
        
        ' .... Private?
        If INI.GetKeyValue("SETTING", "PRIVATE") = "" Then CheckPrivate.Value = 0 Else _
        CheckPrivate.Value = INI.GetKeyValue("SETTING", "PRIVATE")
    
        ' .... Original?
        If INI.GetKeyValue("SETTING", "ORIGINAL") = "" Then CheckOriginal.Value = 0 Else _
        CheckOriginal.Value = INI.GetKeyValue("SETTING", "ORIGINAL")
    
        ' .... VBR?
        If INI.GetKeyValue("SETTING", "VBR") = "" Then CheckVBR.Value = 0 Else _
        CheckVBR.Value = INI.GetKeyValue("SETTING", "VBR")
    
        ' .... Copyright?
        If INI.GetKeyValue("SETTING", "COPYRIGHT") = "" Then CheckCopyright.Value = 0 Else _
        CheckCopyright.Value = INI.GetKeyValue("SETTING", "COPYRIGHT")
        
        ' .... Extra Tags?
        If INI.GetKeyValue("SETTING", "EXTRA TAGS") = "" Then CheckIncludeExtraTags.Value = 0 Else _
        CheckIncludeExtraTags.Value = INI.GetKeyValue("SETTING", "EXTRA TAGS")
        
        ' .... Include Cover?
        If INI.GetKeyValue("SETTING", "INCLUDE COVER") = "" Then CheckIncludeCover.Value = 0 Else _
        CheckIncludeCover.Value = INI.GetKeyValue("SETTING", "INCLUDE COVER")
        
        ' .... Encoded by:
        If INI.GetKeyValue("SETTING", "EXTRA TAGS ENCODED BY") <> "" Then txtEncodedBy.Text = _
        INI.GetKeyValue("SETTING", "EXTRA TAGS ENCODED BY")
        
        ' .... CopyRight:
        If INI.GetKeyValue("SETTING", "EXTRA TAGS COPYRIGHT INFO") <> "" Then txtCopyrightInfo.Text = _
        INI.GetKeyValue("SETTING", "EXTRA TAGS COPYRIGHT INFO")
        
        ' .... Language:
        If INI.GetKeyValue("SETTING", "EXTRA TAGS LANGUAGE") <> "" Then txtLanguage.Text = _
        INI.GetKeyValue("SETTING", "EXTRA TAGS LANGUAGE")
        
        ' .... Replace exist file?
        If INI.GetKeyValue("SETTING", "REPLACE EXIST FILE") = "" Then CheckReplace.Value = 0 Else _
        CheckReplace.Value = INI.GetKeyValue("SETTING", "REPLACE EXIST FILE")
        
        ' .... Default Media path
        If INI.GetKeyValue("SETTING", "DEFAULT_MEDIA_PATH") <> "" Then txtMediaPath.Text = _
        INI.GetKeyValue("SETTING", "DEFAULT_MEDIA_PATH") Else txtMediaPath.Text = App.Path + "\"
        
         ' .... Play all files of List
        If INI.GetKeyValue("SETTING", "PLAY ALL FILE OF LIST") = "" Then CheckPlayAll.Value = 0 Else _
        CheckPlayAll.Value = INI.GetKeyValue("SETTING", "PLAY ALL FILE OF LIST")
        
Exit Sub
ErrorHandler:
        Call WriteErrorLogs(Err.Number, Err.Description, "FormLoad {Sub: Load}", True, True)
    Err.Clear
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim OpenForm As Form
    If OpenError Then Exit Sub
    If asWorking = True Then
            readyToClose = False
    Else
    If MsgBox("Are you sure to Close this Application?", vbYesNo + vbInformation + _
        vbDefaultButton1, "Close Application") = vbYes Then
        readyToClose = True
        readyToCloseII = True
        ' .... UnHooked Application
        If Hooked Then Unhook
        ' .... Release the Deugger
        SetUnhandledExceptionFilter ByVal 0&
        ' .... Release the Library
        Call FreeLibrary(m_hMod)
        ' .... Release the Class GUIDE
        Set objGUIDE = Nothing
        ' .... UnHook the SO
        If Not InIDE() Then SetErrorMode SEM_NOGPFAULTERRORBOX
        ' .... Close all Form's
        For Each OpenForm In Forms
            Unload OpenForm
        Next OpenForm
        ' .... Destroy SyTray
        If GetSysTray(False) Then
            Set m_frmSysTray = Nothing
        End If
        ' .... Close CD if it is Opened
        If OpenDevice Then InitCDAUDIO False
        ' .... Stop All SND
        Call EndPlaySound
        ' .... Save Setting to File *.INI
        SaveSettingINI
        ' .... Unload Class INI
        Set INI = Nothing
        ' .... Stop MP3Player
        cmdpStop = True
    Else
        readyToClose = False
    End If
    End If
    Cancel = Not readyToClose
End Sub
Private Sub Form_Resize()
If asWorking = True Then
        Me.WindowState = vbNormal
    Exit Sub
Else
If Me.WindowState = vbMinimized Then
        Me.Visible = False
        If GetSysTray(True) Then
            SetIcon ICON_PROGRAM
            ShowTip "CD Tracker v1.0.3b now is Hidden in the Tray-Bar." & vbCrLf & "Double click or Right Click for Menu!", "CD Tracker v1.0.3b"
        End If
    ElseIf Me.WindowState = vbNormal Then
        Me.Visible = True
        If GetSysTray(False) Then:
    End If
End If
End Sub

Private Sub Form_Terminate()
    '.... Release the Form
    Set frmMain = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Private Function GetSysTray(ByVal sShowOrsHide As Boolean) As Boolean
    On Local Error GoTo ErrorHandler
    If sShowOrsHide = True Then
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "&Restore CD Tracker v1.0.3b", "Open", True
            .AddMenuItem "-"
            .AddMenuItem "&Netshadows on the Web", "netshadows"
            .AddMenuItem "&Mail to Salvo Cortesiano", "salvocortesiano"
            .AddMenuItem "-"
            .AddMenuItem "&Download the (Full) Project", "Download"
            .AddMenuItem "-"
            .AddMenuItem "&About...", "About"
            .AddMenuItem "-"
            .AddMenuItem "&Close CD Tracker v1.0.3b", "Close"
            .ToolTip = "CD Tracker v1.0.3b"
        End With
    ElseIf sShowOrsHide = False Then
        Unload m_frmSysTray
        Set m_frmSysTray = Nothing
    End If
    GetSysTray = True
Exit Function
ErrorHandler:
        GetSysTray = False
    Err.Clear
End Function

Private Sub SetIcon(sTypeIcon As sSetIcon)
    On Local Error Resume Next
    Select Case sTypeIcon
    Case 0
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(0).Picture.Handle
    Case 1
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(1).Picture.Handle
    Case 2
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(2).Picture.Handle
    Case 3
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(3).Picture.Handle
    Case 4
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(4).Picture.Handle
    Case 5
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(4).Picture.Handle
    Case 6
        m_frmSysTray.IconHandle = m_frmSysTray.Icon
    End Select
End Sub

Private Sub ShowTip(strMessage As String, strTitle As String)
    m_frmSysTray.ShowBalloonTip strMessage, strTitle, NIIF_INFO
End Sub

Private Sub lblPosition_Click()
    lblPosition = "Time: " & MP3.Position
End Sub

Private Sub lstmp3wav_Click()
    On Local Error Resume Next
    strMediaFile = lstmp3wavPath.List(lstmp3wav.ListIndex) + lstmp3wav.List(lstmp3wav.ListIndex)
    If Dir$(strMediaFile) <> "" And LCase(Right$(strMediaFile, 4)) = ".mp3" Or LCase(Right$(strMediaFile, 4)) = ".wma" _
    Or LCase(Right$(strMediaFile, 4)) = ".wav" Or LCase(Right$(strMediaFile, 4)) = ".mid" _
    Or LCase(Right$(strMediaFile, 4)) = ".snd" Or LCase(Right$(strMediaFile, 4)) = ".au" _
    Or LCase(Right$(strMediaFile, 4)) = ".aif" Or LCase(Right$(strMediaFile, 4)) = ".rmi" _
    Or LCase(Right$(strMediaFile, 4)) = ".midi" Or LCase(Right$(strMediaFile, 4)) = ".wmv" _
    Or LCase(Right$(strMediaFile, 4)) = ".mp2" Or LCase(Right$(strMediaFile, 4)) = ".mpeg" _
    Or LCase(Right$(strMediaFile, 4)) = ".mpg" Or LCase(Right$(strMediaFile, 4)) = ".mpa" _
    Or LCase(Right$(strMediaFile, 4)) = ".mpe" Or LCase(Right$(strMediaFile, 4)) = ".asf" _
    Or LCase(Right$(strMediaFile, 4)) = ".mp4" Then
    If lstmp3wav.ListCount > 1 Then
        cmdnNext.Enabled = True
        cmdnpPrev.Enabled = True
        cmdpPlay.Enabled = True
    ElseIf lstmp3wav.ListCount = 1 Then
        cmdnNext.Enabled = False
        cmdnpPrev.Enabled = False
        cmdpPlay.Enabled = True
        cmdLoop.Enabled = True
    End If
    Else
        cmdnNext.Enabled = False
        cmdnpPrev.Enabled = False
        cmdpPlay.Enabled = False
        cmdpStop.Enabled = False
        cmdpPause.Enabled = False
        cmdLoop.Enabled = False
    End If
End Sub

Private Sub lstTracks_ItemClick(ByVal Item As ComctlLib.ListItem)
    If lstTracks.ListItems.Count = 0 Then Exit Sub
    On Local Error Resume Next
    txtTitle.Text = lstTracks.SelectedItem.Text
    txtTrack.Text = Mid$(lstTracks.SelectedItem.SubItems(1), 7, Len(lstTracks.SelectedItem.SubItems(1)))
    If OpenDevice Then
        Lbrano.Caption = "track " & Format(lstTracks.SelectedItem.Index, "00")
        Label14.Caption = str(CDW.GetNumberOfTracks) & " "
        Label18.Caption = CDW.GetTrackLength(CLng(lstTracks.SelectedItem.Index))
        Label12.Caption = CLng(str(lstTracks.SelectedItem.Index)) & "- 00:00 "
    End If
End Sub


Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Dim OpenWeb As Integer
Select Case sKey
    Case "Open"
        Me.WindowState = vbNormal
        Me.Visible = True
        Me.Show
        Me.ZOrder
    Case "Close"
        Unload Me
    Case "About"
        MsgBox "CD Tracker v1.0.3b © 2009 by Salvo Cortesiano!", vbInformation, App.Title
    Case "salvocortesiano"
        If MsgBox("Are you sure to send e-mail to: {salvocortesiano@netshadows.it}?", vbYesNo + vbInformation + _
        vbDefaultButton1, "MailTo: salvocortesiano@netshadows.it") = vbYes Then _
        SendEmail "salvocortesiano@netshadows.it", App.Title, "To-Do " & App.Title, "", ""
    Case "netshadows"
        If MsgBox("Are you sure to visit the web: {www.netshadows.it}?", vbYesNo + vbInformation + _
        vbDefaultButton1, "Open page www.netshadows.it") = vbYes Then _
        OpenWeb = ShellExecute(Me.hwnd, "Open", "http://www.netshadows.it", "", App.Path, 1)
    Case "Download"
        OpenWeb = ShellExecute(Me.hwnd, "Open", "http://www.netshadows.it/CDTracker.rar", "", App.Path, 1)
    End Select
    
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    Me.WindowState = vbNormal
    Me.Visible = True
    Me.Show
    Me.ZOrder
End Sub


Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If (eButton = vbRightButton) Then m_frmSysTray.ShowMenu
End Sub

Private Sub Nero_OnFileSelImage(FileName As String)
    On Local Error GoTo Error_ISO
    With cDialog
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|Nero Image Files (*.nrg)|*.nrg"
        .DialogTitle = "Save ISO image as:"
        .InitDir = App.Path
        .flags = cdlOFNHideReadOnly
        .FilterIndex = 2
        .DefaultExt = ".nrg"
        .ShowOpen
        If .FileName = "" Then Exit Sub
    FileName = .FileName
    End With
Exit Sub
Error_ISO:
        Call WriteErrorLogs(Err.Number, Err.Description, Nero.LastError, True, True)
    Err.Clear
End Sub


Private Sub Nero_OnMegaFatal()
    AddMessage "A fatal error has occurred."
    BurnError = True
    Me.Caption = "CD Tracker v1.0.3b"
End Sub

Private Sub Nero_OnNonEmptyCDRW(Response As NEROLib.NERO_RESPONSE)
    AddMessage "The CD-RW/DVD is not empty!"
    Response = NERO_RETURN_EXIT
End Sub


Private Sub Nero_OnNoTrackFound()
    AddMessage "No Track found!"
End Sub


Private Sub Nero_OnOverburn(Response As NEROLib.NERO_RESPONSE)
    AddMessage "Warning: OverBurn (" & Response & ")"
End Sub


Private Sub Nero_OnOverburn2(ByVal pOverburnInfo As NEROLib.INeroOverburnInfo, Response As NEROLib.NERO_RESPONSE)
    AddMessage "Warning: OverBurn (" & Response & ")"
    AddMessage "Total Blocks on CD (" & pOverburnInfo.TotalBlocksOnCD & ")"
    AddMessage "Total Capacity (" & pOverburnInfo.TotalCapacity & ")"
End Sub


Private Sub Nero_OnRestart()
    AddMessage "The System is being restarted... Now!"
End Sub


Private Sub Nero_OnSettingsRestart(Response As NEROLib.NERO_RESPONSE)
    AddMessage Response
End Sub


Private Sub Nero_OnTempSpace(ByVal bstrCurrentDir As String, ByVal pi64FreeSpace As NEROLib.IInt64, ByVal pi64SpaceNeeded As NEROLib.IInt64, pbstrNewTempDir As String)
    On Local Error Resume Next
    bstrCurrentDir = BrowseFolder("Select temp Dir:", bstrCurrentDir)
        If bstrCurrentDir <> "" And bstrCurrentDir <> "Error!" Then
        bstrCurrentDir = bstrCurrentDir + "\"
    End If
    pbstrNewTempDir = BrowseFolder("Select New temp Dir:", pbstrNewTempDir)
        If pbstrNewTempDir <> "" And pbstrNewTempDir <> "Error!" Then
        pbstrNewTempDir = pbstrNewTempDir + "\"
    End If
End Sub


Private Sub Nero_OnWaitCD(WaitCD As NEROLib.NERO_WAITCD_TYPE, WaitCDLocalizedText As String)
    SplitText WaitCDLocalizedText
End Sub


Private Sub Nero_OnWaitCDDone()
    AddMessage "Done waiting for CD..."
End Sub


Private Sub Nero_OnWaitCDMediaInfo(LastDetectedMedia As NEROLib.NERO_MEDIA_TYPE, LastDetectedMediaName As String, RequestedMedia As NEROLib.NERO_MEDIA_TYPE, RequestedMediaName As String)
    AddMessage "Waiting for a particular media type (" & RequestedMediaName & ")"
End Sub

Private Sub Nero_OnWaitCDReminder()
    AddMessage "Still waiting for CD..."
End Sub


Private Sub OptEncode_Click(Index As Integer)
    If OptEncode(0).Value = True Then
        CheckWriteTag.Enabled = True
    ElseIf OptEncode(1).Value = True Then
        CheckWriteTag.Enabled = False
    End If
End Sub

Private Sub OptionTime_Click(Index As Integer)
    If OpenDevice Then
            MsgBox "Stop the Device first!", vbExclamation, App.Title
        Exit Sub
    End If
    II_ndex = Index
    Select Case Index
        Case 0
            Label12.Caption = "01:00:00:00 "
        Case 1
            Label12.Caption = "1- 00:00 "
        Case 2
            Label12.Caption = "00:00:00 "
        Case 3
            Label12.Caption = "00:00 "
    End Select
End Sub

Private Sub TBS_BeforeClick(Cancel As Integer)
    If asWorking = True Then
            MsgBox "Sorry; you Not select the TABs. Work in progress!", vbExclamation, App.Title
        Cancel = True
    End If
End Sub

Private Sub TBS_Click()
    On Local Error GoTo ErrorHandler
    If tbsKey = TBS.SelectedItem.key Then Exit Sub
    Select Case TBS.SelectedItem.key
        Case "CDExtract"
            txtTextLog.Text = ""
            lblPath.Caption = "n/a"
            picFrame(2).Visible = False
            picFrame(0).Visible = True
            picFrame(1).Visible = False
        Case "Burn"
            txtTextLog.Text = ""
            lblPath.Caption = "n/a"
            picFrame(2).Visible = False
            picFrame(1).Visible = True
            picFrame(0).Visible = False
            If lstmp3wav.ListCount > 0 Then
                lstmp3wav.Selected(0) = True
                lblInfoScan.Caption = "Files: " & lstmp3wav.ListCount & "!"
            End If
        Case "LogCDList"
            If Dir$(App.Path & "\_errs.scl") <> "" Then
                If OpenFile(App.Path & "\_errs.scl") Then lblPath.Caption = App.Path & "\_errs.scl" Else lblPath.Caption = "n/a"
            End If
            picFrame(2).Visible = True
            picFrame(1).Visible = False
            picFrame(0).Visible = False
    End Select
    tbsKey = TBS.SelectedItem.key
Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub



Private Function InitCDInfo() As String
    Dim CDInfo As New clsCDDB
    On Local Error GoTo ErrorHandler
    CDInfo.Init Combo_Lettera.List(Combo_Lettera.ListIndex)
    strCDiD = CDInfo.DiscID
    InitCDInfo = CDInfo.DiscID
    CDInfo.CloseCD
Exit Function
ErrorHandler:
    InitCDInfo = "Error"
        Call WriteErrorLogs(Err.Number, Err.Description, "FormMain {Function: InitCDInfo}", True, True)
    Err.Clear
End Function

Private Sub GetQueryCD()
    Dim CDInfo As New clsCDDB
    Dim temp As String
    On Local Error GoTo ErrorHandler
    CDInfo.Init Combo_Lettera.List(Combo_Lettera.ListIndex)
    strCDiD = CDInfo.DiscID
    strQuery = CDInfo.QueryString()
    If strQuery = "Not Ready" Then Exit Sub
        temp = FindCategory(strQuery)
        strCategorie = Element(temp, 1, "|")
        strCDiD = Element(temp, 2, "|")
    Call FindTracks(strCategorie, strCDiD)
Exit Sub
ErrorHandler:
        
    Err.Clear
End Sub

Private Function Element(ByVal strText As String, ByVal Numero As Integer, ByVal strSep As String) As String
    Dim Debut As Integer, R As Integer, No As Integer
    If Right(strText, Len(strSep)) <> strSep Then strText = strText & strSep
    Debut = 1
    No = 1
Element_0:
    R = InStr(Debut, strText, strSep)
    If R = 0 Then GoTo Element_End
    If Numero = No Then GoTo Element_10
    No = No + 1
    Debut = R + Len(strSep)
    If R >= Len(strText) Then GoTo Element_End
    DoEvents
    GoTo Element_0
    
Element_10:
    Element = Mid$(strText, Debut, R - Debut)
Element_End:
    
End Function

Private Function FindCategory(ByVal Query As String) As String
    Dim temp As String, strLigne As String, strdata As String, i As Integer, R As Integer
    If Query = "" Then Exit Function
    lstResponse.Clear
    temp = strURL & "?cmd=" & Query & strHello
    lstResponse.AddItem "Start Query:"
    lstResponse.AddItem Mid$(temp, 1, 50) & "..."
    lstResponse.Selected(lstResponse.ListCount - 1) = True
    temp = inetConnexion.OpenURL(temp, icString)
    
    txtTemp.Text = temp
    
    strLigne = Element(temp, 1, vbCrLf)
    strdata = Element(strLigne, 1, " ")
    Select Case strdata
        Case "202"
            lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "No data for this CD"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Case "403"
            lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "Query Error! Server Busy?"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Case "409"
            lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "Identifycation Fail!"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
    End Select
    If Left$(strdata, 1) = "5" Then
        lstResponse.AddItem strdata
        txtTemp.Text = strdata
        lstResponse.AddItem "Structure request incorrect!"
        lstResponse.Selected(lstResponse.ListCount - 1) = True
    End If
    If strdata = "200" Then
            FindCategory = Element(temp, 2, " ") & "|" & Element(Query, 3, "+")
        GoTo HendCategory
    End If
    If strdata <> "210" And strdata <> "211" Then
            FindCategory = "Error|Error"
        Exit Function
    End If
    Do While frmCategorie.lstCategories.ListCount > 0
            frmCategorie.lstCategories.RemoveItem 0
        DoEvents
    Loop
    i = 2
Boucle_Propositions:
    DoEvents
    strLigne = Element(temp, i, vbCrLf)
    i = i + 1
    If Left(strLigne, 1) = "." Or strLigne = "" Then GoTo BoucleHend
    frmCategorie.lstCategories.AddItem strLigne
    GoTo Boucle_Propositions
BoucleHend:
    frmCategorie.lstCategories.ListIndex = 0
    With frmMain
        .Tag = ""
        DoEvents
            frmCategorie.Show vbModal, Me
        Do While .Tag = ""
            DoEvents
        Loop
        If .Tag = "Abort" Then
            MsgBox "Command abort by User!", vbExclamation, App.Title
                GoTo HendCategory
            Exit Function
        End If
        FindCategory = Element(.Tag, 1, " ") & "|" & Element(.Tag, 2, " ")
    End With
    
HendCategory:
    On Local Error Resume Next
    Unload frmCategorie
End Function
Private Sub FindTracks(ByVal strCategory As String, ByVal CDiD As String)
    Dim temp As String, strLigne As String, strdata As String, i As Integer, R As Integer
    Dim iTracks As Integer, iTitle As Integer, Max_Len As Long
    If strCategory = "" Or CDiD = "" Then Exit Sub
    temp = strURL & "?cmd=cddb+read+" & strCategory & "+" & CDiD & strHello
    lstResponse.AddItem temp
    lstResponse.AddItem "Send request to Server..."
    lstResponse.Selected(lstResponse.ListCount - 1) = True
    temp = inetConnexion.OpenURL(temp, icString)
    lstResponse.AddItem temp
    lstResponse.Selected(lstResponse.ListCount - 1) = True
    txtTemp.Text = temp
    DoEvents
    ReDim TTracks(30) As sTracksTypes
    strLigne = Element(temp, 1, vbCr)
    strdata = Element(strLigne, 1, " ")
    Select Case strdata
        Case "401"
            lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "No info for this CD"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Case "402"
            lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "Error from Server"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Case "403"
            lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "Error: The Server is busy?"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Case "409"
            lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "No identify the CD ID!"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
    End Select
    If Left(strdata, 1) = "5" Then
        lstResponse.AddItem strdata
        txtTemp.Text = strdata
        lstResponse.AddItem "The structure of request is incorrect"
        lstResponse.Selected(lstResponse.ListCount - 1) = True
    End If
    If strdata <> "210" Then Exit Sub
    i = 1
    lstTracks.ListItems.Clear
    ClearFields
RetriveTitle:
    DoEvents
    strLigne = Element(temp, i, vbCrLf)
    i = i + 1
    
    If Left(strLigne, 1) = "." Or strLigne = "" Then
        lstResponse.AddItem temp
        txtTemp.Text = temp
        lstResponse.AddItem "Problem to analize the {Artist}!"
        lstResponse.Selected(lstResponse.ListCount - 1) = True
        Exit Sub
    End If
    
    If Left(strLigne, 6) <> "DTITLE" Then GoTo RetriveTitle
    strLigne = Element(strLigne, 2, "=")
    txtArtist.Text = Trim$(Element(strLigne, 1, "/"))
    txtAlbum.Text = Trim$(Element(strLigne, 2, "/"))
    
RetriveYear:
    DoEvents
    strLigne = Element(temp, i, vbCrLf)
    i = i + 1
    If Left$(strLigne, 1) = "." Or strLigne = "" Then
        lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "Problem to analize the {Year}!"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Exit Sub
    End If
    If Left(strLigne, 5) <> "DYEAR" Then GoTo RetriveYear
    txtYear.Text = Trim$(Element(strLigne, 2, "="))

Boucle_Genre:
    DoEvents
    strLigne = Element(temp, i, vbCrLf)
    i = i + 1
    If Left(strLigne, 1) = "." Or strLigne = "" Then
        lstResponse.AddItem temp
            txtTemp.Text = temp
            lstResponse.AddItem "Problem to analize the {Genre}!"
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Exit Sub
    End If
    If Left(strLigne, 6) <> "DGENRE" Then GoTo Boucle_Genre
    txtGenre.Text = Trim$(Element(strLigne, 2, "="))

    iTracks = 30
RetriveTracks:
    DoEvents
    strLigne = Element(temp, i, vbCrLf)
    i = i + 1
    Select Case Left$(strLigne, 4)
        Case "TTIT"
            strdata = Element(strLigne, 1, "=")
            iTitle = CInt(Mid(strdata, 7, 3))
            If iTitle > iTracks Then
                    ReDim Preserve TTracks(iTitle) As sTracksTypes
                iTracks = iTitle
            End If
            strdata = Trim$(Element(strLigne, 2, "="))
            If Left$(strdata, 1) Like "[0-9]" Then strdata = " " & strdata
            TTracks(iTitle).Title = strdata
            GoTo RetriveTracks
        Case "EXTT"
            strdata = Element(strLigne, 1, "=")
            iTitle = CInt(Mid(strdata, 5, 3))
            If iTitle > iTracks Then
                    ReDim Preserve TTracks(iTitle) As sTracksTypes
                iTracks = iTitle
            End If
            TTracks(iTitle).Autor = " (" & Trim$(Element(strLigne, 2, "=")) & ")"
            GoTo RetriveTracks
        Case "."
            lstResponse.AddItem "Populate List. Wait please..."
            lstResponse.Selected(lstResponse.ListCount - 1) = True
        Case Else
            GoTo RetriveTracks
    End Select
    lstTracks.View = lvwReport
    lsTracks.Clear
    Max_Len = 0
        PB.Max = iTitle + 1
        For i = 0 To iTitle
            DoEvents
            If TTracks(i).Autor <> " ()" Then
                Set itmX = lstTracks.ListItems.Add(, , CStr(TTracks(i).Title & TTracks(i).Autor))
                itmX.SubItems(1) = "Track " & Format(i + 1, "00")
                lsTracks.AddItem CStr(TTracks(i).Title & TTracks(i).Autor) & ": " & "Track " & Format(i + 1, "00")
                R = Len(TTracks(i).Title & TTracks(i).Autor)
            Else
                Set itmX = lstTracks.ListItems.Add(, , CStr(TTracks(i).Title))
                itmX.SubItems(1) = "Track " & Format(i + 1, "00")
                lsTracks.AddItem CStr(TTracks(i).Title) & ": " & "Track " & Format(i + 1, "00")
                R = Len(TTracks(i).Title)
            End If
            If R > Max_Len Then Max_Len = R
            DoEvents
            PB.Value = i
        Next i
    lstResponse.AddItem "End Query!"
    lstResponse.Selected(lstResponse.ListCount - 1) = True
    lstResponse.AddItem "Save Query info. Wait please..."
    lstResponse.Selected(lstResponse.ListCount - 1) = True
    Dim FL As Integer
    
    Dim strArtist As String, strAlbum, strPath As String, strFileName As String
    FL = FreeFile
    If txtArtist.Text <> "" Then strArtist = txtArtist.Text Else strArtist = Format(Now, "dd-mm-yyyy")
    If txtAlbum.Text <> "" Then strAlbum = txtAlbum.Text Else strAlbum = "Unknow"
    
    ' .... If default Path is nothing display Browser for Folder
    If txtDestPath.Text = "" Then
        strPath = BrowseFolder("Select the deault Folder of extracted tracks:", App.Path)
        If strPath <> "" And strPath <> "Error!" Then
            txtDestPath.Text = strPath + "\"
        Else
            txtDestPath.Text = App.Path + "\"
            strPath = App.Path + "\"
        End If
    Else
        strPath = txtDestPath.Text
    End If
    
    If Not FSO.FolderExists(strPath) Then
        ' .... Create folder CD Query
        If MakeDirectory(strPath + "CD Query") = False Then:
        ' .... Until display the Error now, because if the Folder exist return a Error ;)
    End If
    
    If MakeDirectory(strPath + "CD Query\" + txtArtist.Text) = False Then
    Else
        strPath = strPath + "CD Query\" + txtArtist.Text + "\"
    End If
    
    strFileName = strPath + strArtist + "-" + strAlbum + " (" + txtYear.Text + ").txt"
    Open strFileName For Output As FL
        Print #FL, "#  Generated by " & App.Title & " - " & Format(Now, "Long Date")
        Print #FL, "#  TAG: " & strArtist & "-" & strAlbum
        Print #FL, txtTemp.Text
    Close FL
    lstResponse.AddItem "End!"
    lstResponse.Selected(lstResponse.ListCount - 1) = True
    frmMain.Tag = Empty
    ' .... Retrive the List of Categories
    'txtTemp.Text = ListCategories()
    'RetriveMoreInfo strFileName
    On Local Error Resume Next
    PB.Value = 0
End Sub
Private Sub ClearFields()
    txtArtist = Empty: txtAlbum = Empty: txtYear = Empty: txtGenre = Empty: txtTrack = Empty
    txtTitle = Empty: txtBand = Empty: txtComment = Empty
End Sub

Private Sub SendEmail(Optional Adress As String, Optional Subjet As String, _
                      Optional Content As String, Optional CC As String, Optional CCC As String)
    Dim temp As String
    On Local Error Resume Next
    If Len(Subjet) Then temp = "&Subject=" & Subjet
    If Len(Content) Then temp = temp & "&Body=" & Content
    If Len(CC) Then temp = temp & "&CC=" & CC
    If Len(CCC) Then temp = temp & "&BCC=" & CCC
    If Mid(temp, 1, 1) = "&" Then Mid(temp, 1, 1) = "?"
    temp = "mailto:" & Adress & temp
    Call ShellExecute(Me.hwnd, "open", temp, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Private Function ListCategories() As String
    Dim temp As String
    txtTemp.Text = ""
    temp = strURL & "?cmd=cddb+lscat" & strHello
    temp = inetConnexion.OpenURL(temp, icString)
    ListCategories = temp
End Function

Private Sub RetriveMoreInfo(strFileName As String)
Dim LinesFromFile, NextLine As String
Dim FL As Integer
On Local Error Resume Next
'lstMoreInfo.Clear
    FL = FreeFile
    Open strFileName For Input As FL
        Do Until EOF(FL)
            Line Input #FL, NextLine
            ' .... Assume this Put line of text into LinesFromFile
            'LinesFromFile = LinesFromFile + NextLine + Chr(13) + Chr(10)
            'If UCase$(Mid$(NextLine, 1, 15)) = UCase$("# Disc length: ") Then lstMoreInfo.AddItem "Length: " & Mid$(NextLine, 16, Len(NextLine))
            'If UCase$(Mid$(NextLine, 1, 16)) = UCase$("# Processed by: ") Then lstMoreInfo.AddItem "Processed by: " & Mid$(NextLine, 17, Len(NextLine))
            'If UCase$(Mid$(NextLine, 1, 17)) = UCase$("# Submitted via: ") Then lstMoreInfo.AddItem "Submitted via: " & Mid$(NextLine, 18, Len(NextLine))
            'If UCase$(Mid$(NextLine, 1, 7)) = UCase$("DISCID=") Then lstMoreInfo.AddItem "CD ID: " & Mid$(NextLine, 8, Len(NextLine))
        Loop
    Close FL
End Sub

Private Function InitListView(sListView As ListView, Optional GRIDLINES As Boolean = True, Optional ONECLICKACTIVATE _
As Boolean = True, Optional FULLROWSELECT As Boolean = True, Optional TRACKSELECT As Boolean = True, _
Optional CHECKBOXES As Boolean = True, Optional SUBITEMIMAGES As Boolean = True) As Boolean
    Dim rStyle As Long
    Dim R As Long
    On Error GoTo ErrorHeadler
    ' .... Show Greed = True
    If GRIDLINES Then
        rStyle = SendMessageLong(sListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_GRIDLINES
        R = SendMessageLong(sListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... One click = True
    If ONECLICKACTIVATE Then
        rStyle = SendMessageLong(sListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_ONECLICKACTIVATE
        R = SendMessageLong(sListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... Select all Items = True
    If FULLROWSELECT Then
        rStyle = SendMessageLong(sListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_FULLROWSELECT
        R = SendMessageLong(sListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... Track Select = True
    If TRACKSELECT Then
        rStyle = SendMessageLong(sListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_TRACKSELECT
        R = SendMessageLong(sListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... CheckBox = True
    If CHECKBOXES Then
        rStyle = SendMessageLong(sListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_CHECKBOXES
        R = SendMessageLong(sListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... SubItem Image = True
    If SUBITEMIMAGES Then
        rStyle = SendMessageLong(sListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_SUBITEMIMAGES
        R = SendMessageLong(sListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    InitListView = True
Exit Function
ErrorHeadler:
    InitListView = False
        Call WriteErrorLogs(Err.Number, Err.Description, "Source: Function=InitListView", True, True)
    Err.Clear
End Function

Private Sub TBS2_BeforeClick(Cancel As Integer)
If asWorking = True Then
            MsgBox "Sorry; you Not select the TABs. Work in progress!", vbExclamation, App.Title
        Cancel = True
    End If
End Sub

Private Sub TBS2_Click()
 On Local Error GoTo ErrorHandler
    Select Case TBS2.SelectedItem.key
        Case "CDPlayer"
            PicFrames(0).Visible = True
            PicFrames(1).Visible = False
            PicFrames(2).Visible = False
            lsTracks.Visible = False
        Case "EncodeOption"
            PicFrames(1).Visible = True
            PicFrames(0).Visible = False
            PicFrames(2).Visible = False
            lsTracks.Visible = True
        Case "AdvancedTags"
            PicFrames(2).Visible = True
            PicFrames(1).Visible = False
            PicFrames(0).Visible = False
            lsTracks.Visible = True
    End Select
Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub TCD_Timer()
    On Local Error Resume Next
    If CDW.IsPlaying Then
    Select Case II_ndex
        Case 0
            Label12.Caption = CDW.GetCurrentPosition(Track_Min_Sec_Mil) & " "
        Case 1
            Label12.Caption = CDW.GetCurrentPosition(Track_Min_Sec) & " "
        Case 2
            Label12.Caption = CDW.GetCurrentPosition(Min_Sec_Mil) & " "
        Case 3
            Label12.Caption = CDW.GetCurrentPosition(Min_Sec) & " "
        End Select
        ' ---- OR
        'Label12.Caption = CDW.GetCurrentPosition(Track_Min_Sec_Mil) & " "
        lblTAG.Caption = Mid$(lstTracks.SelectedItem.Text, 1, 25) & "..."
    Else
        statuslabel.Caption = "standby"
    End If
End Sub

Private Sub InitCDAUDIO(strEnableDisable As Boolean)
    Dim i As Integer
    On Local Error Resume Next
    If strEnableDisable Then
        If CDW.OpenCD(Mid$(Combo_Lettera.List(Combo_Lettera.ListIndex), 1, 2)) Then
            Lbrano.Caption = "track " & Format(lstTracks.SelectedItem.Index, "00")
            Label14.Caption = str(CDW.GetNumberOfTracks) & " "
            Label18.Caption = CDW.GetTrackLength(CLng(lstTracks.SelectedItem.Index))
            Label12.Caption = CLng(str(lstTracks.SelectedItem.Index)) & "- 00:00 "
            If lstTracks.ListItems.Count > 0 Then lstTracks.ListItems(1).Selected = True
            TCD.Enabled = True
            OpenDevice = True
            statuslabel.Caption = "device open"
            For i = 0 To 3
                OptionTime(i).Enabled = False
            Next i
        Else
            TCD.Enabled = False
            OpenDevice = False
            statuslabel.Caption = "error"
            For i = 0 To 3
                OptionTime(i).Enabled = False
            Next i
        End If
    Else
        If CDW.StopCD Then
        statuslabel.Caption = "Stopped"
    Else
    End If
    If CDW.CloseCD Then
        statuslabel.Caption = "Standby"
        TCD.Enabled = False
        OpenDevice = False
        For i = 0 To 3
            OptionTime(i).Enabled = True
        Next i
    Else
        MsgBox "Could not close CD device. The CD may not have OPENED successfully!", vbOKOnly + vbInformation, "Error CD Closed!"
        OpenDevice = False
    End If
    End If
End Sub



Private Sub GetDefCover()
    Dim f As Integer
    Dim b() As Byte
    On Local Error Resume Next
    f = FreeFile
        b = LoadResData(101, "GIFT")
        Open App.Path + "\cover.gif" For Binary Access Write Shared As #f
            Put #f, , b
        Close #f
    If Dir$(App.Path + "\cover.gif") <> "" Then Set utcCover.Picture = LoadPicture(App.Path + "\cover.gif")
End Sub






Private Sub SaveSettingINI()
    Dim i As Integer
    On Local Error Resume Next
    
    ' .... display Time Track
    For i = 0 To 3
        If OptionTime(i).Value = True Then
                INI.DeleteKey "SETTING", "DISPLAY_TIME"
                INI.CreateKeyValue "SETTING", "DISPLAY_TIME", OptionTime(i).Index
            Exit For
        End If
    Next i
    
    ' .... Encode To
    i = 0
    For i = 0 To 1
        If OptEncode(i).Value = True Then
            INI.DeleteKey "SETTING", "ENCODE_TO"
            INI.CreateKeyValue "SETTING", "ENCODE_TO", OptEncode(i).Index
        End If
    Next i
    
    ' .... Default path
    INI.DeleteKey "SETTING", "DEFAULT_PATH"
    INI.CreateKeyValue "SETTING", "DEFAULT_PATH", txtDestPath.Text
    
    ' .... Save track mode
    INI.DeleteKey "SETTING", "SAVE_TRACK_MODE"
    INI.CreateKeyValue "SETTING", "SAVE_TRACK_MODE", cmbMode.ListIndex
    
    ' .... Save the Manually Tags
    If lblTagManually.Caption <> Empty Then
        INI.DeleteKey "SETTING", "SAVE_TRACK_MODE_MANUALLY_TAGs"
        INI.CreateKeyValue "SETTING", "SAVE_TRACK_MODE_MANUALLY_TAGs", lblTagManually.Caption
    End If
    
    ' .... Download Cover?
    INI.DeleteKey "SETTING", "DOWNLOAD_COVER"
    INI.CreateKeyValue "SETTING", "DOWNLOAD_COVER", CheckAlbum.Value
    
    ' \*.... Encodeing Options .... '/*
    
    ' .... Write Tags?
    INI.DeleteKey "SETTING", "WRITE_TAGS"
    INI.CreateKeyValue "SETTING", "WRITE_TAGS", CheckWriteTag.Value
    
    ' .... BitRate
    INI.DeleteKey "SETTING", "BIT_RATE"
    INI.CreateKeyValue "SETTING", "BIT_RATE", cmbBitRate.ListIndex
    
    ' .... Private?
    INI.DeleteKey "SETTING", "PRIVATE"
    INI.CreateKeyValue "SETTING", "PRIVATE", CheckPrivate.Value
    
    ' .... Original?
    INI.DeleteKey "SETTING", "ORIGINAL"
    INI.CreateKeyValue "SETTING", "ORIGINAL", CheckOriginal.Value
    
    ' .... VBR?
    INI.DeleteKey "SETTING", "VBR"
    INI.CreateKeyValue "SETTING", "VBR", CheckVBR.Value
    
    ' .... Copyright?
    INI.DeleteKey "SETTING", "COPYRIGHT"
    INI.CreateKeyValue "SETTING", "COPYRIGHT", CheckCopyright.Value
    
    ' .... Extra Tags?
    INI.DeleteKey "SETTING", "EXTRA TAGS"
    INI.CreateKeyValue "SETTING", "EXTRA TAGS", CheckIncludeExtraTags.Value
    
    ' .... Include Cover?
    INI.DeleteKey "SETTING", "INCLUDE COVER"
    INI.CreateKeyValue "SETTING", "INCLUDE COVER", CheckIncludeCover.Value
    
    ' .... Encoded by:
    INI.DeleteKey "SETTING", "EXTRA TAGS ENCODED BY"
    INI.CreateKeyValue "SETTING", "EXTRA TAGS ENCODED BY", txtEncodedBy.Text
    
    ' .... CopyRight:
    INI.DeleteKey "SETTING", "EXTRA TAGS COPYRIGHT INFO"
    INI.CreateKeyValue "SETTING", "EXTRA TAGS COPYRIGHT INFO", txtCopyrightInfo.Text
    
    ' .... Language:
    INI.DeleteKey "SETTING", "EXTRA TAGS LANGUAGE"
    INI.CreateKeyValue "SETTING", "EXTRA TAGS LANGUAGE", txtLanguage.Text
    
    ' .... Replace file?
    INI.DeleteKey "SETTING", "REPLACE EXIST FILE"
    INI.CreateKeyValue "SETTING", "REPLACE EXIST FILE", CheckReplace.Value
    
    ' .... Save default Media path
    INI.DeleteKey "SETTING", "DEFAULT_MEDIA_PATH"
    INI.CreateKeyValue "SETTING", "DEFAULT_MEDIA_PATH", txtMediaPath.Text
    
    ' .... Play all files of List
    INI.DeleteKey "SETTING", "PLAY ALL FILE OF LIST"
    INI.CreateKeyValue "SETTING", "PLAY ALL FILE OF LIST", CheckPlayAll.Value
    
    ' .... Position form
    If Me.WindowState <> vbMinimized Then
        INI.DeleteKey "SETTING", "S_LEFT"
        INI.CreateKeyValue "SETTING", "S_LEFT", frmMain.Left
        INI.DeleteKey "SETTING", "S_TOP"
        INI.CreateKeyValue "SETTING", "S_TOP", frmMain.Top
    End If
    INI.DeleteKey "SETTING", "LAST PATH SCAN"
End Sub

Private Sub AddMessage(ByVal Message As String)
    On Local Error Resume Next
    '/* USE TEXTBOX
    'lst_Messages.Text = lst_Messages.Text + Message + Chr$(13) + Chr$(10)
    'If CheckLog.value = 1 Then WriteLog lst_Messages.Text + Message + Chr$(13) + Chr$(10)
    'lst_Messages.SelStart = Len(lst_Messages.Text)
    '/* OR USE LISTBOX ;)
    lst_Messages.AddItem Message
        If lst_Messages.ListCount <> 0 Then
            lst_Messages.ListIndex = lst_Messages.ListCount - 1
        lst_Messages.Refresh
    End If
End Sub

Private Sub GetAdapters()
    Dim myIndex As Long
    Dim Major_High As Integer
    Dim Major_Low As Integer
    Dim Minor_High As Integer
    Dim Minor_Low As Integer
    Dim ValidVersion As Boolean
    Dim ns As NeroSpeeds
    Dim k As Long
    Dim strBuffer As String
    
    On Local Error GoTo Init_Error
    
    ' init Nero
    Set Nero = New Nero
    
    'Check valid version
    ValidVersion = True
    Nero.APIVersion Major_High, Major_Low, Minor_High, Minor_Low
    If Major_High < 6 Then
        ValidVersion = False
    ElseIf Major_High = 6 And Major_Low < 3 Then
        ValidVersion = False
    ElseIf Major_High = 6 And Major_Low = 3 And Minor_High < 1 Then
        ValidVersion = False
    ElseIf Major_High = 6 And Major_Low = 3 And Minor_High = 1 And Minor_Low < 6 Then
        ValidVersion = False
    End If
    
    ' valid version of Nero?
    If Not ValidVersion Then
            MsgBox "Nero Version 6.3.1.6 Or Greater Required!", vbExclamation, App.Title
    End If
    
    lst_Messages.Clear
    
    ' get Drive Nero version
    AddMessage "Init Nero:"
    AddMessage "Nero Version: " & "v." & Major_High & "." & Major_Low & "." & Minor_High & Minor_Low
    
    ' count available Drives
    Set Drives = Nero.GetDrives(NERO_MEDIA_CDR)
    Combo_Dispositivo.Clear
    Combo_Lettera.Clear
    For myIndex = 0 To Drives.Count - 1
        If Drives(myIndex).DevType = NERO_SCSI_DEVTYPE_WORM And _
            InStr(LCase$(Drives(myIndex).DeviceName), "image recorder") = 0 Then
            Combo_Dispositivo.AddItem Drives(myIndex).DeviceName, myIndex
            Combo_Lettera.AddItem UCase$(Drives(myIndex).DriveLetter) & ":"
        Else
            Combo_Dispositivo.AddItem Drives(myIndex).DeviceName, myIndex
        End If
    
    ' now retrive additional info
    Set Drive = Drives(myIndex)
    If Drives(myIndex).BufUnderrunProtName <> "" Then
            AddMessage "/*"
            AddMessage "Drive: " & myIndex & "-" & Drives(myIndex).DeviceName & " ** Device Ready (" & CStr(Drive.DeviceReady) & ")"
        
        'get read speed
        Set ns = Drive.AvailableSpeeds(NERO_ACCESSTYPE_READ, NERO_MEDIA_CDR + NERO_MEDIA_DVD_ANY)
            AddMessage "Base Read Speed: " & ns.BaseSpeedKBs & " Kb/s"
        For k = 0 To ns.Count - 1
            strBuffer = strBuffer & CStr(ns(k)) & "-"
        Next
            strBuffer = strBuffer & " Kb/s"
            AddMessage "Available Read Speeds: " & strBuffer
        
        'get write speed
        strBuffer = ""
            Set ns = Drive.AvailableSpeeds(NERO_ACCESSTYPE_WRITE, NERO_MEDIA_CDR + NERO_MEDIA_DVD_ANY)
            AddMessage "Base Write Speed: " & ns.BaseSpeedKBs & " Kb/s"
        For k = 0 To ns.Count - 1
            strBuffer = strBuffer & CStr(ns(k)) & "-"
        Next
            strBuffer = strBuffer & " Kb/s"
        ' stamp info
        AddMessage "Available Write Speeds: " & strBuffer
        AddMessage "Buffer Underrun Protection Name: " & Drive.BufUnderrunProtName
        AddMessage "Device Ready: " & CStr(Drive.DeviceReady)
        AddMessage "Drive buffer size: " & CStr(Drive.DriveBufferSize) & " Kb"
        AddMessage "*\"
    Else
        AddMessage "Drive: " & myIndex & "-" & Drives(myIndex).DeviceName & " ** Device Ready (" & CStr(Drive.DeviceReady) & ")"
    End If
    Next myIndex
    Set ns = Nothing
    
    ' use first Drive selected?
    If Combo_Dispositivo.ListCount > 0 Then
            Combo_Dispositivo.ListIndex = 0
        Set Drive = Drives(Combo_Dispositivo.ListIndex)
    End If
    If Combo_Lettera.ListCount > 0 Then Combo_Lettera.ListIndex = 0
    
    Set Drive = Nothing
    Set Drives = Nothing
    Set Nero = Nothing
Exit Sub
Init_Error:
        Call WriteErrorLogs(Err.Number, Err.Description, "FormMain {Form: Initialize}", True, True)
    Err.Clear
End Sub

Private Function NameFromPath(strPath As String) As String
    Dim lngPos As Long
    Dim strPart As String
    Dim blnIncludesFile As Boolean
    lngPos = InStrRev(strPath, "\")
    blnIncludesFile = InStrRev(strPath, ".") > lngPos
    strPart = ""
    If lngPos > 0 Then
        If blnIncludesFile Then
            strPart = Right$(strPath, Len(strPath) - lngPos)
        End If
    End If
    NameFromPath = strPart
End Function

Private Function SplitText(ByVal Data As String)
    Dim temp As String
    Dim i As Integer
    temp = ""
        For i = 1 To Len(Data)
            If Mid$(Data, i, 1) = Chr$(13) Then
                AddMessage Trim$(temp)
                temp = ""
            ElseIf Mid$(Data, i, 1) <> Chr$(10) Then
                temp = temp + Mid$(Data, i, 1)
            End If
        Next
    If temp <> "" Then AddMessage Trim$(temp)
End Function

Private Function BrowseFolder(ByVal strTitle As String, Optional strPath As String = "") As String
    Dim fOlder As String
    On Local Error GoTo ErrorHandler
    If strPath = "" Then
        strPath = App.Path + "\"
    Else
        If Right$(strPath, 1) <> "\" Then strPath = strPath + "\"
    End If
    fOlder = BrowseForFolder(Me.hwnd, strTitle, strPath)
    If fOlder <> "" Then BrowseFolder = fOlder Else BrowseFolder = ""
Exit Function
ErrorHandler:
        BrowseFolder = "Error!"
    Err.Clear
End Function

Private Function MSF_TO_LBA(ByVal Minutes As Long, ByVal Seconds As Long, ByVal Frames As Long) As Long
    MSF_TO_LBA = ((60 * 75 * (Minutes)) + (75 * (Seconds)) + ((Frames) - 150))
End Function

Private Function CDWDVDWToMp3(ByVal sDevice As String, ByVal sTrack As Long, ByVal TrackFileName As String) As Boolean
    Dim conf As BE_CONFIG_LHV1
    Dim info As RAW_READ_INFO
    Dim Toc As CDROM_TOC
    Dim nTrack As Long
    Dim hDev As Long
    Dim n As Long
    Dim i As Long
    Dim trackEnd As Long
    Dim trackStart As Long
    Dim trackSize As Long
    Dim trackPos As Long
    Dim smpl() As Byte
    Dim hbes As Long
    Dim dwSamples As Long
    Dim dwBuffer As Long
    Dim dwWrite As Long
    Dim Buffer() As Byte
    Dim sData() As Byte
    
    On Local Error GoTo Encoded_Error
    
    If Dir$(TrackFileName) <> "" And CheckReplace.Value = 0 Then
        If MsgBox("The file {" & GetFilePath(TrackFileName, Only_FileName_and_Extension) & ") already exists. You want to replace it?", vbYesNo + vbInformation + _
            vbDefaultButton2, App.Title) = vbNo Then
                CDWDVDWToMp3 = False
            Exit Function
        End If
    End If
    
    'Init Lame.dll
    With conf
        .dwReSampleRate = 44100
        .dwSampleRate = 44100
        .dwConfig = BE_CONFIG_LAME
        .dwStructSize = Len(conf)
        .dwMpegVersion = MPEG1
        .dwStructVersion = 1
        .dwMaxBitrate = 160
        .dwBitrate = StripLeft(cmbBitRate.List(cmbBitRate.ListIndex), ":", True)
        .bOriginal = CheckOriginal.Value = 1
        .bCopyright = CheckCopyright.Value = 1
        .bPrivate = CheckPrivate.Value = 1
        .bEnableVBR = CheckVBR.Value = 1
        .bCRC = 1
        .bNoRes = 1
    End With
    
    ' .... Init Drive CD-W/DVDW
    hDev = CreateFile("\\.\" & sDevice, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If (hDev = -1) Then
            CDWDVDWToMp3 = False
                Call WriteErrorLogs("0", "Error to CreateFile!", "Function {CDWDVDWToMp3}!", False, True)
        Exit Function
    End If
    
    ' .... Table of Content"
    If (DeviceIoControl(hDev, IOCTL_CDROM_READ_TOC, ByVal 0&, 0, Toc, Len(Toc), i, ByVal 0&) = 0) Then
                AddMessage "Impossible to access Read the CD-W/DVDW!"
                    CDWDVDWToMp3 = False
                        Call WriteErrorLogs("0", "Impossible to Read the CD-W/DVDW!", "Function {CDWDVDWToMp3}!", False, True)
                CloseHandle hDev
        Exit Function
    End If
    
    ' .... Calcule Tracks TOC
    nTrack = Toc.LastTrack - Toc.FirstTrack + 1
    If (sTrack > nTrack) Then
                AddMessage "The number of Tracks is Not correct!"
                    CDWDVDWToMp3 = False
                        Call WriteErrorLogs("0", "The number of Tracks is Not correct!", "Function {CDWDVDWToMp3!}", False, True)
                CloseHandle hDev
        Exit Function
    End If
    
    ' .... Retrive the info of the Tracks
    trackStart = MSF_TO_LBA(Toc.TrackData(sTrack).Address(1), Toc.TrackData(sTrack).Address(2), Toc.TrackData(sTrack).Address(3))
    trackEnd = MSF_TO_LBA(Toc.TrackData(sTrack + 1).Address(1), Toc.TrackData(sTrack + 1).Address(2), Toc.TrackData(sTrack + 1).Address(3))
    trackSize = (trackEnd - trackStart + 1) * RAW_SECTOR_SIZE
    
    ' .... Init Lame.dll
    Call StdbeInitStream(conf, dwSamples, dwBuffer, hbes)
    ReDim Buffer(dwBuffer - 1)
    
    ' .... Info CD-W/DVDW track length
    n = LARGEST_SECTORS_PER_READ
    ReDim smpl(RAW_SECTOR_SIZE * n - 1)
    PB.Max = trackEnd - trackStart
    PB.Value = 0
    
    ' .... Open the file to write the *.mp3
    trackPos = trackStart
    Open TrackFileName For Binary Access Write Lock Read As #1
    Do While trackPos + n < trackEnd
        
        ' .... Retrive info CD-W/DVDW
        info.DiskOffset.lowpart = trackPos * 2048&
        info.TrackMode = CDDA
        info.SectorCount = n
        
        ' .... Display the Work
        Debug.Print "CD-W/DVDW: " & DeviceIoControl(hDev, IOCTL_CDROM_RAW_READ, info, Len(info), _
        smpl(0), RAW_SECTOR_SIZE * n, i, ByVal 0&)

        ' .... Start Encoding :)
        Call StdbeEncodeChunk(hbes, (UBound(smpl) + 1) / 2, smpl(0), Buffer(0), dwWrite)
        If dwWrite Then
            ReDim sData(dwWrite - 1)
                MemoryCopy sData(0), Buffer(0), dwWrite
            Put #1, , sData
        End If
        
        ' .... Display the working
        trackPos = trackPos + n
        PB.Value = trackPos - trackStart
        DoEvents
        If ABORT_ENCODING = True Then Exit Do
        Loop
    
    ' .... Release the Drive
    CloseHandle hDev
    
    ' .... Close the structure and writing the *.mp3 file
    Call StdbeDeinitStream(hbes, Buffer(0), dwWrite)
    If dwWrite Then
        ReDim sData(dwWrite - 1)
            MemoryCopy sData(0), Buffer(0), dwWrite
        Put #1, , sData
    End If
    Close #1
    
    ' .... Release the Lame.dll
    Call StdbeCloseStream(hbes)
    
    PB.Value = 0
    CDWDVDWToMp3 = True
    
    If CheckWriteTag.Enabled = True Then
        If CheckWriteTag.Value = 1 Then
            If WriteTagOfTrack(lstMedia.List(lstMedia.ListCount - 1)) = False Then: ' Silent Error
        End If
    End If
    
    Exit Function
Encoded_Error:
    CDWDVDWToMp3 = False
    PB.Value = 0
        Call WriteErrorLogs(Err.Number, Err.Description, "Function {CDWDVDWToMp3}!" & vbCr _
        & "To encode Track {" & sTrack & "}.", False, True)
    Err.Clear
End Function

Private Function CDWDVDWToWav(ByVal sDevice As String, ByVal sTrack As Long, ByVal TrackFileName As String) As Boolean
    Dim chk As WAVCHUNKHEADER
    Dim fmt As WAVCHUNKFORMAT
    Dim info As RAW_READ_INFO
    Dim mgc As String * 4
    Dim Toc As CDROM_TOC
    Dim nTrack As Long
    Dim hDev As Long
    Dim trackEnd As Long
    Dim trackStart As Long
    Dim trackSize As Long
    Dim trackPos As Long
    Dim smpl() As Byte
    Dim mSize As Long
    Dim n As Long
    Dim i As Long
    
    On Local Error GoTo Encoded_Error
    
    If Dir$(TrackFileName) <> "" And CheckReplace.Value = 0 Then
        If MsgBox("The file {" & GetFilePath(TrackFileName, Only_FileName_and_Extension) & ") already exists. You want to replace it?", vbYesNo + vbInformation + _
            vbDefaultButton2, App.Title) = vbNo Then
                CDWDVDWToWav = False
            Exit Function
        End If
    End If
    
    ' .... Init Drive CD-W/DVDW
    hDev = CreateFile("\\.\" & sDevice, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If (hDev = -1) Then
        CDWDVDWToWav = False
            Call WriteErrorLogs("0", "Error to CreateFile!", "Function {CDWDVDWToWav}!", False, True)
        Exit Function
    End If
    
    ' .... Table of Content"
    If (DeviceIoControl(hDev, IOCTL_CDROM_READ_TOC, ByVal 0&, 0, Toc, Len(Toc), i, ByVal 0&) = 0) Then
        AddMessage "Impossible to Read the CD-W/DVDW!"
                CDWDVDWToWav = False
                    Call WriteErrorLogs("0", "Impossible to Read the CD-W/DVDW!", "Function {CDWDVDWToWav}!", False, True)
                CloseHandle hDev
        Exit Function
    End If
    
    ' .... Calcule Tracks TOC
    nTrack = Toc.LastTrack - Toc.FirstTrack + 1
    If (sTrack > nTrack) Then
        AddMessage "The number of Tracks is Not correct!"
                CDWDVDWToWav = False
                    Call WriteErrorLogs("0", "The number of Tracks is Not correct!", "Function {CDWDVDWToMp3}!", False, True)
                CloseHandle hDev
        Exit Function
    End If
    
    ' .... Retrive the info of the Tracks
    trackStart = MSF_TO_LBA(Toc.TrackData(sTrack).Address(1), Toc.TrackData(sTrack).Address(2), Toc.TrackData(sTrack).Address(3))
    trackEnd = MSF_TO_LBA(Toc.TrackData(sTrack + 1).Address(1), Toc.TrackData(sTrack + 1).Address(2), Toc.TrackData(sTrack + 1).Address(3))
    trackSize = (trackEnd - trackStart + 1) * RAW_SECTOR_SIZE

    ' .... Open the file to write the *.wav
    Open TrackFileName For Binary Access Write Lock Read As 1
    
    ' .... format the Tag of *.wav
    fmt.wFormatTag = 1
    fmt.wChannels = 2
    fmt.dwSamplesPerSec = 44100
    fmt.dwAvgBytesPerSec = 176400
    fmt.wBlockAlign = 4
    fmt.wBitsPerSample = 16

    ' .... Wav header
    chk.ChunkID = "RIFF"
    chk.ChunkSize = 0
    Put 1, , chk

    ' .... Wave signature
    mgc = "WAVE"
    Put 1, , mgc

    ' .... Format
    chk.ChunkID = "fmt "
    chk.ChunkSize = Len(fmt)
    Put 1, , chk
    Put 1, , fmt

    ' .... Data
    chk.ChunkID = "data"
    chk.ChunkSize = 0
    sndofs = Seek(1)
    Put 1, , chk

    ' .... Info CD-W/DVDW track length
    n = LARGEST_SECTORS_PER_READ
    ReDim smpl(RAW_SECTOR_SIZE * n - 1)
    PB.Max = trackEnd - trackStart
    PB.Value = 0

    trackPos = trackStart
    Do While trackPos + n < trackEnd
        
        ' .... Init CD-W/DVDW
        info.DiskOffset.lowpart = trackPos * 2048&
        info.SectorCount = n
        info.TrackMode = CDDA
        
        ' .... Display the Encoding
        Debug.Print DeviceIoControl(hDev, IOCTL_CDROM_RAW_READ, info, Len(info), smpl(0), _
        RAW_SECTOR_SIZE * n, i, ByVal 0&)
        mSize = mSize + UBound(smpl) + 1
        Put #1, , smpl
        
        ' .... Display the working
        trackPos = trackPos + n
        PB.Value = trackPos - trackStart
        DoEvents
        If ABORT_ENCODING = True Then Exit Do
        Loop

    ' .... Release the Lame.dll
    CloseHandle hDev

    ' .... Data header
    chk.ChunkID = "data"
    chk.ChunkSize = mSize
    Put 1, sndofs, chk

    ' .... File Header
    chk.ChunkID = "RIFF"
    chk.ChunkSize = LOF(1) - 4
    Put 1, 1, chk

    ' .... Close the File
    Close #1
    
    PB.Value = 0
    CDWDVDWToWav = True
Exit Function
Encoded_Error:
        PB.Value = 0
            Call WriteErrorLogs(Err.Number, Err.Description, "Function {CDWDVDWToWav}!" & vbCr _
        & "To encode Track {" & sTrack & "}.", False, True)
        CDWDVDWToWav = False
    Err.Clear
End Function

Private Function StripLeft(strString As String, strChar As String, Optional sLeftsRight As Boolean = True) As String
  On Local Error Resume Next
  Dim i As Integer
    If sLeftsRight Then
        For i = 1 To Len(strString)
            If Mid$(strString, i, 1) = strChar Then
                    StripLeft = Mid$(strString, 1, i - 1)
                Exit For
            End If
        Next
    Else
        For i = (Len(strString)) To 1 Step -1
        If Mid$(strString, i, 1) = strChar Then
                StripLeft = Mid$(strString, i + 2, Len(strString) - i + 1)
            Exit For
        End If
    Next
End If
End Function

Private Function SplitTrack(sString As String, Optional sDelimiter As Variant = "|") As String
    Dim i As Integer
    Dim strTrack As String
    Dim StringArray() As String
    On Local Error GoTo Error_Split
    lblTagManually.Caption = Empty
        StringArray = Split(sString, sDelimiter)
        For i = 0 To UBound(StringArray)
                If StringArray(i) <> "" Then
                    If LCase(StringArray(i)) = LCase(".year") Or LCase(StringArray(i)) = LCase("year") Then
                        lblTagManually.Caption = lblTagManually.Caption + "year|"
                    ElseIf LCase(StringArray(i)) = LCase("album") Then
                        lblTagManually.Caption = lblTagManually.Caption + "album|"
                    ElseIf LCase(StringArray(i)) = LCase("artist") Then
                        lblTagManually.Caption = lblTagManually.Caption + "artist|"
                    ElseIf LCase(StringArray(i)) = LCase("title") Then
                        lblTagManually.Caption = lblTagManually.Caption + "title|"
                    ElseIf LCase(StringArray(i)) = LCase("track") Then
                        lblTagManually.Caption = lblTagManually.Caption + "track|"
                    '/* by (.)
                    ElseIf LCase(StringArray(i)) = LCase(".track") Then
                        lblTagManually.Caption = lblTagManually.Caption + ".track|"
                    ElseIf LCase(StringArray(i)) = LCase(".album") Then
                        lblTagManually.Caption = lblTagManually.Caption + ".album|"
                    ElseIf LCase(StringArray(i)) = LCase(".artist") Then
                        lblTagManually.Caption = lblTagManually.Caption + ".artist|"
                    ElseIf LCase(StringArray(i)) = LCase(".title") Then
                        lblTagManually.Caption = lblTagManually.Caption + ".title|"
                    End If
                End If
            Next i
        lblTagManually.Caption = Mid$(lblTagManually.Caption, 1, Len(lblTagManually.Caption) - 1)
        SplitTrack = lblTagManually.Caption
Exit Function
Error_Split:
    Call WriteErrorLogs(Err.Number, Err.Description, "Function {SplitTrack}!" & vbCr _
        & "To parse string {" & sString & "}.", True, True)
        SplitTrack = "Error!"
    Err.Clear
End Function

Private Function Splitter(SplitString As String, SplitLetter As String) As Variant
    ReDim SplitArray(1 To 1) As Variant
    Dim TempLetter As String
    Dim TempSplit As String
    Dim i As Integer
    Dim x As Integer
    Dim StartPos As Integer
    SplitString = SplitString & SplitLetter
    For i = 1 To Len(SplitString)
        TempLetter = Mid(SplitString, i, Len(SplitLetter))
        If TempLetter = SplitLetter Then
            TempSplit = Mid(SplitString, (StartPos + 1), (i - StartPos) - 1)
            If TempSplit <> "" Then
                x = x + 1
                ReDim Preserve SplitArray(1 To x) As Variant
                SplitArray(x) = TempSplit
            End If
            StartPos = i
        End If
    Next i
    Splitter = SplitArray
End Function

Private Function GetTitleTrack(strString As String) As String
    Dim i As Integer
    Dim strTrack As String
    Dim StringArray() As String
    On Local Error GoTo Error_Split
        If lblTagManually.Caption = Empty Then
                MsgBox "You must set Manually the order of the Tags Tracks!", vbExclamation, App.Title
            Exit Function
        End If
        StringArray = Split(strString, "|")
            For i = 0 To UBound(StringArray)
                If StringArray(i) <> "" Then
                    If LCase(StringArray(i)) = LCase(".year") Or LCase(StringArray(i)) = LCase("year") Then
                        strTrack = strTrack + "." + txtYear.Text
                    ElseIf LCase(StringArray(i)) = LCase("album") Then
                        strTrack = strTrack + "-" + txtAlbum.Text
                    ElseIf LCase(StringArray(i)) = LCase("artist") Then
                        strTrack = strTrack + "-" + txtArtist.Text
                    ElseIf LCase(StringArray(i)) = LCase("title") Then
                        strTrack = strTrack + "-" + txtTitle.Text
                    ElseIf LCase(StringArray(i)) = LCase("track") Then
                        strTrack = strTrack + "-" + txtTrack.Text
                    '/* by (.)
                    ElseIf LCase(StringArray(i)) = LCase(".track") Then
                        strTrack = strTrack + "." + txtTrack.Text
                    ElseIf LCase(StringArray(i)) = LCase(".album") Then
                        strTrack = strTrack + "." + txtAlbum.Text
                    ElseIf LCase(StringArray(i)) = LCase(".artist") Then
                        strTrack = strTrack + "." + txtArtist.Text
                    ElseIf LCase(StringArray(i)) = LCase(".title") Then
                        strTrack = strTrack + "." + txtTitle.Text
                    End If
                End If
            Next i
        GetTitleTrack = Mid$(strTrack, 2, Len(strTrack))
Exit Function
Error_Split:
    Call WriteErrorLogs(Err.Number, Err.Description, "Function {GetTitleTrack}!", False, True)
        GetTitleTrack = ""
    Err.Clear
End Function

Private Function GetFilePath(ByVal FileName As String, strExtract As Extract) As String
    Select Case strExtract
        'Extract only extension of File
    Case 0
         GetFilePath = Mid$(FileName, InStrRev(FileName, ".", , vbTextCompare) + 1)
        'Extract only Filename and Extension
    Case 1
        GetFilePath = Mid$(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
        'Extract only FileName
   Case 2
        GetFilePath = StripString(Mid$(FileName, InStrRev(FileName, "\", , vbTextCompare) + 1))
        'Extract only Path
   Case 3
        GetFilePath = Mid$(FileName, 1, InStrRev(FileName, "\", , vbTextCompare) - 1)
   End Select
End Function

Private Function StripString(ByVal sString As String) As String
    Dim i As Integer
    Dim sTmp As String
    On Error Resume Next
    sTmp = Mid(sString, i + 1, Len(sString))
    For i = 1 To Len(sTmp)
      If Mid(sTmp, i, 1) = "." Then
        Exit For
    Else
        MyString = Mid(sString, i + 2, Len(sString))
    End If
Next
     StripString = Left(sTmp, i - 1)
End Function

Private Function WriteTagOfTrack(sFileName As String) As Boolean
    Dim ID3 As New clsID3
    On Local Error GoTo ErrorWriting
    If modeTrack = True Then Exit Function
    SetAttr sFileName, vbNormal
    With ID3
        .FileName = sFileName
        .Title = txtTitle.Text
        .Artist = txtArtist.Text
        .Album = txtAlbum.Text
        .Year = txtYear.Text
        If txtComment.Text = "" Then .Comments = txtArtist.Text & "/" & txtArtist.Text Else .Comments = txtComment.Text
        .TrackNumber = txtTrack.Text
        If txtBand.Text = "" Then .Band = txtArtist.Text Else .Band = txtBand.Text
        .Genre = txtGenre.Text
        .Composer = txtArtist.Text
        .Copyright = txtArtist.Text
        ' .... Extra Tags
        If CheckIncludeExtraTags.Value = 1 Then
            If txtEncodedBy.Text = "" Then .EncodedBy = "Salvo Cortesiano" Else .EncodedBy = txtEncodedBy.Text
            If txtCopyrightInfo.Text = "" Then .CopyrightInfo = "http://www.netshadows.it" Else .CopyrightInfo = txtCopyrightInfo.Text
            If txtLanguage.Text = "" Then .Languages = "Italian" Else .Languages = txtLanguage.Text
        End If
        .UpdateID3Tags
    End With
    Set ID3 = Nothing
    If CheckIncludeExtraTags.Value = 1 Then
        If CheckIncludeCover.Value = 1 Then
            If CoverOK = True And Dir$(CoverDir) <> "" Then
                'utcCover.LoadID3Cover sFileName
                'Set utcCover.Picture = LoadPicture(CoverDir)
                'utcCover.UpdateID3Cover sFileName
            End If
        End If
    End If
    SetAttr sFileName, vbArchive
    WriteTagOfTrack = True
Exit Function
ErrorWriting:
    WriteTagOfTrack = False
        Call WriteErrorLogs(Err.Number, Err.Description, "FormMain {Function: WriteTagOfTrack}!" & vbCr _
        & "To write tag of File: " & GetFilePath(sFileName, Only_FileName_and_Extension), True, True)
    Err.Clear
End Function

Private Function GetSpecialFolderLocation(CSIDL As Long) As String
    Dim sPath As String
    Dim pidl As Long
    If SHGetSpecialFolderLocation(0&, CSIDL, pidl) = S_OK Then
            sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then _
            GetSpecialFolderLocation = Left(sPath, InStr(sPath, Chr$(0)) - 1)
        Call CoTaskMemFree(pidl)
    End If
End Function

Private Sub tTimer_Timer()
   On Local Error Resume Next
    If MP3.IsPlaying Then lblPosition = "Time: " & MP3.Position
    ' .... Loop
    If MP3.Position = MP3.length And strLoop = True And CheckPlayAll.Value = 0 Then
        cmdpStop = True
        cmdpPlay = True
        lblDuration.Caption = "Total Time: " & MP3.length
    ' .... Play Next File
    ElseIf MP3.Position = MP3.length And strLoop = False And CheckPlayAll.Value = 1 Then
        cmdpStop = True
        cmdnNext = True
        lblDuration.Caption = "Total Time: " & MP3.length
    ' .... Stop Player
    ElseIf MP3.Position = MP3.length And strLoop = False And CheckPlayAll.Value = 0 Then
        cmdpStop = True
        tTimer.Enabled = False
    End If
End Sub
Private Sub AddItem2Array1D(ByRef VarArray As Variant, ByVal VarValue As Variant)
  Dim i  As Long
  Dim iVarType As Integer
  On Local Error Resume Next
  DoEvents
  iVarType = VarType(VarArray) - 8192
  i = UBound(VarArray)
  Select Case iVarType
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
      If VarArray(0) = 0 Then
        i = 0
      Else
        i = i + 1
      End If
    Case vbDate
      If VarArray(0) = "00:00:00" Then
        i = 0
      Else
        i = i + 1
      End If
    Case vbString
      If VarArray(0) = vbNullString Then
        i = 0
      Else
        i = i + 1
      End If
    Case vbBoolean
      If VarArray(0) = False Then
        i = 0
      Else
        i = i + 1
      End If
    Case Else
  End Select
  ReDim Preserve VarArray(i)
  VarArray(i) = VarValue
  DoEvents
End Sub

Private Sub AllFileInFolder(strPath As String, includeSubDirectory As Boolean)
    Dim lFLCount As Long
    Dim j As Integer
    Dim fOlder As Integer
    Dim totF As Integer
    Dim sDirOld As String
    Dim ix As Integer
    On Local Error GoTo ErrorHandler
    ReDim FileList(0) As String
    ix = 0
    FileList = AllFilesInFolders(strPath, includeSubDirectory, "*.*")
    lFLCount = UBound(FileList)
    lstmp3wav.Clear
    lstmp3wavPath.Clear
        For j = 0 To UBound(FileList)
            DoEvents
            If FileList(j) <> "" Then
                If LCase(Right$(FileList(j), 4)) = ".mp3" Or LCase(Right$(FileList(j), 4)) = ".wma" _
                    Or LCase(Right$(FileList(j), 4)) = ".wav" Or LCase(Right$(FileList(j), 4)) = ".mid" _
                    Or LCase(Right$(FileList(j), 4)) = ".snd" Or LCase(Right$(FileList(j), 4)) = ".au" _
                    Or LCase(Right$(FileList(j), 4)) = ".aif" Or LCase(Right$(FileList(j), 4)) = ".rmi" _
                    Or LCase(Right$(FileList(j), 4)) = ".midi" Or LCase(Right$(FileList(j), 4)) = ".wmv" _
                    Or LCase(Right$(FileList(j), 4)) = ".mp2" Or LCase(Right$(FileList(j), 4)) = ".mpeg" _
                    Or LCase(Right$(FileList(j), 4)) = ".mpg" Or LCase(Right$(FileList(j), 4)) = ".mpa" _
                    Or LCase(Right$(FileList(j), 4)) = ".mpe" Or LCase(Right$(FileList(j), 4)) = ".asf" _
                    Or LCase(Right$(FileList(j), 4)) = ".mp4" Then
                    If GetFilePath(FileList(j), Only_Path) <> sDirOld Then fOlder = fOlder + 1
                    ix = ix + 1
                    lstmp3wav.AddItem GetFilePath(FileList(j), Only_FileName_and_Extension)
                    lstmp3wavPath.AddItem GetFilePath(FileList(j), Only_Path) + "\"
                End If
            End If
            DoEvents
            sDirOld = GetFilePath(FileList(j), Only_Path)
            'If STOP_PRESSED = True Then Exit For
        Next
        lblInfoScan.Caption = "Files: " & ix & " in " & fOlder & " Folders!"
  Exit Sub
ErrorHandler:
    Call WriteErrorLogs(Err.Number, Err.Description, "{Sub: AllFileInFolder}", True, True)
    Err.Clear
End Sub

Private Function AllFilesInFolders(ByVal sFolderPath As String, Optional bWithSubFolders As Boolean = True, Optional strFlag As String = "*.*") As String()
    Dim sTemp As String
    Dim sDirIn As String
    ReDim sFilelist(0) As String
    ReDim sSubFolderList(0) As String
    ReDim sToProcessFolderList(0) As String
    Dim i As Integer, j As Integer
    sDirIn = sFolderPath
    If Not (Right$(sDirIn, 1) = "\") Then sDirIn = sDirIn & "\"
    On Local Error Resume Next
    sTemp = Dir$(sDirIn & strFlag)
    Do While sTemp <> ""
    DoEvents
      AddItem2Array1D sFilelist(), sDirIn & sTemp
      sTemp = Dir
      'lblStatus.Caption = GetShortFileName(sTemp)
      DoEvents
      'If STOP_PRESSED = True Then Exit Do
    Loop
    If bWithSubFolders Then
      sTemp = Dir$(sDirIn & strFlag, vbDirectory)
      Do While sTemp <> ""
      DoEvents
         If sTemp <> "." And sTemp <> ".." Then
            If (GetAttr(sDirIn & sTemp) And vbDirectory) = vbDirectory Then
              AddItem2Array1D sToProcessFolderList, sDirIn & sTemp
            End If
         End If
         sTemp = Dir
         'lblStatus.Caption = GetShortFileName(sTemp)
         DoEvents
         'If STOP_PRESSED = True Then Exit Do
      Loop
      If UBound(sToProcessFolderList) > 0 Or UBound(sToProcessFolderList) = 0 And sToProcessFolderList(0) <> "" Then
        For i = 0 To UBound(sToProcessFolderList)
          DoEvents
          sSubFolderList = AllFilesInFolders(sToProcessFolderList(i), bWithSubFolders)
          If UBound(sSubFolderList) > 0 Or UBound(sSubFolderList) = 0 And sSubFolderList(0) <> "" Then
            For j = 0 To UBound(sSubFolderList)
              DoEvents
              AddItem2Array1D sFilelist(), sSubFolderList(j)
              'lblStatus.Caption = GetShortFileName(sSubFolderList(j))
              DoEvents
              'If STOP_PRESSED = True Then Exit For
            Next
          End If
          DoEvents
          'If STOP_PRESSED = True Then Exit For
        Next
      End If
    End If
    AllFilesInFolders = sFilelist
Exit Function
End Function
