VERSION 5.00
Begin VB.Form frmCategorie 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Category:"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5805
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3150
      TabIndex        =   3
      Top             =   2430
      Width           =   1260
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Select"
      Height          =   330
      Left            =   4575
      TabIndex        =   2
      Top             =   2430
      Width           =   1110
   End
   Begin VB.ListBox lstCategories 
      Height          =   1740
      Left            =   60
      TabIndex        =   0
      Top             =   555
      Width           =   5640
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click in the List to Select the Category if you want or Click {&Select}:"
      Height          =   465
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   4980
   End
End
Attribute VB_Name = "frmCategorie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    frmMain.Tag = "Abort"
    Unload Me
End Sub

Private Sub cmdOk_Click()
    frmMain.Tag = lstCategories.List(lstCategories.ListIndex)
    Unload Me
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Me.Caption = " Category: " & lstCategories.List(lstCategories.ListIndex)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Local Error Resume Next
    Set frmCategorie = Nothing
End Sub


Private Sub lstCategories_Click()
    On Local Error Resume Next
    Me.Caption = " Category: " & lstCategories.List(lstCategories.ListIndex)
End Sub

Private Sub lstCategories_DblClick()
    On Local Error Resume Next
    frmMain.Tag = lstCategories.List(lstCategories.ListIndex)
    Unload Me
End Sub


