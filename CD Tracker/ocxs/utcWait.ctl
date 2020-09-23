VERSION 5.00
Begin VB.UserControl utcWait 
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   ScaleHeight     =   1020
   ScaleWidth      =   945
   ToolboxBitmap   =   "utcWait.ctx":0000
   Begin VB.Timer tStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   15
      Top             =   405
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   0
      Picture         =   "utcWait.ctx":0312
      Top             =   0
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   11
      Left            =   6345
      Picture         =   "utcWait.ctx":09FC
      Top             =   135
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   10
      Left            =   6090
      Picture         =   "utcWait.ctx":10E6
      Top             =   120
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   9
      Left            =   5850
      Picture         =   "utcWait.ctx":17D0
      Top             =   120
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   8
      Left            =   5595
      Picture         =   "utcWait.ctx":1EBA
      Top             =   120
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   7
      Left            =   5325
      Picture         =   "utcWait.ctx":25A4
      Top             =   135
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   6
      Left            =   5070
      Picture         =   "utcWait.ctx":2C8E
      Top             =   120
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   5
      Left            =   4800
      Picture         =   "utcWait.ctx":3378
      Top             =   120
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   4
      Left            =   4545
      Picture         =   "utcWait.ctx":3A62
      Top             =   135
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   3
      Left            =   4290
      Picture         =   "utcWait.ctx":414C
      Top             =   150
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   2
      Left            =   4020
      Picture         =   "utcWait.ctx":4836
      Top             =   135
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   1
      Left            =   3765
      Picture         =   "utcWait.ctx":4F20
      Top             =   150
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   0
      Left            =   3510
      Picture         =   "utcWait.ctx":560A
      Top             =   150
      Width           =   360
   End
End
Attribute VB_Name = "utcWait"
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

Private i As Integer

Private m_start As Boolean
Private Const m_def_start As Boolean = False
Private Sub tStart_Timer()
    i = i + 1
    Image1.Picture = img(i).Picture
    If i >= 11 Then i = 0
End Sub

Private Sub UserControl_InitProperties()
    m_start = m_def_start
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_start = PropBag.ReadProperty("Start", m_def_start)
End Sub


Private Sub UserControl_Resize()
    UserControl.Height = 360
    UserControl.Width = 360
End Sub

Public Property Get Start() As Boolean
    Start = m_start
End Property

Public Property Let Start(ByVal NewStart As Boolean)
    m_start = NewStart
    If m_start = True Then
        tStart.Enabled = True
    Else
        tStart.Enabled = False
        Image1.Picture = img(0).Picture
    End If
    PropertyChanged "Start"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Start", m_start, m_def_start)
End Sub


