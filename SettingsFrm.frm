VERSION 5.00
Begin VB.Form SettingsFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Settings"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.HScrollBar PlyScroll 
      Height          =   255
      Left            =   120
      Max             =   6
      Min             =   1
      TabIndex        =   8
      Top             =   1800
      Value           =   2
      Width           =   1815
   End
   Begin VB.HScrollBar WinScroll 
      Height          =   255
      Left            =   120
      Max             =   6
      Min             =   3
      TabIndex        =   5
      Top             =   1080
      Value           =   5
      Width           =   1815
   End
   Begin VB.HScrollBar SizeScroll 
      Height          =   255
      Left            =   120
      Max             =   19
      Min             =   6
      TabIndex        =   2
      Top             =   360
      Value           =   8
      Width           =   1815
   End
   Begin VB.Label PlyLbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      Caption         =   "Plys to search:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label WinLbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      Caption         =   "# needed to win:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label SizeLbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lbls 
      AutoSize        =   -1  'True
      Caption         =   "Board Size"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   765
   End
End
Attribute VB_Name = "SettingsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
PlyScroll.Value = Plys
SizeScroll.Value = BOARDSIZE + 1
WinScroll.Value = WinLength
End Sub

Private Sub OKButton_Click()
BOARDSIZE = SizeScroll.Value - 1
Plys = PlyScroll.Value
WinLength = WinScroll.Value
Unload Me
End Sub

Private Sub PlyScroll_Change()
PlyLbl = CStr(PlyScroll)
End Sub

Private Sub SizeScroll_Change()
SizeLbl = CStr(SizeScroll)
End Sub

Private Sub WinScroll_Change()
WinLbl = CStr(WinScroll)
End Sub

