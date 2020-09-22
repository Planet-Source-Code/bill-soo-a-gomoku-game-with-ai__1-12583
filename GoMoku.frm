VERSION 5.00
Begin VB.Form GoMoku 
   Caption         =   "GoMoku"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame ControlFrame 
      Height          =   4935
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton ExitButton 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton NewButton 
         Caption         =   "&New Game"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label StatusLbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   4200
         Width           =   1815
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "GoMoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MouseX!, MouseY!
Private Done As Boolean
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim r%, c% 'row, column of last move

Private Sub Status(msg$)
StatusLbl = msg$
End Sub

Private Sub Board_Click()
'updates player move. Reads global MouseX,MouseY
'Modifies globals:
'   r%,c%
'   BoardArray
'   Movelist
'   BoardValue
Dim player%
Dim i&, j&

player% = 1
r% = Int(MouseY * (BOARDSIZE + 1) / Board.ScaleHeight)
c% = Int(MouseX * (BOARDSIZE + 1) / Board.ScaleWidth)
i& = r% * (BOARDSIZE + 1&) + CLng(c%)
If BoardArray(r%, c%) = 0 Then 'is valid move?
    BoardArray(r%, c%) = player%  'update board
    For j& = 0 To UBound(MoveList)  'update Movelist
        If MoveList(j&, 0) = i& Then
            MoveList(j&, 1) = player%
            Exit For
        End If
    Next j&
    Done = True
    UpdateBoardValue r%, c%
    DrawBoard
Else 'not valid move
    Beep
End If
End Sub

Private Sub Board_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'stores the mouse coordinates for the Click event
MouseX = X
MouseY = Y
End Sub

Private Sub ExitButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
'initialize settings
Plys = 2
BOARDSIZE = 8
WinLength = 5
Me.Show
NewButton_Click
End Sub

Private Sub DrawBoard()
'draws the BoardArray to the Board
'highlights the last move as defined by r%,c%

Dim i%, r1%, c1%
Dim dx!, dy!, X!, Y!

If AllDone Then Exit Sub
dx! = Board.ScaleWidth / (BOARDSIZE + 1)
dy! = Board.ScaleHeight / (BOARDSIZE + 1)
X! = dx!
Y! = dy!
Board.Cls
Board.ForeColor = vbBlack
For i% = 1 To BOARDSIZE
    Board.Line (X!, 0)-(X!, Board.ScaleHeight)
    Board.Line (0, Y!)-(Board.ScaleWidth, Y!)
    X! = X! + dx!
    Y! = Y! + dy!
Next i%
For r1% = 0 To BOARDSIZE
    For c1% = 0 To BOARDSIZE
        If BoardArray(r1%, c1%) Then
            If BoardArray(r1%, c1%) = 2 Then Board.FillColor = QBColor(15) Else Board.FillColor = QBColor(0)
            X! = dx! * c1% + dx! / 2
            Y! = dy! * r1% + dy! / 2
            If (r1% = r%) And (c1% = c%) Then Board.ForeColor = vbRed Else Board.ForeColor = Board.FillColor
            Board.Circle (X!, Y!), dx! / 2 - 2 * Screen.TwipsPerPixelX
        End If
    Next c1%
Next r1%
Board.Refresh
DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
AllDone = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub NewButton_Click()
'sets up and starts a new game
NewButton.Enabled = False 'prevent double clicking
AllDone = True
DoEvents
SettingsFrm.Show vbModal
ReDim BoardArray(BOARDSIZE, BOARDSIZE) As Byte
ReDim MoveList((BOARDSIZE + 1) * (BOARDSIZE + 1) - 1, 1) As Integer
InitMoveList
For r% = 0 To BOARDSIZE
    For c% = 0 To BOARDSIZE
        BoardArray(r%, c%) = 0
    Next c%
Next r%
AllDone = False
DrawBoard
NewButton.Enabled = True
MainLoop
End Sub

Private Sub MainLoop()
'main game loop. Basically move, then check for victory. Stop when draw or win or exit

Dim player%
Dim best&, v&, move&, total&, possible&
Dim p%, i%
Dim t&

player% = 1 'human goes first
AllDone = False
total& = (BOARDSIZE + 1&) * (BOARDSIZE + 1&) 'total possible moves
BoardValue = 0 'initial empty board is worth 0
Do While Not AllDone
'do player move
    If player% = 1 Then 'move manual
        Board.Enabled = True
        Done = False
        While (Not Done) And (Not AllDone) 'wait for user input
            DoEvents
        Wend
        Board.Enabled = False
    Else 'computer move
        Screen.MousePointer = vbHourglass
        t& = timeGetTime()
        MoveCount = 0
        best& = MINVALUE
        AlphaBetaSearch Plys, best&, player%, move&
        'move& is index of best move
        v& = MoveList(move&, 0)
        MoveList(move&, 1) = player%
        r% = v& \ (BOARDSIZE + 1&)
        c% = v& Mod (BOARDSIZE + 1&)
        BoardArray(r%, c%) = player%
        UpdateBoardValue r%, c%
        DrawBoard
        possible = total
        For i% = 1 To Plys - 1
            possible = possible * (total - i%)
        Next i%
        t& = timeGetTime() - t&
        If MoveCount > 0 Then
            Status CStr(MoveCount) & " of " & possible & " in " & CStr(t& / 1000#) & " seconds (" & Format$(t& / MoveCount, "0.000") & " ms/move)"
        End If
        Screen.MousePointer = vbNormal
    End If
    'check for victory
    p% = VictoryCheck()
    If p% Then
        AllDone = True
        MsgBox "Player" & CStr(p%) & " has won!"
    End If
    If player% = 1 Then player% = 2 Else player% = 1 'toggle players
    total& = total& - 1 'decrement # of remaining moves
    If total = 0 Then
        AllDone = True
        MsgBox "Draw!"
    End If
Loop
End Sub


