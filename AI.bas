Attribute VB_Name = "AI"
Option Explicit
Public Const MAXVALUE = 1000000000 'greater than any possible board value
Public Const MINVALUE = -1000000000

Public BOARDSIZE As Integer 'rows or columns, but -1 since array is zero based
Public WinLength As Integer 'number in a row needed to win
Public Plys As Integer 'depth of search
Public BoardArray() As Byte '15x15 playing board. 0=empty 1=player1 2=player2
Public MoveList() As Integer 'array of valid moves. 2 columns. 1st column is move, 2nd is player
Public AllDone As Boolean 'flag to signal end of game or quit
Private Values() As Long 'value of each combination
Public MoveCount As Long 'number of moves investigated
Public BoardValue As Long 'value of the board. Large is better for player1, small is better for player2

Private Declare Function GetInputState Lib "user32" () As Long

Public Sub UpdateBoardValue(r%, c%)
'changes the global BoardValue based on putting player at r%,c%

Dim i%, player%, j%, wl%
Dim v&
Dim pcnt%, ocnt%

v& = 0
wl% = WinLength - 1
player% = BoardArray(r%, c%)
'check all possible vertical lines that include r%,c%
For i% = r% - wl% To r%
    If (i% >= 0) And ((i% + wl%) <= BOARDSIZE) Then
        pcnt% = 0
        ocnt% = 0
        For j% = i% To i% + wl%
            Select Case BoardArray(j%, c%)
                Case player%
                    If j% <> r% Then pcnt% = pcnt% + 1
                Case 0
                Case Else
                    ocnt% = ocnt% + 1
            End Select
        Next j%
        If ocnt% > 0 Then
            If pcnt% = 0 Then v& = v& + Values(ocnt%) 'get points for denying opponent
        Else
            If pcnt% = wl% Then GoTo BreakOnWin 'if line wins game, then break
            v& = v& + Values(pcnt% + 1) - Values(pcnt%) 'get points for increasing length
        End If
    End If
Next i%
'check all possible horizontal lines that include r%,c%
For i% = c% - wl% To c%
    If (i% >= 0) And ((i% + wl%) <= BOARDSIZE) Then
        pcnt% = 0
        ocnt% = 0
        For j% = i% To i% + wl%
            Select Case BoardArray(r%, j%)
                Case player%
                    If j% <> c% Then pcnt% = pcnt% + 1
                Case 0
                Case Else
                    ocnt% = ocnt% + 1
            End Select
        Next j%
        If ocnt% > 0 Then
            If pcnt% = 0 Then v& = v& + Values(ocnt%) 'get points for denying opponent
        Else
            If pcnt% = wl% Then GoTo BreakOnWin
            v& = v& + Values(pcnt% + 1) - Values(pcnt%) 'get points for increasing length
        End If
    End If
Next i%
'check diagonals from top left to bottom right
For i% = 0 To wl%
    If ((c% - i%) >= 0) And (((c% - i%) + wl%) <= BOARDSIZE) And ((r% - i%) >= 0) And (((r% - i%) + wl%) <= BOARDSIZE) Then
        pcnt% = 0
        ocnt% = 0
        For j% = 0 To wl%
            Select Case BoardArray(r% - i% + j%, c% - i% + j%)
                Case player%
                    If i% <> j% Then pcnt% = pcnt% + 1
                Case 0
                Case Else
                    ocnt% = ocnt% + 1
            End Select
        Next j%
        If ocnt% > 0 Then
            If pcnt% = 0 Then v& = v& + Values(ocnt%) 'get points for denying opponent
        Else
            If pcnt% = wl% Then GoTo BreakOnWin
            v& = v& + Values(pcnt% + 1) - Values(pcnt%) 'get points for increasing length
        End If
    End If
Next i%
'check diagonals from bottom left to top right
For i% = 0 To wl%
    If ((c% - i%) >= 0) And (((c% - i%) + wl%) <= BOARDSIZE) And ((r% + i%) - wl% >= 0) And (((r% + i%)) <= BOARDSIZE) Then
        pcnt% = 0
        ocnt% = 0
        For j% = 0 To wl%
            Select Case BoardArray(r% + i% - j%, c% - i% + j%)
                Case player%
                    If i% <> j% Then pcnt% = pcnt% + 1
                Case 0
                Case Else
                    ocnt% = ocnt% + 1
            End Select
        Next j%
        If ocnt% > 0 Then
            If pcnt% = 0 Then v& = v& + Values(ocnt%) 'get points for denying opponent
        Else
            If pcnt% = wl% Then GoTo BreakOnWin
            v& = v& + Values(pcnt% + 1) - Values(pcnt%) 'get points for increasing length
        End If
    End If
Next i%
If player% = 2 Then v& = -v&  'player 2 likes small boardvalues
BoardValue = BoardValue + v&
Exit Sub
BreakOnWin:
If player% = 2 Then BoardValue = MINVALUE Else BoardValue = MAXVALUE
End Sub

Public Sub InitMoveList()
'load movelist in such a way that centre moves are first in list
Dim r%, c%, i%, n%, v%, lasti%
Dim lv&
ReDim movevalue(1, (BOARDSIZE + 1) * (BOARDSIZE + 1)) As Integer
'uses a kind of linked list structure.
'movevalue(0,x) is the value for square x
'movevalue(1,x) is the index of the next square

'also init the values table
lv& = 0
ReDim Values(WinLength) As Long
For i% = 0 To WinLength
    Values(i%) = lv&
    lv& = lv& * 10 + 1
Next i%

n% = 1
movevalue(1, 0) = 0
For r% = 0 To BOARDSIZE
    For c% = 0 To BOARDSIZE
        v% = (Abs(r% - BOARDSIZE \ 2) + 1) * (Abs(c% - BOARDSIZE \ 2) + 1)
        movevalue(0, n%) = v%
        i% = movevalue(1, 0)
        lasti% = 0
        Do While i%
            If v% < movevalue(0, i%) Then 'insert
                movevalue(1, n%) = i%
                movevalue(1, lasti%) = n%
                Exit Do
            End If
            lasti% = i%
            i% = movevalue(1, i%)
        Loop
        If i% = 0 Then 'no insertion, add to end
            movevalue(1, n%) = 0
            movevalue(1, lasti%) = n%
        End If
        n% = n% + 1
    Next c%
Next r%
i% = movevalue(1, 0)
n% = 0
While i%
    MoveList(n%, 0) = i% - 1
    MoveList(n%, 1) = 0  'not picked
    i% = movevalue(1, i%)
    n% = n% + 1
Wend
End Sub

Public Function AlphaBetaSearch(ByVal ply%, ByVal limit&, ByVal player%, bestmove&) As Long
'returns the value of the board
'aborts if the branch it is searching is worse than the limit
Dim best&, i&, n&, v&, move&
Dim r%, c%, otherplayer%
Dim lastbv&

If AllDone Then Exit Function

If ply% Then 'not at bottom of recursive branch yet...
    n& = (BOARDSIZE + 1&) * (BOARDSIZE + 1&) - 1&
    If player = 1 Then
        best& = MINVALUE
        otherplayer% = 2
    Else
        best& = MAXVALUE
        otherplayer% = 1
    End If
    'check all possible moves from this location...
    For i& = 0 To n&
        If GetInputState() Then DoEvents
        If MoveList(i&, 1) = 0 Then
            r% = MoveList(i&, 0) \ (BOARDSIZE + 1&)
            c% = MoveList(i&, 0) Mod (BOARDSIZE + 1&)
            'we've selected a move...
            MoveList(i&, 1) = player        '...so remove it from list...
            BoardArray(r%, c%) = player     '...and update board...
            lastbv& = BoardValue            '...but save the old boardvalue for later
            UpdateBoardValue r%, c%         'calculate the new boardvalue
            If Abs(BoardValue) = MAXVALUE Then 'break since win was found
                v& = BoardValue
            Else
                v& = AlphaBetaSearch(ply% - 1, best&, otherplayer%, move&) 'and recursively follow branch
            End If
            If AllDone Then Exit For
            'return board and movelist to original state
            BoardValue = lastbv&
            BoardArray(r%, c%) = 0
            MoveList(i&, 1) = 0
            If v& = best Then 'insert a bit of randomness
                If Rnd() < 0.3 Then
                    bestmove& = i&
                End If
            End If
            If player = 1 Then ' Pick largest
                If v& > best& Then
                    best& = v&
                    bestmove& = i&
                End If
                If v& > limit Then 'we are past the limit, meaning there is no point exploring this branch further
                    Exit For
                End If
            Else 'Pick smallest
                If v& < best& Then
                    best& = v&
                    bestmove& = i&
                End If
                If v& < limit Then
                    Exit For
                End If
           End If
        End If
    Next i&
    AlphaBetaSearch = best&
Else  'we are at the bottom of the branch. Just return the board value
    MoveCount = MoveCount + 1
    AlphaBetaSearch = BoardValue
End If
End Function

Function VictoryCheck() As Integer
'checks for victory. Returns 1 or 2 for player 1 or 2 or 0 if no one won.
Dim r%, c%, i%, pl%, r1%, c1%, r2%, c2%, j%, wl%

wl% = WinLength - 1

For r% = 0 To BOARDSIZE
    For c% = 0 To BOARDSIZE
        If r% <= (BOARDSIZE - wl%) Then 'calculate vertical
            pl% = 0
            i% = 0
            For r1% = r% To r% + wl%
                Select Case BoardArray(r1%, c%)
                Case 0 'nothing
                Case 1
                    Select Case pl%
                    Case 0
                        pl% = 1
                        i% = 1
                    Case 1
                        i% = i% + 1
                    Case 2
                        i% = 0
                        Exit For
                    End Select
                Case 2
                    Select Case pl%
                    Case 0
                        pl% = 2
                        i% = 1
                    Case 2
                        i% = i% + 1
                    Case 1
                        i% = 0
                        Exit For
                    End Select
                End Select
            Next r1%
            If i% = WinLength Then
                VictoryCheck = pl%
                Exit Function
            End If
        End If
        If c% <= (BOARDSIZE - wl%) Then 'calculate horizontal
            pl% = 0
            i% = 0
            For c1% = c% To c% + wl%
                Select Case BoardArray(r%, c1%)
                Case 0 'nothing
                Case 1
                    Select Case pl%
                    Case 0
                        pl% = 1
                        i% = 1
                    Case 1
                        i% = i% + 1
                    Case 2
                        i% = 0
                        Exit For
                    End Select
                Case 2
                    Select Case pl%
                    Case 0
                        pl% = 2
                        i% = 1
                    Case 2
                        i% = i% + 1
                    Case 1
                        i% = 0
                        Exit For
                    End Select
                End Select
            Next c1%
            If i% = WinLength Then
                VictoryCheck = pl%
                Exit Function
            End If
        End If
        If (r% <= (BOARDSIZE - wl%)) And (c% <= (BOARDSIZE - wl%)) Then 'diag1
            i% = 0
            pl% = 0
            For j% = 0 To wl%
                Select Case BoardArray(r% + j%, c% + j%)
                Case 0 'nothing
                Case 1
                    Select Case pl%
                    Case 0
                        pl% = 1
                        i% = 1
                    Case 1
                        i% = i% + 1
                    Case 2
                        i% = 0
                        Exit For
                    End Select
                Case 2
                    Select Case pl%
                    Case 0
                        pl% = 2
                        i% = 1
                    Case 2
                        i% = i% + 1
                    Case 1
                        i% = 0
                        Exit For
                    End Select
                End Select
            Next j%
            If i% = WinLength Then
                VictoryCheck = pl%
                Exit Function
            End If
        End If
        If (r% <= (BOARDSIZE - wl%)) And (c% >= wl%) Then  'diag2
            i% = 0
            pl% = 0
            For j% = 0 To wl%
                Select Case BoardArray(r% + j%, c% - j%)
                Case 0 'nothing
                Case 1
                    Select Case pl%
                    Case 0
                        pl% = 1
                        i% = 1
                    Case 1
                        i% = i% + 1
                    Case 2
                        i% = 0
                        Exit For
                    End Select
                Case 2
                    Select Case pl%
                    Case 0
                        pl% = 2
                        i% = 1
                    Case 2
                        i% = i% + 1
                    Case 1
                        i% = 0
                        Exit For
                    End Select
                End Select
            Next j%
            If i% = WinLength Then
                VictoryCheck = pl%
                Exit Function
            End If
        End If
    Next c%
Next r%
DoEvents
VictoryCheck = 0
End Function
