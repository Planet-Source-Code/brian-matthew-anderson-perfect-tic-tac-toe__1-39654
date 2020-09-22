VERSION 5.00
Begin VB.Form TicTacToe 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6945
   DrawWidth       =   3
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   27.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar Intelligence 
      Height          =   255
      Left            =   4560
      Max             =   5
      Min             =   1
      TabIndex        =   14
      Top             =   3960
      Value           =   1
      Width           =   1695
   End
   Begin VB.CheckBox HumanX 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Human plays X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox HumanFirst 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Human goes first"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton PlayAgain 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Label IntelligenceLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Intelligence Level:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   4560
      TabIndex        =   16
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   7
      Visible         =   0   'False
      X1              =   4200
      X2              =   600
      Y1              =   600
      Y2              =   4200
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   6
      Visible         =   0   'False
      X1              =   600
      X2              =   4200
      Y1              =   600
      Y2              =   4200
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   5
      Visible         =   0   'False
      X1              =   3600
      X2              =   3600
      Y1              =   600
      Y2              =   4200
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   4
      Visible         =   0   'False
      X1              =   2400
      X2              =   2400
      Y1              =   600
      Y2              =   4200
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   3
      Visible         =   0   'False
      X1              =   1200
      X2              =   1200
      Y1              =   600
      Y2              =   4200
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   2
      Visible         =   0   'False
      X1              =   480
      X2              =   4320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   1
      Visible         =   0   'False
      X1              =   480
      X2              =   4320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Win 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   0
      Visible         =   0   'False
      X1              =   480
      X2              =   4320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   3120
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   1920
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   720
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   3120
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   1920
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   720
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Square 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tic - Tac - Toe"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.Line Line4 
      X1              =   600
      X2              =   4200
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   600
      X2              =   4200
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   3000
      X2              =   3000
      Y1              =   600
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1800
      Y1              =   600
      Y2              =   4200
   End
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StartSquare As Integer, EndSquare As Integer, Direction As Integer
Dim XO(2) As String
Dim Blank As Integer
Dim CaseCounter As Integer
Dim XsInARow As Integer, OsInARow As Integer
Dim HumanPlayingAsX As Boolean
Dim HumanMovedAlready As Boolean, CompMovedAlready As Boolean
Dim GameOver As Boolean
Dim SquareValue(8) As Integer
Dim RandChoice(8) As Integer
Dim WinLine As Integer
Dim TieGame As Boolean
Dim strCPU As String
Dim strHuman As String


'This program was created by Brian Anderson.
'It is my first real project in VB.
'The comments are easy to follow and will help
'you develop similar programs where low-level A.I. is involved.
'The computer is impossible to beat on the highest
'skill level, but it can be beat at any level lower
'than "genius". If you can't seem to beat it on the "Clever" level
'then you might want to analyze the code. It CAN be beaten at that
'level about 33% of the time if you know what you're doing  ;)

'The only thing I'd like to clean up is the fact that its not
'obvious how the game actually begins. It starts in one of two ways:
'Either the human clicks on a non-occupied square, or the human
'selects the computer to go first.


Private Sub Form_Load()
'Setting up default state of
'all variables.

'Makes all squares clear
Dim Index As Integer

Cls
Randomize

'Clears squares and any red line marking a win on the last game.
For Index = 0 To 8
Square(Index).Caption = ""
SquareValue(Index) = 0
Next Index

For Index = 0 To 7
Win(Index).Visible = False
Next Index

'Keeps Human as playing same type (X or O) as he did in
'previous game. The default side is X.
If HumanX.Value = 1 Then
    HumanPlayingAsX = True
    strCPU = "O"
    strHuman = "X"
Else
    HumanPlayingAsX = False
    strCPU = "X"
    strHuman = "O"
End If

'Defaults Human as not having selected a square yet
HumanMovedAlready = False
CompMovedAlready = False

'CheckBox for choosing side and
'who plays first are visible.
'Human defaults to being first.
'Displays Intelligence level of Computer.
HumanX.Visible = True
HumanFirst.Visible = True
HumanFirst.Value = 1
Call IntelligenceLevelCase

'These variables are true only if the game ends with a win or a cat's game.
GameOver = False
TieGame = False

'Changes the PlayAgain command button caption back to "New Game"
PlayAgain.Caption = "New Game"

End Sub


Private Sub Square_Click(Index As Integer)
Dim Flash100 As Integer

'This Procedure checks for a click on a square
'and determines several things:
    '1) Is the move legal?
    '2) Is it the human player's turn?
'If both conditions are true, the player's
'square is selected - if either condition is not true,
'a beep and screen flash effect alert the player
'that they tried to make an illegal move.
If Square(Index).Caption = "" And HumanMovedAlready = False Then
    HumanMovedAlready = True
    Square(Index).Caption = strHuman
    Call CheckForWin
    
    If GameOver = False Then
        Call CheckForCatsGame
        Call ComputerTurn
    End If
    
'Checks for an illegal move
'Prevents player from clicking to move to an
'already occupied square.
Else
    Beep
    For Flash100 = 1 To 100
        Me.BackColor = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
        Me.BackColor = "&H00C0FFFF"
    Next Flash100
End If
End Sub




Private Sub HumanFirst_Click()
'Skips to computer's turn if human player
'opts to play second
If HumanFirst.Value = 1 Then
    HumanMovedAlready = False
Else
    HumanMovedAlready = True
    Cls
    Call ComputerTurn
End If
End Sub


Private Sub HumanX_Click()
If HumanX.Value = 1 Then
    HumanPlayingAsX = True
    strCPU = "O"
    strHuman = "X"
Else
    HumanPlayingAsX = False
    strCPU = "X"
    strHuman = "O"
End If

End Sub


Private Sub Intelligence_Change()
Dim Red As Integer, Green As Integer, Blue As Integer

'Changes the color of the caption which tells which level of intelligence
'the human player is up against. The color goes from blue (easy) to
'red (hard)
Red = (Intelligence.Value / Intelligence.Max) * 255
Green = 0
Blue = (1 / Intelligence.Value) * 255
Level.ForeColor = RGB(Red, Green, Blue)
Call IntelligenceLevelCase
End Sub


Private Sub PlayAgain_Click()
Dim Index As Integer
Dim OldColor As String

'   This section of code rotates the background color
'of the selectable squares for a nice visual effect.
OldColor = Square(0).BackColor
    For Index = 0 To 7
        Square(Index).BackColor = Square(Index + 1).BackColor
    Next Index
Square(8).BackColor = OldColor

'   And now we start a new game!
Form_Load
End Sub


Private Sub Quit_Click()
Unload Me
End Sub


Private Sub XWins()
Call GameIsOver
MsgBox ("X wins!")
End Sub


Private Sub OWins()
GameOver = True
Call GameIsOver
MsgBox ("O Wins!")
End Sub


Private Sub CatsGame()
TieGame = True
Call GameIsOver
MsgBox ("Cat's Game!")
End Sub


Private Sub GameIsOver()
GameOver = True
PlayAgain.Caption = "Play Again?"
Beep
If TieGame = False Then Win(CaseCounter).Visible = True
End Sub


Private Sub BoardCheck()
'This procedure runs through the CaseCounter routine
'Which we will use in several different areas to check
'for a win by either side.
'The computer A.I. uses this routine to check for
'which moves would be the most advantageous to make.

'Squares are numbered in our array as follows:

'       0|1|2
'       -----
'       3|4|5
'       -----
'       6|7|8

Select Case CaseCounter
               
'These next 3 cases allow our FOR... NEXT loop to
'check the status of the horizontal rows of the board.
    Case 0
        StartSquare = 0
        EndSquare = 2
        Direction = 1
    Case 1
        StartSquare = 3
        EndSquare = 5
        Direction = 1
    Case 2
        StartSquare = 6
        EndSquare = 8
        Direction = 1
            
'These next 3 cases allow our FOR... NEXT loop to
'check the status of the vertical columns of the board
    Case 3
        StartSquare = 0
        EndSquare = 6
        Direction = 3
    Case 4
        StartSquare = 1
        EndSquare = 7
        Direction = 3
    Case 5
        StartSquare = 2
        EndSquare = 8
        Direction = 3
            
'These last 2 cases allow our FOR... NEXT loop to
'check the status of the diagonals of the board.
    Case 6
        StartSquare = 0
        EndSquare = 8
        Direction = 4
    Case 7
        StartSquare = 2
        EndSquare = 6
        Direction = 2
End Select
End Sub


Private Sub XOSelect(XOSquare)
'This Subroutine allows us to check for varying combinations of squares
'that can win the game or save the game from being lost. The Blank variable
'borrows expressions from our CaseCounter Select Case subroutine. It is used
'to identify the open square in a given sequence so that a point value can
'be assigned to it later.

Select Case XOSquare
    Case 1
        XO(0) = "X"
        XO(1) = "X"
        XO(2) = ""
        Blank = EndSquare
    Case 2
        XO(0) = "X"
        XO(1) = ""
        XO(2) = "X"
        Blank = StartSquare + Direction
    Case 3
        XO(0) = ""
        XO(1) = "X"
        XO(2) = "X"
        Blank = StartSquare
    Case 4
        XO(0) = "O"
        XO(1) = "O"
        XO(2) = ""
        Blank = EndSquare
    Case 5
        XO(0) = "O"
        XO(1) = ""
        XO(2) = "O"
        Blank = StartSquare + Direction
    Case 6
        XO(0) = ""
        XO(1) = "O"
        XO(2) = "O"
        Blank = StartSquare
End Select
End Sub


Private Sub IntelligenceLevelCase()
    Select Case Intelligence.Value
        Case 1
            Level.Caption = "Mindless"
        Case 2
            Level.Caption = "Poor"
        Case 3
            Level.Caption = "Average"
        Case 4
            Level.Caption = "Clever"
        Case 5
            Level.Caption = "Genius"
    End Select
End Sub


Private Sub CheckForCatsGame()
'If all the spaces are filled, the game ends.
Dim Index As Integer
Dim Space As Integer

For Index = 0 To 8
    If Square(Index).Caption = "" Then Space = Space + 1
Next Index

If Space = 0 Then
    Call CatsGame
End If
End Sub


Private Sub SpecialAI()
Dim MiddleSideSquare As Integer
Dim Index As Integer

'Apology in advance: I got a little long winded here. Sometimes my
'explanations for what I'm doing are more complicated than
'the actual code I'm trying to describe.  :P

'There are 3 methods by which a human can try to "trick"
'his opponent into winning. They all involve moves which
'give a potential win in two directions at the same time.
'The first method uses the "two diagonal" method.
'The second method uses the "two sides" method.
'The third method uses the "one side, one diagonal" method.

'*************************************************************************
'Method 1 of 3:
'The following code is checking for a special scenario in
'which the human could try to win by taking two diagonals
'early on so that the computer is "fooled" into making a bad move.
'The classic set-up goes like this:
'With human playing X -

' Figure 1:
'           X| |
'           -----
'            |O|
'           -----
'            | |X

'If the computer chooses either diagonal, it will lose because
'the human player will take the opposite diagonal and win the game with:

'           X|*|X
'           -----
'            |O|*
'           -----
'           O| |X
'
'By taking either side marked by an asterisk (*) for the win.

'The solution to this problem is to force the human to respond
'to an immediate attempt by the computer to win on a row or column
'as opposed to a diagonal. This can be done by changing the values
'normally given to the more desirable corner squares to those of
'the middle peripheral squares.

'The way we check this is to add up the number of squares shown in
'the following diagram. (the * symbol signifies the squares we are adding:

'           *| |          | |*
'           -----        -----
'            | |   or     | |
'           -----        -----
'            | |*        *| |

'If 2 of these squares are taken, it means the human is trying to
'win through the method depicted in figure 1. First we check if either
'situation depicted in figure 2 is happening. If so, we make all the
'values of the middle side squares higher than the diagonals.

If (Square(0).Caption = strHuman And Square(8).Caption = strHuman) _
Or (Square(2).Caption = strHuman And Square(6).Caption = strHuman) Then
    
    For Index = 1 To 7 Step 2
        SquareValue(Index) = 10
    Next Index
    
End If


'*************************************************************************
'Method 2 of 3:
'This method is kind of the opposite of method 1. Now we're talking
'about taking two middle side squares instead of two diagonals.
'Here is how it follows: (with computer as "O" and human going first)
'
'   turn 1:  |X|    |  turn 2:  |X|*
'           -----   |          -----
'            |O|    |           |O|X
'           -----   |          -----
'            | |    |          O| |
'
'Notice how the human player can win by taking the square
'marked * by creating the problem of threatening to win in 2 places
'at the same time. Since the computer can block only one square, it
'will lose. The key to overcoming this situation is to tell the
'computer that taking a diagonal on the opposite end of the human's
'two side squares is a bad move.
'
'First things first: We must identify if the human is taking
'two side squares that are close to each other. The following combination
'of squares is what our SELECT CASE procedure will look for:
'
'     1)   |X|    2)  *| |    3)   | |*   4)   |X|
'         -----       -----       -----       -----
'          |O|X        |O|X       X|O|        X|O|
'         -----       -----       -----       -----
'         *| |         |X|         |X|         | |*
'
'If any of these cases are true, then the computer must be
'told to avoid selecting the squares marked by the *.
   
If Square(1).Caption = strHuman And Square(5).Caption = strHuman Then
    SquareValue(6) = 0
End If
    
If Square(5).Caption = strHuman And Square(7).Caption = strHuman Then
    SquareValue(0) = 0
End If
    
If Square(3).Caption = strHuman And Square(7).Caption = strHuman Then
    SquareValue(2) = 0
End If
    
If Square(1).Caption = strHuman And Square(3).Caption = strHuman Then
    SquareValue(8) = 0
End If


'*************************************************************************
'Method 3 of 3:
'The last way of trying to win in 2 places at once is to
'take one diagonal and then a seemingly non-threatening side
'square at the opposite end as shown in the following:
'(with Human player going first and playing as X)
'
'            |X|
'           -----
'            |O|
'           -----
'           *| |X
'
'If the Computer takes the square marked * it will lose with the following:
'
'           *|X|X
'           -----
'            |O|*
'           -----
'           O| |X
'
'The solution to this is to check for the square that isn't in a
'corner (or the center square). Then we simply increase the value of
'the diagonal squares adjacent to the middle square.
'(*) denotes desirable squares to take.
'
'           *|X|*
'           -----
'            |O|
'           -----
'            | |X

'First we check to see if only one middle side square was taken.
'If so, we'll activate the value changes necessary to prevent
'the human player from winning with this strategy.
'This routine will only execute at the highest intelligence level:

If Intelligence.Value > 4 Then
    For Index = 1 To 7 Step 2
        If Square(Index).Caption = strHuman Then
            MiddleSideSquare = MiddleSideSquare + 1
        End If
    Next Index
    
    If MiddleSideSquare = 1 Then
        If Square(1).Caption = strHuman Then
            SquareValue(0) = 10
            SquareValue(2) = 10
        End If
        
        If Square(3).Caption = strHuman Then
            SquareValue(0) = 10
            SquareValue(6) = 10
        End If
        
        If Square(5).Caption = strHuman Then
            SquareValue(2) = 10
            SquareValue(8) = 10
        End If
        
        If Square(7).Caption = strHuman Then
            SquareValue(6) = 10
            SquareValue(8) = 10
        End If
    End If
End If

End Sub


Private Sub CheckForWin()
Dim Index As Integer

'Removes the checkboxes that allow human player to
'change sides or change who goes first after the first
'move is made - no changing sides mid-game!  :)
If HumanX.Visible = True Then Call RemoveOptions

'The rest of this procedure checks every possible direction one could
'win and triggers the end of game procedures if anyone gets 3 in a row.
CaseCounter = 0
While CaseCounter < 8
                
    Call BoardCheck
        XsInARow = 0
        OsInARow = 0
    
    For Index = StartSquare To EndSquare Step Direction
        If Square(Index).Caption = "X" Then XsInARow = XsInARow + 1
        If Square(Index).Caption = "O" Then OsInARow = OsInARow + 1
    Next Index
    
If XsInARow = 3 Then Call XWins
If OsInARow = 3 Then Call OWins
      
CaseCounter = CaseCounter + 1
Wend
End Sub

    

Private Sub RemoveOptions()
    HumanX.Visible = False
    HumanFirst.Visible = False
End Sub


'This section is dedicated to the computer's
'A.I. routine for determining the best move
'to make.
Private Sub CompAI()

Dim XOSquare As Integer
Dim Index As Integer
Dim Space As Integer

'Gives a value of 5 to the open diagonal squares and the center square
'and gives a value of 1 to the remaining open squares.
For Index = 0 To 8
    If Square(Index).Caption = "" Then
        If Index Mod 2 = 0 Then
            SquareValue(Index) = 5
                Else
            SquareValue(Index) = 1
        End If
    Else
'Changes any square that's been taken already to a value of zero.
        SquareValue(Index) = 0
    End If
Next Index

'     Our Special A.I. comes into play on the third move of the game.
'At this critical juncture, the computer needs to analyze the
'strategy that the human may be using to try to win. Each method
'is described in detail within the "SpecialAI" procedure. Note that
'the SpecialAI is only called if the intelligence value is greater
'than 3.
For Index = 0 To 8
    If Square(Index).Caption = "" Then Space = Space + 1
Next Index

If Space = 6 And Intelligence.Value > 3 Then
    Call SpecialAI
End If

CaseCounter = 0
XOSquare = 1

'Center square is given a better value than other squares
'since it allows for a possibility of winning for every direction.
'The lowest intelligence level ignores the rest of the statements
'down to the end of the procedure. Also, note the use of tabbing
'to keep track of where our IF - END IF blocks begin and end.
'At its deepest, there are 5 nested IF - END IF blocks. Its really
'easy to get lost if you don't use tabbing to distinguish how deep
'your logic tree goes!
If Intelligence.Value > 1 Then
    If Square(4).Caption = "" Then
        SquareValue(4) = 10
    End If

'At each itiration, this While/Wend loop checks each row, column and diagonal.
    While CaseCounter < 8


'Calls our case checking procedure. This procedure helps us to examine
'each row, column, and diagonal individually. An open square that would
'block the human's attempt to win is given a value of 500. Of course, if
'the computer can immediately win by taking an open square, that square
'will have the highest value of all - 999!
    Call BoardCheck
    CaseCounter = CaseCounter + 1
        
        For XOSquare = 1 To 6
        Call XOSelect(XOSquare)

            If Square(StartSquare).Caption = XO(0) And _
            Square(StartSquare + Direction).Caption = XO(1) And _
            Square(EndSquare).Caption = XO(2) Then

'Notice the use of the intelligence values here:
'A value of 3 or higher will give the computer the ability to
'discern where to block the human opponent. A value of 2 or
'greater gives the computer the ability to dicern where to
'take a square that will immediately produce a win.
'These following lines are a little hard to follow so I'll break it down:
'Lets take the following block -

'         1       If HumanPlayingAsX = True Then
'         2           If XOSquare <= 3 Then
'         3               If Intelligence.Value > 2 And SquareValue(Blank) < 999 Then
'         4                   SquareValue(Blank) = 500
'         5               End If
'         6           Else

'Line 1: We begin our search only if the human player plays as X.
'Line 2: We isolate the section of our "XOSquare" SELECT CASE procedure
'that we want to look at. The first 3 choices in that procedure allow
'us to check for a winning possibility for X. Since line 1 told us
'that the human player is playing as X, line 2 allows us to study whether
'or not the human is threatening a win.
'Line 3: Two conditions must be true to move on from here - The artificial
'intelligence must be high enough, and the value of the open square we
'are looking at to block must be less than 999. The reason why we give
'the value 999 to a square is because it is of more benefit to take a
'square that could immediately win than a square that merely produces
'a block. If we didn't have the "...And SquareValue(Blank) < 999 Then"
'part of this statement, the computer could potentially overwrite a
'lower value (like 500) with the higher and better value of 999. This
'would happen if the open square could potentially produce a win and
'at the same time block the human's attempt to win. Bottom line: Always
'make sure the best move (represented by the highest SquareValue() number
'is preserved over any lower value.
'Line 4: Assigns a value of 500 to the blank square which would block
'the human's attempt to win.
'Line 5: closes the IF - END IF clause started in line 3
'Line 6: Else refers to line 2. Since line 2 checked for X's, this ELSE
'clause moves on to check the O's for the same thing. The only difference
'here is that since we're checking O's now and the human is playing X, the
'thing the rest of the lines after line 6 check for is different. The rest
'of the lines check for a chance for the computer to immediately win!

                If HumanPlayingAsX = True Then
                    If XOSquare <= 3 Then
                        If Intelligence.Value > 2 And SquareValue(Blank) < 999 Then
                            SquareValue(Blank) = 500
                        End If
                    Else
                        If Intelligence.Value > 1 Then
                            SquareValue(Blank) = 999
                        End If
                    End If
                    
'Checking for same thing with human playing as O.
                Else
                    If XOSquare <= 3 Then
                        If Intelligence.Value > 1 Then
                            SquareValue(Blank) = 999
                        End If
                    Else
                        If Intelligence.Value > 2 And SquareValue(Blank) < 999 Then
                            SquareValue(Blank) = 500
                        End If
                    End If
                End If
            End If
        Next XOSquare
    Wend
End If

'This code will remain in commented mode unless you delete the ' symbols.
'Deleting the comment symbols will allow you to view the square values given
'in the "immediate" window. This was a useful debugging tool for me throughout
'the creation of this program, and it can give you insight as to how the
'artificial intelligence works.

'*************************************************************************
'Debug.Print
'For Index = 0 To 8
'If Index Mod 3 = 0 Then Debug.Print
'Debug.Print SquareValue(Index); " ";
'If SquareValue(Index) < 100 Then Debug.Print " ";
'If SquareValue(Index) < 10 Then Debug.Print " ";
'Next Index
'*************************************************************************
End Sub


Private Sub CompSquareSelection()
Dim Index As Integer
Dim HighVal As Integer
Dim Counter As Integer
Dim FinalChoice As Integer

Counter = 0
HighVal = 0
For Index = 0 To 8
If SquareValue(Index) = HighVal Then
    RandChoice(Counter) = Index
    Counter = Counter + 1
End If

If SquareValue(Index) > HighVal Then
    HighVal = SquareValue(Index)
    For Counter = 1 To 8
    RandChoice(Counter) = 0
    Next Counter
    RandChoice(0) = Index
    Counter = 1
End If
Next Index

'   The check for whether HighVal > 0 is to assure that the computer
'doesn't try to move to an already occupied square when
'the board is full.
'   The rest of the code selects the move the computer ultimately makes.
'If the value of the best choice is equal to the values of other squares,
'this code will choose a square at random from among those choices.

Index = Int(Rnd * Counter)
Square(RandChoice(Index)).Caption = strCPU
HumanMovedAlready = False
End Sub


Private Sub ComputerTurn()

If GameOver = False Then
    Call CompAI
    Call CompSquareSelection
    Call CheckForWin
    Call CheckForCatsGame
End If

End Sub
