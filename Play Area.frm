VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Joe Logic"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   240
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   605
      TabIndex        =   0
      Top             =   120
      Width           =   9075
      Begin VB.Image Image1 
         Height          =   450
         Index           =   11
         Left            =   5160
         Picture         =   "Play Area.frx":0000
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   10
         Left            =   4680
         Picture         =   "Play Area.frx":0802
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   9
         Left            =   4200
         Picture         =   "Play Area.frx":138C
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   8
         Left            =   3720
         Picture         =   "Play Area.frx":1B8E
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   7
         Left            =   3240
         Picture         =   "Play Area.frx":2390
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   6
         Left            =   2760
         Picture         =   "Play Area.frx":2F1A
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   5
         Left            =   2280
         Picture         =   "Play Area.frx":3AA4
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   4
         Left            =   1800
         Picture         =   "Play Area.frx":4256
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   3
         Left            =   1320
         Picture         =   "Play Area.frx":4A58
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   2
         Left            =   840
         Picture         =   "Play Area.frx":5562
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   1
         Left            =   120
         Picture         =   "Play Area.frx":5D64
         Top             =   120
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgExit 
         Height          =   450
         Left            =   2640
         Picture         =   "Play Area.frx":6566
         Top             =   2760
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgTransport 
         Height          =   450
         Index           =   2
         Left            =   2520
         Picture         =   "Play Area.frx":70F0
         Top             =   1440
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgTransport 
         Height          =   450
         Index           =   1
         Left            =   3480
         Picture         =   "Play Area.frx":7C7A
         Top             =   1440
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenLastSave 
         Caption         =   "Open Last Save -------------- [singlelevel.tmp]"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSpace1 
         Caption         =   ""
      End
      Begin VB.Menu mnuFileSaveScreen 
         Caption         =   "Save Screen ------------------ [singlelevel.tmp], also appends to [levels.tmp]"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveMoves 
         Caption         =   "Save Moves ------------------ [moves.tmp]"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFileAppendLevel 
         Caption         =   "Append Saved Screen --- appends [singlelevel.tmp] to [levels.joe]"
      End
      Begin VB.Menu mnuFileAppendMoves 
         Caption         =   "Append Saved Moves ---- appends [moves.tmp] to [levels.joe]"
      End
      Begin VB.Menu mnuFileSpace2 
         Caption         =   ""
      End
      Begin VB.Menu mnuFileRestartLevel 
         Caption         =   "Restart Level ------------------ (the level in memory)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileSpace3 
         Caption         =   ""
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditEditMode 
         Caption         =   "Edit Mode"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "Level"
      Begin VB.Menu mnuLevelUp 
         Caption         =   "Level Up"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuLevelDown 
         Caption         =   "Level Down"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuSolution 
      Caption         =   "Solution"
      Begin VB.Menu mnuSolutionSpeed0 
         Caption         =   "Fast"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSolutionSpeed1 
         Caption         =   "Medium"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSolutionSpeed2 
         Caption         =   "Slow"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSolutionStart 
         Caption         =   "Solution Start"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuSolutionStop 
         Caption         =   "Solution Stop"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim x As Integer, y As Integer
Form1.Show
ChDir App.Path

'Set Your Object Pictures
Set picCell(1) = Image1(1)
Set picCell(2) = Image1(2)
Set picCell(3) = Image1(3)
Set picCell(4) = Image1(4)
Set picCell(5) = Image1(5)
Set picCell(6) = Image1(6)
Set picCell(7) = Image1(7)
Set picCell(8) = Image1(8)
Set picCell(9) = Image1(9)
Set picCell(10) = Image1(10)
Set picCell(11) = Image1(11)

'Pixel Log
'Log the x, y pixel corner of each and every block
For y = 1 To 15
For x = 1 To 20
pxlColumn(x, y) = (x - 1) * 30
pxlRow(x, y) = (y - 1) * 30
'DoEvents
Next x, y


stFileName = "levels.joe" 'Main Play File
sngSolutionDelay = 0 'Set Default Solution Speed
mnuSolutionStart.Enabled = True
mnuSolutionStop.Enabled = False
mnuSolutionSpeed0.Checked = True 'Fast Solution Speed

'Set Level
intLevel = 1 'Initialize Level
rtnInitialize
End Sub

Private Sub mnuFileAppendLevel_Click()
Dim tempTextArray() As String, tempText As String, stChar As String
Dim intChar As Integer, intLine As Integer, tempCount As Integer
Dim levelLocation As Integer
Dim booLevelfound As Boolean

'This procedure will append singlelevel.tmp to the main levels.joe

'***************************************
'***Tell user what this routine does****
'***and prompt for level number*********
'***************************************
Dim intResponse As Integer
Dim strMessage As String
strMessage = "Are you sure?" + Chr(13)
strMessage = strMessage + "Was your Last 'Save Screen' the way you want it?" + Chr(13)
strMessage = strMessage + "This will append your last 'Save Screen' (singlelevel.tmp) to the main (levels.joe)" + Chr(13)
intResponse = MsgBox(strMessage, vbOKCancel, "Append Last Save To The Main (levels.joe)")
If intResponse = vbCancel Then Exit Sub
strMessage = InputBox("Enter the desired level number. No levels will be overwritten, only appended.", "Level Number")
If strMessage = "" Then Exit Sub
'***************************************
'**End Tell user what this routine does*
'***************************************


'***************************************
'*Count number of entries in levels.joe*
'********for redimming array************
'***************************************
tempCount = 0
    Open "levels.joe" For Input As #1
    Do While Not EOF(1)
    tempCount = tempCount + 1
    Line Input #1, tempText
    Loop
    Close #1
ReDim tempTextArray(tempCount)
'***************************************
'*****End Count number of entries*******
'***************************************
    
    
'***************************************
'*****Load a temporary array with*******
'******the contents of levels.joe*******
'***************************************
    Open "levels.joe" For Input As #1
    For intLine = 1 To tempCount
    Line Input #1, tempTextArray(intLine)
    Next intLine
    Close #1
'***************************************
'****End load a temporary array with****
'***************************************
    
'***************************************
'****Now output the loaded array array**
'****to levels.joe and then append******
'***********singlelevel.tmp*************
'***************************************
    Open "levels.joe" For Output As #1
    Open "singlelevel.tmp" For Input As #2
    Line Input #2, tempText 'Skip line with "level1" on it
    
    For intLine = 1 To tempCount
    Print #1, tempTextArray(intLine)
    Next intLine
    Print #1,
    Print #1, "level" & strMessage
    
    Do While Not EOF(2)
    Line Input #2, tempText
    Print #1, tempText
    Loop
        Close
'***************************************
'***End output the loaded array array***
'***************************************

End Sub

Private Sub mnuFileAppendMoves_Click()
Dim tempTextArray() As String, tempText As String, stChar As String
Dim intChar As Integer, intLine As Integer, tempCount As Integer
Dim levelLocation As Integer
Dim booLevelfound As Boolean

'This procedure will append moves.tmp to the main levels.joe

'***************************************
'***Tell user what this routine does****
'***and prompt for level number*********
'***************************************
Dim intResponse As Integer
Dim strMessage As String
strMessage = "Are you sure?" + Chr(13)
strMessage = "Was your Last Moves Save a Complete and Accurate Solution?" + Chr(13)
strMessage = strMessage + "This will append your last moves save (moves.tmp) to the main (levels.joe)" + Chr(13)
intResponse = MsgBox(strMessage, vbOKCancel, "Append Last Moves Save (moves.tmp) To The Main (levels.joe)")
If intResponse = vbCancel Then Exit Sub
strMessage = InputBox("Enter the desired level number for these moves. No moves will be overwritten, only appended.", "Solution/Moves Number")
If strMessage = "" Then Exit Sub
'***************************************
'**End Tell user what this routine does*
'***************************************


'***************************************
'*Count number of entries in levels.joe*
'********for redimming array************
'***************************************
tempCount = 0
    Open "levels.joe" For Input As #1
    Do While Not EOF(1)
    tempCount = tempCount + 1
    Line Input #1, tempText
    Loop
    Close #1
ReDim tempTextArray(tempCount)
'***************************************
'*****End Count number of entries*******
'***************************************
    
    
'***************************************
'*****Load a temporary array with*******
'******the contents of levels.joe*******
'***************************************
    Open "levels.joe" For Input As #1
    For intLine = 1 To tempCount
    Line Input #1, tempTextArray(intLine)
    Next intLine
    Close #1
'***************************************
'****End load a temporary array with****
'***************************************
    
'***************************************
'****Now output the loaded array array**
'****to levels.joe and then append******
'***********moves.tmp*************
'***************************************
    Open "levels.joe" For Output As #1
    Open "moves.tmp" For Input As #2
    
    For intLine = 1 To tempCount
    Print #1, tempTextArray(intLine)
    Next intLine
    Print #1,
    Print #1, "moves" & strMessage
    
    Do While Not EOF(2)
    Line Input #2, tempText
    Print #1, tempText
    Loop
        Close
'***************************************
'***End output the loaded array array***
'***************************************

End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuEditEditMode_Click()
mnuEditEditMode.Checked = Not mnuEditEditMode.Checked
    If mnuEditEditMode.Checked = True Then
    mnuFile.Enabled = False
    mnuLevel.Enabled = False
    mnuSolution.Enabled = False
    Else
    mnuFile.Enabled = True
    mnuLevel.Enabled = True
    mnuSolution.Enabled = True
    End If
EditObject = 1 'cEmpty
rtnRefreshFromArray
rtnAllRollsAndBoxesFall
End Sub

Private Sub mnuFileOpen_Click()
  'Common Dialog Window
  '***************************************************************************
  CommonDialog1.CancelError = True  'Enable on error or cancel GoTo
    On Error GoTo cancelPressed
  CommonDialog1.Flags = cdlOFNHideReadOnly  'disable read only chk box
  CommonDialog1.DialogTitle = "open"        'Title displayed
  CommonDialog1.InitDir = ""                'Start Directory
       
  'Format     object.Filter [= description1 |filter1 |description2 |filter2...]
  'Example    Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico
  CommonDialog1.Filter = "levels (*.joe)|*.joe"

  CommonDialog1.FileName = "levels.joe"
  CommonDialog1.ShowOpen
  stFileName = CommonDialog1.FileName
  '***************************************************************************
intLevel = 1
rtnInitialize
cancelPressed:
Picture1.SetFocus
End Sub

Private Sub mnuFileOpenLastSave_Click()
stFileName = "singlelevel.tmp"
intLevel = 1
rtnInitialize
Exit Sub
End Sub

Private Sub mnuFileRestartLevel_Click()
rtnInitialize
End Sub

Private Sub mnuFileSaveScreen_Click()
Dim tempTextArray() As String, tempText As String, stChar As String
Dim intChar As Integer, intLine As Integer, tempCount As Integer
Dim levelLocation As Integer
Dim booLevelfound As Boolean

'This procedure will save your current screen just as you see it to
'text file singlelevel.tmp and also append it to text file levels.tmp
'These are temporary files used as a work area for creating new levels

'***************************************
'***Tell user what this routine does****
'***************************************
Dim intResponse As Integer
Dim strMessage As String
strMessage = "This will save the screen just as you see it to" + Chr(13)
strMessage = strMessage + "text file (singlelevel.tmp) and also append it to text file (levels.tmp)" + Chr(13)
strMessage = strMessage + "These are temporary files used as a work area for creating new levels" + Chr(13)
intResponse = MsgBox(strMessage, vbOKCancel, "Save and Append Screen to Temporary Files")
If intResponse = vbCancel Then Exit Sub
'***************************************
'**End Tell user what this routine does*
'***************************************


'***************************************
'*Count number of entries in levels.tmp*
'********for redimming array************
'***************************************
tempCount = 0
    Open "levels.tmp" For Input As #1
    Do While Not EOF(1)
    tempCount = tempCount + 1
    Line Input #1, tempText
    Loop
    Close #1
ReDim tempTextArray(tempCount)
'***************************************
'*****End Count number of entries*******
'***************************************
    
    
'***************************************
'*****Load a temporary array with*******
'******the contents of levels.tmp*******
'***************************************
    Open "levels.tmp" For Input As #1
    For intLine = 1 To tempCount
    Line Input #1, tempTextArray(intLine)
    Next intLine
    Close #1
'***************************************
'****End load a temporary array with****
'***************************************
    
'***************************************
'****Now output the loaded array array**
'****to singlelevel.tmp and append it to******
'***********to levels.tmp***************
'***************************************
    Open "levels.tmp" For Output As #1
    Open "singlelevel.tmp" For Output As #2
    Print #1,
    Print #2, "level1"
    
    For intLine = 1 To tempCount
    Print #1, tempTextArray(intLine)
    Next intLine
        
        For intLine = 1 To 15       'y cells
            For intChar = 1 To 20   'x cells
            Print #1, CellCharactr(intChar, intLine);
            Print #2, CellCharactr(intChar, intLine);
            Next intChar
            Print #1,
            Print #2,
        Next intLine
        Close
'***************************************
'***End output the loaded array array***
'***************************************
End Sub

Private Sub mnuFileSaveMoves_Click()
Dim x As Integer, y As Integer, tmpCount As Integer
Picture1.Enabled = False

'***************************************
'***Tell user what this routine does****
'***************************************
Dim intResponse As Integer
Dim strMessage As String
strMessage = "This will save all your moves on this level to temporary file (moves.tmp)"
intResponse = MsgBox(strMessage, vbOKCancel, "Save Moves to Temporary File")
    If intResponse = vbCancel Then
    Picture1.Enabled = True
    Exit Sub
    End If
'***************************************
'**End Tell user what this routine does*
'***************************************

'***************************************
'****Save all moves on current level****
'*********to (moves.tmp)****************
'***************************************
Open "moves.tmp" For Output As #1
    tmpCount = 0
    For y = 1 To 50
        For x = 1 To 200
        tmpCount = tmpCount + 1
        Print #1, Trim(Str(byMoves(tmpCount)));
            If byMoves(tmpCount) = 0 Then
            Print #1,
            Close
            Picture1.Enabled = True
            Exit Sub
            End If
        Next x
    Print #1,
    Next y
    Close
'***************************************
'**End Save all moves on current level**
'***************************************
    Picture1.Enabled = True
End Sub

Private Sub mnuLevelUp_Click()
Dim tempText As String
Dim booFound As Boolean
    'Go Up One Level
    intLevel = intLevel + 1
    
    'Reload level first
    Open stFileName For Input As #1
        Do While Not EOF(1)
        Line Input #1, tempText
            'Check for desired level
            If tempText = "level" & Trim(Str(intLevel)) Then
            booFound = True
            Exit Do
            End If
        Loop
        Close

If booFound = True Then
'Level found so load it
rtnInitialize
Exit Sub

Else
'There wasn't a higher level so load the current one
intLevel = intLevel - 1
rtnInitialize
Exit Sub
End If

End Sub
Private Sub mnuLevelDown_Click()
    'Go Down One Level
intLevel = intLevel - 1

If intLevel > 0 Then
'Now load it
rtnInitialize
Exit Sub

Else
'Already on lowest level so reload it
intLevel = intLevel + 1
rtnInitialize
Exit Sub
End If

End Sub

Private Sub mnuSolutionSpeed0_Click()
'Set Solution Speed
sngSolutionDelay = 0
mnuSolutionSpeed0.Checked = True
mnuSolutionSpeed1.Checked = False
mnuSolutionSpeed2.Checked = False
End Sub

Private Sub mnuSolutionSpeed1_Click()
'Set Solution Speed
sngSolutionDelay = 0.1
mnuSolutionSpeed0.Checked = False
mnuSolutionSpeed1.Checked = True
mnuSolutionSpeed2.Checked = False
End Sub
Private Sub mnuSolutionSpeed2_Click()
'Set Solution Speed
sngSolutionDelay = 0.5
mnuSolutionSpeed0.Checked = False
mnuSolutionSpeed1.Checked = False
mnuSolutionSpeed2.Checked = True
End Sub

Private Sub mnuSolutionStart_Click()
'Display Solution
Dim stLine As String, stChar As String
Dim x As Integer, y As Integer
Dim booSolutionFound As Boolean
Dim sngTimerAppend As Single

'Enable and disable menu to avoid conflicts
Picture1.Enabled = False
mnuFile.Enabled = False
mnuEdit.Enabled = False
mnuEdit.Enabled = False
mnuSolutionStart.Enabled = False
mnuSolutionStop.Enabled = True
mnuLevel.Enabled = False

'First reload level
rtnInitialize

'Now load the main file "levels.joe"
Open stFileName For Input As #1
    booSolutionFound = False 'reset
    
    'Find out if in fact there is a solution to the requested level
    Do While Not EOF(1)
    Input #1, stLine
    If stLine = "moves" & Trim(Str(intLevel)) Then booSolutionFound = True
    Loop
    Close
    
    'Exit here if no solution was included
    If booSolutionFound = False Then
    Close
    GoTo lblExitSub:
    End If


'Ahh! Didn't exit so solution for rquested level was found
'Lets process now.
'Open "levels.joe again"
Open stFileName For Input As #1
    'Find the correct solution. (moves)
    Do While Not EOF(1)
    Input #1, stLine
    If stLine = "moves" & Trim(Str(intLevel)) Then Exit Do
    Loop
    
    'Now the current line is the label. "moves1-?"
    'the next line/lines will contain the solution
    'So let's process
    For y = 1 To 50
    Input #1, stLine
        For x = 1 To 200
        stChar = Mid$(stLine, x, 1) 'Get a solution line
     
            'if we find a zero that our que to exit, end of solution
            If stChar = "0" Then
            Close
            mnuSolutionStart.Enabled = True
            GoTo lblExitSub:
            End If
    
                'Find out which key/direction was used and goto respective procedure.
                'If little joe is carrying a block then send him to respective procedure also.
                Select Case stChar
                Case "1"
                If booWithBlock = False Then
                rtnLeftKey
                Else
                rtnLeftKeyWithBox
                End If
    
                Case "2"
                If booWithBlock = False Then
                rtnDownKey
                Else
                rtnDownKeyWithBox
                End If
    
                Case "3"
                If booWithBlock = False Then
                rtnRightKey
                Else
                rtnRightKeyWithBox
                End If
    
                Case "5"
                If booWithBlock = False Then
                rtnUpKey
                Else
                rtnUpKeyWithBox
                End If
                End Select


'Here is our delay according to user selected speed on solution menu
sngTimerAppend = Timer
Do While Timer < sngTimerAppend + sngSolutionDelay
DoEvents
Loop

Picture1.Refresh
DoEvents
    'If user selects stop solution then exit
    If mnuSolutionStart.Enabled = True Then
    Close
    GoTo lblExitSub:
    End If
Next x
Next y
     Close


'Reset all before leaving
lblExitSub:
mnuSolutionStart.Enabled = True
Picture1.Enabled = True
mnuFile.Enabled = True
mnuEdit.Enabled = True
mnuEdit.Enabled = True
mnuSolutionStart.Enabled = True
mnuSolutionStop.Enabled = False
mnuLevel.Enabled = True
Exit Sub


End Sub

Private Sub mnuSolutionStop_Click()
'Stop Solution
mnuSolutionStart.Enabled = True
mnuSolutionStop.Enabled = False
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
'Heeyyy! A keypress! Lets follow orders and process this keypress!
Dim stChar As String
Dim dblTimerStart As Double
If mnuEditEditMode.Checked = True Then Exit Sub 'Exit if in edit mode
If booBusyFallingOrRolling = True Then Exit Sub 'Exit if busy falling or ball rolling
Picture1.Refresh


'Find out which key/direction was used and goto respective procedure.
'If little joe is carrying a block then send him to respective procedure also.
stChar = Chr(KeyAscii)       ' convert keypress to string and assign to char
Select Case stChar
    
    Case "1"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Save Moves
    If booWithBlock = False Then
    rtnLeftKey
    Exit Sub
    Else
    rtnLeftKeyWithBox
    Exit Sub
    End If
    
    Case "2"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Save Moves
    If booWithBlock = False Then
    rtnDownKey
    Exit Sub
    Else
    rtnDownKeyWithBox
    Exit Sub
    End If

    Case "3"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Save Moves
    If booWithBlock = False Then
    rtnRightKey
    Exit Sub
    Else
    rtnRightKeyWithBox
    Exit Sub
    End If

    Case "5"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Save Moves
    If booWithBlock = False Then
    rtnUpKey
    Exit Sub
    Else
    rtnUpKeyWithBox
    Exit Sub
    End If

End Select

Exit Sub

lblStoreMove:
'Increment move count and store move in case of a Menu Save Moves
intMoveCount = intMoveCount + 1
byMoves(intMoveCount) = Val(stChar)
Return

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Must be in edit mode
'This procedure changes:
'Current object (joe, block, roll, etc) with right button and
'Stores it with left button


If mnuEditEditMode.Checked = False Then Exit Sub 'Not on edit mode. Let's Exit

Dim NewObject As Integer
Dim PreviousLetter As String, NewLetter As String

'Right button was pressed so next object
If Button = 2 Then
EditObject = EditObject + 1 'Next object
  If EditObject > cTotalObjects Then 'Start over
  EditObject = 1
  End If
End If
    
   'Left button was pressed so update arrays etc.
    If Button = 1 Then
       PreviousLetter = CellCharactr(EditX, EditY) 'Letter designation for object previously in the location we wish to change
       NewObject = EditObject 'Temporary hold new object
                    
                    'Contents requested to be change
    Select Case PreviousLetter
       
       Case "T", "t", "J", "j", "X" 'Transport1,Transport2,JoeRight,JoeLeft,Exit
           'Do Nothing
           'You can not write over these.
           'These can only be changed by placing them elsewhere.
       
       Case " ", "B", "O", "=", "#", "." 'Empty,Box,Ball,Rail,Brick,RailWalkOnce
           'These can be changed so lets update arrays
          
          Select Case NewObject
            Case cEmpty
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = " " 'Update new location letter
            Case cBox
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "B" 'Update new location letter
            Case cRoll
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "O" 'Update new location letter
            Case cRail
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "=" 'Update new location letter
            Case cBrick
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "#" 'Update new location letter
            Case cRailWalkOnce
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "." 'Update new location letter
            Case cTransport1
                  'Dummy numbers (100) are initial default values pointing
                  'to an impossible location. They
                  'indicate whether or not transports were
                  'already in use so as to update there previous location.
                  If Transport1X <> 100 And Transport1Y <> 100 Then
                  CellCharactr(Transport1X, Transport1Y) = " " 'Update previous location letter
                  CellContents(Transport1X, Transport1Y) = cEmpty 'Update previous location object
                  End If
              CellContents(EditX, EditY) = cEmpty 'Update new location object
              CellCharactr(EditX, EditY) = "T" 'Update new location letter
              Transport1X = EditX 'Update Transports x matrix value
              Transport1Y = EditY 'Update Transports y matrix value
              Form1.imgTransport(1).Visible = True 'Make Transport visible
              Form1.imgTransport(1).Left = pxlColumn(EditX, EditY) 'Place in proper location
              Form1.imgTransport(1).Top = pxlRow(EditX, EditY) 'Place in proper location
            
            Case cTransport2
                  'Dummy numbers (100) are initial default values pointing
                  'to an impossible location. They
                  'indicate whether or not transports were
                  'already in use so as to update there previous location.
              If Transport2X <> 100 And Transport2Y <> 100 Then
              CellCharactr(Transport2X, Transport2Y) = " " 'Update previous location letter
              CellContents(Transport2X, Transport2Y) = cEmpty 'Update previous location object
              End If
              CellContents(EditX, EditY) = cEmpty 'Update new location object
              CellCharactr(EditX, EditY) = "t" 'Update new location letter
              Transport2X = EditX 'Update Transports x matrix value
              Transport2Y = EditY 'Update Transports y matrix value
              Form1.imgTransport(2).Visible = True 'Make Transport visible
              Form1.imgTransport(2).Left = pxlColumn(EditX, EditY) 'Place in proper location
              Form1.imgTransport(2).Top = pxlRow(EditX, EditY) 'Place in proper location
            
            Case cExit
                  'Dummy numbers (100) are initial default values pointing
                  'to an impossible location. They
                  'indicate whether or not an exit was
                  'already in play area so as to update its previous location.
                  If ExitColumn <> 100 And ExitRow <> 100 Then
                  CellCharactr(ExitColumn, ExitRow) = " " 'Update previous location letter
                  CellContents(ExitColumn, ExitRow) = cEmpty 'Update previous location object
                  End If
              CellContents(EditX, EditY) = cEmpty 'Update new location object
              CellCharactr(EditX, EditY) = "X" 'Update new location letter
              ExitColumn = EditX 'Update Exit x matrix value
              ExitRow = EditY 'Update Exit y matrix value
              Form1.imgExit.Visible = True 'Make Exit visible
              Form1.imgExit.Left = pxlColumn(EditX, EditY) 'Place in proper location
              Form1.imgExit.Top = pxlRow(EditX, EditY) 'Place in proper location
            
            Case cJoeRight
              CellContents(JoeX, JoeY) = cEmpty 'Update previous location object
              CellCharactr(JoeX, JoeY) = " " 'Update previous location letter
              CellContents(EditX, EditY) = cJoeRight 'Update new location object
              CellCharactr(EditX, EditY) = "J" 'Update new location letter
              JoeX = EditX 'Update Joe's x matrix value
              JoeY = EditY 'Update Joe's y matrix value
            
            Case cJoeLeft
              CellContents(JoeX, JoeY) = cEmpty 'Update previous location object
              CellCharactr(JoeX, JoeY) = " " 'Update previous location letter
              CellContents(EditX, EditY) = cJoeLeft 'Update new location object
              CellCharactr(EditX, EditY) = "j" 'Update new location letter
              JoeX = EditX 'Update Joe's x matrix value
              JoeY = EditY 'Update Joe's y matrix value
          
          End Select
    End Select
End If
 '       empty = " " 1
'         box = "B" 2
'        roll = "O" 3
'        rail = "=" 4
'       brick = "#" 5
'   transport = "T" 6
'   transport = "t" 7
'     joeleft = "j" 8
'    joeright = "J" 9
'        exit = "X" 10
'railwalkonce = "." 11
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Must be in edit mode
'This procedure paints and clears an object (joe, block, roll, etc)
'as you move the mouse across the screen.
'It paints object in screen cell block matrix of 20 x 15


If mnuEditEditMode.Checked = False Then Exit Sub 'Not on edit mode. Let's Exit



Dim xx As Integer, yy As Integer
'Set these variables to correspond with the mouse arrow
'and picture plascement
x = x + 15
y = y + 15
x = (x / 30): y = (y / 30)

'Just making sure you are in the play area
xx = x: yy = y
If xx > 20 Then xx = 20
If xx < 1 Then xx = 1
If yy > 15 Then yy = 15
If yy < 1 Then yy = 1
    
    'Repaint only if on another cell block to prevent redundancy
    If xx <> EditX Or yy <> EditY Then
    rtnRefreshFromArray
    Picture1.PaintPicture picCell(EditObject), pxlColumn(xx, yy), pxlRow(xx, yy)
    End If

'Update cell x, y pointers,  20 x 15
EditX = xx: EditY = yy
End Sub







