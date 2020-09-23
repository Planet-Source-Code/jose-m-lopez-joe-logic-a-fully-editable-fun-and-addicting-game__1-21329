Attribute VB_Name = "Module1"
Option Explicit

Public picCell(11) As Picture 'Picture array

Public CellContents(20, 15) As Integer
Public CellCharactr(20, 15) As String
Public pxlColumn(20, 15) As Integer 'Stores Cell corner position
Public pxlRow(20, 15) As Integer 'Stores Cell corner position
Public byMoves(10000) As Byte 'Stores Moves
Public intMoveCount As Integer 'Holds Move Count
Public intLevel As Integer 'Holds current level number
Public JoeX As Integer  'Store Joe's x cell position
Public JoeY As Integer 'Store Joe's y cell position
Public TempColumn As Integer  'Store Temporary x cell position for falling objects
Public TempRow As Integer 'Store Temporary y cell position for falling objects
Public TempObject As Integer  'Holds Temporarily a falling object's value
Public TempCharacter As String   'Holds Temporarily a falling object's Letter

Public EditObject As Integer  'Holds an object's value in Edit Mode
Public EditX As Integer  'Holds an object's value in Edit Mode
Public EditY As Integer  'Holds an object's value in Edit Mode
Public stFileName As String 'Hold current open filename
Public sngSolutionDelay As Double 'Solutions delay value according to solution menu

Public booWithBlock As Boolean 'Whether or not little Joe is carrying a block

'This booBusyFallingOrRolling is needed because without it when Joe is falling down
'an appreciable distance, he can also travel right or left if the respective key
'is held down. This variable prevents this.
Public booBusyFallingOrRolling As Boolean

'*********************************************************
'********Picture Name Constants***************************
'*********************************************************
Public Const cEmpty As Integer = 1
Public Const cBox As Integer = 2
Public Const cRoll  As Integer = 3
Public Const cRail  As Integer = 4
Public Const cBrick  As Integer = 5
Public Const cTransport1  As Integer = 6
Public Const cTransport2  As Integer = 7
Public Const cJoeLeft  As Integer = 8
Public Const cJoeRight  As Integer = 9
Public Const cExit  As Integer = 10
Public Const cRailWalkOnce  As Integer = 11
Public Const cTotalObjects  As Integer = 11
'*********************************************************
'*********************************************************
'*********************************************************


'*********************************************************
'**********Special Location Values************************
'*********************************************************
Public Transport1X   As Integer
Public Transport2X   As Integer
Public Transport1Y   As Integer
Public Transport2Y   As Integer
Public ExitColumn As Integer
Public ExitRow As Integer
'*********************************************************
'*********************************************************
'*********************************************************

Public Sub rtnInitialize()
'Get the text level file and decode it into the CellContents array

Form1.Enabled = False 'Prevent unwanted events while busy
Form1.Caption = "Joe Logic Level " & intLevel 'Title

rtnLevelToArray
Form1.Enabled = False 'Prevent unwanted events while busy

rtnRefreshFromArray
Form1.Enabled = False 'Prevent unwanted events while busy

rtnAllRollsAndBoxesFall
Form1.Enabled = False 'Prevent unwanted events while busy
    
    'Reset moves array
    For intMoveCount = 1 To 10000
    byMoves(intMoveCount) = 0
    Next intMoveCount
intMoveCount = 0 'Reset Count
booWithBlock = False
Form1.Enabled = True
End Sub


Public Sub rtnLevelToArray()
Dim tempText As String, stChar As String
Dim intLine As Integer, intChar As Integer
Dim booFound As Boolean
booFound = False

'Turn off transports
Form1.imgTransport(1).Visible = False
Form1.imgTransport(2).Visible = False
Form1.imgExit.Visible = False

'Impossible values indicating these objects are currently non existent
Transport1X = 100
Transport2X = 100
Transport1Y = 100
Transport2Y = 100
ExitColumn = 100
ExitRow = 100
    
'*****************************************************
'***Determine if next level is available**************
'*****************************************************
    'Look for next lavel
Open stFileName For Input As #1
    Do While Not EOF(1)
     Line Input #1, tempText
        If tempText = "level" & Trim(Str(intLevel)) Then
        booFound = True
        Exit Do
        End If
     Loop

'If not available then exit
If booFound = False Then
Close #1
GoTo lblNoMoreLevels:
End If
'*****************************************************
'*****************************************************
'*****************************************************
        
        
        
'*****************************************************
'***********Read Level Line By Line*******************
'********And GosubTo Processing Routine***************
'*****************************************************
        For intLine = 1 To 15       'y cells
        Line Input #1, tempText
            For intChar = 1 To 20   'x cells
            stChar = Mid$(tempText, intChar, 1)
            GoSub lblGetCellContents:
            Next intChar
        Next intLine
        Close #1

booWithBlock = False
Exit Sub
'*****************************************************
'*****************************************************
'*****************************************************
'*****************************************************


'*****************************************************
'**************Gosub Routine**************************
'**************UpDate Arrays**************************
'*****************************************************
lblGetCellContents:
Select Case stChar
    Case " "
        CellContents(intChar, intLine) = cEmpty
        CellCharactr(intChar, intLine) = stChar
    Case "B"
        CellContents(intChar, intLine) = cBox
        CellCharactr(intChar, intLine) = stChar
    Case "O"
        CellContents(intChar, intLine) = cRoll
        CellCharactr(intChar, intLine) = stChar
    Case "="
        CellContents(intChar, intLine) = cRail
        CellCharactr(intChar, intLine) = stChar
    Case "#"
        CellContents(intChar, intLine) = cBrick
        CellCharactr(intChar, intLine) = stChar
    Case "T"
        CellContents(intChar, intLine) = cEmpty
        CellCharactr(intChar, intLine) = stChar
        Transport1X = intChar
        Transport1Y = intLine
        Form1.imgTransport(1).Visible = True
        Form1.imgTransport(1).Left = pxlColumn(intChar, intLine)
        Form1.imgTransport(1).Top = pxlRow(intChar, intLine)
    Case "t"
        CellContents(intChar, intLine) = cEmpty
        CellCharactr(intChar, intLine) = stChar
        Transport2X = intChar
        Transport2Y = intLine
        Form1.imgTransport(2).Visible = True
        Form1.imgTransport(2).Left = pxlColumn(intChar, intLine)
        Form1.imgTransport(2).Top = pxlRow(intChar, intLine)
    Case "j"
        CellContents(intChar, intLine) = cJoeLeft
        CellCharactr(intChar, intLine) = stChar
        JoeX = intChar
        JoeY = intLine
    Case "J"
        CellContents(intChar, intLine) = cJoeRight
        CellCharactr(intChar, intLine) = stChar
        JoeX = intChar
        JoeY = intLine
    Case "X"
        CellContents(intChar, intLine) = cEmpty
        CellCharactr(intChar, intLine) = stChar
        ExitColumn = intChar
        ExitRow = intLine
        
        Form1.imgExit.Visible = True
        Form1.imgExit.Left = pxlColumn(intChar, intLine)
        Form1.imgExit.Top = pxlRow(intChar, intLine)
    Case "."
        CellContents(intChar, intLine) = cRailWalkOnce
        CellCharactr(intChar, intLine) = stChar
End Select
Return
'*****************************************************
'*****************************************************
'*****************************************************

'       empty = " " 1
'         box = "B" 2
'        roll = "O" 3
'        rail = "=" 4
'       brick = "#" 5
'   transport1 = "T" 6
'   transport2 = "t" 7
'     joeleft = "j" 8
'    joeright = "J" 9
'        exit = "X" 10
'railwalkonce = "." 11


lblNoMoreLevels:
rtnFinale
End
End Sub

Public Sub rtnRefreshFromArray()
Dim x As Integer, y As Integer
Form1.Enabled = False 'Prevent unwanted events while busy

'Paint screen from arrays
'Blocks x = 0 to 19    y = 0 to 14    20x15 Blocks   Total 300 Blocks
For y = 1 To 15
For x = 1 To 20
Form1.Picture1.PaintPicture picCell(CellContents(x, y)), pxlColumn(x, y), pxlRow(x, y)
Next x
Next y

Form1.Picture1.Refresh
Form1.Enabled = True
End Sub

'Enters WITHOUT box
Public Sub rtnRightKey()
With Form1
'***********************************************************************
'if bumping a roll
If CellContents(JoeX + 1, JoeY) = cRoll And CellContents(JoeX, JoeY) = cJoeRight Then
TempObject = cRoll
TempColumn = JoeX + 1
TempRow = JoeY
rtnRollRight
Exit Sub
End If
'***********************************************************************
   
'***********************************************************************
'Just turn Joe Right and Exit if Joe was facing left
If CellContents(JoeX, JoeY) = cJoeLeft Then
'       Paint           This Picture  at   Pixelx   of  Column,  Row    Pixely of  Column,  Row
.Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)

TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
CellContents(JoeX, JoeY) = cJoeRight 'Update Arrays
CellCharactr(JoeX, JoeY) = "J" 'Update Arrays
If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"

Exit Sub
End If
'***********************************************************************
   
'If right block empty then move right
If CellContents(JoeX + 1, JoeY) = cEmpty Then
'       Paint           This Picture  at   Pixelx   of     Column,  Row      Pixely   of   Column,  Row
.Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX + 1, JoeY), pxlRow(JoeX + 1, JoeY)
.Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)

TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
CellContents(JoeX, JoeY) = cEmpty 'Update Arrays
CellCharactr(JoeX, JoeY) = " " 'Update Arrays
If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"

JoeX = JoeX + 1
TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
CellContents(JoeX, JoeY) = cJoeRight 'Update Arrays
CellCharactr(JoeX, JoeY) = "J" 'Update Arrays
If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"
       
   
   'If previous block was a RailWalkOnce then
    If CellContents(JoeX - 1, JoeY + 1) = cRailWalkOnce Then
   'Make it disappear
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX - 1, JoeY + 1), pxlRow(JoeX - 1, JoeY + 1)
    CellContents(JoeX - 1, JoeY + 1) = cEmpty 'Update Arrays
    CellCharactr(JoeX - 1, JoeY + 1) = " " 'Update Arrays
    End If
            
    'Check if Transport and process
    If Transport1X = JoeX And Transport1Y = JoeY Then
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellContents(Transport2X, Transport2Y) = cJoeRight 'Update
    JoeX = Transport2X
    JoeY = Transport2Y
    rtnRefreshFromArray
    Exit Sub
    Else
        If Transport2X = JoeX And Transport2Y = JoeY Then
        CellContents(JoeX, JoeY) = cEmpty 'Update
        CellContents(Transport1X, Transport1Y) = cJoeRight 'Update
        JoeX = Transport1X
        JoeY = Transport1Y
        rtnRefreshFromArray
        Exit Sub
        End If
    End If

    'Check if Exit
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    .Picture1.Enabled = False
    rtnInitialize
    .Picture1.Enabled = True
    Exit Sub
    End If
       
    'Check if should fall
    TempObject = cJoeRight
    TempColumn = JoeX
    TempRow = JoeY
    rtnCheckIfFall
    
    'Check if Exit
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    rtnInitialize
    Exit Sub
    End If
    
End If

End With
End Sub

Public Sub rtnRightKeyWithBox()
'Enters WITH box
With Form1
   '***********************************************************************
'Just turn Joe Right and Exit
'if Joe was facing left
    If CellContents(JoeX, JoeY) = cJoeLeft Then
   '       Paint           This Picture at Pxlx   of     Column,  Row  Pixely  of   Column,  Row
   .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cJoeRight 'Update
    CellCharactr(JoeX, JoeY) = "J"
    Exit Sub
    End If
   '***********************************************************************
    
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************
   'Non obstructed right walk
   'If right block empty then move right            and             right top block is empty
    If CellContents(JoeX + 1, JoeY) = cEmpty And CellContents(JoeX + 1, JoeY - 1) = cEmpty Then
   
   'Paint JoeRight at Adjacent Right cell and Clear Current cell
   .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX + 1, JoeY), pxlRow(JoeX + 1, JoeY)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
   'Paint Box at Adjacent Top Right cell and Clear Current Top cell
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX + 1, JoeY - 1), pxlRow(JoeX + 1, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX + 1, JoeY) = cJoeRight 'Update
    CellCharactr(JoeX + 1, JoeY) = "J"
    
    CellContents(JoeX, JoeY) = cEmpty  'Update
    CellCharactr(JoeX, JoeY) = " "
    
    CellContents(JoeX + 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = "B"
    
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
    JoeX = JoeX + 1
       
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX - 1, JoeY + 1) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX - 1, JoeY + 1), pxlRow(JoeX - 1, JoeY + 1)
             CellContents(JoeX - 1, JoeY + 1) = cEmpty
             CellCharactr(JoeX - 1, JoeY + 1) = " "
        End If
       '***********************************************************************
       '***********************************************************************
       
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
       
       
       '***********************************************************************
       '***********************************************************************
       'If lower block is empty then fall
        If CellContents(JoeX, JoeY + 1) = cEmpty Then
           '***********************************************************************
            booBusyFallingOrRolling = True
            Do While CellContents(JoeX, JoeY + 1) = cEmpty
           '       Paint JoeRight one cell lower
           .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX, JoeY + 1), pxlRow(JoeX, JoeY + 1)
           '       Paint Box one cell lower (where joe was) and clear current block
           .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
           .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
            CellContents(JoeX, JoeY - 1) = cEmpty 'Update
            CellCharactr(JoeX, JoeY - 1) = " "
            
            CellContents(JoeX, JoeY) = cBox 'Update
            CellCharactr(JoeX, JoeY) = "B"
            
            CellContents(JoeX, JoeY + 1) = cJoeRight 'Update
            CellCharactr(JoeX, JoeY + 1) = "J"
            JoeY = JoeY + 1
            Loop
            booBusyFallingOrRolling = False
           '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
        End If
       '***********************************************************************
       '***********************************************************************
    Exit Sub
    End If
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************


   '***********************************************************************
   '***********************************************************************
   '***********************************************************************
   'Top obstructed right walk
   'If right block empty                            and         right top block is not empty
    If CellContents(JoeX + 1, JoeY) = cEmpty And CellContents(JoeX + 1, JoeY - 1) <> cEmpty Then
       'Paint JoeRight at Adjacent Right cell and Clear Current cell
       .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX + 1, JoeY), pxlRow(JoeX + 1, JoeY)
       .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    
        CellContents(JoeX + 1, JoeY) = cJoeRight 'Update
        CellCharactr(JoeX + 1, JoeY) = "J"
        
        CellContents(JoeX, JoeY) = cEmpty  'Update
        CellCharactr(JoeX, JoeY) = " "
        JoeX = JoeX + 1
       
       '***********************************************************************
       'Box fall if nothing underneath
        TempObject = cBox
        TempColumn = JoeX - 1
        TempRow = JoeY - 1
        rtnCheckIfFall
        booWithBlock = False
       '***********************************************************************
            'Check if Transport and process
            If Transport1X = JoeX And Transport1Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport2X, Transport2Y) = cJoeRight 'Update
            JoeX = Transport2X
            JoeY = Transport2Y
            rtnRefreshFromArray
            Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeRight 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            rtnRefreshFromArray
            End If
            End If
       
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
           'Joe fall if nothing underneath
            If CellContents(JoeX, JoeY + 1) = cEmpty Then
            TempObject = cJoeRight
            TempColumn = JoeX
            TempRow = JoeY
            rtnCheckIfFall
            End If
       '***********************************************************************
       '***********************************************************************
    
    Exit Sub
    End If
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************


End With
End Sub

Public Sub rtnLeftKeyWithBox()
'Enters WITH box
With Form1
   '***********************************************************************
   'Just turn joe
   'if left block not empty or joe facing right
   'then just turn Joe if not already facing left
    If CellContents(JoeX - 1, JoeY) <> cEmpty Or CellContents(JoeX, JoeY) = cJoeRight Then
   '       Paint           This Picture at Pxlx   of     Column,  Row  Pixely  of   Column,  Row
   .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cJoeLeft 'Update
    CellCharactr(JoeX, JoeY) = "j"
    Exit Sub
    End If
   '***********************************************************************
    
   '***********************************************************************
   '***********************************************************************3
   '***********************************************************************
   'Non obstructed left walk
   'If left block empty then move left            and             left top block is empty
    If CellContents(JoeX - 1, JoeY) = cEmpty And CellContents(JoeX - 1, JoeY - 1) = cEmpty Then
   
   'Paint Joeleft at Adjacent left cell and Clear Current cell
   .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX - 1, JoeY), pxlRow(JoeX - 1, JoeY)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
   'Paint Box at Adjacent Top left cell and Clear Current Top cell
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX - 1, JoeY - 1), pxlRow(JoeX - 1, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX - 1, JoeY) = cJoeLeft 'Update
    CellCharactr(JoeX - 1, JoeY) = "j"
    
    CellContents(JoeX, JoeY) = cEmpty  'Update
    CellCharactr(JoeX, JoeY) = " "
    
    CellContents(JoeX - 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = "B"
    
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
    JoeX = JoeX - 1
       
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX + 1, JoeY + 1) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX + 1, JoeY + 1), pxlRow(JoeX + 1, JoeY + 1)
             CellContents(JoeX + 1, JoeY + 1) = cEmpty
             CellCharactr(JoeX + 1, JoeY + 1) = " "
        End If
       '***********************************************************************
       '***********************************************************************
       
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
       
       '***********************************************************************
       '***********************************************************************
       'If lower block is empty then fall
        If CellContents(JoeX, JoeY + 1) = cEmpty Then
           '***********************************************************************
            booBusyFallingOrRolling = True
            Do While CellContents(JoeX, JoeY + 1) = cEmpty
           '       Paint Joeleft one cell lower
           .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX, JoeY + 1), pxlRow(JoeX, JoeY + 1)
           '       Paint Box one cell lower (where joe was) and clear current block
           .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
           .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
            CellContents(JoeX, JoeY - 1) = cEmpty 'Update
            CellCharactr(JoeX, JoeY - 1) = " "
            
            CellContents(JoeX, JoeY) = cBox 'Update
            CellCharactr(JoeX, JoeY) = "B"
            
            CellContents(JoeX, JoeY + 1) = cJoeLeft 'Update
            CellCharactr(JoeX, JoeY + 1) = "j"
            JoeY = JoeY + 1
            Loop
            booBusyFallingOrRolling = False
           '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
        End If
       '***********************************************************************
       '***********************************************************************
    Exit Sub
    End If
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************


   '***********************************************************************
   '***********************************************************************
   '***********************************************************************
   'Top obstructed left walk
   'If left block empty                            and         left top block is not empty
    If CellContents(JoeX - 1, JoeY) = cEmpty And CellContents(JoeX - 1, JoeY - 1) <> cEmpty Then
       'Paint Joeleft at Adjacent left cell and Clear Current cell
       .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX - 1, JoeY), pxlRow(JoeX - 1, JoeY)
       .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    
        CellContents(JoeX - 1, JoeY) = cJoeLeft 'Update
        CellCharactr(JoeX - 1, JoeY) = "j"
        
        CellContents(JoeX, JoeY) = cEmpty  'Update
        CellCharactr(JoeX, JoeY) = " "
        JoeX = JoeX - 1
       
       '***********************************************************************
       'Joe fall if nothing underneath
        TempObject = cBox
        TempColumn = JoeX + 1
        TempRow = JoeY - 1
        rtnCheckIfFall
        booWithBlock = False
       '***********************************************************************
           
            'Check if Transport and process
            If Transport1X = JoeX And Transport1Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport2X, Transport2Y) = cJoeLeft 'Update
            JoeX = Transport2X
            JoeY = Transport2Y
            rtnRefreshFromArray
            Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeLeft 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            rtnRefreshFromArray
            End If
            End If
           
       '***********************************************************************
           
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
           'If lower block is empty then fall
            If CellContents(JoeX, JoeY + 1) = cEmpty Then
           'Joe fall if nothing underneath
            TempObject = cJoeLeft
            TempColumn = JoeX
            TempRow = JoeY
            rtnCheckIfFall
            End If
           '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
       '***********************************************************************
    Exit Sub
    End If
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************






End With
End Sub

Public Sub rtnLeftKey()
'Enters WITHOUT box
With Form1
'***********************************************************************
'if bumping a roll
If CellContents(JoeX - 1, JoeY) = cRoll And CellContents(JoeX, JoeY) = cJoeLeft Then
TempObject = cRoll
TempColumn = JoeX - 1
TempRow = JoeY
rtnRollLeft
Exit Sub
End If
'***********************************************************************
   
'***********************************************************************
'Just turn Joe Left and Exit
'if Joe was facing right
 If CellContents(JoeX, JoeY) = cJoeRight Then
' Paint  This Picture at Pxlxof  Column,  Row  Pixely  ofColumn,  Row
.Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
 TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
 CellContents(JoeX, JoeY) = cJoeLeft 'Update
 CellCharactr(JoeX, JoeY) = "j"
 If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
 If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"
 Exit Sub
 End If
'***********************************************************************
 
'***********************************************************************
'***********************************************************************
'***********************************************************************
'If left block empty then move left
 If CellContents(JoeX - 1, JoeY) = cEmpty Then
'Paint  This Picture  atPixelxof  Column,  RowPixelyofColumn,  Row
.Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX - 1, JoeY), pxlRow(JoeX - 1, JoeY)
.Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
 
 TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
 CellContents(JoeX, JoeY) = cEmpty 'Update
 CellCharactr(JoeX, JoeY) = " "
 If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
 If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"
 JoeX = JoeX - 1
 
 TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
 CellContents(JoeX, JoeY) = cJoeLeft 'Update
 CellCharactr(JoeX, JoeY) = "j"
 If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
 If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"

   'If block was a RailWalkOnce
    If CellContents(JoeX + 1, JoeY + 1) = cRailWalkOnce Then
   'Make it disappear
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX + 1, JoeY + 1), pxlRow(JoeX + 1, JoeY + 1)
    CellContents(JoeX + 1, JoeY + 1) = cEmpty
    CellCharactr(JoeX + 1, JoeY + 1) = " "
    End If
           
    'Check if Transport and process
    If Transport1X = JoeX And Transport1Y = JoeY Then
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellContents(Transport2X, Transport2Y) = cJoeLeft 'Update
    JoeX = Transport2X
    JoeY = Transport2Y
    rtnRefreshFromArray
    Exit Sub
    Else
        If Transport2X = JoeX And Transport2Y = JoeY Then
        CellContents(JoeX, JoeY) = cEmpty 'Update
        CellContents(Transport1X, Transport1Y) = cJoeLeft 'Update
        JoeX = Transport1X
        JoeY = Transport1Y
        rtnRefreshFromArray
        Exit Sub
        End If
    End If

    'Check if Exit and process
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    rtnInitialize
    Exit Sub
    End If
       
    'Check if should fall
    TempObject = cJoeLeft
    TempColumn = JoeX
    TempRow = JoeY
    rtnCheckIfFall
    
    'Check if Exit and process
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    rtnInitialize
    Exit Sub
    End If

End If

End With
End Sub

Public Sub rtnUpKey()
'Enters WITHOUT box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'If    Adjacent Right Top Block is Empty             and                  Top Block is empty          and    Adjacent Right Block is not  Empty         then
    If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) <> cEmpty Then
   
   'Then Jump up
   'Paint and clear Top Block
   .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    JoeY = JoeY - 1
    CellContents(JoeX, JoeY) = cJoeRight 'Update
    CellCharactr(JoeX, JoeY) = "J"
   'Paint Adjacent Right Top Block
   .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX + 1, JoeY), pxlRow(JoeX + 1, JoeY)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    JoeX = JoeX + 1
    CellContents(JoeX, JoeY) = cJoeRight 'Update
    CellCharactr(JoeX, JoeY) = "J"
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX - 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX - 1, JoeY + 2), pxlRow(JoeX - 1, JoeY + 2)
             CellContents(JoeX - 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX - 1, JoeY + 2) = " "
        End If
       '***********************************************************************
       '***********************************************************************
       
            'Check if Transport and process
            If Transport1X = JoeX And Transport1Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport2X, Transport2Y) = cJoeRight 'Update
            JoeX = Transport2X
            JoeY = Transport2Y
            rtnRefreshFromArray
            Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeRight 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            rtnRefreshFromArray
            End If
            End If
       
       
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
    'If lower block is empty then fall
    If CellContents(JoeX, JoeY + 1) = cEmpty Then
    'Joe fall if nothing underneath
    TempObject = cJoeRight
    TempColumn = JoeX
    TempRow = JoeY
    rtnCheckIfFall
    End If
    
    End If

  Case cJoeLeft
   'If    Adjacent Left  Top Block is Empty             and                  Top Block is empty          and    Adjacent Left  Block is not  Empty         then
    If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) <> cEmpty Then
   'Paint and clear Top Block
   .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
   JoeY = JoeY - 1
    CellContents(JoeX, JoeY) = cJoeLeft 'Update
    CellCharactr(JoeX, JoeY) = "j"
   'Paint Adjacent Right Top Block
   .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX - 1, JoeY), pxlRow(JoeX - 1, JoeY)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    JoeX = JoeX - 1
    CellContents(JoeX, JoeY) = cJoeLeft 'Update
    CellCharactr(JoeX, JoeY) = "j"
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX + 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX + 1, JoeY + 2), pxlRow(JoeX + 1, JoeY + 2)
             CellContents(JoeX + 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX + 1, JoeY + 2) = " "
        End If
       '***********************************************************************
       '***********************************************************************
           
            'Check if Transport and process
            If Transport1X = JoeX And Transport1Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport2X, Transport2Y) = cJoeLeft 'Update
            JoeX = Transport2X
            JoeY = Transport2Y
            rtnRefreshFromArray
            Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeLeft 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            rtnRefreshFromArray
            End If
            End If
           
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
    'If lower block is empty then fall
    If CellContents(JoeX, JoeY + 1) = cEmpty Then
    'Joe fall if nothing underneath
    TempObject = cJoeLeft
    TempColumn = JoeX
    TempRow = JoeY
    rtnCheckIfFall
    End If
    
    End If

End Select
End With
End Sub

Public Sub rtnUpKeyWithBox()
'Enters WITHOUT box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'If    Adjacent Right Top Top Block is Empty         and       Adjacent Right Top Block is Empty          and          Top Top Block is empty              and    right block is not empty                 then
    If CellContents(JoeX + 1, JoeY - 2) = cEmpty And CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 2) = cEmpty And CellContents(JoeX + 1, JoeY) <> cEmpty Then
   'Paint step 1
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX, JoeY - 2), pxlRow(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
   
   'Paint step 2
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX + 1, JoeY - 2), pxlRow(JoeX + 1, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeRight), pxlColumn(JoeX + 1, JoeY - 1), pxlRow(JoeX + 1, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 2), pxlRow(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
    
    CellContents(JoeX, JoeY - 2) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 2) = " "
    
    CellContents(JoeX + 1, JoeY - 1) = cJoeRight 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = "J"
    
    CellContents(JoeX + 1, JoeY - 2) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY - 2) = "B"
    JoeX = JoeX + 1
    JoeY = JoeY - 1
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX - 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX - 1, JoeY + 2), pxlRow(JoeX - 1, JoeY + 2)
             CellContents(JoeX - 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX - 1, JoeY + 2) = " "
        End If
       '***********************************************************************
       '***********************************************************************
    End If


  Case cJoeLeft
   'If    Adjacent left Top Top Block is Empty         and       Adjacent left Top Block is Empty          and          Top Top Block is empty              and    left block is not empty                 then
    If CellContents(JoeX - 1, JoeY - 2) = cEmpty And CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 2) = cEmpty And CellContents(JoeX - 1, JoeY) <> cEmpty Then
   'Paint step 1
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX, JoeY - 2), pxlRow(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY), pxlRow(JoeX, JoeY)
   
   'Paint step 2
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX - 1, JoeY - 2), pxlRow(JoeX - 1, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeLeft), pxlColumn(JoeX - 1, JoeY - 1), pxlRow(JoeX - 1, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 2), pxlRow(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
    
    CellContents(JoeX, JoeY - 2) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 2) = " "
    
    CellContents(JoeX - 1, JoeY - 1) = cJoeLeft 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = "j"
    
    CellContents(JoeX - 1, JoeY - 2) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY - 2) = "B"
    JoeX = JoeX - 1
    JoeY = JoeY - 1
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX + 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX + 1, JoeY + 2), pxlRow(JoeX + 1, JoeY + 2)
             CellContents(JoeX + 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX + 1, JoeY + 2) = " "
        End If
       '***********************************************************************
       '***********************************************************************
    End If

End Select
End With
End Sub

Public Sub rtnDownKey()
'Enters WITHOUT box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'If    Adjacent Right Top Block is Empty             and                  Top Block is empty          and    Adjacent Right Block is a box         then
    If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) = cBox Then
   
   'Paint Adjacent Right Top Box and Clear Adjacent Right Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX + 1, JoeY - 1), pxlRow(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX + 1, JoeY), pxlRow(JoeX + 1, JoeY)
    CellContents(JoeX + 1, JoeY) = cEmpty 'Update
    CellCharactr(JoeX + 1, JoeY) = " "

   'Paint Top Box and Clear Adjacent Right Top Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX + 1, JoeY - 1), pxlRow(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = " "
    booWithBlock = True
    End If

  Case cJoeLeft
   'If    Adjacent Right Top Block is Empty             and                  Top Block is empty          and    Adjacent Right Block is a box         then
    If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) = cBox Then
   
   'Paint Adjacent Left Top Box and Clear Adjacent Left Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX - 1, JoeY - 1), pxlRow(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX - 1, JoeY), pxlRow(JoeX - 1, JoeY)
    CellContents(JoeX - 1, JoeY) = cEmpty 'Update
    CellCharactr(JoeX - 1, JoeY) = " "
   
   'Paint Top Box and Clear Adjacent Left Top Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX - 1, JoeY - 1), pxlRow(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = " "
    booWithBlock = True
    End If

End Select
End With

End Sub

Public Sub rtnDownKeyWithBox()
'Enters WITH box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'Place box down
   'If    Adjacent Right Top Block is Empty             and    Adjacent Right Block is empty         then
    If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) = cEmpty Then
   
   'Paint Adjacent Right Top Box and Clear Top Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX + 1, JoeY - 1), pxlRow(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
   
   'Paint Adjacent Right Box and Clear Adjacent Right Top Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX + 1, JoeY), pxlRow(JoeX + 1, JoeY)
    CellContents(JoeX + 1, JoeY) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX + 1, JoeY - 1), pxlRow(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = " "
    booWithBlock = False
    
    'Box fall is nothing underneath
    TempObject = cBox
    TempColumn = JoeX + 1
    TempRow = JoeY
    rtnCheckIfFall
    Exit Sub
    End If

       'Place box on top of object
       'If    Adjacent Right Top Block is Empty             and    Adjacent Right Block is not empty         then
        If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) <> cEmpty Then
  
       'Paint Adjacent Right Top Box and Clear Top Box
       .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX + 1, JoeY - 1), pxlRow(JoeX + 1, JoeY - 1)
        CellContents(JoeX + 1, JoeY - 1) = cBox 'Update
        CellCharactr(JoeX + 1, JoeY - 1) = "B"
       .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
        CellContents(JoeX, JoeY - 1) = cEmpty 'Update
        CellCharactr(JoeX, JoeY - 1) = " "
        booWithBlock = False
        Exit Sub
        End If
  
  Case cJoeLeft
   'If    Adjacent Left Top Block is Empty              and    Adjacent Left Block is empty         then
    If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) = cEmpty Then
   
   'Paint Adjacent Left Top Box and Clear Top Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX - 1, JoeY - 1), pxlRow(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
   
   'Paint Adjacent Left Box and Clear Adjacent Left Top Box
   .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX - 1, JoeY), pxlRow(JoeX - 1, JoeY)
    CellContents(JoeX - 1, JoeY) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY) = "B"
   .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX - 1, JoeY - 1), pxlRow(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = " "
    booWithBlock = False
    
    'Box fall is nothing underneath
    TempObject = cBox
    TempColumn = JoeX - 1
    TempRow = JoeY
    rtnCheckIfFall
    Exit Sub
    End If

       'Place box on top of object
       'If    Adjacent Left Top Block is Empty             and    Adjacent Left Block is not empty         then
        If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) <> cEmpty Then
  
       'Paint Adjacent Left Top Box and Clear Top Box
       .Picture1.PaintPicture picCell(cBox), pxlColumn(JoeX - 1, JoeY - 1), pxlRow(JoeX - 1, JoeY - 1)
        CellContents(JoeX - 1, JoeY - 1) = cBox 'Update
        CellCharactr(JoeX - 1, JoeY - 1) = "B"
       .Picture1.PaintPicture picCell(cEmpty), pxlColumn(JoeX, JoeY - 1), pxlRow(JoeX, JoeY - 1)
        CellContents(JoeX, JoeY - 1) = cEmpty 'Update
        CellCharactr(JoeX, JoeY - 1) = " "
        booWithBlock = False
        Exit Sub
        End If



End Select
End With

End Sub


Public Sub rtnCheckIfFall()
'Enter with
'TempObject holding  object # to fall down
'TempColumn holding object's x cell value 1 - 20
'TempRow holding object's y cell value 1 - 15
booBusyFallingOrRolling = True
With Form1
        If CellContents(TempColumn, TempRow + 1) = cEmpty Then
           '***********************************************************************
            Do While CellContents(TempColumn, TempRow + 1) = cEmpty
           '       Paint           This Picture  at   Pixelx of   Column,  Row          Pixely of Column,  Row
           .Picture1.PaintPicture picCell(TempObject), pxlColumn(TempColumn, TempRow + 1), pxlRow(TempColumn, TempRow + 1)
           .Picture1.PaintPicture picCell(cEmpty), pxlColumn(TempColumn, TempRow), pxlRow(TempColumn, TempRow)
            CellContents(TempColumn, TempRow) = cEmpty 'Update
            CellCharactr(TempColumn, TempRow) = " "
            TempRow = TempRow + 1
            If TempObject = cJoeRight Then JoeY = JoeY + 1
            If TempObject = cJoeLeft Then JoeY = JoeY + 1
            CellContents(TempColumn, TempRow) = TempObject 'Update
            
              Select Case TempObject
              Case cBox
              CellCharactr(TempColumn, TempRow) = "B"
              Case cJoeRight
              CellCharactr(TempColumn, TempRow) = "J"
              Case cJoeLeft
              CellCharactr(TempColumn, TempRow) = "j"
              Case cRoll
              CellCharactr(TempColumn, TempRow) = "O"
              End Select
              
            DoEvents
            Loop
           '***********************************************************************
        End If
End With
booBusyFallingOrRolling = False
End Sub

Public Sub rtnAllRollsAndBoxesFall()
Dim x As Integer, y As Integer
'All Boxes Fall Down
'TempObject holding value of object to fall down
'TempColumn holding object's x cell value
'TempRow holding object's y cell value
booBusyFallingOrRolling = True
For y = 15 To 1 Step -1
For x = 1 To 20
If CellContents(x, y) = cRoll Or CellContents(x, y) = cBox Then
   TempObject = CellContents(x, y)
   TempColumn = x
   TempRow = y
   rtnCheckIfFall
End If
Next x
Next y
booBusyFallingOrRolling = False

End Sub

Public Sub rtnRollRight()
'Enter with
'TempObject holding  object # to fall down
'TempColumn holding object's x cell value 1 - 20
'TempRow holding object's y cell value 1 - 15
booBusyFallingOrRolling = True
With Form1
        If CellContents(TempColumn + 1, TempRow) = cEmpty Then
           '***********************************************************************
            Do While CellContents(TempColumn + 1, TempRow) = cEmpty
           '       Paint           This Picture  at   Pixelx of   Column,  Row          Pixely of Column,  Row
           .Picture1.PaintPicture picCell(TempObject), pxlColumn(TempColumn + 1, TempRow), pxlRow(TempColumn + 1, TempRow)
           .Picture1.PaintPicture picCell(cEmpty), pxlColumn(TempColumn, TempRow), pxlRow(TempColumn, TempRow)
            CellContents(TempColumn, TempRow) = cEmpty 'Update
            CellCharactr(TempColumn, TempRow) = " "
            TempColumn = TempColumn + 1
            CellContents(TempColumn, TempRow) = TempObject 'Update
              Select Case TempObject
              Case cBox
              CellCharactr(TempColumn, TempRow) = "B"
              Case cJoeRight
              CellCharactr(TempColumn, TempRow) = "J"
              Case cJoeLeft
              CellCharactr(TempColumn, TempRow) = "j"
              Case cRoll
              CellCharactr(TempColumn, TempRow) = "O"
              End Select
            
               '***********************************************************************
                Do While CellContents(TempColumn, TempRow + 1) = cEmpty
               '       Paint           This Picture  at   Pixelx of   Column,  Row          Pixely of Column,  Row
               .Picture1.PaintPicture picCell(TempObject), pxlColumn(TempColumn, TempRow + 1), pxlRow(TempColumn, TempRow + 1)
               .Picture1.PaintPicture picCell(cEmpty), pxlColumn(TempColumn, TempRow), pxlRow(TempColumn, TempRow)
                CellContents(TempColumn, TempRow) = cEmpty 'Update
                CellCharactr(TempColumn, TempRow) = " "
                TempRow = TempRow + 1
                CellContents(TempColumn, TempRow) = TempObject 'Update
              Select Case TempObject
              Case cBox
              CellCharactr(TempColumn, TempRow) = "B"
              Case cJoeRight
              CellCharactr(TempColumn, TempRow) = "J"
              Case cJoeLeft
              CellCharactr(TempColumn, TempRow) = "j"
              Case cRoll
              CellCharactr(TempColumn, TempRow) = "O"
              End Select
                DoEvents
                Loop
               '***********************************************************************
            
            
            DoEvents
            Loop
           '***********************************************************************
        End If
rtnAllRollsAndBoxesFall
End With
booBusyFallingOrRolling = False
End Sub

Public Sub rtnRollLeft()
'Enter with
'TempObject holding  object # to fall down
'TempColumn holding object's x cell value 1 - 20
'TempRow holding object's y cell value 1 - 15
booBusyFallingOrRolling = True
With Form1
        If CellContents(TempColumn - 1, TempRow) = cEmpty Then
           '***********************************************************************
            Do While CellContents(TempColumn - 1, TempRow) = cEmpty
           '       Paint           This Picture  at   Pixelx of   Column,  Row          Pixely of Column,  Row
           .Picture1.PaintPicture picCell(TempObject), pxlColumn(TempColumn - 1, TempRow), pxlRow(TempColumn - 1, TempRow)
           .Picture1.PaintPicture picCell(cEmpty), pxlColumn(TempColumn, TempRow), pxlRow(TempColumn, TempRow)
            CellContents(TempColumn, TempRow) = cEmpty 'Update
            CellCharactr(TempColumn, TempRow) = " "
            TempColumn = TempColumn - 1
            CellContents(TempColumn, TempRow) = TempObject 'Update
              Select Case TempObject
              Case cBox
              CellCharactr(TempColumn, TempRow) = "B"
              Case cJoeRight
              CellCharactr(TempColumn, TempRow) = "J"
              Case cJoeLeft
              CellCharactr(TempColumn, TempRow) = "j"
              Case cRoll
              CellCharactr(TempColumn, TempRow) = "O"
              End Select
            
               '***********************************************************************
                Do While CellContents(TempColumn, TempRow + 1) = cEmpty
               '       Paint           This Picture  at   Pixelx of   Column,  Row          Pixely of Column,  Row
               .Picture1.PaintPicture picCell(TempObject), pxlColumn(TempColumn, TempRow + 1), pxlRow(TempColumn, TempRow + 1)
               .Picture1.PaintPicture picCell(cEmpty), pxlColumn(TempColumn, TempRow), pxlRow(TempColumn, TempRow)
                CellContents(TempColumn, TempRow) = cEmpty 'Update
                CellCharactr(TempColumn, TempRow) = " "
                TempRow = TempRow + 1
                CellContents(TempColumn, TempRow) = TempObject 'Update
              Select Case TempObject
              Case cBox
              CellCharactr(TempColumn, TempRow) = "B"
              Case cJoeRight
              CellCharactr(TempColumn, TempRow) = "J"
              Case cJoeLeft
              CellCharactr(TempColumn, TempRow) = "j"
              Case cRoll
              CellCharactr(TempColumn, TempRow) = "O"
              End Select
                DoEvents
                Loop
               '***********************************************************************
            
            
            DoEvents
            Loop
           '***********************************************************************
        End If
rtnAllRollsAndBoxesFall
End With
booBusyFallingOrRolling = False
End Sub

Public Sub rtnFinale()
Dim dbltime As Double
With Form1
booBusyFallingOrRolling = True
.Picture1.Cls
.Picture1.BackColor = vbWhite
.Picture1.ForeColor = vbRed
.Picture1.FontSize = 50
.Picture1.Font = "arial"
.Picture1.Print "Congratulations!"
.Picture1.Print "You Win!"
.Picture1.Refresh
dbltime = Timer
Do While Timer < dbltime + 3
Loop
End With
End Sub

