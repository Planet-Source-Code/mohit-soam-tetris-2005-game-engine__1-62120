Attribute VB_Name = "ModCommonFunction"
'   __________                       __   _____
'   \________/ _____  /  |_ ______   |_| /  ___|
'      |  |  _/ ___ \_\____\\_  __ \ _ _ \_____
'      |  |   \  ___/   |  | |  | \/ |  | ____ \
'      \__/    \___    |__|  |__|    |__||_____/
'
' _________                               ____________
' _________ T H E   G A M E   B E G I N S ____________
'
' Tetris - A free 2D real time Logic game engine
' Copyright (C) 2002-2003 Mohit Soam
'----------------------------------------------------
' Project Name : Tetris 2005
' Author: Mohit Soam
' Email : mohit.soam@rediff.com
' Address: III/10,Hydel Colony
'          Veerbhadra
'          Rishikesh (U.A)
'          India 249202

' Date of Creation:  July 1,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.

' -------------------------------------
' --> A Module contains Common Function
' -------------------------------------

Option Explicit

'Purpose : It is used to set Intialize the array
'          NextPosition to -1
Public Sub InitializeNextBrick()
Dim i As Integer
For i = 0 To 5
    NextPosition(i) = -1
Next i

End Sub

'Purpose : It is used to Assign the value of array
'          CurrentPosition to array NextPosition
Public Sub RestoreNextBrick()
Dim i As Integer
For i = 0 To 5
    NextPosition(i) = CurrentPosition(i)
Next i

End Sub

'Purpose : It is used to Increment the values of array
'          CurrentPosition i.e to change the position
'          of Brick and store New values in array
'          NextPosition.
'========== To Move Left call Increment(-1)
'========== To Move Right call Increment(1)
'========== To Move Down call Increment(10)
Public Sub Increment(X As Integer)
Dim i As Integer

RightBrick = False
LeftBrick = False

For i = 0 To 5
    
    If CurrentPosition(i) > 289 Then
'========== if i > 289 the Draw new Brick
        frmMain.Timer1.Interval = 1001 - _
        (frmOption.Slider1.Value * 50) '===} To set
        frmMain.Timer2.Enabled = False '===} Brick
        frmMain.Timer3.Enabled = False '===} Speed
        frmMain.Timer4.Enabled = False '===}

    GAME_OVER = CheckGameOver(frmMain)
    If GAME_OVER Then
        TimeCount = 0
        
        frmOption.Slider1.Value = 1
        frmMain.txtlevel = "1"
        frmOptionSetting
        
        CompareScore
        Exit Sub
     Else
        Call CheckRowBuilt(frmMain)
        If BonusCount = 0 Then
            Call DrawBlock(frmMain) '==========} Draw
        End If                      '          } Brick
        
        START = True '=====================} New
        SelectBrick '======================} Brick
  
        Exit Sub
    End If
         
    End If
    
    If CurrentPosition(i) Mod 10 = 9 And _
        X = 1 Then
        RightBrick = True
        RestoreNextBrick
        Exit Sub
    
    ElseIf CurrentPosition(i) Mod 10 = 0 And _
        X = -1 Then
        LeftBrick = True
        RestoreNextBrick
        Exit Sub
    
    End If
    
    If RightBrick = False And LeftBrick = False Then
        If CurrentPosition(i) <> -1 Then
            NextPosition(i) = CurrentPosition(i) + X
        Else
            NextPosition(i) = -1
        End If
     End If
    
Next i

End Sub

'Purpose : If It is used to Assign the value of array
'          CurrentPosition to array NextPosition
Public Sub MoveBlock()
Dim i As Integer

Call HideBrick(frmMain)

For i = 0 To 5
    If AVILABLE = True Then
        If NextPosition(i) <> -1 Then
          
          frmMain.Imgbox(NextPosition(i)).Picture = _
            frmMain.ImageList1.ListImages(Color).Picture
          
          frmMain.Imgbox(NextPosition(i)).Visible = True

          CurrentPosition(i) = NextPosition(i)
        
        End If
     End If
Next i

Call ShowPositionPointer(frmMain)

AVILABLE = False

End Sub

'Purpose : It is used to count the number of Bricks
'          visible in row if BrickCount = 10 i.e, all
'          the Brick are visible then Hide that row and
'          Move all the Bricks above one step Down.
Public Sub CheckRowBuilt(frm As Form)
Dim i As Integer '=====]
Dim j As Integer '=======> used in For Loop
Dim k As Integer '=====]

BonusCount = 0

For i = 50 To 299 Step 10
    
    For j = 0 To 9
    
        If frm.Imgbox(i + j).Visible = True Then
            BrickCount = BrickCount + 1
        End If
    
    Next j
     
    If BrickCount = 10 Then

'=======To Increase Score.
        frm.txtscore.Text = CLng(frm.txtscore.Text) + 100
        BonusCount = BonusCount + 1
        
'=======To Move all the Bricks above one step Down.
        For j = i To 50 Step -10
            
            For k = 0 To 9
                    
                If (j + k) < 60 Then
                    frm.Imgbox(j + k).Visible = False
                End If
                  
                If frm.Imgbox(j + k - 10).Visible = True Then
                    frm.Imgbox(j + k).Picture = _
                    frm.Imgbox(j + k - 10).Picture
                    frm.Imgbox(j + k).Visible = True
                Else
                    frm.Imgbox(j + k).Visible = False
                End If
             
             Next k
         
         Next j
       
    End If

BrickCount = 0

Next i

End Sub

'Purpose : It is used to find out the status of the game
'          Before drawing new brick we check 0 to 49
'          brick if any one is Visible then game is Over.
Public Function CheckGameOver(frm As Form) As Boolean
Dim i As Integer
    For i = 49 To 0 Step -1
        If frm.Imgbox(i).Visible = True Then
            CheckGameOver = True
            Exit Function
        End If
    Next i
End Function

'Purpose : It is used to check Postion that are stored
'          in array NextPosition i.e to Position after
'          after Increment in CurrentPosition
Public Sub CheckStatus(X As Integer, a As Integer, _
    Optional b, Optional c, Optional d)

MaxBrick = True

If IsMissing(b) And IsMissing(c) And _
        IsMissing(d) Then
      
    MaxBrick = OutOfRange(a)
'Purpose of OutOfRange(a) :
'=========================== It is used to check
'         wether Postion a is greater than or less then
'         290.If greater than it return False Else True.
  
  If MaxBrick Then
    If frmMain.Imgbox(CurrentPosition(a) + X).Visible = False Then
        AVILABLE = True
        Exit Sub
    End If
  End If

ElseIf IsMissing(c) And IsMissing(d) Then
  
  MaxBrick = OutOfRange(a, b)
  
  If MaxBrick Then
    If frmMain.Imgbox(CurrentPosition(a) + X).Visible = False And _
       frmMain.Imgbox(CurrentPosition(b) + X).Visible = False Then
        AVILABLE = True
        Exit Sub
    End If
  End If

ElseIf IsMissing(d) Then
  
  MaxBrick = OutOfRange(a, b, c)
  
  If MaxBrick Then
    If frmMain.Imgbox(CurrentPosition(a) + X).Visible = False And _
       frmMain.Imgbox(CurrentPosition(b) + X).Visible = False And _
       frmMain.Imgbox(CurrentPosition(c) + X).Visible = False Then
        AVILABLE = True
        Exit Sub
    End If
   End If
        
Else
  
    MaxBrick = OutOfRange(a, b, c, d)
   If MaxBrick Then
    If frmMain.Imgbox(CurrentPosition(a) + X).Visible = False And _
       frmMain.Imgbox(CurrentPosition(b) + X).Visible = False And _
       frmMain.Imgbox(CurrentPosition(c) + X).Visible = False And _
       frmMain.Imgbox(CurrentPosition(d) + X).Visible = False Then
        AVILABLE = True
    End If
   End If
End If
End Sub

'Purpose : It is used to check wether Postion 'a' or 'b'
'          'c' or 'd'  is greater than or less then 290
'           If greater than it return False Else True.
Public Function OutOfRange(Optional a, _
    Optional b, Optional c, Optional d) As Boolean

If IsMissing(a) And IsMissing(b) And _
       IsMissing(c) And IsMissing(d) Then
    OutOfRange = True
    Exit Function
   
ElseIf IsMissing(b) And IsMissing(c) And _
        IsMissing(d) Then
   If CurrentPosition(a) > 289 Then
    OutOfRange = False
    Exit Function
   End If
    
ElseIf IsMissing(c) And IsMissing(d) Then
   If CurrentPosition(a) > 289 And _
   CurrentPosition(b) > 289 Then
    OutOfRange = False
    Exit Function
   End If

ElseIf IsMissing(d) Then
    If CurrentPosition(a) > 289 And CurrentPosition(b) > 289 And _
   CurrentPosition(c) > 289 Then
    OutOfRange = False
    Exit Function
   End If
    
Else
   If CurrentPosition(a) > 289 And CurrentPosition(b) > 289 And _
   CurrentPosition(c) > 289 And CurrentPosition(c) > 289 Then
    OutOfRange = False
    Exit Function
   End If
   
End If

OutOfRange = True

End Function

'Purpose : It is used to Set the values of the controls
'          in the frmOption Form
Public Sub frmOptionSetting()

If frmOption.Option1.Value = True Then
    frmOption.Slider1.Value = 10
    frmMain.txtlevel = "Beginner"
    Exit Sub
    
ElseIf frmOption.Option2.Value = True Then
    frmOption.Slider1.Value = 15
    frmMain.txtlevel = "Intermediate"
    Exit Sub
    
ElseIf frmOption.Option3.Value = True Then
    frmOption.Slider1.Value = 20
    frmMain.txtlevel = "Expert"
    Exit Sub
    
ElseIf frmOption.Slider1.Value <= 17 Then
    If TimeCount >= 200 Then
        TimeCount = 0
        frmOption.Slider1.Value = _
            frmOption.Slider1.Value + 1
        frmMain.txtlevel = frmOption.Slider1.Value
        Beep
    Else
        TimeCount = TimeCount + 1
    End If
End If
End Sub

'Purpose : It is used to Load score from the files to
'          the Labels in the frmScore Form.
Public Sub LoadScoreFile(frm As Form)
Dim X As Integer
Dim i As Integer
Dim FNum As Integer
Dim Strlen As Integer
Dim TextLine

Load frmScore

On Error GoTo FileError:

FNum = FreeFile
Open App.Path & "\Score.dat" For Input As #FNum   ' Open file.
X = 1
i = 0
    Do While Not EOF(1) And i <= 6 ' Loop until end of file.
        Line Input #FNum, TextLine   ' Read line into variable.
   
        frm.lblSno(i).Caption = i + 1 & ". "
   
        Strlen = InStr(X, TextLine, ",")
        frm.lblName(i).Caption = Mid(TextLine, 1, Strlen - 1)
   
        TextLine = Mid(TextLine, Strlen + 1, InStr(X, TextLine, ";"))
        Strlen = InStr(X, TextLine, ":")
        frm.lblquote(i).Caption = Mid(TextLine, 1, Strlen - 1)
   
        TextLine = Mid(TextLine, Strlen + 1, InStr(X, TextLine, ";"))
        Strlen = InStr(X, TextLine, ";")
        frm.lblScore(i).Caption = Mid(TextLine, 1, Strlen - 1)
   
        i = i + 1
        Debug.Print TextLine   ' Print to the Immediate window.
    Loop
Close #1   ' Close file.

Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file "
    'OpenFile = ""
End Sub

'Purpose : It is used to Save score from the Labels to
'          the file "Score.dat" in Current Directory.
Public Sub SaveScoreFile(frm As Form)
Dim X As Integer
Dim FNum As Integer

Load frmScore

On Error GoTo FileError:

FNum = FreeFile
Open App.Path & "\Score.dat" For Output As #FNum   ' Open file.
For X = 0 To 6
   Print #FNum, frm.lblName(X).Caption & "," & _
        frm.lblquote(X).Caption & ":" & _
        frm.lblScore(X).Caption & ";"
Next X
Close #1   ' Close file.

Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file "
    'OpenFile = ""
End Sub

'Purpose : It is used to Compare the Current Score of
'          user with saved score.
Public Sub CompareScore()
Dim i As Integer

Load frmScore

i = 0
Do While i < 7
If frmScore.lblScore(i).Caption <> "" Then
    If CLng(frmScore.lblScore(i).Caption) < CLng(frmMain.txtscore) Then
        frmScore.lblScore(i).Caption = frmMain.txtscore.Text
        frmWin.Text1.Text = frmScore.lblName(i).Caption
        frmWin.Text2.Text = frmScore.lblquote(i).Caption
        ScoreIndex = i
        i = 9
        Call RestartNewGame(frmMain)
        frmMain.Hide
        frmWin.Show
        Exit Sub
    Else
        i = i + 1
    End If
End If
Loop

If frmWin.Visible = False Then MsgBox ("Try Again")

Call RestartNewGame(frmMain)

End Sub

'Purpose : It is used to do some setting before
'          starting a new game.
Public Sub RestartNewGame(frm As Form)

With frm
    .imgbutton(1).Visible = False
    .imgbutton(0).Visible = True
    .imgbutton(6).Visible = False
    .imgbutton(7).Visible = False
    .imgbutton(8).Visible = False
    .imgbutton(9).Visible = False
     
    .Imgpausedis.Visible = True
    .Imgmousebutdis.Visible = True
    .Imgstartdis.Visible = False
     
    .Timer1.Enabled = False
End With

End Sub

