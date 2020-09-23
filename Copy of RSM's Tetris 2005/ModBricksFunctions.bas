Attribute VB_Name = "ModBrickFunction"
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
'
' Date of Creation:  July 2,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.
'
' -------------------------------------
' --> A Module to Control Brick
' -------------------------------------

Option Explicit
Dim i As Integer '====== used in For Loop

'Purpose : It is used to Hide all ImgBox(ImageBox)
'          in the frm (frmMain Form).
Public Sub HideBrick(frm As Form)

For i = 0 To 5
    If CurrentPosition(i) <> -1 Then
        frm.Imgbox(CurrentPosition(i)).Visible = False
    End If
Next i

End Sub

'Purpose : It is used to Show Block i.e set the
'          selected Image Box visible according to the
'          position specified in SetCurrentPosition
'          function.
Public Sub DrawBlock(frm As Form)

For i = 0 To 5
    If CurrentPosition(i) <> -1 Then
        frm.Imgbox(CurrentPosition(i)).Picture = _
            frm.ImageList1.ListImages(Color).Picture
        frm.Imgbox(CurrentPosition(i)).Visible = True
     End If
Next i

End Sub
Public Sub SetCurrentPosition(a As Integer, _
    Optional b, Optional c, Optional d, Optional e, _
    Optional f)
    
    CurrentPosition(0) = a
    CurrentPosition(1) = b
    CurrentPosition(2) = c
    CurrentPosition(3) = d
    CurrentPosition(4) = e
    CurrentPosition(5) = f

End Sub
Public Sub SelectBrick()

If START = True Then
    
    Randomize
    
    If Brick = -1 Then
        Brick = Int(17 * Rnd)
    Else
        Brick = Brick1
    End If
    
    Color = Color1
        
    If Brick = 0 Then ' '.'
      Call SetCurrentPosition(45, -1, -1, -1, -1, -1)
'Purpose : It is used to Set the values in an Array
'          SetCurrentPosition i.e to set initial
'          position of the Brick in frm(frmMain Form).
'
'Alternate to call SetCurrentPosition()Funtion :
'        CurrentPosition(0) = 45
'        CurrentPosition(1) = -1
'        CurrentPosition(2) = -1
'        CurrentPosition(3) = -1
'        CurrentPosition(4) = -1
'        CurrentPosition(5) = -1
'
    ElseIf Brick = 1 Then ' ':'
      Call SetCurrentPosition(35, 45, -1, -1, -1, -1)
      
     ElseIf Brick = 2 Then ' '..'
       Call SetCurrentPosition(45, 46, -1, -1, -1, -1)
       
    ElseIf Brick = 3 Then ' '|'
      Call SetCurrentPosition(15, 25, 35, 45, -1, -1)
      
    ElseIf Brick = 4 Then ' "----"
      Call SetCurrentPosition(45, 46, 47, 48, -1, -1)
    
    ElseIf Brick = 5 Then ' '::'
      Call SetCurrentPosition(35, 36, 45, 46, -1, -1)
      
    ElseIf Brick = 6 Then ' '||'
      Call SetCurrentPosition(25, 26, 35, 36, 45, 46)
      
    ElseIf Brick = 7 Then ' ':::'
      Call SetCurrentPosition(35, 36, 37, 45, 46, 47)
      
    ElseIf Brick = 8 Then ' 'Z'
      Call SetCurrentPosition(35, 36, 46, 47, -1, -1)
      
    ElseIf Brick = 9 Then ' Rotate 'Z'
      Call SetCurrentPosition(26, 35, 36, 45, -1, -1)
      
    ElseIf Brick = 10 Then ' Inevrt 'T'
      Call SetCurrentPosition(36, 45, 46, 47, -1, -1)
      
    ElseIf Brick = 11 Then 'Rotate 'T' Left
      Call SetCurrentPosition(25, 35, 36, 45, -1, -1)
      
    ElseIf Brick = 12 Then ' 'T'
      Call SetCurrentPosition(35, 36, 37, 46, -1, -1)
      
    ElseIf Brick = 13 Then ' Rotate 'T' Right
      Call SetCurrentPosition(26, 35, 36, 46, -1, -1)
      
    ElseIf Brick = 14 Then ' 'L'
      Call SetCurrentPosition(25, 35, 45, 46, -1, -1)
      
    ElseIf Brick = 15 Then ' Invert 'L'
      Call SetCurrentPosition(25, 26, 36, 46, -1, -1)
        
    ElseIf Brick = 16 Then 'Rotate 'L' Right
      Call SetCurrentPosition(35, 36, 37, 45, -1, -1)
       
    ElseIf Brick = 17 Then 'Rotate 'L' Left
      Call SetCurrentPosition(37, 45, 46, 47, -1, -1)
      
    End If
    
    Call DrawBlock(frmMain)
'Purpose : It is used to Show Block i.e set the
'          selected Image Box visible according to the
'          position specified in SetCurrentPosition
'          function.
'
'Alternate to call DrawBlock()Funtion :
'To Show the Brick of type "." at position 55
'frmMain.Imgbox(17).Visible = True
    
    Brick1 = Int(17 * Rnd)
    
    If Color = 9 Then
        Color1 = 1
    Else
        Color1 = Color + 1
    End If
           
    Call ShowPreview(frmMain)
    RestoreNextBrick
    START = False
    Exit Sub

End If
End Sub

