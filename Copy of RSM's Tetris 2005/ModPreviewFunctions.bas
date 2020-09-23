Attribute VB_Name = "ModPreviewFunction"
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
' Date of Creation:  July 3,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.
'
' -------------------------------------
' --> A Module to Control Preview Screen
' -------------------------------------

Option Explicit


'Purpose :It is used to Show PicBox (PictureBox) in
' the frm (frmMain Form).
Public Sub ShowPicBox(frm As Form, a As Integer, _
    Optional b, Optional c, Optional d, Optional e, _
    Optional f)

If IsMissing(b) Then
    frm.PicBox(a).Visible = True
    
ElseIf IsMissing(c) Then
    frm.PicBox(a).Visible = True
    frm.PicBox(b).Visible = True

ElseIf IsMissing(d) Then
    frm.PicBox(a).Visible = True
    frm.PicBox(b).Visible = True
    frm.PicBox(c).Visible = True
    
ElseIf IsMissing(e) Then
    frm.PicBox(a).Visible = True
    frm.PicBox(b).Visible = True
    frm.PicBox(c).Visible = True
    frm.PicBox(d).Visible = True
    
ElseIf IsMissing(f) Then
    frm.PicBox(a).Visible = True
    frm.PicBox(b).Visible = True
    frm.PicBox(c).Visible = True
    frm.PicBox(d).Visible = True
    frm.PicBox(e).Visible = True
    
Else
    frm.PicBox(a).Visible = True
    frm.PicBox(b).Visible = True
    frm.PicBox(c).Visible = True
    frm.PicBox(d).Visible = True
    frm.PicBox(e).Visible = True
    frm.PicBox(f).Visible = True
End If

End Sub

'Purpose :It is used to Hide all PicBox (PictureBox)
' in the frm (frmMain Form).
Public Sub HidePicBox(frm As Form)

Dim i As Integer

For i = 0 To 41
    frm.PicBox(i).Visible = False
    frm.PicBox(i).Picture = _
    frm.ImageList1.ListImages(Color1).Picture
Next i

End Sub
Public Sub ShowPreview(frm As Form)

If frmOption.Check1.Value = 1 Then

    Call HidePicBox(frm)
'Purpose : It is used to Hide all PicBox (PictureBox)
'          in the frm (frmMain Form).

    If Brick1 = 0 Then ' '.'
        Call ShowPicBox(frm, 17)
'Purpose : It is used to Show PicBox (PictureBox) in
'          the frm (frmMain Form).
'
'Alternate to call ShowPicBox()Funtion :
'frmMain.PicBox(17).Visible = True
'
    ElseIf Brick1 = 1 Then ' ':'
        Call ShowPicBox(frm, 17, 24)

    ElseIf Brick1 = 2 Then ' '..'
        Call ShowPicBox(frm, 17, 18)
    
    ElseIf Brick1 = 3 Then ' '|'
        Call ShowPicBox(frm, 10, 17, 24, 31)
    
    ElseIf Brick1 = 4 Then ' "----"
        Call ShowPicBox(frm, 15, 16, 17, 18)
    
    ElseIf Brick1 = 5 Then ' '::'
        Call ShowPicBox(frm, 17, 18, 24, 25)
    
    ElseIf Brick1 = 6 Then ' '||'
        Call ShowPicBox(frm, 10, 11, 17, 18, 24, 25)
    
    ElseIf Brick1 = 7 Then ' ':::'
        Call ShowPicBox(frm, 16, 17, 18, 23, 24, 25)

    ElseIf Brick1 = 8 Then ' 'Z'
        Call ShowPicBox(frm, 16, 17, 24, 25)
    
    ElseIf Brick1 = 9 Then ' Rotate 'Z'
        Call ShowPicBox(frm, 10, 16, 17, 23)
    
    ElseIf Brick1 = 10 Then ' Inevrt 'T'
        Call ShowPicBox(frm, 17, 23, 24, 25)
    
    ElseIf Brick1 = 11 Then 'Rotate 'T' Left
        Call ShowPicBox(frm, 10, 17, 18, 24)

    ElseIf Brick1 = 12 Then ' 'T'
        Call ShowPicBox(frm, 16, 17, 18, 24)
    
    ElseIf Brick1 = 13 Then ' Rotate 'T' Right
        Call ShowPicBox(frm, 10, 16, 17, 24)
    
    ElseIf Brick1 = 14 Then ' 'L'
        Call ShowPicBox(frm, 10, 17, 24, 25)

    ElseIf Brick1 = 15 Then ' Invert 'L'
        Call ShowPicBox(frm, 9, 10, 17, 24)
    
    ElseIf Brick1 = 16 Then 'Rotate 'L' Right
        Call ShowPicBox(frm, 16, 17, 18, 23)

    ElseIf Brick1 = 17 Then 'Rotate 'L' Left
        Call ShowPicBox(frm, 18, 23, 24, 25)

    End If
End If

End Sub




