Attribute VB_Name = "ModPosPointerFunction"
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
' Date of Creation:  July 4,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.
'
' -------------------------------------
' --> A Module to Control Postion Pointer
' -------------------------------------

Option Explicit
Dim i As Integer '====== used in For Loop

'Purpose : It is used to set the Postion of Postion
'          Pointer.
Public Sub SetPositionPointer()

For i = 0 To 5
    If CurrentPosition(i) <> -1 Then
        PositionPointer(i) = CurrentPosition(i) Mod 10
    Else
        PositionPointer(i) = -1
    End If
Next i

End Sub

'Purpose : It is used to Hide all PicBox(PictueBox)
'          in the frm (frmMain Form).
Public Sub HidePositionPointer(frm As Form)

For i = 0 To 9
    frm.lposition(i).Visible = False
Next i

End Sub

'Purpose : It is used to Show selected PicBox(PictueBox)
'          in the frm (frmMain Form).
Public Sub ShowPositionPointer(frm As Form)

If frmOption.Check2.Value = 1 Then

    Call SetPositionPointer

    Call HidePositionPointer(frmMain)

    For i = 0 To 5
        If PositionPointer(i) <> -1 Then
            frm.lposition(PositionPointer(i)).Visible = True
        End If
    Next i

End If

End Sub


