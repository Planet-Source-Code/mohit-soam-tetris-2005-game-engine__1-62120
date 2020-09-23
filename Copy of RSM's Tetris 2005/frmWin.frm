VERSION 5.00
Begin VB.Form frmWin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   6615
   ClientTop       =   1950
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Save Screenshot"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      DisabledPicture =   "frmWin.frx":0000
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   2520
      Picture         =   "frmWin.frx":00B2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   600
      TabIndex        =   10
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      DisabledPicture =   "frmWin.frx":0482
      Enabled         =   0   'False
      Height          =   1635
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter Your cool quote:"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   4800
      Width           =   1590
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Enter Your cool Name"
      Height          =   195
      Left            =   1440
      TabIndex        =   9
      Top             =   3960
      Width           =   1560
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Points"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   3000
      TabIndex        =   7
      Top             =   1680
      Width           =   645
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      Caption         =   "5000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tetris 2005 Hall of Shame"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "You have made the"
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   120
      X2              =   4800
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   120
      X2              =   4800
      Y1              =   1460
      Y2              =   1460
   End
   Begin VB.Label lblstop 
      Caption         =   "Congratulations !"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Your Score:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1320
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' Date of Creation:  July 6,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.
'
' -------------------------------------
' frmAbout - A Form to Display Win Window
' -------------------------------------

Option Explicit

Dim X As Integer
Dim Y As Integer
Dim px As Integer
Dim py As Integer

Private Sub Command2_Click()

frmScore.lblName(ScoreIndex).Caption = Me.Text1.Text
frmScore.lblquote(ScoreIndex).Caption = Me.Text2.Text

Call SaveScoreFile(frmScore)

Me.Hide
frmScore.Show

End Sub

Private Sub Command4_Click()

frmSS.PicSaveScreen.Cls
BitBlt frmSS.PicSaveScreen.hDC, 0, 0, Me.ScaleWidth, _
    Me.ScaleHeight, Me.hDC, 0, 0, vbSrcCopy
SavePicture frmSS.PicSaveScreen.Image, App.Path & "\HiScore.Bmp"

End Sub

Private Sub Form_Load()
Me.lblScore.Caption = frmMain.txtscore.Text
End Sub
