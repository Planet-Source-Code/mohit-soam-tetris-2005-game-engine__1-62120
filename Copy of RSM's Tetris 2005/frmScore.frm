VERSION 5.00
Begin VB.Form frmScore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "We proudly presents..."
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmScore.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   2880
      Picture         =   "frmScore.frx":0BDD
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblSno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   29
      Top             =   3150
      Width           =   45
   End
   Begin VB.Label lblSno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   2900
      Width           =   45
   End
   Begin VB.Label lblSno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   27
      Top             =   2650
      Width           =   45
   End
   Begin VB.Label lblSno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   26
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblSno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   25
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label lblSno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label lblSno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
      Height          =   195
      Index           =   4
      Left            =   5160
      TabIndex        =   22
      Top             =   2655
      Width           =   360
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "25000"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   21
      Top             =   2400
      Width           =   450
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   5160
      TabIndex        =   20
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "31050"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   5160
      TabIndex        =   19
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "32000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   5160
      TabIndex        =   18
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label lblquote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oh indeedly doodly!!"
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   17
      Top             =   2655
      Width           =   1440
   End
   Begin VB.Label lblquote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks so much, Sir!"
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   16
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblquote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ha Ha Ha Ha !!!t me!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   15
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblquote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I will come again!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   14
      Top             =   1680
      Width           =   1620
   End
   Begin VB.Label lblquote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You'll never get me!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   13
      Top             =   1320
      Width           =   1710
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mohit Soam"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   12
      Top             =   2655
      Width           =   840
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Soumitra Bhaumik"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   11
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nikhil Mehra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   10
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apurva Chawla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   1680
      Width           =   1290
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rohit Soam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tetris 2005 Hall of Shame"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abhishek Tyagi"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   6
      Top             =   2895
      Width           =   1095
   End
   Begin VB.Label lblquote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cool, man!e again!!!"
      Height          =   195
      Index           =   5
      Left            =   2400
      TabIndex        =   5
      Top             =   2895
      Width           =   1410
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1950"
      Height          =   195
      Index           =   5
      Left            =   5160
      TabIndex        =   4
      Top             =   2895
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anshuman Dhanai"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   3
      Top             =   3150
      Width           =   1305
   End
   Begin VB.Label lblquote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Better luck next time"
      Height          =   195
      Index           =   6
      Left            =   2400
      TabIndex        =   2
      Top             =   3150
      Width           =   1440
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      Height          =   195
      Index           =   6
      Left            =   5160
      TabIndex        =   1
      Top             =   3150
      Width           =   360
   End
End
Attribute VB_Name = "frmScore"
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
' frmAbout - A Form to Display Score Window
' -------------------------------------

Option Explicit

Private Sub Command1_Click()
Me.Hide
frmMain.Show
End Sub

Private Sub Form_Load()

GamePause = True
frmMain.Timer1.Enabled = False

Call LoadScoreFile(Me)

End Sub
