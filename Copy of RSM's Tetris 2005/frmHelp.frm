VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHelp.frx":0000
   ScaleHeight     =   4335
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   2400
      Picture         =   "frmHelp.frx":12C7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   480
      Picture         =   "frmHelp.frx":16B1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   3
      Left            =   1920
      Picture         =   "frmHelp.frx":1A91
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nikhil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rohit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apurv"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1100
      TabIndex        =   8
      Top             =   3120
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   810
      Index           =   0
      Left            =   240
      Picture         =   "frmHelp.frx":2F13
      Top             =   2160
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   990
      Index           =   1
      Left            =   960
      Picture         =   "frmHelp.frx":46F5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   1020
      Index           =   2
      Left            =   2640
      Picture         =   "frmHelp.frx":6D5F
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Label Label6 
      Caption         =   "................"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   870
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4200
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   240
      X2              =   4200
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Idea By :- Rohit Soam"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   2025
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Harish, Yashpal, Mohit         and  Soumitra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3200
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed By :- Mohit Soam"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEDICATED TO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label5 
      Height          =   2055
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
End
Attribute VB_Name = "frmHelp"
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
' Date of Creation:  July 5,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.
'
' -------------------------------------
' frmAbout - A Form to Display Help Window
' -------------------------------------

Option Explicit
Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command2_Click()
frmCreator.Show
End Sub

Private Sub Form_Load()
    GamePause = True
    frmMain.Timer1.Enabled = False
End Sub
