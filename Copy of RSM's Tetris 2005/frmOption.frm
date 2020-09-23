VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOption 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOption.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Skill Level (Constant Speed)"
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
      Begin VB.OptionButton Option3 
         Caption         =   "Expert"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Intermediate"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Beginner"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   2280
      Picture         =   "frmOption.frx":086D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   960
      Picture         =   "frmOption.frx":0C3D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   20
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Background Color :"
      Height          =   435
      Left            =   3240
      TabIndex        =   13
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Speed Skill Level (1-20) :"
      Height          =   435
      Left            =   2760
      TabIndex        =   10
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Postion Pointer"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Preview Screen"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmOption"
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
' frmAbout - A Form to Display Option Window
' -------------------------------------

Option Explicit
Private Sub Check1_Click()
If Check1.Value = 0 Then
    Call HidePicBox(frmMain)
Else
    Call ShowPreview(frmMain)
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
    Call HidePositionPointer(frmMain)
Else
    Call ShowPositionPointer(frmMain)
End If
End Sub

Private Sub Command1_Click()
If Check1.Value = 0 Then
    Call HidePicBox(frmMain)
End If
Me.Hide
End Sub

Private Sub Command2_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    GamePause = True
    frmMain.Timer1.Enabled = False
    lblColor.BackColor = frmMain.Frame1.BackColor
End Sub

Private Sub Label4_Click()

End Sub

Private Sub lblColor_Click()
    CommonDialog1.ShowColor
    lblColor.BackColor = CommonDialog1.Color
    frmMain.Frame1.BackColor = CommonDialog1.Color
End Sub

Private Sub Slider1_Change()
    TimeCount = 0
    frmMain.txtlevel = frmOption.Slider1.Value
    frmMain.Timer1.Interval = 1001 - _
        (frmOption.Slider1.Value * 50)
End Sub

Private Sub Slider1_Click()
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
End Sub
