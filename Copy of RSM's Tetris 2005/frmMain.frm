VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   Caption         =   "RSM's Tetris 2005 "
   ClientHeight    =   6225
   ClientLeft      =   4320
   ClientTop       =   1470
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   0
      Left            =   2760
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   46
      Top             =   2880
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   7
      Left            =   2760
      Picture         =   "frmMain.frx":064C
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   45
      Top             =   3000
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   14
      Left            =   2760
      Picture         =   "frmMain.frx":098E
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   44
      Top             =   3120
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   21
      Left            =   2760
      Picture         =   "frmMain.frx":0CD0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   43
      Top             =   3240
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   28
      Left            =   2760
      Picture         =   "frmMain.frx":1012
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   42
      Top             =   3360
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   35
      Left            =   2760
      Picture         =   "frmMain.frx":1354
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   41
      Top             =   3480
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   1
      Left            =   2880
      Picture         =   "frmMain.frx":1696
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   40
      Top             =   2880
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   2
      Left            =   3000
      Picture         =   "frmMain.frx":19D8
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   39
      Top             =   2880
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   3
      Left            =   3120
      Picture         =   "frmMain.frx":1D1A
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   38
      Top             =   2880
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   4
      Left            =   3240
      Picture         =   "frmMain.frx":205C
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   37
      Top             =   2880
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   5
      Left            =   3360
      Picture         =   "frmMain.frx":239E
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   36
      Top             =   2880
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   6
      Left            =   3480
      Picture         =   "frmMain.frx":26E0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   35
      Top             =   2880
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   8
      Left            =   2880
      Picture         =   "frmMain.frx":2A22
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   34
      Top             =   3000
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   9
      Left            =   3000
      Picture         =   "frmMain.frx":2D64
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   33
      Top             =   3000
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   10
      Left            =   3120
      Picture         =   "frmMain.frx":30A6
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   32
      Top             =   3000
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   11
      Left            =   3240
      Picture         =   "frmMain.frx":33E8
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   31
      Top             =   3000
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   12
      Left            =   3360
      Picture         =   "frmMain.frx":372A
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   30
      Top             =   3000
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   13
      Left            =   3480
      Picture         =   "frmMain.frx":3A6C
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   29
      Top             =   3000
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   15
      Left            =   2880
      Picture         =   "frmMain.frx":3DAE
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   28
      Top             =   3120
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   16
      Left            =   3000
      Picture         =   "frmMain.frx":40F0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   27
      Top             =   3120
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   17
      Left            =   3120
      Picture         =   "frmMain.frx":4432
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   26
      Top             =   3120
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   18
      Left            =   3240
      Picture         =   "frmMain.frx":4774
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   25
      Top             =   3120
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   19
      Left            =   3360
      Picture         =   "frmMain.frx":4AB6
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   24
      Top             =   3120
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   20
      Left            =   3480
      Picture         =   "frmMain.frx":4DF8
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   23
      Top             =   3120
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   22
      Left            =   2880
      Picture         =   "frmMain.frx":513A
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   22
      Top             =   3240
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   23
      Left            =   3000
      Picture         =   "frmMain.frx":547C
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   21
      Top             =   3240
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   24
      Left            =   3120
      Picture         =   "frmMain.frx":57BE
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   20
      Top             =   3240
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   25
      Left            =   3240
      Picture         =   "frmMain.frx":5B00
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   19
      Top             =   3240
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   26
      Left            =   3360
      Picture         =   "frmMain.frx":5E42
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   18
      Top             =   3240
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   27
      Left            =   3480
      Picture         =   "frmMain.frx":6184
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   17
      Top             =   3240
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   29
      Left            =   2880
      Picture         =   "frmMain.frx":64C6
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   16
      Top             =   3360
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   30
      Left            =   3000
      Picture         =   "frmMain.frx":6808
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   15
      Top             =   3360
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   31
      Left            =   3120
      Picture         =   "frmMain.frx":6B4A
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   14
      Top             =   3360
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   32
      Left            =   3240
      Picture         =   "frmMain.frx":6E8C
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   13
      Top             =   3360
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   33
      Left            =   3360
      Picture         =   "frmMain.frx":71CE
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   3360
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   34
      Left            =   3480
      Picture         =   "frmMain.frx":7510
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   11
      Top             =   3360
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   36
      Left            =   2880
      Picture         =   "frmMain.frx":7852
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   10
      Top             =   3480
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   37
      Left            =   3000
      Picture         =   "frmMain.frx":7B94
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   9
      Top             =   3480
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   38
      Left            =   3120
      Picture         =   "frmMain.frx":7ED6
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   8
      Top             =   3480
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   39
      Left            =   3240
      Picture         =   "frmMain.frx":8218
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   7
      Top             =   3480
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   40
      Left            =   3360
      Picture         =   "frmMain.frx":855A
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   3480
      Width           =   100
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   41
      Left            =   3480
      Picture         =   "frmMain.frx":889C
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   3480
      Width           =   100
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   5040
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   4560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   4080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "                                               "
      Height          =   6135
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2425
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   194
         Left            =   960
         Picture         =   "frmMain.frx":8BDE
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   50
         Left            =   0
         Picture         =   "frmMain.frx":8F20
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   51
         Left            =   240
         Picture         =   "frmMain.frx":9262
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   52
         Left            =   480
         Picture         =   "frmMain.frx":95A4
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   53
         Left            =   720
         Picture         =   "frmMain.frx":98E6
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   54
         Left            =   960
         Picture         =   "frmMain.frx":9C28
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   55
         Left            =   1200
         Picture         =   "frmMain.frx":9F6A
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   56
         Left            =   1440
         Picture         =   "frmMain.frx":A2AC
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   57
         Left            =   1680
         Picture         =   "frmMain.frx":A5EE
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   58
         Left            =   1920
         Picture         =   "frmMain.frx":A930
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   59
         Left            =   2160
         Picture         =   "frmMain.frx":AC72
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   60
         Left            =   0
         Picture         =   "frmMain.frx":AFB4
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   61
         Left            =   240
         Picture         =   "frmMain.frx":B2F6
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   62
         Left            =   480
         Picture         =   "frmMain.frx":B638
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   63
         Left            =   720
         Picture         =   "frmMain.frx":B97A
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   64
         Left            =   960
         Picture         =   "frmMain.frx":BCBC
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   65
         Left            =   1200
         Picture         =   "frmMain.frx":BFFE
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   66
         Left            =   1440
         Picture         =   "frmMain.frx":C340
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   67
         Left            =   1680
         Picture         =   "frmMain.frx":C682
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   68
         Left            =   1920
         Picture         =   "frmMain.frx":C9C4
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   69
         Left            =   2160
         Picture         =   "frmMain.frx":CD06
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   70
         Left            =   0
         Picture         =   "frmMain.frx":D048
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   71
         Left            =   240
         Picture         =   "frmMain.frx":D38A
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   72
         Left            =   480
         Picture         =   "frmMain.frx":D6CC
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   73
         Left            =   720
         Picture         =   "frmMain.frx":DA0E
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   74
         Left            =   960
         Picture         =   "frmMain.frx":DD50
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   75
         Left            =   1200
         Picture         =   "frmMain.frx":E092
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   76
         Left            =   1440
         Picture         =   "frmMain.frx":E3D4
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   77
         Left            =   1680
         Picture         =   "frmMain.frx":E716
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   78
         Left            =   1920
         Picture         =   "frmMain.frx":EA58
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   79
         Left            =   2160
         Picture         =   "frmMain.frx":ED9A
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   80
         Left            =   0
         Picture         =   "frmMain.frx":F0DC
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   81
         Left            =   240
         Picture         =   "frmMain.frx":F41E
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   82
         Left            =   480
         Picture         =   "frmMain.frx":F760
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   83
         Left            =   720
         Picture         =   "frmMain.frx":FAA2
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   84
         Left            =   960
         Picture         =   "frmMain.frx":FDE4
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   85
         Left            =   1200
         Picture         =   "frmMain.frx":10126
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   86
         Left            =   1440
         Picture         =   "frmMain.frx":10468
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   87
         Left            =   1680
         Picture         =   "frmMain.frx":107AA
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   88
         Left            =   1920
         Picture         =   "frmMain.frx":10AEC
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   89
         Left            =   2160
         Picture         =   "frmMain.frx":10E2E
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   90
         Left            =   0
         Picture         =   "frmMain.frx":11170
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   91
         Left            =   240
         Picture         =   "frmMain.frx":114B2
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   92
         Left            =   480
         Picture         =   "frmMain.frx":117F4
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   93
         Left            =   720
         Picture         =   "frmMain.frx":11B36
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   94
         Left            =   960
         Picture         =   "frmMain.frx":11E78
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   95
         Left            =   1200
         Picture         =   "frmMain.frx":121BA
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   96
         Left            =   1440
         Picture         =   "frmMain.frx":124FC
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   97
         Left            =   1680
         Picture         =   "frmMain.frx":1283E
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   98
         Left            =   1920
         Picture         =   "frmMain.frx":12B80
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   99
         Left            =   2160
         Picture         =   "frmMain.frx":12EC2
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   100
         Left            =   0
         Picture         =   "frmMain.frx":13204
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   101
         Left            =   240
         Picture         =   "frmMain.frx":13546
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   102
         Left            =   480
         Picture         =   "frmMain.frx":13888
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   103
         Left            =   720
         Picture         =   "frmMain.frx":13BCA
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   104
         Left            =   960
         Picture         =   "frmMain.frx":13F0C
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   105
         Left            =   1200
         Picture         =   "frmMain.frx":1424E
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   106
         Left            =   1440
         Picture         =   "frmMain.frx":14590
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   107
         Left            =   1680
         Picture         =   "frmMain.frx":148D2
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   108
         Left            =   1920
         Picture         =   "frmMain.frx":14C14
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   109
         Left            =   2160
         Picture         =   "frmMain.frx":14F56
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   110
         Left            =   0
         Picture         =   "frmMain.frx":15298
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   111
         Left            =   240
         Picture         =   "frmMain.frx":155DA
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   112
         Left            =   480
         Picture         =   "frmMain.frx":1591C
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   113
         Left            =   720
         Picture         =   "frmMain.frx":15C5E
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   114
         Left            =   960
         Picture         =   "frmMain.frx":15FA0
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   115
         Left            =   1200
         Picture         =   "frmMain.frx":162E2
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   116
         Left            =   1440
         Picture         =   "frmMain.frx":16624
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   117
         Left            =   1680
         Picture         =   "frmMain.frx":16966
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   118
         Left            =   1920
         Picture         =   "frmMain.frx":16CA8
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   119
         Left            =   2160
         Picture         =   "frmMain.frx":16FEA
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   120
         Left            =   0
         Picture         =   "frmMain.frx":1732C
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   121
         Left            =   240
         Picture         =   "frmMain.frx":1766E
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   122
         Left            =   480
         Picture         =   "frmMain.frx":179B0
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   123
         Left            =   720
         Picture         =   "frmMain.frx":17CF2
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   124
         Left            =   960
         Picture         =   "frmMain.frx":18034
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   125
         Left            =   1200
         Picture         =   "frmMain.frx":18376
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   126
         Left            =   1440
         Picture         =   "frmMain.frx":186B8
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   127
         Left            =   1680
         Picture         =   "frmMain.frx":189FA
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   128
         Left            =   1920
         Picture         =   "frmMain.frx":18D3C
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   129
         Left            =   2160
         Picture         =   "frmMain.frx":1907E
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   130
         Left            =   0
         Picture         =   "frmMain.frx":193C0
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   131
         Left            =   240
         Picture         =   "frmMain.frx":19702
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   132
         Left            =   480
         Picture         =   "frmMain.frx":19A44
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   133
         Left            =   720
         Picture         =   "frmMain.frx":19D86
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   134
         Left            =   960
         Picture         =   "frmMain.frx":1A0C8
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   135
         Left            =   1200
         Picture         =   "frmMain.frx":1A40A
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   136
         Left            =   1440
         Picture         =   "frmMain.frx":1A74C
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   137
         Left            =   1680
         Picture         =   "frmMain.frx":1AA8E
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   138
         Left            =   1920
         Picture         =   "frmMain.frx":1ADD0
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   139
         Left            =   2160
         Picture         =   "frmMain.frx":1B112
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   140
         Left            =   0
         Picture         =   "frmMain.frx":1B454
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   141
         Left            =   240
         Picture         =   "frmMain.frx":1B796
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   142
         Left            =   480
         Picture         =   "frmMain.frx":1BAD8
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   143
         Left            =   720
         Picture         =   "frmMain.frx":1BE1A
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   144
         Left            =   960
         Picture         =   "frmMain.frx":1C15C
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   145
         Left            =   1200
         Picture         =   "frmMain.frx":1C49E
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   146
         Left            =   1440
         Picture         =   "frmMain.frx":1C7E0
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   147
         Left            =   1680
         Picture         =   "frmMain.frx":1CB22
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   148
         Left            =   1920
         Picture         =   "frmMain.frx":1CE64
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   149
         Left            =   2160
         Picture         =   "frmMain.frx":1D1A6
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   150
         Left            =   0
         Picture         =   "frmMain.frx":1D4E8
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   151
         Left            =   240
         Picture         =   "frmMain.frx":1D82A
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   152
         Left            =   480
         Picture         =   "frmMain.frx":1DB6C
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   153
         Left            =   720
         Picture         =   "frmMain.frx":1DEAE
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   154
         Left            =   960
         Picture         =   "frmMain.frx":1E1F0
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   155
         Left            =   1200
         Picture         =   "frmMain.frx":1E532
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   156
         Left            =   1440
         Picture         =   "frmMain.frx":1E874
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   157
         Left            =   1680
         Picture         =   "frmMain.frx":1EBB6
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   158
         Left            =   1920
         Picture         =   "frmMain.frx":1EEF8
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   159
         Left            =   2160
         Picture         =   "frmMain.frx":1F23A
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   160
         Left            =   0
         Picture         =   "frmMain.frx":1F57C
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   161
         Left            =   240
         Picture         =   "frmMain.frx":1F8BE
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   162
         Left            =   480
         Picture         =   "frmMain.frx":1FC00
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   163
         Left            =   720
         Picture         =   "frmMain.frx":1FF42
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   164
         Left            =   960
         Picture         =   "frmMain.frx":20284
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   165
         Left            =   1200
         Picture         =   "frmMain.frx":205C6
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   166
         Left            =   1440
         Picture         =   "frmMain.frx":20908
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   167
         Left            =   1680
         Picture         =   "frmMain.frx":20C4A
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   168
         Left            =   1920
         Picture         =   "frmMain.frx":20F8C
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   169
         Left            =   2160
         Picture         =   "frmMain.frx":212CE
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   170
         Left            =   0
         Picture         =   "frmMain.frx":21610
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   171
         Left            =   240
         Picture         =   "frmMain.frx":21952
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   172
         Left            =   480
         Picture         =   "frmMain.frx":21C94
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   173
         Left            =   720
         Picture         =   "frmMain.frx":21FD6
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   174
         Left            =   960
         Picture         =   "frmMain.frx":22318
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   175
         Left            =   1200
         Picture         =   "frmMain.frx":2265A
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   176
         Left            =   1440
         Picture         =   "frmMain.frx":2299C
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   177
         Left            =   1680
         Picture         =   "frmMain.frx":22CDE
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   178
         Left            =   1920
         Picture         =   "frmMain.frx":23020
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   179
         Left            =   2160
         Picture         =   "frmMain.frx":23362
         Top             =   3000
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   180
         Left            =   0
         Picture         =   "frmMain.frx":236A4
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   181
         Left            =   240
         Picture         =   "frmMain.frx":239E6
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   182
         Left            =   480
         Picture         =   "frmMain.frx":23D28
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   183
         Left            =   720
         Picture         =   "frmMain.frx":2406A
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   184
         Left            =   960
         Picture         =   "frmMain.frx":243AC
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   185
         Left            =   1200
         Picture         =   "frmMain.frx":246EE
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   186
         Left            =   1440
         Picture         =   "frmMain.frx":24A30
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   187
         Left            =   1680
         Picture         =   "frmMain.frx":24D72
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   188
         Left            =   1920
         Picture         =   "frmMain.frx":250B4
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   189
         Left            =   2160
         Picture         =   "frmMain.frx":253F6
         Top             =   3240
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   190
         Left            =   0
         Picture         =   "frmMain.frx":25738
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   191
         Left            =   240
         Picture         =   "frmMain.frx":25A7A
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   192
         Left            =   480
         Picture         =   "frmMain.frx":25DBC
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   193
         Left            =   720
         Picture         =   "frmMain.frx":260FE
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   195
         Left            =   1200
         Picture         =   "frmMain.frx":26440
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   196
         Left            =   1440
         Picture         =   "frmMain.frx":26782
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   197
         Left            =   1680
         Picture         =   "frmMain.frx":26AC4
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   198
         Left            =   1920
         Picture         =   "frmMain.frx":26E06
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   199
         Left            =   2160
         Picture         =   "frmMain.frx":27148
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   200
         Left            =   0
         Picture         =   "frmMain.frx":2748A
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   201
         Left            =   240
         Picture         =   "frmMain.frx":277CC
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   202
         Left            =   480
         Picture         =   "frmMain.frx":27B0E
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   203
         Left            =   720
         Picture         =   "frmMain.frx":27E50
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   204
         Left            =   960
         Picture         =   "frmMain.frx":28192
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   205
         Left            =   1200
         Picture         =   "frmMain.frx":284D4
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   206
         Left            =   1440
         Picture         =   "frmMain.frx":28816
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   207
         Left            =   1680
         Picture         =   "frmMain.frx":28B58
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   208
         Left            =   1920
         Picture         =   "frmMain.frx":28E9A
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   209
         Left            =   2160
         Picture         =   "frmMain.frx":291DC
         Top             =   3720
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   210
         Left            =   0
         Picture         =   "frmMain.frx":2951E
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   211
         Left            =   240
         Picture         =   "frmMain.frx":29860
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   212
         Left            =   480
         Picture         =   "frmMain.frx":29BA2
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   213
         Left            =   720
         Picture         =   "frmMain.frx":29EE4
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   214
         Left            =   960
         Picture         =   "frmMain.frx":2A226
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   215
         Left            =   1200
         Picture         =   "frmMain.frx":2A568
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   216
         Left            =   1440
         Picture         =   "frmMain.frx":2A8AA
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   217
         Left            =   1680
         Picture         =   "frmMain.frx":2ABEC
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   218
         Left            =   1920
         Picture         =   "frmMain.frx":2AF2E
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   219
         Left            =   2160
         Picture         =   "frmMain.frx":2B270
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   220
         Left            =   0
         Picture         =   "frmMain.frx":2B5B2
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   221
         Left            =   240
         Picture         =   "frmMain.frx":2B8F4
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   222
         Left            =   480
         Picture         =   "frmMain.frx":2BC36
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   223
         Left            =   720
         Picture         =   "frmMain.frx":2BF78
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   224
         Left            =   960
         Picture         =   "frmMain.frx":2C2BA
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   225
         Left            =   1200
         Picture         =   "frmMain.frx":2C5FC
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   226
         Left            =   1440
         Picture         =   "frmMain.frx":2C93E
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   227
         Left            =   1680
         Picture         =   "frmMain.frx":2CC80
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   228
         Left            =   1920
         Picture         =   "frmMain.frx":2CFC2
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   229
         Left            =   2160
         Picture         =   "frmMain.frx":2D304
         Top             =   4200
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   230
         Left            =   0
         Picture         =   "frmMain.frx":2D646
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   231
         Left            =   240
         Picture         =   "frmMain.frx":2D988
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   232
         Left            =   480
         Picture         =   "frmMain.frx":2DCCA
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   233
         Left            =   720
         Picture         =   "frmMain.frx":2E00C
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   234
         Left            =   960
         Picture         =   "frmMain.frx":2E34E
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   235
         Left            =   1200
         Picture         =   "frmMain.frx":2E690
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   236
         Left            =   1440
         Picture         =   "frmMain.frx":2E9D2
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   237
         Left            =   1680
         Picture         =   "frmMain.frx":2ED14
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   238
         Left            =   1920
         Picture         =   "frmMain.frx":2F056
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   239
         Left            =   2160
         Picture         =   "frmMain.frx":2F398
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   240
         Left            =   0
         Picture         =   "frmMain.frx":2F6DA
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   241
         Left            =   240
         Picture         =   "frmMain.frx":2FA1C
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   242
         Left            =   480
         Picture         =   "frmMain.frx":2FD5E
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   243
         Left            =   720
         Picture         =   "frmMain.frx":300A0
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   244
         Left            =   960
         Picture         =   "frmMain.frx":303E2
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   245
         Left            =   1200
         Picture         =   "frmMain.frx":30724
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   246
         Left            =   1440
         Picture         =   "frmMain.frx":30A66
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   247
         Left            =   1680
         Picture         =   "frmMain.frx":30DA8
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   248
         Left            =   1920
         Picture         =   "frmMain.frx":310EA
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   249
         Left            =   2160
         Picture         =   "frmMain.frx":3142C
         Top             =   4680
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   250
         Left            =   0
         Picture         =   "frmMain.frx":3176E
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   251
         Left            =   240
         Picture         =   "frmMain.frx":31AB0
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   252
         Left            =   480
         Picture         =   "frmMain.frx":31DF2
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   253
         Left            =   720
         Picture         =   "frmMain.frx":32134
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   254
         Left            =   960
         Picture         =   "frmMain.frx":32476
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   255
         Left            =   1200
         Picture         =   "frmMain.frx":327B8
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   256
         Left            =   1440
         Picture         =   "frmMain.frx":32AFA
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   257
         Left            =   1680
         Picture         =   "frmMain.frx":32E3C
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   258
         Left            =   1920
         Picture         =   "frmMain.frx":3317E
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   259
         Left            =   2160
         Picture         =   "frmMain.frx":334C0
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   260
         Left            =   0
         Picture         =   "frmMain.frx":33802
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   261
         Left            =   240
         Picture         =   "frmMain.frx":33B44
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   262
         Left            =   480
         Picture         =   "frmMain.frx":33E86
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   263
         Left            =   720
         Picture         =   "frmMain.frx":341C8
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   264
         Left            =   960
         Picture         =   "frmMain.frx":3450A
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   265
         Left            =   1200
         Picture         =   "frmMain.frx":3484C
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   266
         Left            =   1440
         Picture         =   "frmMain.frx":34B8E
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   267
         Left            =   1680
         Picture         =   "frmMain.frx":34ED0
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   268
         Left            =   1920
         Picture         =   "frmMain.frx":35212
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   269
         Left            =   2160
         Picture         =   "frmMain.frx":35554
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   270
         Left            =   0
         Picture         =   "frmMain.frx":35896
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   271
         Left            =   240
         Picture         =   "frmMain.frx":35BD8
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   272
         Left            =   480
         Picture         =   "frmMain.frx":35F1A
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   273
         Left            =   720
         Picture         =   "frmMain.frx":3625C
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   274
         Left            =   960
         Picture         =   "frmMain.frx":3659E
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   275
         Left            =   1200
         Picture         =   "frmMain.frx":368E0
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   276
         Left            =   1440
         Picture         =   "frmMain.frx":36C22
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   277
         Left            =   1680
         Picture         =   "frmMain.frx":36F64
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   278
         Left            =   1920
         Picture         =   "frmMain.frx":372A6
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   279
         Left            =   2160
         Picture         =   "frmMain.frx":375E8
         Top             =   5400
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   280
         Left            =   0
         Picture         =   "frmMain.frx":3792A
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   281
         Left            =   240
         Picture         =   "frmMain.frx":37C6C
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   282
         Left            =   480
         Picture         =   "frmMain.frx":37FAE
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   283
         Left            =   720
         Picture         =   "frmMain.frx":382F0
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   284
         Left            =   960
         Picture         =   "frmMain.frx":38632
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   285
         Left            =   1200
         Picture         =   "frmMain.frx":38974
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   286
         Left            =   1440
         Picture         =   "frmMain.frx":38CB6
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   287
         Left            =   1680
         Picture         =   "frmMain.frx":38FF8
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   288
         Left            =   1920
         Picture         =   "frmMain.frx":3933A
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   289
         Left            =   2160
         Picture         =   "frmMain.frx":3967C
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   290
         Left            =   0
         Picture         =   "frmMain.frx":399BE
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   291
         Left            =   240
         Picture         =   "frmMain.frx":39D00
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   292
         Left            =   480
         Picture         =   "frmMain.frx":3A042
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   293
         Left            =   720
         Picture         =   "frmMain.frx":3A384
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   294
         Left            =   960
         Picture         =   "frmMain.frx":3A6C6
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   295
         Left            =   1200
         Picture         =   "frmMain.frx":3AA08
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   296
         Left            =   1440
         Picture         =   "frmMain.frx":3AD4A
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   297
         Left            =   1680
         Picture         =   "frmMain.frx":3B08C
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   298
         Left            =   1920
         Picture         =   "frmMain.frx":3B3CE
         Top             =   5880
         Width           =   240
      End
      Begin VB.Image Imgbox 
         Height          =   240
         Index           =   299
         Left            =   2160
         Picture         =   "frmMain.frx":3B710
         Top             =   5880
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BA52
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BDA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C0FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C44E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C7A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CAF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CE4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D1A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D4FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   3600
   End
   Begin VB.TextBox txtscore 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtlevel 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Image Imgpausedis 
      Height          =   345
      Left            =   2760
      Picture         =   "frmMain.frx":3DBAA
      Top             =   480
      Width           =   840
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   2745
      Top             =   2865
      Width           =   855
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   49
      Left            =   2280
      Picture         =   "frmMain.frx":3DFD3
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   48
      Left            =   2040
      Picture         =   "frmMain.frx":3E315
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   47
      Left            =   1800
      Picture         =   "frmMain.frx":3E657
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   46
      Left            =   1560
      Picture         =   "frmMain.frx":3E999
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   45
      Left            =   1320
      Picture         =   "frmMain.frx":3ECDB
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   44
      Left            =   1080
      Picture         =   "frmMain.frx":3F01D
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   43
      Left            =   840
      Picture         =   "frmMain.frx":3F35F
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   42
      Left            =   600
      Picture         =   "frmMain.frx":3F6A1
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   41
      Left            =   360
      Picture         =   "frmMain.frx":3F9E3
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   40
      Left            =   120
      Picture         =   "frmMain.frx":3FD25
      Top             =   1080
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   33
      Left            =   840
      Picture         =   "frmMain.frx":40067
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   32
      Left            =   600
      Picture         =   "frmMain.frx":403A9
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   31
      Left            =   360
      Picture         =   "frmMain.frx":406EB
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   30
      Left            =   120
      Picture         =   "frmMain.frx":40A2D
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   23
      Left            =   840
      Picture         =   "frmMain.frx":40D6F
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   22
      Left            =   600
      Picture         =   "frmMain.frx":410B1
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   21
      Left            =   360
      Picture         =   "frmMain.frx":413F3
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   20
      Left            =   120
      Picture         =   "frmMain.frx":41735
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   39
      Left            =   2280
      Picture         =   "frmMain.frx":41A77
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   38
      Left            =   2040
      Picture         =   "frmMain.frx":41DB9
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   37
      Left            =   1800
      Picture         =   "frmMain.frx":420FB
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   36
      Left            =   1560
      Picture         =   "frmMain.frx":4243D
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   35
      Left            =   1320
      Picture         =   "frmMain.frx":4277F
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   34
      Left            =   1080
      Picture         =   "frmMain.frx":42AC1
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   29
      Left            =   2280
      Picture         =   "frmMain.frx":42E03
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   28
      Left            =   2040
      Picture         =   "frmMain.frx":43145
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   27
      Left            =   1800
      Picture         =   "frmMain.frx":43487
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   26
      Left            =   1560
      Picture         =   "frmMain.frx":437C9
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   25
      Left            =   1320
      Picture         =   "frmMain.frx":43B0B
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   24
      Left            =   1080
      Picture         =   "frmMain.frx":43E4D
      Top             =   600
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   19
      Left            =   2280
      Picture         =   "frmMain.frx":4418F
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   18
      Left            =   2040
      Picture         =   "frmMain.frx":444D1
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   17
      Left            =   1800
      Picture         =   "frmMain.frx":44813
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   16
      Left            =   1560
      Picture         =   "frmMain.frx":44B55
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   15
      Left            =   1320
      Picture         =   "frmMain.frx":44E97
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   14
      Left            =   1080
      Picture         =   "frmMain.frx":451D9
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   13
      Left            =   840
      Picture         =   "frmMain.frx":4551B
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   12
      Left            =   600
      Picture         =   "frmMain.frx":4585D
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   11
      Left            =   360
      Picture         =   "frmMain.frx":45B9F
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   10
      Left            =   120
      Picture         =   "frmMain.frx":45EE1
      Top             =   360
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   9
      Left            =   2280
      Picture         =   "frmMain.frx":46223
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   8
      Left            =   2040
      Picture         =   "frmMain.frx":46565
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   7
      Left            =   1800
      Picture         =   "frmMain.frx":468A7
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   6
      Left            =   1560
      Picture         =   "frmMain.frx":46BE9
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   5
      Left            =   1320
      Picture         =   "frmMain.frx":46F2B
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   4
      Left            =   1080
      Picture         =   "frmMain.frx":4726D
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":475AF
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":478F1
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":47C33
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Imgbox 
      Height          =   240
      Index           =   3
      Left            =   840
      Picture         =   "frmMain.frx":47F75
      Top             =   120
      Width           =   240
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   9
      X1              =   2280
      X2              =   2520
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   8
      X1              =   2040
      X2              =   2280
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   7
      X1              =   1800
      X2              =   2040
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   6
      X1              =   1560
      X2              =   1800
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   5
      X1              =   1320
      X2              =   1560
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   4
      X1              =   1080
      X2              =   1320
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   840
      X2              =   1080
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   600
      X2              =   840
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   360
      X2              =   600
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line lposition 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   360
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.Image imgbutton 
      Height          =   360
      Index           =   9
      Left            =   3000
      Picture         =   "frmMain.frx":482B7
      Top             =   4065
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgbutton 
      Height          =   360
      Index           =   8
      Left            =   3360
      Picture         =   "frmMain.frx":48642
      Top             =   3705
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgbutton 
      Height          =   360
      Index           =   7
      Left            =   3000
      Picture         =   "frmMain.frx":489CF
      Top             =   3705
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgbutton 
      Height          =   360
      Index           =   6
      Left            =   2640
      Picture         =   "frmMain.frx":48D61
      Top             =   3705
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Line LT2 
      BorderColor     =   &H80000003&
      X1              =   4680
      X2              =   5880
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Line LT1 
      BorderColor     =   &H80000005&
      X1              =   4680
      X2              =   5880
      Y1              =   1560
      Y2              =   2040
   End
   Begin VB.Line LL2 
      BorderColor     =   &H80000003&
      X1              =   4800
      X2              =   6000
      Y1              =   1800
      Y2              =   2280
   End
   Begin VB.Line LL1 
      BorderColor     =   &H80000005&
      X1              =   4680
      X2              =   5880
      Y1              =   1440
      Y2              =   1920
   End
   Begin VB.Image imgbutton 
      Height          =   420
      Index           =   5
      Left            =   2640
      Picture         =   "frmMain.frx":490ED
      ToolTipText     =   "Exit Biricks '2005"
      Top             =   2400
      Width           =   990
   End
   Begin VB.Image imgbutton 
      Height          =   420
      Index           =   4
      Left            =   2685
      Picture         =   "frmMain.frx":49512
      ToolTipText     =   "Show help.."
      Top             =   1920
      Width           =   990
   End
   Begin VB.Image imgbutton 
      Height          =   420
      Index           =   3
      Left            =   2685
      Picture         =   "frmMain.frx":4990D
      ToolTipText     =   "Highscore history.."
      Top             =   1440
      Width           =   990
   End
   Begin VB.Image imgbutton 
      Height          =   420
      Index           =   2
      Left            =   2685
      Picture         =   "frmMain.frx":49D23
      ToolTipText     =   "Game options.."
      Top             =   960
      Width           =   990
   End
   Begin VB.Image imgbutton 
      Height          =   420
      Index           =   1
      Left            =   2685
      Picture         =   "frmMain.frx":4A147
      ToolTipText     =   "Pause game to take breath..."
      Top             =   480
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgbutton 
      Height          =   420
      Index           =   0
      Left            =   2685
      Picture         =   "frmMain.frx":4A564
      ToolTipText     =   "Start a new game"
      Top             =   0
      Width           =   990
   End
   Begin VB.Image Imgmousebutdis 
      Height          =   675
      Left            =   2640
      Picture         =   "frmMain.frx":4A982
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Image Imgstartdis 
      Height          =   360
      Left            =   2760
      Picture         =   "frmMain.frx":4ADB3
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuquit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuoption 
      Caption         =   "&Option"
   End
   Begin VB.Menu mnuscores 
      Caption         =   "&Scores"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnudesigner 
         Caption         =   "About &Designer "
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About &Bricks 2005"
      End
   End
End
Attribute VB_Name = "frmMain"
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
' Date of Creation:  July 1,2005
' Last Update:  August 1,2005
'
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.
'
' -------------------------------------
' frmMain - A Form to Display Main Game Window
' -------------------------------------

Option Explicit

Private Sub Form_Load()
Dim i As Integer

For i = 0 To 299
    Imgbox(i).Visible = False
Next i

For i = 0 To 41
    PicBox(i).Visible = False
Next i

' Here we do some Initial Setting before Starting a
' New Game.
START = True
GamePause = True
LeftBrick = False
RightBrick = False
AVILABLE = False
GAME_OVER = False

txtscore.Text = 0
Me.txtlevel = 1

BrickCount = 0
Level = 1
Brick = -1
Color = 1
Color1 = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        X = MsgBox("(  PLEASE 'RATE' THIS CODE  ).I want to know how do you rate this code.", vbInformation + vbOKCancel, "ThankYou")
    Else
        X = MsgBox("( PLEASE GIVE FEEDBACK ) to improve this code.", vbInformation + vbOKCancel, "Please Give FeedBack")
    End If

End

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If GamePause = True Then Exit Sub

If KeyCode = 40 Then
    Timer1.Interval = 1
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    
    Exit Sub

ElseIf KeyCode = 37 Then
    Increment (-1)
    
    If LeftBrick = True Then
        Exit Sub
    End If

ElseIf KeyCode = 39 Then
    Increment (1)
    
    If RightBrick = True Then
        Exit Sub
    End If

End If

If Brick = 0 Then ' '.'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0)
    
    ElseIf KeyCode = 38 Then
        Exit Sub
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 0)
    
    End If

ElseIf Brick = 1 Then ' ':'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 1)
    
    ElseIf KeyCode = 38 Then
        
        Call HideBrick(frmMain)
        
        If CurrentPosition(0) Mod 10 <> 9 And _
            NextPosition(0) <> -1 Then
            NextPosition(1) = NextPosition(0) + 1
        
            Brick = 2
        End If
        
        Call CheckStatus(10, 0, 1)
        Call DrawBlock(frmMain)
        
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 0, 1)
    
    End If
    
ElseIf Brick = 2 Then ' '..'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0)
    
    ElseIf KeyCode = 38 Then
        
        Call HideBrick(frmMain)
        
        If NextPosition(0) <> -1 Then
            NextPosition(1) = NextPosition(0) + 10
        End If
        
        Call CheckStatus(10, 1)
        Call DrawBlock(frmMain)
        
        Brick = 1
        
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 1)
    
    End If
    
ElseIf Brick = 3 Then ' '|'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 1, 2, 3)
    
    ElseIf KeyCode = 38 Then
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 7 And _
            NextPosition(0) <> -1 Then
            NextPosition(1) = NextPosition(0) + 1
            NextPosition(2) = NextPosition(0) + 2
            NextPosition(3) = NextPosition(0) + 3
            
            Brick = 4
        End If
        
        Call CheckStatus(10, 0, 1, 2, 3)
        Call DrawBlock(frmMain)
        
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 0, 1, 2, 3)
    
    End If
    
ElseIf Brick = 4 Then ' "----"
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) <> -1 Then
        NextPosition(1) = NextPosition(0) + 10
        NextPosition(2) = NextPosition(0) + 20
        NextPosition(3) = NextPosition(0) + 30
        
        End If
        
        Call CheckStatus(10, 3)
        Call DrawBlock(frmMain)
        
        Brick = 3
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 3)
    
    End If
    
ElseIf Brick = 5 Then ' '::'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 2)
    
    ElseIf KeyCode = 38 Then
        Exit Sub
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 1, 3)
    
    End If
    
ElseIf Brick = 6 Then ' '||'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 2, 4)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
           
        If NextPosition(0) Mod 10 < 8 And _
            NextPosition(0) <> -1 Then
            NextPosition(2) = NextPosition(0) + 2
            NextPosition(3) = NextPosition(0) + 10
            NextPosition(4) = NextPosition(0) + 11
            NextPosition(5) = NextPosition(0) + 12
        
            Brick = 7
        End If
        
        Call CheckStatus(10, 3, 4, 5)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 1, 3, 5)
    
    End If
    
ElseIf Brick = 7 Then ' ':::'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 3)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 8 And _
            NextPosition(0) <> -1 Then
            NextPosition(2) = NextPosition(0) + 10
            NextPosition(3) = NextPosition(0) + 11
            NextPosition(4) = NextPosition(0) + 20
            NextPosition(5) = NextPosition(0) + 21
        
            Brick = 6
        End If
        
        Call CheckStatus(10, 4, 5)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 2, 5)
    End If
    
ElseIf Brick = 8 Then ' 'Z'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 8 And _
            NextPosition(0) <> -1 Then
            NextPosition(0) = NextPosition(0) + 1
            NextPosition(1) = NextPosition(0) + 9
            NextPosition(3) = NextPosition(0) + 19
        
            Brick = 9
         End If
        
        Call CheckStatus(10, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 3)
    
    End If
    
ElseIf Brick = 9 Then ' Rotate 'Z'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 1)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 9 And _
            NextPosition(0) <> -1 Then
            NextPosition(0) = NextPosition(0) - 1
            NextPosition(1) = NextPosition(0) + 1
            NextPosition(3) = NextPosition(0) + 12
            
            Brick = 8
         End If
        
        Call CheckStatus(10, 0, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 2, 3)
    
    End If
    
ElseIf Brick = 10 Then ' Inevrt 'T'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 1)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) <> -1 Then
            NextPosition(0) = NextPosition(0) - 1
            NextPosition(3) = NextPosition(0) + 20
        
            Brick = 11
        End If
        
        Call CheckStatus(10, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 3)
    
    End If
    
ElseIf Brick = 11 Then 'Rotate 'T' Left
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 1, 3)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 8 And _
            NextPosition(0) <> -1 Then
            NextPosition(1) = NextPosition(0) + 1
            NextPosition(2) = NextPosition(0) + 2
            NextPosition(3) = NextPosition(0) + 11
            
            Brick = 12
        End If
        
        Call CheckStatus(10, 0, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 2)
    
    End If
    
ElseIf Brick = 12 Then ' 'T'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 3)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) <> -1 Then
            NextPosition(0) = NextPosition(0) + 1
            NextPosition(1) = NextPosition(0) + 9
            NextPosition(2) = NextPosition(0) + 10
            NextPosition(3) = NextPosition(0) + 20
        
            Brick = 13
        End If
        
        Call CheckStatus(10, 1, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 2, 3)
    
    End If
    
ElseIf Brick = 13 Then 'Rotate 'T' Right
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 1, 3)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 9 And _
        NextPosition(0) <> -1 Then
            NextPosition(3) = NextPosition(0) + 11
            
            Brick = 10
        End If
        
        Call CheckStatus(10, 1, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 0, 2, 3)
    
    End If
    
ElseIf Brick = 14 Then ' 'L'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 1, 2)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 8 And _
            NextPosition(0) <> -1 Then
            NextPosition(0) = NextPosition(0) + 2
            NextPosition(1) = NextPosition(0) + 8
            NextPosition(2) = NextPosition(0) + 9
            NextPosition(3) = NextPosition(0) + 10
        
            Brick = 17
        End If
        
        Call CheckStatus(10, 1, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 0, 1, 3)
    
    End If
    
ElseIf Brick = 15 Then ' Invert 'L'
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 2, 3)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) Mod 10 < 8 And _
            NextPosition(0) <> -1 Then
            NextPosition(2) = NextPosition(0) + 2
            NextPosition(3) = NextPosition(0) + 10
            
            Brick = 16
        End If
        
        Call CheckStatus(10, 1, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 1, 2, 3)
    
    End If
    
ElseIf Brick = 16 Then 'Rotate 'L' Right
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 3)
    
    ElseIf KeyCode = 38 Then
        
         Call HideBrick(frmMain)
        
        If NextPosition(0) <> -1 Then
            NextPosition(1) = NextPosition(0) + 10
            NextPosition(2) = NextPosition(0) + 20
            NextPosition(3) = NextPosition(0) + 21
        
            Brick = 14
        End If
        
        Call CheckStatus(10, 2, 3)
        Call DrawBlock(frmMain)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 2, 3)
    
    End If

ElseIf Brick = 17 Then 'Rotate 'L' Left
    If KeyCode = 37 Then
        Call CheckStatus(-1, 0, 1)
    
    ElseIf KeyCode = 38 Then
         Call HideBrick(frmMain)
        
        If NextPosition(0) <> -1 Then
            NextPosition(0) = NextPosition(0) - 2
            NextPosition(1) = NextPosition(0) + 1
            NextPosition(2) = NextPosition(0) + 11
            NextPosition(3) = NextPosition(0) + 21
        
            Brick = 15
        End If
        
        Call CheckStatus(10, 0, 3)
    
    ElseIf KeyCode = 39 Then
        Call CheckStatus(1, 0, 3)
    
    End If

End If

MoveBlock

End Sub

Private Sub mnuabout_Click()
    Call imgbutton_MouseDown(4, 1, 0, 645, 280)
End Sub

Private Sub mnudesigner_Click()
    frmCreator.Show
End Sub

Private Sub mnuoption_Click()
    Call imgbutton_MouseDown(2, 1, 0, 645, 315)
End Sub

Private Sub mnuquit_Click()
    Call Form_Unload(0)
End Sub

Private Sub mnuscores_Click()
    Call imgbutton_MouseDown(3, 1, 0, 705, 150)
End Sub

Private Sub Timer1_Timer()

frmOptionSetting

Increment (10)

If Brick = 0 Then ' '.'
    Call CheckStatus(10, 0)
    
ElseIf Brick = 1 Then ' ':'
    Call CheckStatus(10, 1)
    
ElseIf Brick = 2 Then ' '..'
    Call CheckStatus(10, 0, 1)
    
ElseIf Brick = 3 Then ' '|'
    Call CheckStatus(10, 3)
    
ElseIf Brick = 4 Then ' "----"
    Call CheckStatus(10, 0, 1, 2, 3)
    
ElseIf Brick = 5 Then ' '::'
    Call CheckStatus(10, 2, 3)
    
ElseIf Brick = 6 Then ' '||'
    Call CheckStatus(10, 4, 5)
    
ElseIf Brick = 7 Then ' ':::'
    Call CheckStatus(10, 3, 4, 5)
    
ElseIf Brick = 8 Then ' 'Z'
    Call CheckStatus(10, 0, 2, 3)
    
ElseIf Brick = 9 Then ' Rotate 'Z'
    Call CheckStatus(10, 2, 3)
    
ElseIf Brick = 10 Then ' Inevrt 'T'
    Call CheckStatus(10, 1, 2, 3)
    
ElseIf Brick = 11 Then 'Rotate 'T' Left
    Call CheckStatus(10, 2, 3)
    
ElseIf Brick = 12 Then ' 'T'
    Call CheckStatus(10, 0, 2, 3)
    
ElseIf Brick = 13 Then 'Rotate 'T' Right
    Call CheckStatus(10, 1, 3)
    
ElseIf Brick = 14 Then ' 'L'
    Call CheckStatus(10, 2, 3)
    
ElseIf Brick = 15 Then ' Invert 'L'
    Call CheckStatus(10, 0, 3)
    
ElseIf Brick = 16 Then 'Rotate 'L' Right
    Call CheckStatus(10, 1, 2, 3)
    
ElseIf Brick = 17 Then 'Rotate 'L' Left
    Call CheckStatus(10, 1, 2, 3)
 
End If

If AVILABLE = False Then
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

MoveBlock
'DoEvents
End Sub
Private Sub Timer2_Timer()
    Call Timer1_Timer
End Sub

Private Sub Timer3_Timer()
    Call Timer2_Timer
End Sub

Private Sub Timer4_Timer()
    Call Timer3_Timer
End Sub

'-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/
'--------------------------------------------------- -*/
' Sub procedures below are used to Display 2D Button.-*/
' These sample code just draws colored Lines around  -*/
' image.                                             -*/
'--------------------------------------------------- -*/
'-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/-*/

Private Sub imgbutton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If imgbutton(Index).Visible = True Then

LL1.X1 = imgbutton(Index).Left
LT1.X1 = imgbutton(Index).Left
LL1.Y1 = imgbutton(Index).Top
LT1.Y1 = imgbutton(Index).Top

LL1.X2 = imgbutton(Index).Left
LT1.X2 = imgbutton(Index).Left + imgbutton(Index).Width
LL1.Y2 = imgbutton(Index).Top + imgbutton(Index).Height
LT1.Y2 = imgbutton(Index).Top

LL2.X2 = imgbutton(Index).Left + imgbutton(Index).Width
LT2.X2 = imgbutton(Index).Left
LL2.Y2 = imgbutton(Index).Top
LT2.Y2 = imgbutton(Index).Top + imgbutton(Index).Height

LL2.X1 = imgbutton(Index).Left + imgbutton(Index).Width
LT2.X1 = imgbutton(Index).Left + imgbutton(Index).Width
LL2.Y1 = imgbutton(Index).Top + imgbutton(Index).Height
LT2.Y1 = imgbutton(Index).Top + imgbutton(Index).Height

End If
End Sub

Private Sub imgbutton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Integer

  imgbutton(Index).Left = imgbutton(Index).Left + 40
  imgbutton(Index).Top = imgbutton(Index).Top + 40

Select Case Index

Case 0
    For i = 0 To 299
        Imgbox(i).Visible = False
    Next i
     SelectBrick
      
     imgbutton(1).Visible = True
     imgbutton(0).Visible = False
     imgbutton(6).Visible = True
     imgbutton(7).Visible = True
     imgbutton(8).Visible = True
     imgbutton(9).Visible = True
     
     Imgpausedis.Visible = False
     Imgmousebutdis.Visible = False
     Imgstartdis.Visible = True
     
      Me.txtscore.Text = "0"
   
      Timer1.Enabled = True
      GamePause = False
    
Case 1
    If Timer1.Enabled = False Then
        Timer1.Enabled = True
        GamePause = False
    Else
        Timer1.Enabled = False
        GamePause = True
    End If
    
Case 2
    frmOption.Show
    Timer1.Enabled = False
    GamePause = True
Case 3
    frmScore.Show
    Timer1.Enabled = False
    GamePause = True
Case 4
    frmHelp.Show
    Timer1.Enabled = False
    GamePause = True
Case 5
    Call Form_Unload(0)

Case 6
    Call Form_KeyDown(37, 0)

Case 7
    Call Form_KeyDown(38, 0)

Case 8
    Call Form_KeyDown(39, 0)

Case 9
    Call Form_KeyDown(40, 0)

End Select

End Sub

Private Sub imgbutton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  imgbutton(Index).Left = imgbutton(Index).Left - 40
  imgbutton(Index).Top = imgbutton(Index).Top - 40

End Sub
Private Sub imgbutton_DblClick(Index As Integer)
    
  imgbutton(Index).Left = imgbutton(Index).Left + 40
  imgbutton(Index).Top = imgbutton(Index).Top + 40

End Sub

