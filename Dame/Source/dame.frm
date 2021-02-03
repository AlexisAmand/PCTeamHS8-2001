VERSION 5.00
Begin VB.Form frmDames 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dames"
   ClientHeight    =   8112
   ClientLeft      =   156
   ClientTop       =   444
   ClientWidth     =   8808
   Icon            =   "dame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8112
   ScaleWidth      =   8808
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Quitter"
      Height          =   1095
      Left            =   7200
      TabIndex        =   104
      Top             =   1560
      Width           =   1572
   End
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminer le tour"
      Height          =   1095
      Left            =   7200
      TabIndex        =   103
      Top             =   480
      Width           =   1572
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   99
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   99
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   98
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   98
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   97
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   97
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   96
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   96
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   95
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   95
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   94
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   94
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   93
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   93
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   92
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   92
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   91
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   91
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   90
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   90
      Top             =   6960
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   89
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   89
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   88
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   88
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   87
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   87
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   86
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   86
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   85
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   85
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   84
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   84
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   83
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   83
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   82
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   82
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   81
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   81
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   80
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   80
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   79
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   79
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   78
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   78
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   77
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   77
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   76
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   76
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   75
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   75
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   74
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   74
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   73
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   73
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   72
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   72
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   71
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   71
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   70
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   70
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   69
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   69
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   68
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   68
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   67
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   67
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   66
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   66
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   65
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   65
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   64
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   64
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   63
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   63
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   62
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   62
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   61
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   61
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   60
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   60
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   59
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   59
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   58
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   58
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   57
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   57
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   56
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   56
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   55
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   55
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   54
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   54
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   53
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   53
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   52
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   52
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   51
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   51
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   50
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   50
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   49
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   49
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   48
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   48
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   47
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   47
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   46
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   46
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   45
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   45
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   44
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   44
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   43
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   43
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   42
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   42
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   41
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   41
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   40
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   40
      Top             =   3360
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   39
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   39
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   38
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   38
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   37
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   37
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   36
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   36
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   35
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   35
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   34
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   34
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   33
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   33
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   32
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   32
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   31
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   31
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   30
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   30
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   29
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   29
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   28
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   28
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   27
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   27
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   26
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   26
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   25
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   25
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   24
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   24
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   23
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   23
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   22
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   22
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   21
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   21
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   20
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   20
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   19
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   19
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   18
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   18
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   17
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   16
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   15
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   14
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   14
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   13
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   12
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   12
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   11
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   10
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   9
      Left            =   6480
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   8
      Left            =   5760
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   7
      Left            =   5040
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   6
      Left            =   4320
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   5
      Left            =   3600
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   4
      Left            =   2880
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   2160
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   1440
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   720
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Cases 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   0
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Joueur 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   105
      Top             =   7680
      Width           =   6372
   End
   Begin VB.Label tourLbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tour : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   7200
      TabIndex        =   102
      Top             =   2640
      Width           =   1572
   End
   Begin VB.Label Joueur2Lbl 
      Caption         =   "Joueur 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Joueur1Lbl 
      Caption         =   "Joueur 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   100
      Top             =   9000
      Width           =   6375
   End
End
Attribute VB_Name = "frmDames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type Piece
    Contient As Integer
    PremierCoup As Boolean
    Reine As Boolean
End Type

Private Type point
    i As Integer
    j As Integer
End Type

Dim plateau(0 To 9, 0 To 9) As Piece
Dim NumJeton(0 To 1) As Integer
Dim joueur As Integer
Dim coup() As point
Dim init As Boolean
Dim i1 As Integer
Dim i2 As Integer
Dim j1 As Integer
Dim j2 As Integer

Private Sub Cases_Click(index As Integer)
    Dim max As Integer
    
    If init = False Then
        j1 = Int(index / 10)
        i1 = index - j1 * 10
        If plateau(i1, j1).Contient = joueur Then
            coup(0).i = i1
            coup(0).j = j1
            init = True
            Cases(index).Line (1, 1)-(1, 45), vbRed
            Cases(index).Line (1, 45)-(45, 45), vbRed
            Cases(index).Line (45, 45)-(45, 1), vbRed
            Cases(index).Line (45, 1)-(1, 1), vbRed
        End If
    Else
        j2 = Int(index / 10)
        i2 = index - j2 * 10
        max = UBound(coup)
        ' Pour annuler un deplacement, on reclique sur la derniere case choisie
        If (j2 <> coup(max).j) And (i2 <> coup(max).i) Then
            ReDim Preserve coup(0 To max + 1) As point
            coup(max + 1).i = i2
            coup(max + 1).j = j2
            Cases(index).Line (1, 1)-(1, 45), vbBlue
            Cases(index).Line (1, 45)-(45, 45), vbBlue
            Cases(index).Line (45, 45)-(45, 1), vbBlue
            Cases(index).Line (45, 1)-(1, 1), vbBlue
        Else
            ' On change le jeton de depart
            If max = 0 Then
                init = False
                Cases(coup(0).i + coup(0).j * 10).Line (1, 1)-(1, 45), Cases(coup(0).i + coup(0).j * 10).BackColor
                Cases(coup(0).i + coup(0).j * 10).Line (1, 45)-(45, 45), Cases(coup(0).i + coup(0).j * 10).BackColor
                Cases(coup(0).i + coup(0).j * 10).Line (45, 45)-(45, 1), Cases(coup(0).i + coup(0).j * 10).BackColor
                Cases(coup(0).i + coup(0).j * 10).Line (45, 1)-(1, 1), Cases(coup(0).i + coup(0).j * 10).BackColor
                ReDim coup(0 To 0) As point
            Else
                ReDim Preserve coup(0 To max - 1) As point
                Cases(index).Line (1, 1)-(1, 45), Cases(index).BackColor
                Cases(index).Line (1, 45)-(45, 45), Cases(index).BackColor
                Cases(index).Line (45, 45)-(45, 1), Cases(index).BackColor
                Cases(index).Line (45, 1)-(1, 1), Cases(index).BackColor
            End If
        End If
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTerminate_Click()
    Dim l As Integer
    Dim index As Integer
    
    For l = 0 To UBound(coup) - 1
        index = coup(l).i + 10 * coup(l).j
        deplacer coup(l).i, coup(l).j, coup(l + 1).i, coup(l + 1).j
        Cases(index).FillColor = Cases(index).BackColor
        Cases(index).Line (1, 1)-(1, 45), Cases(index).BackColor
        Cases(index).Line (1, 45)-(45, 45), Cases(index).BackColor
        Cases(index).Line (45, 45)-(45, 1), Cases(index).BackColor
        Cases(index).Line (45, 1)-(1, 1), Cases(index).BackColor
    Next l
    index = coup(l).i + 10 * coup(l).j
    Cases(index).FillColor = Cases(index).BackColor
        Cases(index).Line (1, 1)-(1, 45), Cases(index).BackColor
        Cases(index).Line (1, 45)-(45, 45), Cases(index).BackColor
        Cases(index).Line (45, 45)-(45, 1), Cases(index).BackColor
        Cases(index).Line (45, 1)-(1, 1), Cases(index).BackColor
    ReDim coup(0 To 0) As point
    init = False
    
    If NumJeton((joueur + 1) Mod 2) = 0 Then
        tourLbl.Caption = "joueur" & (joueur + 1) & " a gagne "
        cmdTerminate.Enabled = False
        Exit Sub
    End If
    joueur = (joueur + 1) Mod 2
    tourLbl.Caption = "Tour : Joueur " & joueur + 1
    affichePlateau
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    
    ' initialisation des deux premieres rangees de chaque joueur
    For i = 0 To 9
        For j = 0 To 9
            plateau(i, j).Contient = 2
        Next j
    Next i
    For i = 1 To 9 Step 2
        plateau(i, 0).Reine = False
        plateau(i, 0).PremierCoup = True
        plateau(i, 0).Contient = 1
    Next i
    For i = 0 To 9 Step 2
        plateau(i, 1).Reine = False
        plateau(i, 1).PremierCoup = True
        plateau(i, 1).Contient = 1
    Next i
    For i = 1 To 9 Step 2
        plateau(i, 8).Reine = False
        plateau(i, 8).PremierCoup = True
        plateau(i, 8).Contient = 0
    Next i
    For i = 0 To 9 Step 2
        plateau(i, 9).Reine = False
        plateau(i, 9).PremierCoup = True
        plateau(i, 9).Contient = 0
    Next i
    init = False
    
    ' Nombre total de pieces
    NumJeton(0) = 10
    NumJeton(1) = 10
    joueur = 0
    tourLbl.Caption = "Tour : Joueur " & joueur + 1
    ReDim coup(0 To 0) As point
    affichePlateau
End Sub

Private Sub deplacer(ByVal i1 As Integer, ByVal j1 As Integer, ByVal i2 As Integer, ByVal j2 As Integer)
    Dim sensi As Integer
    Dim sensj As Integer
    Dim incI As Integer
    Dim IncJ As Integer
    Dim k As Integer
    Dim m As Integer
    
    If plateau(i2, j2).Contient = 2 Then
        If plateau(i1, j1).Reine = False Then
            placerPion i1, j1, i2, j2
        Else
            sensi = i2 - i1
            sensj = j2 - j1
            If Abs(sensi) <= 2 And Abs(sensj) <= 2 Then
                placerPion i1, j1, i2, j2
            Else
                incI = Sgn(sensi)
                IncJ = Sgn(sensj)
                'on regarde si le chemin est libre pour la reine
                For k = i1 + incI To i2 - 2 * incI Step incI
                    For m = j1 + IncJ To j2 - 2 * IncJ Step IncJ
                        If Abs(k / m) = 1 Then
                            If plateau(k, m).Contient <> 2 Then
                                Exit Sub
                            End If
                        End If
                    Next m
                Next k
                'si le chemin est une diagonale
                If Abs((i2 - i1) / (j2 - j1)) = 1 Then
                    poserPion i1, j1, i2 - 2 * Sgn(sensi), j2 - 2 * Sgn(sensj)
                    placerPion i2 - 2 * Sgn(sensi), j2 - 2 * Sgn(sensj), i2, j2
                End If
            End If
        End If
    End If
    affichePlateau
End Sub

Private Sub placerPion(i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer)
    Dim sensi As Integer
    Dim sensj As Integer
    
    If joueur = 0 Then
        sensi = i2 - i1
        sensj = j1 - j2
        ' Si on avance d une case
        If Abs(sensi) = 1 And Abs(sensj) = 1 Then
           If plateau(i1, j1).PremierCoup = True Then
               plateau(i1, j1).PremierCoup = False
           End If
           poserPion i1, j1, i2, j2
           ' Si on avance de deux cases
        ElseIf Abs(sensi) = 2 And Abs(sensj) = 2 Then
           ' On mange une piece de l'adversaire?
           If (plateau(i1 + Sgn(sensi), j1 - Sgn(sensj)).Contient = ((joueur + 1) Mod 2)) Then
               plateau(i1 + Sgn(sensi), j1 - Sgn(sensj)).Contient = 2
               NumJeton(1) = NumJeton(1) - 1
               poserPion i1, j1, i2, j2
           ' On a le droit de sauter deux cases?
           ElseIf plateau(i1 + Sgn(sensi), j1 - Sgn(sensj)).Contient = 2 And ((plateau(i1, j1).PremierCoup = True And sensj = 2) Or plateau(i1, j1).Reine = True) Then
               plateau(i1, j1).PremierCoup = False
               poserPion i1, j1, i2, j2
           End If
           
        End If
    Else ' Même chose dans le cas du joueur deux
        sensi = i2 - i1
        sensj = j2 - j1
        If Abs(sensi) = 1 And Abs(sensj) = 1 Then
           If plateau(i1, j1).PremierCoup = True Then
               plateau(i1, j1).PremierCoup = False
           End If
           poserPion i1, j1, i2, j2
        ElseIf Abs(sensi) = 2 And Abs(sensj) = 2 Then
           If (plateau(i1 + Sgn(sensi), j1 + Sgn(sensj)).Contient = ((joueur + 1) Mod 2)) Then
               plateau(i1 + Sgn(sensi), j1 + Sgn(sensj)).Contient = 2
               NumJeton(0) = NumJeton(0) - 1
               poserPion i1, j1, i2, j2
           ElseIf plateau(i1 + Sgn(sensi), j1 + Sgn(sensj)).Contient = 2 And ((plateau(i1, j1).PremierCoup = True And sensj = 2) Or plateau(i1, j1).Reine = True) Then
               plateau(i1, j1).PremierCoup = False
               poserPion i1, j1, i2, j2
           End If
        End If
    End If
End Sub

Private Sub poserPion(i1 As Integer, j1 As Integer, i2 As Integer, j2 As Integer)
    plateau(i2, j2).Contient = plateau(i1, j1).Contient
    plateau(i2, j2).Reine = plateau(i1, j1).Reine
    plateau(i2, j2).PremierCoup = plateau(i1, j1).PremierCoup
    plateau(i1, j1).Contient = 2
    If joueur = 0 And j2 = 0 Then
        plateau(i2, j2).Reine = True
    ElseIf joueur = 1 And j2 = 9 Then
        plateau(i2, j2).Reine = True
    End If
End Sub

Private Sub affichePlateau()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    For i = 0 To 9
        For j = 0 To 9
            k = j * 10 + i
            frmDames.Cases(k).FillColor = frmDames.Cases(k).BackColor
            frmDames.Cases(k).Circle (23, 23), 20, frmDames.Cases(k).BackColor
            frmDames.Cases(k).FillStyle = vbSolid
            If plateau(i, j).Contient = 1 Then
                frmDames.Cases(k).FillColor = vbBlack
                DrawDame k
                ' Dessin de la reine
                 If plateau(i, j).Reine = True Then
                    frmDames.Cases(k).Line (15, 30)-(15, 10), vbWhite
                    frmDames.Cases(k).Line (15, 10)-(25, 10), vbWhite
                    frmDames.Cases(k).Line (25, 10)-(25, 20), vbWhite
                    frmDames.Cases(k).Line (25, 20)-(15, 20), vbWhite
                    frmDames.Cases(k).Line (20, 20)-(25, 30), vbWhite
                End If
            ElseIf plateau(i, j).Contient = 0 Then
                frmDames.Cases(k).FillColor = vbWhite
                DrawDame k
                ' Dessin de la reine
                If plateau(i, j).Reine = True Then
                    frmDames.Cases(k).Line (15, 30)-(15, 10), vbBlack
                    frmDames.Cases(k).Line (15, 10)-(25, 10), vbBlack
                    frmDames.Cases(k).Line (25, 10)-(25, 20), vbBlack
                    frmDames.Cases(k).Line (25, 20)-(15, 20), vbBlack
                    frmDames.Cases(k).Line (20, 20)-(25, 30), vbBlack
                End If
            End If
        Next j
    Next i
End Sub

Private Sub DrawDame(k As Integer)
    frmDames.Cases(k).Circle (23, 23), 20
End Sub

