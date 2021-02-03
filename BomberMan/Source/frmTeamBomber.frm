VERSION 5.00
Begin VB.Form frmTeamBomber 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TeamBomber 1.0"
   ClientHeight    =   7800
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   11028
   Icon            =   "frmTeamBomber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTeamBomber.frx":08CA
   ScaleHeight     =   7800
   ScaleWidth      =   11028
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ExplosionFinH2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   7680
      Picture         =   "frmTeamBomber.frx":390C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   91
      Top             =   5280
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox ExplosionFinH1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   6960
      Picture         =   "frmTeamBomber.frx":6950
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   90
      Top             =   4560
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox ExplosionFinV2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   6000
      Picture         =   "frmTeamBomber.frx":9994
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   89
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox ExplosionFinV1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   5880
      Picture         =   "frmTeamBomber.frx":C9D8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   88
      Top             =   4200
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox ExplosionCorpsV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   5040
      Picture         =   "frmTeamBomber.frx":FA1C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   87
      Top             =   4200
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox ExplosionCorpsH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   4200
      Picture         =   "frmTeamBomber.frx":12A60
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   86
      Top             =   4200
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox ExplosionCentre 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   3360
      Picture         =   "frmTeamBomber.frx":15AA4
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   85
      Top             =   4200
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Index           =   1
      Left            =   6720
      Picture         =   "frmTeamBomber.frx":18AE8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   80
      Top             =   4200
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Index           =   0
      Left            =   5880
      Picture         =   "frmTeamBomber.frx":1BB2C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   79
      Top             =   4200
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Quitter"
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   7320
      Width           =   2892
   End
   Begin VB.CommandButton cmdNewRound 
      Caption         =   "Nouveau round"
      Height          =   492
      Left            =   2760
      TabIndex        =   78
      Top             =   7320
      Width           =   3972
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   3720
      Picture         =   "frmTeamBomber.frx":1EB70
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   77
      Top             =   7920
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   4800
      Picture         =   "frmTeamBomber.frx":21BB4
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   76
      Top             =   7440
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox MaisonOrange2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   2280
      Picture         =   "frmTeamBomber.frx":24BF8
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   75
      Top             =   3960
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox MaisonOrange1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   2280
      Picture         =   "frmTeamBomber.frx":27C3C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   74
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox MaisonVerte2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   1560
      Picture         =   "frmTeamBomber.frx":2AC80
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   73
      Top             =   3960
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox MaisonVerte1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   1560
      Picture         =   "frmTeamBomber.frx":2DCC4
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   72
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   23
      Left            =   7680
      Picture         =   "frmTeamBomber.frx":30D08
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   71
      Top             =   3600
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   22
      Left            =   7320
      Picture         =   "frmTeamBomber.frx":33D4C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   70
      Top             =   3600
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   21
      Left            =   7080
      Picture         =   "frmTeamBomber.frx":36D90
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   69
      Top             =   3600
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   20
      Left            =   6720
      Picture         =   "frmTeamBomber.frx":39DD4
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   68
      Top             =   3600
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   19
      Left            =   6360
      Picture         =   "frmTeamBomber.frx":3CE18
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   67
      Top             =   3600
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   18
      Left            =   6000
      Picture         =   "frmTeamBomber.frx":3FE5C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   66
      Top             =   3600
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   17
      Left            =   7800
      Picture         =   "frmTeamBomber.frx":42EA0
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   65
      Top             =   2640
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   16
      Left            =   7440
      Picture         =   "frmTeamBomber.frx":45EE4
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   64
      Top             =   2640
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   15
      Left            =   7200
      Picture         =   "frmTeamBomber.frx":48F28
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   63
      Top             =   2640
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   14
      Left            =   6840
      Picture         =   "frmTeamBomber.frx":4BF6C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   62
      Top             =   2640
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   13
      Left            =   6480
      Picture         =   "frmTeamBomber.frx":4EFB0
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   61
      Top             =   2640
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   12
      Left            =   6120
      Picture         =   "frmTeamBomber.frx":51FF4
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   60
      Top             =   2640
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   11
      Left            =   7800
      Picture         =   "frmTeamBomber.frx":55038
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   59
      Top             =   1680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   10
      Left            =   7440
      Picture         =   "frmTeamBomber.frx":5807C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   58
      Top             =   1680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   9
      Left            =   7200
      Picture         =   "frmTeamBomber.frx":5B0C0
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   57
      Top             =   1680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   8
      Left            =   6840
      Picture         =   "frmTeamBomber.frx":5E104
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   56
      Top             =   1680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   7
      Left            =   6480
      Picture         =   "frmTeamBomber.frx":61148
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   55
      Top             =   1680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   6
      Left            =   6120
      Picture         =   "frmTeamBomber.frx":6418C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   54
      Top             =   1680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   5
      Left            =   7800
      Picture         =   "frmTeamBomber.frx":671D0
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   53
      Top             =   840
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   4
      Left            =   7440
      Picture         =   "frmTeamBomber.frx":6A214
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   52
      Top             =   840
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   3
      Left            =   7200
      Picture         =   "frmTeamBomber.frx":6D258
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   51
      Top             =   840
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   2
      Left            =   6840
      Picture         =   "frmTeamBomber.frx":7029C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   50
      Top             =   840
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   1
      Left            =   6480
      Picture         =   "frmTeamBomber.frx":732E0
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   49
      Top             =   840
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   0
      Left            =   6120
      Picture         =   "frmTeamBomber.frx":76324
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   48
      Top             =   840
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox MurDestructible 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   8160
      Picture         =   "frmTeamBomber.frx":79368
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   47
      Top             =   7440
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox pctReady 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   3480
      Picture         =   "frmTeamBomber.frx":7C3AC
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   46
      Top             =   5280
      Visible         =   0   'False
      Width           =   2352
   End
   Begin VB.PictureBox pctGo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Left            =   5520
      Picture         =   "frmTeamBomber.frx":853F0
      ScaleHeight     =   768
      ScaleWidth      =   2304
      TabIndex        =   45
      Top             =   7320
      Visible         =   0   'False
      Width           =   2352
   End
   Begin VB.Timer TimerCommencer 
      Interval        =   1000
      Left            =   2040
      Top             =   7560
   End
   Begin VB.Timer Timer2 
      Index           =   3
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Index           =   2
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Index           =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Index           =   0
      Left            =   5040
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Index           =   3
      Left            =   6000
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Index           =   2
      Left            =   5880
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Index           =   1
      Left            =   6000
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Index           =   0
      Left            =   5880
      Top             =   5040
   End
   Begin VB.CommandButton cmdCommencer 
      Caption         =   "Commencer une nouvelle partie"
      Height          =   492
      Left            =   0
      TabIndex        =   44
      Top             =   7320
      Width           =   2772
   End
   Begin VB.PictureBox Bombe 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Index           =   2
      Left            =   5520
      Picture         =   "frmTeamBomber.frx":8E434
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   43
      Top             =   3840
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.PictureBox Bombe 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Index           =   1
      Left            =   5520
      Picture         =   "frmTeamBomber.frx":91178
      ScaleHeight     =   768
      ScaleWidth      =   720
      TabIndex        =   42
      Top             =   2880
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.PictureBox Bombe 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Index           =   0
      Left            =   5520
      Picture         =   "frmTeamBomber.frx":93EBC
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   41
      Top             =   2040
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.PictureBox Fond 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   5520
      ScaleHeight     =   924
      ScaleWidth      =   684
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox RouteBD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   3240
      Picture         =   "frmTeamBomber.frx":96C00
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   39
      Top             =   5880
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox RouteBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   2880
      Picture         =   "frmTeamBomber.frx":97029
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   38
      Top             =   5880
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox RouteHD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   2520
      Picture         =   "frmTeamBomber.frx":9745A
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   37
      Top             =   5880
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox RouteHG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   2160
      Picture         =   "frmTeamBomber.frx":97888
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   36
      Top             =   5880
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox RouteCroisement 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   1680
      Picture         =   "frmTeamBomber.frx":97CC1
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   35
      Top             =   5880
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox RouteV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   1200
      Picture         =   "frmTeamBomber.frx":980E5
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   34
      Top             =   5880
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox RouteH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   720
      Picture         =   "frmTeamBomber.frx":984E9
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   33
      Top             =   5880
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Mort 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   4
      Left            =   4560
      Picture         =   "frmTeamBomber.frx":98917
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Mort 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   3
      Left            =   4440
      Picture         =   "frmTeamBomber.frx":9B6A9
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   31
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Mort 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   2
      Left            =   3960
      Picture         =   "frmTeamBomber.frx":9DFDA
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Mort 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   1
      Left            =   3720
      Picture         =   "frmTeamBomber.frx":A0D7C
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Mort 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   0
      Left            =   3000
      Picture         =   "frmTeamBomber.frx":A36AD
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Maison2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   816
      Left            =   840
      Picture         =   "frmTeamBomber.frx":A644F
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Maison1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   816
      Left            =   840
      Picture         =   "frmTeamBomber.frx":A9493
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   26
      Top             =   4680
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   23
      Left            =   4080
      Picture         =   "frmTeamBomber.frx":A9A13
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   22
      Left            =   3360
      Picture         =   "frmTeamBomber.frx":ACA57
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   21
      Left            =   2760
      Picture         =   "frmTeamBomber.frx":AFA9B
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   20
      Left            =   2040
      Picture         =   "frmTeamBomber.frx":B2ADF
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   19
      Left            =   1560
      Picture         =   "frmTeamBomber.frx":B5B23
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   18
      Left            =   840
      Picture         =   "frmTeamBomber.frx":B8B67
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   17
      Left            =   4080
      Picture         =   "frmTeamBomber.frx":BBBAB
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   16
      Left            =   3360
      Picture         =   "frmTeamBomber.frx":BEBEF
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   15
      Left            =   2760
      Picture         =   "frmTeamBomber.frx":C1C33
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   14
      Left            =   2040
      Picture         =   "frmTeamBomber.frx":C4C77
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   13
      Left            =   1560
      Picture         =   "frmTeamBomber.frx":C7CBB
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   12
      Left            =   840
      Picture         =   "frmTeamBomber.frx":CACFF
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   11
      Left            =   4080
      Picture         =   "frmTeamBomber.frx":CDD43
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   10
      Left            =   3360
      Picture         =   "frmTeamBomber.frx":D0D87
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   9
      Left            =   2760
      Picture         =   "frmTeamBomber.frx":D3DCB
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   8
      Left            =   2040
      Picture         =   "frmTeamBomber.frx":D6E0F
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   7
      Left            =   1560
      Picture         =   "frmTeamBomber.frx":D9E53
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   6
      Left            =   840
      Picture         =   "frmTeamBomber.frx":DCE97
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   5
      Left            =   4080
      Picture         =   "frmTeamBomber.frx":DFEDB
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   4
      Left            =   3360
      Picture         =   "frmTeamBomber.frx":E2F1F
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   3
      Left            =   2760
      Picture         =   "frmTeamBomber.frx":E5F63
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   2
      Left            =   2040
      Picture         =   "frmTeamBomber.frx":E8FA7
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   1
      Left            =   1560
      Picture         =   "frmTeamBomber.frx":EBFEB
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox Bomber1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   816
      Index           =   0
      Left            =   840
      Picture         =   "frmTeamBomber.frx":EF02F
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   816
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7248
      Left            =   0
      Picture         =   "frmTeamBomber.frx":F2073
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   0
      Width           =   9648
      Begin VB.Timer Timer3 
         Interval        =   500
         Left            =   6720
         Top             =   5640
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9.6
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   3
         Left            =   8760
         TabIndex        =   84
         Top             =   6360
         Width           =   732
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9.6
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   2
         Left            =   8760
         TabIndex        =   83
         Top             =   4800
         Width           =   732
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9.6
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   1
         Left            =   8760
         TabIndex        =   82
         Top             =   3240
         Width           =   732
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9.6
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Index           =   0
         Left            =   8760
         TabIndex        =   81
         Top             =   1680
         Width           =   732
      End
   End
End
Attribute VB_Name = "frmTeamBomber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Bombe
    x As Integer
    y As Integer
    Délai As Integer
    TypeDélai As Integer
    Puissance As Integer
    Active As Boolean
    Util As Boolean
End Type

Private Type Bomber
    x As Integer
    y As Integer
    Bombes(0 To 3) As Bombe
    TypePos As Integer
    Pos As Integer
    Mort As Boolean
    Score As Integer
End Type

Dim Joueur(1 To 2) As Bomber
Dim Tableau(0 To 10, 0 To 8) As Integer
Dim TableBombe(0 To 10, 0 To 8) As Integer
Dim TableauMaison(0 To 10, 0 To 8) As Integer
Dim TableauBonus(0 To 10, 0 To 8) As Integer
Dim TableauPix(0 To 703, 0 To 575) As Integer
Dim CouleurTrans As Long
Dim BordureX As Integer
Dim BordureY As Integer
Dim ValeurDébut As Integer
Dim DéplacementJoueur1 As Integer
Dim DéplacementJoueur2 As Integer
Dim BombeJoueur1 As Integer
Dim BombeJoueur2 As Integer
Dim TimerExpiration As Integer

Private Sub cmdCommencer_Click()
    Dim i As Integer
        
    cmdNewRound_Click
        
    For i = LBound(Joueur) To UBound(Joueur)
        With Joueur(i)
            .Score = 0
        End With
    Next
    
End Sub

Private Sub cmdNewRound_Click()
    Dim i As Integer
    Dim j As Integer
        
    For i = LBound(Joueur) To UBound(Joueur)
        With Joueur(i)
            .Bombes(0).Active = True
            .Bombes(0).Délai = 3
            .Bombes(0).Puissance = 1
            .Bombes(0).TypeDélai = 0
            .Bombes(1).Active = False
            .Bombes(1).Délai = 3
            .Bombes(1).Puissance = 1
            .Bombes(1).TypeDélai = 0
            .Bombes(2).Active = False
            .Bombes(2).Délai = 3
            .Bombes(2).Puissance = 1
            .Bombes(2).TypeDélai = 0
            .Bombes(3).Active = False
            .Bombes(3).Délai = 3
            .Bombes(3).Puissance = 1
            .Bombes(3).TypeDélai = 0
            
            .Mort = False
        End With
    Next
    
    For i = 0 To 3
        Timer1(i).Interval = 1500
        Timer1(i).Enabled = False
        Timer2(i).Interval = 1500
        Timer2(i).Enabled = False
    Next
    Joueur(2).x = 10 * 64
    Joueur(2).y = 8 * 64
    InitTables
    ConstructionFond
    InitTableCaisses
    InitLblScore
    DoAffichage
    ValeurDébut = 0
    TimerCommencer.Enabled = True
    Timer3.Enabled = True
End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim diff As Single
    Dim i As Integer
    diff = pct.Width - pct.ScaleWidth
    Fond.Width = 704 * Screen.TwipsPerPixelX + diff
    diff = pct.Height - pct.ScaleHeight
    Fond.Height = 576 * Screen.TwipsPerPixelY + diff
    BordureX = 13
    BordureY = 13
    pct.ScaleMode = 3
    pct.Top = 0
    pct.Left = 0
    Me.Width = pct.Width + diff
    cmdCommencer_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Arret = True
End Sub

Private Sub ConstructionFond()
    Dim i As Integer
    Dim j  As Integer
    Dim k As Integer
    Dim l As Integer
    Dim coul As Integer
    
    Randomize
    Fond.ScaleMode = 3
    CouleurTrans = QBColor(3)
    For j = 1 To 7 Step 2
        For i = 1 To 9 Step 2
            coul = Int(3 * Rnd + 1)
            Tableau(i, j) = 50 + coul
            Select Case Tableau(i, j)
                Case 51
                    Fond.PaintPicture Maison1.Picture, i * 64, j * 64
                Case 52
                    Fond.PaintPicture MaisonOrange1.Picture, i * 64, j * 64
                Case 53
                    Fond.PaintPicture MaisonVerte1.Picture, i * 64, j * 64
            End Select
            For k = 0 To 63
                For l = 0 To 63
                    TableauPix(i * 64 + k, j * 64 + l) = 50 + coul
                Next
            Next
        Next
    Next
    For j = 0 To 8 Step 2
        For i = 1 To 9 Step 2
            Fond.PaintPicture RouteH.Picture, i * 64, j * 64
        Next
    Next
    For j = 1 To 9 Step 2
        For i = 0 To 8 Step 2
            Fond.PaintPicture RouteV.Picture, i * 64, j * 64
        Next
    Next
    For j = 2 To 7 Step 2
        For i = 2 To 8 Step 2
            Fond.PaintPicture RouteCroisement.Picture, i * 64, j * 64
        Next
    Next
    For i = 1 To 9
        Fond.PaintPicture RouteH.Picture, i * 64, 0
    Next
    For i = 1 To 9
        Fond.PaintPicture RouteH.Picture, i * 64, 8 * 64
    Next
    For j = 1 To 7
        Fond.PaintPicture RouteV.Picture, 0, j * 64
    Next
    For j = 1 To 7
        Fond.PaintPicture RouteV.Picture, 10 * 64, j * 64
    Next
    Fond.PaintPicture RouteHG.Picture, 0, 0
    Fond.PaintPicture RouteHD.Picture, 10 * 64, 0
    Fond.PaintPicture RouteBG.Picture, 0, 8 * 64
    Fond.PaintPicture RouteBD.Picture, 10 * 64, 8 * 64
    Fond.Picture = Fond.Image
    BitBlt pct.hdc, BordureX, BordureY, 704, 576, Fond.hdc, 0, 0, SRCCOPY
    pct.Refresh
    pct.Picture = pct.Image
End Sub

Private Sub DoAffichage()
    Dim i As Integer
    Dim j As Integer
    Dim r As RECT
    Dim r2 As RECT
    Dim TotalMorts As Integer
    Dim MortStr As String
    Dim TempStr As String
    Dim TempCent As Integer
    Dim Valeur As Integer
    pct.Cls
    If ValeurDébut < 0 Then
        For i = 0 To 10
            For j = 0 To 8
                r.Left = i * 64 + BordureX
                r.Bottom = r.Top + 59
                r.Right = r.Left + 59
                r.Top = j * 64 + BordureY
                r2.Left = 0
                r2.Bottom = 59
                r2.Right = 59
                r2.Top = 0
                If Tableau(i, j) = 10 Then ' Bombe 1
                    TransparentBlt pct.hdc, r, Bombe(0).hdc, r2, CouleurTrans
                    Bombe(0).Cls
                ElseIf Tableau(i, j) = 11 Then ' Bombe 2
                    TransparentBlt pct.hdc, r, Bombe(1).hdc, r2, CouleurTrans
                    Bombe(1).Cls
                ElseIf Tableau(i, j) = 12 Then ' Bombe 3
                    TransparentBlt pct.hdc, r, Bombe(2).hdc, r2, CouleurTrans
                    Bombe(2).Cls
                ElseIf Tableau(i, j) = 20 Then
                    BitBlt pct.hdc, i * 64 + BordureX, j * 64 + BordureY, 64, 64, MurDestructible.hdc, 0, 0, SRCCOPY
                ElseIf Tableau(i, j) >= 100 Then ' Explosion
                    r.Left = i * 64 + BordureX
                    r.Bottom = r.Top + 64
                    r.Right = r.Left + 64
                    r.Top = j * 64 + BordureY
                    r2.Left = 0
                    r2.Bottom = 64
                    r2.Right = 64
                    r2.Top = 0
                    TempStr = CStr(Tableau(i, j))
                    TempCent = CInt(Left(TempStr, 1)) * 100
                    Valeur = Tableau(i, j) - TempCent
                    Select Case Valeur
                        Case 0 ' Croix
                            TransparentBlt pct.hdc, r, ExplosionCentre.hdc, r2, CouleurTrans
                            ExplosionCentre.Cls
                        Case 1 ' Bordure Gauche
                            TransparentBlt pct.hdc, r, ExplosionFinH1.hdc, r2, CouleurTrans
                            ExplosionFinH1.Cls
                        Case 2 ' Bordure Droite
                            TransparentBlt pct.hdc, r, ExplosionFinH2.hdc, r2, CouleurTrans
                            ExplosionFinH2.Cls
                        Case 3 ' Corps horizontal
                            TransparentBlt pct.hdc, r, ExplosionCorpsH.hdc, r2, CouleurTrans
                            ExplosionCorpsH.Cls
                        Case 4 ' Bordure Haute
                            TransparentBlt pct.hdc, r, ExplosionFinV1.hdc, r2, CouleurTrans
                            ExplosionFinV1.Cls
                        Case 5 ' Bordure basse
                            TransparentBlt pct.hdc, r, ExplosionFinV2.hdc, r2, CouleurTrans
                            ExplosionFinV2.Cls
                        Case 6 ' Corps Vertical
                            TransparentBlt pct.hdc, r, ExplosionCorpsV.hdc, r2, CouleurTrans
                            ExplosionCorpsV.Cls
                    End Select
                ElseIf Tableau(i, j) = 0 And TableauBonus(i, j) > 0 Then ' Bonus
                    TransparentBlt pct.hdc, r, Bonus(TableauBonus(i, j) - 1).hdc, r2, CouleurTrans
                    Bonus(TableauBonus(i, j) - 1).Cls
                End If
            Next
        Next
        'Affichage Joueurs
        For j = LBound(Joueur) To UBound(Joueur)
            With Joueur(j)
                If .Mort = False Then
                    Select Case .TypePos
                        Case 0
                            i = 0
                        Case 1
                            i = 6
                        Case 2
                            i = 12
                        Case 3
                            i = 18
                    End Select
                    r.Left = .x + BordureX
                    r.Bottom = .y + 63
                    r.Right = .x + 63
                    r.Top = .y + BordureY
                    r2.Left = 0
                    r2.Bottom = 63
                    r2.Right = 63
                    r2.Top = 0
                    If j = 1 Then
                        TransparentBlt pct.hdc, r, Bomber1(i + .Pos).hdc, r2, CouleurTrans
                        Bomber1(i + .Pos).Cls
                    ElseIf j = 2 Then
                        TransparentBlt pct.hdc, r, Bomber2(i + .Pos).hdc, r2, CouleurTrans
                        Bomber2(i + .Pos).Cls
                    End If
                Else
                    TotalMorts = TotalMorts + 1
                End If
            End With
        Next
        ' Affichage Toit
        For j = 1 To 7 Step 2
            For i = 1 To 9 Step 2
                r.Left = i * 64 + BordureX
                r.Bottom = r.Top + 64
                r.Right = r.Left + 64
                r.Top = (j - 1) * 64 + BordureY
                r2.Left = 0
                r2.Bottom = 64
                r2.Right = 64
                r2.Top = 0
                Select Case Tableau(i, j)
                    Case 51
                        TransparentBlt pct.hdc, r, Maison2.hdc, r2, CouleurTrans
                        Maison2.Cls
                    Case 52
                        TransparentBlt pct.hdc, r, MaisonOrange2.hdc, r2, CouleurTrans
                        MaisonOrange2.Cls
                    Case 53
                        TransparentBlt pct.hdc, r, MaisonVerte2.hdc, r2, CouleurTrans
                        MaisonVerte2.Cls
                End Select
            Next
        Next
        If TotalMorts = UBound(Joueur) - LBound(Joueur) Or TotalMorts = UBound(Joueur) - LBound(Joueur) + 1 Then
            pct.Refresh
            For j = LBound(Joueur) To UBound(Joueur)
                MortStr = MortStr & "Joueur numéro " & CStr(j) & " : "
                If Joueur(j).Mort = True Then
                    MortStr = MortStr & "Mort ..."
                Else
                    MortStr = MortStr & "GAGNANT !"
                    Joueur(j).Score = Joueur(j).Score + 1
                End If
                MortStr = MortStr & Chr(13)
            Next
            MsgBox MortStr, vbOKOnly + vbExclamation, "Fin du round"
            InitLblScore
            Timer3.Enabled = False
            DéplacementJoueur1 = 0
            DéplacementJoueur2 = 0
            BombeJoueur1 = 0
            BombeJoueur2 = 0
            ValeurDébut = -1
        End If
    Else
        ' Affichage Toit
        For j = 1 To 7 Step 2
            For i = 1 To 9 Step 2
                r.Left = i * 64 + BordureX
                r.Bottom = r.Top + 64
                r.Right = r.Left + 64
                r.Top = (j - 1) * 64 + BordureY
                r2.Left = 0
                r2.Bottom = 64
                r2.Right = 64
                r2.Top = 0
                Select Case Tableau(i, j)
                    Case 51
                        TransparentBlt pct.hdc, r, Maison2.hdc, r2, CouleurTrans
                        Maison2.Cls
                    Case 52
                        TransparentBlt pct.hdc, r, MaisonOrange2.hdc, r2, CouleurTrans
                        MaisonOrange2.Cls
                    Case 53
                        TransparentBlt pct.hdc, r, MaisonVerte2.hdc, r2, CouleurTrans
                        MaisonVerte2.Cls
                End Select
            Next
        Next
        r.Left = 250 + BordureX
        r.Bottom = r.Top + 64
        r.Right = r.Left + 3 * 64
        r.Top = 250 + BordureY
        r2.Left = 0
        r2.Bottom = 63
        r2.Right = 3 * 64 - 1
        r2.Top = 0
        If ValeurDébut = 0 Or ValeurDébut = 1 Then
            TransparentBlt pct.hdc, r, pctReady.hdc, r2, CouleurTrans
            pctReady.Cls
        Else
            TransparentBlt pct.hdc, r, pctGo.hdc, r2, CouleurTrans
            pctGo.Cls
        End If
    End If
    pct.Refresh
End Sub

Private Sub pct_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        BombeJoueur1 = 1
    ElseIf KeyCode = 32 Then
        BombeJoueur2 = 1
    ElseIf KeyCode > 32 And KeyCode < 80 Then
        DéplacementJoueur1 = KeyCode
    Else
        DéplacementJoueur2 = KeyCode
    End If
End Sub

Private Sub pct_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        BombeJoueur1 = 0
    ElseIf KeyCode = 32 Then
        BombeJoueur2 = 0
    ElseIf KeyCode > 32 And KeyCode < 80 Then
        DéplacementJoueur1 = 0
    Else
        DéplacementJoueur2 = 0
    End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
    Joueur(1).Bombes(Index).Délai = Joueur(1).Bombes(Index).Délai - 1
    If Joueur(1).Bombes(Index).Délai = 0 Then
        If Joueur(1).Bombes(Index).TypeDélai = 0 Then
            InitTableBombes
            FaireExploser 1, Index
        Else ' Explosion déjà en cours
            Tableau(Joueur(1).Bombes(Index).x, Joueur(1).Bombes(Index).y) = 0
            Joueur(1).Bombes(Index).TypeDélai = 0
            Joueur(1).Bombes(Index).Délai = 3
            Joueur(1).Bombes(Index).Util = False
            SupprimerExplosion 1, Index
            Timer1(Index).Enabled = False
        End If
    Else
        If Joueur(1).Bombes(Index).TypeDélai = 0 Then
            Tableau(Joueur(1).Bombes(Index).x, Joueur(1).Bombes(Index).y) = 13 - Joueur(1).Bombes(Index).Délai
        End If
    End If
End Sub

Private Sub Timer2_Timer(Index As Integer)
    Joueur(2).Bombes(Index).Délai = Joueur(2).Bombes(Index).Délai - 1
    If Joueur(2).Bombes(Index).Délai = 0 Then
        If Joueur(2).Bombes(Index).TypeDélai = 0 Then
            InitTableBombes
            FaireExploser 2, Index
        Else ' Explosion déjà en cours
            Tableau(Joueur(2).Bombes(Index).x, Joueur(2).Bombes(Index).y) = 0
            Joueur(2).Bombes(Index).TypeDélai = 0
            Joueur(2).Bombes(Index).Délai = 3
            Joueur(2).Bombes(Index).Util = False
            SupprimerExplosion 2, Index
            Timer2(Index).Enabled = False
        End If
    Else
        If Joueur(2).Bombes(Index).TypeDélai = 0 Then
            Tableau(Joueur(2).Bombes(Index).x, Joueur(2).Bombes(Index).y) = 13 - Joueur(2).Bombes(Index).Délai
        End If
    End If
End Sub

Private Sub Timer3_Timer()
    TimerExpiration = 1
End Sub

Private Sub TimerCommencer_Timer()
    ValeurDébut = ValeurDébut + 1
    If ValeurDébut = 5 Then
        ValeurDébut = -1
        TimerCommencer.Enabled = False
    End If
    DoAffichage
End Sub

Private Sub FaireExploser(noJ As Integer, nb As Integer)
    Dim Puissance As Integer
    Dim noJ2 As Integer
    Dim nb2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim OrigX As Integer
    Dim OrigY As Integer
    Dim m As Integer
    Dim n As Integer
    Dim j2 As Integer
    
    TableBombe(Joueur(noJ).Bombes(nb).x, Joueur(noJ).Bombes(nb).y) = 1
    Puissance = Joueur(noJ).Bombes(nb).Puissance
    For k = -Puissance To Puissance Step 1
        i = k + Joueur(noJ).Bombes(nb).x
        j = Joueur(noJ).Bombes(nb).y
        If i >= 0 And i <= 10 And j >= 0 And j <= 8 Then
            If Tableau(i, j) >= 10 And Tableau(i, j) < 15 Then
                If TableBombe(i, j) = 0 Then
                    noJ2 = RechercherBombeJ(i, j)
                    nb2 = RechercherBombeB(i, j)
                    FaireExploser noJ2, nb2
                Else
                    Tableau(i, j) = noJ * 100 + nb * 10
                End If
            ElseIf (Tableau(i, j) = 0 Or Tableau(i, j) = 20) Then
                If Tableau(i, j) = 20 Then
                    For m = 0 To 63
                        For n = 0 To 63
                            TableauPix(i * 64 + m, j * 64 + n) = 0
                        Next
                    Next
                End If
                If i = Joueur(noJ).Bombes(nb).x And j = Joueur(noJ).Bombes(nb).y Then
                    Tableau(i, j) = noJ * 100 + nb * 10 ' Centre de la bombe (Croix)
                ElseIf i = Joueur(noJ).Bombes(nb).x - Puissance Then
                    Tableau(i, j) = noJ * 100 + nb * 10 + 1 ' Terminaison Gauche
                ElseIf i = Joueur(noJ).Bombes(nb).x + Puissance Then
                    Tableau(i, j) = noJ * 100 + nb * 10 + 2 ' Terminaison Droite
                Else
                    Tableau(i, j) = noJ * 100 + nb * 10 + 3 ' Corps Normal
                End If
            End If
            For j2 = LBound(Joueur) To UBound(Joueur)
                With Joueur(j2)
                    If (.x + 32) \ 64 = i And (.y + 32) \ 64 = j Then
                        .Mort = True
                    End If
                End With
            Next
        End If
    Next
    For k = -Puissance To Puissance Step 1
        i = Joueur(noJ).Bombes(nb).x
        j = k + Joueur(noJ).Bombes(nb).y
        If i >= 0 And i <= 10 And j >= 0 And j <= 8 Then
            If Tableau(i, j) >= 10 And Tableau(i, j) < 15 Then
                If TableBombe(i, j) = 0 Then
                    noJ2 = RechercherBombeJ(i, j)
                    nb2 = RechercherBombeB(i, j)
                    FaireExploser noJ2, nb2
                Else
                    Tableau(i, j) = noJ * 100 + nb * 10
                End If
            ElseIf (Tableau(i, j) = 0 Or Tableau(i, j) = 20) Then
                If Tableau(i, j) = 20 Then
                    For m = 0 To 63
                        For n = 0 To 63
                            TableauPix(i * 64 + m, j * 64 + n) = 0
                        Next
                    Next
                End If
                If i = Joueur(noJ).Bombes(nb).x And j = Joueur(noJ).Bombes(nb).y Then
                    Tableau(i, j) = noJ * 100 + nb * 10 ' Centre de la bombe (Croix)
                ElseIf j = Joueur(noJ).Bombes(nb).y - Puissance Then
                    Tableau(i, j) = noJ * 100 + nb * 10 + 4 ' Terminaison Haute
                ElseIf j = Joueur(noJ).Bombes(nb).y + Puissance Then
                    Tableau(i, j) = noJ * 100 + nb * 10 + 5 ' Terminaison Bas
                Else
                    Tableau(i, j) = noJ * 100 + nb * 10 + 6 ' Corps Normal Vertical
                End If
            End If
            For j2 = LBound(Joueur) To UBound(Joueur)
                With Joueur(j2)
                    If (.x + 32) \ 64 = i And (.y + 32) \ 64 = j Then
                        .Mort = True
                    End If
                End With
            Next
        End If
    Next
    
    If noJ = 1 Then
        Timer1(nb).Enabled = True
    Else
        Timer2(nb).Enabled = True
    End If
    Joueur(noJ).Bombes(nb).Délai = 1
    Joueur(noJ).Bombes(nb).TypeDélai = 1
End Sub

Private Function RechercherBombeJ(i As Integer, j As Integer) As Integer
    Dim x As Integer
    Dim z As Integer
    Dim Resul As Integer
    Resul = 0
    For z = 1 To 2
        For x = 0 To 3
            If Joueur(z).Bombes(x).Active = True Then
                If Joueur(z).Bombes(x).Util = True Then
                    If Joueur(z).Bombes(x).x = i And Joueur(z).Bombes(x).y = j Then
                        Resul = z
                        Exit For
                    End If
                End If
            End If
        Next
        If Resul <> 0 Then Exit For
    Next
    RechercherBombeJ = Resul
End Function

Private Function RechercherBombeB(i As Integer, j As Integer) As Integer
    Dim x As Integer
    Dim z As Integer
    Dim Resul As Integer
    Resul = -1
    For z = 1 To 2
        For x = 0 To 3
            If Joueur(z).Bombes(x).Active = True Then
                If Joueur(z).Bombes(x).Util = True Then
                    If Joueur(z).Bombes(x).x = i And Joueur(z).Bombes(x).y = j Then
                        Resul = x
                        Exit For
                    End If
                End If
            End If
        Next
        If Resul <> -1 Then Exit For
    Next
    RechercherBombeB = Resul
End Function

Private Sub InitTables()
    Dim i As Integer
    Dim j As Integer
    InitTableBombes
    For i = 0 To 10
        For j = 0 To 8
            Tableau(i, j) = 0
        Next
    Next
End Sub

Private Sub InitTableBombes()
    Dim i As Integer
    Dim j As Integer
    For i = 0 To 10
        For j = 0 To 8
            TableBombe(i, j) = 0
        Next
    Next
End Sub

Private Sub InitTableCaisses()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim bonus1 As Integer
    Dim bonus2 As Integer
    Dim Resul As Integer
    Randomize
    Do While Resul < 20
        i = Int(11 * Rnd)
        j = Int(9 * Rnd)
        If Tableau(i, j) = 0 And (i <> 0 And j <> 0) And (i <> UBound(Tableau, 1) And j <> UBound(Tableau, 2)) Then
            Tableau(i, j) = 20
            bonus2 = Int(2 * Rnd)
            If bonus2 = 1 Then
                bonus1 = Int(2 * Rnd)
                If bonus1 = 1 Then ' Nouvelle bombe
                    TableauBonus(i, j) = 1
                Else ' Puissance
                    TableauBonus(i, j) = 2
                End If
            End If
            Resul = Resul + 1
        End If
    Loop
    For i = 0 To 10
        For j = 0 To 8
            If Tableau(i, j) = 20 Then
                For k = 0 To 63
                    For l = 0 To 63
                        TableauPix(i * 64 + k, j * 64 + l) = 20
                    Next
                Next
            End If
        Next
    Next
End Sub

Private Sub SupprimerExplosion(noJ As Integer, nb As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    temp = noJ * 100 + nb * 10
    For i = 0 To 10
        For j = 0 To 8
            If Tableau(i, j) >= temp And Tableau(i, j) <= (noJ * 100 + (nb + 1) * 10) Then
                Tableau(i, j) = 0
            End If
        Next
    Next
End Sub

Public Sub Déplacement()
    Dim i As Integer

    If ValeurDébut < 0 Then
        If DéplacementJoueur1 > 0 Then
            Select Case DéplacementJoueur1
                Case 37 ' Gauche Joueur 1
                    If Joueur(1).TypePos <> 2 Then
                        Joueur(1).TypePos = 2
                        Joueur(1).Pos = 0
                    Else
                        Joueur(1).Pos = (Joueur(1).Pos + 1) Mod 6
                    End If
                    If Joueur(1).x - 8 > 0 Then
                        If TableauPix((Joueur(1).x - 8), (Joueur(1).y + 50)) < 10 Then
                            Joueur(1).x = Joueur(1).x - 8
                            CheckBonus 1
                        End If
                    End If
                Case 39 ' Joueur 1 droite
                    If Joueur(1).TypePos <> 0 Then
                        Joueur(1).TypePos = 0
                        Joueur(1).Pos = 0
                    Else
                        Joueur(1).Pos = (Joueur(1).Pos + 1) Mod 6
                    End If
                    If Joueur(1).x + 64 + 8 < 11 * 64 Then
                        If TableauPix((Joueur(1).x + 64 + 8), (Joueur(1).y + 50)) < 10 Then
                            Joueur(1).x = Joueur(1).x + 8
                            CheckBonus 1
                        End If
                    End If
                Case 40 ' Joueur 1 Bas
                    If Joueur(1).TypePos <> 1 Then
                        Joueur(1).TypePos = 1
                        Joueur(1).Pos = 1
                    Else
                        Joueur(1).Pos = (Joueur(1).Pos + 1) Mod 6
                    End If
                    If Joueur(1).y + 64 + 8 < 9 * 64 Then
                        If TableauPix((Joueur(1).x + 40), (Joueur(1).y + 64 + 8)) < 10 Then
                            Joueur(1).y = Joueur(1).y + 8
                            CheckBonus 1
                        End If
                    End If
                Case 38 ' Joueur 1 Haut
                    If Joueur(1).TypePos <> 3 Then
                        Joueur(1).TypePos = 3
                        Joueur(1).Pos = 3
                    Else
                        Joueur(1).Pos = (Joueur(1).Pos + 1) Mod 6
                    End If
                    If Joueur(1).y - 8 > 0 Then
                        If TableauPix((Joueur(1).x + 40), (Joueur(1).y - 8)) < 10 Then
                            Joueur(1).y = Joueur(1).y - 8
                            CheckBonus 1
                        End If
                    End If
            End Select
        End If
        
        If BombeJoueur1 > 0 Then ' Poser une bombe Joueur1
            If Tableau((Joueur(1).x + 32) \ 64, (Joueur(1).y + 32) \ 64) = 0 Then
                For i = 0 To 3
                    If Joueur(1).Bombes(i).Active = True Then
                        If Joueur(1).Bombes(i).Util = False Then
                            Joueur(1).Bombes(i).Util = True
                            Joueur(1).Bombes(i).x = (Joueur(1).x + 32) \ 64
                            Joueur(1).Bombes(i).y = (Joueur(1).y + 32) \ 64
                            Tableau(Joueur(1).Bombes(i).x, Joueur(1).Bombes(i).y) = 10
                            Timer1(i).Enabled = True
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
            
        If DéplacementJoueur2 > 0 Then
            Select Case DéplacementJoueur2
                Case 100 ' Gauche Joueur 2
                    If Joueur(2).TypePos <> 2 Then
                        Joueur(2).TypePos = 2
                        Joueur(2).Pos = 0
                    Else
                        Joueur(2).Pos = (Joueur(2).Pos + 1) Mod 6
                    End If
                    If Joueur(2).x - 8 > 0 Then
                        If TableauPix((Joueur(2).x - 8), (Joueur(2).y + 50)) < 10 Then
                            Joueur(2).x = Joueur(2).x - 8
                            CheckBonus 2
                        End If
                    End If
                Case 102 ' Joueur 2 droite
                    If Joueur(2).TypePos <> 0 Then
                        Joueur(2).TypePos = 0
                        Joueur(2).Pos = 0
                    Else
                        Joueur(2).Pos = (Joueur(2).Pos + 1) Mod 6
                    End If
                    If Joueur(2).x + 64 + 8 < 11 * 64 Then
                        If TableauPix((Joueur(2).x + 64 + 8), (Joueur(2).y + 50)) < 10 Then
                            Joueur(2).x = Joueur(2).x + 8
                            CheckBonus 2
                        End If
                    End If
                Case 98 ' Joueur 2 Bas
                    If Joueur(2).TypePos <> 1 Then
                        Joueur(2).TypePos = 1
                        Joueur(2).Pos = 1
                    Else
                        Joueur(2).Pos = (Joueur(2).Pos + 1) Mod 6
                    End If
                    If Joueur(2).y + 64 + 8 < 9 * 64 Then
                        If TableauPix((Joueur(2).x + 40), (Joueur(2).y + 64 + 8)) < 10 Then
                            Joueur(2).y = Joueur(2).y + 8
                            CheckBonus 2
                        End If
                    End If
                Case 104 ' Joueur 2 Haut
                    If Joueur(2).TypePos <> 3 Then
                        Joueur(2).TypePos = 3
                        Joueur(2).Pos = 3
                    Else
                        Joueur(2).Pos = (Joueur(2).Pos + 1) Mod 6
                    End If
                    If Joueur(2).y - 8 > 0 Then
                        If TableauPix((Joueur(2).x + 40), (Joueur(2).y - 8)) < 10 Then
                            Joueur(2).y = Joueur(2).y - 8
                            CheckBonus 2
                        End If
                    End If
            End Select
        End If
                
        If BombeJoueur2 > 0 Then ' Poser une bombe Joueur 2
            If Tableau((Joueur(2).x + 32) \ 64, (Joueur(2).y + 32) \ 64) = 0 Then
                For i = 0 To 3
                    If Joueur(2).Bombes(i).Active = True Then
                        If Joueur(2).Bombes(i).Util = False Then
                            Joueur(2).Bombes(i).Util = True
                            Joueur(2).Bombes(i).x = (Joueur(2).x + 32) \ 64
                            Joueur(2).Bombes(i).y = (Joueur(2).y + 32) \ 64
                            Tableau(Joueur(2).Bombes(i).x, Joueur(2).Bombes(i).y) = 10
                            Timer2(i).Enabled = True
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        If DéplacementJoueur1 > 0 Or DéplacementJoueur2 > 0 Or BombeJoueur1 > 0 Or BombeJoueur2 > 0 Or TimerExpiration = 1 Then
            DoAffichage
            TimerExpiration = 0
        End If
    End If
End Sub

Private Sub CheckBonus(noJ As Integer)
    Dim i As Integer
    With Joueur(noJ)
        If TableauBonus((.x + 32) \ 64, (.y + 32) \ 64) = 1 Then
            For i = 1 To 3
                If Joueur(noJ).Bombes(i).Active = False Then
                    Joueur(noJ).Bombes(i).Active = True
                    Exit Sub
                End If
            Next
            TableauBonus((.x + 32) \ 64, (.y + 32) \ 64) = 0
        ElseIf TableauBonus((.x + 32) \ 64, (.y + 32) \ 64) = 2 Then
            For i = 0 To 3
                Joueur(noJ).Bombes(i).Puissance = Joueur(noJ).Bombes(i).Puissance + 1
            Next
            TableauBonus((.x + 32) \ 64, (.y + 32) \ 64) = 0
        End If
    End With
End Sub

Private Sub InitLblScore()
    Dim i As Integer
    For i = 1 To UBound(Joueur)
        lblScore(i - 1).Caption = Joueur(i).Score
    Next
End Sub
