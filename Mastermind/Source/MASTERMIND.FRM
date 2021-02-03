VERSION 5.00
Begin VB.Form frmMastermind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mastermind"
   ClientHeight    =   6912
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   6684
   Icon            =   "Mastermind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6912
   ScaleWidth      =   6684
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Gestion"
      Height          =   6732
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   2412
      Begin VB.CommandButton cmdValider 
         Caption         =   "Valider le tour"
         Height          =   372
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   2172
      End
      Begin VB.CommandButton cmdCoul 
         Height          =   372
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5640
         Width           =   2172
      End
      Begin VB.CommandButton cmdCoul 
         Height          =   372
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5040
         Width           =   2172
      End
      Begin VB.CommandButton cmdCoul 
         Height          =   372
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4440
         Width           =   2172
      End
      Begin VB.CommandButton cmdCoul 
         Height          =   372
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   2172
      End
      Begin VB.CommandButton cmdCoul 
         Height          =   372
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3240
         Width           =   2172
      End
      Begin VB.CommandButton cmdCoul 
         Height          =   372
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   2172
      End
      Begin VB.PictureBox pct 
         AutoRedraw      =   -1  'True
         Height          =   372
         Left            =   120
         ScaleHeight     =   324
         ScaleWidth      =   2124
         TabIndex        =   6
         Top             =   1440
         Width           =   2172
      End
      Begin VB.CommandButton cmdCoul 
         BackColor       =   &H00C0C0C0&
         Height          =   372
         Index           =   0
         Left            =   120
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   2172
      End
      Begin VB.CommandButton cmdQuitter 
         Caption         =   "Quitter"
         Height          =   372
         Left            =   120
         TabIndex        =   3
         Top             =   6240
         Width           =   2172
      End
      Begin VB.CommandButton cmdNouvellePartie 
         Caption         =   "Nouvelle partie"
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2172
      End
      Begin VB.Label lblCouleur 
         Caption         =   "Choix de la couleur :"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2172
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jeu"
      Height          =   6732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3972
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   23
         Left            =   3480
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   65
         Top             =   1200
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   22
         Left            =   3240
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   21
         Left            =   3000
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   20
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   19
         Left            =   3480
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   61
         Top             =   2280
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   18
         Left            =   3240
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   60
         Top             =   2280
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   17
         Left            =   3000
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   59
         Top             =   2280
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   16
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   58
         Top             =   2280
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   15
         Left            =   3480
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   57
         Top             =   3360
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   14
         Left            =   3240
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   56
         Top             =   3360
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   13
         Left            =   3000
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   55
         Top             =   3360
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   12
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   54
         Top             =   3360
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   11
         Left            =   3480
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   53
         Top             =   4440
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   10
         Left            =   3240
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   52
         Top             =   4440
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   9
         Left            =   3000
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   51
         Top             =   4440
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   8
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   50
         Top             =   4440
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   7
         Left            =   3480
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   49
         Top             =   5400
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   6
         Left            =   3240
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   48
         Top             =   5400
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   5
         Left            =   3000
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   47
         Top             =   5400
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   4
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   46
         Top             =   5400
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   3
         Left            =   3480
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   45
         Top             =   6240
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   2
         Left            =   3240
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   44
         Top             =   6240
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   1
         Left            =   3000
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   43
         Top             =   6240
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctValidResul 
         Height          =   132
         Index           =   0
         Left            =   2760
         ScaleHeight     =   84
         ScaleWidth      =   84
         TabIndex        =   42
         Top             =   6240
         Visible         =   0   'False
         Width           =   132
      End
      Begin VB.PictureBox pctResul 
         AutoRedraw      =   -1  'True
         Height          =   372
         Index           =   3
         Left            =   2040
         ScaleHeight     =   324
         ScaleWidth      =   324
         TabIndex        =   39
         Top             =   360
         Width           =   372
      End
      Begin VB.PictureBox pctResul 
         AutoRedraw      =   -1  'True
         Height          =   372
         Index           =   2
         Left            =   1440
         ScaleHeight     =   324
         ScaleWidth      =   324
         TabIndex        =   38
         Top             =   360
         Width           =   372
      End
      Begin VB.PictureBox pctResul 
         AutoRedraw      =   -1  'True
         Height          =   372
         Index           =   1
         Left            =   840
         ScaleHeight     =   324
         ScaleWidth      =   324
         TabIndex        =   37
         Top             =   360
         Width           =   372
      End
      Begin VB.PictureBox pctResul 
         AutoRedraw      =   -1  'True
         Height          =   372
         Index           =   0
         Left            =   240
         ScaleHeight     =   324
         ScaleWidth      =   324
         TabIndex        =   36
         Top             =   360
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   23
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1080
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   22
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1080
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   21
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1080
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   20
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1080
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   19
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   18
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   17
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2160
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   16
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2160
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   15
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3240
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   14
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3240
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   13
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3240
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   12
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3240
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   11
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4320
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   10
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4320
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   9
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4320
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   8
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4320
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   7
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5280
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   6
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5280
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   5
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5280
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5280
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   3
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6120
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6120
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6120
         Width           =   372
      End
      Begin VB.CommandButton cmdResul 
         Height          =   372
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6120
         Width           =   372
      End
   End
End
Attribute VB_Name = "frmMastermind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tour As Integer
Dim Coul(0 To 6) As Long
Dim Resul(0 To 3) As Long
Dim TotalPlacé As Integer
Dim TotalCouleur As Integer
Dim FirstCoul As Long

Private Sub cmdCoul_Click(Index As Integer)
    pct.BackColor = Coul(Index)
End Sub

Private Sub cmdNouvellePartie_Click()
    Dim i As Integer
    
    cmdCoul_Click (0)
    Tour = 0
    
    For i = 0 To cmdResul.UBound
        cmdResul(i).BackColor = QBColor(7)
        pctValidResul(i).Visible = False
    Next
    For i = 4 To cmdResul.UBound
        cmdResul(i).Enabled = False
    Next
    For i = 0 To 3
        cmdResul(i).Enabled = True
    Next
    
    Randomize
    
    For i = 0 To 3
        Resul(i) = QBColor(Int(7 * Rnd))
        pctResul(i).BackColor = FirstCoul
        pctResul(i).Enabled = False
    Next
    cmdValider.Enabled = True
End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Private Sub cmdResul_Click(Index As Integer)
    cmdResul(Index).BackColor = pct.BackColor
End Sub

Private Sub cmdValider_Click()
    Dim i As Integer
    Dim j As Integer
    
    TotalCouleur = 0
    TotalPlacé = 0
    
    For i = 0 To 3
        If cmdResul(Tour * 4 + i).BackColor = Resul(i) Then
            TotalCouleur = TotalCouleur + 1
        Else
            For j = 0 To 3
                If i <> j Then
                    If cmdResul(Tour * 4 + i).BackColor = Resul(j) Then
                        TotalPlacé = TotalPlacé + 1
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    
    For i = 0 To cmdResul.UBound
        cmdResul(i).Enabled = False
    Next
    
    If TotalCouleur = 4 Then ' Partie terminée !
        For i = 0 To 3
            pctResul(i).BackColor = Resul(i)
        Next
        MsgBox "Gagné !", vbOKOnly + vbExclamation, "Gagné !"
    Else
        If TotalCouleur > 0 Then
            For i = 1 To TotalCouleur
                pctValidResul(Tour * 4 + i - 1).BackColor = vbBlack
                pctValidResul(Tour * 4 + i - 1).Visible = True
            Next
        End If
        If TotalPlacé > 0 Then
            For i = 1 To TotalPlacé
                pctValidResul(Tour * 4 + i - 1 + TotalCouleur).BackColor = vbWhite
                pctValidResul(Tour * 4 + i - 1 + TotalCouleur).Visible = True
            Next
        End If
        
        If Tour = 5 Then
            For i = 0 To 3
                pctResul(i).BackColor = Resul(i)
            Next
            MsgBox "Perdu !", vbOKOnly + vbExclamation, "Perdu..."
            cmdValider.Enabled = False
        Else
            For i = (Tour + 1) * 4 To (Tour + 1) * 4 + 3
                cmdResul(i).Enabled = True
            Next
        End If
    End If
    Tour = Tour + 1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 6
        Coul(i) = QBColor(i)
        cmdCoul(i).BackColor = Coul(i)
    Next
    FirstCoul = pctResul(0).BackColor
    cmdNouvellePartie_Click
End Sub

