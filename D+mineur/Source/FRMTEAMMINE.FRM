VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TeamMine 1.0"
   ClientHeight    =   3768
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   2436
   Icon            =   "frmTeamMine.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   120
   End
   Begin VB.CommandButton cmdTime 
      Height          =   372
      Left            =   1560
      TabIndex        =   66
      Top             =   960
      Width           =   612
   End
   Begin VB.CommandButton cmdNew 
      Height          =   372
      Left            =   960
      Picture         =   "frmTeamMine.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   492
   End
   Begin VB.CommandButton cmdNBMines 
      Caption         =   "10"
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   612
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   63
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   62
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   61
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   60
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   59
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   58
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   57
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   56
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3240
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   55
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   54
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   53
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   52
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   51
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   50
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   49
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   48
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3000
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   47
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   46
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   45
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   44
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   43
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   42
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   41
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   40
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2760
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   39
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   38
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   37
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   36
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   35
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   34
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   33
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   32
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2520
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   31
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   30
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   29
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   28
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   27
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   26
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   25
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   24
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2280
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   23
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   22
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   21
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   20
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   19
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   18
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   17
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   16
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2040
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   15
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   14
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   13
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   12
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   11
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   10
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   9
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   8
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   7
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   6
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   5
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   4
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   3
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   252
   End
   Begin VB.CommandButton cmdMine 
      Height          =   252
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   252
   End
   Begin VB.Line Line24 
      BorderColor     =   &H80000014&
      X1              =   10
      X2              =   190
      Y1              =   121
      Y2              =   121
   End
   Begin VB.Line Line23 
      BorderColor     =   &H8000000C&
      X1              =   10
      X2              =   189
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Image imgLogo 
      Height          =   540
      Left            =   240
      Picture         =   "frmTeamMine.frx":0DF8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1920
   End
   Begin VB.Image imgMauvaisChoix 
      Height          =   180
      Left            =   3000
      Picture         =   "frmTeamMine.frx":629A
      Top             =   2280
      Width           =   168
   End
   Begin VB.Image imgOk 
      Height          =   252
      Left            =   3000
      Picture         =   "frmTeamMine.frx":6570
      Top             =   1920
      Width           =   240
   End
   Begin VB.Image imgGagné 
      Height          =   240
      Left            =   3000
      Picture         =   "frmTeamMine.frx":6A9E
      Top             =   1560
      Width           =   252
   End
   Begin VB.Image imgoh 
      Height          =   240
      Left            =   3000
      Picture         =   "frmTeamMine.frx":6FE0
      Top             =   1200
      Width           =   228
   End
   Begin VB.Image imgPerdu 
      Height          =   252
      Left            =   3000
      Picture         =   "frmTeamMine.frx":74D2
      Top             =   840
      Width           =   228
   End
   Begin VB.Image imgDrapeau 
      Height          =   132
      Left            =   3000
      Picture         =   "frmTeamMine.frx":7A00
      Top             =   600
      Width           =   132
   End
   Begin VB.Image imgExplosé 
      Height          =   180
      Left            =   3000
      Picture         =   "frmTeamMine.frx":7BCE
      Top             =   360
      Width           =   168
   End
   Begin VB.Image imgBombe 
      Height          =   168
      Left            =   3000
      Picture         =   "frmTeamMine.frx":7EA4
      Top             =   120
      Width           =   156
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FFFFFF&
      X1              =   2
      X2              =   2
      Y1              =   5
      Y2              =   310
   End
   Begin VB.Line Line21 
      BorderColor     =   &H8000000C&
      X1              =   1
      X2              =   1
      Y1              =   5
      Y2              =   310
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FFFFFF&
      X1              =   200
      X2              =   200
      Y1              =   5
      Y2              =   310
   End
   Begin VB.Line Line19 
      BorderColor     =   &H8000000C&
      X1              =   199
      X2              =   199
      Y1              =   5
      Y2              =   310
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   200
      Y1              =   311
      Y2              =   311
   End
   Begin VB.Line Line17 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   200
      Y1              =   310
      Y2              =   310
   End
   Begin VB.Line Line16 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   200
      Y1              =   5
      Y2              =   5
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   200
      Y1              =   6
      Y2              =   6
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000C&
      X1              =   189
      X2              =   189
      Y1              =   59
      Y2              =   13
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000C&
      X1              =   9
      X2              =   9
      Y1              =   59
      Y2              =   13
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000014&
      X1              =   190
      X2              =   190
      Y1              =   14
      Y2              =   60
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000014&
      X1              =   10
      X2              =   10
      Y1              =   14
      Y2              =   60
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000014&
      X1              =   10
      X2              =   190
      Y1              =   14
      Y2              =   14
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000C&
      X1              =   10
      X2              =   190
      Y1              =   13
      Y2              =   13
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   189
      X2              =   189
      Y1              =   299
      Y2              =   70
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   9
      X2              =   9
      Y1              =   299
      Y2              =   70
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   10
      X2              =   190
      Y1              =   299
      Y2              =   299
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      X1              =   10
      X2              =   10
      Y1              =   300
      Y2              =   70
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   10
      X2              =   190
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   190
      X2              =   190
      Y1              =   300
      Y2              =   70
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   10
      X2              =   189
      Y1              =   69
      Y2              =   69
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   10
      X2              =   190
      Y1              =   70
      Y2              =   70
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' voici les deux tableaux de la carte de jeu : un qui indique si la case est marquée d'un drapeau
' et l'autre qui contient les emplacements des bombes et les valeurs de chaque case
Dim Drapeau(0 To 63) As Integer '0 = rien, 1 = drapeau
Dim Carte(0 To 63) As Integer   '0=rien, 1=1 bombe à côté, ... , 10 = 1 bombe
Dim Perdu As Boolean
Dim Victoire As Boolean
Dim Début As Boolean
Dim TempsEcoulé As Integer
Dim Table(0 To 7, 0 To 7) As Integer

' L'utilisateur enfonce un bouton sur la carte : on analyse ici seulement le bouton droit de la souris
Private Sub cmdMine_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NombreMine As Integer
    
    ' Si on commence une partie, on met en route le Timer
    If Début = True Then
        Début = False
        Timer1.Interval = 1000
        Timer1.Enabled = True
    End If
    
    If (Perdu = False And Victoire = False) Then
        ' on change l'image du bouton central
        cmdNew.Picture = imgoh.Picture
        ' si le bouton enfoncé est le droit
        If Button And vbRightButton Then
            ' s'il n'y a pas de drapeau, on en pose un, et on supprime 1 au nombre de drapeaux restants
            ' sinon, on ôte le drapeau
            If Drapeau(Index) = 0 Then
                If Val(cmdNBMines.Caption) > 0 Then
                    ' affichage du drapeau
                    cmdMine(Index).Picture = imgDrapeau.Picture
                    Drapeau(Index) = 1
                    NombreMine = Val(cmdNBMines.Caption) - 1
                    If NombreMine = 0 Then TestVictoire
                    cmdNBMines.Caption = Str(NombreMine)
                End If
             Else
                ' on ôte le drapeau de l'affichage et de la carte
                cmdMine(Index).Picture = Nothing
                Drapeau(Index) = 0
                ' on rajoute +1 au nombre de mines restantes
                NombreMine = Val(cmdNBMines.Caption) + 1
                cmdNBMines.Caption = Str(NombreMine)
            End If
        End If
    End If
End Sub

' L'utilisateur relâche le bouton de la souris
Private Sub cmdMine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si on commence une partie, on met en route le Timer
    If Début = True Then
        Début = False
        Timer1.Interval = 1000
        Timer1.Enabled = True
    End If
    If (Perdu = False And Victoire = False) Then
        ' on rechange l'image du bouton central...
        cmdNew.Picture = imgOk.Picture
        ' Bouton gauche uniquement
        If Button And vbLeftButton Then
            If Drapeau(Index) = 0 Then
                ' s'il n'y a pas de bombe, on affiche la case, sinon on termine la partie
                If Table(Index Mod 8, Index \ 8) <> 10 Then
                    TestVoisinageBis Index Mod 8, Index \ 8
                    cmdNew.SetFocus
                Else
                    Call FinDePartie(Index)
                End If
            End If
        End If
    End If
End Sub

' Nouvelle partie
Private Sub cmdNew_Click()
    cmdNBMines.Caption = "10"
    Victoire = False
    Perdu = False
    InitDémineur
End Sub

' Chargement de la feuille
Private Sub Form_Load()
    InitDémineur
End Sub

' Procédure d'initialisation de la carte
Private Sub InitDémineur()
    Dim Place As Integer
    Dim Compteur As Integer
    Dim TotalMine As Integer
    Dim x0 As Integer
    Dim y0 As Integer
    Dim x1 As Integer
    Dim y1 As Integer
    Dim Largeur As Integer
    Dim Hauteur As Integer
    ' Phase 1 : Dessin des lignes sous les boutons de commande
    Début = True
    TempsEcoulé = 0
    cmdTime.Caption = "0"
    Largeur = cmdMine(0).Width
    Hauteur = cmdMine(0).Height
    x0 = cmdMine(0).Left
    y0 = cmdMine(0).Top
    x1 = cmdMine(63).Left + Largeur - 1
    y1 = cmdMine(63).Top + Hauteur - 1
    Form1.Cls
    For Compteur = 1 To 7
        Form1.Line (x0 + Compteur * (Largeur - 1), y0)-(x0 + Compteur * (Largeur - 1), y1), vbBlack
        Form1.Line (x0, y0 + Compteur * (Hauteur - 1))-(x1, y0 + Compteur * (Hauteur - 1)), vbBlack
    Next
    Form1.Line (x0, y0)-(x1, y1), vbBlack, B
    cmdNew.Picture = imgOk.Picture
    ' on remet tout à 0
    For Compteur = 0 To 63
        Carte(Compteur) = 0
        Drapeau(Compteur) = 0
        cmdMine(Compteur).Visible = True
        cmdMine(Compteur).Picture = Nothing
    Next
    Compteur = 0
    Randomize
    ' Phase 2 :Initialisation de la carte
    Do
        Place = Int(63 * Rnd)
        If Carte(Place) = 0 Then
            Carte(Place) = 10
            Compteur = Compteur + 1
        End If
    Loop While Compteur <= 9
    Colorise
End Sub

' Cette fonction compte le nombre de mines adjacentes à une case,
' puis écrit avec une couleur donnée sur la fenêtre
Private Sub Colorise()
Dim i As Integer, j As Integer
    For i = 0 To 63
        Table(i Mod 8, i \ 8) = Carte(i)
    Next
    For i = 0 To 7
        For j = 0 To 7
            If Table(i, j) <> 10 Then
                Table(i, j) = CompteMines(i, j)
            End If
            Form1.CurrentX = cmdMine(i + j * 8).Left + 3
            Form1.CurrentY = cmdMine(i + j * 8).Top + 2
            If Table(i, j) <> 0 And Table(i, j) <> 10 Then
                If Table(i, j) = 1 Then Form1.ForeColor = vbBlue
                If Table(i, j) = 2 Then Form1.ForeColor = vbRed
                If Table(i, j) = 3 Then Form1.ForeColor = vbGreen
                If Table(i, j) = 4 Then Form1.ForeColor = vbYellow
                If Table(i, j) = 5 Then Form1.ForeColor = vbCyan
                If Table(i, j) = 6 Then Form1.ForeColor = vbMagenta
                If Table(i, j) = 7 Then Form1.ForeColor = vbWhite
                If Table(i, j) = 8 Then Form1.ForeColor = 2
                Form1.Print Table(i, j)
            End If
        Next
    Next
End Sub

' Cette fonction compte le nombre de mines qui entourent
' la case (i,j)
Public Function CompteMines(i As Integer, j As Integer) As Integer
    Dim TempVar As Integer
    If i > 0 Then
        TempVar = IIf(Table(i - 1, j) = 10, TempVar + 1, TempVar)
    End If
    If i < 7 Then
        TempVar = IIf(Table(i + 1, j) = 10, TempVar + 1, TempVar)
    End If
    If j > 0 Then
        TempVar = IIf(Table(i, j - 1) = 10, TempVar + 1, TempVar)
    End If
    If j < 7 Then
        TempVar = IIf(Table(i, j + 1) = 10, TempVar + 1, TempVar)
    End If
    If i > 0 And j > 0 Then
        TempVar = IIf(Table(i - 1, j - 1) = 10, TempVar + 1, TempVar)
    End If
    If i < 7 And j > 0 Then
        TempVar = IIf(Table(i + 1, j - 1) = 10, TempVar + 1, TempVar)
    End If
    If j < 7 And i > 0 Then
        TempVar = IIf(Table(i - 1, j + 1) = 10, TempVar + 1, TempVar)
    End If
    If i < 7 And j < 7 Then
        TempVar = IIf(Table(i + 1, j + 1) = 10, TempVar + 1, TempVar)
    End If
    CompteMines = TempVar
End Function

' Fin de la partie : on vérifie pour chaque drapeau s'il y avait une bombe en dessous ou non
Private Sub FinDePartie(Nombre As Integer)
Dim Compteur As Integer
    For Compteur = 0 To 63
        If Carte(Compteur) = 10 Then cmdMine(Compteur).Picture = imgBombe.Picture
        If (Carte(Compteur) <> 10 And Drapeau(Compteur) = 1) Then cmdMine(Compteur).Picture = imgMauvaisChoix.Picture
    Next
    cmdMine(Nombre).Picture = imgExplosé.Picture
    cmdNew.Picture = imgPerdu.Picture
    Timer1.Enabled = False
    Perdu = True
End Sub

' Les dix drapeaux ont été posés : on vérifie si ils sont au bon endroit
Private Sub TestVictoire()
Dim Compteur As Integer
Dim TotalJuste As Integer
    For Compteur = 0 To 63
        If (Carte(Compteur) = 10 And Drapeau(Compteur) = 1) Then TotalJuste = TotalJuste + 1
    Next
    If TotalJuste = 10 Then
        Timer1.Enabled = False
        For Compteur = 0 To 63
            If Carte(Compteur) <> 10 Then cmdMine(Compteur).Visible = False
        Next
        cmdNew.Picture = imgGagné.Picture
        Victoire = True
    End If
End Sub

' Evènement Timer : l'intervalle donné (1000 ms) s'est écoulé : on rajoute 1s à l'affichage du temps
Private Sub Timer1_Timer()
    TempsEcoulé = TempsEcoulé + 1
    ' on remet à jour le contrôle affichant le temps
    cmdTime.Caption = TempsEcoulé
End Sub

' Fonction récursive : c'est elle qui découvre les cases
Public Sub TestVoisinageBis(i As Integer, j As Integer)
    cmdMine(i + j * 8).Visible = False
    If Table(i, j) <> 0 Then Exit Sub
    ' On teste ici la position de la case (i,j) pour voir si elle
    ' se situe ou non sur une bordure, ceci afin de tester dans les
    ' huit directions
    If i > 0 Then ' Gauche
        ' On vérifie préalablement si la case n'est pas découverte
        ' comme à chaque fois
        If cmdMine(i - 1 + j * 8).Visible = True Then
            If Table(i - 1, j) <> 10 Then
                cmdMine(i - 1 + j * 8).Visible = False
                If Table(i - 1, j) = 0 Then TestVoisinageBis i - 1, j
            End If
        End If
    End If
    If i < 7 Then ' Droite
        If cmdMine(i + 1 + j * 8).Visible = True Then
            If Table(i + 1, j) <> 10 Then
                cmdMine(i + 1 + j * 8).Visible = False
                If Table(i + 1, j) = 0 Then TestVoisinageBis i + 1, j
            End If
        End If
    End If
    If j > 0 Then ' Haut
        If cmdMine(i + (j - 1) * 8).Visible = True Then
            If Table(i, j - 1) <> 10 Then
                cmdMine(i + (j - 1) * 8).Visible = False
                If Table(i, j - 1) = 0 Then TestVoisinageBis i, j - 1
            End If
        End If
    End If
    If j < 7 Then ' Bas
        If cmdMine(i + (j + 1) * 8).Visible = True Then
            If Table(i, j + 1) <> 10 Then
                cmdMine(i + (j + 1) * 8).Visible = False
                If Table(i, j + 1) = 0 Then TestVoisinageBis i, j + 1
            End If
        End If
    End If
    If i > 0 And j > 0 Then ' Haut gauche
        If cmdMine(i - 1 + (j - 1) * 8).Visible = True Then
            If Table(i - 1, j - 1) <> 10 Then
                cmdMine(i - 1 + (j - 1) * 8).Visible = False
                If Table(i - 1, j - 1) = 0 Then TestVoisinageBis i - 1, j - 1
            End If
        End If
    End If
    If i < 7 And j > 0 Then ' Haut droite
        If cmdMine(i + 1 + (j - 1) * 8).Visible = True Then
            If Table(i + 1, j - 1) <> 10 Then
                cmdMine(i + 1 + (j - 1) * 8).Visible = False
                If Table(i + 1, j - 1) = 0 Then TestVoisinageBis i + 1, j - 1
            End If
        End If
    End If
    If j < 7 And i > 0 Then ' Bas gauche
        If cmdMine(i - 1 + (j + 1) * 8).Visible = True Then
            If Table(i - 1, j + 1) <> 10 Then
                cmdMine(i - 1 + (j + 1) * 8).Visible = False
                If Table(i - 1, j + 1) = 0 Then TestVoisinageBis i - 1, j + 1
            End If
        End If
    End If
    If i < 7 And j < 7 Then ' Bas droite
        If cmdMine(i + 1 + (j + 1) * 8).Visible = True Then
            If Table(i + 1, j + 1) <> 10 Then
                cmdMine(i + 1 + (j + 1) * 8).Visible = False
                If Table(i + 1, j + 1) = 0 Then TestVoisinageBis i + 1, j + 1
            End If
        End If
    End If
End Sub

