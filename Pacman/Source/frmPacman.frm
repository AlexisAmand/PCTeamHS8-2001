VERSION 5.00
Begin VB.Form frmPacman 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pacman"
   ClientHeight    =   5964
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5220
   Icon            =   "frmPacman.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5964
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pacgum1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   156
      Left            =   2160
      Picture         =   "frmPacman.frx":08CA
      ScaleHeight     =   108
      ScaleWidth      =   108
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   156
   End
   Begin VB.PictureBox Ghost 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   336
      Index           =   1
      Left            =   2760
      Picture         =   "frmPacman.frx":0A08
      ScaleHeight     =   288
      ScaleWidth      =   276
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.PictureBox Ghost 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   336
      Index           =   0
      Left            =   2400
      Picture         =   "frmPacman.frx":110A
      ScaleHeight     =   288
      ScaleWidth      =   276
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.PictureBox Pac 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   384
      Index           =   1
      Left            =   1440
      Picture         =   "frmPacman.frx":180C
      ScaleHeight     =   336
      ScaleWidth      =   324
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Quitter"
      Height          =   372
      Left            =   3240
      TabIndex        =   2
      Top             =   5520
      Width           =   1812
   End
   Begin VB.CommandButton cmdNouvellePartie 
      Caption         =   "Nouvelle partie"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   1692
   End
   Begin VB.PictureBox fond 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5436
      Left            =   0
      Picture         =   "frmPacman.frx":217E
      ScaleHeight     =   5388
      ScaleWidth      =   5184
      TabIndex        =   0
      Top             =   0
      Width           =   5232
      Begin VB.Timer Timer1 
         Left            =   3960
         Top             =   3360
      End
      Begin VB.PictureBox Pac 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   384
         Index           =   7
         Left            =   2760
         Picture         =   "frmPacman.frx":902D2
         ScaleHeight     =   336
         ScaleWidth      =   324
         TabIndex        =   14
         Top             =   3840
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.PictureBox Pac 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   384
         Index           =   6
         Left            =   2400
         Picture         =   "frmPacman.frx":90C44
         ScaleHeight     =   336
         ScaleWidth      =   324
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.PictureBox Pac 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   384
         Index           =   5
         Left            =   1920
         Picture         =   "frmPacman.frx":915B6
         ScaleHeight     =   336
         ScaleWidth      =   324
         TabIndex        =   12
         Top             =   3600
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.PictureBox Pac 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   384
         Index           =   4
         Left            =   1560
         Picture         =   "frmPacman.frx":91F28
         ScaleHeight     =   336
         ScaleWidth      =   324
         TabIndex        =   11
         Top             =   3480
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.PictureBox Pac 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   384
         Index           =   3
         Left            =   2160
         Picture         =   "frmPacman.frx":9289A
         ScaleHeight     =   336
         ScaleWidth      =   324
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.PictureBox Pac 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   384
         Index           =   2
         Left            =   1800
         Picture         =   "frmPacman.frx":9320C
         ScaleHeight     =   336
         ScaleWidth      =   324
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.PictureBox Pacgum2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   336
         Left            =   2760
         Picture         =   "frmPacman.frx":93B7E
         ScaleHeight     =   288
         ScaleWidth      =   276
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   324
      End
      Begin VB.PictureBox Pac 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   384
         Index           =   0
         Left            =   1080
         Picture         =   "frmPacman.frx":94282
         ScaleHeight     =   336
         ScaleWidth      =   324
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   372
      End
   End
End
Attribute VB_Name = "frmPacman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DecX As Integer
Dim DecY As Integer
Dim PacX As Integer
Dim PacY As Integer
Dim GhostX As Integer
Dim GhostY As Integer
Dim Tableau(0 To 8, 0 To 8) As Integer
Dim TableauGum(0 To 8, 0 To 8) As Integer
Dim TableauMurs(0 To 16, 0 To 16) As Integer
Dim PacPos As Integer
Dim GhostPos As Integer
Dim TypePos As Integer
Dim CouleurTrans As Long
Dim Direction As Integer
Dim DirectionTime As Integer
Dim ClavierBloqué As Boolean

Private Sub cmdNouvellePartie_Click()
    Randomize
    PacX = 0
    PacY = 0
    GhostX = 8
    GhostY = 8
    PacPos = 0
    TypePos = 0
    GhostPos = 0
    DirectionTime = 1
    ClavierBloqué = False
    InitTableauGum
    DoAffichage
    Timer1.Enabled = True
End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Private Sub fond_KeyDown(KeyCode As Integer, Shift As Integer)
    If ClavierBloqué = False Then
        Select Case KeyCode
            Case 37 ' Gauche
                If TypePos <> 0 Then
                    TypePos = 0
                    PacPos = 0
                Else
                    PacPos = 1 - PacPos
                End If
                If PacX - 1 >= 0 Then
                    If TableauMurs((PacX - 1) * 2 + 1, PacY * 2) = 0 Then
                        PacX = PacX - 1
                        TableauGum(PacX, PacY) = 0
                    End If
                End If
            Case 39 ' Droite
                If TypePos <> 1 Then
                    TypePos = 1
                    PacPos = 0
                Else
                    PacPos = 1 - PacPos
                End If
                If PacX + 1 <= 8 Then
                    If TableauMurs(PacX * 2 + 1, PacY * 2) = 0 Then
                        PacX = PacX + 1
                        TableauGum(PacX, PacY) = 0
                    End If
                End If
            Case 38 ' Haut
                If TypePos <> 2 Then
                    TypePos = 2
                    PacPos = 0
                Else
                    PacPos = 1 - PacPos
                End If
                If PacY - 1 >= 0 Then
                    If TableauMurs(PacX * 2, (PacY - 1) * 2 + 1) = 0 Then
                        PacY = PacY - 1
                        TableauGum(PacX, PacY) = 0
                    End If
                End If
            Case 40 ' Bas
                If TypePos <> 3 Then
                    TypePos = 3
                    PacPos = 0
                Else
                    PacPos = 1 - PacPos
                End If
                If PacY + 1 <= 8 Then
                    If TableauMurs(PacX * 2, PacY * 2 + 1) = 0 Then
                        PacY = PacY + 1
                        TableauGum(PacX, PacY) = 0
                    End If
                End If
        End Select
        DoAffichage
        CheckFin
    End If
End Sub

Private Sub Form_Load()
    Dim diff As Integer
    
    fond.ScaleMode = 3
    Pac(0).ScaleMode = 3
    Pac(1).ScaleMode = 3
    Pacgum1.ScaleMode = 3
    Pacgum2.ScaleMode = 3
    Ghost(0).ScaleMode = 3
    Ghost(1).ScaleMode = 3
    cmdNouvellePartie.Top = fond.Height
    cmdQuitter.Top = fond.Height
    diff = Me.Width - Me.ScaleWidth
    Me.Width = fond.Width + diff
    diff = Me.Height - Me.ScaleHeight
    Me.Height = fond.Height + diff + cmdNouvellePartie.Height
    

    CouleurTrans = QBColor(3)

    DecX = 16
    DecY = 24
    
    Timer1.Interval = 400
    Timer1.Enabled = False
    ConstructionTableauMurs
    
    cmdNouvellePartie_Click
End Sub

Private Sub DoAffichage()
    Dim i As Integer
    Dim j As Integer
    Dim r As RECT
    Dim r2 As RECT

    fond.Cls
    
    r.Left = PacX * 46 + DecX
    r.Right = r.Left + Pac(0).ScaleWidth
    r.Top = PacY * 46 + DecY
    r.Bottom = r.Top + Pac(0).ScaleHeight
    r2.Left = 0
    r2.Top = 0
    r2.Right = Pac(0).ScaleWidth
    r2.Bottom = Pac(0).ScaleHeight
    TransparentBlt fond.hdc, r, Pac(TypePos * 2 + PacPos).hdc, r2, CouleurTrans
    Pac(TypePos * 2 + PacPos).Cls
    
    r.Left = GhostX * 46 + DecX
    r.Right = r.Left + Ghost(GhostPos).ScaleWidth
    r.Top = GhostY * 46 + DecY
    r.Bottom = r.Top + Ghost(GhostPos).ScaleHeight
    r2.Left = 0
    r2.Top = 0
    r2.Right = Ghost(GhostPos).ScaleWidth
    r2.Bottom = Ghost(GhostPos).ScaleHeight
    TransparentBlt fond.hdc, r, Ghost(GhostPos).hdc, r2, CouleurTrans
    Ghost(GhostPos).Cls
    
    For i = 0 To 8
        For j = 0 To 8
            If TableauGum(i, j) <> 0 Then
                r.Left = i * 46 + DecX + 10
                r.Right = r.Left + Pacgum1.ScaleWidth
                r.Top = j * 46 + DecY + 10
                r.Bottom = r.Top + Pacgum1.ScaleHeight
                r2.Left = 0
                r2.Top = 0
                r2.Right = Pacgum1.ScaleWidth
                r2.Bottom = Pacgum1.ScaleHeight
                TransparentBlt fond.hdc, r, Pacgum1.hdc, r2, CouleurTrans
                Pacgum1.Cls
            End If
        Next
    Next
End Sub

Private Sub ConstructionTableauMurs()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 16
        For j = 0 To 16
            TableauMurs(i, j) = 0
        Next
    Next
    TableauMurs(6, 1) = 1
    TableauMurs(8, 1) = 1
    TableauMurs(10, 1) = 1
    TableauMurs(14, 1) = 1
    TableauMurs(2, 3) = 1
    TableauMurs(8, 3) = 1
    TableauMurs(0, 5) = 1
    TableauMurs(4, 5) = 1
    TableauMurs(6, 5) = 1
    TableauMurs(10, 5) = 1
    TableauMurs(12, 5) = 1
    TableauMurs(16, 5) = 1
    TableauMurs(0, 7) = 1
    TableauMurs(2, 7) = 1
    TableauMurs(14, 7) = 1
    TableauMurs(16, 7) = 1
    TableauMurs(2, 9) = 1
    TableauMurs(4, 9) = 1
    TableauMurs(6, 9) = 1
    TableauMurs(8, 9) = 1
    TableauMurs(10, 9) = 1
    TableauMurs(12, 9) = 1
    TableauMurs(4, 11) = 1
    TableauMurs(12, 11) = 1
    TableauMurs(14, 11) = 1
    TableauMurs(0, 13) = 1
    TableauMurs(4, 13) = 1
    TableauMurs(6, 13) = 1
    TableauMurs(10, 13) = 1
    TableauMurs(12, 13) = 1
    TableauMurs(16, 13) = 1
    TableauMurs(2, 15) = 1
    TableauMurs(4, 15) = 1
    TableauMurs(8, 15) = 1
    TableauMurs(12, 15) = 1
    TableauMurs(14, 15) = 1
    TableauMurs(1, 2) = 1
    TableauMurs(3, 2) = 1
    TableauMurs(5, 2) = 1
    TableauMurs(11, 2) = 1
    TableauMurs(13, 2) = 1
    TableauMurs(15, 2) = 1
    TableauMurs(1, 6) = 1
    TableauMurs(5, 6) = 1
    TableauMurs(11, 6) = 1
    TableauMurs(15, 6) = 1
    TableauMurs(5, 8) = 1
    TableauMurs(11, 8) = 1
    TableauMurs(1, 10) = 1
    TableauMurs(3, 10) = 1
    TableauMurs(5, 10) = 1
    TableauMurs(11, 10) = 1
    TableauMurs(13, 10) = 1
    TableauMurs(15, 10) = 1
    TableauMurs(7, 12) = 1
    TableauMurs(9, 12) = 1
    TableauMurs(3, 14) = 1
    TableauMurs(5, 14) = 1
    TableauMurs(11, 14) = 1
    TableauMurs(13, 14) = 1
    TableauMurs(7, 16) = 1
    TableauMurs(9, 16) = 1
End Sub

Private Sub InitTableauGum()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 8
        For j = 0 To 8
            If Not (i = 0 And j = 0) And Not (i = 8 And j = 8) Then
                TableauGum(i, j) = 1
            End If
        Next
    Next
End Sub

Private Sub Timer1_Timer()
    
    If DirectionTime = 1 Then
        Randomize
        Direction = Int(4 * Rnd + 1)
        DirectionTime = 0
    Else
        DirectionTime = 1
    End If
    Select Case Direction
        Case 1 ' Gauche
            If GhostX - 1 >= 0 Then
                GhostX = GhostX - 1
            End If
        Case 2 ' Droite
            If GhostX + 1 <= 8 Then
                GhostX = GhostX + 1
            End If
        Case 3 ' Haut
            If GhostY - 1 >= 0 Then
                GhostY = GhostY - 1
            End If
        Case 4 ' Bas
            If GhostY + 1 <= 8 Then
                GhostY = GhostY + 1
            End If
    End Select
    GhostPos = 1 - GhostPos
    TableauGum(GhostX, GhostY) = 0
    DoAffichage
    CheckFin
End Sub

Private Sub CheckFin()
    If PacX = GhostX And PacY = GhostY Then
        Timer1.Enabled = False
        ClavierBloqué = True
        MsgBox "Essaye encore petit scarabée !", vbOKOnly + vbExclamation, "Perdu ..."
    End If
End Sub
