VERSION 5.00
Begin VB.Form frmArkanobeer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arkonobeer"
   ClientHeight    =   8076
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   9540
   Icon            =   "frmArkanobeer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8076
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Quitter"
      Height          =   612
      Left            =   7680
      TabIndex        =   6
      Top             =   600
      Width           =   1812
   End
   Begin VB.CommandButton cmdCommencer 
      Caption         =   "Commencer une nouvelle partie"
      Height          =   612
      Left            =   7680
      TabIndex        =   5
      Top             =   0
      Width           =   1812
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   5760
   End
   Begin VB.PictureBox Balle 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   252
      Left            =   360
      Picture         =   "frmArkanobeer.frx":08CA
      ScaleHeight     =   204
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   228
   End
   Begin VB.PictureBox Capsule 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   612
      Left            =   360
      Picture         =   "frmArkanobeer.frx":0C3E
      ScaleHeight     =   564
      ScaleWidth      =   1320
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1368
   End
   Begin VB.PictureBox Bouteille 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   372
      Left            =   1440
      Picture         =   "frmArkanobeer.frx":4976
      ScaleHeight     =   324
      ScaleWidth      =   384
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.PictureBox Fond 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8088
      Left            =   0
      Picture         =   "frmArkanobeer.frx":5FBD
      ScaleHeight     =   8040
      ScaleWidth      =   7680
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   7728
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3012
      Left            =   0
      ScaleHeight     =   2964
      ScaleWidth      =   2124
      TabIndex        =   0
      Top             =   480
      Width           =   2172
   End
End
Attribute VB_Name = "frmArkanobeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Depl
    Angle As Double
    X As Integer
    Y As Integer
End Type

Dim CouleurTrans As Long
Dim Début As Boolean
Dim CapsuleX As Integer
Dim CapsuleY As Integer
Dim BalleX As Integer
Dim BalleY As Integer
Dim Tableau(0 To 9, 0 To 4)  As Integer
Dim OldDeplX As Integer
Dim OldDeplY As Integer
Dim TableauDepl(0 To 15) As Depl
Dim BalleDepl As Integer
Dim TableauBouteille(0 To 4, 0 To 4) As POINTAPI
Dim TableauBouteilleV(0 To 4, 0 To 4) As Boolean
Dim TableauPix() As Integer

Private Sub cmdCommencer_Click()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim DecBouteilleX As Integer
    Dim DecBouteilleY As Integer
    Dim IncBouteilleX As Integer
    Dim IncBouteilleY As Integer
    
    BalleX = CapsuleX - Balle.ScaleWidth / 4
    CapsuleY = pct.ScaleHeight - 50
    BalleY = CapsuleY
    Début = True
    DecBouteilleX = 143
    DecBouteilleY = 50
    IncBouteilleX = 75
    IncBouteilleY = 50

    For i = 0 To 4
        For j = 0 To 4
            ' Tableau des coordonnées des bouteilles
            TableauBouteille(i, j).X = i * IncBouteilleX + DecBouteilleX
            TableauBouteille(i, j).Y = j * IncBouteilleY + DecBouteilleY
            ' Tableau qui indique si les bouteilles sont ou non visibles
            TableauBouteilleV(i, j) = True
            ' Enfin, on crée un tableau de la taille de l'image
            ' de fon et on met à 1 les pixels où se trouvent
            ' les bouteilles
            For k = 0 To Bouteille.ScaleWidth - 1
                For l = 0 To Bouteille.ScaleHeight - 1
                    TableauPix(TableauBouteille(i, j).X + k, TableauBouteille(i, j).Y + l) = 1
                Next
            Next
        Next
    Next
End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Public Sub Form_Load()
    Dim i As Integer

    pct.Top = 0
    pct.Left = 0
    pct.Width = Fond.Width
    pct.Height = Fond.Height
    pct.ScaleMode = 3
    Capsule.ScaleMode = 3
    Bouteille.ScaleMode = 3
    Balle.ScaleMode = 3
    CouleurTrans = QBColor(3)
    Timer1.Interval = 10
    Timer1.Enabled = False
    
    ' On copie l'image de fond
    BitBlt pct.hdc, 0, 0, pct.ScaleWidth, pct.ScaleHeight, Fond.hdc, 0, 0, SRCCOPY
    pct.Picture = pct.Image
    
    ReDim TableauPix(0 To pct.ScaleWidth - 1, 0 To pct.ScaleHeight - 1)
    
    ' Construction de la table des déplacements
    For i = 0 To 15
        TableauDepl(i).Angle = PI - i * PI / 8
        TableauDepl(i).X = 10 * Cos(TableauDepl(i).Angle)
        TableauDepl(i).Y = 10 * Sin(TableauDepl(i).Angle)
    Next
    
    cmdCommencer_Click
End Sub

' LA procédure d'affichage
Public Sub DoAffichage()
    Dim i As Integer
    Dim j As Integer
    Dim r As RECT
    Dim r2 As RECT
    
    pct.Cls
    
    ' On copie les canettes visibles sur le fond
    For i = 0 To 4
        For j = 0 To 4
            If TableauBouteilleV(i, j) = True Then
                BitBlt pct.hdc, TableauBouteille(i, j).X, TableauBouteille(i, j).Y, Bouteille.ScaleWidth, Bouteille.ScaleHeight, Bouteille.hdc, 0, 0, SRCCOPY
            End If
        Next
    Next
    
    ' On copie la capsule en transparence car elle
    ' n'est pas rectangulaire
    r.Left = CapsuleX - Capsule.ScaleWidth / 2
    r.Right = r.Left + Capsule.ScaleWidth
    r.Top = CapsuleY
    r.Bottom = r.Top + Capsule.ScaleHeight
    r2.Left = 0
    r2.Top = 0
    r2.Right = Capsule.ScaleWidth
    r2.Bottom = Capsule.ScaleHeight
    TransparentBlt pct.hdc, r, Capsule.hdc, r2, CouleurTrans
    Capsule.Cls
    
    ' De même pour la balle
    r.Left = BalleX
    r.Right = r.Left + Balle.ScaleWidth
    r.Top = BalleY
    r.Bottom = r.Top + Balle.ScaleHeight
    r2.Left = 0
    r2.Top = 0
    r2.Right = Balle.ScaleWidth
    r2.Bottom = Balle.ScaleHeight
    TransparentBlt pct.hdc, r, Balle.hdc, r2, CouleurTrans
    Balle.Cls
    pct.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Lorsqu'on termine, on met la condition d'arrêt à True
    Arret = True
End Sub

' Déplacement de la capsule
Private Sub pct_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        If Début = True Then
            BalleDepl = 4
            Timer1.Enabled = True
            Début = False
        End If
    End If
End Sub

Private Sub pct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CapsuleX = X
    If CapsuleX < Capsule.ScaleWidth / 2 Then
        CapsuleX = Capsule.ScaleWidth / 2
    ElseIf X + Capsule.ScaleWidth / 2 > pct.ScaleWidth Then
        CapsuleX = pct.ScaleWidth - Capsule.ScaleWidth / 2
    End If
    If Début = True Then
        BalleX = CapsuleX
    End If
End Sub

' Déplacement de la balle
Private Sub Timer1_Timer()
    Dim res As Boolean
    BalleX = BalleX + 1 * TableauDepl(BalleDepl).X
    BalleY = BalleY - 1 * TableauDepl(BalleDepl).Y
    res = CheckCollision
    If res = True Then
        BalleX = BalleX + 1 * TableauDepl(BalleDepl).X
        BalleY = BalleY - 1 * TableauDepl(BalleDepl).Y
    End If
End Sub

' La grosse procédure pour gérer les collisions
Private Function CheckCollision() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim TypeCollision As Integer
    Dim res As Boolean
    Dim l As Integer
    Dim m As Integer
    Dim s As Integer
    Dim t As Integer

    ' Premier cas : on regarde si on rentre en collision
    ' avec la capsule
    If BalleY >= CapsuleY And BalleX >= CapsuleX - Capsule.ScaleWidth / 2 And BalleX + Balle.ScaleWidth <= CapsuleX + Capsule.ScaleWidth / 2 Then
        j = BalleX + Balle.ScaleWidth / 2
        If j <= CapsuleX - Capsule.ScaleWidth / 8 Then
            BalleDepl = 1
        ElseIf j <= CapsuleX - Capsule.ScaleWidth / 4 Then
            BalleDepl = 2
        ElseIf j <= CapsuleX - Capsule.ScaleWidth / 2 Then
            BalleDepl = 3
        ElseIf j = CapsuleX - Capsule.ScaleWidth / 2 - Balle.ScaleWidth / 4 Then
            BalleDepl = 4
        ElseIf j <= CapsuleX + Capsule.ScaleWidth / 8 Then
            BalleDepl = 5
        ElseIf j <= CapsuleX - Capsule.ScaleWidth / 4 Then
            BalleDepl = 6
        ElseIf j <= CapsuleX + Capsule.ScaleWidth / 2 Then
            BalleDepl = 7
        End If
        CheckCollision = True
        Exit Function
    End If
    
    ' Sinon, on regarde  sur les bords et sur les canettes
    For i = BalleX To BalleX + Balle.ScaleWidth - 1
        For j = BalleY To BalleY + Balle.ScaleHeight - 1
            If j < 0 Then ' Bordure haute
                TypeCollision = 1
                Exit For
            ElseIf j >= pct.ScaleHeight Then  ' Bordure basse
                TypeCollision = 2
                Exit For
            Else
                If i >= 0 And i <= pct.ScaleWidth - 1 Then
                    If TableauPix(i, j) = 1 Then
                        ' Trouver la bouteille
                        For l = i To 0 Step -1
                            If TableauPix(l, j) = 0 Then
                                l = l + 1
                                Exit For
                            End If
                        Next
                        For m = j To 0 Step -1
                            If TableauPix(i, m) = 0 Then
                                m = m + 1
                                Exit For
                            End If
                        Next
                        For s = 0 To 4
                            For t = 0 To 4
                                If TableauBouteille(s, t).X = l And TableauBouteille(s, t).Y = m Then
                                    TableauBouteilleV(s, t) = False
                                    Exit For
                                End If
                            Next
                        Next
                        For s = l To Bouteille.ScaleWidth + l - 1
                            For t = m To Bouteille.ScaleHeight + m - 1
                                TableauPix(s, t) = 0
                            Next
                        Next
                        TypeCollision = 3
                        Exit For
                    End If
                Else
                    TypeCollision = 3
                    Exit For
                End If
            End If
        Next
        If TypeCollision > 0 Then
            res = True
            Exit For
        End If
    Next
    
    If TypeCollision > 0 Then
        Select Case TypeCollision
            Case 1 ' Bordure haute
                BalleDepl = 16 - BalleDepl
            Case 2 ' Bordure basse
                MsgBox "Perdu !", vbOKOnly + vbExclamation, "Perdu..."
                Timer1.Enabled = False
            Case 3 ' Autre collision
                If (BalleDepl > 0 And BalleDepl < 4) Or (BalleDepl > 4 And BalleDepl < 8) Then
                    BalleDepl = 8 - BalleDepl
                ElseIf (BalleDepl > 8 And BalleDepl < 12) Or (BalleDepl > 12 And BalleDepl < 16) Then
                    BalleDepl = 16 - BalleDepl + 8
                ElseIf BalleDepl = 4 Then BalleDepl = 12
                ElseIf BalleDepl = 12 Then BalleDepl = 4
                ElseIf BalleDepl = 0 Then BalleDepl = 8
                ElseIf BalleDepl = 8 Then BalleDepl = 0
                End If
        End Select
    End If
    CheckCollision = res
End Function
