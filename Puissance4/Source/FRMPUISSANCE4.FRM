VERSION 5.00
Begin VB.Form frmPuissance4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Puissance 4"
   ClientHeight    =   5640
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9780
   Icon            =   "frmPuissance4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Joueur 2"
      Height          =   5412
      Left            =   7680
      TabIndex        =   16
      Top             =   120
      Width           =   2052
      Begin VB.PictureBox pctPièces 
         AutoRedraw      =   -1  'True
         Height          =   3732
         Index           =   1
         Left            =   240
         ScaleHeight     =   3675
         ScaleWidth      =   1395
         TabIndex        =   17
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label Label4 
         Caption         =   "Score :"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre de pièces :"
         Height          =   252
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1812
      End
      Begin VB.Label lblNbPièces 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   1452
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5412
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   5292
      Begin VB.CommandButton cmdCol 
         Caption         =   "7"
         Height          =   372
         Index           =   6
         Left            =   4080
         TabIndex        =   9
         Top             =   720
         Width           =   612
      End
      Begin VB.CommandButton cmdCol 
         Caption         =   "6"
         Height          =   372
         Index           =   5
         Left            =   3480
         TabIndex        =   10
         Top             =   720
         Width           =   612
      End
      Begin VB.CommandButton cmdCol 
         Caption         =   "5"
         Height          =   372
         Index           =   4
         Left            =   2880
         TabIndex        =   11
         Top             =   720
         Width           =   612
      End
      Begin VB.CommandButton cmdCol 
         Caption         =   "4"
         Height          =   372
         Index           =   3
         Left            =   2280
         TabIndex        =   12
         Top             =   720
         Width           =   612
      End
      Begin VB.CommandButton cmdCol 
         Caption         =   "3"
         Height          =   372
         Index           =   2
         Left            =   1680
         TabIndex        =   13
         Top             =   720
         Width           =   612
      End
      Begin VB.CommandButton cmdCol 
         Caption         =   "2"
         Height          =   372
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   720
         Width           =   612
      End
      Begin VB.CommandButton cmdCol 
         Caption         =   "1"
         Height          =   372
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Width           =   612
      End
      Begin VB.PictureBox pct 
         AutoRedraw      =   -1  'True
         Height          =   4212
         Left            =   480
         ScaleHeight     =   4155
         ScaleWidth      =   4155
         TabIndex        =   8
         Top             =   1080
         Width           =   4212
      End
      Begin VB.CommandButton cmdInit 
         Caption         =   "Nouvelle partie"
         Height          =   372
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Joueur 1"
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2052
      Begin VB.PictureBox pctPièces 
         AutoRedraw      =   -1  'True
         Height          =   3732
         Index           =   0
         Left            =   240
         ScaleHeight     =   3675
         ScaleWidth      =   1395
         TabIndex        =   5
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label lblNbPièces 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre de pièces :"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1812
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Score :"
         Height          =   252
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1692
      End
   End
End
Attribute VB_Name = "frmPuissance4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NbPièces(0 To 1) As Integer
Dim Joueur As Integer
Dim Tableau(0 To 6, 0 To 6) As Integer
Dim TaillePièce As Integer

' Click sur un des Command_button :
' Posé d'une pièce
Private Sub cmdCol_Click(Index As Integer)
    PoserPièce (Index)
    DoAffichage
    DoAffichagePile
    ChangerJoueur
End Sub

Private Sub cmdInit_Click()
    ReInit
End Sub

Public Sub Form_Load()
    ReInit
End Sub

' dessin du plateau de jeu et des pièces
Public Sub DoAffichage()
    Dim Ratio As Single
    Dim i As Integer
    Dim j As Integer

    pct.ScaleMode = 3
    pct.Cls
    Ratio = pct.ScaleWidth / 7
    pct.Line (0, 0)-(pct.ScaleWidth, 0), vbBlue
    pct.Line (0, 0)-(0, pct.ScaleHeight), vbBlue
    pct.Line (0, pct.ScaleHeight - 1)-(pct.ScaleWidth, pct.ScaleHeight - 1), vbBlue
    pct.Line (pct.ScaleWidth - 1, pct.ScaleHeight - 1)-(pct.ScaleWidth - 1, 0), vbBlue
    For i = 0 To 6
        pct.Line (0, i * Ratio)-(pct.ScaleWidth, i * Ratio), vbBlue
        pct.Line (i * Ratio, 0)-(i * Ratio, pct.ScaleHeight), vbBlue
    Next
    For i = 0 To 6
        For j = 0 To 6
            If Tableau(i, j) = 1 Then
                pct.FillColor = vbRed
                pct.FillStyle = vbFSSolid
                pct.Circle (i * Ratio + Ratio / 2, j * Ratio + Ratio / 2), Ratio / 2 - 1, vbRed
            ElseIf Tableau(i, j) = 2 Then
                pct.FillColor = vbYellow
                pct.FillStyle = vbFSSolid
                pct.Circle (i * Ratio + Ratio / 2, j * Ratio + Ratio / 2), Ratio / 2 - 1, vbYellow
            End If
        Next
    Next
End Sub

' Changement de tour
Private Sub ChangerJoueur()
    If Joueur = 1 Then
        Joueur = 2
    Else
        Joueur = 1
    End If
End Sub

' Le joueur souhaite poser une pièce :
' on vérifie préalablement que c'est possible
' et le cas échéant, on choisit la bonne case
Private Sub PoserPièce(col As Integer)
    Dim i As Integer
    
    For i = 6 To 0 Step -1
        If Tableau(col, i) = 0 Then
            Exit For
        End If
    Next
    If i <> -1 Then
        Tableau(col, i) = Joueur
        NbPièces(Joueur - 1) = NbPièces(Joueur - 1) - 1
        CheckGagnant col, i
    Else
        Beep
    End If
End Sub

' Procédure clé : il faut vérifier
' si quatre pions de la même couleur sont alignés !
Private Sub CheckGagnant(l, c)
    Dim somme As Integer
    Dim i As Integer
    Dim j As Integer
    
    ' Diagonale 1
    somme = 0
    For i = -3 To 3 Step 1
        If (l + i) >= 0 And (l + i) <= 6 And (c + i) >= 0 And (c + i) <= 6 Then
            If Tableau(l + i, c + i) = Joueur Then
                somme = somme + 1
                If somme = 4 Then Exit For
            Else
                somme = 0
            End If
        End If
    Next
    If somme = 4 Then
        Winner
        Exit Sub
    End If
    
    ' Diagonale 2 : Ligne
    somme = 0
    For i = -3 To 3 Step 1
        If (l + i) >= 0 And (l + i) <= 6 Then
            If Tableau(l + i, c) = Joueur Then
                somme = somme + 1
                If somme = 4 Then Exit For
            Else
                somme = 0
            End If
        End If
    Next
    If somme = 4 Then
        Winner
        Exit Sub
    End If
    
    ' Diagonale 3 : Ligne en bas
    If c <= 3 Then
        somme = 4
        For i = 1 To 3
            If Tableau(l, c + i) <> Joueur Then
                somme = 0
                Exit For
            End If
        Next
        If somme = 4 Then
            Winner
            Exit Sub
        End If
    End If

    ' Diagonale 4 : idem diagonale 1, mais dans l'autre sens
    somme = 0
    For i = -3 To 3 Step 1
        If (l - i) >= 0 And (l - i) <= 6 And (c + i) >= 0 And (c + i) <= 6 Then
            If Tableau(l - i, c + i) = Joueur Then
                somme = somme + 1
                If somme = 4 Then Exit For
            Else
                somme = 0
            End If
        End If
    Next
    If somme = 4 Then
        Winner
        Exit Sub
    End If
End Sub

' Initialisation du jeu
Private Sub ReInit()
    Dim i As Integer
    Dim j As Integer
    
    Joueur = 1
        
    NbPièces(0) = 24
    NbPièces(1) = 24
    lblNbPièces(0) = 24
    lblNbPièces(1) = 24
    
    For i = 0 To 6
        For j = 0 To 6
            Tableau(i, j) = 0
        Next
    Next
    For i = 0 To 6
        cmdCol(i).Enabled = True
    Next
    pctPièces(0).ScaleMode = 3
    TaillePièce = (pctPièces(0).ScaleHeight - 20) / NbPièces(0)
    
    DoAffichage
    DoAffichagePile
    Joueur = 2
    DoAffichagePile
    Joueur = 1
End Sub

' On va ici afficher la pile de pièces du joueur qui
' vient d'en poser une
Private Sub DoAffichagePile()
    Dim TailleH As Integer
    Dim i As Integer

    pctPièces(Joueur - 1).Cls
    If Joueur = 1 Then
        pctPièces(Joueur - 1).FillStyle = vbFSSolid
        pctPièces(Joueur - 1).FillColor = vbRed
    Else
        pctPièces(Joueur - 1).FillStyle = vbFSSolid
        pctPièces(Joueur - 1).FillColor = vbYellow
    End If
    lblNbPièces(Joueur - 1).Caption = CStr(NbPièces(Joueur - 1))
    With pctPièces(Joueur - 1)
        .ScaleMode = 3
        ' On laisse 10 pixels de marge en haut et en bas
        TailleH = .ScaleHeight
        For i = 1 To NbPièces(Joueur - 1)
            pctPièces(Joueur - 1).Line (5, TailleH - TaillePièce * i)-(.ScaleWidth - 5, TailleH - TaillePièce * (i + 1)), vbBlue, B
        Next
    End With
End Sub

' Ca y est, un joueur a gagné !
Private Sub Winner()
    Dim i As Integer
    
    DoAffichage
    DoAffichagePile
    MsgBox "Le joueur " & CStr(Joueur) & " a gagné !", vbOKOnly
    
    ' On bloque tout
    For i = 0 To 6
        cmdCol(i).Enabled = False
    Next
    ' On icrémente le score
    lblScore(Joueur - 1).Caption = CStr(CInt(lblScore(Joueur - 1).Caption) + 1)
End Sub

