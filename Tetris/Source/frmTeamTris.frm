VERSION 5.00
Begin VB.Form frmTeamTris 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TeamTris"
   ClientHeight    =   6768
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6132
   Icon            =   "frmTeamTris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   564
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   511
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6750
      Left            =   0
      ScaleHeight     =   6708
      ScaleWidth      =   4452
      TabIndex        =   1
      Top             =   0
      Width           =   4500
   End
   Begin VB.CommandButton cmdCommencer 
      Caption         =   "Commencer une partie"
      Height          =   612
      Left            =   4560
      TabIndex        =   0
      Top             =   0
      Width           =   1572
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   120
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score :"
      Height          =   732
      Left            =   4560
      TabIndex        =   2
      Top             =   720
      Width           =   1572
   End
End
Attribute VB_Name = "frmTeamTris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Pièce1
    x As Integer
    y As Integer
    Largeur As Integer
    Hauteur As Integer
    Dessin() As Integer
    Couleur As Long
End Type

Dim Pièce As Pièce1
Dim PièceNum As Integer
Dim TableauPièces(0 To 14) As Pièce1
Dim Ratio As Single
Dim Tableau() As Integer
Dim Fini As Boolean
Dim Score As Integer

Private Sub cmdCommencer_Click()
    NouvellePièce
    InitTableauPartie
    DoAffichage
    Timer1.Interval = 1000
    Timer1.Enabled = True
    pct.SetFocus
End Sub

Private Sub Form_Load()
    pct.ScaleMode = 3
    Ratio = pct.ScaleWidth / 10
    InitTableauPièces
    ReDim Tableau(0 To 9, 0 To 14) As Integer
End Sub

Private Sub InitTableauPièces()
    ' 0 : La barre verticale
    TableauPièces(0).Largeur = 0
    TableauPièces(0).Hauteur = 3
    ReDim TableauPièces(0).Dessin(0 To 0, 0 To 3) As Integer
    TableauPièces(0).Dessin(0, 0) = 1
    TableauPièces(0).Dessin(0, 1) = 1
    TableauPièces(0).Dessin(0, 2) = 1
    TableauPièces(0).Dessin(0, 3) = 1

    ' 1 : La barre horizontale
    TableauPièces(1).Largeur = 3
    TableauPièces(1).Hauteur = 0
    ReDim TableauPièces(1).Dessin(0 To 3, 0 To 0) As Integer
    TableauPièces(1).Dessin(0, 0) = 1
    TableauPièces(1).Dessin(1, 0) = 1
    TableauPièces(1).Dessin(2, 0) = 1
    TableauPièces(1).Dessin(3, 0) = 1
    
    ' 2 : Le L
    TableauPièces(2).Largeur = 1
    TableauPièces(2).Hauteur = 2
    ReDim TableauPièces(2).Dessin(0 To 1, 0 To 2) As Integer
    TableauPièces(2).Dessin(0, 0) = 1
    TableauPièces(2).Dessin(1, 0) = 0
    TableauPièces(2).Dessin(0, 1) = 1
    TableauPièces(2).Dessin(1, 1) = 0
    TableauPièces(2).Dessin(0, 2) = 1
    TableauPièces(2).Dessin(1, 2) = 1

    ' 3 : Le L retourné une première fois
    TableauPièces(3).Largeur = 2
    TableauPièces(3).Hauteur = 1
    ReDim TableauPièces(3).Dessin(0 To 2, 0 To 1) As Integer
    TableauPièces(3).Dessin(0, 0) = 0
    TableauPièces(3).Dessin(0, 1) = 1
    TableauPièces(3).Dessin(1, 0) = 0
    TableauPièces(3).Dessin(1, 1) = 1
    TableauPièces(3).Dessin(2, 0) = 1
    TableauPièces(3).Dessin(2, 1) = 1
    
    ' 4 : Le L retourné une troisième fois
    TableauPièces(4).Largeur = 1
    TableauPièces(4).Hauteur = 2
    ReDim TableauPièces(4).Dessin(0 To 1, 0 To 2) As Integer
    TableauPièces(4).Dessin(0, 0) = 1
    TableauPièces(4).Dessin(1, 0) = 1
    TableauPièces(4).Dessin(0, 1) = 0
    TableauPièces(4).Dessin(1, 1) = 1
    TableauPièces(4).Dessin(0, 2) = 0
    TableauPièces(4).Dessin(1, 2) = 1
    
    ' 5 : Le L retourné une quatrième fois
    TableauPièces(5).Largeur = 2
    TableauPièces(5).Hauteur = 1
    ReDim TableauPièces(5).Dessin(0 To 2, 0 To 1) As Integer
    TableauPièces(5).Dessin(0, 0) = 1
    TableauPièces(5).Dessin(0, 1) = 1
    TableauPièces(5).Dessin(1, 0) = 1
    TableauPièces(5).Dessin(1, 1) = 0
    TableauPièces(5).Dessin(2, 0) = 1
    TableauPièces(5).Dessin(2, 1) = 0
    
    ' 6 : Le cube
    TableauPièces(6).Largeur = 1
    TableauPièces(6).Hauteur = 1
    ReDim TableauPièces(6).Dessin(0 To 1, 0 To 1) As Integer
    TableauPièces(6).Dessin(0, 0) = 1
    TableauPièces(6).Dessin(0, 1) = 1
    TableauPièces(6).Dessin(1, 0) = 1
    TableauPièces(6).Dessin(1, 1) = 1
    
    ' 7 : L'espèce de croix...
    TableauPièces(7).Largeur = 1
    TableauPièces(7).Hauteur = 2
    ReDim TableauPièces(7).Dessin(0 To 1, 0 To 2) As Integer
    TableauPièces(7).Dessin(0, 0) = 1
    TableauPièces(7).Dessin(1, 0) = 0
    TableauPièces(7).Dessin(0, 1) = 1
    TableauPièces(7).Dessin(1, 1) = 1
    TableauPièces(7).Dessin(0, 2) = 0
    TableauPièces(7).Dessin(1, 2) = 1

    ' 8 : La même retournée
    TableauPièces(8).Largeur = 2
    TableauPièces(8).Hauteur = 1
    ReDim TableauPièces(8).Dessin(0 To 2, 0 To 1) As Integer
    TableauPièces(8).Dessin(0, 0) = 0
    TableauPièces(8).Dessin(0, 1) = 1
    TableauPièces(8).Dessin(1, 0) = 1
    TableauPièces(8).Dessin(1, 1) = 1
    TableauPièces(8).Dessin(2, 0) = 1
    TableauPièces(8).Dessin(2, 1) = 0
    
    ' 9 : L'autre croix
    TableauPièces(9).Largeur = 1
    TableauPièces(9).Hauteur = 2
    ReDim TableauPièces(9).Dessin(0 To 1, 0 To 2) As Integer
    TableauPièces(9).Dessin(0, 0) = 0
    TableauPièces(9).Dessin(1, 0) = 1
    TableauPièces(9).Dessin(0, 1) = 1
    TableauPièces(9).Dessin(1, 1) = 1
    TableauPièces(9).Dessin(0, 2) = 1
    TableauPièces(9).Dessin(1, 2) = 0

    ' 10 : La même retournée
    TableauPièces(10).Largeur = 2
    TableauPièces(10).Hauteur = 1
    ReDim TableauPièces(10).Dessin(0 To 2, 0 To 1) As Integer
    TableauPièces(10).Dessin(0, 0) = 1
    TableauPièces(10).Dessin(0, 1) = 0
    TableauPièces(10).Dessin(1, 0) = 1
    TableauPièces(10).Dessin(1, 1) = 1
    TableauPièces(10).Dessin(2, 0) = 0
    TableauPièces(10).Dessin(2, 1) = 1

    ' 11 : Le t
    TableauPièces(11).Largeur = 1
    TableauPièces(11).Hauteur = 2
    ReDim TableauPièces(11).Dessin(0 To 1, 0 To 2) As Integer
    TableauPièces(11).Dessin(0, 0) = 1
    TableauPièces(11).Dessin(1, 0) = 0
    TableauPièces(11).Dessin(0, 1) = 1
    TableauPièces(11).Dessin(1, 1) = 1
    TableauPièces(11).Dessin(0, 2) = 1
    TableauPièces(11).Dessin(1, 2) = 0

    ' 12 : t retourné 1
    TableauPièces(12).Largeur = 2
    TableauPièces(12).Hauteur = 1
    ReDim TableauPièces(12).Dessin(0 To 2, 0 To 1) As Integer
    TableauPièces(12).Dessin(0, 0) = 0
    TableauPièces(12).Dessin(0, 1) = 1
    TableauPièces(12).Dessin(1, 0) = 1
    TableauPièces(12).Dessin(1, 1) = 1
    TableauPièces(12).Dessin(2, 0) = 0
    TableauPièces(12).Dessin(2, 1) = 1
    
    ' 13 : t retourné 2
    TableauPièces(13).Largeur = 1
    TableauPièces(13).Hauteur = 2
    ReDim TableauPièces(13).Dessin(0 To 1, 0 To 2) As Integer
    TableauPièces(13).Dessin(0, 0) = 0
    TableauPièces(13).Dessin(1, 0) = 1
    TableauPièces(13).Dessin(0, 1) = 1
    TableauPièces(13).Dessin(1, 1) = 1
    TableauPièces(13).Dessin(0, 2) = 0
    TableauPièces(13).Dessin(1, 2) = 1

    ' 14 : t retourné 3
    TableauPièces(14).Largeur = 2
    TableauPièces(14).Hauteur = 1
    ReDim TableauPièces(14).Dessin(0 To 2, 0 To 1) As Integer
    TableauPièces(14).Dessin(0, 0) = 1
    TableauPièces(14).Dessin(0, 1) = 0
    TableauPièces(14).Dessin(1, 0) = 1
    TableauPièces(14).Dessin(1, 1) = 1
    TableauPièces(14).Dessin(2, 0) = 1
    TableauPièces(14).Dessin(2, 1) = 0
End Sub

Private Sub pct_KeyDown(KeyCode As Integer, Shift As Integer)
    If Fini = False Then
        TestCollision KeyCode
    End If
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    Dim Collision As Boolean
    
    Collision = TestCollision(40)
    If Collision = True Then
        CollerPièceTableau
        CheckLignes
        Fini = CheckFini
        If Fini Then
            MsgBox "perdu"
            Timer1.Enabled = False
        End If
        NouvellePièce
    End If
    DoAffichage
End Sub

Private Sub DoAffichage()
    Dim i As Integer
    Dim j As Integer
    Dim Coul As Long
    
    pct.Cls
    
    For i = 0 To 9
        For j = 0 To 14
            If Tableau(i, j) <> 0 Then
                Coul = SelectCoul(i, j)
                pct.Line (i * Ratio, j * Ratio)-((i + 1) * Ratio, (j + 1) * Ratio), QBColor(Tableau(i, j)), BF
            End If
        Next
        pct.Line (i * Ratio, 0)-(i * Ratio, (pct.ScaleHeight)), RGB(96, 96, 96)
    Next
    For j = 0 To 14
        pct.Line (0, j * Ratio)-(pct.ScaleWidth, j * Ratio), RGB(96, 96, 96)
    Next
    For i = 0 To Pièce.Largeur
        For j = 0 To Pièce.Hauteur
            If Pièce.Dessin(i, j) <> 0 Then
                pct.Line ((Pièce.x + i) * Ratio, (Pièce.y + j) * Ratio)-((Pièce.x + i + 1) * Ratio, (Pièce.y + j + 1) * Ratio), QBColor(Pièce.Couleur), BF
            End If
        Next
    Next
End Sub

Private Function SelectCoul(x As Integer, y As Integer) As Long
    Dim res As Long
    
    Select Case Tableau(x, y)
        Case 1
            res = vbRed
        Case 2
            res = vbYellow
        Case 3
            res = vbBlack
        Case 4
            res = vbGreen
    End Select
    SelectCoul = res
End Function

Private Sub CopierPièce(Optional NoNewXY As Boolean)
    Dim i As Integer
    Dim j As Integer
    
    If NoNewXY = False Then
        Pièce.x = TableauPièces(PièceNum).x
        Pièce.y = TableauPièces(PièceNum).y
    End If
    Pièce.Largeur = TableauPièces(PièceNum).Largeur
    Pièce.Hauteur = TableauPièces(PièceNum).Hauteur
    ReDim Pièce.Dessin(0 To Pièce.Largeur, 0 To Pièce.Hauteur)
    
    For i = 0 To Pièce.Largeur
        For j = 0 To Pièce.Hauteur
            Pièce.Dessin(i, j) = TableauPièces(PièceNum).Dessin(i, j)
        Next
    Next
End Sub

Private Sub CollerPièceTableau()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To Pièce.Largeur
        For j = 0 To Pièce.Hauteur
            If Pièce.Dessin(i, j) <> 0 Then
                Tableau(Pièce.x + i, Pièce.y + j) = Pièce.Couleur
            End If
        Next
    Next
    DoAffichage
End Sub

Private Function TestCollision(KeyCode As Integer) As Boolean
    Dim Cond As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim OldPièceNum As Integer
    
    ' Déplacement à droite
    If KeyCode = 39 Then
        If Pièce.x + Pièce.Largeur + 2 <= 10 Then
            Cond = False
            For i = 0 To Pièce.Largeur
                For j = 0 To Pièce.Hauteur
                    If Pièce.Dessin(i, j) <> 0 Then
                        If Tableau(Pièce.x + i + 1, Pièce.y + j) <> 0 Then
                            Cond = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If Cond = False Then
                Pièce.x = Pièce.x + 1
            End If
        End If
        DoAffichage
    ' Déplacement à gauche
    ElseIf KeyCode = 37 Then
        If Pièce.x - 1 >= 0 Then
            Cond = False
            For i = 0 To Pièce.Largeur
                For j = 0 To Pièce.Hauteur
                    If Pièce.Dessin(i, j) <> 0 Then
                        If Tableau(Pièce.x + i - 1, Pièce.y + j) <> 0 Then
                            Cond = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If Cond = False Then
                Pièce.x = Pièce.x - 1
            End If
        End If
        DoAffichage
    ' Déplacement en bas
    ElseIf KeyCode = 40 Then
        If Pièce.y + Pièce.Hauteur + 2 <= 15 Then
            Cond = False
            For i = 0 To Pièce.Largeur
                For j = 0 To Pièce.Hauteur
                    If Pièce.Dessin(i, j) <> 0 Then
                        If Tableau(Pièce.x + i, Pièce.y + j + 1) <> 0 Then
                            Cond = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If Cond = False Then
                Pièce.y = Pièce.y + 1
            End If
        Else
            Cond = True
        End If
        DoAffichage
    ElseIf KeyCode = 13 Then ' Entrée = Retournement
        OldPièceNum = PièceNum
        Retourner
        If Pièce.y + Pièce.Hauteur + 1 <= 15 And Pièce.x + Pièce.Largeur + 1 <= 10 Then
            Cond = False
            For i = 0 To Pièce.Largeur
                For j = 0 To Pièce.Hauteur
                    If Pièce.Dessin(i, j) <> 0 Then
                        If Tableau(Pièce.x + i, Pièce.y + j) <> 0 Then
                            Cond = True
                            Exit For
                        End If
                    End If
                Next
            Next
        Else
            Cond = True
        End If
        If Cond = True Then
            PièceNum = OldPièceNum
            CopierPièce True
        Else
            DoAffichage
        End If
    End If
    TestCollision = Cond
End Function

Private Sub InitTableauPartie()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 9
        For j = 0 To 14
            Tableau(i, j) = 0
        Next
    Next
    
    Fini = False
End Sub

Private Sub NouvellePièce()
    Dim Resul As Integer

    Randomize
    Resul = Int(6 * Rnd)
    Select Case Resul
        Case 0
            PièceNum = 0
        Case 1
            PièceNum = 2
        Case 2
            PièceNum = 6
        Case 3
            PièceNum = 7
        Case 4
            PièceNum = 9
        Case 5
            PièceNum = 11
    End Select
    CopierPièce
    
    Resul = Int(6 * Rnd + 1)
    Pièce.Couleur = Resul
End Sub

Private Sub CheckLignes()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Arret As Boolean
    Dim TabSupp(0 To 14) As Boolean
    
    ' On cherche les lignes à supprimer
    For j = 1 To 14
        Arret = False
        For i = 0 To 9
            If Tableau(i, j) = 0 Then
                Arret = True
                Exit For
            End If
        Next
        If Arret = False Then
            TabSupp(j) = True
        End If
    Next
    
    ' On les supprime
    For j = 1 To 14
        If TabSupp(j) = True Then
            Score = Score + 1
            For k = j - 1 To 1 Step -1
                For i = 0 To 9
                    Tableau(i, k + 1) = Tableau(i, k)
                    Tableau(i, k) = 0
                Next
            Next
        End If
    Next
    lblScore.Caption = "Score : " & Chr(13) & CStr(Score)
End Sub

Private Function CheckFini() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim Resul As Boolean
    
    Resul = False
    For i = 0 To 9
        If Tableau(i, 0) <> 0 Then
            Resul = True
            Exit For
        End If
    Next
    
    CheckFini = Resul
End Function

Private Sub Retourner()
    Dim Num As Integer

    Select Case PièceNum
        Case 0
            Num = 1
        Case 1
            Num = 0
        Case 2
            Num = 3
        Case 3
            Num = 4
        Case 4
            Num = 5
        Case 5
            Num = 2
        Case 6
            Num = 6
        Case 7
            Num = 8
        Case 8
            Num = 7
        Case 9
            Num = 10
        Case 10
            Num = 9
        Case 11
            Num = 12
        Case 12
            Num = 13
        Case 13
            Num = 14
        Case 14
            Num = 11
    End Select
    
    PièceNum = Num
    CopierPièce True
End Sub
