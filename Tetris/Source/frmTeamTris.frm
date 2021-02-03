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

Private Type Pi�ce1
    x As Integer
    y As Integer
    Largeur As Integer
    Hauteur As Integer
    Dessin() As Integer
    Couleur As Long
End Type

Dim Pi�ce As Pi�ce1
Dim Pi�ceNum As Integer
Dim TableauPi�ces(0 To 14) As Pi�ce1
Dim Ratio As Single
Dim Tableau() As Integer
Dim Fini As Boolean
Dim Score As Integer

Private Sub cmdCommencer_Click()
    NouvellePi�ce
    InitTableauPartie
    DoAffichage
    Timer1.Interval = 1000
    Timer1.Enabled = True
    pct.SetFocus
End Sub

Private Sub Form_Load()
    pct.ScaleMode = 3
    Ratio = pct.ScaleWidth / 10
    InitTableauPi�ces
    ReDim Tableau(0 To 9, 0 To 14) As Integer
End Sub

Private Sub InitTableauPi�ces()
    ' 0 : La barre verticale
    TableauPi�ces(0).Largeur = 0
    TableauPi�ces(0).Hauteur = 3
    ReDim TableauPi�ces(0).Dessin(0 To 0, 0 To 3) As Integer
    TableauPi�ces(0).Dessin(0, 0) = 1
    TableauPi�ces(0).Dessin(0, 1) = 1
    TableauPi�ces(0).Dessin(0, 2) = 1
    TableauPi�ces(0).Dessin(0, 3) = 1

    ' 1 : La barre horizontale
    TableauPi�ces(1).Largeur = 3
    TableauPi�ces(1).Hauteur = 0
    ReDim TableauPi�ces(1).Dessin(0 To 3, 0 To 0) As Integer
    TableauPi�ces(1).Dessin(0, 0) = 1
    TableauPi�ces(1).Dessin(1, 0) = 1
    TableauPi�ces(1).Dessin(2, 0) = 1
    TableauPi�ces(1).Dessin(3, 0) = 1
    
    ' 2 : Le L
    TableauPi�ces(2).Largeur = 1
    TableauPi�ces(2).Hauteur = 2
    ReDim TableauPi�ces(2).Dessin(0 To 1, 0 To 2) As Integer
    TableauPi�ces(2).Dessin(0, 0) = 1
    TableauPi�ces(2).Dessin(1, 0) = 0
    TableauPi�ces(2).Dessin(0, 1) = 1
    TableauPi�ces(2).Dessin(1, 1) = 0
    TableauPi�ces(2).Dessin(0, 2) = 1
    TableauPi�ces(2).Dessin(1, 2) = 1

    ' 3 : Le L retourn� une premi�re fois
    TableauPi�ces(3).Largeur = 2
    TableauPi�ces(3).Hauteur = 1
    ReDim TableauPi�ces(3).Dessin(0 To 2, 0 To 1) As Integer
    TableauPi�ces(3).Dessin(0, 0) = 0
    TableauPi�ces(3).Dessin(0, 1) = 1
    TableauPi�ces(3).Dessin(1, 0) = 0
    TableauPi�ces(3).Dessin(1, 1) = 1
    TableauPi�ces(3).Dessin(2, 0) = 1
    TableauPi�ces(3).Dessin(2, 1) = 1
    
    ' 4 : Le L retourn� une troisi�me fois
    TableauPi�ces(4).Largeur = 1
    TableauPi�ces(4).Hauteur = 2
    ReDim TableauPi�ces(4).Dessin(0 To 1, 0 To 2) As Integer
    TableauPi�ces(4).Dessin(0, 0) = 1
    TableauPi�ces(4).Dessin(1, 0) = 1
    TableauPi�ces(4).Dessin(0, 1) = 0
    TableauPi�ces(4).Dessin(1, 1) = 1
    TableauPi�ces(4).Dessin(0, 2) = 0
    TableauPi�ces(4).Dessin(1, 2) = 1
    
    ' 5 : Le L retourn� une quatri�me fois
    TableauPi�ces(5).Largeur = 2
    TableauPi�ces(5).Hauteur = 1
    ReDim TableauPi�ces(5).Dessin(0 To 2, 0 To 1) As Integer
    TableauPi�ces(5).Dessin(0, 0) = 1
    TableauPi�ces(5).Dessin(0, 1) = 1
    TableauPi�ces(5).Dessin(1, 0) = 1
    TableauPi�ces(5).Dessin(1, 1) = 0
    TableauPi�ces(5).Dessin(2, 0) = 1
    TableauPi�ces(5).Dessin(2, 1) = 0
    
    ' 6 : Le cube
    TableauPi�ces(6).Largeur = 1
    TableauPi�ces(6).Hauteur = 1
    ReDim TableauPi�ces(6).Dessin(0 To 1, 0 To 1) As Integer
    TableauPi�ces(6).Dessin(0, 0) = 1
    TableauPi�ces(6).Dessin(0, 1) = 1
    TableauPi�ces(6).Dessin(1, 0) = 1
    TableauPi�ces(6).Dessin(1, 1) = 1
    
    ' 7 : L'esp�ce de croix...
    TableauPi�ces(7).Largeur = 1
    TableauPi�ces(7).Hauteur = 2
    ReDim TableauPi�ces(7).Dessin(0 To 1, 0 To 2) As Integer
    TableauPi�ces(7).Dessin(0, 0) = 1
    TableauPi�ces(7).Dessin(1, 0) = 0
    TableauPi�ces(7).Dessin(0, 1) = 1
    TableauPi�ces(7).Dessin(1, 1) = 1
    TableauPi�ces(7).Dessin(0, 2) = 0
    TableauPi�ces(7).Dessin(1, 2) = 1

    ' 8 : La m�me retourn�e
    TableauPi�ces(8).Largeur = 2
    TableauPi�ces(8).Hauteur = 1
    ReDim TableauPi�ces(8).Dessin(0 To 2, 0 To 1) As Integer
    TableauPi�ces(8).Dessin(0, 0) = 0
    TableauPi�ces(8).Dessin(0, 1) = 1
    TableauPi�ces(8).Dessin(1, 0) = 1
    TableauPi�ces(8).Dessin(1, 1) = 1
    TableauPi�ces(8).Dessin(2, 0) = 1
    TableauPi�ces(8).Dessin(2, 1) = 0
    
    ' 9 : L'autre croix
    TableauPi�ces(9).Largeur = 1
    TableauPi�ces(9).Hauteur = 2
    ReDim TableauPi�ces(9).Dessin(0 To 1, 0 To 2) As Integer
    TableauPi�ces(9).Dessin(0, 0) = 0
    TableauPi�ces(9).Dessin(1, 0) = 1
    TableauPi�ces(9).Dessin(0, 1) = 1
    TableauPi�ces(9).Dessin(1, 1) = 1
    TableauPi�ces(9).Dessin(0, 2) = 1
    TableauPi�ces(9).Dessin(1, 2) = 0

    ' 10 : La m�me retourn�e
    TableauPi�ces(10).Largeur = 2
    TableauPi�ces(10).Hauteur = 1
    ReDim TableauPi�ces(10).Dessin(0 To 2, 0 To 1) As Integer
    TableauPi�ces(10).Dessin(0, 0) = 1
    TableauPi�ces(10).Dessin(0, 1) = 0
    TableauPi�ces(10).Dessin(1, 0) = 1
    TableauPi�ces(10).Dessin(1, 1) = 1
    TableauPi�ces(10).Dessin(2, 0) = 0
    TableauPi�ces(10).Dessin(2, 1) = 1

    ' 11 : Le t
    TableauPi�ces(11).Largeur = 1
    TableauPi�ces(11).Hauteur = 2
    ReDim TableauPi�ces(11).Dessin(0 To 1, 0 To 2) As Integer
    TableauPi�ces(11).Dessin(0, 0) = 1
    TableauPi�ces(11).Dessin(1, 0) = 0
    TableauPi�ces(11).Dessin(0, 1) = 1
    TableauPi�ces(11).Dessin(1, 1) = 1
    TableauPi�ces(11).Dessin(0, 2) = 1
    TableauPi�ces(11).Dessin(1, 2) = 0

    ' 12 : t retourn� 1
    TableauPi�ces(12).Largeur = 2
    TableauPi�ces(12).Hauteur = 1
    ReDim TableauPi�ces(12).Dessin(0 To 2, 0 To 1) As Integer
    TableauPi�ces(12).Dessin(0, 0) = 0
    TableauPi�ces(12).Dessin(0, 1) = 1
    TableauPi�ces(12).Dessin(1, 0) = 1
    TableauPi�ces(12).Dessin(1, 1) = 1
    TableauPi�ces(12).Dessin(2, 0) = 0
    TableauPi�ces(12).Dessin(2, 1) = 1
    
    ' 13 : t retourn� 2
    TableauPi�ces(13).Largeur = 1
    TableauPi�ces(13).Hauteur = 2
    ReDim TableauPi�ces(13).Dessin(0 To 1, 0 To 2) As Integer
    TableauPi�ces(13).Dessin(0, 0) = 0
    TableauPi�ces(13).Dessin(1, 0) = 1
    TableauPi�ces(13).Dessin(0, 1) = 1
    TableauPi�ces(13).Dessin(1, 1) = 1
    TableauPi�ces(13).Dessin(0, 2) = 0
    TableauPi�ces(13).Dessin(1, 2) = 1

    ' 14 : t retourn� 3
    TableauPi�ces(14).Largeur = 2
    TableauPi�ces(14).Hauteur = 1
    ReDim TableauPi�ces(14).Dessin(0 To 2, 0 To 1) As Integer
    TableauPi�ces(14).Dessin(0, 0) = 1
    TableauPi�ces(14).Dessin(0, 1) = 0
    TableauPi�ces(14).Dessin(1, 0) = 1
    TableauPi�ces(14).Dessin(1, 1) = 1
    TableauPi�ces(14).Dessin(2, 0) = 1
    TableauPi�ces(14).Dessin(2, 1) = 0
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
        CollerPi�ceTableau
        CheckLignes
        Fini = CheckFini
        If Fini Then
            MsgBox "perdu"
            Timer1.Enabled = False
        End If
        NouvellePi�ce
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
    For i = 0 To Pi�ce.Largeur
        For j = 0 To Pi�ce.Hauteur
            If Pi�ce.Dessin(i, j) <> 0 Then
                pct.Line ((Pi�ce.x + i) * Ratio, (Pi�ce.y + j) * Ratio)-((Pi�ce.x + i + 1) * Ratio, (Pi�ce.y + j + 1) * Ratio), QBColor(Pi�ce.Couleur), BF
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

Private Sub CopierPi�ce(Optional NoNewXY As Boolean)
    Dim i As Integer
    Dim j As Integer
    
    If NoNewXY = False Then
        Pi�ce.x = TableauPi�ces(Pi�ceNum).x
        Pi�ce.y = TableauPi�ces(Pi�ceNum).y
    End If
    Pi�ce.Largeur = TableauPi�ces(Pi�ceNum).Largeur
    Pi�ce.Hauteur = TableauPi�ces(Pi�ceNum).Hauteur
    ReDim Pi�ce.Dessin(0 To Pi�ce.Largeur, 0 To Pi�ce.Hauteur)
    
    For i = 0 To Pi�ce.Largeur
        For j = 0 To Pi�ce.Hauteur
            Pi�ce.Dessin(i, j) = TableauPi�ces(Pi�ceNum).Dessin(i, j)
        Next
    Next
End Sub

Private Sub CollerPi�ceTableau()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To Pi�ce.Largeur
        For j = 0 To Pi�ce.Hauteur
            If Pi�ce.Dessin(i, j) <> 0 Then
                Tableau(Pi�ce.x + i, Pi�ce.y + j) = Pi�ce.Couleur
            End If
        Next
    Next
    DoAffichage
End Sub

Private Function TestCollision(KeyCode As Integer) As Boolean
    Dim Cond As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim OldPi�ceNum As Integer
    
    ' D�placement � droite
    If KeyCode = 39 Then
        If Pi�ce.x + Pi�ce.Largeur + 2 <= 10 Then
            Cond = False
            For i = 0 To Pi�ce.Largeur
                For j = 0 To Pi�ce.Hauteur
                    If Pi�ce.Dessin(i, j) <> 0 Then
                        If Tableau(Pi�ce.x + i + 1, Pi�ce.y + j) <> 0 Then
                            Cond = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If Cond = False Then
                Pi�ce.x = Pi�ce.x + 1
            End If
        End If
        DoAffichage
    ' D�placement � gauche
    ElseIf KeyCode = 37 Then
        If Pi�ce.x - 1 >= 0 Then
            Cond = False
            For i = 0 To Pi�ce.Largeur
                For j = 0 To Pi�ce.Hauteur
                    If Pi�ce.Dessin(i, j) <> 0 Then
                        If Tableau(Pi�ce.x + i - 1, Pi�ce.y + j) <> 0 Then
                            Cond = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If Cond = False Then
                Pi�ce.x = Pi�ce.x - 1
            End If
        End If
        DoAffichage
    ' D�placement en bas
    ElseIf KeyCode = 40 Then
        If Pi�ce.y + Pi�ce.Hauteur + 2 <= 15 Then
            Cond = False
            For i = 0 To Pi�ce.Largeur
                For j = 0 To Pi�ce.Hauteur
                    If Pi�ce.Dessin(i, j) <> 0 Then
                        If Tableau(Pi�ce.x + i, Pi�ce.y + j + 1) <> 0 Then
                            Cond = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If Cond = False Then
                Pi�ce.y = Pi�ce.y + 1
            End If
        Else
            Cond = True
        End If
        DoAffichage
    ElseIf KeyCode = 13 Then ' Entr�e = Retournement
        OldPi�ceNum = Pi�ceNum
        Retourner
        If Pi�ce.y + Pi�ce.Hauteur + 1 <= 15 And Pi�ce.x + Pi�ce.Largeur + 1 <= 10 Then
            Cond = False
            For i = 0 To Pi�ce.Largeur
                For j = 0 To Pi�ce.Hauteur
                    If Pi�ce.Dessin(i, j) <> 0 Then
                        If Tableau(Pi�ce.x + i, Pi�ce.y + j) <> 0 Then
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
            Pi�ceNum = OldPi�ceNum
            CopierPi�ce True
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

Private Sub NouvellePi�ce()
    Dim Resul As Integer

    Randomize
    Resul = Int(6 * Rnd)
    Select Case Resul
        Case 0
            Pi�ceNum = 0
        Case 1
            Pi�ceNum = 2
        Case 2
            Pi�ceNum = 6
        Case 3
            Pi�ceNum = 7
        Case 4
            Pi�ceNum = 9
        Case 5
            Pi�ceNum = 11
    End Select
    CopierPi�ce
    
    Resul = Int(6 * Rnd + 1)
    Pi�ce.Couleur = Resul
End Sub

Private Sub CheckLignes()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Arret As Boolean
    Dim TabSupp(0 To 14) As Boolean
    
    ' On cherche les lignes � supprimer
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

    Select Case Pi�ceNum
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
    
    Pi�ceNum = Num
    CopierPi�ce True
End Sub
