VERSION 5.00
Begin VB.Form frmLoto 
   Caption         =   "TeamLoto"
   ClientHeight    =   3600
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5148
   Icon            =   "frmLoto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   5148
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Loto"
      Height          =   1332
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4932
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   4
         Left            =   3960
         TabIndex        =   12
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   852
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Résultat du tirage :"
      Height          =   2052
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4932
      Begin VB.TextBox txtResul 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   4
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "TextResul"
         Top             =   1080
         Width           =   852
      End
      Begin VB.TextBox txtResul 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   3
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "TextResul"
         Top             =   1080
         Width           =   852
      End
      Begin VB.TextBox txtResul 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   2
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "TextResul"
         Top             =   1080
         Width           =   852
      End
      Begin VB.TextBox txtResul 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "TextResul"
         Top             =   1080
         Width           =   852
      End
      Begin VB.TextBox txtResul 
         Alignment       =   2  'Center
         Height          =   288
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "TextResul"
         Top             =   1080
         Width           =   852
      End
      Begin VB.CommandButton cmdTirage 
         Caption         =   "Effectuer le tirage"
         Height          =   372
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1812
      End
   End
End
Attribute VB_Name = "frmLoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Tirer les numéros au sort et remplir
' les cases résultat
Private Sub cmdTirage_Click()
    Dim i As Integer
    Dim TResul As Integer
    Dim CompteurOK As Integer
    Dim Utilisé As Boolean
    Dim NBOk As Integer
    Dim mesg As String
    
    CompteurOK = 0
    Randomize
    Do
        Utilisé = False
        TResul = Int(100 * Rnd)
        For i = 0 To CompteurOK - 1
            If (Int(txtResul(i).Text) = TResul) Then
                Utilisé = True
            End If
        Next
        If Utilisé = False Then
            txtResul(CompteurOK).Text = Str(TResul)
            CompteurOK = CompteurOK + 1
        End If
    Loop While CompteurOK <= 4
    
    CompteurOK = 0
    For i = 0 To 4
        If txtResul(i).Text = txtNum(i).Text Then
            CompteurOK = CompteurOK + 1
        End If
    Next
    
    Select Case CompteurOK
        Case 0
            mesg = "Essaye encore petit scarabée !"
        Case Else
            mesg = "Vous avez " & CStr(CompteurOK) & " bons numéro(s) !"
    End Select
    MsgBox mesg, vbOKOnly + vbExclamation, "Résultats"
End Sub

' Chargement de la fenêtre :
' c'est par là que tout commence
Private Sub Form_Load()
    Dim i As Integer
    ' Phase d'initialisation :
    ' on remplit tout avec des 0
    For i = 0 To 4
        txtResul(i).Text = "0"
        txtNum(i).Text = "0"
    Next
End Sub
