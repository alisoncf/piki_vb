VERSION 5.00
Begin VB.Form frmRelFluxoGeral 
   Caption         =   "Análise de Fluxo Geral"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   0
      Width           =   9630
      Begin VB.PictureBox pcReferencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   1
         Top             =   0
         Width           =   2535
         Begin VB.TextBox txtDe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtAte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Até:"
            Height          =   195
            Left            =   1320
            TabIndex        =   4
            Top             =   0
            Width           =   285
         End
      End
   End
End
Attribute VB_Name = "frmRelFluxoGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
