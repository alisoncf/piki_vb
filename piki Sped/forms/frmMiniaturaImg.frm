VERSION 5.00
Begin VB.Form frmMiniaturaImg 
   Caption         =   "Imagens do Produto"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcQuadro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   6345
      TabIndex        =   4
      Top             =   360
      Width           =   6375
      Begin VB.PictureBox pcImg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1785
         ScaleWidth      =   3825
         TabIndex        =   5
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.PictureBox pcNavegaImgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdDir 
         Caption         =   ">"
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEsq 
         Caption         =   "<"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblCapImg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Imagem 0 de 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMiniaturaImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
