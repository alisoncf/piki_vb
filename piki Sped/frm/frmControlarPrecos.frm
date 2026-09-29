VERSION 5.00
Begin VB.Form frmControlarPrecos 
   Caption         =   "Controlar Preços:"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14820
   Icon            =   "frmControlarPrecos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcNota 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   13305
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   285
         Left            =   4320
         Picture         =   "frmControlarPrecos.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtbxNum_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         MaxLength       =   12
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxSerie_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxEntrada_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxEmissao_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtChaveNfe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9000
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Participante:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Número:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4680
         TabIndex        =   12
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Série:"
         Height          =   195
         Left            =   5760
         TabIndex        =   11
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "E&ntrada:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         TabIndex        =   10
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Emissão: "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   9
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chave NFE:"
         Height          =   195
         Left            =   9000
         TabIndex        =   8
         Top             =   120
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmControlarPrecos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
