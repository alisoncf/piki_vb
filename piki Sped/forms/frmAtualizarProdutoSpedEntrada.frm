VERSION 5.00
Begin VB.Form frmAtualizarProdutoSpedEntrada 
   Caption         =   "Informação de Produtos"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7965
   Icon            =   "frmAtualizarProdutoSpedEntrada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   7665
      TabIndex        =   10
      Top             =   2520
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "Atualizar"
         Height          =   495
         Left            =   6240
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtEan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   390
      End
      Begin VB.Label lblEan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   7665
      TabIndex        =   5
      Top             =   1440
      Width           =   7695
      Begin VB.TextBox txtNcm 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Atualizar"
         Height          =   495
         Left            =   6240
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblNcm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "Atualizar"
         Height          =   495
         Left            =   6240
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Produto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1485
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmAtualizarProdutoSpedEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
