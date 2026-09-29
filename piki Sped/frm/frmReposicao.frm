VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReposicao 
   Caption         =   "Reposição e Mercadorias"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PcProd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10080
      TabIndex        =   3
      Top             =   5010
      Width           =   10110
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   7
         Top             =   360
         Width           =   750
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton cmdBuscaProd 
         Height          =   285
         Left            =   4920
         Picture         =   "frmReposicao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
         Height          =   195
         Left            =   6480
         TabIndex        =   9
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd:"
         Height          =   195
         Left            =   5400
         TabIndex        =   8
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   600
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1695
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2990
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   10110
      TabIndex        =   1
      Top             =   5745
      Width           =   10110
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   10110
      TabIndex        =   0
      Top             =   0
      Width           =   10110
   End
End
Attribute VB_Name = "frmReposicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

