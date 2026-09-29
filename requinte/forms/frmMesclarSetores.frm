VERSION 5.00
Begin VB.Form frmMesclarSetores 
   Caption         =   "Mesclar Setores"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Mesclar"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtEnt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txtSai 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Height          =   285
      Left            =   4080
      Picture         =   "frmMesclarSetores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   285
      Left            =   4080
      Picture         =   "frmMesclarSetores.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblCidEnt 
      AutoSize        =   -1  'True
      Caption         =   "cidade"
      Height          =   195
      Left            =   4560
      TabIndex        =   5
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Setor que entra"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblCidSai 
      AutoSize        =   -1  'True
      Caption         =   "cidade"
      Height          =   195
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Setor que sai"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmMesclarSetores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    buscaCidadeSai
End Sub
Sub buscaCidadeSai()
    Dim s As setor
    s = Setores.F_BuscaSetor(txtSai.Text)
    If s.cod <> "" Then
        txtSai.Text = s.nome
        txtSai.tag = s.cod
        lblCidSai.caption = s.Cidade
    End If
    
    
End Sub
Sub buscaCidadeEnt()
    Dim s As setor
    s = Setores.F_BuscaSetor(txtEnt.Text)
    If s.cod <> "" Then
        txtEnt.Text = s.nome
        txtEnt.tag = s.cod
        lblCidEnt.caption = s.Cidade
    End If
    
    
End Sub

Private Sub Command1_Click()
    Dim strSQL As String
    strSQL = "update tab_clientes "
    
    


End Sub

Private Sub Command2_Click()
    buscaCidadeEnt
End Sub
