VERSION 5.00
Begin VB.Form frmServico 
   Caption         =   "servico"
   ClientHeight    =   3780
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   7995
   Icon            =   "frmServico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   7995
   Begin VB.TextBox txtComissao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Text            =   "0,00"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   7455
   End
   Begin VB.TextBox txtTipo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7995
      TabIndex        =   11
      Top             =   2925
      Width           =   7995
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   4080
         Picture         =   "frmServico.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   5880
         Picture         =   "frmServico.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdTipo 
      Height          =   285
      Left            =   4080
      Picture         =   "frmServico.frx":1015
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão:"
      Height          =   195
      Index           =   5
      Left            =   1200
      TabIndex        =   9
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "valor:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1305
      Width           =   465
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Serviço:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   945
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "id:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   165
   End
End
Attribute VB_Name = "frmServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdTipo_Click()
    buscaTipo
End Sub
Sub buscaTipo()
    Dim t As tipo_servico
    
    t = servico.BuscaTipoServico(txtTipo.Text)
    If t.id <> "" Then
        txtTipo.tag = t.nome
        txtTipo.tag = t.id
        txtNome.SetFocus
    End If
    
    
End Sub
Private Sub txtComissao_GotFocus()
    ganhaFoco txtComissao
End Sub

Private Sub txtComissao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtComissao_LostFocus()
    PerdeFoco txtComissao
End Sub

Private Sub txtNome_GotFocus()
    ganhaFoco txtNome
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtNome_LostFocus()
    AjustaTexto txtNome
End Sub

Private Sub txtValor_GotFocus()
    ganhaFoco txtValor
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtValor_LostFocus()
    PerdeFoco txtValor
End Sub
