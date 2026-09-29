VERSION 5.00
Begin VB.Form frmCadPermissoes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão de Permissões"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbxGrupo 
      Height          =   285
      Left            =   480
      MaxLength       =   40
      TabIndex        =   3
      Top             =   960
      Width           =   4815
   End
   Begin VB.TextBox txtbxObs 
      Height          =   285
      Left            =   480
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1560
      Width           =   5055
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   3840
      Picture         =   "frmCadPermissoes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   615
      Left            =   2040
      Picture         =   "frmCadPermissoes.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   480
      MaxLength       =   40
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton cmndbtnBuscar 
      Height          =   285
      Left            =   5280
      Picture         =   "frmCadPermissoes.frx":08D5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Grupo"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Observações:"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome da Permissão:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCadPermissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim g As grupoPer
    g = F_BuscaGrupoPer(txtbxGrupo.Text)
    If g.Codigo <> "" Then
        txtbxGrupo.Text = g.nome
        txtbxGrupo.tag = g.Codigo
    End If
    
    
End Sub



Private Sub cmndbtnSalvar_Click()
    If Me.tag = "" Then
        If permissao.F_GravaPermissao(txtbxNome.Text, txtbxObs.Text, txtbxGrupo.tag) <> "" Then
            MsgBox "Permissão gravada!", vbInformation
            txtbxNome.Text = ""
            txtbxNome.SetFocus
            Exit Sub
        End If
    Else
        
    End If
End Sub

Private Sub txtbxGrupo_GotFocus()
    P_SelecTexto txtbxGrupo
End Sub

Private Sub txtbxGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub

Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub

Private Sub txtbxObs_GotFocus()
    P_SelecTexto txtbxObs
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    P_AjustaTexto txtbxObs
End Sub
