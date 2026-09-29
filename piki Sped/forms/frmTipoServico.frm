VERSION 5.00
Begin VB.Form frmTipoServico 
   Caption         =   "Tipo de Serviço"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6540
   Icon            =   "frmTipoServico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComissao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "0,00"
      Top             =   1680
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
      ScaleWidth      =   6540
      TabIndex        =   4
      Top             =   2145
      Width           =   6540
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   4320
         Picture         =   "frmTipoServico.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   615
         Left            =   2520
         Picture         =   "frmTipoServico.frx":1E65
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   5775
   End
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   165
   End
End
Attribute VB_Name = "frmTipoServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirmar_Click()
    If txtNome.Text = "" Then
        MsgBox "informe um nome!", vbExclamation
        Exit Sub
    End If
    If Me.tag = "" Then
        salva
    Else
        altera
    
    End If
End Sub
Sub salva()
    If F_BuscaUmCampo("nome", "tipo_servico", "nome='" & txtNome.Text & "'") <> "" Then
        MsgBox "Já existe um tipo com este nome!", vbExclamation
        Exit Sub
    End If
    If servico.GravaTipo_Servico(txtNome.Text, txtComissao.Text) = True Then
        MsgBox "Tipo gravado!", vbInformation
        txtNome.SetFocus
    Else
        MsgBox "Erro", vbExclamation
    End If
End Sub
Sub altera()
    If F_BuscaUmCampo("nome", "tipo_servico", "nome='" & txtNome.Text & "' and id<>" & Me.tag) <> "" Then
        MsgBox "Já existe um tipo com este nome!", vbExclamation
        Exit Sub
    End If
    If servico.AlteraTipo_Servico(Me.tag, txtNome.Text, txtComissao.Text) = True Then
        MsgBox "Tipo alterado!", vbInformation
        txtNome.SetFocus
    Else
        MsgBox "Erro", vbExclamation
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
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
