VERSION 5.00
Begin VB.Form frmCadContaCob 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Conta Cobrança"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   7080
      Picture         =   "frmCadContaCob.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   5640
      Picture         =   "frmCadContaCob.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtbxNome 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da Conta:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmCadContaCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSalvar_Click()
    If Me.tag <> "" Then
        grava
    Else
        altera
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
Sub grava()

End Sub
Sub altera()

End Sub

