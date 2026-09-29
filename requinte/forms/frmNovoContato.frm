VERSION 5.00
Begin VB.Form frmNovoContato 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Novo Contato"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtTel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmNovoContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    If Me.tag = "" Then
        Unload Me
        
    Else
        Me.Hide
    End If
End Sub

Private Sub cmdSalvar_Click()
    If Me.tag = "" Then
        grava
    Else
        altera
    End If
End Sub

Sub grava()
    Dim strSQL As String
    gravaContato txtNome.Text, txtTel
    MsgBox "Dados gravados!", vbInformation

End Sub
Sub altera()
    Dim strSQL As String
    alteraContato Me.tag, txtNome.Text, txtTel.Text
    MsgBox "Dados alterados!", vbInformation
    Unload Me
    
End Sub
