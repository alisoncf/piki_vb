VERSION 5.00
Begin VB.Form frmSenhaRepr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Senha de Representante"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmSenhaRepr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnCancelar 
      Cancel          =   -1  'True
      Caption         =   "Esc - Cancelar"
      Height          =   975
      Left            =   4920
      Picture         =   "frmSenhaRepr.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Confirmar"
      Height          =   975
      Left            =   3000
      Picture         =   "frmSenhaRepr.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtbxSenha 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   6495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha do Representante:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "É necessária confirmação para validação da operação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   6840
   End
End
Attribute VB_Name = "frmSenhaRepr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnCancelar_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("senha_rep", "tab_representantes", "cod_rep='" & Me.tag & "'")
    If UCase(strSQL) = UCase(txtbxSenha.Text) Then
        Me.tag = "x"
        Me.Hide
    Else
        MsgBox "Senha incorreta!", vbExclamation
        
    End If
End Sub

Private Sub txtbxSenha_GotFocus()
    P_SelecTexto txtbxSenha
End Sub

Private Sub txtbxSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnOk_Click
    End If
End Sub

Private Sub txtbxSenha_LostFocus()
    P_AjustaTexto txtbxSenha
End Sub
