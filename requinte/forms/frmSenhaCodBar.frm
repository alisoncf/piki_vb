VERSION 5.00
Begin VB.Form frmSenhaCodBar 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "      **********AUTENTICAÇÃO**********"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtbxUsu 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frmSenhaCodBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnOk_Click()
    Dim sUsu As String
    Dim sSenha As String
    Dim strSQL As String
    sUsu = F_BuscaUmCampo("cod_fun", "tab_usuarios", "cod_fun='" & _
        Left(txtbxUsu.Text, 3) & "'")
    If sUsu = "" Then
        MsgBox "Autenticação inválida!", vbExclamation
        txtbxUsu.SetFocus
        Exit Sub
    End If
    sSenha = F_BuscaUmCampo("senha_usu", "tab_usuarios", "cod_fun='" & _
        Left(txtbxUsu.Text, 3) & "'")
        
    If sSenha <> F_GeraPW(Mid(txtbxUsu.Text, 4, Len(txtbxUsu.Text))) Then
        MsgBox "Autenticação inválida!", vbExclamation
        txtbxUsu.SetFocus
        Exit Sub
    End If
    
    Me.Tag = sUsu
    Me.Hide
End Sub

Private Sub cmndbtnSair_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub txtbxUsu_GotFocus()
    P_SelecTexto txtbxUsu
End Sub

Private Sub txtbxUsu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_AjustaTexto txtbxUsu
        cmndbtnOk_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    End If
End Sub

Private Sub txtbxUsu_LostFocus()
    P_AjustaTexto txtbxUsu
End Sub
