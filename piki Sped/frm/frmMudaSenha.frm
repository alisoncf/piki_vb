VERSION 5.00
Begin VB.Form frmMudaSenha 
   Caption         =   "Mudar Senha:"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtbxSenha_Par 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmndbtnMudaSenha 
      Caption         =   "Mu&dar Senha"
      Height          =   525
      Left            =   120
      Picture         =   "frmMudaSenha.frx":0000
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtbxConfSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N&ova Senha:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Con&firmar Senha:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmMudaSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnMudaSenha_Click()
     Dim sSenha As String
    Dim strSQL As String
    
        frmLogin.tag = "senha"
        On Error Resume Next
        frmLogin.txtbxNome_Usu.Text = strNome_Usu
        frmLogin.txtbxNome_Usu.Locked = True
        frmLogin.Show 1
        If frmLogin.tag = "" Then
            Unload frmLogin
            Exit Sub
        End If
        On Error GoTo trata_erro
        
        
    
        If txtbxSenha_Par.Text <> txtbxConfSenha.Text Then
            MsgBox "As senhas são diferentes!", vbExclamation
            txtbxSenha_Par.SetFocus
            Exit Sub
        Else
            sSenha = F_GeraPW(txtbxSenha_Par.Text)
            strSQL = "update tab_usuarios set senha_usu='" & sSenha & "' "
            strSQL = strSQL & " where cod_fun = '" & strCod_Usu & "' "
            conn.Execute strSQL
            MsgBox "Senha alterada com sucesso!", vbInformation
            txtbxSenha_Par.Enabled = False
            Label15.Enabled = False
            
            Label19.Enabled = False
            txtbxConfSenha.Enabled = False
            cmndbtnMudaSenha.caption = "Mudar Senha"
            
        End If
        Unload frmLogin
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub txtbxConfSenha_GotFocus()
    ganhaFoco txtbxConfSenha
End Sub

Private Sub txtbxConfSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxConfSenha_LostFocus()
    AjustaTexto txtbxConfSenha
    If txtbxConfSenha.Text <> txtbxSenha_Par.Text Then
        MsgBox "A senha não confere!", vbExclamation
    End If
End Sub

Private Sub txtbxSenha_Par_GotFocus()
    ganhaFoco txtbxSenha_Par
End Sub

Private Sub txtbxSenha_Par_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxSenha_Par_LostFocus()
    AjustaTexto txtbxSenha_Par
End Sub
