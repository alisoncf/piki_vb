VERSION 5.00
Begin VB.Form frmSenhaCodBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2265
   ClientLeft      =   15
   ClientTop       =   -15
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtbxUsu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Acesso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   2535
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
        P_SelecTexto txtbxUsu
        txtbxUsu.SetFocus
        Exit Sub
    End If
    sSenha = F_BuscaUmCampo("senha_usu", "tab_usuarios", "cod_fun='" & _
        Left(txtbxUsu.Text, 3) & "'")
        
    If sSenha <> F_GeraPW(Mid(txtbxUsu.Text, 4, Len(txtbxUsu.Text))) Then
        MsgBox "Autenticação inválida!", vbExclamation
        P_SelecTexto txtbxUsu
        txtbxUsu.SetFocus
        Exit Sub
    End If
    
    Me.tag = sUsu
    Me.Hide
End Sub

Private Sub cmndbtnSair_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub txtbxUsu_GotFocus()
    P_SelecTexto txtbxUsu
End Sub

Private Sub txtbxUsu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AjustaTexto txtbxUsu
        cmndbtnOk_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    End If
End Sub

Private Sub txtbxUsu_LostFocus()
    AjustaTexto txtbxUsu
End Sub
