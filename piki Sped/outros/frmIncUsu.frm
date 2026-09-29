VERSION 5.00
Begin VB.Form frmIncUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "frmIncUsu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3525
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Mudar Senha"
      Height          =   735
      Left            =   120
      Picture         =   "frmIncUsu.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ComboBox cmbxFuncionario 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbxTipo_Usu 
         Height          =   315
         ItemData        =   "frmIncUsu.frx":0454
         Left            =   240
         List            =   "frmIncUsu.frx":0464
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
      Begin VB.TextBox txtbxSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Senha:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   510
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   2280
      Picture         =   "frmIncUsu.frx":0497
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   1320
      Picture         =   "frmIncUsu.frx":0821
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   3135
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmIncUsu.frx":0DAB
         Left            =   240
         List            =   "frmIncUsu.frx":0DB5
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmIncUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxFuncionario_Click()
    cmbxFuncionario.Tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & cmbxFuncionario.Text & "'")
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If cmbxFuncionario.Tag = "" Then
        MsgBox "Informe o nome do Usuário      ", vbExclamation
        cmbxFuncionario.SetFocus
        Exit Sub
    ElseIf cmbxTipo_Usu.Text = "" Then
        MsgBox "Informe o Tipo do Usuário      ", vbExclamation
        cmbxTipo_Usu.SetFocus
        Exit Sub
    Else
        
    End If
    If Me.Tag = "" Then
        strSQL = F_BuscaUmCampo("cod_FUN", "Tab_usuarios", "COD_FUN='" & cmbxFuncionario.Tag & "'")
        If strSQL <> "" And Me.Tag = "" Then
            MsgBox "Este usuário já existe!", vbExclamation
            Exit Sub
        End If
        P_Grava
    Else
        P_Altera
    End If
End Sub
Sub P_Altera()
    Dim strSQL As String
    
    
    strSQL = " update tab_usuarios set "
    
    strSQL = strSQL & "nivel_usu='" & Left(cmbxTipo_Usu.Text, 1) & "',"
    strSQL = strSQL & "status_usu='" & Left(cmbxStatus.Text, 1) & "' "
    
    strSQL = strSQL & " where cod_fun='" & Me.Tag & "' "
    Banco.Execute strSQL
    MsgBox "Usuário Cadastrado!", vbInformation
    
    If IsLoad("frmEdiUsu") Then
        RefreshDataPosition frmEdiUsu.dtProd
    End If
    Unload Me
End Sub
Sub P_Grava()
    Dim strSQL As String
    Dim sCod As String
    sCod = cmbxFuncionario.Tag
    
    strSQL = "insert into tab_usuarios(cod_fun,nivel_usu,status_usu,senha_usu)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & Left(cmbxTipo_Usu.Text, 1) & "',"
    strSQL = strSQL & "'1',"
    strSQL = strSQL & "'" & F_GeraPW(txtbxSenha.Text) & "') "
    Banco.Execute strSQL
    frmDefUsu.Tag = sCod
    frmDefUsu.caption = sCod & "-" & cmbxFuncionario.Text
    frmDefUsu.Show 1
    MsgBox "Usuário Cadastrado!", vbInformation
    cmbxFuncionario.ListIndex = "-1"
    txtbxSenha.Text = ""
    
    If IsLoad("frmEdiUsu") Then
        RefreshDataPosition frmEdiUsu.dtProd
    End If
    cmbxFuncionario.SetFocus
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    'Me.Height = 4290
    'Me.Width = 3345
    P_PreencheFunc cmbxFuncionario, 1
    P_CentraForm Me
End Sub

Private Sub txtbxNome_Usu_GotFocus()
    'P_SelecTexto txtbxNome_Usu
End Sub

Private Sub txtbxNome_Usu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Usu_LostFocus()
    'P_AjustaTexto txtbxNome_Usu
End Sub


Private Sub txtbxSenha_GotFocus()
    P_SelecTexto txtbxSenha
End Sub

Private Sub txtbxSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSenha_LostFocus()
    P_AjustaTexto txtbxSenha
End Sub
