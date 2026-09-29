VERSION 5.00
Begin VB.Form frmIncUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "frmIncUsu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5505
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.ComboBox cmbxFuncionario 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4455
      End
      Begin VB.ComboBox cmbxTipo_Usu 
         Height          =   315
         ItemData        =   "frmIncUsu.frx":0CCA
         Left            =   240
         List            =   "frmIncUsu.frx":0CDA
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   4455
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
      Width           =   4935
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
      Caption         =   "&Sair"
      Height          =   735
      Left            =   3720
      Picture         =   "frmIncUsu.frx":0D0D
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   2280
      Picture         =   "frmIncUsu.frx":0E78
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4935
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmIncUsu.frx":1320
         Left            =   240
         List            =   "frmIncUsu.frx":132A
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
Dim bActive As Boolean
Private Sub cmbxFuncionario_Click()
    cmbxFuncionario.tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & cmbxFuncionario.Text & "'")
End Sub


Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub carrega()
    Dim strSQL As String
    Dim r As New ADODB.Recordset
    strSQL = "select "
    
    
End Sub
Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If strCadFun_Usu <> "S" Then
        MsgBox "Você não tem permissão para cadastrar funcionários", vbInformation
        Exit Sub
    End If
    
    If cmbxFuncionario.tag = "" Then
        MsgBox "Informe o nome do Usuário      ", vbExclamation
        cmbxFuncionario.SetFocus
        Exit Sub
    ElseIf cmbxTipo_Usu.Text = "" Then
        MsgBox "Informe o Tipo do Usuário      ", vbExclamation
        cmbxTipo_Usu.SetFocus
        Exit Sub
    Else
        
    End If
    If Int(strNivel_Usu) < Left(cmbxTipo_Usu.Text, 1) Then
        MsgBox "Você não pode adicionar ou alterar um usuário para nível superior a si.", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        strSQL = F_BuscaUmCampo("cod_FUN", "Tab_usuarios", "COD_FUN='" & cmbxFuncionario.tag & "'")
        If strSQL <> "" And Me.tag = "" Then
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
    
    strSQL = strSQL & " where cod_fun='" & Me.tag & "' "
    
    
    
    Conn.Execute strSQL
    MsgBox "Usuário Cadastrado!", vbInformation
    
    If IsLoad("frmEdiUsu") Then
        'RefreshDataPosition frmEdiUsu.adc
    End If
    Unload Me
End Sub
Sub P_Grava()
    Dim strSQL As String
    Dim sCod As String
    sCod = cmbxFuncionario.tag
    
    strSQL = "insert into tab_usuarios(cod_fun,nivel_usu,status_usu,senha_usu)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & Left(cmbxTipo_Usu.Text, 1) & "',"
    strSQL = strSQL & "'1',"
    strSQL = strSQL & "'" & F_GeraPW(txtbxSenha.Text) & "') "
    Conn.Execute strSQL
    frmDefUsu.tag = sCod
    frmDefUsu.caption = sCod & "-" & cmbxFuncionario.Text
    frmDefUsu.Show 1
    MsgBox "Usuário Cadastrado!", vbInformation
    cmbxFuncionario.ListIndex = "-1"
    txtbxSenha.Text = ""
    
    If IsLoad("frmEdiUsu") Then
        'RefreshDataPosition frmEdiUsu.adc
    End If
    cmbxFuncionario.SetFocus
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        
    End If
    bActive = True
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
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxSenha_LostFocus()
    P_AjustaTexto txtbxSenha
End Sub
