VERSION 5.00
Begin VB.Form frmInFun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Funcionários"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "frmInFun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4260
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtbxCpf 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtbxGuerra 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbxCargo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox txtbxTelefone_Fun 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtbxRG_Fun 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtbxEndereco_Fun 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox txtbxNome_Fun 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CPF:"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N&ome:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F&unção:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Telefone:"
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RG:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereco:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome Completo:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "S&tatus:"
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   3600
      Width           =   2055
      Begin VB.ComboBox cmbxStatus 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmInFun.frx":030A
         Left            =   120
         List            =   "frmInFun.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2160
      TabIndex        =   17
      Top             =   3600
      Width           =   2055
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Salvar"
         Height          =   735
         Left            =   120
         Picture         =   "frmInFun.frx":033A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   1080
         Picture         =   "frmInFun.frx":08C4
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   180
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmInFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbxCargo_Click()
    cmbxCargo.tag = F_BuscaUmCampo("cod_car", "tab_cargos", "nome_car='" & cmbxCargo.Text & "'")
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Me.tag = "" Then
        If txtbxNome_Fun.Text = "" Then
            MsgBox "Informe o nome completo do funcionário"
            Exit Sub
        End If
        If txtbxGuerra.Text = "" Then
            MsgBox "Informe o nome do funcionário"
            Exit Sub
        End If
        If F_BuscaUmCampo("nome_fun", TAB_FUNCIONARIOS, "guerra_fun='" & _
            txtbxGuerra.Text & "'") <> "" Then
            MsgBox "Já existe um funcionário com este nome!", vbExclamation
            txtbxGuerra.SetFocus
            Exit Sub
            
        End If
        If cmbxCargo.tag = "" Then
            MsgBox "Informe o Cargo do funcionário"
            cmbxCargo.SetFocus
            Exit Sub
        End If
        P_Salva
    Else
        If txtbxNome_Fun.Text = "" Then
            MsgBox "Informe o nome completo do funcionário"
            Exit Sub
        End If
        If txtbxGuerra.Text = "" Then
            MsgBox "Informe o nome do funcionário"
            Exit Sub
        End If
        If cmbxCargo.tag = "" Then
            MsgBox "Informe o Cargo do funcionário"
            cmbxCargo.SetFocus
            Exit Sub
        End If
        If F_BuscaUmCampo("nome_fun", TAB_FUNCIONARIOS, "guerra_fun='" & _
            txtbxGuerra.Text & "' and cod_fun <>'" & Me.tag & "'") <> "" Then
            MsgBox "Já existe um funcionário com este nome!", vbExclamation
            txtbxGuerra.SetFocus
            Exit Sub
            
        End If
        P_Altera
    End If
End Sub
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    sCod = F_BuscaUmCampo("MAX(Cod_Fun)", TAB_FUNCIONARIOS, "")
    If sCod = "" Then
        sCod = "001"
    Else
        sCod = Format(Int(sCod) + 1, "00#")
    End If
    strSQL = "insert into tab_funcionarios(cod_fun,nome_fun,guerra_fun,rg_fun,cpf_fun,"
    strSQL = strSQL & "endereco_fun,telefone_Fun,cod_car,status_fun)Values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxNome_Fun.Text & "',"
    strSQL = strSQL & "'" & txtbxGuerra.Text & "',"
    strSQL = strSQL & "'" & txtbxRG_Fun.Text & "',"
    strSQL = strSQL & "'" & txtbxCpf.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco_Fun.Text & "', "
    strSQL = strSQL & "'" & txtbxTelefone_Fun.Text & "',"
    strSQL = strSQL & "'" & cmbxCargo.tag & "',"
    strSQL = strSQL & "'1')"
    conn.Execute strSQL
    MsgBox "Funcionário Cadastrado!", vbInformation
    txtbxEndereco_Fun.Text = ""
    cmbxCargo.ListIndex = -1
    txtbxNome_Fun.Text = ""
    txtbxTelefone_Fun.Text = ""
    txtbxRG_Fun.Text = ""
    If IsLoad("frmEdiFun") Then
        
    End If
    txtbxNome_Fun.SetFocus
    
End Sub
Sub P_Altera()
    strSQL = "update tab_funcionarios set "
    
    strSQL = strSQL & "nome_fun='" & txtbxNome_Fun.Text & "',"
    strSQL = strSQL & "guerra_fun='" & txtbxGuerra.Text & "',"
    strSQL = strSQL & "rg_fun='" & txtbxRG_Fun.Text & "',"
    strSQL = strSQL & "cpf_fun='" & txtbxRG_Fun.Text & "',"
    strSQL = strSQL & "endereco_fun='" & txtbxEndereco_Fun.Text & "', "
    strSQL = strSQL & "telefone_fun='" & txtbxTelefone_Fun.Text & "',"
    strSQL = strSQL & "cod_car='" & cmbxCargo.tag & "',"
    strSQL = strSQL & "status_fun='" & Left(cmbxStatus.Text, 1) & "' "
    strSQL = strSQL & " where cod_fun='" & Me.tag & "' "
    conn.Execute strSQL
    MsgBox "Funcionário Cadastrado!", vbInformation
    
    
    If IsLoad("frmEdiFun") Then
        
    End If
    If IsLoad("frmEdiUsu") Then
        
    End If
    Unload Me
End Sub
Private Sub Form_Activate()
    If Me.tag = "" Then
        cmbxStatus.ListIndex = 0
        cmbxStatus.Enabled = False
    End If
    'MudaAparencia Me
End Sub

Private Sub Form_Load()
    'Me.Width = 4740
    'Me.Height = 4500
    P_PreencheCargos cmbxCargo
    P_CentraForm Me
End Sub

Private Sub txtbxCpf_GotFocus()
    ganhaFoco txtbxCpf
End Sub

Private Sub txtbxCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCPF_LostFocus()
    txtbxCpf.Text = F_SoNumero(txtbxCpf.Text)
    If Not F_CPF(txtbxCpf.Text) Then
        MsgBox "CPF Inválido!", vbExclamation
        txtbxCpf.Text = ""
    Else
        P_FormataCPF txtbxCpf
    End If
    
End Sub

Private Sub txtbxEndereco_Fun_GotFocus()
    ganhaFoco txtbxEndereco_Fun
End Sub

Private Sub txtbxEndereco_Fun_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If

End Sub

Private Sub txtbxEndereco_Fun_LostFocus()
    AjustaTexto txtbxEndereco_Fun
End Sub


Private Sub txtbxGuerra_GotFocus()
    ganhaFoco txtbxGuerra
End Sub

Private Sub txtbxGuerra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxGuerra_LostFocus()
    AjustaTexto txtbxGuerra
End Sub

Private Sub txtbxNome_Fun_GotFocus()
    ganhaFoco txtbxNome_Fun
End Sub

Private Sub txtbxNome_Fun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNome_Fun_LostFocus()
    AjustaTexto txtbxNome_Fun
End Sub

Private Sub txtbxRG_Fun_GotFocus()
    ganhaFoco txtbxRG_Fun
End Sub

Private Sub txtbxRG_Fun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If

End Sub

Private Sub txtbxRG_Fun_LostFocus()
    AjustaTexto txtbxRG_Fun
End Sub

Private Sub txtbxTelefone_Fun_GotFocus()
    ganhaFoco txtbxTelefone_Fun
End Sub

Private Sub txtbxTelefone_Fun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If

End Sub

Private Sub txtbxTelefone_Fun_LostFocus()
    txtbxTelefone_Fun.Text = F_AjustaTelefone(txtbxTelefone_Fun.Text)
End Sub
