VERSION 5.00
Begin VB.Form frmCadCliRapido 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pré-Cadastro de Clientes:"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optnPessoa 
      Caption         =   "Pessoa &Jurídica"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optnPessoa 
      Caption         =   "Pessoa &Física"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtbxRazao 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2160
      Width           =   5775
   End
   Begin VB.TextBox txtbxCNPJ 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      MaxLength       =   18
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6930
      TabIndex        =   8
      Top             =   0
      Width           =   6930
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   480
         Left            =   0
         Picture         =   "frmCadCliRapido.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Novo Cliente"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   600
         TabIndex        =   9
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fecha&r"
      Height          =   735
      Left            =   4920
      Picture         =   "frmCadCliRapido.frx":08DF
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "Sal&var"
      Height          =   735
      Left            =   3360
      Picture         =   "frmCadCliRapido.frx":0A4A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome:"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label lblCpf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&CPF:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   345
   End
End
Attribute VB_Name = "frmCadCliRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    sCod = F_BuscaUmCampo(" max(cod_Cli)", "tab_Clientes", "")
    If sCod = "" Then       'pega só 4 dígitos para compor o dígito verificador
        sCod = "0001"       'que é o último número do cod_cli
    Else
        sCod = Left(sCod, 4)
        sCod = Format(Int(sCod) + 1, "000#")
    End If
    sCod = F_GeraDig(sCod)
    
    
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,cpf_cli,tipo_cli) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJ.Text & "',"
    
    If optnPessoa(0).Value = True Then
        strSQL = strSQL & "'F')"
    Else
        strSQL = strSQL & "'J')"
    End If
    
    banco.Execute strSQL
    Me.Tag = sCod
End Sub


Private Sub cmndbtnSair_Click()
    Me.Hide
End Sub

Private Sub cmndbtnSalvar_Click()
    If txtbxRazao.Text = "" Then
        MsgBox "Informe um nome para o Cliente!", vbExclamation
        Exit Sub
        
    ElseIf txtbxCNPJ.Text = "" Then
        If optnPessoa(0).Value = True Then
            MsgBox "Informe o CPF do Cliente!", vbExclamation
            Exit Sub
        Else
            MsgBox "Informe o CNPJ do Cliente!", vbExclamation
            Exit Sub
        End If
    ElseIf bCNPJ = False Then
        If optnPessoa(0).Value = True Then
            MsgBox "Informe um CPF Válido para Cliente!", vbExclamation
            Exit Sub
        Else
            MsgBox "Informe um CNPJ Válido para Cliente!", vbExclamation
            Exit Sub
        End If
    
    End If
    
    If Me.Tag = "" Then
        If F_BuscaUmCampo("cod_cli", "tab_clientes", "cpf_cli='" & txtbxCNPJ.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este CPF/CNPJ!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_cli", "tab_clientes", "nome_cli='" & txtbxRazao.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este Nome/Razão Social!", vbCritical
            Exit Sub
        
        End If
    
        P_Salva
        MsgBox "Cliente Cadastrado com sucesso!", vbInformation
        
        Me.Hide

        
    Else
        
    End If
    
End Sub



Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub



Private Sub optnPessoa_Click(Index As Integer)
    If Index = 0 Then
        lblCpf.caption = "&CPF:"
        'lblRg.caption = "&RG:"
        txtbxCNPJ.MaxLength = 14
    Else
        lblCpf.caption = "&CNPJ:"
        'lblRg.caption = "&Insc Est:"
        txtbxCNPJ.MaxLength = 18
    End If
    'txtbxCNPJ.Text = ""
End Sub

Private Sub txtbxCNPJ_GotFocus()
    P_SelecTexto txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    
    
    If optnPessoa(0).Value Then
        If Not F_CPF(txtbxCNPJ.Text) Then
            'MsgBox "CPF Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
            
        Else
            bCNPJ = True
            P_FormataCPF txtbxCNPJ
        End If
        
    Else
        If Not F_CGC(txtbxCNPJ.Text) Then
            'MsgBox "CNPJ Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
            
        End If
        P_FormataCGC txtbxCNPJ
    End If
    
    
End Sub




Private Sub txtbxRazao_GotFocus()
    P_SelecTexto txtbxRazao
End Sub

Private Sub txtbxRazao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxRazao_LostFocus()
    P_AjustaTexto txtbxRazao
End Sub


