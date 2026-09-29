VERSION 5.00
Begin VB.Form frmIncClientes 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "frmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11475
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11445
      TabIndex        =   58
      Top             =   6345
      Width           =   11475
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Fecha&r"
         Height          =   735
         Left            =   9840
         Picture         =   "frmClientes.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "Sal&var"
         Height          =   735
         Left            =   8280
         Picture         =   "frmClientes.frx":0E14
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox pcOutrosDados 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5805
      Left            =   7200
      ScaleHeight     =   5775
      ScaleWidth      =   4185
      TabIndex        =   33
      Top             =   480
      Width           =   4215
      Begin VB.TextBox txtbxLimite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   37
         Text            =   "0,00"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmClientes.frx":0F5E
         Left            =   120
         List            =   "frmClientes.frx":0F68
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtbxObs 
         Appearance      =   0  'Flat
         Height          =   3285
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restante:"
         Height          =   195
         Left            =   3120
         TabIndex        =   54
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label lblRestante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   2880
         TabIndex        =   53
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em Uso:"
         Height          =   195
         Left            =   1800
         TabIndex        =   52
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblUso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1560
         TabIndex        =   51
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de Crédito:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação do Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Observações/Anotações:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7200
      ScaleHeight     =   375
      ScaleWidth      =   4185
      TabIndex        =   48
      Top             =   120
      Width           =   4215
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outros Dados:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   49
         Top             =   0
         Width           =   1650
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   6900
      TabIndex        =   46
      Top             =   4440
      Width           =   6930
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   47
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   6900
      TabIndex        =   44
      Top             =   2520
      Width           =   6930
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   45
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   6900
      TabIndex        =   42
      Top             =   120
      Width           =   6930
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   55
         Top             =   0
         Width           =   4575
         Begin VB.OptionButton optnPessoa 
            BackColor       =   &H000000C0&
            Caption         =   "Pessoa &Jurídica"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   57
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton optnPessoa 
            BackColor       =   &H000000C0&
            Caption         =   "Pessoa &Física"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   56
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identificação:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   43
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   6900
      TabIndex        =   24
      Top             =   4800
      Width           =   6930
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   16
         TabIndex        =   32
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txtbxFax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         MaxLength       =   16
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   16
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxCel_cli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Email:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Left            =   3480
         TabIndex        =   29
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Celular:"
         Height          =   195
         Left            =   1800
         TabIndex        =   27
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   6900
      TabIndex        =   13
      Top             =   2880
      Width           =   6930
      Begin VB.TextBox txtbxCEP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   14
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   21
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtbxBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   19
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtbxEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   15
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Height          =   195
         Left            =   5880
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CE&P:"
         Height          =   195
         Left            =   5520
         TabIndex        =   16
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         Height          =   195
         Left            =   3480
         TabIndex        =   20
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bairro: (Setor, Vila)"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço: (Rua, Av, Logradouro, Complemento, etc)"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   3720
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   6900
      TabIndex        =   0
      Top             =   480
      Width           =   6930
      Begin VB.TextBox txtbxMae 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   12
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox txtbxInscProdRural 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   16
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtbxIE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   16
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtbxCNPJ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   18
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtbxNasc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   18
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxRazao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblMae 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da &Mãe"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Insc. Prod. Rural:"
         Height          =   195
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lblRg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RG:"
         Height          =   195
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   285
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CPF:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblNasc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Nasc:"
         Height          =   195
         Left            =   5520
         TabIndex        =   3
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmIncClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL("tab_clientes", "cod_cli", "", "0000#")
    
    
    
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,cpf_cli,rg_cli,inscprodrur_cli, endereco_Cli, "
    strSQL = strSQL & "bairro_cli,cidade_Cli,uf_cli,cep_cli,telefone_Cli,fax_Cli,"
    strSQL = strSQL & "nasc_cli,mae_cli,"
    strSQL = strSQL & "status_cli,tipo_cli,cel_cli,obs_Cli,limite_cli) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "'" & txtbxIE.Text & "',"
    strSQL = strSQL & "'" & txtbxInscProdRural.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "'" & txtbxBairro.Text & "',"
    strSQL = strSQL & "'" & txtbxCidade.Text & "',"
    strSQL = strSQL & "'" & cmbxUF.Text & "',"
    strSQL = strSQL & "'" & txtbxCEP.Text & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "',"
    strSQL = strSQL & "'" & txtbxFax.Text & "',"
    strSQL = strSQL & F_GravaData(txtbxNasc.Text) & ","
   strSQL = strSQL & "'" & txtbxMae.Text & "',"
 
    
    strSQL = strSQL & "'" & Left(cmbxSituacao.Text, 1) & "',"
    If optnPessoa(0).Value = True Then
        strSQL = strSQL & "'F',"
    Else
        strSQL = strSQL & "'J',"
    End If
    strSQL = strSQL & "'" & txtbxCel_cli.Text & "',"
    strSQL = strSQL & "'" & txtbxObs.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxLImite.Text) & ") "
    banco.Execute strSQL
    
End Sub
Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_clientes set "
    strSQL = strSQL & "nome_cli='" & txtbxRazao.Text & "',"
    strSQL = strSQL & "cpf_Cli='" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "rg_cli='" & txtbxIE.Text & "',"
    strSQL = strSQL & "inscprodrur_cli='" & txtbxInscProdRural.Text & "',"
    strSQL = strSQL & "endereco_Cli='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "bairro_Cli='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "cidade_Cli='" & txtbxCidade.Text & "',"
    strSQL = strSQL & "uf_Cli='" & cmbxUF.Text & "',"
    strSQL = strSQL & "cep_Cli='" & txtbxCEP.Text & "',"
    strSQL = strSQL & "telefone_Cli='" & txtbxTel.Text & "',  "
     strSQL = strSQL & "fax_Cli='" & txtbxFax.Text & "',  "
    strSQL = strSQL & "cel_Cli='" & txtbxCel_cli.Text & "',  "
    strSQL = strSQL & "obs_Cli='" & txtbxObs.Text & "',  "
    strSQL = strSQL & "NASC_CLI=" & F_GravaData(txtbxNasc.Text) & ","
    strSQL = strSQL & "limite_cli=" & F_GravaDec(txtbxLImite.Text) & ","
    
    strSQL = strSQL & "mae_cli='" & txtbxMae.Text & "',"


    strSQL = strSQL & "status_cli='" & Left(cmbxSituacao.Text, 1) & "' "
         
         
    strSQL = strSQL & " where cod_Cli = '" & Me.Tag & "'"
    banco.Execute strSQL
End Sub

Private Sub cmbxEstadoCiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
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
            'Exit Sub
        Else
            MsgBox "Informe um CNPJ Válido para Cliente!", vbExclamation
            'Exit Sub
        End If
    ElseIf txtbxEndereco.Text = "" Then
        MsgBox "Informe o endereço do Cliente!", vbExclamation
        Exit Sub
    ElseIf txtbxBairro.Text = "" Then
        MsgBox "informe o Bairro do Cliente!,vbExclamation  "
        Exit Sub
    ElseIf txtbxCidade.Text = "" Then
        MsgBox "Informe a Cidade do Cliente!", vbExclamation
        Exit Sub
    ElseIf cmbxUF.Text = "" Then
        MsgBox "Informe o Estado de Origem do Cliente!", vbExclamation
        Exit Sub
    End If
    
    If Me.Tag = "" Then
        If F_BuscaUmCampo("cod_cli", "tab_clientes", "cpf_cli='" & txtbxCNPJ.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este CPF/CNPJ!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_cli", "tab_clientes", "nome_cli='" & txtbxRazao.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este Nome/Razão Social!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_cli", "tab_clientes", "rg_cli='" & txtbxIE.Text & "'") <> "" Then
            If txtbxIE.Text <> "" Then
                MsgBox "Já existe um cliente com esta Inscrição Estadual /RG! ", vbCritical
                Exit Sub
            End If
        End If
    
        P_Salva
        MsgBox "Cliente Cadastrado com sucesso!", vbInformation
        P_Limpa

        txtbxRazao.Text = ""
        optnPessoa(0).SetFocus
    Else
        If MsgBox("Confirma alteração do cliente " & Me.Tag & " ?", vbYesNo + vbQuestion) = vbYes Then
            P_Altera
            MsgBox "Cliente Alterado com sucesso!", vbInformation
            Unload Me
        End If
    End If
    
End Sub
Sub P_Limpa()
    Dim i As Integer
    For i = 0 To Me.Controls.count - 1
        If LCase(Left(Me.Controls(i).Name, 3)) = "txt" Then
            Me.Controls(i).Text = ""
        End If
    Next
   
    cmbxUF.ListIndex = -1
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Sub P_StatusCli()
    Dim strSQL As String
    Dim rPesq As Recordset
    lblPrimCompr.caption = F_BuscaUmCampo("min(data_ven)", "tab_vendas", "cod_cli='" & Me.Tag & "'")
    lblUtlCompr.caption = F_BuscaUmCampo("max(dataefet_ven)", "tab_vendas", "cod_cli='" & Me.Tag & "'")
    lblMaiorComp.caption = "R$ " & FormatVal(F_BuscaUmCampo("max(valor_ven-desconto_ven)", "tab_vendas", "cod_cli='" & Me.Tag & "'"))
    lblCompras.caption = F_BuscaUmCampo("sum(valor_ven-desconto_ven)", "tab_vendas", "cod_cli='" & Me.Tag & "'")
    lblParc.caption = F_BuscaUmCampo("count(*)", "tab_itdoc1", "venc_itd<=" & F_GravaData(Date) & " and reman_itd>0 and cod_ter='" & Me.Tag & "'")
    lblParcAb.caption = F_BuscaUmCampo("count(*)", "tab_itdoc1", "reman_itd>0 and cod_ter='" & Me.Tag & "'")
    
    
End Sub
Private Sub Form_Activate()
    Dim rPesq As Recordset
    Dim strSQL As String
    P_PreencheUF cmbxUF
    'P_PreencheUF cmbxUfFirma
    
    If Me.Tag <> "" Then
        strSQL = "select * from tab_clientes where cod_cli = '" & Me.Tag & "' "
        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
            Do Until .EOF
                txtbxBairro.Text = Format(!bairro_cli)
                txtbxCEP.Text = Format(!cep_cli)
                txtbxCidade.Text = Format(!cidade_cli)
                txtbxCNPJ.Text = !cpf_cli
                txtbxEndereco.Text = Format(!endereco_cli)
                txtbxIE.Text = Format(!rg_cli)
                txtbxRazao.Text = !nome_Cli
                txtbxTel.Text = Format(!telefone_cli)
                txtbxObs.Text = Format(!obs_cli)
                txtbxNasc.Text = Format(!nasc_cli)
                txtbxCel_cli.Text = Format(!cel_cli)
                txtbxLImite.Text = FormatVal(!limite_cli)
                lblUso.caption = FormatVal(F_RetornaLimiteUsoCli(Me.Tag))
                lblRestante.caption = FormatVal(!limite_cli - CCur(lblUso.caption))
                If !uf_cli <> "" Then cmbxUF.Text = !uf_cli
                
                txtbxFax.Text = Format(!fax_cli)
                txtbxInscProdRural.Text = !inscprodrur_cli
                
                
                If !tipo_cli = "F" Then
                    optnPessoa(0).Value = True
                Else
                    optnPessoa(1).Value = True
                End If
                'optnPessoa(0).Enabled = False
                'optnPessoa(1).Enabled = False
                txtbxMae.Text = Format(!mae_cli)
                txtbxNasc.Text = Format(!nasc_cli)
                
                


                On Error Resume Next
                
                
                
                
                P_LocalizaOperacao !status_cli, cmbxSituacao
                P_StatusCli
                .MoveNext
            Loop
            .Close
        End With
    End If
'    MsgBox strHora & "-" & Time
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 7605 'cmndbtnSair.Height + cmndbtnSair.Top + 500
    Me.Width = pcOutrosDados.Left + pcOutrosDados.Width + 300
    Me.Width = 11595
        
    If Me.Tag = "" Then
        bCNPJ = True
    End If
    'resizeForm Me
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad(frmEdiClientes.Name) Then
        RefreshDataPosition frmEdiClientes.dtClientes
    End If
End Sub

Private Sub optnPessoa_Click(Index As Integer)
    If Index = 0 Then
        lblCpf.caption = "&CPF:"
        lblRg.caption = "&RG:"
        txtbxCNPJ.MaxLength = 14
        txtbxNasc.Enabled = True
        txtbxMae.Enabled = True
        lblNasc.Enabled = True
        lblMae.Enabled = True
    Else
        lblCpf.caption = "&CNPJ:"
        lblRg.caption = "&Insc Est:"
        txtbxCNPJ.MaxLength = 18
        txtbxNasc.Enabled = False
        txtbxMae.Enabled = False
        lblNasc.Enabled = False
        lblMae.Enabled = False
    End If
    'txtbxCNPJ.Text = ""
End Sub

Private Sub txtbxCel_cli_GotFocus()
    P_SelecTexto txtbxCel_cli
End Sub


Private Sub txtbxBairro_GotFocus()
    P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    P_AjustaTexto txtbxBairro
End Sub


Private Sub txtbxCel_cli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCel_cli_LostFocus()
    txtbxCel_cli.Text = F_AjustaTelefone(txtbxCel_cli.Text)
End Sub

Private Sub txtbxCEP_GotFocus()
    P_SelecTexto txtbxCEP
End Sub

Private Sub txtbxCEP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCEP_LostFocus()
    P_CEPLostFocus txtbxCEP
End Sub


Private Sub txtbxCidade_GotFocus()
    P_SelecTexto txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    P_AjustaTexto txtbxCidade
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
            MsgBox "CPF Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCPF txtbxCNPJ
    Else
        If Not F_CGC(txtbxCNPJ.Text) Then
            MsgBox "CNPJ Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCGC txtbxCNPJ
    End If
    
    
End Sub



Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
End Sub





Private Sub txtbxFax_GotFocus()
    P_SelecTexto txtbxFax
End Sub

Private Sub txtbxFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxFax_LostFocus()
    P_AjustaFone txtbxFax
End Sub

Private Sub txtbxIE_GotFocus()
    P_SelecTexto txtbxIE
End Sub

Private Sub txtbxIE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIE_LostFocus()
    P_AjustaTexto txtbxIE
End Sub

Private Sub txtbxInscProdRural_GotFocus()
    P_SelecTexto txtbxInscProdRural
End Sub

Private Sub txtbxInscProdRural_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxInscProdRural_LostFocus()
    P_AjustaTexto txtbxInscProdRural
End Sub

Private Sub txtbxLimite_GotFocus()
    P_SelecTexto txtbxLImite
End Sub

Private Sub txtbxLimite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxLimite_LostFocus()
    P_ValorLostFocus txtbxLImite
End Sub

Private Sub txtbxMae_GotFocus()
    P_SelecTexto txtbxMae
End Sub

Private Sub txtbxMae_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        SendKeys Chr(9)
        
    End If
End Sub

Private Sub txtbxMae_LostFocus()
    P_AjustaTexto txtbxMae
    
End Sub

Private Sub txtbxNasc_GotFocus()
    P_SelecTexto txtbxNasc

End Sub

Private Sub txtbxNasc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNasc_LostFocus()
    P_AjustaData txtbxNasc
End Sub



Private Sub txtbxObs_GotFocus()
    P_SelecTexto txtbxObs
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then
    '    SendKeys Chr(9)
    'End If
End Sub

Private Sub txtbxObs_LostFocus()
    P_AjustaTexto txtbxObs


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


Private Sub txtbxTel_GotFocus()
    P_SelecTexto txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        
        SendKeys Chr(9)

        
    End If
    
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
    
End Sub

