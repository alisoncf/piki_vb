VERSION 5.00
Begin VB.Form frmCadFun 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Funcionários"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   Icon            =   "frmCadFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10935
   Begin VB.ComboBox cmbxUF 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtbxNatural 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   30
      Top             =   3840
      Width           =   3735
   End
   Begin VB.ComboBox cmbxStatus 
      Height          =   315
      ItemData        =   "frmCadFun.frx":058A
      Left            =   120
      List            =   "frmCadFun.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   5640
      Width           =   3735
   End
   Begin VB.ComboBox cmbxCar 
      Height          =   315
      ItemData        =   "frmCadFun.frx":05B0
      Left            =   120
      List            =   "frmCadFun.frx":05BD
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   5040
      Width           =   3735
   End
   Begin VB.TextBox txtbxTel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   23
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtbxConjuge 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   32
      Top             =   4440
      Width           =   3735
   End
   Begin VB.ComboBox cmbxSexo 
      Height          =   315
      ItemData        =   "frmCadFun.frx":05F8
      Left            =   3360
      List            =   "frmCadFun.frx":0602
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtbxEndPais 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   28
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox txtbxPai 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   26
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txtbxMae 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   24
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox txtbxDataNasc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   40
      TabIndex        =   11
      Top             =   840
      Width           =   1110
   End
   Begin VB.TextBox txtbxCEP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      MaxLength       =   14
      TabIndex        =   19
      Text            =   "75.000-000"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtbxCidade 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   17
      Text            =   "Anápolis"
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox txtbxBairro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   20
      TabIndex        =   15
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtbxEndereco 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   40
      TabIndex        =   13
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox txtbxRg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      MaxLength       =   18
      TabIndex        =   5
      Top             =   240
      Width           =   1095
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
      Left            =   5040
      MaxLength       =   18
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtbxGuerra 
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
      TabIndex        =   7
      Top             =   840
      Width           =   3135
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
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   10905
      TabIndex        =   37
      Top             =   6915
      Width           =   10935
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "F7 - Sal&var"
         Height          =   735
         Left            =   7320
         Picture         =   "frmCadFun.frx":061B
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "ESC - &Fecha&r"
         Height          =   735
         Left            =   9120
         Picture         =   "frmCadFun.frx":0777
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   5520
         Picture         =   "frmCadFun.frx":08E2
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Histórico"
         Height          =   735
         Left            =   3720
         Picture         =   "frmCadFun.frx":0CEF
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnFoto 
         Caption         =   "Fotos"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         Picture         =   "frmCadFun.frx":0DD9
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnPerm 
         Caption         =   "Permissões"
         Height          =   735
         Left            =   1920
         Picture         =   "frmCadFun.frx":0E8E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Label lblIdade 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6240
      TabIndex        =   46
      Top             =   855
      Width           =   615
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Idade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6240
      TabIndex        =   45
      Top             =   600
      Width           =   450
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "&UF:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5040
      TabIndex        =   20
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefones:"
      Height          =   195
      Left            =   5760
      TabIndex        =   22
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Côn&juge:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   630
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "&Data Nasc:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5040
      TabIndex        =   10
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&RG:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6960
      TabIndex        =   4
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Car&go:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Chama-se"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "S&ituação:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   5400
      Width           =   675
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Naturalidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   3600
      Width           =   945
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "&Sexo:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "End. Pais:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label lblMae 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da &Mãe:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "CE&P:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   18
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ci&dade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Bairro:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5040
      TabIndex        =   14
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblCpf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&CPF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5040
      TabIndex        =   2
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Nome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   555
   End
   Begin VB.Label lblPai 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do &Pai:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3960
      TabIndex        =   25
      Top             =   2400
      Width           =   960
   End
End
Attribute VB_Name = "frmCadFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Dim bActive As Boolean


Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim cFile As New Collection

Dim Indice As Integer




Sub P_Salva()
    P_GravaTelefones
End Sub

Sub P_GravaTelefones()
    Dim strSQL As String
    Dim iCount As Integer
    Dim sTipo As String
    Dim sFone As String
    strSQL = "delete from tab_telefones where cod_cli='" & Me.tag & "' "
    strSQL = strSQL & " and tipo_ter='U' "
    conn.Execute strSQL
    With msflxgrdTel
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0: sFone = .Text
            .col = 1: sTipo = Left(.Text, 1)
            If F_GravaTelefone(Me.tag, sFone, sTipo, "U") Then
            End If
            
        Next
    End With
End Sub



Private Sub cmbxEstadoCiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub cmbxCar_Click()
    cmbxCar.tag = F_BuscaUmCampo("cod_Car", "tab_cargos", "nome_car='" & cmbxCar.Text & "'")
    
    
End Sub

Private Sub cmbxCar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxStatus_click()
    If cmbxStatus.ListIndex <> 0 Then
        
    Else
        
    End If
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub










Private Sub cmndbtnHist_Click()
    If Me.tag = "" Then
    Else
        P_histCli Me.tag
    End If
End Sub


Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    Dim iCount As Integer
    If strCadFun_Usu <> "S" Then
        MsgBox "Você não tem permissão para cadastrar funcionários", vbExclamation
        Exit Sub
    End If
    
    If txtbxRazao.Text = "" Then
        MsgBox "Informe um nome para o funcionario!", vbExclamation
        Exit Sub
        
    ElseIf F_Vazio(txtbxCNPJ.Text) And F_Vazio(txtbxRg.Text) Then
        
            MsgBox "Informe o CPF  ou CI do funcionario !", vbExclamation
            txtbxCNPJ.SetFocus
            Exit Sub
    ElseIf cmbxCar.tag = "" Then
        MsgBox "Informe o cargo do Funcionário", vbExclamation
        cmbxCar.SetFocus
        Exit Sub
    ElseIf txtbxEndereco.Text = "" Then
        MsgBox "Informe o endereço do funcionario!", vbExclamation
        Exit Sub
    ElseIf txtbxBairro.Text = "" Then
        MsgBox "informe o Bairro do funcionario!", vbExclamation
        Exit Sub
    ElseIf txtbxCidade.Text = "" Then
        MsgBox "Informe a Cidade do funcionario!", vbExclamation
        Exit Sub
    ElseIf cmbxUF.Text = "" Then
        MsgBox "Informe o Estado de Origem do funcionario!", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("cod_fun", TAB_FUNCIONARIOS, "cpf_fun='" & txtbxCNPJ.Text & "'") <> "" Then
            MsgBox "Já existe um funcionario com este CPF/CNPJ!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_fun", TAB_FUNCIONARIOS, "nome_fun='" & txtbxRazao.Text & "'") <> "" Then
            MsgBox "Já existe um funcionario com este Nome!", vbCritical
            Exit Sub
        
        End If
    
        strSQL = funcionario.F_GravaFun(txtbxRazao.Text, txtbxGuerra, txtbxEndereco, _
            txtbxBairro.Text, txtbxCidade.Text, cmbxUF.Text, txtbxCEP.Text, txtbxDataNasc.Text, txtbxNatural.Text, _
            txtbxRg.Text, txtbxCNPJ.Text, Left(cmbxSexo.Text, 1), txtbxConjuge.Text, txtbxPai.Text, _
            txtbxMae.Text, txtbxEndPais.Text, "", cmbxCar.tag, txtbxTel.Text)
        
        
'        P_GravaTelefones
        
        MsgBox "funcionario Cadastrado com sucesso!", vbInformation
        P_Limpa
        
        txtbxRazao.Text = ""
        Me.tag = ""
    Else
        If MsgBox("Confirma alteração do funcionario " & Me.tag & " ?", vbYesNo + vbQuestion) = vbYes Then
            If funcionario.F_AlteraFun(Me.tag, txtbxRazao.Text, txtbxGuerra.Text, txtbxEndereco.Text, txtbxBairro.Text, _
                 txtbxCidade.Text, cmbxUF.Text, txtbxCEP.Text, txtbxDataNasc.Text, txtbxNatural.Text, _
                 txtbxRg.Text, txtbxCNPJ.Text, Left(cmbxSexo.Text, 1), txtbxConjuge.Text, _
                    txtbxPai.Text, txtbxMae.Text, txtbxEndPais.Text, "", Left(cmbxStatus.Text, 1), cmbxCar.tag, txtbxTel.Text) Then
                MsgBox "funcionario alterado com sucesso!", vbInformation
            End If
                
            Unload Me
        End If
    End If
    
End Sub
Sub P_Limpa()
    Dim i As Integer
    For i = 0 To Me.Controls.Count - 1
        If LCase(Left(Me.Controls(i).Name, 3)) = "txt" Then
            Me.Controls(i).Text = ""
        End If
    Next
   
    cmbxUF.ListIndex = -1
End Sub







Private Sub Form_Activate()
    
    If bActive Then Exit Sub
    
    
   ' MudaAparencia Me
    
    
    
    
    If Me.tag <> "" Then
        Carrega
        
    End If
    
    
    
    
    
bActive = True
    
End Sub
Sub Carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select * from tab_funcionarios where cod_fun = '" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do Until .EOF
            txtbxBairro.Text = Format(!bairro_fun)
            txtbxCEP.Text = Format(!cep_fun)
            txtbxCidade.Text = Format(!cidade_fun)
            txtbxCNPJ.Text = Format(!cpf_fun)
            txtbxRg.Text = Format(!rg_fun)
            txtbxEndereco.Text = Format(!endereco_fun)
            txtbxMae.Text = Format(!mae_fun)
            txtbxPai.Text = Format(!pai_fun)
            txtbxRazao.Text = !nome_fun
            txtbxDataNasc.Text = Format(!nasc_fun)
            txtbxEndPais.Text = Format(!endpais_fun)
            txtbxNatural.Text = Format(!natural_fun)
            P_LocalizaOperacao Format(!sexo_fun), cmbxSexo
            txtbxConjuge.Text = Format(!conjuge_fun)
            calcIdade
            
            txtbxRg.Text = Format(!rg_fun)
            
            txtbxGuerra.Text = !guerra_fun
            cmbxCar.Text = F_BuscaUmCampo("nome_car", "tab_cargos", "cod_car='" & !cod_Car & "'")
            If !uf_fun <> "" Then cmbxUF.Text = !uf_fun
            
            
            
            
            
            


            
            txtbxTel.Text = Format(!telefone_fun)
            P_LocalizaOperacao !status_fun, cmbxStatus
            'P_CarregaTelefone
                    
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    bActive = False
    'Me.Height = 6540
    'Me.Width = 6255
   ' P_CentraForm Me
    
    'resizeForm Me
    P_PreencheCargos cmbxCar
    
    If Me.tag = "" Then
        bCNPJ = True
    End If
    'stb.Tab = 0
    
    
    
    
    If strContrMens_usu <> "S" Then
'        chkMens.Enabled = False
'        txtbxLimite.Enabled = False
    Else
'        chkMens.Enabled = True
'        txtbxLimite.Enabled = True
    End If
'    With msflxgrdTel
'        .FixedCols = 0
'        .rows = 1
'        .Cols = 2
'        .col = 0: .Text = "Fone:"
'        .col = 1: .Text = "Tipo:"
'        .ColWidth(0) = 1500
'        .ColWidth(1) = 1200
'        .ColAlignment(0) = 1
'        .ColAlignment(1) = 1
'    End With
    On Error Resume Next
    P_PreencheUF cmbxUF
    cmbxUF.Text = "GO"
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad(frmEdiFun.Name) Then
        'RefreshDataPosition frmEdiFun.adc
    End If
End Sub







Private Sub txtbxBairro_GotFocus()
    P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    P_AjustaTexto txtbxBairro
End Sub



Private Sub txtbxCEP_GotFocus()
    P_SelecTexto txtbxCEP
End Sub

Private Sub txtbxCEP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    
    
    
        If Not F_CPF(txtbxCNPJ.Text) Then
            MsgBox "CPF Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCPF txtbxCNPJ
    
End Sub



Private Sub txtbxConjuge_GotFocus()
    P_SelecTexto txtbxConjuge
End Sub

Private Sub txtbxConjuge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxConjuge_LostFocus()
    P_AjustaTexto txtbxConjuge
End Sub

Private Sub txtbxDataNasc_GotFocus()
    P_SelecTexto txtbxDataNasc
End Sub

Private Sub txtbxDataNasc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataNasc_LostFocus()
    P_AjustaData txtbxDataNasc
    calcIdade
End Sub
Sub calcIdade()
    On Error Resume Next
    If txtbxDataNasc.Text = "" Then
        Exit Sub
    End If
    lblIdade.caption = Format(DateDiff("M", txtbxDataNasc.Text, Date) / 12, "##0.0")
End Sub
Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
End Sub


Private Sub txtbxEndPais_GotFocus()
    P_SelecTexto txtbxEndPais
End Sub

Private Sub txtbxEndPais_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEndPais_LostFocus()
    P_AjustaTexto txtbxEndPais
End Sub

Private Sub txtbxGuerra_GotFocus()
    P_SelecTexto txtbxGuerra
End Sub

Private Sub txtbxGuerra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxGuerra_LostFocus()
    P_AjustaTexto txtbxGuerra
End Sub

Private Sub txtbxMae_GotFocus()
    P_SelecTexto txtbxMae
End Sub

Private Sub txtbxMae_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMae_LostFocus()
    P_AjustaTexto txtbxMae
End Sub

Private Sub txtbxNatural_GotFocus()
    P_SelecTexto txtbxNatural
End Sub

Private Sub txtbxNatural_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNatural_LostFocus()
    P_AjustaTexto txtbxNatural
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSalvar.SetFocus
    End If
End Sub

Private Sub txtbxPai_GotFocus()
    P_SelecTexto txtbxPai
End Sub

Private Sub txtbxPai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPai_LostFocus()
    P_AjustaTexto txtbxPai
End Sub

Private Sub txtbxRazao_GotFocus()
    P_SelecTexto txtbxRazao
End Sub

Private Sub txtbxRazao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxRazao_LostFocus()
    P_AjustaTexto txtbxRazao
End Sub


Private Sub txtbxRg_GotFocus()
    P_SelecTexto txtbxRg
End Sub

Private Sub txtbxRg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxRg_LostFocus()
    P_AjustaTexto txtbxRg
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
End Sub

Sub P_CarregaTelefone()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipo As String
    msflxgrdTel.rows = 1
    strSQL = "select * from tab_telefones where cod_CLI='" & Me.tag & "' and tipo_ter='U' "
    rPesq.Open strSQL, conn
    With rPesq
        Do While Not .EOF
            If !tipo_tel = "1" Then
                sTipo = "1 - Residencial"
            ElseIf !tipo_tel = "2" Then
                sTipo = "2 - Comercial"
            ElseIf !tipo_tel = "3" Then
                sTipo = "3 - Celular"
            ElseIf !tipo_tel = "4" Then
                sTipo = "4 - Recado"
            ElseIf !tipo_tel = "5" Then
                sTipo = "5 - Fax"
            End If
            strSQL = !num_tel & Chr(9) & sTipo
            msflxgrdTel.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub
Private Sub txtbxTel_GotFocus()
    P_SelecTexto txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub


Private Sub cmbxTipoTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub




Sub deletaArquivos()
    On Error Resume Next
    Dim iCount As Integer
    For iCount = 1 To cFile.Count
        Kill cFile(iCount)
    Next
    Set cFile = Nothing
End Sub

