VERSION 5.00
Begin VB.Form frmIncFor 
   AutoRedraw      =   -1  'True
   Caption         =   "Inclusão de Fornecedores"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   Icon            =   "frmIncTer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   5265
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   19
      Top             =   3970
      Width           =   5175
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   240
         Picture         =   "frmIncTer.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   2640
         Picture         =   "frmIncTer.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Fornecedor:"
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtbxRazao 
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   2
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox txtbxCNPJ 
         Height          =   285
         Left            =   240
         MaxLength       =   18
         TabIndex        =   4
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtbxEndereco 
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox txtbxIE 
         Height          =   285
         Left            =   2640
         MaxLength       =   16
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtbxBairro 
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2280
         Width           =   4695
      End
      Begin VB.TextBox txtbxCidade 
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   12
         Top             =   2880
         Width           =   3975
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   4320
         TabIndex        =   14
         Text            =   "GO"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtbxCEP 
         Height          =   285
         Left            =   240
         MaxLength       =   16
         TabIndex        =   16
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtbxTel 
         Height          =   285
         Left            =   2640
         MaxLength       =   16
         TabIndex        =   18
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Razão Social:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&NPJ:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Inscrição Estadual:"
         Height          =   195
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bairro:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&UF"
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   2640
         Width           =   210
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CE&P:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Telefone:"
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         Top             =   3240
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmIncFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean

Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_erro
    sCod = F_BuscaUmCampo(" max(cod_for)", "tab_fornecedores", "")
    If sCod = "" Then       'pega só 4 dígitos para compor o dígito verificador
        sCod = "0001"       'que é o último número do cod_cli
    Else
        sCod = Left(sCod, 4)
        sCod = Format(Int(sCod) + 1, "000#")
    End If
    sCod = F_GeraDig(sCod)
    
    
    
    strSQL = "insert into tab_fornecedores(cod_for,nome_for,cnpj_for,ie_for,endereco_for, "
    strSQL = strSQL & "bairro_for,cidade_for,uf_for,cep_for,telefone_for) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "'" & txtbxIE.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "'" & txtbxBairro.Text & "',"
    strSQL = strSQL & "'" & txtbxCidade.Text & "',"
    strSQL = strSQL & "'" & cmbxUF.Text & "',"
    strSQL = strSQL & "'" & txtbxCEP.Text & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "')"
    
    banco.Execute strSQL
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_fornecedor set "
    strSQL = strSQL & "nome_for='" & txtbxRazao.Text & "',"
    strSQL = strSQL & "cnpj_for='" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "ie_for='" & txtbxIE.Text & "',"
    strSQL = strSQL & "endereco_for='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "bairro_for='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "cidade_for='" & txtbxCidade.Text & "',"
    strSQL = strSQL & "uf_for='" & cmbxUF.Text & "',"
    strSQL = strSQL & "cep_for='" & txtbxCEP.Text & "',"
    strSQL = strSQL & "telefone_for='" & txtbxTel.Text & "' "
    
    strSQL = strSQL & " where cod_for = '" & Me.Tag & "'"
    banco.Execute strSQL
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    
    
     If txtbxRazao.Text = "" Then
        MsgBox "Informe um nome para o Fornecedor!", vbExclamation
        Exit Sub
    ElseIf txtbxCNPJ.Text = "" Then
        MsgBox "Informe o CNPJ do Fornecedor!", vbExclamation
        Exit Sub
    ElseIf bCNPJ = False Then
        MsgBox "Informe um CNPJ Válido para Fornecedor!", vbExclamation
        Exit Sub
    ElseIf txtbxEndereco.Text = "" Then
        MsgBox "Informe o endereço do Fornecedor!", vbExclamation
        Exit Sub
    ElseIf txtbxBairro.Text = "" Then
        MsgBox "informe o Bairro do Fornecedor!,vbExclamation  "
        Exit Sub
    ElseIf txtbxCidade.Text = "" Then
        MsgBox "Informe a Cidade do Fornecedor!", vbExclamation
        Exit Sub
    ElseIf txtbxCEP.Text = "" Then
        MsgBox "Informe o CEP do Fornecedor!", vbExclamation
        Exit Sub
   
    ElseIf cmbxUF.Text = "" Then
        MsgBox "Informe o Estado de Origem do Fornecedor!", vbExclamation
        Exit Sub
    End If
    If Me.Tag = "" Then
        If F_BuscaUmCampo("cod_For", "tab_fornecedores", "cnpj_for='" & txtbxCNPJ.Text & "'") <> "" Then
            MsgBox "Já existe um Fornecedor com este CNPJ!", vbCritical
            Unload Me
        ElseIf F_BuscaUmCampo("cod_for", "tab_Fornecedor", "nome_for='" & txtbxRazao.Text & "'") <> "" Then
            MsgBox "Já existe um fornecedor com esta Razão Social!", vbCritical
            Unload Me
        ElseIf F_BuscaUmCampo("cod_for", "tab_fornecer", "IE_for='" & txtbxIE.Text & "'") <> "" Then
            If txtbxIE.Text <> "" Then
                MsgBox "Já existe um fornecedor com esta Inscrição Estadual!", vbCritical
                Unload Me
            End If
        End If
    
        P_Salva
        MsgBox "Cliente Cadastrado com sucesso!", vbInformation
        Unload Me
    Else
        P_Altera
        MsgBox "Cliente Alterado com sucesso!", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    P_PreencheUF cmbxUF
     If Me.Tag <> "" Then
        strSQL = "select * from tab_FORNECEDORES where cod_FOR = '" & Me.Tag & "' "
        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
           
                txtbxBairro.Text = !bairro_FOR
                txtbxCEP.Text = !cep_for
                txtbxCidade.Text = !cidade_for
                txtbxCNPJ.Text = !cnpj_for
                txtbxEndereco.Text = !endereco_for
                txtbxIE.Text = !ie_for
                txtbxRazao.Text = !nome_for
                txtbxTel.Text = !telefone_for
                cmbxUF.Text = !uf_for
               
                
           
        End With
    End If
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Me.Height = 5520
    Me.Width = 5385
    P_CentraForm Me
    If Me.Tag = "" Then
        bCNPJ = True
    End If
   
End Sub

Private Sub Form_Resize()
    P_CentraObejct Frame1, Me, False
    P_CentraObejct Frame2, Me, False
End Sub

Private Sub Form_Unload(cancel As Integer)
    If IsLoad(frmEdiFor.Name) Then
        frmEdiFor.dtClientes.Refresh
    End If
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

Private Sub txtbxCNpj_GotFocus()
    P_SelecTexto txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNpj_LostFocus()
    If Not F_CGC(txtbxCNPJ.Text) Then
        MsgBox "CNPJ Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
        bCNPJ = False
    Else
        bCNPJ = True
    End If
    P_FormataCGC txtbxCNPJ
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

Private Sub txtbxTel_gotFocus()
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

