VERSION 5.00
Begin VB.Form frmIncFor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão de Fornecedores"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8340
   Icon            =   "frmIncFor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   5760
      Picture         =   "frmIncFor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   3240
      Picture         =   "frmIncFor.frx":0537
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1125
         Left            =   1440
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3840
         Width           =   5895
      End
      Begin VB.TextBox txtbxContato 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   26
         Top             =   3480
         Width           =   5895
      End
      Begin VB.TextBox txtbxEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   22
         Top             =   3120
         Width           =   5895
      End
      Begin VB.TextBox txtbxFax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   16
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtbxFan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox txtbxRazao 
         Appearance      =   0  'Flat
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
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   2
         Top             =   240
         Width           =   5895
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
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtbxEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox txtbxIE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   16
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtbxBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   14
         Top             =   2040
         Width           =   4815
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtbxCEP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   18
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtbxTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   20
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblAno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Anotações:"
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   165
         TabIndex        =   28
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Co&ntato"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   25
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Emai&l"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   21
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblFax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fa&x"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4200
         TabIndex        =   23
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fan&tasia:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Razão Social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C&NPJ:"
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
         Height          =   285
         Left            =   165
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Endereço:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Insc. Est:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Bairro:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ci&dade:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   13
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&UF:"
         Height          =   195
         Left            =   6360
         TabIndex        =   15
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CE&P:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Telefone:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   165
         TabIndex        =   19
         Top             =   2760
         Width           =   1215
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
    On Error GoTo trata_Erro
    sCod = F_BuscaUmCampo(" max(cod_for)", "tab_fornecedores", "")
    If sCod = "" Then       'pega só 4 dígitos para compor o dígito verificador
        sCod = "0001"       'que é o último número do cod_cli
    Else
        sCod = Left(sCod, 4)
        sCod = Format(Int(sCod) + 1, "000#")
    End If
    sCod = F_GeraDig(sCod)
    
    
    
    strSQL = "insert into tab_fornecedores(cod_for,nome_for,fantasia_for,cnpj_for,ie_for,endereco_for, "
    strSQL = strSQL & "bairro_for,cidade_for,uf_for,cep_for,telefone_for,fax_for,contato_for,email_for) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & "'" & txtbxFan.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "'" & txtbxIE.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "'" & txtbxBairro.Text & "',"
    strSQL = strSQL & "'" & txtbxCidade.Text & "',"
    strSQL = strSQL & "'" & cmbxUF.Text & "',"
    strSQL = strSQL & "'" & txtbxCEP.Text & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "',"
    strSQL = strSQL & "'" & txtbxFax.Text & "',"
    strSQL = strSQL & "'" & txtbxContato.Text & "',"
    strSQL = strSQL & "'" & txtbxEmail.Text & "')"
    
    banco.Execute strSQL
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_fornecedores set "
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
    ElseIf txtbxFan.Text = "" Then
        MsgBox "Informe um Nome Fantasia para o Fornecedor!", vbExclamation
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
        MsgBox "Fornecedor Cadastrado com sucesso!", vbInformation
        txtbxBairro.Text = ""
        txtbxCEP.Text = ""
        txtbxCidade.Text = ""
        txtbxCNPJ.Text = ""
        txtbxContato.Text = ""
        txtbxEmail.Text = ""
        txtbxEndereco.Text = ""
        txtbxFan.Text = ""
        txtbxFax.Text = ""
        txtbxIE.Text = ""
        txtbxRazao.Text = ""
        txtbxTel.Text = ""
        
    Else
        P_Altera
        MsgBox "Fornecedor Alterado com sucesso!", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
     Dim rPesq As Recordset
     Dim strSQL As String
     On Error Resume Next
     If Me.Tag <> "" Then
        strSQL = "select * from tab_FORNECEDORES where cod_FOR = '" & Me.Tag & "' "
        Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
           
                txtbxBairro.Text = !bairro_for
                txtbxCEP.Text = !cep_for
                txtbxCidade.Text = !cidade_For
                txtbxCNPJ.Text = !cnpj_for
                txtbxEndereco.Text = !endereco_for
                txtbxIE.Text = !ie_for
                txtbxRazao.Text = !nome_for
                txtbxTel.Text = !telefone_for
                txtbxContato.Text = !contato_for
                txtbxEmail.Text = !email_for
                txtbxFan.Text = !fantasia_for
                txtbxTel.Text = !tel_for
                txtbxFax.Text = !fax_for
                cmbxUF.Text = !uf_for
               
                
            .Close
        End With
    End If
    
End Sub

Private Sub Form_Load()
    P_PreencheUF cmbxUF
   ' Me.Height = 5520
   ' Me.Width = 5385
    P_CentraForm Me
    If Me.Tag = "" Then
        bCNPJ = True
    End If
   
End Sub

Private Sub Form_Resize()
'    P_CentraObejct Frame1, Me, False
'    P_CentraObejct Frame2, Me, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
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

Private Sub txtbxCNPJ_GotFocus()
    P_SelecTexto txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
'    If Not F_CGC(txtbxCnpj.Text) Then
'        MsgBox "CNPJ Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
'        bCNPJ = False
'    Else
'        bCNPJ = True
'    End If
    P_FormataCGC txtbxCNPJ
End Sub

Private Sub txtbxContato_GotFocus()
    P_SelecTexto txtbxContato
End Sub

Private Sub txtbxContato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxContato_LostFocus()
    P_AjustaTexto txtbxContato
End Sub

Private Sub txtbxEmail_GotFocus()
    P_SelecTexto txtbxEmail
End Sub

Private Sub txtbxEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEmail_LostFocus()
    P_AjustaTexto txtbxEmail
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

Private Sub txtbxFan_GotFocus()
    P_SelecTexto txtbxFan
End Sub

Private Sub txtbxFan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFan_LostFocus()
    P_AjustaTexto txtbxFan
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
     txtbxFax.Text = F_AjustaTelefone(txtbxFax.Text)
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

