VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIncFor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inclusão de Fornecedores"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7485
   Icon            =   "frmIncFor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   5040
      Picture         =   "frmIncFor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   2640
      Picture         =   "frmIncFor.frx":0537
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Fornecedor:"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtObs 
         Height          =   525
         Left            =   1320
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   3120
         Width           =   5415
      End
      Begin VB.TextBox txtbxContato 
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   21
         Top             =   2760
         Width           =   5415
      End
      Begin VB.TextBox txtbxEmail 
         Height          =   285
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   19
         Top             =   2400
         Width           =   5415
      End
      Begin VB.TextBox txtbxFan 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox txtbxRazao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox txtbxCNPJ 
         Height          =   285
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   6
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtbxEndereco 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox txtbxIE 
         Height          =   285
         Left            =   4440
         MaxLength       =   16
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtbxBairro 
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   12
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtbxCidade 
         Height          =   285
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   16
         Top             =   2040
         Width           =   4695
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtbxCEP 
         Height          =   285
         Left            =   5760
         MaxLength       =   16
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblAno 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Obs:"
         Height          =   195
         Left            =   900
         TabIndex        =   22
         Top             =   3240
         Width           =   330
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co&ntato:"
         Height          =   195
         Left            =   630
         TabIndex        =   20
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emai&l:"
         Height          =   195
         Left            =   810
         TabIndex        =   18
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fan&tasia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   510
         TabIndex        =   3
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Razão Social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&NPJ:"
         Height          =   195
         Left            =   780
         TabIndex        =   5
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         Height          =   195
         Left            =   495
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&IE:"
         Height          =   195
         Left            =   4080
         TabIndex        =   7
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bairro:"
         Height          =   195
         Left            =   780
         TabIndex        =   11
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         Height          =   195
         Left            =   690
         TabIndex        =   15
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CE&P:"
         Height          =   195
         Left            =   5310
         TabIndex        =   13
         Top             =   1680
         Width           =   360
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Telefones:"
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   7095
      Begin VB.TextBox txtbxTel 
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxTipoTel 
         Height          =   315
         ItemData        =   "frmIncFor.frx":09DF
         Left            =   1320
         List            =   "frmIncFor.frx":09F2
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnAddTel 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnExclTel 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   1440
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrdTel 
         Height          =   1695
         Left            =   2880
         TabIndex        =   31
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2990
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº:"
         Height          =   195
         Left            =   840
         TabIndex        =   25
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmIncFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim bCNPJ As Boolean

Sub P_CarregaTelefone()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipo As String
    msflxgrdTel.rows = 1
    strSQL = "select * from tab_telefones where cod_cli='" & Me.tag & "' and tipo_ter='F' "
    abreObjetoAdo rPesq, strSQL
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

Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    Dim iCount As Integer
    Dim sTel As String
    Dim sTipo As String
    
    sCod = MaxKeySQL("tab_fornecedores", "cod_for", "", "0000#")
    
    
    
    
    strSQL = "insert into tab_fornecedores(cod_for,nome_for,fantasia_for,cnpj_for,ie_for,endereco_for, "
    strSQL = strSQL & "bairro_for,cidade_for,uf_for,cep_for,contato_for,email_for) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & "'" & txtbxFan.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "'" & txtbxIE.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "'" & txtbxBairro.Text & "',"
    strSQL = strSQL & "'" & txtbxCidade.Text & "',"
    strSQL = strSQL & "'" & cmbxUf.Text & "',"
    strSQL = strSQL & "'" & txtbxCEP.Text & "',"
    'strSQL = strSQL & "'" & txtbxTel.Text & "',"
    'strSQL = strSQL & "'" & txtbxFax.Text & "',"
    strSQL = strSQL & "'" & txtbxContato.Text & "',"
    strSQL = strSQL & "'" & txtbxEmail.Text & "')"
    
    Conn.Execute strSQL
    
    Me.tag = sCod
    P_gravatel
    
Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_gravatel()
    Dim iCount As Integer
    Dim sTel As String
    Dim sTipo As String
    strSQL = "delete from tab_telefones where cod_cli='" & Me.tag & "' "
    strSQL = strSQL & " and tipo_ter='F' "
    Conn.Execute strSQL
    For iCount = 1 To msflxgrdTel.rows - 1
        msflxgrdTel.row = iCount
        
        msflxgrdTel.col = 0: sTel = msflxgrdTel.Text
        msflxgrdTel.col = 1: sTipo = msflxgrdTel.Text
        If Not Clientes.F_GravaTelefone(Me.tag, sTel, sTipo, "F") Then
            MsgBox "Telefone não gravado!"
        End If
    Next
End Sub
Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_fornecedores set "
    strSQL = strSQL & "nome_for='" & txtbxRazao.Text & "',"
    strSQL = strSQL & "fantasia_for='" & txtbxFan.Text & "',"
    strSQL = strSQL & "cnpj_for='" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "ie_for='" & txtbxIE.Text & "',"
    strSQL = strSQL & "endereco_for='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "bairro_for='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "cidade_for='" & txtbxCidade.Text & "',"
    strSQL = strSQL & "uf_for='" & cmbxUf.Text & "',"
    strSQL = strSQL & "cep_for='" & txtbxCEP.Text & "',"
    strSQL = strSQL & "telefone_for='" & txtbxTel.Text & "', "
    'strSQL = strSQL & "fax_for='" & txtbxfax.Text & "', "
    strSQL = strSQL & "contato_for='" & txtbxContato.Text & "', "
    strSQL = strSQL & "email_for='" & txtbxEmail.Text & "'"
    strSQL = strSQL & " where cod_for = '" & Me.tag & "'"
    Conn.Execute strSQL
    P_gravatel
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
    'ElseIf txtbxCNPJ.Text = "" Then
    '    MsgBox "Informe o CNPJ do Fornecedor!", vbExclamation
    '    Exit Sub
    ElseIf bCNPJ = False Then
        MsgBox "Informe um CNPJ Válido para Fornecedor!", vbExclamation
        Exit Sub
    ElseIf txtbxEndereco.Text = "" Then
'        MsgBox "Informe o endereço do Fornecedor!", vbExclamation
'        Exit Sub
    ElseIf txtbxBairro.Text = "" Then
'        MsgBox "informe o Bairro do Fornecedor!,vbExclamation  "
'        Exit Sub
    ElseIf txtbxCidade.Text = "" Then
 '       MsgBox "Informe a Cidade do Fornecedor!", vbExclamation
 '       Exit Sub
    ElseIf cmbxUf.Text = "" Then
 '       MsgBox "Informe o Estado de Origem do Fornecedor!", vbExclamation
 '       Exit Sub
    End If
    If Me.tag = "" Then
        If F_BuscaUmCampo("cod_For", "tab_fornecedores", "cnpj_for='" & txtbxCNPJ.Text & "'") <> "" And txtbxCNPJ.Text <> "" Then
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
        'txtbxFax.Text = ""
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
     On Error GoTo trata_Erro
     Dim strSQL As String
     Dim rPesq As New ADODB.Recordset
     If bActive Then Exit Sub
     If Me.tag <> "" Then
        strSQL = "select * from tab_FORNECEDORES where cod_FOR = '" & Me.tag & "' "
        abreObjetoAdo rPesq, strSQL
        With rPesq
           
                txtbxBairro.Text = !bairro_for
                txtbxCEP.Text = !cep_for
                txtbxCidade.Text = !cidade_For
                txtbxCNPJ.Text = !cnpj_for
                txtbxEndereco.Text = !endereco_for
                txtbxIE.Text = !ie_for
                txtbxRazao.Text = !nome_for
                
                txtbxContato.Text = !contato_for
                txtbxEmail.Text = !email_for
                txtbxFan.Text = !fantasia_for
                
                cmbxUf.Text = !uf_for
                If Not IsNull(!obs_for) Then
                    txtObs.Text = !obs_for
                End If
                P_CarregaTelefone
                
           .Close
        End With
    End If
    bActive = True
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Form_Load()
    bActive = False
    P_PreencheUF cmbxUf
   ' Me.Height = 5520
   ' Me.Width = 5385
    P_CentraForm Me
    If Me.tag = "" Then
        bCNPJ = True
    End If
   With msflxgrdTel
        .FixedCols = 0
        .rows = 1
        .Cols = 2
        .col = 0: .Text = "Fone:"
        .col = 1: .Text = "Tipo:"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1200
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
End Sub
Private Sub cmndbtnAddTel_Click()
    If F_Vazio(txtbxTel.Text) Then
        MsgBox "Infome o nº do Telefone!", vbExclamation
        txtbxTel.SetFocus
        Exit Sub
    End If
    If cmbxTipoTel.Text = "" Then
        MsgBox "Informe o tipo do Telefone!", vbExclamation
        cmbxTipoTel.SetFocus
        Exit Sub
    End If
    If F_InFlxgrd(msflxgrdTel, txtbxTel.Text, 0) > -1 Then
        MsgBox "Este telefone já se encontra Adicionado!", vbExclamation
        txtbxTel.SetFocus
        Exit Sub
    End If
    If cmndbtnAddTel.tag = "" Then
        msflxgrdTel.AddItem txtbxTel.Text & Chr(9) & cmbxTipoTel.Text
    Else
        msflxgrdTel.col = 0: msflxgrdTel.Text = txtbxTel.Text
        msflxgrdTel.col = 1: msflxgrdTel.Text = cmbxTipoTel.Text
    End If
    cmndbtnAddTel.tag = ""
    txtbxTel.SetFocus
    
    
    
End Sub

Private Sub cmndbtnEditTel_Click()
    With msflxgrdTel
        If .row <= 0 Then
            Exit Sub
        End If
        .col = 0
        txtbxTel.Text = .Text
        .col = 1
        P_LocalizaOperacao .Text, cmbxTipoTel
        cmndbtnAddTel.tag = .row
    End With
End Sub

Private Sub cmndbtnExclTel_Click()
    With msflxgrdTel
        If .row <= 0 Then
            Exit Sub
        End If
        If MsgBox("Deseja excluir o telefone?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        If .rows = 2 Then
            .rows = 1
        Else
            .RemoveItem .row
        End If
    End With
End Sub

Private Sub Form_Resize()
'    P_CentraObejct Frame1, Me, False
'    P_CentraObejct Frame2, Me, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad(frmEdiFor.Name) Then
        
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
    txtbxCEP.Text = F_SoNumero(txtbxCEP.Text)
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
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFan_LostFocus()
    P_AjustaTexto txtbxFan
End Sub

Private Sub txtbxFax_GotFocus()
    'P_SelecTexto txtbxFax
End Sub

Private Sub txtbxFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFax_LostFocus()
     txtbxfax.Text = F_AjustaTelefone(txtbxfax.Text)
End Sub

Private Sub txtbxIE_GotFocus()
    P_SelecTexto txtbxIE
End Sub

Private Sub txtbxIE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
End Sub

