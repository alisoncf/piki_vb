VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVendas 
   Caption         =   "Edição de Vendas"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13020
   Icon            =   "frmEdiVendas2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   13020
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   2055
      TabIndex        =   22
      Top             =   855
      Width           =   2055
      Begin VB.CommandButton cmdRomaneio 
         Caption         =   "&Romaneio Entrega"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Reabrir Venda"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":053E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir (A4)"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":1208
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":1592
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5520
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":17DC
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Nova"
         Enabled         =   0   'False
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":1926
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnVendasEfet 
         Caption         =   "&Vendas a Efetivar"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":1CB0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSepd 
         Caption         =   "Gerar &NFe"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":1DFA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar a Venda"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnReem 
         Caption         =   "&Reemitir CUPOM"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas2.frx":2246
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3720
         Width           =   1935
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12990
      TabIndex        =   21
      Top             =   0
      Width           =   13020
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         ItemData        =   "frmEdiVendas2.frx":2390
         Left            =   9360
         List            =   "frmEdiVendas2.frx":23A0
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbxSituacao_ven 
         Height          =   315
         ItemData        =   "frmEdiVendas2.frx":23CB
         Left            =   3480
         List            =   "frmEdiVendas2.frx":23DB
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiVendas2.frx":2418
         Left            =   1800
         List            =   "frmEdiVendas2.frx":2428
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   40
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiVendas2.frx":2453
         Left            =   120
         List            =   "frmEdiVendas2.frx":2463
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Pesquisar"
         Height          =   615
         Left            =   11040
         Picture         =   "frmEdiVendas2.frx":248E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         MaxLength       =   40
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   9360
         TabIndex        =   26
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação da Venda:"
         Height          =   195
         Left            =   3480
         TabIndex        =   4
         Top             =   120
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   7200
         TabIndex        =   8
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   8280
         TabIndex        =   10
         Top             =   120
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiVendas2.frx":2587
      Height          =   4695
      Left            =   2160
      TabIndex        =   24
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Vendas Localizadas"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "num_ven"
         Caption         =   "Número"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nomecli_ven"
         Caption         =   "Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "data_ven"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "dataefet_ven"
         Caption         =   "Efetivação"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "valor_ven"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "situacao_ven"
         Caption         =   "Situação"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "guerra_fun"
         Caption         =   "Funcionário"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "nome"
         Caption         =   "Fechado por:"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "desconto_ven"
         Caption         =   "Desconto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2369,764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   120
      Top             =   5760
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEdiVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxBusca_Click()
    P_Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxCaixa_Click()
    If Not bActive Then Exit Sub
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrd_Click()
    P_Busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxSituacao_ven_Click()
   cmndbtnBuscar_Click
End Sub

Private Sub cmdRomaneio_Click()
'    On Error GoTo trata_Erro
    Vendas.ImprimeComprovanteNaoFiscal adc.Recordset!num_ven
    Exit Sub
trata_Erro:
    MostraErro

End Sub

Private Sub cmndbtnBuscar_Click()
    P_Busca
End Sub

Private Sub cmndbtnCancel_Click()
    If strExclVenda_usu = "S" Then
        If MsgBox("Deseja realmente cancelar todos os lancamentos dessa venda?", vbYesNo + vbQuestion) = vbYes Then
            
            
            If adc.Recordset!situacao_ven < "1" Then
                MsgBox "A venda não pode ser cancelada por aqui!", vbCritical
            End If
            If Not F_CancelaVendaAposCupom(adc.Recordset!num_ven) Then
                MsgBox "Venda cancelada!", vbExclamation
            End If
        End If
    Else
        MsgBox "Você não tem permissão para isso!", vbCritical
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmVenda
    P_Status "Montando venda..."
    With adc.Recordset
        fr.tag = !num_ven
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        fr.Atacado = True
        fr.lblValor.caption = FormatVal(!valor_ven)
        fr.caption = "Venda " & !num_ven & "-" & !nome_Cli
        fr.txtbxNomeCli.Text = Format(!nomecli_ven)
        fr.txtbxNomeCli.tag = Format(!cod_cli, "0000#")
        fr.txtbxCod_Cli.Text = Format(!cod_cli)
        fr.operacao = "B"
        fr.pcECF.BackColor = vbYellow
        'frmVenda.txtbxObs.Text = Format(!obs_ven)
        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
        fr.Show 0
        fr.lblAtacado.tag = "X" 'indica que a venda foi editada
    End With
    P_Status ""
    adc.Refresh
End Sub

Private Sub cmndbtnImpr_Click()
    P_Status "Gerando impressão... "
    P_ImprimeVenda frmVisualiza.Pic, adc.Recordset!num_ven, False
    P_Status ""
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        
        P_ImprimeVenda Printer, adc.Recordset!num_ven, True
               
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
   Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = "Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    frmVenda.Show 0, mdiFrm
End Sub

Private Sub cmndbtnReem_Click()
    frmReemissao.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSepd_Click()
    On Error GoTo trata_Erro
    P_notaSEPD adc.Recordset!num_ven, ""
    Exit Sub
trata_Erro:
    MostraErro
End Sub

Private Sub cmndbtnVendasEfet_Click()
    frmEfetivaVen.Show 0
End Sub

Private Sub Command1_Click()
    If strExclVenda_usu = "S" Then
        If MsgBox("Deseja realmente cancelar todos os lancamentos dessa venda?", vbYesNo + vbQuestion) = vbYes Then
            
            
            If adc.Recordset!situacao_ven <> "2" Then
                MsgBox "A venda não pode ser reaberta!", vbCritical
            End If
            strSQL = "update " & TAB_VENDAS & " set situacao_ven='1' where num_ven='" & adc.Recordset!num_ven & "' "
            conn.Execute strSQL
            adc.Refresh
        End If
    Else
        MsgBox "Você não tem permissão para isso!", vbCritical
    End If
End Sub

Private Sub dtgrd_DblClick()
    If cmndbtnEdit.Enabled = True Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    On Error Resume Next
    If adc.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        cmndbtnImpr.Enabled = True
        cmndbtnEdit.Enabled = True
    End If
    If adc.Recordset!situacao_ven > 1 Then
        cmndbtnEdit.Enabled = False
        
    End If
End Sub

Private Sub Form_Activate()
    
    
    'MudaAparencia Me
    bActive = True
    
End Sub

Private Sub Form_Load()
    'Me.Width = 10725
    'Me.Height = 6135
    resizeForm Me
    adc.ConnectionString = banco.Connect
    P_PreencheCaixas cmbxCaixa
    P_LocalizaOperacao strCod_Cai, cmbxCaixa
    
    If strExclcai_usu <> "S" Then
        cmbxCaixa.Enabled = False
    End If
    txtbxAte.Text = Date
    txtbxDe.Text = Date
    
    cmbxSituacao_ven.ListIndex = 0
    If strAltprod_usu <> "S" Then
        cmbxSituacao_ven.Enabled = False
    End If
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 1
    
    P_Busca
    bActive = False
    
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    
    If cmbxBusca.ListIndex = 0 Then
        sBusca = "Nome_Cli"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sBusca = "num_ven"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sBusca = "data_ven"
    Else
        sBusca = "data_ven"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "Nome_Cli"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "num_ven"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "data_ven"
    Else
        sOrd = "dataefet_ven"
    End If
    
    
    strSQL = "select tc.nome_cli,tv.*,tf.guerra_fun,tf2.guerra_fun as nome "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " " & TAB_VENDAS & " as tv,tab_funcionarios as tf, tab_funcionarios as tf2 "
    If sBusca <> "dataefet_ven" And sBusca <> "data_ven" Then
        strSQL = strSQL & " where " & sBusca & " like '" & txtbxExpr.Text & "%' "
    Else
        strSQL = strSQL & " where " & sBusca & " >= " & F_GravaData(txtbxExpr.Text) & " "
    End If
    strSQL = strSQL & " and tf.cod_fun=tv.cod_fun "
    strSQL = strSQL & " and tf2.cod_fun=tv.cod_usu "
    strSQL = strSQL & " and tv.situacao_ven='" & Left(cmbxSituacao_ven.Text, 1) & "' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    
    
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and tv.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
      
    End If
    
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(txtbxDe.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(txtbxAte.Text & " 23:59:59")
    End If
        
  
        
    strSQL = strSQL & " order by " & sOrd & " desc "
    adc.RecordSource = strSQL
    adc.Refresh
End Sub


Private Sub Form_Resize()
    dtgrd.Width = Me.Width - dtgrd.Left - 100
    dtgrd.Height = Me.Height - dtgrd.Top - 500
End Sub

Private Sub txtbxAte_GotFocus()
    ganhaFoco txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxDe_GotFocus()
    ganhaFoco txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub

Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub

Private Sub txtbxExpr_GotFocus()
    ganhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        P_Busca
        If adc.Recordset.RecordCount > 0 Then
            dtgrd.SetFocus
        End If
    End If
    If KeyCode = vbKeyDown Then
        If adc.Recordset.RecordCount > 0 Then
            dtgrd.SetFocus
        End If
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    AjustaTexto txtbxExpr
End Sub
