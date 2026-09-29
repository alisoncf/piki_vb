VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiDEv 
   Caption         =   "Devoluções realizadas"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16815
   Icon            =   "frmEdiDev.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   16815
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   16785
      TabIndex        =   15
      Top             =   7155
      Width           =   16815
      Begin VB.CommandButton cmdMotivo 
         Caption         =   "Motivo"
         Height          =   735
         Left            =   7080
         Picture         =   "frmEdiDev.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Imprimir"
         Height          =   705
         Left            =   11040
         Picture         =   "frmEdiDev.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdImprDev 
         Caption         =   "Imprimir Devolução"
         CausesValidation=   0   'False
         Height          =   705
         Left            =   12960
         Picture         =   "frmEdiDev.frx":0A7A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   14880
         Picture         =   "frmEdiDev.frx":0F68
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   9120
         Picture         =   "frmEdiDev.frx":14B6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblSoma 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   600
         TabIndex        =   21
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Soma"
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   120
         Width           =   405
      End
   End
   Begin MSDataGridLib.DataGrid gridDev 
      Height          =   6015
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   10610
      _Version        =   393216
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
      Caption         =   "Devoluções"
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "seq_dev"
         Caption         =   "Nº Dev."
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
         DataField       =   "num_ven"
         Caption         =   "Nº Venda"
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
      BeginProperty Column03 
         DataField       =   "data_dev"
         Caption         =   "Data Devolução"
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
         DataField       =   "motivo"
         Caption         =   "Motivo"
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
         DataField       =   "cod_pro"
         Caption         =   "código"
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
         DataField       =   "codfab_pro"
         Caption         =   "Cód.Fab"
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
         DataField       =   "desc_ite"
         Caption         =   "Produto"
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
         DataField       =   "qtd_ite"
         Caption         =   "Qtd Saída"
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
      BeginProperty Column09 
         DataField       =   "devolvido"
         Caption         =   "Qtd Devolvida"
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
      BeginProperty Column10 
         DataField       =   "total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "guerra_fun"
         Caption         =   "Vendedor"
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
      BeginProperty Column12 
         DataField       =   "nome_mar"
         Caption         =   "Marca"
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
         BeginProperty Column00 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3135,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2684,977
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3284,788
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1934,929
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   3105,071
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   16785
      TabIndex        =   0
      Top             =   0
      Width           =   16815
      Begin VB.TextBox txtVendedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         MaxLength       =   40
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Pesquisar"
         Height          =   705
         Left            =   10560
         Picture         =   "frmEdiDev.frx":19A7
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdCliente 
         Height          =   285
         Left            =   6360
         Picture         =   "frmEdiDev.frx":1AA0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdVendedor 
         Height          =   285
         Left            =   9360
         Picture         =   "frmEdiDev.frx":1DF0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         Height          =   195
         Left            =   6720
         TabIndex        =   11
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código do Produto"
         Height          =   195
         Left            =   2280
         TabIndex        =   6
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Left            =   3720
         TabIndex        =   8
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   285
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmEdiDEv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmbxBusca_Click()
    P_Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
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
   P_Busca
End Sub

Private Sub cmndbtnEdit_Click()
    
    P_Status "Montando venda..."
    With dtVendas.Recordset
        frmDevolucao.tag = !Num_DEV
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", tab_funcionarios, _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        'frmDevolucao.lblValor.caption = FormatVal(!valor_dev)
        'frmDevolucao.caption = "DEVOLUÇÃO " & !Num_DEV & "-" & !nome_Cli
        'frmDevolucao.Show 1
    End With
    P_Status ""
    dtVendas.Refresh
End Sub

Private Sub cmndbtnNew_Click()
       Dim strSQL As String
    
    strSQL = F_GeraNovaDevolucao("00000", strCod_Usu)
    
    
    
    frmDevolucoes.caption = "STYLO PRÓPRIO -Devolução a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmDevolucoes.tag = strSQL
    On Error Resume Next
    frmDevolucoes.Show 1
End Sub

Private Sub cmdCliente_Click()
    buscaCli
End Sub
Sub buscaCli()
    Dim c As cliente
    
    c = F_BuscaCliente(txtCliente.Text)
    If c.cod_cli <> "" Then
        txtCliente.Text = c.nome_Cli
        txtCliente.tag = c.cod_cli
        busca
    End If
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV gridDev, ado
End Sub

Private Sub cmdImprDev_Click()
    On Error GoTo trata_Erro
    ImprimeDevolucao ado!seq_dev
    
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmdMotivo_Click()
    strSQL = InputBox("Informe um novo motivo pra devolução", "Motivo da Devolução", ado!motivo)
    
    If strSQL <> ado!motivo Then
        If MsgBox("Confirma alteração do motivo para " & strSQL & "?", vbYesNo + vbQuestion) = vbYes Then
            Devolucao.AlteraMotivo ado!seq_dev, CStr(strSQL)
            ado.Requery
        End If
        
    End If
    
End Sub

Private Sub cmdPrint_Click()
    P_DtgrdToHtml "Devoluções " & txtbxDe.Text & " até " & txtbxAte.Text, gridDev, ado
End Sub

Private Sub cmdVendedor_Click()
    Dim f As funcionario.funcionario
    
    f = f_BuscaFun(txtVendedor.Text)
    If f.cod_fun <> "" Then
        txtVendedor.Text = f.guerra_fun
        txtVendedor.tag = f.cod_fun
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    
    busca
End Sub
Sub busca()
    Dim strSQL As String
    P_Status "carregando..."
    strSQL = "select td.*,guerra_fun,abrevenda_usu,nomecli_ven,tm.nome_mar,dataefet_ven,(qtd_dev) as devolvido,(total_dev) as total,tp.codfab_pro,tv.data_ven,dataefet_ven,"
    strSQL = strSQL & "tc.cnpj_cli,tc.cpf_cli,td.*,ti.* from " & TAB_VENDAS & " as tv "
    strSQL = strSQL & " inner join " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " inner join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar"
    strSQL = strSQL & " inner join devolucoes as td on td.num_ven=tv.num_ven "
    strSQL = strSQL & " inner join itdev as tid on td.seq_dev=tid.seq_dev "
    strSQL = strSQL & " inner join " & TAB_FUNCIONARIOS & " as tf on tf.cod_fun=tv.abrevenda_usu"
    
    strSQL = strSQL & " where 1 > 0 "
    
    If txtProd.Text <> "" Then
        strSQL = strSQL & " and (codfab_pro='" & txtProd.Text & "' or tp.cod_pro='" & txtProd.Text & "')"
    End If
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and data_dev>=" & F_GravaData(txtbxDe.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and data_dev<=" & F_GravaData(txtbxAte.Text)
    End If
    If txtCliente.tag <> "" Then
        strSQL = strSQL & "  and tc.cod_cli='" & txtCliente.tag & "' "
    End If
    If txtVendedor.tag <> "" Then
        strSQL = strSQL & " AND abrevenda_usu='" & txtVendedor.tag & "' "
    End If
    strSQL = strSQL & " and tid.cod_pro=ti.cod_pro "
    'strSQL = strSQL & " group by td.seq_dev "
       
    
    abreObjetoAdo ado, strSQL, gridDev
    soma
    P_Status ""
End Sub
Sub soma()
    Dim soma As Currency
    lblSoma.caption = "0,00"
    gridDev.Visible = False
    
    With ado
        Do While Not .EOF
            soma = soma + !total
            .MoveNext
        Loop
        lblSoma = FormatVal(soma)
        If .RecordCount > 0 Then .MoveFirst
    End With
    gridDev.Visible = True
End Sub
Private Sub cmndbtnSair_Click()

    Unload Me
End Sub

Private Sub DBGrid1_DblClick()
    If cmndbtnEdit.Enabled = True Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DBGrid1_DblClick
    End If
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub dtVendas_Reposition()
    On Error Resume Next
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        cmndbtnImpr.Enabled = True
        cmndbtnEdit.Enabled = True
    End If
    If dtVendas.Recordset!situacao_DEV > 1 Then
        cmndbtnEdit.Enabled = False
        
    End If
End Sub

Private Sub DataGrid1_Click()

End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    
    'MudaAparencia Me
    
    busca
End Sub

Private Sub Form_Load()
    
    
    Set ado = New ADODB.Recordset
    
    txtbxDe.Text = Date
    txtbxAte.Text = Date
    
    
    bActive = False
    resizeForm Me
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    
    busca
End Sub


Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub


Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        P_Busca
        If dtVendas.Recordset.RecordCount > 0 Then
            DBGrid1.SetFocus
        End If
    End If
    If KeyCode = vbKeyDown Then
        If dtVendas.Recordset.RecordCount > 0 Then
            DBGrid1.SetFocus
        End If
    End If
End Sub



Private Sub Form_Resize()
    PosicionaGrade gridDev, Me, pcBot
    posicionaADireita cmdSair, pcBot
End Sub

Private Sub gridDev_DblClick()
    cmdMotivo_Click
End Sub

Private Sub gridDev_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        gridDev_DblClick
    End If
    
End Sub

Private Sub txtbxAte_GotFocus()
        P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub

Private Sub txtbxVendedor_Change()

End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtCliente_GotFocus()
    P_SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCliente_Click
    End If
End Sub

Private Sub txtProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub

Private Sub txtVendedor_Change()
    txtVendedor.tag = ""
End Sub

Private Sub txtVendedor_GotFocus()
    P_SelecTexto txtVendedor
End Sub

Private Sub txtVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdVendedor_Click
    End If
End Sub
