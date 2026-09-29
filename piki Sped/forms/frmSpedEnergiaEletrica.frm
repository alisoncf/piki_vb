VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSpedEnergiaEletrica 
   Caption         =   "Escolha uma Nota Fiscal"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14475
   Icon            =   "frmSpedEnergiaEletrica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   14475
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14445
      TabIndex        =   17
      Top             =   6660
      Width           =   14475
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   690
         Left            =   120
         Picture         =   "frmSpedEnergiaEletrica.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   690
         Left            =   1920
         Picture         =   "frmSpedEnergiaEletrica.frx":07CC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   690
         Left            =   3720
         Picture         =   "frmSpedEnergiaEletrica.frx":0916
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   690
         Left            =   11880
         Picture         =   "frmSpedEnergiaEletrica.frx":0A60
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcBuscar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   14445
      TabIndex        =   0
      Top             =   5925
      Width           =   14475
      Begin VB.CommandButton cmdPesquisar 
         Appearance      =   0  'Flat
         Caption         =   "Pesquisar"
         Height          =   525
         Left            =   13200
         Picture         =   "frmSpedEnergiaEletrica.frx":0BCB
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10680
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmSpedEnergiaEletrica.frx":0F76
         Left            =   3720
         List            =   "frmSpedEnergiaEletrica.frx":0F89
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmSpedEnergiaEletrica.frx":0FF9
         Left            =   11760
         List            =   "frmSpedEnergiaEletrica.frx":1009
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         Picture         =   "frmSpedEnergiaEletrica.frx":1040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9240
         Picture         =   "frmSpedEnergiaEletrica.frx":13EB
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   5520
         TabIndex        =   16
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   3720
         TabIndex        =   15
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10680
         TabIndex        =   14
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9600
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   11760
         TabIndex        =   10
         Top             =   0
         Width           =   510
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frmSpedEnergiaEletrica.frx":1796
      Height          =   5535
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   9763
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
      Caption         =   "Entradas Encontradas"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "Id"
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
         DataField       =   "num_doc"
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
      BeginProperty Column02 
         DataField       =   "nome"
         Caption         =   "Participante"
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
         DataField       =   "dt_doc"
         Caption         =   "Emissão"
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
         DataField       =   "dt_e_s"
         Caption         =   "Data Entrada"
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
         DataField       =   "vl_doc"
         Caption         =   "Valor"
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
      BeginProperty Column06 
         DataField       =   "vl_icms"
         Caption         =   "ICMS"
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
      BeginProperty Column07 
         DataField       =   "status"
         Caption         =   "Status"
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
         DataField       =   "tipo"
         Caption         =   "Tipo"
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
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3690,142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSpedEnergiaEletrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub cmbxSitua_Click()
    busca
End Sub

Private Sub cmndbtnCancelar_Click()

End Sub

Private Sub cmndbtnEditar_Click()
    Dim fr As New frmEntradaSped
    fr.operacao = "I"
    If dtNot.Recordset.RecordCount <= 0 Or dbgrd.Row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With dtNot.Recordset
        If !status_not <> "1" Then
            MsgBox "A Nota Não está Aberta!", vbExclamation
            Exit Sub
        End If
        fr.operacao = "E"
        fr.tag = !seq_not
        fr.dtgrd.tag = fr.tag
        fr.caption = "Finalização de Nota Fiscal: " & !num_not & " de " & !nome_for
        
        fr.txtbxDesconto.Text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.Text = !num_not
        fr.txtbxNum_Not.tag = !num_not
        fr.txtbxSerie_Not.Text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.Text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.Text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.Text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.Text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.Text = !emissao_not
        fr.txtbxEntrada_Not.Text = !entrada_not
        fr.txtbxICMS_Not.Text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.Text = Format(!ipi_Not, "###,###,##0.00")
        fr.txtForn.Text = !nome_for
        fr.txtForn.tag = !cod_for
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.Text = FormatVal(!valor_not)
        
        'P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        'P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        If fr.operacao = "E" Then
            fr.Show 0
        Else
            fr.Show 0, mdiFrm
        End If
        
    End With
End Sub

Private Sub cmbxProd_Click()

End Sub

Private Sub cmdForn_Click()
    buscaForn
End Sub
Sub buscaForn()
    Dim f As fornecedor.fornecedor
    If f.cod_for <> "" Then
        txtForn.Text = f.nome_for
        txtForn.tag = f.cod_for
        busca
    End If
    
    
    
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnEdit_Click()
    With ado
        If !status <> "1" Then
            MsgBox "Esta nota não pode ser editada!", vbExclamation
            Exit Sub
        End If
        frmEntradaSped.tag = !id
        frmEntradaSped.Show 0, mdiFrm
        
    End With
    
        
    
    
End Sub

Private Sub cmndbtnEfetivar_Click1()
    Dim rPesq As New ADODB.Recordset
    Dim rForn As New ADODB.Recordset
    Dim strSQL As String
    Dim cCustoMed As Currency
    Dim cEst As Currency 'o estoque antigo
    Dim cCusto As Currency 'o custo antigo
    Dim sPro As String
    Dim sPreco As String
    Dim cNovoPreco As Currency
    'On Error Resume Next
    If MsgBox("Confirma Conclusão da Nota 'Fiscal'? ", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With dtNot.Recordset
        
        strSQL = "update tab_cmp set status_not = '3' "
        strSQL = strSQL & " where cod_for='" & !cod_for & "' "
        strSQL = strSQL & " and serie_not='" & !serie_not & "'"
        strSQL = strSQL & " and num_not='" & !num_not & "' "
        
        conn.Execute strSQL
        
            strSQL = "update tab_compras set status_not = '4' where "
        
        strSQL = strSQL & " seq_not ='" & !seq_not & "' "
        conn.Execute strSQL
        'itens
        strSQL = "select * from tab_itcompras "
        strSQL = strSQL & " where seq_not='" & F_BuscaUmCampo("seq_not", "tab_compras", "cod_for='" & !cod_for & "' and serie_not='" & !serie_not & "' and num_not='" & !num_not & "'") & "' "
        'strSQL = strSQL & " where seq_not='" & dtNot.Recordset!seq_not & "'"
        
        rPesq.Open strSQL, conn_client
        With rPesq
            Do While Not .EOF
                cEst = CCur((F_BuscaUmCampo("ef_Pro", "tab_produtos", " cod_pro='" & !cod_pro & "'")))
                cCusto = CCur((F_BuscaUmCampo("cf_pro ", "tab_produtos", " cod_pro='" & !cod_pro & "'")))
                
                
                
                strSQL = "update tab_produtos set ef_pro=ef_pro+" & F_GravaDec(!qtd_ite * !qtd_emb)
                strSQL = strSQL & ", cf_pro=" & F_GravaDec(!custo_ite)
                
                
                

                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                conn.Execute strSQL
                
                
                .MoveNext
            Loop
        
        End With
        
        'itens cmp
        strSQL = "select ti.*, tp.pv_pro,tp.nome_pro,tp.er_pro,tp.cr_pro from tab_itcmp as ti , tab_produtos as tp "
        strSQL = strSQL & " where ti.seq_not='" & F_BuscaUmCampo("seq_not", "tab_cmp", "cod_for='" & dtNot.Recordset!cod_for & "' and serie_not='" & dtNot.Recordset!serie_not & "' and num_not='" & dtNot.Recordset!num_not & "'") & "' "
        strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
        
        rPesq.Open strSQL, conn_client
    End With

       On Error GoTo trata_erro
        With rPesq
            Do While Not .EOF
                cEst = !er_pro
                cCusto = !cr_pro
                If (cEst + !qtd_ite * !qtd_emb) Then
                    cCustoMed = ((cEst * cCusto) + (!qtd_ite * !qtd_emb * !custo_ite)) / (cEst + !qtd_ite * !qtd_emb)
                End If
                strSQL = "update tab_produtos set er_pro=er_pro+" & F_GravaDec(!qtd_ite * !qtd_emb)
               
                
                    strSQL = strSQL & ", cr_pro=" & F_GravaDec(!custo_ite)
                
                
                
                
                cNovoPreco = Round(!precofuturo_ite, 2)
                
                 
                strSQL = strSQL & ", pv_pro=" & F_GravaDec(cNovoPreco)
                    
                
                

                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                conn.Execute strSQL
                P_AltPreco !cod_pro, cNovoPreco, !pv_pro
                .MoveNext
            Loop
            
        End With
        MsgBox "Nota Efetivada com Sucesso!", vbInformation
        dtNot.Refresh
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnEfetivar_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    If strEfetnota_usu <> "S" Then
        MsgBox "Você não tem permissão para esta ação!", vbExclamation
        Exit Sub
    End If
    
    
    With ado
        If !status = "2" Then
            EfetivaSpedNf !id
        Else
            MsgBox "Esta nota não está apta para ser fechada!"
        End If
    End With
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    frmSpedEnergiaEletricaIncluir.Show 1
    busca

End Sub

Private Sub cmndbtnNova_Click()
    Dim fr As New frmEntradaSped
    fr.operacao = "E"
    fr.Show 0, mdiFrm
End Sub

Private Sub cmndbtnReabrir_Click()
    Dim strSQL As String
    With dtNot.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If MsgBox("Deseja reabrir a Nota Fiscal?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update tab_compras set status_not='1' where "
            strSQL = strSQL & "num_not='" & !num_not & "' "
            strSQL = strSQL & " and cod_for='" & !cod_for & "' "
            strSQL = strSQL & " and serie_not = '" & !serie_not & "' "
            conn.Execute strSQL
            If banco.RecordsAffected = 0 Then
                MsgBox "Esta não é uma nota complementada!", vbExclamation
                Exit Sub
            End If
            dtNot.Refresh
        End If
    End With
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTransf_Click()
    
    frmSpedImportarXMLNF.Show 1
    
End Sub
Sub transfere()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sSeq As String
    With dtNot.Recordset
        Dim fr As New frmEntradaSped
        fr.operacao = "E"
        fr.tag = !seq_not
        fr.dtgrd.tag = fr.tag
        fr.caption = "Nota Fiscal: " & !num_not & " de " & !nome_for
        fr.txtbxDesconto.Text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.Text = !num_not
        fr.txtbxNum_Not.tag = !num_not
        fr.txtbxSerie_Not.Text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.Text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.Text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.Text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.Text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.Text = !emissao_not
        fr.txtbxEntrada_Not.Text = !entrada_not
        fr.txtbxICMS_Not.Text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.Text = Format(!ipi_Not, "###,###,##0.00")
        fr.txtForn.Text = !nome_for
        fr.txtForn.tag = !cod_for
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.Text = FormatVal(!valor_not)
        
        'P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        'P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        sSeq = MaxKeySQL("tab_cmp", "seq_not", "", "000000#")
        strSQL = "insert into tab_cmp(seq_not,Cod_For,Num_not,Serie_not,desconto_not,valorprod_not,Valor_not,Bc_Not,"
        strSQL = strSQL & "frete_not,icms_not,ipi_not,despesas_not,operacao_not,emissao_not,"
        strSQL = strSQL & "entrada_not,cod_usu,status_not,seq_ent)Values("
        strSQL = strSQL & "'" & sSeq & "',"
        strSQL = strSQL & "'" & !cod_for & "', "
        strSQL = strSQL & "'" & !num_not & "',"
        strSQL = strSQL & "'" & !serie_not & "',"
        strSQL = strSQL & F_GravaDec(!desconto_not) & ","
        strSQL = strSQL & F_GravaDec(!valorprod_not) & ","
        strSQL = strSQL & F_GravaDec(!valor_not) & ","
        strSQL = strSQL & F_GravaDec(!bc_not) & ","
        strSQL = strSQL & F_GravaDec(!frete_not) & ","
        strSQL = strSQL & F_GravaDec(!icms_not) & ","
        strSQL = strSQL & F_GravaDec(!ipi_Not) & ","
        strSQL = strSQL & F_GravaDec(!despesas_not) & ","
        strSQL = strSQL & "'" & (!operacao_not) & "',"
        strSQL = strSQL & F_GravaData(!emissao_not) & ","
        strSQL = strSQL & F_GravaData(!entrada_not) & ","
        strSQL = strSQL & "'" & strCod_Usu & "','1', " 'em aberto
        strSQL = strSQL & "'" & !seq_not & "') "
        conn.Execute strSQL

        
        strSQL = "select * from tab_itcompras where seq_not='" & !seq_not & "' "
        strSQL = strSQL & " order by seq_ite"
        rPesq.Open strSQL, conn_client
        With rPesq
            
            Do While Not .EOF
                
                strSQL = "Insert into tab_itcmp(seq_not,seq_ite,cod_pro,unit_ite,qtd_ite,qtd_emb,ipi_ite,"
                strSQL = strSQL & "gastos_ite,custo_ite,sujerido_ite,precofuturo_ite,lucrosuj_ite,"
                strSQL = strSQL & "lucroliq_ite,custosai_ite,icms_ite)values("
                strSQL = strSQL & "'" & sSeq & "',"
                strSQL = strSQL & !seq_ite & ","
                strSQL = strSQL & "'" & !cod_pro & "',"
                strSQL = strSQL & F_GravaDec(!unit_ite) & ","
                strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
                strSQL = strSQL & F_GravaDec(!qtd_emb) & ","
                strSQL = strSQL & F_GravaDec(!ipi_ite) & ","
                strSQL = strSQL & F_GravaDec(!gastos_ite) & ","
                strSQL = strSQL & F_GravaDec(!custo_ite) & ","
                strSQL = strSQL & F_GravaDec(!sujerido_ite) & ","
                strSQL = strSQL & F_GravaDec(!precofuturo_ite) & ","
                strSQL = strSQL & F_GravaDec(!lucrosuj_ite) & ","
                strSQL = strSQL & F_GravaDec(!lucroliq_ite) & ","
                strSQL = strSQL & F_GravaDec(!custosai_ite) & ","
                strSQL = strSQL & F_GravaDec(!icms_ite) & ")"
                conn.Execute strSQL
                
                .MoveNext
            Loop
            
        End With
        
    End With
    
     With dtNot.Recordset
        
        strSQL = "update tab_compras set status_not = '3' where seq_not='" & !seq_not & "' "
        
        conn.Execute strSQL
        
        fr.caption = "Finalização de NF: " & !num_not & " - " & !nome_for
        fr.tag = sSeq
        fr.Show 0
    End With
    
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If dtNot.Recordset.RecordCount = 0 Then Exit Sub
    With dtNot.Recordset
        If KeyCode = vbKeyP And Shift = 2 Then
        
            P_ImprNotaFiscal frmVisualiza.Pic, !num_not, _
         !cod_for, !serie_not, 0
         frmVisualiza.Show 1
            Unload frmVisualiza
        End If
        If KeyCode = vbKeyE And Shift = vbCtrlMask Then
            ''P_EspelhoNota !seq_Not
            
            
        End If
    End With
End Sub

Private Sub dtNot_Reposition()
    With dtNot.Recordset
            cmndbtnCancelar.Enabled = False
            cmndbtnEditar.Enabled = False
            cmndbtnEfetivar.Enabled = False
            cmndbtnReabrir.Enabled = False
            cmndbtnTransf.Enabled = False
        If .RecordCount = 0 Then
            
            
        Else
            If cmbxSitua.ListIndex = 0 Then
                cmndbtnCancelar.Enabled = False
                cmndbtnEditar.Enabled = False
                cmndbtnEfetivar.Enabled = True
                cmndbtnReabrir.Enabled = True
                cmndbtnTransf.Enabled = True
            ElseIf cmbxSitua.ListIndex = 1 Then
                cmndbtnCancelar.Enabled = True
                cmndbtnEditar.Enabled = True
                cmndbtnEfetivar.Enabled = False
                cmndbtnReabrir.Enabled = False
                cmndbtnTransf.Enabled = False
            ElseIf cmbxSitua.ListIndex = 2 Then
                cmndbtnCancelar.Enabled = True
                cmndbtnEditar.Enabled = False
                cmndbtnEfetivar.Enabled = True
                cmndbtnReabrir.Enabled = False
                cmndbtnTransf.Enabled = False
            ElseIf cmbxSitua.ListIndex = 3 Then
                cmndbtnCancelar.Enabled = False
                cmndbtnEditar.Enabled = False
                cmndbtnEfetivar.Enabled = False
                cmndbtnReabrir.Enabled = False
                cmndbtnTransf.Enabled = False
            
            End If
        End If
        
    End With
End Sub

Private Sub cmdCancel_Click()
        Dim strSQL As String
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status = "4" Then
            If MsgBox("Deseja realmente cancelar a nota fiscal e voltar todo o estoque dos produtos envolvidos?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
            End If
                P_cancelanotaFechada !seq_not, "R"
                busca
        Else
            If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !num_doc & _
                " ?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "delete from sped_nf where id = '" & !id & "'"
                
                conn.Execute strSQL
                strSQL = "delete from sped_nf_itens where id = '" & !id & "'"
                conn.Execute strSQL
            
                ado.Requery
            End If
        End If
    End With
    
End Sub



Private Sub Form_Load()
        Set ado = New ADODB.Recordset
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select tf.nome, "
    strSQL = strSQL & " te.* "
    strSQL = strSQL & " from sped_nf_energia as te left join sped_participante as tf  "
    strSQL = strSQL & " on tf.codigo=te.COD_PART "
    strSQL = strSQL & " where 1>0 "
    If cmbxStatus.Text <> "" Then
        strSQL = strSQL & " and status='" & Left(cmbxStatus.Text, 1) & "'"
    End If
        

    
    abreObjetoAdo ado, strSQL, grid
    
End Sub
Sub soma()
    Dim cTotal As Currency
    With dtNot.Recordset
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            cTotal = cTotal + !valor_not
            
            .MoveNext
        Loop
        
        
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    lblTotal.caption = FormatVal(cTotal)
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcBuscar
End Sub

Private Sub txtbxdataFin_GotFocus()
    GanhaFoco txtbxdataFin
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    GanhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
End Sub
