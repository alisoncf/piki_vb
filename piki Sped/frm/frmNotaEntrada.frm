VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmNotaEntrada 
   Caption         =   "Escolha uma Nota Fiscal"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "frmNotaEntrada.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   13560
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13530
      TabIndex        =   17
      Top             =   6720
      Width           =   13560
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferir"
         Height          =   690
         Left            =   3720
         Picture         =   "frmNotaEntrada.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnReabrir 
         Caption         =   "&Reabrir"
         Height          =   690
         Left            =   7320
         Picture         =   "frmNotaEntrada.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Efetivar"
         Height          =   690
         Left            =   5520
         Picture         =   "frmNotaEntrada.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   690
         Left            =   120
         Picture         =   "frmNotaEntrada.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   690
         Left            =   1920
         Picture         =   "frmNotaEntrada.frx":0A72
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   690
         Left            =   9120
         Picture         =   "frmNotaEntrada.frx":0BBC
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   690
         Left            =   11880
         Picture         =   "frmNotaEntrada.frx":0D06
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
      ScaleWidth      =   13530
      TabIndex        =   0
      Top             =   5985
      Width           =   13560
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11280
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmNotaEntrada.frx":0E71
         Left            =   3720
         List            =   "frmNotaEntrada.frx":0E84
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
         ItemData        =   "frmNotaEntrada.frx":0EF4
         Left            =   12360
         List            =   "frmNotaEntrada.frx":0F04
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         Picture         =   "frmNotaEntrada.frx":0F3B
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmbxProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9840
         Picture         =   "frmNotaEntrada.frx":12E6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
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
         Left            =   11280
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
         Left            =   10200
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
         Left            =   12360
         TabIndex        =   10
         Top             =   0
         Width           =   510
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmNotaEntrada.frx":1691
      Height          =   5535
      Left            =   0
      TabIndex        =   22
      Top             =   -120
      Width           =   12015
      _ExtentX        =   21193
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "num_not"
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
         DataField       =   "nome_for"
         Caption         =   "Fornecedor"
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
         DataField       =   "emissao_not"
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
      BeginProperty Column03 
         DataField       =   "data_not"
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
      BeginProperty Column04 
         DataField       =   "valor_not"
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
         DataField       =   "icms_not"
         Caption         =   "ICMS"
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
         DataField       =   "status_not"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2789,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   2760
      Top             =   5400
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
Attribute VB_Name = "frmNotaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




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
        
'        P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        'P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        If fr.operacao = "E" Then
            fr.Show 0
        Else
            fr.Show 0, mdiFrm
        End If
        
    End With
End Sub

Private Sub cmndbtnEfetivar_Click()
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

Private Sub cmndbtnNew_Click()
    Dim fr As New frmEntradaSped
    On Error Resume Next
    fr.operacao = "E"
    fr.Show 0
    fr.SetFocus

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
    
        If MsgBox("Confirma Transferência?", vbYesNo + vbQuestion) = vbYes Then
            transfere
        End If
    
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
    With dtNot.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "4" Then
            If MsgBox("Deseja realmente cancelar a nota fiscal e voltar todo o estoque dos produtos envolvidos?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
            End If
                P_cancelanotaFechada !seq_not, "R"
                busca
        Else
            If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !num_not & _
                " ?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "delete from tab_cmp where num_not = '" & !num_not & "'"
                strSQL = strSQL & " and cod_for='" & !cod_for & "' "
                strSQL = strSQL & " and serie_not ='" & !serie_not & "' "
                conn.Execute strSQL
                strSQL = "delete from tab_itcmp where seq_not = '" & !seq_not & "'"
            
            
                conn.Execute strSQL
                strSQL = "update tab_compras set status_not='2' where num_not = '" & !num_not & "'"
                strSQL = strSQL & " and cod_for='" & !cod_for & "' "
                strSQL = strSQL & " and serie_not ='" & !serie_not & "' "
                conn.Execute strSQL
                dtNot.Refresh
            End If
        End If
    End With
    
End Sub

Private Sub Form_Load()
        adc.ConnectionString = strConnect
End Sub
Sub busca()
    Dim strSQL As String
    If cmbxSitua.ListIndex = 0 Then
        strSQL = "select tc.*,tf.nome_for,tf.uf_for from tab_fornecedores as tf, "
        strSQL = strSQL & "tab_compras as tc "
        strSQL = strSQL & " where tc.cod_for=tf.cod_for "
    
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and tc.entrada_not < " & F_PesqData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and tc.entrada_not > " & F_PesqData(txtbxDataIni.Text)
        End If
    
        strSQL = strSQL & " and tc.status_not='2' "
    Else
        strSQL = "select tc.*,tf.nome_for,tf.uf_for from tab_fornecedores as tf, "
        strSQL = strSQL & "tab_cmp as tc "
        strSQL = strSQL & " where tc.cod_for=tf.cod_for "
    
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and tc.entrada_not < " & F_PesqData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and tc.entrada_not > " & F_PesqData(txtbxDataIni.Text)
        End If
        If cmbxSitua.ListIndex = 1 Then
            strSQL = strSQL & " and tc.status_not='1' "
        ElseIf cmbxSitua.ListIndex = 2 Then
            strSQL = strSQL & " and tc.status_not='2' "
        ElseIf cmbxSitua.ListIndex = 3 Then
            strSQL = strSQL & " and tc.status_not='3' "
        End If
    End If
    dtNot.RecordSource = strSQL
    dtNot.Refresh
    soma
    
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
    lblTOtal.caption = FormatVal(cTotal)
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
