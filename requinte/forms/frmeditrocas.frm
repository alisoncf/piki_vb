VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiTrocas 
   Caption         =   "Trocas Efetuadas"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16890
   Icon            =   "frmeditrocas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   16890
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   16860
      TabIndex        =   16
      Top             =   5760
      Width           =   16890
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir Troca"
         Height          =   615
         Left            =   13800
         Picture         =   "frmeditrocas.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPvenda 
         Caption         =   "Imprimir Venda"
         Height          =   615
         Left            =   12240
         Picture         =   "frmeditrocas.frx":10D7
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   4680
         Picture         =   "frmeditrocas.frx":14E4
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         TabIndex        =   17
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   15480
         Picture         =   "frmeditrocas.frx":186E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "Imprimir Grade"
         Height          =   615
         Left            =   10680
         Picture         =   "frmeditrocas.frx":19D9
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Visualizar"
         Height          =   615
         Left            =   9120
         Picture         =   "frmeditrocas.frx":1DE6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   375
         Width           =   2055
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmeditrocas.frx":1F30
         Left            =   120
         List            =   "frmeditrocas.frx":1F43
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   375
         Width           =   975
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   7
         Top             =   375
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   7440
         Picture         =   "frmeditrocas.frx":1F70
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5880
         TabIndex        =   8
         Top             =   120
         Width           =   1215
         Begin VB.OptionButton optnVenda 
            Caption         =   "Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optnTroca 
            Caption         =   "Troca"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdVendedor 
         Height          =   315
         Left            =   7680
         Picture         =   "frmeditrocas.frx":22FA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5040
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         Caption         =   "&Expressão:"
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&De:"
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   4680
         TabIndex        =   6
         Top             =   120
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmeditrocas.frx":2684
      Height          =   5775
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "seq_tro"
         Caption         =   "Número Troca"
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
         DataField       =   "data_tro"
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
         DataField       =   "num_ven"
         Caption         =   "Venda"
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
         DataField       =   "dataefet_ven"
         Caption         =   "Data Venda"
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
         DataField       =   "abat_tro"
         Caption         =   "vlr abatido"
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
         DataField       =   "acres_tro"
         Caption         =   "Vlr acrescido"
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
         DataField       =   "final"
         Caption         =   "Diferença"
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
      BeginProperty Column08 
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
      BeginProperty Column09 
         DataField       =   "situacao_ven"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3660,095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1530,142
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiTrocas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Private Sub cmbxBusca_Click()
    If cmbxBusca.ListIndex = 0 Then
        lblExpr.caption = "&Venda:"
    ElseIf cmbxBusca.ListIndex = 1 Then
        lblExpr.caption = "&Troca:"
    ElseIf cmbxBusca.ListIndex = 2 Then
        lblExpr.caption = "&Cliente:"
    ElseIf cmbxBusca.ListIndex = 3 Then
        lblExpr.caption = "&Produto:"
    End If
End Sub

Private Sub cmdPvenda_Click()
  On Error GoTo trata_Erro
  
   ImprVenda ado!num_ven
   Exit Sub
trata_Erro:
   
   
End Sub

Private Sub cmdVendedor_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectVendedor(txtVendedor.Text)
    If f.cod_fun <> "" Then
        txtVendedor.Text = f.guerra_fun
        txtVendedor.tag = f.cod_fun
        cmbxBusca_Click
    End If
End Sub
Private Sub cmndbtnLocaliza_Click()
    buscaProduto
End Sub
Sub buscaProduto()
    Dim p As Produto
    p = F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        busca
        
        
        
        
    End If
    
End Sub
Private Sub cmndbtnBuscar_Click()
    
    busca
    
End Sub
Sub busca()
    Dim strSQL As String
    On Error GoTo trata_Erro
    P_Msg "Aguarde..."
    strSQL = "select distinct tt.*,situacao_ven,tv.num_ven,tv.dataefet_ven,tv.nomecli_ven, guerra_fun,"
    strSQL = strSQL & "tv.valor_ven, (tt.acres_tro - tt.abat_tro) as final "
    strSQL = strSQL & "from tab_vendas as tv left join tab_itvendas as ti on tv.num_ven=ti.num_ven, "
    strSQL = strSQL & "tab_trocas as tt   "
    strSQL = strSQL & " left join tab_funcionarios as tu on tt.cod_usu=tu.cod_fun "
    
    strSQL = strSQL & " where tt.num_ven=tv.num_ven "
    'strSQL = strSQL & " and tv.num_ven=ti.num_ven "
    
    If optnVenda.Value Then
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & "and tv.dataefet_ven >= " & F_GravaData(txtbxDataIni.Text)
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & "and tv.dataefet_ven <= " & F_GravaData(txtbxdataFin.Text)
        End If
    Else
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & "and tt.data_tro >= " & F_GravaDataTempo(txtbxDataIni.Text)
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & "and tt.data_tro <= " & F_GravaDataTempo(txtbxdataFin.Text & " 23:59:59")
        End If
    End If
    If txtVendedor.tag <> "" Then
        strSQL = strSQL & " and tt.cod_usu='" & txtVendedor.tag & "' "
    End If
    If txtProduto.tag <> "" Then
        strSQL = strSQL & " and ti.cod_pro='" & txtProduto.tag & "' "
    Else
        If txtProduto.Text <> "" Then
            strSQL = strSQL & " and desc_ite like '" & txtProduto.Text & "%' "
        End If
    End If
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " and tv.num_ven like '" & txtbxExpr.Text & "%' "
        strSQL = strSQL & " order by tv.num_ven,seq_tro "
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " and tt.seq_tro like '" & txtbxExpr.Text & "%' "
        strSQL = strSQL & " order by seq_tro,tv.num_ven"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " and tv.nomecli_ven like ='%" & txtbxExpr.tag & "%' "
        strSQL = strSQL & " order by nomecli_ven,tv.num_ven,seq_tro "
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & " and ti.cod_pro ='" & txtbxExpr.tag & "' "
        strSQL = strSQL & " order by seq_tro,tv.num_ven"
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & " and tt.cod_usu ='" & txtbxExpr.tag & "' "
        strSQL = strSQL & " order by seq_tro,tv.num_ven"
    End If
    abreObjetoAdo ado, strSQL, dtg
    'soma
    P_Msg
trata_Erro:
    P_MostraErro
End Sub
Sub soma()
    Dim cTotalTro As Currency
    Dim cTotalVen As Currency
    With ado
        Do While Not .EOF
            cTotalTro = cTotalTro + !final
            cTotalVen = cTotalVen + !valor_ven
            .MoveNext
        Loop
        
    End With
    lblTro.caption = FormatVal(cTotalTro)
    lblVen.caption = FormatVal(cTotalVen)
    lblList.caption = ado.RecordCount
End Sub

Private Sub cmndbtnImpr_Click()
    P_DbgrdToHtml "Trocas Efetuadas", dtg, ado, "De: " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
    
    
    
    
    
    
    
    Dim fr As New frmVisualizaTro
    On Error Resume Next
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        fr.caption = "Troca: " & !seq_tro & " - Venda: " & !num_ven & _
        " - Cliente: " & !nomecli_ven
        fr.tag = ado!seq_tro
        fr.lblAbat.caption = FormatVal(!abat_tro)
        fr.lblAcr.caption = FormatVal(!acres_tro)
        fr.lblTotal.caption = FormatVal(!acres_tro - !abat_tro)
        fr.Show 0
        
    End With
End Sub

Private Sub Command2_Click()
   ImprimeTroca ado!seq_tro
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    resizeForm Me
    Set ado = New ADODB.Recordset
    cmbxBusca.ListIndex = 0
    
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtg, Me, pcBot
    posicionaADireita cmndbtnSair, pcBot
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub
Sub buscaCli()
    
    
    
    
    
End Sub
Sub buscaProd()
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxExpr.Text)
    txtbxExpr.Text = p.nome_pro
    txtbxExpr.tag = p.cod_pro
    
End Sub
Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        If cmbxBusca.ListIndex = 2 Then
             
        ElseIf cmbxBusca.ListIndex = 3 Then
            buscaProd
            
        End If
        busca
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub
