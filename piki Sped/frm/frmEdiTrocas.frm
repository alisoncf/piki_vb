VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiTrocas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trocas Efetuadas"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   Icon            =   "frmEdiTrocas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8955
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   7320
      ScaleHeight     =   1065
      ScaleWidth      =   1545
      TabIndex        =   14
      Top             =   4680
      Width           =   1575
      Begin VB.Label lblVen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   25
         Top             =   600
         Width           =   120
      End
      Begin VB.Label lblTro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   24
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas:"
         Height          =   195
         Left            =   20
         TabIndex        =   23
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trocas:"
         Height          =   195
         Left            =   20
         TabIndex        =   22
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   18
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Listados:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   20
         TabIndex        =   17
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   5880
         TabIndex        =   19
         Top             =   120
         Width           =   1215
         Begin VB.OptionButton optnTroca 
            Caption         =   "Troca"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optnVenda 
            Caption         =   "Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7440
         Picture         =   "frmEdiTrocas.frx":0CCA
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   7320
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiTrocas.frx":0E14
         Left            =   120
         List            =   "frmEdiTrocas.frx":0E24
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Até:"
         Height          =   195
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&De:"
         Height          =   195
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         Caption         =   "&Expressão:"
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   7320
      TabIndex        =   0
      Top             =   960
      Width           =   1575
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiTrocas.frx":0E48
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Rel. Analítico"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiTrocas.frx":0F92
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rel. Simplificado"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiTrocas.frx":10DC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Visualizar"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiTrocas.frx":1226
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1320
      Top             =   6840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiTrocas.frx":1370
      Height          =   4695
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiTrocas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmndbtnBuscar_Click()
    busca
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select distinct tt.*,tv.num_ven,tv.dataefet_ven,tv.nomecli_ven, "
    strSQL = strSQL & "tvf.valorfinal_ven, (tt.acres_tro - tt.abat_tro) as final "
    strSQL = strSQL & "from " & TAB_VENDAS & " as tv, "
    strSQL = strSQL & "tab_trocas as tt," & TAB_ITVENDAS & " as ti,  "
    strSQL = strSQL & " " & TAB_VENFEC & "  as tvf "
    strSQL = strSQL & "where tt.num_ven=tv.num_ven "
    strSQL = strSQL & "and tv.num_ven=ti.num_ven and "
    strSQL = strSQL & "tv.num_ven = tvf.num_ven "
    If optnVenda.Value Then
        strSQL = strSQL & "and tv.dataefet_ven >= " & F_GravaData(txtbxDataIni.Text)
        strSQL = strSQL & "and tv.dataefet_ven <= " & F_GravaData(txtbxDataFin.Text)
    Else
        strSQL = strSQL & "and tt.data_tro >= " & F_GravaDataTempo(txtbxDataIni.Text)
        strSQL = strSQL & "and tt.data_tro <= " & F_GravaDataTempo(txtbxDataFin.Text & " 23:59:59")
    End If
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & " and tv.num_ven like '" & txtbxExpr.Text & "%' "
        strSQL = strSQL & " order by tv.num_ven,seq_tro "
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & " and tt.seq_tro like '" & txtbxExpr.Text & "%' "
        strSQL = strSQL & " order by seq_tro,tv.num_ven"
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & " and tv.cod_cli ='" & txtbxExpr.tag & "' "
        strSQL = strSQL & " order by nomecli_ven,tv.num_ven,seq_tro "
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & " and tv.cod_cli ='" & txtbxExpr.tag & "' "
        strSQL = strSQL & " order by seq_tro,tv.num_ven"
    End If
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
    soma
End Sub
Sub soma()
    Dim cTotalTro As Currency
    Dim cTotalVen As Currency
    With dtVendas.Recordset
        Do While Not .EOF
            cTotalTro = cTotalTro + !Final
            cTotalVen = cTotalVen + !valorfinal_ven
            .MoveNext
        Loop
        dtVendas.Refresh
    End With
    lblTro.caption = FormatVal(cTotalTro)
    lblVen.caption = FormatVal(cTotalVen)
    lblList.caption = dtVendas.Recordset.RecordCount
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim fr As New frmVisualizaTro
    On Error Resume Next
    With dtVendas.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        fr.caption = "Troca: " & !seq_tro & " - Venda: " & !num_ven & _
        " - Cliente: " & !nomecli_ven
        fr.tag = dtVendas.Recordset!seq_tro
        fr.lblAbat.caption = FormatVal(!abat_tro)
        fr.lblAcr.caption = FormatVal(!acres_tro)
        fr.lblTotal.caption = FormatVal(!acres_tro - !abat_tro)
        fr.Show 0
        
    End With
End Sub

Private Sub Form_Activate()
    'MudaAparencia Me
End Sub

Private Sub Form_Load()
    dtVendas.ConnectionString = strConnect
    cmbxBusca.ListIndex = 0
    P_CentraForm Me
End Sub

Private Sub txtbxdatafin_GotFocus()
    ganhaFoco txtbxDataFin
End Sub

Private Sub txtbxdatafin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxdatafin_LostFocus()
    P_AjustaData txtbxDataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()
    ganhaFoco txtbxExpr
End Sub
Sub buscaCli()
    frmSelectCli.txtbxNome.tag = "C"
    frmSelectCli.txtbxNome.Text = txtbxExpr.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    
    
    
    txtbxExpr.Text = frmSelectCli.adoR!nome
    txtbxExpr.tag = frmSelectCli.adoR!cod
    
    
    
    
    Unload frmSelectCli
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
             buscaCli
        ElseIf cmbxBusca.ListIndex = 3 Then
            buscaProd
        End If
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    AjustaTexto txtbxExpr
End Sub
