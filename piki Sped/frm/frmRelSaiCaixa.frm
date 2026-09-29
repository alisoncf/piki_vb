VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelSaiCaixa 
   Caption         =   "Relatório de Saídas"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12915
   Icon            =   "frmRelSaiCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   12915
   Begin TabDlg.SSTab sstb 
      Height          =   6735
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Agrupado"
      TabPicture(0)   =   "frmRelSaiCaixa.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdExportar"
      Tab(0).Control(1)=   "gridAgrud"
      Tab(0).Control(2)=   "lblTotal"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Vendas"
      TabPicture(1)   =   "frmRelSaiCaixa.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTotalVendas"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "gridVendas"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdExporVendas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdExporVendas 
         Caption         =   "Exportar"
         Height          =   360
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   420
         Width           =   1335
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   360
         Left            =   -63960
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   420
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid gridAgrud 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   14
         Top             =   900
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9763
         _Version        =   393216
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
         Caption         =   "Saída de Produtos"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "cod_pro"
            Caption         =   "Código"
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
            DataField       =   "codfab_pro"
            Caption         =   "Ref"
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
            DataField       =   "nome_pro"
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
         BeginProperty Column03 
            DataField       =   "saida"
            Caption         =   "Saída"
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
            DataField       =   "devolvido"
            Caption         =   "Devolvido"
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
            DataField       =   "qtd"
            Caption         =   "Total"
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
            DataField       =   "total"
            Caption         =   "Valor Gerado"
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
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4545,071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridVendas 
         Height          =   5775
         Left            =   120
         TabIndex        =   17
         Top             =   900
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   10186
         _Version        =   393216
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
         Caption         =   "Saída de Produtos"
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "cod_pro"
            Caption         =   "Código"
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
            DataField       =   "codfab_pro"
            Caption         =   "Ref"
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
            DataField       =   "nome_pro"
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
         BeginProperty Column03 
            DataField       =   "cod_emp"
            Caption         =   "Emp"
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
         BeginProperty Column05 
            DataField       =   "qtd_ite"
            Caption         =   "qtd"
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
            DataField       =   "devolvido_ite"
            Caption         =   "Devolvido"
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
            DataField       =   "qtd"
            Caption         =   "Total"
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
            DataField       =   "unit_ite"
            Caption         =   "Unit_ite"
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
            DataField       =   "final_ite"
            Caption         =   "Final Item"
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
            Caption         =   "Valor Gerado"
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
         BeginProperty Column11 
            DataField       =   "vendedor"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4545,071
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTotalVendas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   15
         Top             =   540
         Width           =   495
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   12885
      TabIndex        =   11
      Top             =   8085
      Width           =   12915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6000
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Fe&char (Esc)"
         Height          =   615
         Left            =   11040
         Picture         =   "frmRelSaiCaixa.frx":0902
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   8640
         Picture         =   "frmRelSaiCaixa.frx":0A6D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         ItemData        =   "frmRelSaiCaixa.frx":0D7F
         Left            =   4320
         List            =   "frmRelSaiCaixa.frx":0D8C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         ItemData        =   "frmRelSaiCaixa.frx":0DAF
         Left            =   2520
         List            =   "frmRelSaiCaixa.frx":0DBC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Turno:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4320
         TabIndex        =   7
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Caixa"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   120
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmRelSaiCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim WithEvents AdoVendas As ADODB.Recordset
Attribute AdoVendas.VB_VarHelpID = -1
Dim WithEvents adoDev As ADODB.Recordset
Attribute adoDev.VB_VarHelpID = -1
Private Sub cmbxCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxCaixa.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxOrderBY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxTurno.ListIndex = -1
    End If
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV gridAgrud, ado
End Sub

Private Sub cmdExporVendas_Click()
    ExportaGridToCSV gridVendas, AdoVendas
End Sub

Private Sub cmndbtnBuscar_Click()
    
    BuscaAgrupada
    buscaVendas ""
    soma
End Sub
Sub BuscaAgrupada(Optional bAgrup As Boolean = True)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency

    
    strSQL = "select tp.cod_pro,tp.nome_pro, codfab_pro,tp.unidade_pro, "
    
    strSQL = strSQL & "sum(ti.qtd_Ite) as saida,"
    strSQL = strSQL & "sum(devolvido_ite) as devolvido,"
    strSQL = strSQL & "sum(ti.qtd_Ite-devolvido_ite) as qtd,"
    strSQL = strSQL & "sum((qtd_ite-devolvido_ite)*(final_ite/qtd_ite)) as total, "
    strSQL = strSQL & "sum((desconto_ite/qtd_ite)*(qtd_ite-devolvido_ite)) As desconto"
    
        
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv  inner join " & TAB_VENFEC & " as tc "
    strSQL = strSQL & " on tc.num_ven=tv.num_ven "
    strSQL = strSQL & " inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    
    
    strSQL = strSQL & " where 1>0 "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.Text)
    strSQL = strSQL & " and tv.situacao_ven=4 "
    If txtProduto.Text <> "" Then
        strSQL = strSQL & " and (nome_pro like '" & txtProduto.Text & "%' or codfab_pro='" & txtProduto.Text & "' or tp.cod_pro='" & txtProduto.Text & "')  "
    End If
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    If bAgrup Then
        strSQL = strSQL & " group by cod_pro "
    End If
    strSQL = strSQL & " order by total desc"
    
    abreObjetoAdo ado, strSQL, gridAgrud
        
        
    
    
    

    lblTotal.caption = "Total...........: " & FormatVal(cTotal)
End Sub
Sub buscaVendas(order As String)
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency

    
    strSQL = "select tv.num_ven,tp.cod_pro,tp.nome_pro, codfab_pro,tp.unidade_pro,guerra_fun as vendedor, nome_mar,e.cod_emp,"
    
    strSQL = strSQL & "ti.qtd_Ite,unit_ite,final_ite,"
    strSQL = strSQL & "devolvido_ite,"
    strSQL = strSQL & "(ti.qtd_Ite-devolvido_ite) as qtd,"
    strSQL = strSQL & "(final_ite/qtd_ite)*(ti.qtd_Ite-devolvido_ite) as total"
    
        
    
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on ti.cod_pro=tp.cod_pro"
    strSQL = strSQL & " inner join " & TAB_FUNCIONARIOS & " as tf on tv.abrevenda_usu=tf.cod_fun "
    strSQL = strSQL & " inner join estoque as e on ti.cod_pro=e.cod_pro and ti.cod_emp=e.cod_emp "
    strSQL = strSQL & " inner join " & TAB_MARCAS & " as tm on tm.cod_mar=tp.cod_mar     "
    strSQL = strSQL & " where 1> 0 "
    strSQL = strSQL & " and tv.dataefet_ven>=" & F_GravaData(txtbxDataIni.Text)
    strSQL = strSQL & " and tv.dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.Text)
    
    strSQL = strSQL & " and tv.situacao_ven=4 "
    If cmbxCaixa.Text <> "" Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    End If
    If cmbxTurno.Text <> "" Then
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    End If
    If txtProduto.Text <> "" Then
        strSQL = strSQL & " and (nome_pro like '" & txtProduto.Text & "%' or codfab_pro='" & txtProduto.Text & "' or tp.cod_pro='" & txtProduto.Text & "')  "
    End If
    If order = "" Then
    
        strSQL = strSQL & " order by nome_pro,ti.num_ven "
    Else
        strSQL = strSQL & " order by " & order
    End If
    abreObjetoAdo AdoVendas, strSQL, gridVendas
        
        
    
    

End Sub
Sub soma()
    Dim cTotal As Currency
    gridAgrud.Visible = False
    gridVendas.Visible = False
    With ado
        Do While Not .EOF
            cTotal = cTotal + !total
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        lblTotal.caption = "Total...:" & FormatVal(cTotal)
    End With
    cTotal = 0
    With AdoVendas
        Do While Not .EOF
            cTotal = cTotal + !total
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        lblTotalVendas.caption = "Total...:" & FormatVal(cTotal)
    End With
    gridAgrud.Visible = True
    gridVendas.Visible = True
End Sub
Sub imprime(obj As Object)
    Dim iCount As Integer
    Dim iCol As Integer
    Dim sWith As Single
    P_CabRelTimbr "Relatório de Saídas de " & txtbxDataIni.Text & " até " & txtbxDAtaFIn.Text, obj
    obj.fontsize = 12
    P_ImprX obj, "Caixa: " & cmbxCaixa, 0
    P_ImprX obj, "Turno: " & cmbxTurno, 60, False
    obj.fontsize = 10
    P_ImprLinha obj
    
        With msflxgrd
            For iCol = 0 To .Cols - 1
                .row = 0
                .col = iCol
                
                P_ImprX obj, .Text, ScaleX(sWith, 1, 6)
                sWith = sWith + .CellWidth
            Next
            P_ImprX obj, "", 0, False
            sWith = 0
            If obj.CurrentY >= obj.ScaleHeight - 30 Then
                P_FimRel obj, "Continua..."
                For iCol = 0 To .Cols - 1
                    .row = 0
                    .col = iCol
                
                    P_ImprX obj, .Text, ScaleX(sWith, 1, 6)
                    sWith = sWith + .CellWidth
                    Next
                    P_ImprX obj, "", 0, False
            End If
            sWith = 0
            For iCount = 1 To .rows - 1
                .row = iCount
                For iCol = 0 To .Cols - 1
                    .col = iCol
                
                    P_ImprX obj, .Text, ScaleX(sWith, 1, 6)
                sWith = sWith + .CellWidth
            Next
            sWith = 0
            P_ImprX obj, "", 0, False
            Next
            P_ImprLinha obj
            P_ImprX obj, lblTotal.caption, 0, False
            P_FimRel obj, strNome_Emp
        End With
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    resizeForm Me
    P_PreencheCaixas cmbxCaixa
    P_PreencheTurnos cmbxTurno
    txtbxDAtaFIn.Text = Date
    txtbxDataIni.Text = Date
    Set ado = New ADODB.Recordset
    Set AdoVendas = New ADODB.Recordset
    Set adoDev = New ADODB.Recordset
    sstb.Tab = 0
End Sub


Private Sub optnAgrupada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
End Sub


Private Sub optnSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar.SetFocus
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    PosicionaGrade sstb, Me, pcTop
    
    gridVendas.height = sstb.height - sstb.Top - 400
    gridVendas.width = sstb.width - sstb.Left - 400
    
    gridAgrud.height = gridVendas.height
    gridAgrud.width = gridVendas.width
    
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
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

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub
