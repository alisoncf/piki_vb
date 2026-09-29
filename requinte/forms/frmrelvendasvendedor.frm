VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelVendasVendedor 
   Caption         =   "Relatório de Vendas do Vendedor"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   14685
   Icon            =   "frmrelvendasvendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   14685
   Begin TabDlg.SSTab sstb 
      Height          =   6975
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Produto/Fabricante"
      TabPicture(0)   =   "frmrelvendasvendedor.frx":27A2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sintético: Fabricante/Funcionário"
      TabPicture(1)   =   "frmrelvendasvendedor.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridFab"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Produtos por Venda"
      TabPicture(2)   =   "frmrelvendasvendedor.frx":27DA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "gridVenda"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid gridFab 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   19
         Top             =   480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   9975
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
         Caption         =   "Sintético"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cod_fun"
            Caption         =   "Cód"
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
            DataField       =   "nome_fun"
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
         BeginProperty Column02 
            DataField       =   "nome_for"
            Caption         =   "Fabricante"
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
            DataField       =   "total"
            Caption         =   "Total (Preço de Venda)"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2264,882
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1890,142
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grid 
         Height          =   6015
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   12735
         _ExtentX        =   22463
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
         Caption         =   "Listagem por Produtos/Marca"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "nome_fun"
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
         BeginProperty Column01 
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
            DataField       =   "pv_pro"
            Caption         =   "Preço de Venda"
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
            DataField       =   "qtd"
            Caption         =   "Qtd Vendida"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4305,26
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1890,142
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridVenda 
         Height          =   6255
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   11033
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         Caption         =   "Listagem por Produtos/Venda"
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "nome_fun"
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
         BeginProperty Column01 
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
            DataField       =   "qtd"
            Caption         =   "Qtd Vendida"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "telefonecli_ven"
            Caption         =   "Telefone"
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
            DataField       =   "dataefet_ven"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3674,835
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1049,953
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   14685
      TabIndex        =   15
      Top             =   8235
      Width           =   14685
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   840
         TabIndex        =   17
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   405
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
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   0
      Width           =   14685
      Begin VB.TextBox txtFabricante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   40
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   11880
         Picture         =   "frmrelvendasvendedor.frx":27F6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnGraf 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   10440
         Picture         =   "frmrelvendasvendedor.frx":28EB
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   615
         Left            =   9000
         Picture         =   "frmrelvendasvendedor.frx":2A35
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   7560
         Picture         =   "frmrelvendasvendedor.frx":2E42
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxVendedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFun 
         Height          =   285
         Left            =   3960
         Picture         =   "frmrelvendasvendedor.frx":31CC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdFabr 
         Height          =   285
         Left            =   6600
         Picture         =   "frmrelvendasvendedor.frx":3556
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
         Height          =   195
         Left            =   4320
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário:"
         Height          =   195
         Left            =   2400
         TabIndex        =   5
         Top             =   0
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmRelVendasVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim WithEvents adoVenda As ADODB.Recordset
Attribute adoVenda.VB_VarHelpID = -1
Dim WithEvents adoFab As ADODB.Recordset
Attribute adoFab.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmbxFunc_Click()
    cmbxFunc.tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & cmbxFunc.text & "'")
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyDelete Then
        cmbxFunc.ListIndex = -1
    End If
End Sub

Private Sub cmdFabr_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtFabricante.text)
    If f.cod_For <> "" Then
        txtFabricante.text = f.nome_for
        txtFabricante.tag = f.cod_For
        
    End If
    
    
    
End Sub

Private Sub cmdFun_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectVendedor(txtbxVendedor.text)
    If f.cod_fun <> "" Then
        txtbxVendedor.text = f.nome_fun
        txtbxVendedor.tag = f.cod_fun
    End If
    
End Sub

Sub buscaProd()
    Dim strSQL As String
    strSQL = "select tf.nome_fun,tp.cod_pro,sum(ti.qtd_ite) as qtd, sum(final_ite) as total,tp.nome_pro from tab_funcionarios as tf,"
    strSQL = strSQL & " tab_produtos as tp, tab_vendas tv,"
    strSQL = strSQL & " tab_itvendas ti "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro"
    strSQL = strSQL & " and tv.num_ven=ti.num_ven"
    strSQL = strSQL & " and tv.abrevenda_usu=tf.cod_fun "
    strSQL = strSQL & " and situacao_ven=4 "
    strSQL = strSQL & " and tipo_operacao=1 "
    If txtbxVendedor.tag <> "" Then
        strSQL = strSQL & "  and tv.abrevenda_usu='" & txtbxVendedor.tag & "' "
    End If
    If txtFabricante.tag <> "" Then
        strSQL = strSQL & "  and tp.cod_fab='" & txtFabricante.tag & "' "
    End If
    
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxDAtaFIn.text <> "" Then
        strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.text)
    End If
    strSQL = strSQL & " group by tv.abrevenda_usu,tp.cod_fab,tp.cod_pro "
    
    
    
    abreObjetoAdo ado, strSQL, grid

End Sub
Sub buscaVenda()
    Dim strSQL As String
    strSQL = "select tf.nome_fun,tp.cod_pro,ti.qtd_ite as qtd,dataefet_ven, nomecli_ven,telefonecli_ven,tv.num_ven,final_ite as total,"
    strSQL = strSQL & " tp.nome_pro from tab_funcionarios as tf,"
    strSQL = strSQL & " tab_produtos as tp, tab_vendas tv,"
    strSQL = strSQL & " tab_itvendas ti "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro"
    strSQL = strSQL & " and tv.num_ven=ti.num_ven"
    strSQL = strSQL & " and tv.abrevenda_usu=tf.cod_fun "
    strSQL = strSQL & " and situacao_ven=4 "
    strSQL = strSQL & " and tipo_operacao=1 "
    If txtbxVendedor.tag <> "" Then
        strSQL = strSQL & "  and tv.abrevenda_usu='" & txtbxVendedor.tag & "' "
    End If
    If txtFabricante.tag <> "" Then
        strSQL = strSQL & "  and tp.cod_fab='" & txtFabricante.tag & "' "
    End If
    
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxDAtaFIn.text <> "" Then
        strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.text)
    End If
    strSQL = strSQL & " order by dataefet_ven,nome_fun,nome_pro     "
    
    
    abreObjetoAdo adoVenda, strSQL, gridVenda

End Sub

Sub BuscaFab()
    Dim strSQL As String
    strSQL = "select tf.nome_fun,tf.cod_fun,tfa.nome_for,sum(final_ite) as total from tab_funcionarios as tf,"
    strSQL = strSQL & " tab_produtos as tp,tab_fornecedores as tfa, tab_vendas tv,"
    strSQL = strSQL & " tab_itvendas ti "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro"
    strSQL = strSQL & " and tp.cod_fab=tfa.cod_for"
    strSQL = strSQL & " and tv.num_ven=ti.num_ven"
    strSQL = strSQL & " and tv.abrevenda_usu=tf.cod_fun "
    strSQL = strSQL & " and situacao_ven=4 "
    strSQL = strSQL & " and tipo_operacao=1"
    If txtbxVendedor.tag <> "" Then
        strSQL = strSQL & "  and tv.abrevenda_usu='" & txtbxVendedor.tag & "' "
    End If
    If txtFabricante.tag <> "" Then
        strSQL = strSQL & "  and tp.cod_fab='" & txtFabricante.tag & "' "
    End If
    If txtbxDataIni.text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtbxDataIni.text)
    End If
    If txtbxDAtaFIn.text <> "" Then
        strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxDAtaFIn.text)
    End If
    
    strSQL = strSQL & " group by tp.cod_fab,tv.abrevenda_usu "
    strSQL = strSQL & " order by total desc"
    
    
    abreObjetoAdo adoFab, strSQL, gridFab
End Sub
Private Sub cmndbtnBuscar_Click()
 
    If sstb.Tab = 0 Then
        buscaProd
        somaProd
    ElseIf sstb.Tab = 1 Then
        BuscaFab
        somaFab
    ElseIf sstb.Tab = 2 Then
        buscaVenda
        somaVenda
    End If
    
    
    

    
    
End Sub
Sub somaVenda()
    On Error GoTo trata_Erro
    Dim cTotal As Currency
        
    gridVenda.Visible = False
    
    With adoVenda
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            cTotal = cTotal + !total
            
            .MoveNext
        Loop
        lblTotal.caption = FormatVal(cTotal)
        adoVenda.Requery
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    gridVenda.Visible = True
trata_Erro:
    P_MostraErro
    gridVenda.Visible = True
End Sub
Sub somaProd()
    On Error GoTo trata_Erro
    Dim cTotal As Currency
    
    grid.Visible = False
    
    With ado
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !total
            
            .MoveNext
        Loop
        lblTotal.caption = FormatVal(cTotal)
        ado.Requery
        .MoveFirst
    End With
    grid.Visible = True
trata_Erro:
    P_MostraErro
    grid.Visible = True
End Sub
Sub somaFab()
    On Error GoTo trata_Erro
    Dim cTotal As Currency
    gridFab.Visible = False
    
    
    With adoFab
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !total
            
            .MoveNext
        Loop
        lblTotal.caption = FormatVal(cTotal)
        adoFab.Requery
        .MoveFirst
    End With
    gridFab.Visible = True
trata_Erro:
    P_MostraErro
    gridFab.Visible = True
    
End Sub


Private Sub cmndbtnGraf_Click()
    If sstb.Tab = 0 Then
        ExportaGridToCSV grid, ado
    ElseIf sstb.Tab = 1 Then
        ExportaGridToCSV gridFab, adoFab
    ElseIf sstb.Tab = 2 Then
        ExportaGridToCSV gridVenda, adoVenda
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    If sstb.Tab = 0 Then
        P_DtgrdToHtml "Relatório de Vendas por Fabricante", grid, ado, "Fabricante: " & txtFabricante.text & _
        " - Vendedor: " & txtbxVendedor.text, "3,4", "Vendedor: " & txtbxVendedor.tag
    ElseIf sstb.Tab = 1 Then
        P_DtgrdToHtml "Relatório de Vendas por Fabricante: Sintético", gridFab, adoFab, "Fabricante: " & txtFabricante.text & _
        " de " & txtbxDataIni.text & " até " & txtbxDAtaFIn.text, "2"
    ElseIf sstb.Tab = 2 Then
        P_DtgrdToHtml "Relatório de Vendas por Fabricante: Produtos", gridVenda, adoVenda, "Fabricante: " & txtFabricante.text & _
        " de " & txtbxDataIni.text & " até " & txtbxDAtaFIn.text, "2"
    
    End If
End Sub
Sub imprHTML()
    Dim sNomeArq As String
    Dim sCodFun As String
    Dim cTotal As Currency
    Dim cTotalComiss As Currency
    Dim cTotalTotalComiss As Currency
    Dim cTotalTotal As Currency
    sNomeArq = F_AbreHtml("rel_comiss")
    P_HtmlAbreHead
    P_HtmlAbreTitle 1, "Relatório de Comissões", True
    P_HtmlFechaHead
    P_HtmlAbreBody
    P_HtmlAbreTable 1, 1, 700
    
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Relatório de " & IIf(optnVendas.Value, "Comissões", "Diferenças de Trocas") & ". Período de " & txtbxDataIni.text & " Até " & txtbxDAtaFIn.text, , True, , "4"
    P_HtmlFechaLinha
    
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            If sCodFun <> !cod_fun Then
                If .AbsolutePosition > 0 Then
                    P_HtmlAbreLinha
                        P_HtmlAbreColuna 1, "Total"
                        P_HtmlAbreColuna 1, FormatVal(cTotal)
                        P_HtmlAbreColuna 1, FormatVal(cTotalComiss)
                        P_HtmlAbreColuna 1, "."
                    P_HtmlFechaLinha
                End If
                
                sCodFun = !cod_fun
                cTotal = 0
                cTotalComiss = 0
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, !nome_fun & " (Cargo: " & !nome_car & ")", , True, , "4"
                P_HtmlFechaLinha
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, "Tipo Pag."
                    P_HtmlAbreColuna 1, "Vlr Vendido."
                    P_HtmlAbreColuna 1, "Comissão"
                    P_HtmlAbreColuna 1, "%"
                P_HtmlFechaLinha
            End If
            P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, !tipo_cai
                    P_HtmlAbreColuna 1, FormatVal(!total)
                    P_HtmlAbreColuna 1, FormatVal(!comiss)
                    
                    P_HtmlAbreColuna 1, !comiss_car
            P_HtmlFechaLinha
            
            cTotal = cTotal + !total
            cTotalComiss = cTotalComiss + !comiss
            cTotalTotal = cTotalTotal + !total
            cTotalTotalComiss = cTotalTotalComiss + !comiss
            .MoveNext
        Loop
        
    End With
    
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Soma", , True, , "4"
    P_HtmlFechaLinha
    P_HtmlAbreLinha
        P_HtmlAbreColuna 1, "Geral"
        P_HtmlAbreColuna 1, FormatVal(cTotalTotal)
        P_HtmlAbreColuna 1, FormatVal(cTotalTotalComiss)
        P_HtmlAbreColuna 1, "."
    P_HtmlFechaLinha
    
    
    P_HtmlFechaTable 1
    P_HtmlFechaBody
    P_FechaHtml , True, sNomeArq
    
    
    
    
End Sub
    
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub
Sub imprime(obj As Object)
    P_CabRelTimbr "Relatório de Comissões de " & txtbxDataIni.text & " até " & txtbxDAtaFIn.text, obj
    Dim sCodFun As String
    P_ImprLinha obj, "-"
    
    
    
    With dt.Recordset
        .MoveFirst
        Do While Not .EOF
            If !cod_fun <> sCodFun Then
                sCodFun = !cod_fun
                P_ImprX obj, "Nome", 0
                P_ImprLinha obj, "-"
                P_ImprX obj, "TipoCai", 10
                P_ImprX obj, "Vendido", 70, True, True
                P_ImprX obj, "Comissão", 100, True, True
                P_ImprX obj, "Perc.", 115, False, True
                P_ImprLinha obj, "-"
            End If
            
            P_ImprX obj, !guerra_fun, 10
            P_ImprX obj, FormatVal(!total), 70, True, True
            P_ImprX obj, FormatVal(!comiss), 100, True, True
            P_ImprX obj, FormatVal(!total / CCur(lblVendas.caption) * 100), 115, False, True

            .MoveNext
        Loop
        P_RodapeVis obj, constNome_sis
    End With
    
    
End Sub


Private Sub Form_Activate()
    bActive = True
    
End Sub

Private Sub Form_Load()
    
    resizeForm Me
    
    Set ado = New ADODB.Recordset
    Set adoFab = New ADODB.Recordset
    Set adoVenda = New ADODB.Recordset
    
    sstb.Tab = 0
    sstb_Click 0
End Sub

Private Sub Option1_Click()
    dbgrd.caption = "Diferenças de Troca no Período"
    cmndbtnBuscar_Click
End Sub

Private Sub optnVendas_Click()
    dbgrd.caption = "Comissões de Vendas no Período"
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Resize()
    gridVenda.Left = 10
    gridFab.Left = 10
    grid.Left = 10
    PosicionaGrade sstb, Me, pcBot
    PosicionaGradeVal grid, sstb.width, sstb.height
    PosicionaGradeVal gridFab, sstb.width, sstb.height
    PosicionaGradeVal gridVenda, sstb.width, sstb.height - 100
    
End Sub

Private Sub sstb_Click(PreviousTab As Integer)
    If sstb.Tab = 0 Then
        grid.Visible = True
        gridVenda.Visible = False
        gridFab.Visible = False
    ElseIf sstb.Tab = 1 Then
        grid.Visible = False
        gridVenda.Visible = False
        gridFab.Visible = True
    Else
        grid.Visible = False
        gridVenda.Visible = True
        gridFab.Visible = False
    End If
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
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

Private Sub txtbxVendedor_Change()
    txtbxVendedor.tag = ""
End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFun_Click
    End If
End Sub

Private Sub txtFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdFabr_Click
    End If
End Sub
