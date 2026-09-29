VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHistProd 
   Caption         =   "Histórico de Movimentação de Produtos"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17280
   Icon            =   "frmHistProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8700
   ScaleWidth      =   17280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8400
      TabIndex        =   9
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   735
      Left            =   15480
      Picture         =   "frmHistProd.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtAte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtDe 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtProduto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton cmdPesquisar 
      Height          =   285
      Left            =   5640
      Picture         =   "frmHistProd.frx":0898
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   315
   End
   Begin MSDataGridLib.DataGrid gridCompras 
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
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
      Caption         =   "Compras"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "data"
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
      BeginProperty Column01 
         DataField       =   "nome_participante"
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
      BeginProperty Column02 
         DataField       =   "ind_oper"
         Caption         =   "E/S"
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
         Caption         =   "Qtd"
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
         DataField       =   "qtd_convert"
         Caption         =   "qtd_convert"
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
         DataField       =   "unid"
         Caption         =   "Un."
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
         DataField       =   "unit"
         Caption         =   "Unitário"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3539,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1214,929
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid gridVendas 
      Height          =   4815
      Left            =   8400
      TabIndex        =   10
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8493
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
      Caption         =   "Vendas"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "num_ven"
         Caption         =   "Número"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
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
      BeginProperty Column03 
         DataField       =   "qtd_ite"
         Caption         =   "Qtd"
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
         DataField       =   "unit"
         Caption         =   "Unitário"
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
         DataField       =   "Desconto"
         Caption         =   "Desc.%"
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
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3720,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCliente 
      Height          =   285
      Left            =   13080
      Picture         =   "frmHistProd.frx":0BE8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   315
   End
   Begin MSDataGridLib.DataGrid gridDevolucao 
      Height          =   2775
      Left            =   8400
      TabIndex        =   13
      Top             =   5640
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
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
      Caption         =   "Devolução"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "data_dev"
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
      BeginProperty Column01 
         DataField       =   "qtd_dev"
         Caption         =   "Qtd"
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
         DataField       =   "guerra_fun"
         Caption         =   "Usuário"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2610,142
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid gridContagem 
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
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
      Caption         =   "Contagem"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "data"
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
      BeginProperty Column01 
         DataField       =   "hora"
         Caption         =   "Hora"
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
         DataField       =   "qtd_ant"
         Caption         =   "De"
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
         Caption         =   "Para"
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
         DataField       =   "guerra_fun"
         Caption         =   "Usuário"
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
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8400
      TabIndex        =   12
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Até:"
      Height          =   195
      Left            =   7320
      TabIndex        =   5
      Top             =   0
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "de:"
      Height          =   195
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Produto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frmHistProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adocompras As ADODB.Recordset
Attribute adocompras.VB_VarHelpID = -1
Dim WithEvents AdoVendas As ADODB.Recordset
Attribute AdoVendas.VB_VarHelpID = -1
Dim WithEvents AdoContagens As ADODB.Recordset
Attribute AdoContagens.VB_VarHelpID = -1
Dim WithEvents AdoDevolucao As ADODB.Recordset
Attribute AdoDevolucao.VB_VarHelpID = -1

Private Sub cmdCliente_Click()
    buscaCli
End Sub
Sub buscaCli()
    Dim c As cliente
    c = Clientes.F_BuscaCliente(txtCliente.text)
    If c.cod_cli <> "" Then
        txtCliente.text = c.nome_Cli
        txtCliente.tag = c.cod_cli
        busca
    End If
    
    
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdPesquisar_Click()
    Dim p As Produto
    p = Produtos.F_BuscaProdSaida(txtProduto.text)
    If p.cod_pro <> "" Then
        txtProduto.text = p.nome_pro
        txtProduto.tag = p.cod_pro
        
        
        busca
    End If
    
    
    
    
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from tab_produtos where cod_pro='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        txtProduto.text = !nome_pro
        txtProduto.tag = !cod_pro
        busca
        .Close
    End With
    Exit Sub
trata_Erro:
End Sub


Private Sub Form_Activate()
    If Me.tag <> "" Then
        busca
    End If
End Sub

Private Sub Form_Load()
    Set adocompras = New ADODB.Recordset
    Set AdoVendas = New ADODB.Recordset
    Set AdoContagens = New ADODB.Recordset
    Set AdoDevolucao = New ADODB.Recordset
    
    txtDe.text = DateAdd("M", -60, Date)
    txtDe_LostFocus
    
    txtAte.text = Date
    txtAte_LostFocus
    
End Sub


Sub busca(Optional ordem As String = "")
    buscaCompras
    buscaVendas
    buscaContagens
    buscaDevolucoes
End Sub
Sub buscaContagens()
    Dim strSQL As String
    strSQL = "select c.*,tf.guerra_fun from contagem as c inner join " & TAB_FUNCIONARIOS
    strSQL = strSQL & " as tf on tf.cod_fun=c.cod_usu "
    strSQL = strSQL & " where cod_pro='" & txtProduto.tag & "' "
    If txtDe.text <> "" Then
        strSQL = strSQL & " and data>= " & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and data<= " & F_GravaData(txtAte.text)
    End If
    abreObjetoAdo AdoContagens, strSQL, gridContagem
    
End Sub
Sub buscaDevolucoes()
    Dim strSQL As String
    strSQL = "select d.data_dev,i.*,tf.guerra_fun,tv.num_ven,tv.nomecli_ven from devolucoes as d inner join itdev as i "
    strSQL = strSQL & " on d.seq_dev=i.seq_dev inner join " & TAB_FUNCIONARIOS
    strSQL = strSQL & " as tf on tf.cod_fun=d.cod_usu "
    strSQL = strSQL & " inner join " & TAB_VENDAS & "  as tv on tv.num_ven=d.num_ven "
    strSQL = strSQL & " where cod_pro='" & txtProduto.tag & "' "
    If txtDe.text <> "" Then
        strSQL = strSQL & " and data_dev>= " & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and data_dev<= " & F_GravaData(txtAte.text)
    End If
    strSQL = strSQL & " and tipo_operacao=1"
    abreObjetoAdo AdoDevolucao, strSQL, gridDevolucao
    
End Sub

Sub buscaCompras(Optional ordem As String = "")
    Dim strSQL As String
    
    strSQL = "select p.nome as nome_participante,nf.status, vl_item/qtd_convert as unit, nf.data,ind_oper,nf.dt_e_s, i.* from sped_nf as nf"
    strSQL = strSQL & " left join sped_participante as p on nf.cod_part=p.codigo "
    strSQL = strSQL & " inner join sped_nf_itens  as i on nf.id=i.id_nf "
    strSQL = strSQL & " where cod_pro='" & txtProduto.tag & "' "
    
    If txtDe.text <> "" Then
        strSQL = strSQL & " and data>=" & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and data<=" & F_GravaData(txtAte.text)
    End If
    
    
    If ordem = "" Then
        strSQL = strSQL & " order by data desc"
    End If
    
    
    abreObjetoAdo adocompras, strSQL, gridCompras
End Sub
Sub buscaVendas()
    Dim strSQL As String
    strSQL = "select tc.nome_cli,tc.cpf_cli,tv.nomecli_ven, tv.dataefet_ven,(final_ite/qtd_ite) as unit,100-((final_ite/qtd_ite)/unit_ite*100) as desconto, "
    strSQL = strSQL & " ti.* from " & TAB_VENDAS & " tv "
    strSQL = strSQL & " inner join " & TAB_CLIENTES & " as tc on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and ti.cod_pro='" & txtProduto.tag & "' "
    If txtCliente.tag <> "" Then
        strSQL = strSQL & " and tc.cod_cli='" & txtCliente.tag & "' "
    Else
        If txtCliente.text <> "" Then
        strSQL = strSQL & " and nomecli_ven like '%" & txtCliente.text & " '%"
        End If
    End If
    If txtDe.text <> "" Then
        strSQL = strSQL & " and dataefet_ven>=" & F_GravaData(txtDe.text)
    End If
    If txtAte.text <> "" Then
        strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtAte.text)
    End If
    strSQL = strSQL & " and (situacao_ven=4 or situacao_ven=2) "
    strSQL = strSQL & " and tipo_operacao=1"
    strSQL = strSQL & " order by dataefet_ven desc "
    abreObjetoAdo AdoVendas, strSQL, gridVendas
End Sub

Private Sub gridCompras_DblClick()
    On Error GoTo trata_Erro
    VisualizarSpedNF adocompras!Id_nf
    
    Exit Sub
trata_Erro:
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_LostFocus()

End Sub

Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub
Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub

Private Sub txtProduto_Change()
    txtProduto.tag = ""
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
    
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdPesquisar_Click
    End If
End Sub
