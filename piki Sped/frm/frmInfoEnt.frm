VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInfoEnt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informações de Compra"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtProduto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   8055
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14160
      TabIndex        =   3
      Top             =   7770
      Width           =   14190
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   10320
         Picture         =   "frmInfoEnt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnFechar 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   12240
         Picture         =   "frmInfoEnt.frx":04F1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   8400
         Picture         =   "frmInfoEnt.frx":073B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.OptionButton optnGer 
      Caption         =   "Usar &Filtro"
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton optnFiscal 
      Caption         =   "Informações &Gerais"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   330
      Left            =   1680
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSDataGridLib.DataGrid dbgrd 
      Bindings        =   "frmInfoEnt.frx":0885
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   11033
      _Version        =   393216
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
      Caption         =   "Registros"
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
   Begin VB.CommandButton cmdProd 
      Height          =   450
      Left            =   8160
      Picture         =   "frmInfoEnt.frx":0897
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblCodFab 
      AutoSize        =   -1  'True
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8760
      TabIndex        =   11
      Top             =   240
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frmInfoEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim bActive As Boolean


Private Sub cmdExportar_Click()
    ExportaGridToCSV dbgrd, adc.Recordset
End Sub

Private Sub cmdProd_Click()
    buscaProd
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DtgrdToHtml "Histórico de Produtos", dbgrd, adc
End Sub

Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    busca
    bActive = True
End Sub
Sub busca()
    Dim strSQL As String
    Dim sTabCab As String
    Dim sTabIte As String
    
    
    strSQL = "select codfab_pro, tp.cod_pro,tp.nome_pro,tf.nome as Fornecedor,status,data,num_doc as nota,ti.id,ti.qtd_convert as qtd, "
    strSQL = strSQL & " ti.custo,ti.cfop, "
    strSQL = strSQL & " ti.vl_item "
    strSQL = strSQL & " from sped_nf as tc inner join sped_nf_itens as ti on tc.id=ti.id_nf "
    strSQL = strSQL & " left join sped_participante as tf on tf.codigo=tc.cod_part "
    
    strSQL = strSQL & " inner join " & TAB_PRODUTOS & " as tp on ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " where ti.cod_pro='" & txtProduto.tag & "' "
    
    strSQL = strSQL & " group by ti.id_nf ,ti.cod_pro "
    strSQL = strSQL & " order by data desc"
    adc.RecordSource = strSQL
    adc.Refresh
    With adc.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        Me.caption = Me.tag & " - " & !nome_pro & " [Informações de Entradas]  "
        txtProduto.Text = !nome_pro
        txtProduto.tag = !cod_pro
        lblCodFab.caption = !codfab_pro
        dbgrd.caption = .RecordCount & " Registros"
    End With
    soma
End Sub
Sub soma()
    Dim cTotal As Currency
    Dim cValor As Currency
    dbgrd.Visible = False
    On Error GoTo trata_Erro
    With adc.Recordset
        Do While Not .EOF
            cTotal = cTotal + !qtd
            cValor = cValor + (!custo * !qtd)
            .MoveNext
        Loop
        lblTotal.caption = (cTotal) & " itens  - R$ " & FormatVal(cValor) & " - Custo Médio: R$ " & FormatVal(cValor / cTotal)
        .MoveFirst
    End With
    dbgrd.Visible = True
trata_Erro:
dbgrd.Visible = True
End Sub
Private Sub Form_Load()
    P_CentraForm Me
    adc.ConnectionString = strConnect
    
End Sub

Private Sub optnFiscal_Click()
    busca
End Sub

Private Sub optnGer_Click()
    busca
End Sub

Private Sub txtProduto_Change()
    txtProduto.tag = ""
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub
Sub buscaProd()
    Dim p As Produto
    p = Produtos.F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        lblCodFab.caption = p.codfab_pro
        busca
    End If
    
End Sub
