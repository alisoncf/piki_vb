VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiPedidoCompra 
   Caption         =   "Pedidos de Compra"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13770
   Icon            =   "frmEdiPedidoCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   13770
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13740
      TabIndex        =   15
      Top             =   7785
      Width           =   13770
      Begin VB.CommandButton cmndbtnGradeProd 
         Caption         =   "Grade de Produtos"
         Height          =   690
         Left            =   7920
         Picture         =   "frmEdiPedidoCompra.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   690
         Left            =   4800
         Picture         =   "frmEdiPedidoCompra.frx":09A7
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   690
         Left            =   12240
         Picture         =   "frmEdiPedidoCompra.frx":0DB4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   690
         Left            =   1680
         Picture         =   "frmEdiPedidoCompra.frx":0EFE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   690
         Left            =   120
         Picture         =   "frmEdiPedidoCompra.frx":1048
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   690
         Left            =   3240
         Picture         =   "frmEdiPedidoCompra.frx":13D2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprGrade 
         Caption         =   "&Grade de Pedidos"
         Height          =   690
         Left            =   6360
         Picture         =   "frmEdiPedidoCompra.frx":151C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   13740
      TabIndex        =   0
      Top             =   0
      Width           =   13770
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   450
         Left            =   12120
         Picture         =   "frmEdiPedidoCompra.frx":1929
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmEdiPedidoCompra.frx":1CB3
         Left            =   120
         List            =   "frmEdiPedidoCompra.frx":1CC0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxFornecedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   40
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiPedidoCompra.frx":1CE9
         Left            =   10320
         List            =   "frmEdiPedidoCompra.frx":1CF9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         Picture         =   "frmEdiPedidoCompra.frx":1D21
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9840
         Picture         =   "frmEdiPedidoCompra.frx":20AB
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   7200
         TabIndex        =   14
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   4200
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   10320
         TabIndex        =   7
         Top             =   0
         Width           =   885
      End
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   1200
      Top             =   7200
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
      Bindings        =   "frmEdiPedidoCompra.frx":2435
      Height          =   6735
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11880
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
   Begin MSAdodcLib.Adodc adcPed 
      Height          =   375
      Left            =   8760
      Top             =   8040
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
   Begin MSDataGridLib.DataGrid dtgrdPed 
      Bindings        =   "frmEdiPedidoCompra.frx":2447
      Height          =   7215
      Left            =   7560
      TabIndex        =   25
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   12726
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
Attribute VB_Name = "frmEdiPedidoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = pedidoCompra.F_SelectPedidos("", txtbxDataIni.Text, txtbxDAtaFIn.Text, txtbxFornecedor.tag, _
        txtbxProduto.tag, Left(cmbxTipo.Text, 1))
        
    dtNot.RecordSource = strSQL
    dtNot.Refresh
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_erro
    frmPedidoCompra.tag = dtNot.Recordset!seq_ped
    frmPedidoCompra.Show 0
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnGradeProd_Click()
    'P_DtgrdToHtml "Produtos", dbgrdPed, dtPed, , "2,3,4"
End Sub

Private Sub cmndbtnImpr_Click()
    On Error GoTo trata_erro
    ImprimePedidoCompra dtNot.Recordset!seq_ped
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnImprGrade_Click()
    P_DbgrdToHtml "Pedidos", dtgrd, adc
    
End Sub

Private Sub cmndbtnNew_Click()
    frmPedidoCompra.Show 0
End Sub

Private Sub cmndbtnProd_Click()
    Dim p As Produtos.Produto
    p = F_BuscaProdSaida(txtbxProduto.Text)
    If p.cod_pro <> "" Then
        txtbxProduto.Text = p.nome_pro
        txtbxProduto.tag = p.cod_pro
        cmndbtnBuscar_Click
        cmbxOrder.SetFocus
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtbxFornecedor.Text)
    If f.cod_for <> "" Then
        txtbxFornecedor.Text = f.nome_for
        txtbxFornecedor.tag = f.cod_for
        cmndbtnBuscar_Click
        txtbxProduto.SetFocus
    End If
End Sub

Private Sub dtNot_Reposition()
   ' On Error Resume Next
    busca_prod
End Sub
Sub busca_prod()
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = F_SelectItPed(dtNot.Recordset!seq_ped)
    dtPed.RecordSource = strSQL
    dtPed.Refresh
    Exit Sub
trata_erro:

End Sub
Private Sub Form_Activate()
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Load()
    adc.ConnectionString = strConnect
    adcPed.ConnectionString = strConnect
    resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Height = pcTop.Top - dbgrd.Top - 200
    dbgrdPed.Width = Me.Width - dbgrdPed.Left - 200
    dbgrdPed.Height = dbgrd.Height
    
End Sub

Private Sub txtbxdataFin_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    AjustaTexto txtbxDAtaFIn
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

Private Sub txtbxFornecedor_Change()
    txtbxFornecedor.tag = ""
End Sub

Private Sub txtbxFornecedor_GotFocus()
    ganhaFoco txtbxFornecedor
End Sub

Private Sub txtbxFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxProduto_Change()
    txtbxProduto.tag = ""
End Sub

Private Sub txtbxProduto_GotFocus()
    ganhaFoco txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnProd_Click
    End If
End Sub
