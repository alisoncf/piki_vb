VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
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
   Begin VB.Data dtPed 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pcTop 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13740
      TabIndex        =   16
      Top             =   7785
      Width           =   13770
      Begin VB.CommandButton cmndbtnGradeProd 
         Caption         =   "Grade de Produtos"
         Height          =   690
         Left            =   7920
         Picture         =   "frmEdiPedidoCompra.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   690
         Left            =   4800
         Picture         =   "frmEdiPedidoCompra.frx":09A7
         Style           =   1  'Graphical
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   690
         Left            =   1680
         Picture         =   "frmEdiPedidoCompra.frx":0EFE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   690
         Left            =   120
         Picture         =   "frmEdiPedidoCompra.frx":1048
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   690
         Left            =   3240
         Picture         =   "frmEdiPedidoCompra.frx":13D2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprGrade 
         Caption         =   "&Grade de Pedidos"
         Height          =   690
         Left            =   6360
         Picture         =   "frmEdiPedidoCompra.frx":151C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Data dtNot 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
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
         TabIndex        =   23
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   7200
         TabIndex        =   15
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
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiPedidoCompra.frx":2435
      Height          =   6495
      Left            =   0
      OleObjectBlob   =   "frmEdiPedidoCompra.frx":2449
      TabIndex        =   12
      Top             =   840
      Width           =   5655
   End
   Begin MSDBGrid.DBGrid dbgrdPed 
      Bindings        =   "frmEdiPedidoCompra.frx":3342
      Height          =   6495
      Left            =   5760
      OleObjectBlob   =   "frmEdiPedidoCompra.frx":3356
      TabIndex        =   24
      Top             =   840
      Width           =   7935
   End
End
Attribute VB_Name = "frmEdiPedidoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    strSQL = pedidoCompra.F_SelectPedidos("", txtbxDataIni.Text, txtbxdataFin.Text, txtbxFornecedor.Tag, _
        txtbxProduto.Tag, Left(cmbxTipo.Text, 1))
        
    dtNot.RecordSource = strSQL
    dtNot.Refresh
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro
    frmPedidoCompra.Tag = dtNot.Recordset!seq_ped
    frmPedidoCompra.Show 0
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnGradeProd_Click()
    P_DbgrdToHtml "Produtos", dbgrdPed, dtPed, , "2,3,4"
End Sub

Private Sub cmndbtnImpr_Click()
    On Error GoTo trata_Erro
    ImprimePedidoCompra dtNot.Recordset!seq_ped
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImprGrade_Click()
    P_DbgrdToHtml "Pedidos", dbgrd, dtNot
    
End Sub

Private Sub cmndbtnNew_Click()
    frmPedidoCompra.Show 0
End Sub

Private Sub cmndbtnProd_Click()
    Dim p As Produtos.Produto
    p = F_BuscaProdSaida(txtbxProduto.Text)
    If p.cod_pro <> "" Then
        txtbxProduto.Text = p.nome_pro
        txtbxProduto.Tag = p.cod_pro
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
        txtbxFornecedor.Tag = f.cod_for
        cmndbtnBuscar_Click
        txtbxProduto.SetFocus
    End If
End Sub

Private Sub dtNot_Reposition()
   ' On Error Resume Next
    busca_prod
End Sub
Sub busca_prod()
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = F_SelectItPed(dtNot.Recordset!seq_ped)
    dtPed.RecordSource = strSQL
    dtPed.Refresh
    Exit Sub
trata_Erro:

End Sub
Private Sub Form_Activate()
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Load()
    dtNot.Connect = banco.Connect
    dtPed.Connect = banco.Connect
    resizeForm Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Height = pcTop.Top - dbgrd.Top - 200
    dbgrdPed.Width = Me.Width - dbgrdPed.Left - 200
    dbgrdPed.Height = dbgrd.Height
    
End Sub

Private Sub txtbxdataFin_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxdataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxdataFin_LostFocus()
    P_AjustaTexto txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxFornecedor_Change()
    txtbxFornecedor.Tag = ""
End Sub

Private Sub txtbxFornecedor_GotFocus()
    P_SelecTexto txtbxFornecedor
End Sub

Private Sub txtbxFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxProduto_Change()
    txtbxProduto.Tag = ""
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnProd_Click
    End If
End Sub
