VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Pedidos"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7500
   Begin VB.Frame frControle 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Data dtVendas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2055
         Begin VB.ComboBox cmbxSituacao_ven 
            Height          =   315
            ItemData        =   "frmEdiPed.frx":0000
            Left            =   120
            List            =   "frmEdiPed.frx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
         Begin VB.ComboBox cmbxLoja 
            Height          =   315
            ItemData        =   "frmEdiPed.frx":004D
            Left            =   120
            List            =   "frmEdiPed.frx":005D
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Situação do Pedido:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Loja:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   345
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmEdiPed.frx":0088
         Height          =   3855
         Left            =   2280
         OleObjectBlob   =   "frmEdiPed.frx":009F
         TabIndex        =   11
         Top             =   240
         Width           =   5055
      End
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Imprimir"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiPed.frx":0DD6
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiPed.frx":1160
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiPed.frx":14EA
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo Pedido"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiPed.frx":1634
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmEdiPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxBusca_Click()
    'P_Busca
End Sub

Private Sub cmbxOrd_Click()
    'P_Busca
End Sub

Private Sub cmbxLoja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_ven_Click()
   P_Busca
End Sub

Private Sub cmbxSituacao_ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    
    P_Status "Montando venda..."
    With dtVendas.Recordset
        If !situacao_ped <> "1" Then
            MsgBox "Este pedido não pode ser editado!", vbExclamation
            P_Status ""
            Exit Sub
        End If
        frmPedido.Tag = !Num_Ped
        'fr.lblAbert.caption = !abertura_ven
        P_LocalizaOperacao !cod_loj, frmPedido.cmbxLoja
        
        frmPedido.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_usu", "tab_funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.lblValor.caption = FormatVal(!valor_ped)
        frmPedido.caption = "Pedido " & !Num_Ped & "-" & !nome_loj
        frmPedido.lblPedido.caption = !Num_Ped
        frmPedido.txtbxEndereco.Text = Format(!endereco_cli)
        frmPedido.txtbxNome_Cli.Text = Format(!nome_cli)
        frmPedido.Show 1
    End With
    P_Status ""
End Sub

Private Sub cmndbtnImpr_Click()
    P_Status "Gerando impressão... "
    If dtVendas.Recordset!situacao_ped <> "0" Then
        P_ImprimePedido frmVisualiza.Pic, dtVendas.Recordset!Num_Ped, False
    End If
    P_Status ""
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        
        If dtVendas.Recordset!situacao_ped = 0 Then
            P_ImprimePedidoCancelado Printer, dtVendas.Recordset!Num_Ped, True
        Else
            P_ImprimePedido Printer, dtVendas.Recordset!Num_Ped, True
        End If
        
                
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnNew_Click()
    
    Dim strSQL As String
    If MsgBox("Confirma abertura de Pedido? ", vbYesNo + vbQuestion) = vbYes Then
        strSQL = F_GeraNovoPedido(strCod_Loj)
        If strSQL <> "" Then
            frmPedido.lblPedido.caption = strSQL
            frmPedido.Tag = strSQL
            On Error Resume Next
            frmPedido.cmbxLoja.ListIndex = 0
            frmPedido.Show 1
        End If
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub DBGrid1_DblClick()
    If cmndbtnEdit.Enabled = True Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DBGrid1_DblClick
    End If
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub dtVendas_Reposition()
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        If dtVendas.Recordset!situacao_ped = "1" Then
            cmndbtnImpr.Enabled = False
            cmndbtnEdit.Enabled = True
        Else
            cmndbtnImpr.Enabled = True
            cmndbtnEdit.Enabled = False
        End If
    End If
    
    
    If dtVendas.Recordset!situacao_ped > 1 Then
        cmndbtnEdit.Enabled = False
        
    End If
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    'Me.Width = 10725
    'Me.Height = 6135
    P_CentraForm Me
    P_PreencheLojas cmbxLoja, False
    dtVendas.Connect = banco.Connect
    
    
    'cmbxBusca.ListIndex = 0
    'cmbxOrd.ListIndex = 0
    cmbxSituacao_ven.ListIndex = 0
    P_Busca
    bActive = True
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    On Error GoTo trata_erro
    
    
    
    strSQL = "select tc.nome_loj,tv.* from tab_lojas as tc, "
    strSQL = strSQL & " tab_pedidos as tv "
    
    strSQL = strSQL & " where tv.situacao_ped='" & Left(cmbxSituacao_ven.Text, 1) & "' "
    strSQL = strSQL & " and tc.cod_loj=tv.cod_loj "
    strSQL = strSQL & " order by num_ped "
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
trata_erro:
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub


Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    P_Busca
End Sub



