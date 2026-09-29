VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiDEv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devoluções a Concluir"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frmEdiDev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frControle 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2055
         Begin VB.ComboBox cmbxSituacao_ven 
            Height          =   315
            ItemData        =   "frmEdiDev.frx":0442
            Left            =   120
            List            =   "frmEdiDev.frx":044C
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1560
            Width           =   1815
         End
         Begin VB.ComboBox cmbxOrd 
            Height          =   315
            ItemData        =   "frmEdiDev.frx":046B
            Left            =   120
            List            =   "frmEdiDev.frx":047B
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtbxExpr 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   9
            Top             =   2160
            Width           =   1815
         End
         Begin VB.ComboBox cmbxBusca 
            Height          =   315
            ItemData        =   "frmEdiDev.frx":04A6
            Left            =   120
            List            =   "frmEdiDev.frx":04B6
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "&Situação da Venda:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Ordenar:"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "E&xpressão:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1920
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Buscar Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   825
         End
      End
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
      Begin VB.Frame Frame2 
         Height          =   2655
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   2055
         Begin VB.CommandButton cmndbtnImpr 
            Caption         =   "&Imprimir"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiDev.frx":04E1
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiDev.frx":086B
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiDev.frx":0BF5
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Nova"
            Height          =   535
            Left            =   120
            Picture         =   "frmEdiDev.frx":0D3F
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmEdiDev.frx":10C9
         Height          =   5055
         Left            =   2280
         OleObjectBlob   =   "frmEdiDev.frx":10E0
         TabIndex        =   15
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmEdiDEv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxBusca_Click()
    P_Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrd_Click()
    P_Busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_ven_Click()
   P_Busca
End Sub

Private Sub cmndbtnEdit_Click()
    
    P_Status "Montando venda..."
    With dtVendas.Recordset
        frmDevolucao.tag = !Num_DEV
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        frmDevolucao.lblValor.caption = FormatVal(!valor_Dev)
        frmDevolucao.caption = "DEVOLUÇÃO " & !Num_DEV & "-" & !nome_Cli
        frmDevolucao.Show 1
    End With
    P_Status ""
    dtVendas.Refresh
End Sub

Private Sub cmndbtnNew_Click()
       Dim strSQL As String
    
    strSQL = F_GeraNovaDevolucao("00000", strCod_Usu)
    
    
    
    frmDevolucoes.caption = "STYLO PRÓPRIO -Devolução a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmDevolucoes.tag = strSQL
    On Error Resume Next
    frmDevolucoes.Show 1
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
    On Error Resume Next
    If dtVendas.Recordset.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        cmndbtnImpr.Enabled = True
        cmndbtnEdit.Enabled = True
    End If
    If dtVendas.Recordset!situacao_DEV > 1 Then
        cmndbtnEdit.Enabled = False
        
    End If
End Sub

Private Sub Form_Activate()
    
    MudaAparencia Me
    
    P_Busca
End Sub

Private Sub Form_Load()
    'Me.Width = 10725
    'Me.Height = 6135
    P_CentraForm Me
    dtVendas.Connect = Conn.ConnectionString
    
    cmbxSituacao_ven.ListIndex = 0
    cmbxSituacao_ven.Enabled = False
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 1
    
    P_Busca
    bActive = True
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    
    If cmbxBusca.ListIndex = 0 Then
        sBusca = "Nome_Cli"
    ElseIf cmbxBusca.ListIndex = 1 Then
        sBusca = "num_DEV"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sBusca = "data_DEV"
    Else
        sBusca = "data_dEV"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "Nome_Cli"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "num_dev"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "data_Dev"
    Else
        sOrd = "data_dev"
    End If
    
    
    strSQL = "select tc.nome_cli,tv.*,tf.guerra_fun,tf2.guerra_fun as nome "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " tab_devolucoes as tv,tab_funcionarios as tf, tab_funcionarios as tf2 "
    
    strSQL = strSQL & " where " & sBusca & " like '" & txtbxExpr.Text & "%' "
    strSQL = strSQL & " and tf.cod_fun=tv.cod_USU "
    strSQL = strSQL & " and tf2.cod_fun=tv.cod_usu "
    strSQL = strSQL & " and tv.situacao_dev='" & Left(cmbxSituacao_ven.Text, 1) & "' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " order by " & sOrd
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub


Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        P_Busca
        If dtVendas.Recordset.RecordCount > 0 Then
            DBGrid1.SetFocus
        End If
    End If
    If KeyCode = vbKeyDown Then
        If dtVendas.Recordset.RecordCount > 0 Then
            DBGrid1.SetFocus
        End If
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub


