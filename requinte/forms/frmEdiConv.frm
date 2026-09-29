VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiConv 
   Caption         =   "Edição de Convênios"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   Icon            =   "frmEdiConv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   10935
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   10905
      TabIndex        =   7
      Top             =   0
      Width           =   10935
      Begin VB.TextBox txtbxExpr 
         Height          =   315
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiConv.frx":1EFA
         Left            =   240
         List            =   "frmEdiConv.frx":1F0D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiConv.frx":1F35
         Left            =   1800
         List            =   "frmEdiConv.frx":1F48
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   3480
         TabIndex        =   13
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox pcBotoes 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6420
      Left            =   9360
      ScaleHeight     =   6390
      ScaleWidth      =   1545
      TabIndex        =   0
      Top             =   615
      Width           =   1575
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "&Débitos"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConv.frx":1F78
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnConv 
         Caption         =   "Con&veniados"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConv.frx":20C2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConv.frx":2177
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConv.frx":2A41
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConv.frx":2B8B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiConv.frx":2CD5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Data dtTel 
      Caption         =   "Edição de Gêneros"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Data dtClientes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiConv.frx":305F
      Height          =   3975
      Left            =   0
      OleObjectBlob   =   "frmEdiConv.frx":3078
      TabIndex        =   15
      Top             =   720
      Width           =   9615
   End
   Begin MSDBGrid.DBGrid dbgrdFones 
      Bindings        =   "frmEdiConv.frx":446D
      Height          =   1215
      Left            =   0
      OleObjectBlob   =   "frmEdiConv.frx":4481
      TabIndex        =   16
      Top             =   4800
      Width           =   9255
   End
End
Attribute VB_Name = "frmEdiConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DBGrid1_Click()

End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnCard_Click()
    'P_Cartao dtClientes.Recordset!cod_cli
End Sub

Private Sub cmndbtnDeb_Click()
    On Error GoTo trata_Erro
    frmParcelasTer.txtbxNomeCli.Text = dtClientes.Recordset!nome_Cli
    frmParcelasTer.txtbxNomeCli.tag = dtClientes.Recordset!cod_cli
    frmParcelasTer.tag = dtClientes.Recordset!cod_cli
    frmParcelasTer.Show 0
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadConv
    If dtClientes.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    strHora = Time
    fr.tag = dtClientes.Recordset!cod_cli
    fr.txtbxCliente.tag = dtClientes.Recordset!cod_cli
    fr.txtbxCliente.Text = dtClientes.Recordset!nome_Cli
    fr.txtbxDesc.Text = dtClientes.Recordset!desconto_con
    P_LocalizaOperacao dtClientes.Recordset!status_con, fr.cmbxStatus
    fr.Show 0
End Sub

Private Sub cmndbtnHist_Click()
'    On Error GoTo trata_erro
    With dtClientes.Recordset
        If .RecordCount = 0 Then Exit Sub
        'P_HistCli !cod_cli
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    On Error GoTo trata_Erro
    strSQL = "select tc.*,tco.desconto_con,tco.status_con from tab_clientes as tc,tab_convenios as tco where "
    strSQL = strSQL & " tc.cod_cli=tco.cod_cli and "
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & "Nome_cli like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & "cod_cli like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & "cpf_cli like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & "rg_cli like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 4 Then
        strSQL = "select tc.* from tab_clientes as tc, tab_autorizados as ta"
        strSQL = strSQL & " where ta.cod_cli=tc.cod_cli "
        strSQL = strSQL & " and nome_aut like '" & txtbxExpr.Text & "%' "
    End If
    strSQL = strSQL & " and tc.cod_cli <> 0 " 'venda a consumidor
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "tc.cod_cli"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_Cli,cpf_Cli"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "CPF_CLI"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "Cidade_Cli,uf_Cli,NOme_Cli"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "UF_CLI,Cidade_CLi"
    End If
    strSQL = strSQL & " order by " & sOrderBy
    dtClientes.RecordSource = strSQL
    dtClientes.Refresh
    If dtClientes.Recordset.RecordCount = 0 Then
        'MsgBox "Não foi encontrado nenhum cliente na sua busca!", vbExclamation
    End If
    dbgrd.caption = "Cadastrados...: " & F_BuscaUmCampo("count(cod_Cli)", "tab_clientes", "")
    dbgrd.caption = dbgrd.caption & " - Listados...: " & dtClientes.Recordset.RecordCount
trata_Erro:
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadConv
    fr.Show 0
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmndbtnPremiacao_Click()
    On Error GoTo trata_Erro
    'P_PremiacoesCli dtClientes.Recordset!cod_cli
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnPromocoes_Click()
    On Error GoTo trata_Erro
    'P_PontosCli dtClientes.Recordset!cod_cli
    Exit Sub
trata_Erro:
End Sub

Private Sub cmndbtnRes_Click()
    frmReserva.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim c As Clientes.Cliente
    c = Clientes.F_BuscaCliente("")
    
    MsgBox c.t(0).num_tel & c.t(1).num_tel
    
    
    
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dtClientes_Reposition()
    Dim strSQL As String
    On Error Resume Next
    With dtClientes.Recordset
     
        strSQL = "select * from tab_telefones where cod_cli='" & !cod_cli & "'"
        dtTel.RecordSource = strSQL
        dtTel.Refresh
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    txtbxExpr.SetFocus
End Sub

Private Sub Form_GotFocus()
    SendKeys Chr(9)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    resizeForm Me
    dtClientes.Connect = Banco.Connect
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    cmndbtnLocaliza_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Width = pcBotoes.Left - 100
    
    dbgrdAut.Top = Me.Height - dbgrdAut.Height - 500
    dbgrdFones.Top = dbgrdAut.Top
    dbgrd.Height = Me.Height - dbgrd.Top - dbgrdAut.Height - 500
    dbgrdFones.Width = dbgrd.Width / 2
    
End Sub

Private Sub pcBotoes_GotFocus()
5    SendKeys Chr(9)
End Sub

Private Sub txtbxExpr_Change()
    cmndbtnLocaliza_Click
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
    If KeyCode = vbKeyDown Then
        dbgrd.SetFocus
    End If
    
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    If cmbxBusca.Text = "CNPJ" Then
       
    End If
End Sub

