VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiVendas 
   Caption         =   "Edição de Vendas"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   Icon            =   "frmEdiVendas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10035
   Begin VB.Data dtVendas 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5880
      Left            =   0
      ScaleHeight     =   5880
      ScaleWidth      =   1935
      TabIndex        =   23
      Top             =   735
      Width           =   1935
      Begin VB.CommandButton Command1 
         Caption         =   "&Reabrir Venda"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":110C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":1496
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":15E0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Nova"
         Enabled         =   0   'False
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":172A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnVendasEfet 
         Caption         =   "&Vendas a Efetivar"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":1AB4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSepd 
         Caption         =   "&SEPD"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":1BFE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar a Venda"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnReem 
         Caption         =   "&Reemitir CUPOM"
         Height          =   585
         Left            =   120
         Picture         =   "frmEdiVendas.frx":1E44
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3120
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10005
      TabIndex        =   22
      Top             =   0
      Width           =   10035
      Begin VB.ComboBox cmbxSituacao_ven 
         Height          =   315
         ItemData        =   "frmEdiVendas.frx":1F8E
         Left            =   3480
         List            =   "frmEdiVendas.frx":1F98
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiVendas.frx":1FB7
         Left            =   1800
         List            =   "frmEdiVendas.frx":1FC7
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxExpr 
         Height          =   285
         Left            =   5160
         MaxLength       =   40
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiVendas.frx":1FF2
         Left            =   120
         List            =   "frmEdiVendas.frx":2002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   255
         Left            =   9360
         Picture         =   "frmEdiVendas.frx":202D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtbxDe 
         Height          =   285
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Height          =   285
         Left            =   8280
         MaxLength       =   40
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação da Venda:"
         Height          =   195
         Left            =   3480
         TabIndex        =   4
         Top             =   120
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   7200
         TabIndex        =   8
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   8280
         TabIndex        =   10
         Top             =   120
         Width           =   285
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmEdiVendas.frx":2126
      Height          =   5415
      Left            =   2040
      OleObjectBlob   =   "frmEdiVendas.frx":213D
      TabIndex        =   21
      Top             =   840
      Width           =   7575
   End
End
Attribute VB_Name = "frmEdiVendas"
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

Private Sub cmndbtnBuscar_Click()
    P_Busca
End Sub

Private Sub cmndbtnCancel_Click()
    If strExclVenda_usu = "S" Then
        If MsgBox("Deseja realmente cancelar todos os lancamentos dessa venda?", vbYesNo + vbQuestion) = vbYes Then
            
            
            If dtVendas.Recordset!situacao_ven <> "4" Then
                MsgBox "A venda não pode ser cancelada por aqui!", vbCritical
            End If
            If Not F_CancelaVendaAposCupom(dtVendas.Recordset!num_ven) Then
            End If
        End If
    Else
        MsgBox "Você não tem permissão para isso!", vbCritical
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmVenda
    P_Status "Montando venda..."
    With dtVendas.Recordset
        fr.Tag = !num_ven
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        fr.Atacado = True
        fr.lblValor.caption = FormatVal(!valor_ven)
        fr.caption = "Venda " & !num_ven & "-" & !nome_cli
        fr.txtbxNomeCli.Text = Format(!nomecli_ven)
        fr.txtbxNomeCli.Tag = Format(!cod_cli, "0000#")
        fr.txtbxCod_Cli.Text = Format(!cod_cli)
        fr.operacao = "B"
        fr.pcECF.BackColor = vbYellow
        'frmVenda.txtbxObs.Text = Format(!obs_ven)
        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
        fr.Show 0
        fr.lblAtacado.Tag = "X" 'indica que a venda foi editada
    End With
    P_Status ""
    dtVendas.Refresh
End Sub

Private Sub cmndbtnImpr_Click()
    P_Status "Gerando impressão... "
    P_ImprimeVenda frmVisualiza.Pic, dtVendas.Recordset!num_ven, False
    P_Status ""
    frmVisualiza.Show 1
    If frmVisualiza.Tag <> "" Then
        
        P_ImprimeVenda Printer, dtVendas.Recordset!num_ven, True
               
    End If
    Unload frmVisualiza
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
   Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = "Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.Tag = strSQL
    On Error Resume Next
    frmVenda.Show 0, mdiFrm
End Sub

Private Sub cmndbtnReem_Click()
    frmReemissao.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSepd_Click()
    On Error GoTo trata_erro
    P_notaSEPD dtVendas.Recordset!num_ven, ""
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnVendasEfet_Click()
    frmEfetivaVen.Show 0
End Sub

Private Sub Command1_Click()
    If strExclVenda_usu = "S" Then
        If MsgBox("Deseja realmente cancelar todos os lancamentos dessa venda?", vbYesNo + vbQuestion) = vbYes Then
            
            
            If dtVendas.Recordset!situacao_ven <> "2" Then
                MsgBox "A venda não pode ser reaberta!", vbCritical
            End If
            strSQL = "update tab_vendas set situacao_ven='1' where num_ven='" & dtVendas.Recordset!num_ven & "' "
            Banco.Execute strSQL
            dtVendas.Refresh
        End If
    Else
        MsgBox "Você não tem permissão para isso!", vbCritical
    End If
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
    If dtVendas.Recordset!situacao_ven > 1 Then
        cmndbtnEdit.Enabled = False
        
    End If
End Sub

Private Sub Form_Activate()
    
    
    MudaAparencia Me
    
    P_Busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        If cmbxSituacao_ven.ListCount <= 2 Then
            cmbxSituacao_ven.AddItem "3- Canceladas"
            
            cmbxSituacao_ven.AddItem "4- Efetivadas"
        End If
    End If
End Sub

Private Sub Form_Load()
    'Me.Width = 10725
    'Me.Height = 6135
    resizeForm Me
    dtVendas.Connect = Banco.Connect
    
    
    cmbxSituacao_ven.ListIndex = 0
    If strAltprod_usu <> "S" Then
        cmbxSituacao_ven.Enabled = False
    End If
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
        sBusca = "num_ven"
    ElseIf cmbxBusca.ListIndex = 2 Then
        sBusca = "data_ven"
    Else
        sBusca = "data_ven"
    End If
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "Nome_Cli"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "num_ven"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "data_ven"
    Else
        sOrd = "dataefet_ven"
    End If
    
    
    strSQL = "select tc.nome_cli,tv.*,tf.guerra_fun,tf2.guerra_fun as nome "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " tab_vendas as tv,tab_funcionarios as tf, tab_funcionarios as tf2 "
    If sBusca <> "dataefet_ven" And sBusca <> "data_ven" Then
        strSQL = strSQL & " where " & sBusca & " like '" & txtbxExpr.Text & "%' "
    Else
        strSQL = strSQL & " where " & sBusca & " >= " & F_GravaData(txtbxExpr.Text) & " "
    End If
    strSQL = strSQL & " and tf.cod_fun=tv.cod_fun "
    strSQL = strSQL & " and tf2.cod_fun=tv.cod_usu "
    strSQL = strSQL & " and tv.situacao_ven='" & Left(cmbxSituacao_ven.Text, 1) & "' "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
     If UCase(App.Title) = "PIKIBOX_CX" Then
        strSQL = strSQL & " and tv.cod_cai='" & strCod_Cai & "' "
      
    End If
    If txtbxDe.Text <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaData(txtbxDe.Text)
    End If
    If txtbxAte.Text <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaData(txtbxAte.Text)
    End If
        
  
        
    strSQL = strSQL & " order by " & sOrd & " desc "
    dtVendas.RecordSource = strSQL
    dtVendas.Refresh
End Sub


Private Sub Form_Resize()
    DBGrid1.Width = Me.Width - DBGrid1.Left - 100
    DBGrid1.Height = Me.Height - DBGrid1.Top - 500
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxDe_GotFocus()
    P_SelecTexto txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
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
