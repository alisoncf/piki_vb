VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiFrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "stt SYS - Topa Tudo  da Construção"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11610
   Icon            =   "mdiFrmEst.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6210
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Picture         =   "mdiFrmEst.frx":030A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "sTT- Super Topa Tudo SYS"
            TextSave        =   "sTT- Super Topa Tudo SYS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6175
            MinWidth        =   6175
            Picture         =   "mdiFrmEst.frx":09DE
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "03/06/2003"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "ex¶ert°¨"
            TextSave        =   "ex¶ert°¨"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":15D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":1EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":278E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":306A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":3946
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":4AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":53DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListPrint 
      Left            =   2640
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   184
      ImageHeight     =   152
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrmEst.frx":5CB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnPrinc 
      Caption         =   "&Principal"
      Begin VB.Menu mnEditCli 
         Caption         =   "&Clientes"
         Shortcut        =   {F2}
      End
      Begin VB.Menu div6 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnForn 
         Caption         =   "&Fornecedores"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnProd 
         Caption         =   "&Produtos"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnCadCat 
         Caption         =   "C&ategorias"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnCadMar 
         Caption         =   "&Marcas"
         Enabled         =   0   'False
      End
      Begin VB.Menu div4 
         Caption         =   "-"
      End
      Begin VB.Menu mnCadCargos 
         Caption         =   "Ca&rgos"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnFun 
         Caption         =   "&Funcionários"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnEdiUsu 
         Caption         =   "&Usuários"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnDiv6 
         Caption         =   "-"
      End
      Begin VB.Menu mnFreteiros 
         Caption         =   "F&reteiros"
         Enabled         =   0   'False
      End
      Begin VB.Menu diva1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair do SGF"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnCadTer 
      Caption         =   "&Compras"
      Begin VB.Menu mnIncMer 
         Caption         =   "&Nova Compra"
      End
      Begin VB.Menu mnEdiComp 
         Caption         =   "&Edição de Compras"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnEntrada 
         Caption         =   "E&ntrada de Produtos"
      End
   End
   Begin VB.Menu mnMov 
      Caption         =   "&Vendas"
      Begin VB.Menu mnNovaVenda 
         Caption         =   "&Nova Venda"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnEdiVen 
         Caption         =   "Edição de Vendas"
         Shortcut        =   {F5}
      End
      Begin VB.Menu div10 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnOrc 
         Caption         =   "&Orçamentos"
      End
      Begin VB.Menu mnPesq 
         Caption         =   "Pesquisa de Preços"
      End
      Begin VB.Menu mnDiv25 
         Caption         =   "-"
      End
      Begin VB.Menu mnPed 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu mnEntr 
         Caption         =   "Entregas"
      End
      Begin VB.Menu mnDEv 
         Caption         =   "&Devoluções"
      End
   End
   Begin VB.Menu mnFinan 
      Caption         =   "&Financeiro"
      Begin VB.Menu mnContasPagar 
         Caption         =   "&Edição de Documentos"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnCalGuias 
         Caption         =   "&Lançamento de Documentos"
      End
      Begin VB.Menu mnCheq 
         Caption         =   "Controle de C&heques"
      End
      Begin VB.Menu div11 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnBan 
         Caption         =   "&Bancos"
      End
      Begin VB.Menu mnCOn 
         Caption         =   "&Contas"
      End
      Begin VB.Menu mnFinanceiras 
         Caption         =   "&Financeiras"
      End
      Begin VB.Menu mnCadPla 
         Caption         =   "&Plano de Receitas e Despesas"
      End
   End
   Begin VB.Menu mnCaixa 
      Caption         =   "Cai&xa"
      Begin VB.Menu mnLanCaixa 
         Caption         =   "&Edição de Lançamentos"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnFechaCaixa 
         Caption         =   "&Lançamentos"
      End
      Begin VB.Menu mnEfetivVen 
         Caption         =   "E&fetivação de Vendas"
      End
   End
   Begin VB.Menu mnRel 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnRelVend 
         Caption         =   "Relatório de &Vendas"
      End
      Begin VB.Menu mnRelComp 
         Caption         =   "Relatório de &Compras"
      End
      Begin VB.Menu mnInvent 
         Caption         =   "&Inventário"
      End
   End
   Begin VB.Menu mnParam 
      Caption         =   "Parâ&metros"
      Begin VB.Menu mnConfig 
         Caption         =   "&Configurações do Sistema"
      End
      Begin VB.Menu mnLojas 
         Caption         =   "&Lojas"
      End
      Begin VB.Menu mnUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnMudSen 
         Caption         =   "&Mudar Senha"
      End
      Begin VB.Menu mnTransf 
         Caption         =   "&Transferência"
      End
      Begin VB.Menu div5 
         Caption         =   "-"
      End
      Begin VB.Menu mnFormasPag 
         Caption         =   "&Formas de Pagamento"
      End
      Begin VB.Menu mnComDes 
         Caption         =   "&Comissões Decrescentes"
      End
      Begin VB.Menu mnApar 
         Caption         =   "&Aparência"
      End
   End
   Begin VB.Menu mnJan 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
      Begin VB.Menu mnCasc 
         Caption         =   "&Cascata"
      End
      Begin VB.Menu mnLado 
         Caption         =   "&Lado a Lado"
      End
      Begin VB.Menu mnTodasJan 
         Caption         =   "&Todas"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "A&juda"
      Begin VB.Menu mnAjuda 
         Caption         =   "&Indices da Ajuda"
      End
      Begin VB.Menu mnCred 
         Caption         =   "&Créditos"
      End
   End
End
Attribute VB_Name = "mdiFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    'TaskBar False
    If F_BuscaUmCampo("cod_loj", "tab_parametros", "") = "" Then
        MsgBox "Os Dados desta Loja não estão definidos corretamente!" & Chr(13) & _
            "Informe esta mensagem ao Responsável!", vbExclamation
        P_Fim
    End If
    
End Sub

Private Sub MDIForm_DblClick()
    frmVenda.Show 0, Me
End Sub

Private Sub MDIForm_Load()
        
    
    
    'frmLogo.Show 0
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnPrinc
    End If
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
    'TaskBar True
    P_Fim
End Sub

Private Sub mnAjuda_Click()
    P_Status "Preparando Ajuda"
    'frmHelp.Show 0
    P_Status ""
End Sub

Private Sub mnApar_Click()
    frmConfig.Show 0
End Sub

Private Sub mnBan_Click()
    frmEdiBan.Show 0
End Sub

Private Sub mnCadDiv_Click()
    frmEdiDiv.Show 0
End Sub

Private Sub mnCadEsp_Click()
    frmEdiMar.Show 0
End Sub

Private Sub mnCadGen_Click()
    frmEdiCat.Show 0
End Sub

Private Sub mnCadCargos_Click()
    frmEdiCargos.Show 0
End Sub

Private Sub mnCadCat_Click()
    frmEdiCat.Show 0
End Sub

Private Sub mnCadMar_Click()
    frmEdiMar.Show 0
End Sub

Private Sub mnCadPla_Click()
    frmEdiPlanoContas.Show 0
    frmEdiPlanoContas.SetFocus
End Sub

Private Sub mnCalGuias_Click()
    frmLanDoc.Show 0
End Sub

Private Sub mnCasc_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnCheq_Click()
    frmControleCheques.Show 0
End Sub

Private Sub mnComDes_Click()
    frmEdiDes.Show 0
End Sub

Private Sub mnCOn_Click()
    frmEdiCon.Show 0
End Sub

Private Sub mnConfig_Click()
    frmParametros.Show 0
End Sub

Private Sub mnContasPagar_Click()
    frmEditaMov.Show 0
End Sub

Private Sub mnCred_Click()
    On Error Resume Next
    frmCred.Show 1
End Sub

Private Sub mnDEv_Click()
    frmDevolucoes.Show 0
    frmDevolucoes.SetFocus
End Sub

Private Sub mnEdiComp_Click()
    frmEdiEntr.Show 0
End Sub

Private Sub mnEditCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub mnEditLan_Click()
     If Not F_SelectCli(frmEditaMov) Then Exit Sub
    frmEditaMov.Show 0
End Sub

Private Sub mnEdiUsu_Click()
    frmEdiUsu.Show 0
End Sub

Private Sub mnEdiVen_Click()
    frmEdiVendas.Show 0
End Sub

Private Sub mnEfetivVen_Click()
    If IsLoad("frmEfetivaVen") Then
        frmEfetivaVen.SetFocus
    Else
        frmEfetivaVen.Show 0
    End If
End Sub

Private Sub mnEntr_Click()
    frmEntrega.Show 0
End Sub

Private Sub mnEntrada_Click()
    frmNotaEntrada.Show 1
End Sub

Private Sub mnFechaCaixa_Click()
    Dim fr As New frmLanCaixa
    fr.Show 0
End Sub

Private Sub mnFinanceiras_Click()
    frmEdiFinan.Show 0
End Sub

Private Sub mnFormasPag_Click()
    frmEdiFormasPag.Show 0
End Sub

Private Sub mnForn_Click()
    frmEdiFor.Show 0
End Sub

Private Sub mnIncCli_Click()
    frmIncClientes.Show 0
End Sub

Private Sub mnIncTer_Click()
    frmIncTer.Show 0
End Sub



Private Sub mnLivIcms_Click()
    frmLivIcms.Show 0
End Sub

Private Sub mnLivIpi_Click()
    frmLivIpi.Show 0
End Sub

Private Sub mnMudaCli_Click()
    On Error Resume Next
    frmSelectCliente.Show 1
End Sub

Private Sub mnPer_Click()
    frmPerfis.Show 0
End Sub

Private Sub mnFreteiros_Click()
    frmEdiFreteiros.Show 0
    frmEdiFreteiros.SetFocus
End Sub

Private Sub mnFun_Click()
    frmEdiFun.Show 0
End Sub

Private Sub mnIncMer_Click()
    Dim fr As New frmEntradas
    fr.Show 0
End Sub

Private Sub mnInvent_Click()
    frmInventario.Show 0
End Sub

Private Sub mnLado_Click()
    Me.Arrange vbVertical
End Sub

Private Sub mnLanCaixa_Click()
    frmEdiCai.Show 0
End Sub

Private Sub mnLojas_Click()
    frmEdiLoj.Show 0
End Sub

Private Sub mnMudSen_Click()
    On Error Resume Next
    frmMudaSenha.Show 1
End Sub

Private Sub mnNovaVenda_Click()
    On Error Resume Next
    frmAbreVenda.Show 1
End Sub

Private Sub mnPed_Click()
    frmEdiPed.Show 0
End Sub

Private Sub mnProd_Click()
    frmEdiPro.Show 0
End Sub

Private Sub mnRelGuias_Click()
    frmRelGuias.Show 0
End Sub

Private Sub mnRelVend_Click()
  
    frmRelVendas.Tag = "S"
    frmRelVendas.Show 0
    
End Sub

Private Sub mnTer_Click()
    If strCOD_CLI = "" Then
        If MsgBox("Não há cliente selecionado." & Chr(13) & "Deseja selecionar um agora?" _
            , vbYesNo + vbQuestion) = vbYes Then
            On Error Resume Next
            frmSelectCliente.Show 1
            If strCOD_CLI = "" Then
                MsgBox "Não foi selecionado nehum cliente", vbExclamation
            Else
                frmSelectTer.Show
            End If
            
        Else
            MsgBox "Não posso manipular terceiros sem que você escolha um cliente", vbCritical
            Exit Sub
        End If
    Else
        frmSelectTer.Show 0
    End If
    
End Sub

Private Sub mnSair_Click()
    If MsgBox("Deseja Realmente Sair?", vbYesNo + vbQuestion) = vbYes Then
        P_Fim
    End If
End Sub

Private Sub mnServicos_Click()
    frmEdiServ.Show 0
End Sub

Private Sub mnTodasJan_Click()
    'On Error Resume Next
   
   If Forms.count > 1 Then
        frmForms.Show , mdiFrm
    End If
   
End Sub

Private Sub mnTransf_Click()
    On Error Resume Next
    frmTransf.Show 1
End Sub

Private Sub mnUsu_Click()
    Dim fr As New frmIncUsu
    fr.Show 0
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
     If Panel.Index = 1 Then
        mnMudaCli_Click
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim iCount As Integer
    If Button.Index = 1 Then
        mnEditCli_Click
    ElseIf Button.Index = 2 Then
        mnForn_Click
    ElseIf Button.Index = 3 Then
        mnProd_Click
    ElseIf Button.Index = 4 Then
        mnEdiVen_Click
    ElseIf Button.Index = 5 Then
        mnNovaVenda_Click
    ElseIf Button.Index = 6 Then
        mnLanCaixa_Click
    ElseIf Button.Index = 7 Then
        mnEdiComp_Click
    ElseIf Button.Index = 8 Then
        mnServicos_Click
    ElseIf Button.Index = 9 Then
        If Forms.count = 1 Then
            mnSair_Click
        Else
            'For iCount = 1 To Forms.count - 1
                
               ' MsgBox Forms(iCount).Name
                 Unload Screen.ActiveForm
            'Next
        End If
    End If
    
End Sub



