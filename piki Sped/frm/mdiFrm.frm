VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiFrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00404000&
   Caption         =   "Hiper Útil - O Show Room do Seu Lar..."
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   Icon            =   "mdiFrm.frx":0000
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
            Picture         =   "mdiFrm.frx":068A
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
            Picture         =   "mdiFrm.frx":0D5E
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "23/06/2003"
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
      Left            =   8280
      Top             =   1920
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
            Picture         =   "mdiFrm.frx":107A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":1956
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":2232
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":2B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":33EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":3CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":45A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":4E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":575A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListPrint 
      Left            =   7320
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   438
      ImageHeight     =   318
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm.frx":6036
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   1191
      ButtonWidth     =   609
      ButtonHeight    =   1032
      Appearance      =   1
      _Version        =   393216
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11955
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   630
            Left            =   0
            ScaleHeight     =   630
            ScaleWidth      =   1350
            TabIndex        =   12
            Top             =   0
            Width           =   1350
            Begin VB.Image Image1 
               Height          =   570
               Left            =   120
               Picture         =   "mdiFrm.frx":B26A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1170
            End
         End
         Begin VB.CommandButton cmndbtnProd 
            Caption         =   "&Produtos"
            Height          =   615
            Left            =   1440
            TabIndex        =   3
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnECF 
            Caption         =   "Controle EC&F"
            Height          =   615
            Left            =   2520
            TabIndex        =   4
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnVendas 
            Caption         =   "&Vendas"
            Height          =   615
            Left            =   3600
            TabIndex        =   5
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnNovaVenda 
            Caption         =   "&Nova Venda"
            Default         =   -1  'True
            Height          =   615
            Left            =   4680
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnCaixa 
            Caption         =   "Cai&xa"
            Enabled         =   0   'False
            Height          =   615
            Left            =   5760
            TabIndex        =   7
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnFechar 
            Caption         =   "Fec&hamento"
            Enabled         =   0   'False
            Height          =   615
            Left            =   6840
            TabIndex        =   8
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnUsu 
            Caption         =   "Usuá&rio"
            Height          =   615
            Left            =   7920
            TabIndex        =   9
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnJan 
            Caption         =   "&Janelas"
            Height          =   615
            Left            =   9000
            TabIndex        =   10
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnSair 
            Caption         =   "&Sair"
            Height          =   615
            Left            =   10080
            TabIndex        =   11
            Top             =   0
            Width           =   1095
         End
      End
   End
   Begin VB.Menu mnPrinc 
      Caption         =   "&Principal"
      Visible         =   0   'False
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
      End
      Begin VB.Menu mnCadMar 
         Caption         =   "&Marcas"
      End
      Begin VB.Menu div4 
         Caption         =   "-"
      End
      Begin VB.Menu mnCadCargos 
         Caption         =   "Ca&rgos"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnFun 
         Caption         =   "&Funcionários"
      End
      Begin VB.Menu mnEdiUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnDiv6 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair do SGF"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnCadTer 
      Caption         =   "&Compras"
      Visible         =   0   'False
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
   Begin VB.Menu mnAj 
      Caption         =   "A&justes"
      Visible         =   0   'False
      Begin VB.Menu mnEstIni 
         Caption         =   "&Estoque Inicial"
      End
      Begin VB.Menu mnSaiDia 
         Caption         =   "&Saida Diária"
      End
      Begin VB.Menu mnPisos 
         Caption         =   "&Pisos"
      End
   End
   Begin VB.Menu mnMov 
      Caption         =   "&Vendas"
      Visible         =   0   'False
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
      Begin VB.Menu mnDEv 
         Caption         =   "&Devoluções"
      End
   End
   Begin VB.Menu mnFinan 
      Caption         =   "&Financeiro"
      Visible         =   0   'False
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
      Visible         =   0   'False
      Begin VB.Menu mnLanCaixa 
         Caption         =   "&Edição de Lançamentos"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnFechaCaixa 
         Caption         =   "&Lançamentos"
      End
      Begin VB.Menu mnCOntEcf 
         Caption         =   "Controle ECF"
      End
      Begin VB.Menu mnEfetivVen 
         Caption         =   "E&fetivação de Vendas"
      End
   End
   Begin VB.Menu mnRel 
      Caption         =   "&Relatórios"
      Visible         =   0   'False
      Begin VB.Menu mnRelVend 
         Caption         =   "Relatório de &Vendas"
      End
      Begin VB.Menu mnRelComp 
         Caption         =   "Relatório de &Compras"
      End
      Begin VB.Menu mnInvent 
         Caption         =   "&Inventário de Conf."
      End
   End
   Begin VB.Menu mnParam 
      Caption         =   "Parâ&metros"
      Visible         =   0   'False
      Begin VB.Menu mnConfig 
         Caption         =   "&Configurações do Sistema"
      End
      Begin VB.Menu mnUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnMudSen 
         Caption         =   "&Mudar Senha"
      End
      Begin VB.Menu div5 
         Caption         =   "-"
      End
      Begin VB.Menu mnFormasPag 
         Caption         =   "&Formas de Pagamento"
      End
      Begin VB.Menu mnApar 
         Caption         =   "&Aparência"
      End
   End
   Begin VB.Menu mnJan 
      Caption         =   "&Janelas"
      Visible         =   0   'False
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
      Visible         =   0   'False
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

Private Sub cmndbtnECF_Click()
    On Error Resume Next
    frmImpressoraECF.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmndbtnJan_Click()
    Me.PopupMenu mnJan
End Sub

Private Sub cmndbtnNovaVenda_Click()
    mnNovaVenda_Click
End Sub

Private Sub cmndbtnSair_Click()
    mnSair_Click
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0, mdiFrm
End Sub

Sub mnCOntEcf_Click()
    frmImpressoraECF.Show 1
End Sub

Private Sub cmndbtnCLi_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub cmndbtnDev_Click()
    'frmDevolucoes.Show 0
End Sub

Private Sub cmndbtnNewVEnda_Click()
    On Error Resume Next
    frmVenda.Show 1
End Sub

Private Sub cmndbtnPEd_Click()
    frmEdiPed.Show 0
End Sub

Private Sub cmndbtnProd_Click()
    frmEdiPro.Show 0
    frmEdiPro.SetFocus
End Sub

Private Sub cmndbtnVendas_Click()
    frmEdiVendas.Show 0
    frmEdiVendas.SetFocus
End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    'TaskBar False
   ' If F_BuscaUmCampo("cod_loj", "tab_parametros", "") = "" Then
    '    MsgBox "Os Dados desta Loja não estão definidos corretamente!" & Chr(13) & _
     '       "Informe esta mensagem ao Responsável!", vbExclamation
      '  P_Fim
        'End If
    
End Sub

Private Sub MDIForm_DblClick()
    'frmVenda.Show 0, Me
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

Private Sub mnEntrada_Click()
    frmNotaEntrada.Show 1
End Sub

Private Sub mnEntrPed_Click()
    frmEfetivaPed.Show 0
End Sub

Private Sub mnEntrVen_Click()
    frmEntrega.Show 0
End Sub

Private Sub mnEstIni_Click()
    On Error Resume Next
    frmAjEstIni.Show 1
End Sub

Private Sub mnEti_Click()
    P_ImprxEstiquetas frmVisualiza.Pic
    frmVisualiza.Show 1
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
    Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = "HIPER ÚTIL -Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.Tag = strSQL
    On Error Resume Next
    frmVenda.Show 1
    
End Sub

Private Sub mnPed_Click()
    frmEdiPed.Show 0
End Sub

Private Sub mnPisos_Click()
    frmEstoqueIni.Show 1
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

Private Sub mnSaiDia_Click()
    On Error Resume Next
    frmSaidaDiaria.Show 1
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



