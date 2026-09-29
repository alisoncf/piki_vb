VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiFrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00755C33&
   Caption         =   "Piki Box 1.0 - Sistema de Controle de Supermercados - Caixa"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11865
   Icon            =   "mdiFrm_finan.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcMenu 
      Align           =   3  'Align Left
      Height          =   9615
      Left            =   0
      ScaleHeight     =   9555
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   0
      Width           =   1560
      Begin VB.PictureBox pcCalend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   1425
         TabIndex        =   12
         Top             =   8160
         Width           =   1455
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   600
            Picture         =   "mdiFrm_finan.frx":0CCA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   13
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblDia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   420
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblNumDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   825
            Left            =   540
            TabIndex        =   16
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblMes 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   495
            TabIndex        =   15
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label lblAno 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2007"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   390
            TabIndex        =   14
            Top             =   1440
            Width           =   675
         End
      End
      Begin VB.CommandButton cmndbtnDocs 
         Caption         =   "&Documentos"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":1002
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1540
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnUsu 
         Caption         =   "&Usuário"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":1A04
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6080
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnCompras 
         Caption         =   "&Novo Doc."
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":22CE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2300
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnUsu1 
         Caption         =   "&Fechamento"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":2858
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5320
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnCaixa 
         Caption         =   "&Mov. Contas"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":29B1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4565
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnNovaVenda 
         Caption         =   "&Cheques"
         Default         =   -1  'True
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":327B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3815
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnVendas 
         Caption         =   "&Monitoramento"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":3805
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3065
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnCli 
         Caption         =   "&Clientes"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":4207
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   770
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnProd 
         Caption         =   "&Fornecedores"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":4ED1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1500
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Sair"
         Height          =   735
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Picture         =   "mdiFrm_finan.frx":6BCB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6820
         Width           =   1500
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":6FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":7026
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":7084
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":70E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":7140
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":719E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":71FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":725A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":72B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":7316
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":7374
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":73D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":7430
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":748E
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
      ImageWidth      =   57
      ImageHeight     =   72
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_finan.frx":74EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   9615
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5116
            MinWidth        =   5116
            Text            =   "Piki - Xpert Sol. Proc. Dados Ltda"
            TextSave        =   "Piki - Xpert Sol. Proc. Dados Ltda"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6704
            MinWidth        =   6704
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2116
            MinWidth        =   2116
            TextSave        =   "9/2/2007"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   2858
            MinWidth        =   2858
            Text            =   "Xpert Softwares"
            TextSave        =   "Xpert Softwares"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnPrinc 
      Caption         =   "&Principal"
      Begin VB.Menu mnEditCli 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu div6 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnForn 
         Caption         =   "&Fornecedores"
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnCadCargos 
         Caption         =   "Ca&rgos"
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
         Caption         =   "&Sair do Piki Box"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnCadTer 
      Caption         =   "&Entradas"
      Begin VB.Menu mnIncMer 
         Caption         =   "&Nova Compra"
      End
      Begin VB.Menu mnEdiComp 
         Caption         =   "&Edição de Compras"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnEntrada 
         Caption         =   "E&ntrada de Produtos"
         Visible         =   0   'False
      End
      Begin VB.Menu div10 
         Caption         =   "-"
      End
      Begin VB.Menu mnEdiDev 
         Caption         =   "&Devoluções"
         Visible         =   0   'False
      End
      Begin VB.Menu mnDEv 
         Caption         =   "Nova Dev&olução"
      End
      Begin VB.Menu mnDivVen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnProdProm 
         Caption         =   "&Produtos em Promoção"
      End
      Begin VB.Menu mnDivVen3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnFinan 
      Caption         =   "&Financeiro"
      Begin VB.Menu mnMovFin 
         Caption         =   "&Movimento Financeiro"
      End
      Begin VB.Menu mnContasPagar 
         Caption         =   "&Edição de Documentos"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnCalGuias 
         Caption         =   "&Lançamento de Documentos"
      End
      Begin VB.Menu mnEdiMovFin 
         Caption         =   "Movimento de &Financeiras"
      End
      Begin VB.Menu mnCheq 
         Caption         =   "Controle de C&heques"
      End
      Begin VB.Menu mnCons 
         Caption         =   "&Consolidação Diária Financeira"
      End
      Begin VB.Menu div11 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnBan 
         Caption         =   "&Bancos"
      End
      Begin VB.Menu mnCon 
         Caption         =   "&Contas"
      End
      Begin VB.Menu mnFinanceiras 
         Caption         =   "&Financeiras"
      End
      Begin VB.Menu mnCadPla 
         Caption         =   "&Plano de Receitas e Despesas"
      End
      Begin VB.Menu mnBol 
         Caption         =   "&Boletos"
      End
      Begin VB.Menu mnFormasPag 
         Caption         =   "&Formas de Pagamento"
      End
   End
   Begin VB.Menu mnCaixa 
      Caption         =   "Controle de Cai&xas"
      Begin VB.Menu mnLanCaixa 
         Caption         =   "&Edição de Lançamentos"
      End
      Begin VB.Menu mnMonit 
         Caption         =   "&Monitoramento de Caixas"
      End
      Begin VB.Menu mnTurnos 
         Caption         =   "Edição de &Turnos"
      End
      Begin VB.Menu mnEdiCaixas 
         Caption         =   "Edição de &Caixas"
      End
      Begin VB.Menu mnMovUsu 
         Caption         =   "&Movimentação por Usuário"
      End
      Begin VB.Menu mnFecTur 
         Caption         =   "&Fechamento de Turno"
      End
      Begin VB.Menu mnEfetVen 
         Caption         =   "&Efetivação de Vendas"
      End
   End
   Begin VB.Menu mnRel 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnRelVend 
         Caption         =   "Relatório de &Vendas"
      End
      Begin VB.Menu mnRelDev 
         Caption         =   "&Relatório de Devoluções"
      End
      Begin VB.Menu mnRelProdForn 
         Caption         =   "&Relatório de Produtos por Fornecedor"
      End
      Begin VB.Menu mnRelCmpFor 
         Caption         =   "Relatório de &Compras por Fornecedor"
      End
      Begin VB.Menu mnInvent 
         Caption         =   "&Inventário de Conferência"
      End
      Begin VB.Menu mnConsPre 
         Caption         =   "Consulta de Pr&eços"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnRelProdLucr 
         Caption         =   "Realtório de Produtos por &Lucratividade"
      End
      Begin VB.Menu mnRelMMD 
         Caption         =   "Relatório de &Media Movimentação Diária"
      End
      Begin VB.Menu div12 
         Caption         =   "-"
      End
      Begin VB.Menu mnHistCheCPF 
         Caption         =   "&Histórico de Cheques por CPF"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnRelCOmiss 
         Caption         =   "Relatório de &Comissões"
      End
      Begin VB.Menu mnRelFatVenda 
         Caption         =   "Relatório de &Faturamento por Vendas"
      End
      Begin VB.Menu mnRelAnaFat 
         Caption         =   "&Análise de Faturamento"
      End
      Begin VB.Menu mnRelGiro 
         Caption         =   "Relatório de Giro Periódico"
      End
      Begin VB.Menu mnCompProCli 
         Caption         =   "Relatório de Comp. de Prod. p/ C&liente"
      End
      Begin VB.Menu mnCOmport 
         Caption         =   "Relatório de &Comportamento de Produtos"
      End
      Begin VB.Menu div13 
         Caption         =   "-"
      End
      Begin VB.Menu mnRelPagReceb 
         Caption         =   "Relatório de Contas a Pagar/Receber"
      End
      Begin VB.Menu mnRelChq 
         Caption         =   "Relatório de Cheques a Pagar/Receber"
      End
      Begin VB.Menu mnRelAtvCli 
         Caption         =   "Relatório de &Atividade de Clientes por Vendas"
      End
      Begin VB.Menu mNRelAtvPro 
         Caption         =   "Relatório de &Atividade de Clientes por Produtos"
      End
   End
   Begin VB.Menu mnParam 
      Caption         =   "Parâ&metros"
      Begin VB.Menu mnAdm 
         Caption         =   "&Aministração do Sistema"
         Enabled         =   0   'False
         Begin VB.Menu mnSQLPlu 
            Caption         =   "&SQL Plus - Console SQL"
         End
         Begin VB.Menu mnBAckUp 
            Caption         =   "&BackUp"
         End
         Begin VB.Menu mnTransf 
            Caption         =   "&Utilitário de Transferência de Dados"
         End
         Begin VB.Menu mnMigra 
            Caption         =   "&Utilitário de Migração"
         End
      End
      Begin VB.Menu mnConfig 
         Caption         =   "&Configurações do Sistema"
      End
      Begin VB.Menu mnUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnProm 
         Caption         =   "&Promoções"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnMudSen 
         Caption         =   "&Mudar Senha"
      End
      Begin VB.Menu div5 
         Caption         =   "-"
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
Dim bActive As Boolean
Private Sub mnCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub cmndbtnCaixa_Click()
    frmEdiFin.Show 0
End Sub

Private Sub cmndbtnCompras_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub

Private Sub cmndbtnECF_Click()
    On Error Resume Next
    frmImpressoraECF.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmndbtnFechar_Click()

End Sub

Private Sub cmndbtnJan_Click()
    frmConsPreco.Show 0
    frmConsPreco.SetFocus
'    Me.PopupMenu mnJan
End Sub

Private Sub cmndbtnDocs_Click()
    mnContasPagar_Click
End Sub

Private Sub cmndbtnNovaVenda_Click()
    frmControleCheques.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    mnSair_Click
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0, mdiFrm
End Sub

Private Sub mnAjEstRes_Click()
    P_AjustaEstoqueReservado
End Sub

Private Sub mnArr_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiArr.Show 0
    frmEdiArr.SetFocus
End Sub

Private Sub mnAbreCx_Click()
    frmAberturaCaixa.Show 0, mdiFrm
End Sub

Private Sub cmndbtnUsu1_Click()
    mnFecTur_Click
End Sub

Private Sub mnBol_Click()
    frmBoletos.Show 0
End Sub

Private Sub mnBuscaOrc_Click()
    frmEdiOrc.Show 0
    frmEdiOrc.SetFocus
End Sub

Private Sub mnCalMed_Click()
    If strViscusto_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para esta operação!", vbExclamation
        Exit Sub
        
    End If
    frmMediaConsumo_Mer.Show 1
End Sub

Private Sub mnCOmport_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmComport.Show 0
End Sub

Private Sub mnCompProCli_Click()
    frmComportProCli.Show 0
    frmComportProCli.SetFocus
End Sub

Private Sub mnCons_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmFechaFinan.Show 1
    End If
End Sub

Private Sub mnConsPre_Click()
    frmConsPreco.Show 0
    frmConsPreco.SetFocus
End Sub

Private Sub mnCont_Click()
    frmContagens.Show 0
    frmContagens.SetFocus
End Sub

Sub mnCOntEcf_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmImpressoraECF.Show 1
End Sub

Private Sub cmndbtnCli_Click()
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
    frmEdiFor.Show 0
    frmEdiFor.SetFocus
End Sub

Private Sub cmndbtnVendas_Click()
    If strFechames_usu = "S" Then
        frmMonitoramentoCaixas.Show 0, mdiFrm
        frmMonitoramentoCaixas.SetFocus
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
End Sub

Private Sub MDIForm_Activate()
     On Error Resume Next
    Me.caption = constNome_sis & " " & constVersao_Sis & " - " & strNome_Emp & " - " & strCidade_Emp & "/" & strUF_Emp & " .:. CNPJ.: " & strCNPJ_Emp
    'TaskBar False
   ' If F_BuscaUmCampo("cod_loj", "tab_parametros", "") = "" Then
    '    MsgBox "Os Dados desta Loja não estão definidos corretamente!" & Chr(13) & _
     '       "Informe esta mensagem ao Responsável!", vbExclamation
      '  P_Fim
        'End If
    StatusBar1.Panels(6).Text = Winsock1.LocalIP & "-" & Winsock1.LocalHostName
    strIP_Usu = Winsock1.LocalIP
    strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixas", "ip_cai='" & strIP_Usu & "'")
    If strCod_Cai = "" Then
        strCod_Cai = "00"
    End If
    strCod_Tur = F_BuscaUmCampo("Cod_tur", "tab_caixas", "cod_Cai='" & strCod_Cai & "'")
    
    bActive = True

    lblDia.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    lblNumDia.caption = Day(Date)
    lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    lblAno.caption = Year(Date)
    
    
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

Private Sub MDIForm_Unload(Cancel As Integer)
    'TaskBar True
    P_Fim
End Sub

Private Sub mnAjuda_Click()
    MsgBox "em breve"
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
    If strLandoc_usu = "S" Or strBaixadoc_usu = "S" Then
        frmEditaMov.Show 0
        frmEditaMov.SetFocus
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
End Sub

Private Sub mnCred_Click()
    On Error Resume Next
    frmCred.Show 1
End Sub

Private Sub mnDePara_Click()
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    frmDePara.Show 1
End Sub

Private Sub mnDEv_Click()
    Dim strSQL As String
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    
    
    
    
    'frmDevolucoes.caption = "STYLO PRÓPRIO -Devolução a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
  
    On Error Resume Next
    frmDevolucoes.Show 1
    
End Sub

Private Sub mnEdiComp_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiEntr.Show 0
End Sub


Private Sub mnEdiMovFin_Click()
    
    frmEdiMovFin.Show 0
End Sub

Private Sub mnEdiPer_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiPerdas.Show 0
    frmEdiPerdas.SetFocus
End Sub

Private Sub mnEdiProcomp_Click()
    frmEdiProcomp.Show 0
    frmEdiProcomp.SetFocus
End Sub

Private Sub mnEditCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub mnEditLan_Click()
     
    frmEditaMov.Show 0
End Sub

Private Sub mnEdiTransf_Click()
    frmEdiTransf.Show 0
    frmEdiTransf.SetFocus
End Sub

Private Sub mnEdiTro_Click()
    frmEdiTrocas.Show 0
    frmEdiTrocas.SetFocus
End Sub

Private Sub mnEdiUsu_Click()
    frmEdiUsu.Show 0
End Sub

Private Sub mnEdiVen_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiVendas.Show 0
End Sub

Private Sub mnEfetivVen_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    If IsLoad("frmEfetivaVen") Then
        frmEfetivaVen.SetFocus
    Else
        frmEfetivaVen.Show 0
    End If
End Sub

Private Sub mnEdiVenAtac_Click()
    frmEdiVendas.Show 0, mdiFrm
End Sub

Private Sub mnEfetVen_Click()
    frmEfetivaVen.Show 0
End Sub

Private Sub mnEntrada_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
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
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmLanCaixa
    fr.Show 0
End Sub

Private Sub mnFecTur_Click()
    If strFechacai_usu = "S" Then
        
        frmMovDiarioCx.Show 1
        
    Else
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
    End If

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

Private Sub mnHistCheCPF_Click()
    frmHistoricoChe.Show 0, mdiFrm
End Sub

Private Sub mnIncMer_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmEntradas
    fr.Show 0, mdiFrm
End Sub

Private Sub mnInvent_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmInventario.Show 0, mdiFrm
    frmInventario.SetFocus
End Sub

Private Sub mnLado_Click()
    Me.Arrange vbVertical
End Sub

Private Sub mnLanCaixa_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiCai.Show 0
End Sub

Private Sub mnLojas_Click()
    frmEdiLoj.Show 0
End Sub

Private Sub mnLanCont_Click()
    frmLanCont.Show 1
End Sub

Private Sub mnLoj_Click()
    frmEdiLoj.Show 0
    frmEdiLoj.SetFocus
End Sub

Private Sub mnMonit_Click()
    frmMonitoramentoCaixas.Show 0, mdiFrm
    frmMonitoramentoCaixas.SetFocus
End Sub

Private Sub mnMovFin_Click()
    frmEdiFin.Show 0
    frmEdiFin.SetFocus
End Sub

Private Sub mnMovUsu_Click()
    frmFechaCaixaUsu.Show 0, mdiFrm
End Sub

Private Sub mnMudSen_Click()
    On Error Resume Next
    frmMudaSenha.Show 1
End Sub

Private Sub mnNovaVenda_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.Tag = strSQL
    On Error Resume Next
    frmVenda.Atacado = False
    frmVenda.Show 0, mdiFrm
    
End Sub

Private Sub mnPed_Click()
    frmEdiPed.Show 0
End Sub

Private Sub mnPisos_Click()
    frmEstoqueIni.Show 1
End Sub

Private Sub mnNovoOrc_Click()
    Dim fr As New frmOrcamento
    fr.Show 0
End Sub

Private Sub mnProd_Click()
    frmEdiPro.Show 0
End Sub

Private Sub mnRelGuias_Click()
    frmRelGuias.Show 0
End Sub

Private Sub mnRelComp_Click()
If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
End Sub

Private Sub mnProdProm_Click()
    
        frmProdProm.Show 0
        frmProdProm.SetFocus
    
    
End Sub

Private Sub mnProm_Click()
    If strAltpreco_usu = "S" Then
        If Not F_senha Then
            Exit Sub
        End If
        frmProm.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnRelAnaFat_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFatMensal.Show 0
        frmRelFatMensal.SetFocus
    End If
End Sub

Private Sub mnRelAtvCli_Click()
    If strVisrelfin_usu = "S" Then
        frmRelAtvCli.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mNRelAtvPro_Click()
    If strVisrelfin_usu = "S" Then
        frmRelAtvCliProd.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnRelChq_Click()
    frmRelCheques.Show 0
    frmRelCheques.SetFocus
End Sub

Private Sub mnRelCmpFor_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelProdFor.Show 0
    frmRelProdFor.SetFocus
End Sub

Private Sub mnRelCOmiss_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Você não tem permissão para este aesso!", vbExclamation
    Else
        frmRelComiss.Show 0
        frmRelComiss.SetFocus
    End If
End Sub

Private Sub mnRelDev_Click()
    frmRelDevProd.Show 0
End Sub

Private Sub mnRelFatVenda_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFat.Show 0
        frmRelFat.SetFocus
    End If
End Sub

Private Sub mnRelGiro_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelGiro.Show 0
    frmRelGiro.SetFocus
End Sub

Private Sub mnRelMMD_Click()
    frmRelMMD.Show 0
    frmRelMMD.SetFocus
End Sub

Private Sub mnRelPagReceb_Click()
    frmRelMov.Show 0
    frmRelMov.SetFocus
End Sub

Private Sub mnRelProdForn_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelFornProd.Show 0
    frmRelFornProd.SetFocus
End Sub

Private Sub mnRelProdLucr_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbCritical
        Exit Sub
    End If
    frmRelProdLucr.Show 0
End Sub

Private Sub mnRelVend_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelSaiCaixa.Show 0
    frmRelSaiCaixa.SetFocus
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
    Else
        Exit Sub
    End If
End Sub

Private Sub mnServicos_Click()
    frmEdiServ.Show 0
End Sub

Private Sub mnSujComp_Click()
    If strVisrelest_usu = "S" Then
        frmSujComp.Show 0, mdiFrm
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
    
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

Private Sub mnTroca_Click()
    If strAbresaida_usu = "S" Then
        On Error Resume Next
        frmTroca.Show 1
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnUsu_Click()
    Dim fr As New frmIncUsu
    fr.Show 0
End Sub

Private Sub mnVendaAtac_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.Tag = strSQL
    On Error Resume Next
    frmVenda.Atacado = True
    frmVenda.Show 0, mdiFrm
    

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



