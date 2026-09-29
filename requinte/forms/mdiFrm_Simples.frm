VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiFrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00EEDFD5&
   Caption         =   "Piki"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   840
   ClientWidth     =   14370
   Icon            =   "mdiFrm_Simples.frx":0000
   LinkTopic       =   "Geral"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTempoSemUso 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   720
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14340
      TabIndex        =   5
      Top             =   0
      Width           =   14370
      Begin MSComctlLib.TabStrip tbst 
         Height          =   390
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   688
         Style           =   2
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
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
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   840
   End
   Begin VB.PictureBox pcStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   14370
      TabIndex        =   0
      Top             =   8805
      Width           =   14370
      Begin VB.PictureBox pcUsu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "mdiFrm_Simples.frx":3282
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox txtbxShell 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4575
      End
      Begin VB.PictureBox pcRig 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   11280
         ScaleHeight     =   495
         ScaleWidth      =   3735
         TabIndex        =   2
         Top             =   240
         Width           =   3735
         Begin VB.Label lblDia 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Dia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   8
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblAno 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2007"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   0
            TabIndex        =   7
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblHora 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2595
            TabIndex        =   3
            Top             =   120
            Width           =   1005
         End
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblEmpresa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   480
         TabIndex        =   10
         Top             =   600
         Width           =   675
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5610
         TabIndex        =   1
         Top             =   360
         Width           =   195
      End
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   4320
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnPrinc 
      Caption         =   "&Principal"
      Begin VB.Menu mnEditCli 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu mnrep 
         Caption         =   "&Representantes Comerciais"
      End
      Begin VB.Menu divp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnBuscaTel 
         Caption         =   "Pesquisar Telefone"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnGruCli 
         Caption         =   "Grupos de Clientes"
      End
      Begin VB.Menu mnOrgaoGru 
         Caption         =   "Orgãos de Grupos de Clientes"
      End
      Begin VB.Menu div6 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnForn 
         Caption         =   "&Fornecedores"
      End
      Begin VB.Menu mnFabr 
         Caption         =   "Fa&bricantes"
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnProd 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mnSec 
         Caption         =   "Seções"
      End
      Begin VB.Menu mnSubSec 
         Caption         =   "S&ubseções"
      End
      Begin VB.Menu div4 
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
      Begin VB.Menu mnSetores 
         Caption         =   "Setores"
      End
      Begin VB.Menu mnCidades 
         Caption         =   "Cidades"
      End
      Begin VB.Menu mnDiv6 
         Caption         =   "-"
      End
      Begin VB.Menu mnRecados2 
         Caption         =   "Recados"
      End
      Begin VB.Menu mnAgenda 
         Caption         =   "Agenda"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair do Piki Box"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnSelecionar 
      Caption         =   "&Selecionar"
      Visible         =   0   'False
      Begin VB.Menu mnSelectAll 
         Caption         =   "Selecionar Todos"
      End
      Begin VB.Menu mnSelecNenhum 
         Caption         =   "Selecionar Nenhum"
      End
   End
   Begin VB.Menu mnEntradas 
      Caption         =   "&Entrada de Produtos"
      Begin VB.Menu mnIncMer 
         Caption         =   "&Nova Compra"
      End
      Begin VB.Menu mnEdiComp 
         Caption         =   "&Edição de Compras"
      End
      Begin VB.Menu mnEntrada 
         Caption         =   "E&ntrada de Produtos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnEntradaAntiga 
         Caption         =   "Entrada Antiga"
      End
      Begin VB.Menu mnEdiPer 
         Caption         =   "&Perdas"
      End
      Begin VB.Menu mnEdiProcomp 
         Caption         =   "&Edição de Processos de Compra"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnSujComp 
         Caption         =   "&Sugestão de Compras"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnContagem 
         Caption         =   "Contagem"
      End
      Begin VB.Menu div10 
         Caption         =   "-"
      End
      Begin VB.Menu mnEdiDev 
         Caption         =   "&Devoluções"
         Visible         =   0   'False
      End
      Begin VB.Menu mnDev 
         Caption         =   "Nova Dev&olução"
      End
   End
   Begin VB.Menu mnTransferencias 
      Caption         =   "Movimentação Interna"
      Begin VB.Menu mnInserirTransf 
         Caption         =   "Inserir Arquivos Transferidos"
      End
      Begin VB.Menu mnTrans 
         Caption         =   "Transferência"
      End
   End
   Begin VB.Menu mnConsig 
      Caption         =   "&Consignado"
      Begin VB.Menu mnConsignacoes 
         Caption         =   "Consignações"
      End
      Begin VB.Menu MnTransfCons 
         Caption         =   "Realizar Transferência"
      End
      Begin VB.Menu mnDevolCons 
         Caption         =   "Realizar &Devolução"
      End
      Begin VB.Menu mnAcerto 
         Caption         =   "Realizar Acerto"
      End
      Begin VB.Menu mnAcertoPromocional 
         Caption         =   "&Acerto Promocional"
      End
      Begin VB.Menu DIVCONS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnDevAbertos 
         Caption         =   "Editar Devoluções"
      End
      Begin VB.Menu mnTransfAbert 
         Caption         =   "Editar Transferências"
      End
      Begin VB.Menu mnAcertosAbertos 
         Caption         =   "EditarAcertos"
      End
      Begin VB.Menu mnVarreduraIncons 
         Caption         =   "Varredura de Inconsistências"
      End
      Begin VB.Menu mnRastrearProduto 
         Caption         =   "Rastrear Produto"
      End
      Begin VB.Menu mnRastrEntradasSaidasProd 
         Caption         =   "Rastrear Entradas e Saídas de Produtos"
      End
      Begin VB.Menu mnRastrearProdutoConsig 
         Caption         =   "Rastrear Produtos em um Consignado"
      End
      Begin VB.Menu mnComissoesPermantes 
         Caption         =   "Comissões Permanentes"
      End
      Begin VB.Menu mnDifAce 
         Caption         =   "Diferenças de acertos"
      End
      Begin VB.Menu mnResumoRep 
         Caption         =   "Resumo de Representantes"
      End
   End
   Begin VB.Menu mnVendas 
      Caption         =   "&Vendas"
      Begin VB.Menu mnVendaAtac 
         Caption         =   "&Nova Venda"
      End
      Begin VB.Menu mnEdiVenAtac 
         Caption         =   "&Edição de Vendas"
      End
      Begin VB.Menu divOrc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnCartaoPresente 
         Caption         =   "Cartão-Presente"
      End
      Begin VB.Menu mnEdiCartaoPres 
         Caption         =   "&Acompanhar lançamentos de Cartão-presente"
      End
   End
   Begin VB.Menu mnTrocas 
      Caption         =   "&Trocas"
      Begin VB.Menu mnNovaTroca 
         Caption         =   "&Nova Troca"
      End
      Begin VB.Menu mnEdiTrocas 
         Caption         =   "Edição de Trocas"
      End
      Begin VB.Menu mnTrocaAv 
         Caption         =   "Troca &Avulsa"
      End
      Begin VB.Menu mnTrocasAvulsas 
         Caption         =   "&Visualizar Trocas Avulsas"
      End
   End
   Begin VB.Menu mnSaidas 
      Caption         =   "&Saídas"
      Begin VB.Menu mnEdiNotaSai 
         Caption         =   "Edição de Outras Saídas/Devoluções"
      End
      Begin VB.Menu divRemessa 
         Caption         =   "-"
      End
      Begin VB.Menu mnDevolucao 
         Caption         =   "Gerar Devolução"
      End
      Begin VB.Menu mnManifesto 
         Caption         =   "Gerar Remessa/Manifesto"
      End
      Begin VB.Menu divSai1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEdiMovInt 
         Caption         =   "Edição de &Movimento Interno"
      End
      Begin VB.Menu mnNovoMovInt 
         Caption         =   "Novo Movimento &Interno"
      End
      Begin VB.Menu mnDest 
         Caption         =   "Destinos de Movimento Interno"
      End
   End
   Begin VB.Menu mnNfeSped 
      Caption         =   "NFe/Sped"
      Begin VB.Menu mnNfe 
         Caption         =   "Notas Fiscais Eletrônicas"
      End
      Begin VB.Menu mnInfoAdd 
         Caption         =   "Informações Adicionais"
      End
      Begin VB.Menu mnSpedParticipantes 
         Caption         =   "Participantes"
      End
      Begin VB.Menu mnSpedEmpresa 
         Caption         =   "Sped Empresa"
      End
   End
   Begin VB.Menu mnFinan 
      Caption         =   "&Financeiro"
      Begin VB.Menu mnRecebCarne 
         Caption         =   "&Receber Carnês"
      End
      Begin VB.Menu mnCarnesRecebidos 
         Caption         =   "Carnês recebidos"
      End
      Begin VB.Menu mnEfetVen 
         Caption         =   "&Efetivação de Vendas"
      End
      Begin VB.Menu mnMovFin 
         Caption         =   "&Movimento Financeiro"
      End
      Begin VB.Menu mnContasPagar 
         Caption         =   "&Edição de Documentos"
      End
      Begin VB.Menu mnContasPag 
         Caption         =   "&Contas A Pagar"
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
      Begin VB.Menu mnVinc 
         Caption         =   "&Vinculação de Documentos"
      End
      Begin VB.Menu div11 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnCartoes 
         Caption         =   "Cartões"
      End
      Begin VB.Menu mnBan 
         Caption         =   "&Bancos"
      End
      Begin VB.Menu mnCobradores 
         Caption         =   "Cobradores"
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
      Begin VB.Menu mnFormaPag 
         Caption         =   "Formas de Pagamento"
      End
   End
   Begin VB.Menu mnCobranca 
      Caption         =   "Cobranca"
      Begin VB.Menu mnCobSelec 
         Caption         =   "Selecionar"
      End
      Begin VB.Menu mnCobrSelTodos 
         Caption         =   "Selecionar Todos"
      End
      Begin VB.Menu mnCobrDiv 
         Caption         =   "-"
      End
      Begin VB.Menu mnCobrEdit 
         Caption         =   "&Editar Parcela"
      End
   End
   Begin VB.Menu mnCaixa 
      Caption         =   "Cai&xas"
      Begin VB.Menu mnLanCaixa 
         Caption         =   "&Edição de Lançamentos"
      End
      Begin VB.Menu mnRelVendasEfetNaoCaixa 
         Caption         =   "&Vendas Efetivadas não Lançadas"
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
      Begin VB.Menu mnAbreCx 
         Caption         =   "&Abertura de Caixa"
      End
      Begin VB.Menu mnMovUsu 
         Caption         =   "&Movimentação por Usuário"
      End
      Begin VB.Menu mnEcf 
         Caption         =   "EC&F"
      End
      Begin VB.Menu mnRecebNaoFiscal 
         Caption         =   "Recebimento Não Fiscal"
      End
      Begin VB.Menu mnTEF 
         Caption         =   "&TEF"
      End
      Begin VB.Menu mnVales 
         Caption         =   "Vales"
      End
      Begin VB.Menu mnImpressoraPadraoNFCe 
         Caption         =   "Impressora Padrão NFCe"
      End
   End
   Begin VB.Menu mnRel 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnRelCons 
         Caption         =   "&Consignações"
         Begin VB.Menu mnRelAcertos 
            Caption         =   "Relatório de Acertos"
         End
         Begin VB.Menu mnCentralRelCons 
            Caption         =   "Outros Relatórios"
         End
         Begin VB.Menu mnResumoDiaCons 
            Caption         =   "&Resumo Diário"
         End
      End
      Begin VB.Menu mnRelVendas 
         Caption         =   "&Vendas"
         Begin VB.Menu mnRelVend 
            Caption         =   "Relatório de &Vendas"
         End
         Begin VB.Menu mnInfoVend 
            Caption         =   "Informações de Vendas"
         End
         Begin VB.Menu mnRelDev 
            Caption         =   "&Relatório de Devoluções"
         End
         Begin VB.Menu mnRelComiss 
            Caption         =   "&Vendas por Fabricante/Vendedor"
         End
         Begin VB.Menu mnRelComissGeral 
            Caption         =   "Relatório Geral de Comissões"
         End
         Begin VB.Menu mnPosVenda 
            Caption         =   "Pós-Venda"
         End
         Begin VB.Menu mnPosVendaFxEtaria 
            Caption         =   "Análise de Pós Venda por Faixa Etária"
         End
         Begin VB.Menu mnRelCliNaoDeb 
            Caption         =   "Relatório de Clientes (sem débitos)"
         End
         Begin VB.Menu mnPosVendaAdimplentes 
            Caption         =   "Pós-venda adimplentes"
         End
      End
      Begin VB.Menu mnRelComp 
         Caption         =   "&Compras"
         Begin VB.Menu mnRelProdForn 
            Caption         =   "&Relatório de Produtos por Fornecedor"
         End
         Begin VB.Menu mnRelCmpFor 
            Caption         =   "Relatório de &Compras por Fornecedor"
         End
      End
      Begin VB.Menu div12 
         Caption         =   "-"
      End
      Begin VB.Menu mnRelFatur 
         Caption         =   "&Faturamento"
         Begin VB.Menu mnRelFatVenda 
            Caption         =   "Relatório de &Faturamento por Vendas"
         End
         Begin VB.Menu mnRelAnaFat 
            Caption         =   "&Análise de Faturamento"
         End
         Begin VB.Menu mnRelAnaIndFluxo 
            Caption         =   "Relatório de Análise Direcionada de Fl&uxo"
         End
      End
      Begin VB.Menu mnRelEst 
         Caption         =   "&Análise de Estoque"
         Begin VB.Menu mnRelAltPre 
            Caption         =   "&Relatório de Alterações de Preços"
         End
         Begin VB.Menu mnRelProdLucr 
            Caption         =   "Realtório de Produtos por &Lucratividade"
         End
         Begin VB.Menu mnRelMMD 
            Caption         =   "Relatório de &Media Movimentação Diária"
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
         Begin VB.Menu mnRelCompEst 
            Caption         =   "&Listagem de Comparativo de Estoque"
         End
         Begin VB.Menu mnRelLucrComp 
            Caption         =   "Lucratividade Por &Componente"
         End
         Begin VB.Menu mnDemoInv 
            Caption         =   "&Demonstrativo de Inventário"
         End
         Begin VB.Menu div13 
            Caption         =   "-"
         End
         Begin VB.Menu mnRelListaProd 
            Caption         =   "Lista de Produtos &Vendidos"
         End
         Begin VB.Menu mnListaPerma 
            Caption         =   "&Lista Permanente"
         End
      End
      Begin VB.Menu mnRelFinan 
         Caption         =   "&Financeiro"
         Begin VB.Menu mnResumo 
            Caption         =   "&Resumo de Documentos a Pagar"
         End
         Begin VB.Menu mnHistCheCPf 
            Caption         =   "&Histórico de Cheques por CPF"
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
         Begin VB.Menu mnRecCli 
            Caption         =   "Relação de Carnês Baixados"
         End
         Begin VB.Menu mnRelReceitasDesp 
            Caption         =   "&Relatório de Receitas e Despesas"
         End
      End
      Begin VB.Menu mnInvent 
         Caption         =   "&Inventário de Conferência"
      End
   End
   Begin VB.Menu mnParam 
      Caption         =   "Parâ&metros"
      Begin VB.Menu mnAdm 
         Caption         =   "&Aministração do Sistema"
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
      Begin VB.Menu mnUtilECF 
         Caption         =   "&Utilitário de ECF"
      End
      Begin VB.Menu mnImp 
         Caption         =   "&Impostos"
      End
      Begin VB.Menu mnCont 
         Caption         =   "C&ontagens"
      End
      Begin VB.Menu mnCaixas 
         Caption         =   "&Caixas"
      End
      Begin VB.Menu mnDePara 
         Caption         =   "&De Para"
      End
      Begin VB.Menu mnCalMed 
         Caption         =   "&Cálculo de Média Diária"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnConfig 
         Caption         =   "&Configurações do Sistema"
      End
      Begin VB.Menu mnUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnProm 
         Caption         =   "&Promoções"
      End
      Begin VB.Menu mnMescComp 
         Caption         =   "&Mesclar Componentes"
      End
      Begin VB.Menu mnMudSen 
         Caption         =   "&Mudar Senha"
      End
      Begin VB.Menu mnDesc 
         Caption         =   "&Descontos"
      End
      Begin VB.Menu mnGeraEti 
         Caption         =   "&Gerar Etiquetas"
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
      Begin VB.Menu mnCalc 
         Caption         =   "Ca&lculadora"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnRecados 
         Caption         =   "Recados"
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
      Begin VB.Menu mnDescanso 
         Caption         =   "Descanso de Tela"
      End
      Begin VB.Menu mnVisao 
         Caption         =   "&Visão"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnOptCor 
         Caption         =   "Opções de Cor"
         Begin VB.Menu mnMudaCorFundo 
            Caption         =   "Mudar cor de Fundo"
         End
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
Option Explicit
 Dim bActive As Boolean
Public Sub MudaCorFundo(cor)
    Me.BackColor = cor
End Sub

Public Sub muda_visao(cod As tipo_visao)
        mnParam.Visible = False
    If cod = tCAIXA Then
        
        strVisao = "Caixa"
        mnPrinc.Visible = True
        mnEditCli.Visible = True
        mnrep.Visible = True
        mnGruCli.Visible = True
        mnOrgaoGru.Visible = True
        mnForn.Visible = True
        mnFabr.Visible = True
        mnProd.Visible = True
        mnSec.Visible = True
        mnSubSec.Visible = True
        mnCadCargos.Visible = True
        mnFun.Visible = True
        mnCidades.Visible = True
        mnSetores.Visible = True
        mnEdiUsu.Visible = True
        mnTrocas.Visible = True
    
        mnEntradas.Visible = False
        mnCaixa.Visible = True
        mnTransferencias.Visible = False
        mnConsig.Visible = False
        mnVendas.Visible = True
        mnSaidas.Visible = True
        mnFinan.Visible = True
        mnCobranca.Visible = True
        mnParam.Visible = True
        
        'relatorios
        mnRelCons.Visible = False
        mnRelVendas.Visible = True
        mnRelComp.Visible = False
        mnRelFatur.Visible = True
        mnRelEst.Visible = False
        mnRelFinan.Visible = True
        mnInvent.Visible = False
        
        
        frmMenu.pcClientes.Visible = True
        frmMenu.pcProduto.Visible = False
        frmMenu.pcConsign.Visible = False
        frmMenu.pcCompras.Visible = False
        frmMenu.pcVendas.Visible = False
        frmMenu.pcCaixa.Visible = True
        frmMenu.pcFinanceiro.Visible = True
        
        
        
        
    ElseIf cod = tConsignado Then
        
        strVisao = "Consignado"
        mnTrocas.Visible = True
        mnPrinc.Visible = True
        mnEditCli.Visible = True
        mnrep.Visible = True
        mnGruCli.Visible = True
        mnOrgaoGru.Visible = True
        mnForn.Visible = True
        mnFabr.Visible = True
        mnProd.Visible = True
        mnSec.Visible = True
        mnSubSec.Visible = True
        mnCadCargos.Visible = True
        mnFun.Visible = True
        mnCidades.Visible = True
        mnSetores.Visible = True
        mnEdiUsu.Visible = True
    
        mnEntradas.Visible = False
        mnCaixa.Visible = True
        mnTransferencias.Visible = True
        mnConsig.Visible = True
        mnVendas.Visible = True
        mnSaidas.Visible = False
        mnFinan.Visible = True
        mnCobranca.Visible = True
        mnParam.Visible = False
        
        'relatorios
        mnRelCons.Visible = True
        mnRelVendas.Visible = False
        mnRelComp.Visible = False
        mnRelFatur.Visible = False
        mnRelEst.Visible = False
        mnRelFinan.Visible = True
        mnInvent.Visible = False
        
        
        
        frmMenu.pcClientes.Visible = True
        frmMenu.pcProduto.Visible = True
        frmMenu.pcConsign.Visible = True
        frmMenu.pcCompras.Visible = False
        frmMenu.pcVendas.Visible = False
        frmMenu.pcCaixa.Visible = False
        frmMenu.pcFinanceiro.Visible = True
        
    ElseIf cod = tCrediario Then
        strVisao = "Crediário"
        mnTrocas.Visible = False
        mnPrinc.Visible = True
        mnEditCli.Visible = True
        mnrep.Visible = True
        mnGruCli.Visible = True
        mnOrgaoGru.Visible = True
        mnForn.Visible = True
        mnFabr.Visible = True
        mnProd.Visible = True
        mnSec.Visible = True
        mnSubSec.Visible = True
        mnCadCargos.Visible = True
        mnFun.Visible = True
        mnCidades.Visible = True
        mnSetores.Visible = True
        mnEdiUsu.Visible = True
    
        mnEntradas.Visible = False
        mnCaixa.Visible = True
        mnTransferencias.Visible = False
        mnConsig.Visible = False
        mnVendas.Visible = True
        mnSaidas.Visible = True
        mnFinan.Visible = True
        mnCobranca.Visible = True
        mnParam.Visible = False
        
        'relatorios
        mnRelCons.Visible = False
        mnRelVendas.Visible = True
        mnRelComp.Visible = False
        mnRelFatur.Visible = True
        mnRelEst.Visible = False
        mnRelFinan.Visible = True
        mnInvent.Visible = False
        
        
        frmMenu.pcClientes.Visible = True
        frmMenu.pcProduto.Visible = False
        frmMenu.pcConsign.Visible = False
        frmMenu.pcCompras.Visible = False
        frmMenu.pcVendas.Visible = True
        frmMenu.pcCaixa.Visible = True
        frmMenu.pcFinanceiro.Visible = True
        'frmMenu.pcUsuario.Visible = True
        
    ElseIf cod = tEstoque Then
        strVisao = "Estoque"
        mnPrinc.Visible = True
        mnEditCli.Visible = False
        mnrep.Visible = False
        mnGruCli.Visible = False
        mnOrgaoGru.Visible = False
        mnForn.Visible = True
        mnFabr.Visible = True
        mnProd.Visible = True
        mnSec.Visible = True
        mnSubSec.Visible = True
        mnCadCargos.Visible = False
        mnFun.Visible = False
        mnCidades.Visible = False
        mnSetores.Visible = False
        mnEdiUsu.Visible = False
        mnTrocas.Visible = False
    
        mnEntradas.Visible = True
        mnCaixa.Visible = False
        mnTransferencias.Visible = True
        mnConsig.Visible = False
        mnVendas.Visible = False
        mnSaidas.Visible = False
        mnFinan.Visible = False
        mnCobranca.Visible = False
        mnParam.Visible = False
        
        'relatorios
        mnRelCons.Visible = False
        mnRelVendas.Visible = False
        mnRelComp.Visible = True
        mnRelFatur.Visible = False
        mnRelEst.Visible = True
        mnRelFinan.Visible = False
        mnInvent.Visible = True
        
        
        
        
        
        frmMenu.pcClientes.Visible = False
        frmMenu.pcProduto.Visible = True
        frmMenu.pcConsign.Visible = False
        frmMenu.pcCompras.Visible = True
        frmMenu.pcVendas.Visible = False
        frmMenu.pcCaixa.Visible = False
        frmMenu.pcFinanceiro.Visible = False
        'frmMenu.pcUsuario.Visible = True
    ElseIf cod = tPacote Then
        strVisao = "Pacote"
        mnPrinc.Visible = True
        
        mnEditCli.Visible = True
        mnrep.Visible = False
        mnGruCli.Visible = True
        mnOrgaoGru.Visible = True
        mnForn.Visible = False
        mnFabr.Visible = False
        mnProd.Visible = True
        mnSec.Visible = False
        mnSubSec.Visible = False
        mnCadCargos.Visible = False
        mnFun.Visible = False
        mnCidades.Visible = False
        mnSetores.Visible = False
        mnEdiUsu.Visible = False
    
        mnTrocas.Visible = False
        mnEntradas.Visible = False
        mnCaixa.Visible = False
        mnTransferencias.Visible = False
        mnConsig.Visible = False
        mnVendas.Visible = False
        mnSaidas.Visible = False
        mnFinan.Visible = False
        mnCobranca.Visible = False
        mnParam.Visible = False
        
        'relatorios
        mnRelCons.Visible = False
        mnRelVendas.Visible = False
        mnRelComp.Visible = False
        mnRelFatur.Visible = False
        mnRelEst.Visible = False
        mnRelFinan.Visible = False
        mnInvent.Visible = False
        
        
        frmMenu.pcClientes.Visible = True
        frmMenu.pcProduto.Visible = False
        frmMenu.pcConsign.Visible = False
        frmMenu.pcCompras.Visible = False
        frmMenu.pcVendas.Visible = False
        frmMenu.pcCaixa.Visible = False
        frmMenu.pcFinanceiro.Visible = False
        'frmMenu.pcUsuario.Visible = True
        'frmMenu.pcPacote.Visible = True
        'frmMenu.pcProduto.Visible = True
        
        
        
    ElseIf cod = TVendas Then
        strVisao = "Vendas"
        mnPrinc.Visible = True
        mnEditCli.Visible = False
        mnrep.Visible = False
        mnGruCli.Visible = False
        mnOrgaoGru.Visible = False
        mnForn.Visible = False
        mnFabr.Visible = False
        mnProd.Visible = True
        mnSec.Visible = False
        mnSubSec.Visible = False
        mnCadCargos.Visible = False
        mnFun.Visible = False
        mnCidades.Visible = False
        mnSetores.Visible = False
        mnEdiUsu.Visible = False
    
        mnEntradas.Visible = False
        mnCaixa.Visible = False
        mnTransferencias.Visible = False
        mnConsig.Visible = False
        mnVendas.Visible = True
        mnTrocas.Visible = False
        
        mnSaidas.Visible = False
        mnFinan.Visible = False
        mnCobranca.Visible = False
        mnParam.Visible = False
        
        'relatorios
        mnRelCons.Visible = False
        mnRelVendas.Visible = False
        mnRelComp.Visible = False
        mnRelFatur.Visible = False
        mnRelEst.Visible = False
        mnRelFinan.Visible = False
        mnInvent.Visible = False
        
        
        frmMenu.pcClientes.Visible = False
        frmMenu.pcProduto.Visible = False
        frmMenu.pcConsign.Visible = False
        frmMenu.pcCompras.Visible = False
        frmMenu.pcVendas.Visible = True
        frmMenu.pcCaixa.Visible = False
        frmMenu.pcFinanceiro.Visible = False
        'frmMenu.pcUsuario.Visible = True
    ElseIf cod = tFinanceiro Then
        strVisao = "Financeiro"
        mnTrocas.Visible = True
        mnPrinc.Visible = True
        mnEditCli.Visible = True
        mnrep.Visible = True
        mnGruCli.Visible = True
        mnOrgaoGru.Visible = True
        mnForn.Visible = True
        mnFabr.Visible = True
        mnProd.Visible = True
        mnSec.Visible = True
        mnSubSec.Visible = True
        mnCadCargos.Visible = True
        mnFun.Visible = True
        mnCidades.Visible = True
        mnSetores.Visible = True
        mnEdiUsu.Visible = True
    
        mnEntradas.Visible = False
        mnCaixa.Visible = True
        mnTransferencias.Visible = True
        mnConsig.Visible = True
        mnVendas.Visible = True
        mnSaidas.Visible = False
        mnFinan.Visible = True
        mnCobranca.Visible = True
        mnParam.Visible = True
        
        'relatorios
        mnRelCons.Visible = True
        mnRelVendas.Visible = False
        mnRelComp.Visible = False
        mnRelFatur.Visible = False
        mnRelEst.Visible = False
        mnRelFinan.Visible = True
        mnInvent.Visible = False
        
        
        frmMenu.pcClientes.Visible = True
        frmMenu.pcProduto.Visible = False
        frmMenu.pcConsign.Visible = False
        frmMenu.pcCompras.Visible = True
        frmMenu.pcVendas.Visible = True
        frmMenu.pcCaixa.Visible = True
        frmMenu.pcFinanceiro.Visible = True
        'frmMenu.pcUsuario.Visible = True
    Else
        strVisao = "Geral"
    
    End If
    
    If bUsaConsignado = False Then
        frmMenu.pcConsign.Visible = False
        mnConsig.Visible = False
        mnCons.Visible = False
    
    End If
    If bUsaCobranca = False Then
        mnCobradores.Visible = False
        mnCobranca.Visible = False
        frmMenu.pcBotao(20).Visible = False
        
    End If
        If bUsaCrediario = False Then
            mnCred.Visible = False
            frmMenu.pcBotao(10).Visible = False
        End If
        
        
    
End Sub
Private Sub mnCli_Click()
    frmEdiClientes.Show 0, mdiFrm
End Sub

Private Sub cmndbtnCaixa_Click()
    'frmSaidas.Show 0
    'frmSaidas.SetFocus
End Sub

Private Sub cmndbtnCompras_Click()
    frmEdiEntr.Show 0
End Sub

Private Sub cmndbtnECF_Click()
    On Error Resume Next
    frmECF.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmndbtnFechar_Click()

End Sub

Private Sub cmndbtnDocs_Click()
    frmEditaMov.Show 0, mdiFrm
End Sub

Private Sub cmndbtnJan_Click()
    frmEdiFin.Show 0
End Sub

Private Sub cmndbtnNovaVenda_Click()
    mnEdiVen_Click
End Sub

Private Sub cmndbtnSair_Click()
    mnSair_Click
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0, mdiFrm
End Sub

Private Sub mnAjEstRes_Click()
    
End Sub



Private Sub lblUs_Click()
    frmOptUsu.Show 0
End Sub

Private Sub lblEmpresa_Click()
    buscaEmpresa
    MDIForm_Activate
End Sub

Private Sub MDIForm_Resize()
    pcRig.Left = Me.width - pcRig.width - 100
    lblMsg.Left = (Me.width / 2) - (lblMsg.width / 2)
    txtbxShell.width = Me.width - 100
    tbst.width = Me.width - 100
    
    
End Sub

Private Sub mnAbreCx_Click()
    frmAberturaCaixa.Show 0, mdiFrm
End Sub

Private Sub mnAcerto_Click()
    frmAcerto.acerto_promocao = False
    frmAcerto.Show 0
    frmAcerto.SetFocus
    
End Sub

Private Sub mnAcertoPromocional_Click()
    frmAcerto.acerto_promocao = True
    frmAcerto.Show 0
    
End Sub

Private Sub mnAcertosAbertos_Click()
    frmEdiAcertos.Show 0
End Sub

Private Sub mnAgenda_Click()
    frmAgenda.Show 0
End Sub

Private Sub mnBAckUp_Click()
    frmBackUp.Show 0
End Sub


Private Sub mnBuscaTel_Click()
    On Error Resume Next
    frmBuscarTelefone.Show
    frmBuscarTelefone.SetFocus
End Sub

Private Sub mnCadCargos_Click()
    If strCadusu_usu = "S" Then
        frmEdiCargos.Show 0, mdiFrm
    Else
    End If
End Sub

Private Sub mnCaixas_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbCritical
        Exit Sub
    
    End If
    frmCaixas.Show 0
End Sub

Private Sub mnCalc_Click()
    On Error Resume Next
    shell "C:\windows\calc.exe", vbNormalFocus
    shell "C:\windows\system32\calc.exe", vbNormalFocus
End Sub

Private Sub mnCalMed_Click()
    If strViscusto_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para esta operação!", vbExclamation
        Exit Sub
        
    End If
    frmMediaConsumo_Mer.Show 1
End Sub

Private Sub mnCarnesRecebidos_Click()
    frmRelCarnesRecebidos.Show 0, mdiFrm
    
End Sub

Private Sub mnCartaoPresente_Click()
    frmCartaoPresente.Show 1
End Sub

Private Sub mnCartoes_Click()
    frmCartaoMovimento.Show 0, mdiFrm
End Sub

Private Sub mnCentralRelCons_Click()
    frmCentralRelatorios.Show 0
End Sub

Private Sub mnCidades_Click()
    frmEdiCid.Show 0
    frmEdiCid.SetFocus
End Sub

Private Sub mnCobradores_Click()
    frmEdiContaCob.Show 0, mdiFrm
End Sub

Private Sub mnComissDet_Click()
    If Not F_senha Then Exit Sub
    If strVisrelfin_usu <> "S" Then
        MsgBox "Você não tem permissão para este aesso!", vbExclamation
    Else
        frmRelComiss.Show 0
        frmRelComiss.SetFocus
    End If
End Sub

Private Sub mnComissoesPermantes_Click()
    frmComissParc.Show 1
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


Private Sub mnConsignacoes_Click()
    frmConsign.Show 0
    frmConsign.SetFocus
End Sub

Private Sub mnCont_Click()
    frmContagem.Show 0
    frmContagem.SetFocus
End Sub

Sub mnCOntEcf_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmECF.Show 1
End Sub

Private Sub cmndbtnCli_Click()
    frmEdiClientes.Show 0, mdiFrm
End Sub

Private Sub cmndbtnDev_Click()
    'frmDevolucoes.Show 0
End Sub

Private Sub cmndbtnNewVEnda_Click()
    On Error Resume Next
    frmVenda.Show 1
End Sub


Private Sub cmndbtnProd_Click()
    frmEdiPro.Show 0
    frmEdiPro.SetFocus
End Sub

Private Sub cmndbtnVendas_Click()
    mnNovaVenda_Click
End Sub

Sub verificaModular()
    If bUsaCobranca = False Then
        
    End If
End Sub
Sub ShowMenu(Visao As String)
    If Visao = "estoque" Then
    
    ElseIf Visao = "vendas" Then
    
    ElseIf Visao = "caixa" Then
    
    ElseIf Visao = "crediario" Then
    
    ElseIf Visao = "simples" Then
    
    End If
    
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    
    e = sped_empresa.SelecionarSped_Empresa("")
    If e.codigo <> "" Then
        InicializaEmpresa e.codigo
    End If
    
End Sub
Private Sub MDIForm_Activate()
     On Error Resume Next
   ' VerificaAtualizacao
    'verificaModular
    VerificaAtualizacao
    MudaCorFundo strCorFundoUsu
    frmMenu.BackColor = strCor_Emp
    'mdiFrm.lblUsu.caption = "" & strNome_Usu
    'mdiFrm.lblCodUsu.caption = strCod_Usu
    Me.caption = constNome_sis & " " & constVersao_Sis & " - " & strFantasia_Emp & " - " & strCidade_Emp & "/" & strUF_Emp  '& " .:. CNPJ.: " & strCNPJ_Emp
    'TaskBar False
   ' If F_BuscaUmCampo("cod_loj", "tab_parametros", "") = "" Then
    '    MsgBox "Os Dados desta Loja não estão definidos corretamente!" & Chr(13) & _
     '       "Informe esta mensagem ao Responsável!", vbExclamation
      '  P_Fim
        'End If
   'lblIp.caption = Winsock1.LocalIP & "-" & Winsock1.LocalHostName
   strIP_Usu = Winsock1.LocalIP
   strHost_Usu = Winsock1.LocalHostName
   
   frmMenu.Show 0
   
   strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixasIP", "ip_cai='" & strHost_Usu & "'")
   
   If strCod_Cai = "" Then
    
    'strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixasIP", "ip_cai='" & strIP_Usu & "'")
   End If
   
   strSerie_Nfe = F_BuscaUmCampo("serie_nfe", "tab_caixas", "cod_Cai='" & strCod_Cai & "'")
   If strSerie_Nfe = "" Then
    strSerie_Nfe = "0"
   End If
   
   strImprFiscal = F_BuscaUmCampo("impressora_cai", "tab_caixas", "ip_cai='" & _
    strIP_Usu & "' or ip_cai='" & strHost_Usu & "'")
   
   If strImprFiscal = "" Then
        strImprFiscal = "B"
   End If
   If strCod_Cai = "" Then
        strCod_Cai = "00"
   End If
   strCod_Tur = F_BuscaUmCampo("Cod_tur", "tab_caixas", "cod_Cai='" & strCod_Cai & "'")
    If bActive <> True Then
'        frmMenu.Show 0, Me
'        frmMenu.Left = 0
'        frmMenu.Top = 600
'        frmMenu.Height = StatusBar1.Top - 50
        bActive = True
        
    End If
    'pcCalend.Top = pcStatus.Top - pcCalend.Height
    
    
   lblDia.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase) & ", " & Day(Date) & " de " & _
    StrConv(MonthName(Month(Date)), vbProperCase) & " de " & Year(Date)
    lblEmpresa.caption = strCod_Emp & "-" & strNome_Emp
    lblUsuario.caption = strCod_Usu & "-" & strNome_Usu
    
End Sub

Private Sub MDIForm_DblClick()
    'frmVenda.Show 0, Me
End Sub

Private Sub MDIForm_Load()
'    pcCalend.Top = pcStatus.Top - pcCalend.Height
    
    'frmLogo.Show 0
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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




Private Sub mnCadPla_Click()
    frmEdiPlanoContas.Show 0
    frmEdiPlanoContas.SetFocus
End Sub

Private Sub mnCalGuias_Click()
    frmLanDoc.Show 1
End Sub

Private Sub mnCasc_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnCheq_Click()
    frmControleChequesReceber.Show 0
End Sub


Private Sub mnCOn_Click()
    frmEdiCon.Show 0
End Sub

Private Sub mnConfig_Click()
    frmParametros.Show 0
End Sub

Private Sub mnContagem_Click()
    frmContagem.Show 0, mdiFrm
End Sub

Private Sub mnContasPag_Click()
    frmContasPagar.Show 0
    frmContasPagar.SetFocus
End Sub

Private Sub mnContasPagar_Click()
    On Error Resume Next
    frmEditaMov.Show 0, mdiFrm
End Sub



Private Sub mnCred_Click()
    On Error Resume Next
    frmCred.Show 1
End Sub

Private Sub mnDemoInv_Click()
    frmDemonstrativoInvent.Show 0
    
End Sub

Private Sub mnDePara_Click()
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    frmDePara.Show 1
End Sub



Private Sub mnDescanso_Click()
    frmCaixaLivre.Show 1
End Sub

Private Sub mnDest_Click()
    frmEdiUmCampo.operacao = "5"
    frmEdiUmCampo.Show 0
End Sub

Private Sub mnDEv_Click()
    Dim strSQL As String
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    
    
    
    
    
    
    
  
    On Error Resume Next
    
    
End Sub

Private Sub mnDevAbertos_Click()
    frmEdiDevConsign.Show
End Sub

Private Sub mnDevolCons_Click()
    frmDevolucaoConsign.Show 0
    frmDevolucaoConsign.SetFocus
End Sub

Private Sub mnDevolucao_Click()
     Dim strSQL As String
    
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
    
    strSQL = Vendas.F_GeraNovaVenda("00000", tipo_operacao_venda.Devolucao)
    
    
    
    'frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    Dim fr As New frmVenda
    
    fr.tag = strSQL
    
    
    fr.Show 1
End Sub

Private Sub mnDifAce_Click()
    
    
    frmRelDifAcerto.Show 0
End Sub

Private Sub mnEcf_Click()
    frmECF.Show 0, mdiFrm
End Sub

Private Sub mnEdiCaixas_Click()
    
    frmCaixas.Show 0
    
End Sub

Private Sub mnEdiCartaoPres_Click()
    frmEdiCartao_Presente.Show 1
End Sub

Private Sub mnEdiComp_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEntradasSped.Show 0, mdiFrm
End Sub


Private Sub mnEdiMovFin_Click()
    
    frmEdiMovFin.Show 0
End Sub

Private Sub mnEdiMovInt_Click()
    frmEdiMovInt.Show 0
    frmEdiMovInt.SetFocus
End Sub

Private Sub mnEdiNotaSai_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    frmEdiVendas.Tipo_Operacao = Remessa
    frmEdiVendas.Show 0, mdiFrm
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
    frmEdiClientes.Show 0, mdiFrm
End Sub

Private Sub mnEditLan_Click()
     
    frmEditaMov.Show 0, mdiFrm
End Sub

Private Sub mnEdiTransf_Click()
    frmEdiMovInt.Show 0
    frmEdiMovInt.ZOrder
End Sub

Private Sub mnEdiTro_Click()
    frmEdiTrocas.Show 0
    frmEdiTrocas.SetFocus
End Sub

Private Sub mnEdiTrocas_Click()
    frmEdiTrocas.Show 0
End Sub

Private Sub mnEdiUsu_Click()
    frmEdiUsu.Show 0
End Sub

Private Sub mnEdiVen_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiVendas.Show 0, mdiFrm
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



Private Sub mnEntradaAntiga_Click()
    frmEdiEntr.Show 0
End Sub

Private Sub mnFinanceiras_Click()
    frmEdiFinan.Show 0
End Sub

Private Sub mnFormaPag_Click()
    frmEdiFormasPag.Show 1
    
End Sub

Private Sub mnFormasPag_Click()
    frmEdiFormasPag.Show 0
End Sub

Private Sub mnForn_Click()
    frmEdiFor.Show 0
End Sub


Private Sub mnFun_Click()
    frmEdiFun.Show 0
End Sub

Private Sub mnGeraEti_Click()
    frmGeraEtiquetas.Show 0
    frmGeraEtiquetas.SetFocus
End Sub


Private Sub mnGruCli_Click()
    frmEdiGru.Show 0
    frmEdiGru.SetFocus
End Sub

Private Sub mnHistCheCPf_Click()
    frmHistoricoChe.Show 0, mdiFrm
End Sub


Private Sub mnImp_Click()
    frmImpost.Show 1
End Sub

Private Sub mnImpressoraPadraoNFCe_Click()
    MudaImpressoraPadraoNFCe
End Sub

Private Sub mnIncMer_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmEntradaSped
    fr.Show 0, mdiFrm
End Sub

Private Sub mnInserirTransf_Click()
    frmMigraMovInt.Show 1
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




Private Sub mnManifesto_Click()
    Dim strSQL As String
    
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
    
    strSQL = Vendas.F_GeraNovaVenda("00000", Remessa)
    
    
    
    'frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    Dim fr As New frmVenda
    
    fr.tag = strSQL
    
    
    fr.Show 1
    
End Sub

Private Sub mnMescComp_Click()
    frmMesclaComp.Show 0, mdiFrm
End Sub

Private Sub mnMigra_Click()
    frmDump.Show 1
End Sub

Private Sub mnMovFin_Click()
    frmEdiFin.Show 0
    frmEdiFin.SetFocus
End Sub

Private Sub mnMovUsu_Click()
'    frmFechaCaixaUsu.Show 0, mdiFrm
End Sub

Public Sub mnMudaCorFundo_Click()
    
    Dim cor
    mdiFrm.cmndlg.ShowColor
    cor = mdiFrm.cmndlg.color
    
    
    frmMenu.mudaCor cor
    If MsgBox("Deseja utilizar essa cor na próxima vez que acessar o sistema?", vbYesNo + vbQuestion) = vbYes Then
        If Not F_senha Then Exit Sub
        funcionario.AlteraCorFundoUsu strCod_Usu, cor
    End If
    
    

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
    
    strSQL = F_GeraNovaVenda("00000")
    
    
    
    
    
    
    frmVenda.tag = strSQL
    'On Error Resume Next
    frmVenda.Atacado = False
    frmVenda.Show 0, mdiFrm
    
End Sub



Private Sub mnNovaTroca_Click()
    If F_IsCaixa = False Then
        MsgBox "Esse computador não é um caixa. Portanto não pode realizar trocas.", vbExclamation
        Exit Sub
    End If
    
    frmTroca.Show 0, mdiFrm
End Sub

Private Sub mnNovoMovInt_Click()
    Dim fr As New frmMovInt
    fr.Show 0
End Sub

Private Sub mnNovoOrc_Click()
    'Dim fr As New frmOrcamento
    'fr.Show 0
End Sub


Private Sub mnOrgaoGru_Click()
    frmEdiOrgao.Show 0, mdiFrm
    frmEdiOrgao.SetFocus
End Sub

Private Sub mnPrincipio_Click()
    principio.P_EditaPrincipio
End Sub

Private Sub mnPosVenda_Click()
    frmRelPosVenda.Show 0
End Sub

Private Sub mnPosVendaAdimplentes_Click()
    
    frmRelPosVendaClientesAdimplentes.Show 1
End Sub

Private Sub mnPosVendaFxEtaria_Click()
    frmRelPosVendaFaixa.Show 0, mdiFrm
End Sub

Private Sub mnProd_Click()
    frmEdiPro.Show 0
End Sub


Private Sub mnRastrearProduto_Click()
    frmProdutosPorRepresentante.Show 0
End Sub

Private Sub mnRastrearProdutoConsig_Click()
    frmHistProdConsignado.Show 0, mdiFrm
End Sub

Private Sub mnRastrEntradasSaidasProd_Click()
    frmHistProdAndamento.Show 0, mdiFrm
End Sub

Private Sub mnRecados_Click()
    frmRecados.Show 1
End Sub

Private Sub mnRecados2_Click()
    frmMensagens.Show 0
End Sub

Private Sub mnRecCli_Click()
    frmRelCarnesRecebidos.Show 0
    frmRelCarnesRecebidos.SetFocus
End Sub

Private Sub mnRecebCarne_Click()
    frmBaixaDocReq.Show 1
End Sub

Private Sub mnRecebNaoFiscal_Click()
    frmRecebNaoFiscal.Show 1
End Sub

Private Sub mnRelAcertos_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
        
    End If
    frmRelAcertos.Show 0
End Sub

Private Sub mnRelAltPre_Click()
    frmAltPreco.Show 0, mdiFrm
End Sub

Private Sub mnRelCliNaoDeb_Click()
    frmRelClientesNaoDebito.Show 0
End Sub

Private Sub mnRelComissGeral_Click()
    frmRelComissGeral.Show 1
    
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
    
        
        frmProdProm.Show 0
    
End Sub

Private Sub mnRelAnaFat_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFatMensal.Show 0
        frmRelFatMensal.SetFocus
    End If
End Sub

Private Sub mnRelAnaIndFluxo_Click()
    If strVisrelest_usu = "S" Then
        frmRelAnaliseIndFluxo.Show 0
        frmRelAnaliseIndFluxo.SetFocus
    Else
        MsgBox "Acesso Inválido!", vbCritical
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
        frmRelVendasVendedor.Show 1
        
    End If
End Sub

Private Sub mnRelCompEst_Click()
    If strVisrelest_usu = "S" Then
        frmRelEstFis.Show
    End If
End Sub

Private Sub mnRelDev_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
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

Private Sub mnRelFluxo_Click()
    If strVisrelest_usu = "S" Then
        frmRelAnaliseFluxo.Show 0
        frmRelAnaliseFluxo.SetFocus
    Else
        MsgBox "Acesso Inválido!", vbCritical
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

Private Sub mnRelListaProd_Click()
    frmRelAtvProdu.Show 0, mdiFrm
End Sub

Private Sub mnRelLucrComp_Click()
    If strVisrelest_usu = "S" Then
        frmRelLucrComp.Show 0
        frmRelLucrComp.SetFocus
    End If
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

Private Sub mnRelReceitasDesp_Click()
    frmRelReceitasDespesas.Show 0
End Sub

Private Sub mnRelVend_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelSaiCaixa.Show 0
    frmRelSaiCaixa.SetFocus
End Sub


Private Sub mnRelVendasEfetNaoCaixa_Click()
    frmRelVendasNaoLancadasEmCaixa.Show 1
End Sub

Private Sub mnrep_Click()
    frmEdiRep.Show 0
End Sub

Private Sub mnResumo_Click()
    frmResumoContasPagar.Show 0
End Sub

Private Sub mnResumoRep_Click()
    If strBaixarComRepr_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    frmRelResumoRep.Show
    
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

Private Sub mnSec_Click()
    secao.P_EditaSecao
End Sub

Private Sub mnSetores_Click()
    frmEdiSet.Show 0
    frmEdiSet.SetFocus
End Sub

Private Sub mnSpedEmpresa_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    frmEdiSpedEmpresa.Show 0, mdiFrm
End Sub

Private Sub mnSpedParticipantes_Click()
    If strCadcli_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para esse acesso", vbExclamation
        
        Exit Sub
    End If
    frmEdiParticipante.Show 0, mdiFrm
End Sub

Private Sub mnSubSec_Click()
    frmEdiSubSec.Show 0, mdiFrm
    frmEdiSubSec.SetFocus
End Sub

Private Sub mnSujComp_Click()
    If strVisrelest_usu = "S" Then
        frmSujComp.Show 0
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
    
End Sub

Private Sub mnTEF_Click()
    P_AdmTEF
End Sub

Private Sub mnTodasJan_Click()
    'On Error Resume Next
   
   If Forms.Count > 1 Then
        frmForms.Show , mdiFrm
    End If
   
End Sub

Private Sub mnTrans_Click()
    frmEdiMovInt.Show 0
End Sub

Private Sub mnTransf_Click()
    
'    frmExportar.Show 0, mdiFrm
End Sub

Private Sub mnTroca_Click()
    
End Sub

Private Sub mnTransfComp_Click()
    frmTransfComp.Show 0, mdiFrm
    
End Sub

Private Sub mnTransfAbert_Click()
    If strCorrigeEstCons_Usu <> "S" Then
        Exit Sub
    End If
    
    frmEdiTransAbertas.Show 0
    frmEdiTransAbertas.SetFocus
End Sub

Private Sub MnTransfCons_Click()
    frmTransfConsign.Show 0
    frmTransfConsign.SetFocus
End Sub

Private Sub mnTrocaAv_Click()
    frmTrocaAvulsa.Show 1
End Sub

Private Sub mnTrocas_Click()
'    frmEdiTrocas.Show 0
End Sub

Private Sub mnTrocasAvulsas_Click()
    frmEdiTrocasAvulsas.Show 0
End Sub

Private Sub mnUsu_Click()
    Dim fr As New frmIncUsu
    fr.Show 0
End Sub

Private Sub mnUtilECF_Click()
    
    frmUtilitarioECF.Show 1
End Sub

Private Sub mnVales_Click()
    If strLancaixa_usu = "S" Then
        frmEdiVales.Show
        
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
End Sub

Private Sub mnVarreduraIncons_Click()
    frmVerificaInconsistenciaCons.Show 1
End Sub

Private Sub mnVendaAtac_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = Vendas.F_GeraNovaVenda("00000")
    
    
    
    'frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    
    frmVenda.Show 0, mdiFrm
    

End Sub

Private Sub mnVinc_Click()
    frmVinculaDoc.Show 1
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
        If Forms.Count = 1 Then
            mnSair_Click
        Else
            'For iCount = 1 To Forms.count - 1
                
               ' MsgBox Forms(iCount).Name
                 Unload Screen.ActiveForm
            'Next
        End If
    End If
    
End Sub



Private Sub mnVisao_Click()
    frmEscolherVisao.Show 1
End Sub

Private Sub pcUsu_Click()
    frmOptUsu.Show 0
End Sub

Private Sub tbst_Click()
    Dim iCount As Integer
    On Error GoTo trata_ERRO
    For iCount = 1 To Forms.Count - 1
        If Forms(iCount).Name = tbst.SelectedItem.tag Then
            Forms(iCount).Visible = True
            Forms(iCount).SetFocus
            
           ' Me.Arrange 0
            Forms(iCount).ZOrder = True
            Exit For
        End If
    Next
    Exit Sub
trata_ERRO:
End Sub


Private Sub Timer1_Timer()
    lblHora.caption = Time
    buscaFrm
End Sub

Private Sub txtbxShell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_shell txtbxShell.text
    End If
    
End Sub
Sub buscaFrm()
    On Error Resume Next
    tbst.Tabs.Clear
    
    Dim iCount As Integer
    For iCount = 1 To Forms.Count - 1
        If Forms(iCount).Visible = True Then
            tbst.Tabs.Add , , Forms(iCount).caption
            tbst.Tabs(iCount).tag = Forms(iCount).Name
        End If
    Next
End Sub
