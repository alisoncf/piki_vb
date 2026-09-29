VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDefUsu 
   Caption         =   "Definições de Usuário"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   9795
   Icon            =   "frmDefUsu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab stb 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Financeiro"
      TabPicture(0)   =   "frmDefUsu.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Venda"
      TabPicture(1)   =   "frmDefUsu.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Caixa"
      TabPicture(2)   =   "frmDefUsu.frx":1D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cadastros"
      TabPicture(3)   =   "frmDefUsu.frx":1D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Estoque"
      TabPicture(4)   =   "frmDefUsu.frx":1D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Controle"
      TabPicture(5)   =   "frmDefUsu.frx":1D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Consignado"
      TabPicture(6)   =   "frmDefUsu.frx":1D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Visões"
      TabPicture(7)   =   "frmDefUsu.frx":1D8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "frVisoes"
      Tab(7).ControlCount=   1
      Begin VB.Frame frVisoes 
         Caption         =   "Visões"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74880
         TabIndex        =   111
         Top             =   360
         Width           =   9135
         Begin VB.CheckBox chkCaixa 
            Caption         =   "Caixa"
            Height          =   195
            Left            =   120
            TabIndex        =   119
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkConsignado 
            Caption         =   "Consignado"
            Height          =   195
            Left            =   120
            TabIndex        =   118
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox chkCrediario 
            Caption         =   "Crediário"
            Height          =   195
            Left            =   120
            TabIndex        =   117
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkFinanceiro 
            Caption         =   "Financeiro"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkPacote 
            Caption         =   "Pacote"
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkVendas 
            Caption         =   "Vendas"
            Height          =   195
            Left            =   120
            TabIndex        =   114
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkGeral 
            Caption         =   "Geral"
            Height          =   195
            Left            =   1800
            TabIndex        =   113
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkEstoque 
            Caption         =   "Estoque"
            Height          =   195
            Left            =   1800
            TabIndex        =   112
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Consignado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74880
         TabIndex        =   91
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbxTransfCons 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1DAA
            Left            =   3240
            List            =   "frmDefUsu.frx":1DB4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbxDevCons 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1DC2
            Left            =   3240
            List            =   "frmDefUsu.frx":1DCC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cmbxAcertoCons 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1DDA
            Left            =   3240
            List            =   "frmDefUsu.frx":1DE4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ComboBox cmbxAlteraComsRepr 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1DF2
            Left            =   3240
            List            =   "frmDefUsu.frx":1DFC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ComboBox cmbxBaixaComRepr 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1E0A
            Left            =   3240
            List            =   "frmDefUsu.frx":1E14
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   2160
            Width           =   1095
         End
         Begin VB.ComboBox cmbxExclAcertoCons 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1E22
            Left            =   3240
            List            =   "frmDefUsu.frx":1E2C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cmbxCorrigeEstCons 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1E3A
            Left            =   3240
            List            =   "frmDefUsu.frx":1E44
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   104
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efetuar Transferência"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efetuar Devolução:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   720
            Width           =   1380
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efetuar Acerto"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alterar Comissão de Representante"
            Height          =   195
            Left            =   120
            TabIndex        =   100
            Top             =   1800
            Width           =   2490
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Baixar comissão de representante"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   2160
            Width           =   2385
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir Acerto"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corrigir Estoque Consignado:"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   2520
            Width           =   2040
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Controle:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   -74880
         TabIndex        =   80
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbxVisFin 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1E52
            Left            =   2640
            List            =   "frmDefUsu.frx":1E5C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox cmbxVisRel 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1E6A
            Left            =   2640
            List            =   "frmDefUsu.frx":1E74
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbxVisCusto 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1E82
            Left            =   2640
            List            =   "frmDefUsu.frx":1E8C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cmbxContrPre 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1E9A
            Left            =   2640
            List            =   "frmDefUsu.frx":1EA4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   1680
            Width           =   1095
         End
         Begin VB.ComboBox cmbxFechaMes 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1EB2
            Left            =   2640
            List            =   "frmDefUsu.frx":1EBC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visualizar relatórios financeiros"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   1020
            Width           =   2145
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visualizar relatórios de estoque"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   660
            Width           =   2175
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visualizar Custos"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Controlar Preços"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   1740
            Width           =   1170
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fechar Mês"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   1380
            Width           =   840
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estoque:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74880
         TabIndex        =   71
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbxCadPro 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1ECA
            Left            =   1920
            List            =   "frmDefUsu.frx":1ED4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cmbxAltPro 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1EE2
            Left            =   1920
            List            =   "frmDefUsu.frx":1EEC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cmbxAltPrecos 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1EFA
            Left            =   1920
            List            =   "frmDefUsu.frx":1F04
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   960
            Width           =   855
         End
         Begin VB.ComboBox cmbxTransfInt 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1F12
            Left            =   1920
            List            =   "frmDefUsu.frx":1F1C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar Produto"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alterar Produto"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   660
            Width           =   1050
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alterar Preços"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1020
            Width           =   990
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transferência Interna"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   1380
            Width           =   1515
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cadastro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   -74880
         TabIndex        =   108
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbxCadCli 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1F2A
            Left            =   2040
            List            =   "frmDefUsu.frx":1F34
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cmbxCadUsu 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1F42
            Left            =   2040
            List            =   "frmDefUsu.frx":1F4C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbxCadFun 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1F5A
            Left            =   2040
            List            =   "frmDefUsu.frx":1F64
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cmbxCadFor 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1F72
            Left            =   2040
            List            =   "frmDefUsu.frx":1F7C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox cmbxCadRepr 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1F8A
            Left            =   2040
            List            =   "frmDefUsu.frx":1F94
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar Clientes"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar Usuários"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar Funcionários"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   1140
            Width           =   1620
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar Fornecedores"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   1500
            Width           =   1695
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar Representante"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   1860
            Width           =   1770
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Caixa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   -74880
         TabIndex        =   52
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbxExclCai 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1FA2
            Left            =   2280
            List            =   "frmDefUsu.frx":1FAC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cmbxEdiCai 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1FBA
            Left            =   2280
            List            =   "frmDefUsu.frx":1FC4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ComboBox cmbxFechaCai 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1FD2
            Left            =   2280
            List            =   "frmDefUsu.frx":1FDC
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cmbxLanCai 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":1FEA
            Left            =   2280
            List            =   "frmDefUsu.frx":1FF4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir Caixa"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   1500
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Editar Caixa"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1140
            Width           =   840
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fechar Caixa"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   780
            Width           =   930
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fazer Lançamento no Caixa"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   420
            Width           =   1980
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Venda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbxDesbloquearVenda 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2002
            Left            =   1920
            List            =   "frmDefUsu.frx":200C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   2880
            Width           =   855
         End
         Begin VB.ComboBox cmbxAbreVenda 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":201A
            Left            =   1920
            List            =   "frmDefUsu.frx":2024
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cmbxEfetVenda 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2032
            Left            =   1920
            List            =   "frmDefUsu.frx":203C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cmbxAbreNota 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":204A
            Left            =   1920
            List            =   "frmDefUsu.frx":2054
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cmbxVisualizarVenda 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2062
            Left            =   1920
            List            =   "frmDefUsu.frx":206C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox cmbxExclVenda 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":207A
            Left            =   1920
            List            =   "frmDefUsu.frx":2084
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   1800
            Width           =   855
         End
         Begin VB.ComboBox cmbxDesconto 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2092
            Left            =   1920
            List            =   "frmDefUsu.frx":20A5
            Sorted          =   -1  'True
            TabIndex        =   47
            Text            =   "cmbxDesconto"
            Top             =   2160
            Width           =   855
         End
         Begin VB.ComboBox cmbxCrediario 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":20BD
            Left            =   1920
            List            =   "frmDefUsu.frx":20C7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reabrir Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abrir Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efetivar Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abrir Nota Fiscal"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visualizar Vendas"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   1440
            Width           =   1245
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Venda"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   2160
            Width           =   1200
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Crediário"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   2520
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Financeiro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9135
         Begin VB.ComboBox cmbxLanDoc 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":20D5
            Left            =   2040
            List            =   "frmDefUsu.frx":20DF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cmbxBaixarDoc 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":20ED
            Left            =   2040
            List            =   "frmDefUsu.frx":20F7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cmbxExclDoc 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2105
            Left            =   2040
            List            =   "frmDefUsu.frx":210F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cmbxBaixarChe 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":211D
            Left            =   2040
            List            =   "frmDefUsu.frx":2127
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2160
            Width           =   855
         End
         Begin VB.ComboBox cmbxLanChe 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2135
            Left            =   2040
            List            =   "frmDefUsu.frx":213F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1800
            Width           =   855
         End
         Begin VB.ComboBox cmbxCadFormasPag 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":214D
            Left            =   2040
            List            =   "frmDefUsu.frx":2157
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   2880
            Width           =   855
         End
         Begin VB.ComboBox cmbxCadCCO 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2165
            Left            =   2040
            List            =   "frmDefUsu.frx":216F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   3240
            Width           =   855
         End
         Begin VB.ComboBox cmbxAltFormasPag 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":217D
            Left            =   2040
            List            =   "frmDefUsu.frx":2187
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3600
            Width           =   855
         End
         Begin VB.ComboBox cmbxAlterarVlrParcela 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2195
            Left            =   2040
            List            =   "frmDefUsu.frx":219F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox cmbxProrrogarParcela 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":21AD
            Left            =   2040
            List            =   "frmDefUsu.frx":21B7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   3960
            Width           =   855
         End
         Begin VB.ComboBox cmbxAlterarTerceiroDoc 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":21C5
            Left            =   5640
            List            =   "frmDefUsu.frx":21CF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cmbxExclCheque 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":21DD
            Left            =   2040
            List            =   "frmDefUsu.frx":21E7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2520
            Width           =   855
         End
         Begin VB.ComboBox cmbxLanFin 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":21F5
            Left            =   5640
            List            =   "frmDefUsu.frx":21FF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cmbxEdiFin 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":220D
            Left            =   5640
            List            =   "frmDefUsu.frx":2217
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cmbxExclFin 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2225
            Left            =   5640
            List            =   "frmDefUsu.frx":222F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1440
            Width           =   855
         End
         Begin VB.ComboBox cmbxDescontoBaixa 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":223D
            Left            =   5640
            List            =   "frmDefUsu.frx":2250
            Sorted          =   -1  'True
            TabIndex        =   32
            Text            =   "cmbxDesconto"
            Top             =   1800
            Width           =   855
         End
         Begin VB.ComboBox cmbxLiberarCli 
            Height          =   315
            ItemData        =   "frmDefUsu.frx":2268
            Left            =   5640
            List            =   "frmDefUsu.frx":2272
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fazer Lanç. Doc"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Baixar Doc."
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir Doc."
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Baixar Cheque"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fazer Lanç. Cheque"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   1440
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar Frms Pgto"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   2880
            Width           =   1425
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cadastrar C. Corr."
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   3240
            Width           =   1245
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alterar. Formas Pgto"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   3600
            Width           =   1425
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alterar Valor Parcela"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   1440
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prorrogar Parcela:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   3960
            Width           =   1275
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alterar Terceiro Parcela:"
            Height          =   195
            Left            =   3720
            TabIndex        =   23
            Top             =   360
            Width           =   1710
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir Cheque"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lançar Financeiro"
            Height          =   195
            Left            =   3720
            TabIndex        =   25
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Editar Financeiro"
            Height          =   195
            Left            =   3720
            TabIndex        =   27
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir Financeiro"
            Height          =   195
            Left            =   3720
            TabIndex        =   29
            Top             =   1440
            Width           =   1245
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Baixa Parc."
            Height          =   195
            Left            =   3720
            TabIndex        =   31
            Top             =   1800
            Width           =   1545
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liberar/Bloquear Cliente"
            Height          =   195
            Left            =   3720
            TabIndex        =   33
            Top             =   2160
            Width           =   1710
         End
      End
   End
   Begin VB.CommandButton cmdVisao 
      Caption         =   "&Visão"
      Height          =   735
      Left            =   3360
      Picture         =   "frmDefUsu.frx":2280
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   6480
      Picture         =   "frmDefUsu.frx":2B27
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   7800
      Width           =   3135
   End
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   240
      Picture         =   "frmDefUsu.frx":2C92
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   7800
      Width           =   3015
   End
End
Attribute VB_Name = "frmDefUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub cmbxEscVend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxDesconto_LostFocus()
    P_ValorLostFocus cmbxDesconto
End Sub

Private Sub cmbxDescontoBaixa_LostFocus()
    P_ValorLostFocus cmbxDesconto, "###"
End Sub

Private Sub cmdVisao_Click()
    frmVisao.Show 1
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If Not F_senha() Then
        MsgBox "não logado!"
        Exit Sub
    End If
    If MsgBox("Confirma alteração das definições de usuário?", vbYesNo + vbQuestion) = vbYes Then
        
        grava
        P_BuscaDefUsu
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    bActive = True
    busca
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_defusu where cod_usu = '" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            Exit Sub
        End If
        P_LocalizaOperacao !abrevenda_usu, cmbxAbreVenda
        'P_LocalizaOperacao !efetvenda_usu, cmbxEfetvenda
        'P_LocalizaOperacao !baixavenda_usu, cmbxBaixarVenda
        'P_LocalizaOperacao !efetped_usu, cmbxEfetPed
        'P_LocalizaOperacao !abreped_usu, cmbxAbrPed
        P_LocalizaOperacao !abrenota_usu, cmbxAbreNota
        'P_LocalizaOperacao !efetnota_usu, cmbxEfetNota
        'P_LocalizaOperacao !abresaida_usu, cmbxAbreSaida
        P_LocalizaOperacao !lancaixa_usu, cmbxLanCai
        P_LocalizaOperacao !landoc_usu, cmbxLanDoc
        P_LocalizaOperacao !lanche_usu, cmbxLanChe
        P_LocalizaOperacao !baixadoc_usu, cmbxBaixarDoc
        P_LocalizaOperacao !baixache_usu, cmbxBaixarChe
        P_LocalizaOperacao !edicai_usu, cmbxEdiCai
        P_LocalizaOperacao !exclcai_usu, cmbxExclCai
        P_LocalizaOperacao !excldoc_usu, cmbxExclDoc
        P_LocalizaOperacao !cadprod_usu, cmbxCadPro
        P_LocalizaOperacao !altprod_usu, cmbxAltPro
        P_LocalizaOperacao !altpreco_usu, cmbxAltPrecos
        P_LocalizaOperacao !cadusu_usu, cmbxCadUsu
        P_LocalizaOperacao !cadcli_usu, cmbxCadCli
        P_LocalizaOperacao !cadcco_usu, cmbxCadCCO
        P_LocalizaOperacao !cadformaspag_usu, cmbxCadFormasPag
        P_LocalizaOperacao !exclvenda_usu, cmbxExclVenda
        P_LocalizaOperacao !altformaspag_usu, cmbxAltFormasPag
        P_LocalizaOperacao !corrigeestcons_usu, cmbxCorrigeEstCons
        
        P_LocalizaOperacao !viscusto_usu, cmbxVisCusto
        P_LocalizaOperacao !visrelest_usu, cmbxVisRel
        P_LocalizaOperacao !visrelfin_usu, cmbxVisFin
        P_LocalizaOperacao !fechacai_usu, cmbxFechaCai
        P_LocalizaOperacao !fechames_usu, cmbxFechaMes
        P_LocalizaOperacao !contrpre_usu, cmbxContrPre
        P_LocalizaOperacao !alterarvlrparcela_usu, cmbxAlterarVlrParcela
        P_LocalizaOperacao !excluircheque_usu, cmbxExclCheque
        P_LocalizaOperacao !prorrogarparcela_usu, cmbxProrrogarParcela
        P_LocalizaOperacao !efetivarvenda_usu, cmbxEfetVenda
        P_LocalizaOperacao !crediario_usu, cmbxCrediario
        P_LocalizaOperacao !visualizarvenda_usu, cmbxVisualizarVenda
        P_LocalizaOperacao !lancarfin_usu, cmbxLanFin
        P_LocalizaOperacao !editarfin_usu, cmbxEdiFin
        P_LocalizaOperacao !transfinterna_usu, cmbxTransfInt
        
        P_LocalizaOperacao !efetuardevcons_usu, cmbxDevCons
        P_LocalizaOperacao !efetuartransfcons_usu, cmbxTransfCons
        P_LocalizaOperacao !efetuaracerto_usu, cmbxAcertoCons
        P_LocalizaOperacao !excluiracerto_usu, cmbxExclAcertoCons
        P_LocalizaOperacao !alterarcomrepr_usu, cmbxAlteraComsRepr
        P_LocalizaOperacao !baixarcomrepr_usu, cmbxBaixaComRepr
        P_LocalizaOperacao !cadastrarrepr_usu, cmbxCadRepr
        P_LocalizaOperacao !alterarterceirodoc_usu, cmbxAlterarTerceiroDoc
        P_LocalizaOperacao !excluirfin_usu, cmbxExclFin
        P_LocalizaOperacao !cadastrarfun_usu, cmbxCadFun
        P_LocalizaOperacao !cadastrarfor_usu, cmbxCadFor
        P_LocalizaOperacao !liberarcliente_usu, cmbxLiberarCli
        
        P_LocalizaOperacao !desbloquearvenda_usu, cmbxDesbloquearVenda
        
        
        
        If !visaocaixa_usu = "S" Then
            chkCaixa.Value = 1
        End If
        If (!visaoconsignado_usu = "S") Then
            chkConsignado.Value = 1
        End If
        
        
        If Not IsNull(!visaoestoque_usu) Then
            If !visaocrediario_usu = "S" Then
                chkCrediario.Value = 1
            End If
            
        End If
        If !visaoestoque_usu = "S" Then
            chkEstoque.Value = 1
        End If
        If (!visaofinanceiro_usu = "S") Then
            chkFinanceiro.Value = 1
        End If
        If !visaovendas_usu = "S" Then
            chkVendas.Value = 1
        End If
        If !visaopacote_usu = "S" Then
            chkPacote.Value = 1
        End If
        If !visaogeral_usu = "S" Then
            chkGeral.Value = 1
        End If
        
                
        cmbxDesconto.text = !desconto_usu
        cmbxDescontoBaixa.text = !descontobaixa_usu
        'P_LocalizaOperacao
        
        
        
    End With
End Sub
Sub grava()
      Dim strSQL As String
      On Error GoTo trata_Erro
    strSQL = "delete from tab_defusu where cod_usu='" & Me.tag & "'"
    Conn.Execute strSQL
    strSQL = "insert into tab_defusu(abrevenda_usu,"
    
    
    
    
    strSQL = strSQL & "abrenota_usu,"
    
    strSQL = strSQL & "lancaixa_usu,"
    strSQL = strSQL & "landoc_usu,"
    strSQL = strSQL & "lanche_usu,"
    strSQL = strSQL & "baixadoc_usu,"
    strSQL = strSQL & "baixache_usu,"
    strSQL = strSQL & "edicai_usu,"
    strSQL = strSQL & "exclcai_usu,"
    strSQL = strSQL & "excldoc_usu,"
    strSQL = strSQL & "cadprod_usu,"
    strSQL = strSQL & "altprod_usu,"
    strSQL = strSQL & "altpreco_usu,"
    strSQL = strSQL & "cadusu_usu,"
    strSQL = strSQL & "cadcli_usu,"
    strSQL = strSQL & "cadcco_usu,"
    strSQL = strSQL & "cadformaspag_usu,"
    strSQL = strSQL & "altformaspag_usu,"
    strSQL = strSQL & "exclvenda_usu,"
    strSQL = strSQL & "viscusto_usu,"
    strSQL = strSQL & "visrelest_usu,"
    strSQL = strSQL & "visrelfin_usu,"
    strSQL = strSQL & "fechacai_usu,"
    strSQL = strSQL & "fechames_usu,"
    strSQL = strSQL & "contrpre_usu,"
    strSQL = strSQL & "alterarvlrparcela_usu,"
    strSQL = strSQL & "prorrogarparcela_usu,"
    strSQL = strSQL & "alterarterceirodoc_usu,"
    strSQL = strSQL & "baixarcheque_usu,"
    strSQL = strSQL & "excluircheque_usu,"
    strSQL = strSQL & "efetivarvenda_usu,"
    strSQL = strSQL & "crediario_usu,"
    strSQL = strSQL & "visualizarvenda_usu,"
    strSQL = strSQL & "lancarfin_usu,"
    strSQL = strSQL & "editarfin_usu,"
    strSQL = strSQL & "excluirfin_usu,"
    strSQL = strSQL & "transfinterna_usu,"
    strSQL = strSQL & "efetuardevcons_usu,"
    strSQL = strSQL & "efetuartransfcons_usu,"
    strSQL = strSQL & "efetuaracerto_usu,"
    strSQL = strSQL & "excluiracerto_usu,"
    strSQL = strSQL & "alterarcomrepr_usu,"
    strSQL = strSQL & "baixarcomrepr_usu,"
    strSQL = strSQL & "cadastrarrepr_usu,"
    strSQL = strSQL & "cadastrarfun_usu,"
    strSQL = strSQL & "cadastrarfor_usu,"
    strSQL = strSQL & "desconto_usu,"
    strSQL = strSQL & "descontobaixa_usu,"
    strSQL = strSQL & "corrigeestcons_usu,"
    strSQL = strSQL & "visaocaixa_usu,"
    strSQL = strSQL & "visaoconsignado_usu,"
    strSQL = strSQL & "visaocrediario_usu,"
    strSQL = strSQL & "visaovendas_usu,"
    strSQL = strSQL & "visaopacote_usu,"
    strSQL = strSQL & "visaogeral_usu,"
    strSQL = strSQL & "visaofinanceiro_usu,"
    strSQL = strSQL & "visaoestoque_usu,"
    strSQL = strSQL & "liberarcliente_usu,"
    strSQL = strSQL & "Desbloquearvenda_usu,"
    
    strSQL = strSQL & "cod_usu)values("
    strSQL = strSQL & "'" & Left(cmbxAbreVenda.text, 1) & "', "
    



    strSQL = strSQL & "'" & Left(cmbxAbreNota.text, 1) & "',"
        
    strSQL = strSQL & "'" & Left(cmbxLanCai.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanDoc.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanChe.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarDoc.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarChe.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEdiCai.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclCai.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclDoc.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadPro.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltPro.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltPrecos.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadUsu.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadCli.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadCCO.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadFormasPag.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltFormasPag.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclVenda.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisCusto.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisRel.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisFin.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxFechaCai.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxFechaMes.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxContrPre.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAlterarVlrParcela.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxProrrogarParcela.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAlterarTerceiroDoc.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarChe.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclCheque.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEfetVenda.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCrediario.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisualizarVenda.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanFin.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEdiFin.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclFin.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxTransfInt.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxDevCons.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxTransfCons.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAcertoCons.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclAcertoCons.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAlteraComsRepr.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixaComRepr.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadRepr.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadFun.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadFor.text, 1) & "',"
    strSQL = strSQL & F_GravaDec(cmbxDesconto.text) & ","
    strSQL = strSQL & F_GravaDec(cmbxDescontoBaixa.text) & ","
    strSQL = strSQL & "'" & Left(cmbxCorrigeEstCons.text, 1) & "',"
    
    strSQL = strSQL & IIf(chkCaixa.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkConsignado.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkCrediario.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkVendas.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkPacote.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkGeral.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkFinanceiro.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkEstoque.Value = 1, "'S'", "'N'") & ","
    
    strSQL = strSQL & "'" & Left(cmbxLiberarCli.text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxDesbloquearVenda.text, 1) & "',"
    strSQL = strSQL & "'" & Me.tag & "') "
    Conn.Execute strSQL
    MsgBox "Definições de Usuário Gravadas!", vbInformation
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub altera()
    
End Sub

Private Sub Form_Load()
    bActive = False
    stb.Tab = 0
End Sub

