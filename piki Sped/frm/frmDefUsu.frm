VERSION 5.00
Begin VB.Form frmDefUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definições de Usuário"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   Icon            =   "frmDefUsu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   13305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnSalvar 
      Caption         =   "&Confirmar"
      Height          =   735
      Left            =   120
      Picture         =   "frmDefUsu.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7680
      Width           =   3015
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   6360
      Picture         =   "frmDefUsu.frx":0C54
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   7680
      Width           =   3135
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
      Height          =   3855
      Left            =   9600
      TabIndex        =   111
      Top             =   0
      Width           =   3015
      Begin VB.ComboBox cmbxCancelarVenda 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0DBF
         Left            =   1920
         List            =   "frmDefUsu.frx":0DC9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox cmbxCancelarItem 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0DD7
         Left            =   1920
         List            =   "frmDefUsu.frx":0DE1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox cmbxAbreVenda 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0DEF
         Left            =   1920
         List            =   "frmDefUsu.frx":0DF9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxEfetVenda 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0E07
         Left            =   1920
         List            =   "frmDefUsu.frx":0E11
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbxAbreNota 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0E1F
         Left            =   1920
         List            =   "frmDefUsu.frx":0E29
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbxVisualizarVenda 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0E37
         Left            =   1920
         List            =   "frmDefUsu.frx":0E41
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbxExclVenda 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0E4F
         Left            =   1920
         List            =   "frmDefUsu.frx":0E59
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbxDesconto 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0E67
         Left            =   1920
         List            =   "frmDefUsu.frx":0E7A
         Sorted          =   -1  'True
         TabIndex        =   49
         Text            =   "cmbxDesconto"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbxCrediario 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0E92
         Left            =   1920
         List            =   "frmDefUsu.frx":0E9C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   120
         Top             =   3120
         Width           =   1140
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancela Item"
         Height          =   195
         Left            =   120
         TabIndex        =   119
         Top             =   2760
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   118
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efetivar Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   117
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir Nota Fiscal"
         Height          =   195
         Left            =   120
         TabIndex        =   116
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visualizar Vendas"
         Height          =   195
         Left            =   120
         TabIndex        =   115
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto Venda"
         Height          =   195
         Left            =   120
         TabIndex        =   113
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crediário"
         Height          =   195
         Left            =   120
         TabIndex        =   112
         Top             =   2400
         Width           =   615
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
      Height          =   1815
      Left            =   6360
      TabIndex        =   36
      Top             =   0
      Width           =   3135
      Begin VB.ComboBox cmbxLanCai 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0EAA
         Left            =   2040
         List            =   "frmDefUsu.frx":0EB4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxFechaCai 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0EC2
         Left            =   2040
         List            =   "frmDefUsu.frx":0ECC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbxEdiCai 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0EDA
         Left            =   2040
         List            =   "frmDefUsu.frx":0EE4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbxExclCai 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0EF2
         Left            =   2040
         List            =   "frmDefUsu.frx":0EFC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fazer Lanç. Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   110
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fechar Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Editar Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   1380
         Width           =   900
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
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ComboBox cmbxLanDoc 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0F0A
         Left            =   2040
         List            =   "frmDefUsu.frx":0F14
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxBaixarDoc 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0F22
         Left            =   2040
         List            =   "frmDefUsu.frx":0F2C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbxExclDoc 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0F3A
         Left            =   2040
         List            =   "frmDefUsu.frx":0F44
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbxBaixarChe 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0F52
         Left            =   2040
         List            =   "frmDefUsu.frx":0F5C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbxLanChe 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0F6A
         Left            =   2040
         List            =   "frmDefUsu.frx":0F74
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbxCadFormasPag 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0F82
         Left            =   2040
         List            =   "frmDefUsu.frx":0F8C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox cmbxCadCCO 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0F9A
         Left            =   2040
         List            =   "frmDefUsu.frx":0FA4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox cmbxAltFormasPag 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0FB2
         Left            =   2040
         List            =   "frmDefUsu.frx":0FBC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3480
         Width           =   855
      End
      Begin VB.ComboBox cmbxAlterarVlrParcela 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0FCA
         Left            =   2040
         List            =   "frmDefUsu.frx":0FD4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbxProrrogarParcela 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0FE2
         Left            =   2040
         List            =   "frmDefUsu.frx":0FEC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox cmbxAlterarTerceiroDoc 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":0FFA
         Left            =   2040
         List            =   "frmDefUsu.frx":1004
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4200
         Width           =   855
      End
      Begin VB.ComboBox cmbxExclCheque 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1012
         Left            =   2040
         List            =   "frmDefUsu.frx":101C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox cmbxLanFin 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":102A
         Left            =   2040
         List            =   "frmDefUsu.frx":1034
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4560
         Width           =   855
      End
      Begin VB.ComboBox cmbxEdiFin 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1042
         Left            =   2040
         List            =   "frmDefUsu.frx":104C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4920
         Width           =   855
      End
      Begin VB.ComboBox cmbxExclFin 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":105A
         Left            =   2040
         List            =   "frmDefUsu.frx":1064
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   5280
         Width           =   855
      End
      Begin VB.ComboBox cmbxDescontoBaixa 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1072
         Left            =   2040
         List            =   "frmDefUsu.frx":1085
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "cmbxDesconto"
         Top             =   5640
         Width           =   855
      End
      Begin VB.ComboBox cmbxLiberarCli 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":109D
         Left            =   2040
         List            =   "frmDefUsu.frx":10A7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fazer Lanç. Doc"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baixar Doc."
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir Doc."
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baixar Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fazer Lanç. Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   104
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Frms Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   103
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar C. Corr."
         Height          =   195
         Left            =   120
         TabIndex        =   102
         Top             =   3120
         Width           =   1245
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar. Formas Pgto"
         Height          =   195
         Left            =   120
         TabIndex        =   101
         Top             =   3480
         Width           =   1425
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar Valor Parcela"
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prorrogar Parcela:"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   3840
         Width           =   1275
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar Terceiro Parcela:"
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   4200
         Width           =   1710
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lançar Financeiro"
         Height          =   195
         Left            =   120
         TabIndex        =   96
         Top             =   4560
         Width           =   1275
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Editar Financeiro"
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   4920
         Width           =   1185
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir Financeiro"
         Height          =   195
         Left            =   120
         TabIndex        =   94
         Top             =   5280
         Width           =   1245
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto Baixa Parc."
         Height          =   195
         Left            =   120
         TabIndex        =   93
         Top             =   5640
         Width           =   1545
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Liberar/Bloquear Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   6000
         Width           =   1710
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
      Height          =   2775
      Left            =   3120
      TabIndex        =   85
      Top             =   4800
      Width           =   3135
      Begin VB.ComboBox cmbxCadPro 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":10B5
         Left            =   1920
         List            =   "frmDefUsu.frx":10BF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxAltPro 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":10CD
         Left            =   1920
         List            =   "frmDefUsu.frx":10D7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbxAltPrecos 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":10E5
         Left            =   1920
         List            =   "frmDefUsu.frx":10EF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbxTransfInt 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":10FD
         Left            =   1920
         List            =   "frmDefUsu.frx":1107
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   90
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   660
         Width           =   1050
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar Preços"
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transferência Interna"
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   1380
         Width           =   1515
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
      Height          =   2055
      Left            =   3240
      TabIndex        =   79
      Top             =   2760
      Width           =   3015
      Begin VB.ComboBox cmbxVisFin 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1115
         Left            =   1920
         List            =   "frmDefUsu.frx":111F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbxVisRel 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":112D
         Left            =   1920
         List            =   "frmDefUsu.frx":1137
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbxVisCusto 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1145
         Left            =   1920
         List            =   "frmDefUsu.frx":114F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxContrPre 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":115D
         Left            =   1920
         List            =   "frmDefUsu.frx":1167
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbxFechaMes 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1175
         Left            =   1920
         List            =   "frmDefUsu.frx":117F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vis. Relat. Financ."
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   1020
         Width           =   1290
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vis. Relat. de Est."
         Height          =   195
         Left            =   120
         TabIndex        =   83
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visualizar Custos"
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controlar Preços"
         Height          =   195
         Left            =   120
         TabIndex        =   81
         Top             =   1740
         Width           =   1170
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fechar Mês"
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1380
         Width           =   840
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
      Height          =   2055
      Left            =   6360
      TabIndex        =   73
      Top             =   1800
      Width           =   3135
      Begin VB.ComboBox cmbxCadCli 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":118D
         Left            =   2040
         List            =   "frmDefUsu.frx":1197
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbxCadUsu 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":11A5
         Left            =   2040
         List            =   "frmDefUsu.frx":11AF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxCadFun 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":11BD
         Left            =   2040
         List            =   "frmDefUsu.frx":11C7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbxCadFor 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":11D5
         Left            =   2040
         List            =   "frmDefUsu.frx":11DF
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbxCadRepr 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":11ED
         Left            =   2040
         List            =   "frmDefUsu.frx":11F7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Clientes"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Usuários"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Funcionários"
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   1020
         Width           =   1620
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Fornecedores"
         Height          =   195
         Left            =   120
         TabIndex        =   75
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastrar Representante"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   1740
         Width           =   1770
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
      Height          =   2775
      Left            =   3240
      TabIndex        =   61
      Top             =   0
      Width           =   3015
      Begin VB.ComboBox cmbxTransfCons 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1205
         Left            =   1920
         List            =   "frmDefUsu.frx":120F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbxDevCons 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":121D
         Left            =   1920
         List            =   "frmDefUsu.frx":1227
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbxAcertoCons 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1235
         Left            =   1920
         List            =   "frmDefUsu.frx":123F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbxAlteraComsRepr 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":124D
         Left            =   1920
         List            =   "frmDefUsu.frx":1257
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbxBaixaComRepr 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1265
         Left            =   1920
         List            =   "frmDefUsu.frx":126F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbxExclAcertoCons 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":127D
         Left            =   1920
         List            =   "frmDefUsu.frx":1287
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbxCorrigeEstCons 
         Height          =   315
         ItemData        =   "frmDefUsu.frx":1295
         Left            =   1920
         List            =   "frmDefUsu.frx":129F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efetuar Transferência"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efetuar Devolução:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efetuar Acerto"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar Comissão Repr."
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baixar Comissão Repr."
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   2040
         Width           =   1590
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir Acerto"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corrigir Estoque Cons:"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   2400
         Width           =   1560
      End
   End
   Begin VB.Frame frVisoes 
      Caption         =   "Visões"
      Height          =   1935
      Left            =   9720
      TabIndex        =   53
      Top             =   5640
      Width           =   3135
      Begin VB.CheckBox chkCaixa 
         Caption         =   "Caixa"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkConsignado 
         Caption         =   "Consignado"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkCrediario 
         Caption         =   "Crediário"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkFinanceiro 
         Caption         =   "Financeiro"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkPacote 
         Caption         =   "Pacote"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkVendas 
         Caption         =   "Vendas"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox chkGeral 
         Caption         =   "Geral"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkEstoque 
         Caption         =   "Estoque"
         Height          =   195
         Left            =   1800
         TabIndex        =   62
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdVisao 
      Caption         =   "&Visão"
      Height          =   735
      Left            =   3240
      Picture         =   "frmDefUsu.frx":12AD
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7680
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
    PerdeFoco cmbxDesconto
End Sub

Private Sub cmbxDescontoBaixa_LostFocus()
    PerdeFoco cmbxDesconto, "###"
End Sub

Private Sub cmdVisao_Click()
 '   frmVisao.Show 1
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
        'P_Buscadef
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
    strSQL = "select * from " & TAB_DEFUSU & "  where cod_usu = '" & Me.tag & "' "
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
        P_LocalizaOperacao !cancelarvenda_usu, cmbxCancelarVenda
        P_LocalizaOperacao !cancelaritem_usu, cmbxCancelarItem
        
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
        
                
        cmbxDesconto.Text = !desconto_usu
        cmbxDescontoBaixa.Text = !descontobaixa_usu
        'P_LocalizaOperacao
        
        
        
    End With
End Sub
Sub grava()
      Dim strSQL As String
      On Error GoTo trata_Erro
    strSQL = "delete from " & TAB_DEFUSU & "  where cod_usu='" & Me.tag & "'"
    conn.Execute strSQL
    strSQL = "insert into " & TAB_DEFUSU & " (abrevenda_usu,"
    
    
    
    
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
    strSQL = strSQL & "cancelarvenda_usu,"
    strSQL = strSQL & "cancelaritem_usu,"
    strSQL = strSQL & "cod_usu)values("
    strSQL = strSQL & "'" & Left(cmbxAbreVenda.Text, 1) & "', "
    



    strSQL = strSQL & "'" & Left(cmbxAbreNota.Text, 1) & "',"
        
    strSQL = strSQL & "'" & Left(cmbxLanCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanDoc.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanChe.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarDoc.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarChe.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEdiCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclDoc.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadPro.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltPro.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltPrecos.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadUsu.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadCli.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadCCO.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadFormasPag.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAltFormasPag.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclVenda.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisCusto.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisRel.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisFin.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxFechaCai.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxFechaMes.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxContrPre.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAlterarVlrParcela.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxProrrogarParcela.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAlterarTerceiroDoc.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixarChe.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclCheque.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEfetVenda.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCrediario.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxVisualizarVenda.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxLanFin.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxEdiFin.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclFin.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxTransfInt.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxDevCons.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxTransfCons.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAcertoCons.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxExclAcertoCons.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxAlteraComsRepr.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxBaixaComRepr.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadRepr.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadFun.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCadFor.Text, 1) & "',"
    strSQL = strSQL & F_GravaDec(cmbxDesconto.Text) & ","
    strSQL = strSQL & F_GravaDec(cmbxDescontoBaixa.Text) & ","
    strSQL = strSQL & "'" & Left(cmbxCorrigeEstCons.Text, 1) & "',"
    
    strSQL = strSQL & IIf(chkCaixa.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkConsignado.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkCrediario.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkVendas.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkPacote.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkGeral.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkFinanceiro.Value = 1, "'S'", "'N'") & ","
    strSQL = strSQL & IIf(chkEstoque.Value = 1, "'S'", "'N'") & ","
    
    strSQL = strSQL & "'" & Left(cmbxLiberarCli.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCancelarVenda.Text, 1) & "',"
    strSQL = strSQL & "'" & Left(cmbxCancelarItem.Text, 1) & "',"
    strSQL = strSQL & "'" & Me.tag & "') "
    conn.Execute strSQL
    MsgBox "Definições de Usuário Gravadas!", vbInformation
    
    Unload Me
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub altera()
    
End Sub

Private Sub Form_Load()
    bActive = False
End Sub


