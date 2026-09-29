VERSION 5.00
Begin VB.Form frmConfigNF 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuração de NF"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndbtnCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   10080
      TabIndex        =   167
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnAltera 
      Caption         =   "Salvar Alterações"
      Height          =   495
      Left            =   10080
      TabIndex        =   166
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdBaixo 
      Height          =   375
      Left            =   10440
      Picture         =   "frmConfigNF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmdEsq 
      Height          =   375
      Left            =   10080
      Picture         =   "frmConfigNF.frx":0311
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdDir 
      Height          =   375
      Left            =   10800
      Picture         =   "frmConfigNF.frx":0626
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdCima 
      Height          =   375
      Left            =   10440
      Picture         =   "frmConfigNF.frx":0922
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtMover 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   161
      Text            =   "1"
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   3075
      TabIndex        =   158
      Top             =   6840
      Width           =   3135
      Begin VB.TextBox txtDistProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   160
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Distância entre Produtos"
         Height          =   195
         Left            =   240
         TabIndex        =   159
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1185
      ScaleWidth      =   3105
      TabIndex        =   148
      Top             =   5280
      Width           =   3135
      Begin VB.TextBox txtNumX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   157
         Text            =   "0"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtNumY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   156
         Text            =   "0"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtNumRodaPeX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   155
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtNumRodaPeY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   154
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtMarcaX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   153
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtMarcaY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   152
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Roda Pé:"
         Height          =   195
         Left            =   120
         TabIndex        =   151
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   150
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca X"
         Height          =   195
         Left            =   120
         TabIndex        =   149
         Top             =   840
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   6840
      ScaleHeight     =   7425
      ScaleWidth      =   2865
      TabIndex        =   97
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtDadosAddX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   147
         Text            =   "0"
         Top             =   5640
         Width           =   495
      End
      Begin VB.TextBox txtDadosAddY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   146
         Text            =   "0"
         Top             =   5640
         Width           =   495
      End
      Begin VB.TextBox txtPliqFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   145
         Text            =   "0"
         Top             =   5280
         Width           =   495
      End
      Begin VB.TextBox txtPliqFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   144
         Text            =   "0"
         Top             =   5280
         Width           =   495
      End
      Begin VB.TextBox txtPbFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   143
         Text            =   "0"
         Top             =   4920
         Width           =   495
      End
      Begin VB.TextBox txtPBFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   142
         Text            =   "0"
         Top             =   4920
         Width           =   495
      End
      Begin VB.TextBox txtNomeFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   139
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtNomeFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   138
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtTipoFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   137
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtTipoFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   136
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtPlacaFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   135
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtPlacaFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   134
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtUfVeiX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   133
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtUfVeiY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   132
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtCNPJFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   131
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtCNPJFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   130
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtEndFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   129
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtEndFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   128
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtMuniFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   127
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtMuniFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   126
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtUfFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   125
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtUfFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   124
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtIEFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   123
         Text            =   "0"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtIeFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   122
         Text            =   "0"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtQtdFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   121
         Text            =   "0"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtQtdFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   120
         Text            =   "0"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtEspFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   119
         Text            =   "0"
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox txtEspFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   118
         Text            =   "0"
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox txtMarcafreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   117
         Text            =   "0"
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox txtMarcaFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   116
         Text            =   "0"
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox txtNumFreX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   115
         Text            =   "0"
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox txtNumFreY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   114
         Text            =   "0"
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Left            =   1560
         TabIndex        =   141
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         Height          =   195
         Left            =   2160
         TabIndex        =   140
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   113
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   112
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   111
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF Veículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   110
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ Frete"
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   108
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Município Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   107
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IE Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   3120
         Width           =   600
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTD Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   104
         Top             =   3480
         Width           =   795
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   103
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   102
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   101
         Top             =   4560
         Width           =   1005
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Bruto:"
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   4920
         Width           =   825
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Líquido:"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   5280
         Width           =   990
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dados ADD:"
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   5640
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   3600
      ScaleHeight     =   3825
      ScaleWidth      =   3105
      TabIndex        =   64
      Top             =   3720
      Width           =   3135
      Begin VB.TextBox txtBcX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   94
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtBCY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   93
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtVlrIcmsX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   92
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtVlrIcmsY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   91
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtBcSubsX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   90
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtBCSubsY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   89
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtVlrSubsX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   88
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtVlrSubsY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   87
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtVlrProdX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   86
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtVlrProdY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   85
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtVlrFreteX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   84
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtVlrFreteY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   83
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtVlrSegX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   82
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtVlrSegY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   81
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtOutrasDespX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   80
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtOutrasDespY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   79
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtVlrIpiX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   78
         Text            =   "0"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtVlrIpiY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   77
         Text            =   "0"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtVlrTotalx 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   76
         Text            =   "0"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtVlrTotalY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   75
         Text            =   "0"
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Left            =   2040
         TabIndex        =   96
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         Height          =   195
         Left            =   2640
         TabIndex        =   95
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Cálculo:"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor ICMS"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Cálculo Subst:"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Substituição:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Produtos:"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Seguro:"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outras Despesas:"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor IPI:"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   3480
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   3600
      ScaleHeight     =   3465
      ScaleWidth      =   3105
      TabIndex        =   43
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtAliqProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   63
         Text            =   "0"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtTotalProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   62
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtCodProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   61
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtDescProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   60
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtStProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   59
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtUnid 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   58
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtQtdProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   57
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtUnitProd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   56
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtIniProdY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   53
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Left            =   2040
         TabIndex        =   55
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         Height          =   195
         Left            =   2640
         TabIndex        =   54
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Início dos Produtos:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Produto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição Produto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ST Produto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade Produto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   870
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alíquota:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0071CACE&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   5025
      ScaleWidth      =   3105
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtNatopx 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtNatopY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   26
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtCfopx 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtCfopY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   24
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtIex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIey 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   22
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtNomeDestX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   21
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtNomeDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtCnpjX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtCnpjY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtEmissX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtEmissY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Text            =   "0"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtEndDestx 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtEndDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtBaidestX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtBaiDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtCepDestX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Text            =   "0"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtCepDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Text            =   "0"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtMuniDestX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Text            =   "0"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtMuniDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Text            =   "0"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtFoneDestX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Text            =   "0"
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox txtFoneDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Text            =   "0"
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox txtUfDestX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "0"
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox txtUfDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Text            =   "0"
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox txtIeDestX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Text            =   "0"
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox txtIeDestY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Text            =   "0"
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Natureza da Operação:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   195
         Left            =   2040
         TabIndex        =   41
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         Height          =   195
         Left            =   2640
         TabIndex        =   40
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IE Sub. Trib:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Destinatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2760
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cep Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Município Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   1125
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   4200
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IE Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   4560
         Width           =   570
      End
   End
   Begin VB.CommandButton cmndbtnFechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   10080
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
End
Attribute VB_Name = "frmConfigNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBaixo_Click()
    moveVert CInt("-" & txtMover.Text)
    
End Sub
Sub moveVert(valor As Integer)
    
        txtMarcaY.Text = Int(txtMarcaY.Text) + valor
        txtNatopY.Text = Int(txtNatopY.Text) + valor
        
        txtBaiDestY.Text = Int(txtBaiDestY.Text) + valor
        txtCfopY.Text = Int(txtCfopY.Text) + valor
        txtIey.Text = Int(txtIey.Text) + valor
        txtNomeDestY.Text = Int(txtNomeDestY.Text) + valor
        txtCnpjY.Text = Int(txtCnpjY.Text) + valor
        txtEmissY.Text = Int(txtEmissY.Text) + valor
        txtEndDestY.Text = Int(txtEndDestY.Text) + valor
        
        txtCepDestY.Text = Int(txtCepDestY.Text) + valor
        txtMuniDestY.Text = Int(txtMuniDestY.Text) + valor
        txtFoneDestY.Text = Int(txtFoneDestY.Text) + valor
        txtUfDestY.Text = Int(txtUfDestY.Text) + valor
        txtIeDestY.Text = Int(txtIeDestY.Text) + valor
        txtNumY.Text = Int(txtNumY.Text) + valor
        txtNumRodaPeY.Text = Int(txtNumRodaPeY.Text) + valor
        
        txtIniProdY.Text = Int(txtIniProdY.Text) + valor
        
        txtBCY.Text = Int(txtBCY.Text) + valor
        txtVlrIcmsY.Text = Int(txtVlrIcmsY.Text) + valor
        txtBCSubsY.Text = Int(txtBCSubsY.Text) + valor
        txtVlrSubsY.Text = Int(txtVlrSubsY.Text) + valor
        txtVlrProdY.Text = Int(txtVlrProdY.Text) + valor
        txtVlrFreteY.Text = Int(txtVlrFreteY.Text) + valor
        txtVlrSegY.Text = Int(txtVlrSegY) + valor
        txtOutrasDespY.Text = Int(txtOutrasDespY.Text) + valor
        txtVlrIpiY.Text = Int(txtVlrIpiY.Text) + valor
        txtVlrTotalY.Text = Int(txtVlrTotalY.Text) + valor
        txtNomeFreY.Text = Int(txtNomeFreY.Text) + valor
        txtTipoFreY.Text = Int(txtTipoFreY.Text) + valor
        txtCNPJFreY.Text = Int(txtCNPJFreY.Text) + valor
        txtMarcaFreY.Text = Int(txtMarcaFreY.Text) + valor
        txtPlacaFreY.Text = Int(txtPlacaFreY.Text) + valor
        txtUfVeiY.Text = Int(txtUfVeiY.Text) + valor
        txtIeFreY.Text = Int(txtIeFreY.Text) + valor
        txtUfFreY.Text = Int(txtUfFreY.Text) + valor
        txtEndFreY.Text = Int(txtEndFreY.Text) + valor
        txtMuniFreY.Text = Int(txtMuniFreY.Text) + valor
        txtQtdFreY.Text = Int(txtQtdFreY.Text) + valor
        txtEspFreY.Text = Int(txtEspFreY.Text) + valor
        txtNumFreY.Text = Int(txtNumFreY.Text) + valor
        txtDadosAddY.Text = Int(txtDadosAddY.Text) + valor
        txtPBFreY.Text = Int(txtPBFreY.Text) + valor
        txtPliqFreY.Text = Int(txtPliqFreY.Text) + valor
        
        
        
End Sub
Sub MoveHoriz(valor As Integer)
    txtMarcaX.Text = Int(txtMarcaX.Text) + valor
        txtNatopx.Text = Int(txtNatopx.Text) + valor
        
        txtBaidestX.Text = Int(txtBaidestX.Text) + valor
        txtCfopx.Text = Int(txtCfopx.Text) + valor
        txtIex.Text = Int(txtIex.Text) + valor
        txtNomeDestX.Text = Int(txtNomeDestX.Text) + valor
        txtCnpjX.Text = Int(txtCnpjX.Text) + valor
        txtEmissX.Text = Int(txtEmissX.Text) + valor
        txtEndDestx.Text = Int(txtEndDestx.Text) + valor
        
        txtCepDestX.Text = Int(txtCepDestX.Text) + valor
        txtMuniDestX.Text = Int(txtMuniDestX.Text) + valor
        txtFoneDestX.Text = Int(txtFoneDestX.Text) + valor
        txtUfDestX.Text = Int(txtUfDestX.Text) + valor
        txtIeDestX.Text = Int(txtIeDestX.Text) + valor
        txtNumX.Text = Int(txtNumX.Text) + valor
        txtNumRodaPeX.Text = Int(txtNumRodaPeX.Text) + valor
        
        'txtIniProdx.Text = Int(txtIniProdx.Text) + valor
        
        txtStProd.Text = Int(txtStProd.Text) + valor
        txtUnid.Text = Int(txtUnid.Text) + valor
        txtCodProd.Text = Int(txtCodProd.Text) + valor
        txtDescProd.Text = Int(txtDescProd.Text) + valor
        txtAliqProd.Text = Int(txtAliqProd.Text) + valor
        txtQtdProd.Text = Int(txtQtdProd.Text) + valor
        txtUnitProd.Text = Int(txtUnitProd.Text) + valor
        txtTotalProd.Text = Int(txtTotalProd.Text) + valor
        
        
        txtBcX.Text = Int(txtBcX.Text) + valor
        txtVlrIcmsX.Text = Int(txtVlrIcmsX.Text) + valor
        txtBcSubsX.Text = Int(txtBcSubsX.Text) + valor
        txtVlrSubsX.Text = Int(txtVlrSubsX.Text) + valor
        txtVlrProdX.Text = Int(txtVlrProdX.Text) + valor
        txtVlrFreteX.Text = Int(txtVlrFreteX.Text) + valor
        txtVlrSegX.Text = Int(txtVlrSegX) + valor
        txtOutrasDespX.Text = Int(txtOutrasDespX.Text) + valor
        txtVlrIpiX.Text = Int(txtVlrIpiX.Text) + valor
        txtVlrTotalx.Text = Int(txtVlrTotalx.Text) + valor
        txtNomeFreX.Text = Int(txtNomeFreX.Text) + valor
        txtTipoFreX.Text = Int(txtTipoFreX.Text) + valor
        txtCNPJFreX.Text = Int(txtCNPJFreX.Text) + valor
        txtMarcafreX.Text = Int(txtMarcafreX.Text) + valor
        txtPlacaFreX.Text = Int(txtPlacaFreX.Text) + valor
        txtUfVeiX.Text = Int(txtUfVeiX.Text) + valor
        txtIEFreX.Text = Int(txtIEFreX.Text) + valor
        txtUfFreX.Text = Int(txtUfFreX.Text) + valor
        txtEndFreX.Text = Int(txtEndFreX.Text) + valor
        txtMuniFreX.Text = Int(txtMuniFreX.Text) + valor
        txtQtdFreX.Text = Int(txtQtdFreX.Text) + valor
        txtEspFreX.Text = Int(txtEspFreX.Text) + valor
        txtNumFreX.Text = Int(txtNumFreX.Text) + valor
        txtDadosAddX.Text = Int(txtDadosAddX.Text) + valor
        txtPbFreX.Text = Int(txtPbFreX.Text) + valor
        txtPliqFreX.Text = Int(txtPliqFreX.Text) + valor
        
End Sub
Private Sub cmdCima_Click()
    moveVert Int(txtMover.Text)
End Sub

Private Sub cmdDir_Click()
    MoveHoriz txtMover.Text
End Sub

Private Sub cmdEsq_Click()
    MoveHoriz "-" & txtMover.Text
End Sub

Private Sub cmndbtnAltera_Click()
        Dim strSQL As String
        strSQL = "update tab_confignf set "
        strSQL = strSQL & "marcax_con=" & txtMarcaX.Text & ","
        strSQL = strSQL & "marcay_con=" & txtMarcaY.Text & ","
        strSQL = strSQL & "natopx_con=" & txtNatopx.Text & ","
        strSQL = strSQL & "natopy_con=" & txtNatopY.Text & ","
        strSQL = strSQL & "aliqprod_con=" & txtAliqProd.Text & ","
        strSQL = strSQL & "baidestx_con=" & txtBaidestX.Text & ","
        strSQL = strSQL & "baidesty_con=" & txtBaiDestY.Text & ","
        strSQL = strSQL & "cfopy_con=" & txtCfopY.Text & ","
        strSQL = strSQL & "cfopx_con=" & txtCfopx.Text & ","
        strSQL = strSQL & "iex_con=" & txtIex.Text & ","
        strSQL = strSQL & "iey_con=" & txtIey.Text & ","
        strSQL = strSQL & "nomedestx_con=" & txtNomeDestX.Text & ","
        strSQL = strSQL & "nomedesty_con=" & txtNomeDestY.Text & ", "
        strSQL = strSQL & " cnpjx_con=" & txtCnpjX.Text & ","
        strSQL = strSQL & " cnpjy_con=" & txtCnpjY.Text & ","
        strSQL = strSQL & " emissx_con=" & txtEmissX.Text & ","
        strSQL = strSQL & " emissy_con=" & txtEmissY.Text & ","
        strSQL = strSQL & " enddestx_con=" & txtEndDestx.Text & ","
        strSQL = strSQL & " enddesty_con=" & txtEndDestY.Text & ","
        strSQL = strSQL & " baidestx_con=" & txtBaidestX.Text & ","
        strSQL = strSQL & " baidesty_con=" & txtBaiDestY.Text & ","
        strSQL = strSQL & " cepdestx_con=" & txtCepDestX.Text & ","
        strSQL = strSQL & " cepdesty_con=" & txtCepDestY.Text & ","
        strSQL = strSQL & " munidestx_con=" & txtMuniDestX.Text & ","
        strSQL = strSQL & " munidesty_con=" & txtMuniDestY.Text & ","
        strSQL = strSQL & " fonedestx_con=" & txtFoneDestX.Text & ","
        strSQL = strSQL & " fonedesty_con=" & txtFoneDestY.Text & ","
        strSQL = strSQL & " ufdestx_con=" & txtUfDestX.Text & ","
        strSQL = strSQL & "ufdesty_con=" & txtUfDestY.Text & ","
        strSQL = strSQL & " iedestx_con=" & txtIeDestX.Text & ","
        strSQL = strSQL & " iedesty_con=" & txtIeDestY.Text & ","
        strSQL = strSQL & " numx_con=" & txtNumX.Text & ","
        strSQL = strSQL & " numy_con=" & txtNumY.Text & ","
        strSQL = strSQL & " numrodapex_con=" & txtNumRodaPeX.Text & ","
        strSQL = strSQL & " numrodapey_con=" & txtNumRodaPeY.Text & ","
        strSQL = strSQL & " distprod_con=" & txtDistProd.Text & ","
        strSQL = strSQL & " iniprody_con=" & txtIniProdY.Text & ","
        strSQL = strSQL & " codprod_con=" & txtCodProd.Text & ","
        strSQL = strSQL & " descprod_con=" & txtDescProd.Text & ","
        strSQL = strSQL & " unid_con=" & txtUnid.Text & ","
        strSQL = strSQL & " qtdprod_con=" & txtQtdProd.Text & ","
        strSQL = strSQL & " stprod_con=" & txtStProd.Text & ","
        strSQL = strSQL & " unitprod_con=" & txtUnitProd.Text & ","
        strSQL = strSQL & " totalprod_con=" & txtTotalProd.Text & ","
        strSQL = strSQL & " aliqprod_con=" & txtAliqProd.Text & ","
        strSQL = strSQL & " bcx_con=" & txtBcX.Text & ","
        strSQL = strSQL & " bcy_con=" & txtBCY.Text & ","
        strSQL = strSQL & " vlricmsx_con=" & txtVlrIcmsX.Text & ","
        strSQL = strSQL & " vlricmsy_con=" & txtVlrIcmsY.Text & ","
        strSQL = strSQL & " bcsubsx_con=" & txtBcSubsX.Text & ","
        strSQL = strSQL & " bcsubsy_con=" & txtBCSubsY.Text & ","
        strSQL = strSQL & " vlrsubsx_con=" & txtVlrSubsX.Text & ","
        strSQL = strSQL & " vlrsubsy_con=" & txtVlrSubsY.Text & ","
        strSQL = strSQL & " vlrprodx_con=" & txtVlrProdX.Text & ","
        strSQL = strSQL & " vlrprody_con=" & txtVlrProdY.Text & ","
        strSQL = strSQL & " vlrfretex_con=" & txtVlrFreteX.Text & ","
        strSQL = strSQL & " vlrfretey_con=" & txtVlrFreteY.Text & ","
        strSQL = strSQL & " vlrsegx_con=" & txtVlrSegX.Text & ","
        strSQL = strSQL & " vlrsegy_con=" & txtVlrSegY.Text & ","
        strSQL = strSQL & " outrasdespx_con=" & txtOutrasDespX.Text & ","
        strSQL = strSQL & " outrasdespy_con=" & txtOutrasDespY.Text & ","
        strSQL = strSQL & " vlripix_con=" & txtVlrIpiX.Text & ","
        strSQL = strSQL & " vlripiy_con=" & txtVlrIpiY.Text & ","
        strSQL = strSQL & " vlrtotalx_con=" & txtVlrTotalx.Text & ","
        strSQL = strSQL & " vlrtotaly_con=" & txtVlrTotalY.Text & ","
        strSQL = strSQL & " nomefrex_con=" & txtNomeFreX.Text & ","
        strSQL = strSQL & " nomefrey_con=" & txtNomeFreY.Text & ","
        strSQL = strSQL & " tipofrey_con=" & txtTipoFreX.Text & ","
        strSQL = strSQL & " tipofrey_con=" & txtTipoFreY.Text & ","
        strSQL = strSQL & " cnpjfrex_con=" & txtCNPJFreX.Text & ","
        strSQL = strSQL & " cnpjfrey_con=" & txtCNPJFreY.Text & ","
        strSQL = strSQL & " marcafrex_con=" & txtMarcafreX.Text & ","
        strSQL = strSQL & " marcafrey_con=" & txtMarcaFreY.Text & ","
        strSQL = strSQL & " placafrex_con=" & txtPlacaFreX.Text & ","
        strSQL = strSQL & " placafrey_con=" & txtPlacaFreY.Text & ","
        strSQL = strSQL & " ufveix_con=" & txtUfVeiX.Text & ","
        strSQL = strSQL & " ufveiy_con=" & txtUfVeiY.Text & ","
        strSQL = strSQL & " iefrex_con=" & txtIEFreX.Text & ","
        strSQL = strSQL & " iefrey_con=" & txtIeFreY.Text & ","
        strSQL = strSQL & " uffrex_con=" & txtUfFreX.Text & ","
        strSQL = strSQL & " uffrey_con=" & txtUfFreY.Text & ","
        strSQL = strSQL & " endfrex_con=" & txtEndFreX.Text & ","
        strSQL = strSQL & " endfrey_con=" & txtEndFreY.Text & ","
        strSQL = strSQL & " munifrex_con=" & txtMuniFreX.Text & ","
        strSQL = strSQL & " munifrey_con=" & txtMuniFreY.Text & ","
        strSQL = strSQL & " qtdfrex_con=" & txtQtdFreX.Text & ","
        strSQL = strSQL & " qtdfrey_con=" & txtQtdFreY.Text & ","
        strSQL = strSQL & " espfrex_con=" & txtEspFreX.Text & ","
        strSQL = strSQL & " espfrey_con=" & txtEspFreY.Text & ","
        strSQL = strSQL & " numfrex_con=" & txtNumFreX.Text & ","
        strSQL = strSQL & " numfrey_con=" & txtNumFreY.Text & ","
        strSQL = strSQL & " dadosaddx_con=" & txtDadosAddX.Text & ","
        strSQL = strSQL & " dadosaddy_con=" & txtDadosAddY.Text & ","
        strSQL = strSQL & " pbrutofrex_con=" & txtPbFreX.Text & ","
        strSQL = strSQL & " pbrutofrey_con=" & txtPBFreY.Text & ","
        strSQL = strSQL & " pliqfrex_con=" & txtPliqFreX.Text & ","
        strSQL = strSQL & " pliqfrey_con=" & txtPliqFreY.Text
        conn.Execute strSQL
        
    
End Sub

Private Sub cmndbtnCancelar_Click()
    busca
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    busca
    
    
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_confignf "
    
    rPesq.Open strSQL, conn_client
    With rPesq
        txtMarcaX.Text = !marcax_con
        txtMarcaY.Text = !marcay_con
        txtNatopx.Text = !natopx_con
        txtNatopY.Text = !natopy_con
        txtAliqProd.Text = !aliqprod_con
        txtBaidestX.Text = !baidestx_con
        txtBaiDestY.Text = !baidesty_con
        txtCfopY.Text = !cfopy_con
        txtCfopx.Text = !cfopx_con
        txtIex.Text = !iex_con
        txtIey.Text = !iey_con
        txtNomeDestX.Text = !nomedestx_con
        txtNomeDestY.Text = !nomedesty_con
        txtCnpjX.Text = !cnpjx_con
        txtCnpjY.Text = !cnpjy_con
        txtEmissX.Text = !emissx_con
        txtEmissY.Text = !emissy_con
        txtEndDestx.Text = !enddestx_con
        txtEndDestY.Text = !enddesty_con
        txtBaidestX.Text = !baidestx_con
        txtBaiDestY.Text = !baidesty_con
        txtCepDestX.Text = !cepdestx_con
        txtCepDestY.Text = !cepdesty_con
        txtMuniDestX.Text = !munidestx_con
        txtMuniDestY.Text = !munidesty_con
        txtFoneDestX.Text = !fonedestx_con
        txtFoneDestY.Text = !fonedesty_con
        txtUfDestX.Text = !ufdestx_con
        txtUfDestY.Text = !ufdesty_con
        txtIeDestX.Text = !iedestx_con
        txtIeDestY.Text = !iedesty_con
        txtNumX.Text = !numx_con
        txtNumY.Text = !numy_con
        txtNumRodaPeX.Text = !numrodapex_con
        txtNumRodaPeY.Text = !numrodapey_con
        txtDistProd.Text = !distprod_con
        txtIniProdY.Text = !iniprody_con
        txtCodProd.Text = !codprod_con
        txtDescProd.Text = !descprod_con
        txtUnid.Text = !unid_con
        txtQtdProd.Text = !qtdprod_con
        txtStProd.Text = !stprod_con
        txtUnitProd.Text = !unitprod_con
        txtTotalProd.Text = !totalprod_con
        txtAliqProd.Text = !aliqprod_con
        txtBcX.Text = !bcx_con
        txtBCY.Text = !bcy_con
        txtVlrIcmsX.Text = !vlricmsx_con
        txtVlrIcmsY.Text = !vlricmsy_con
        txtBcSubsX.Text = !bcsubsx_con
        txtBCSubsY.Text = !bcsubsy_con
        txtVlrSubsX.Text = !vlrsubsx_con
        txtVlrSubsY.Text = !vlrsubsy_con
        txtVlrProdX.Text = !vlrprodx_con
        txtVlrProdY.Text = !vlrprody_con
        txtVlrFreteX.Text = !vlrfretex_con
        txtVlrFreteY.Text = !vlrfretey_con
        txtVlrSegX.Text = !vlrsegx_con
        txtVlrSegY.Text = !vlrsegy_Con
        txtOutrasDespX.Text = !outrasdespx_con
        txtOutrasDespY.Text = !outrasdespy_con
        txtVlrIpiX.Text = !vlripix_con
        txtVlrIpiY.Text = !vlripiy_con
        txtVlrTotalx.Text = !vlrtotalx_con
        txtVlrTotalY.Text = !vlrtotaly_con
        txtNomeFreX.Text = !nomefrex_con
        txtNomeFreY.Text = !nomefrey_con
        txtTipoFreX.Text = !tipofrey_con
        txtTipoFreY.Text = !tipofrey_con
        txtCNPJFreX.Text = !cnpjfrex_con
        txtCNPJFreY.Text = !cnpjfrey_con
        txtMarcafreX.Text = !marcafrex_con
        txtMarcaFreY.Text = !marcafrey_con
        txtPlacaFreX.Text = !placafrex_con
        txtPlacaFreY.Text = !placafrey_con
        txtUfVeiX.Text = !ufveix_con
        txtUfVeiY.Text = !ufveiy_con
        txtIEFreX.Text = !iefrex_con
        txtIeFreY.Text = !iefrey_con
        txtUfFreX.Text = !uffrex_con
        txtUfFreY.Text = !uffrey_con
        txtEndFreX.Text = !endfrex_con
        txtEndFreY.Text = !endfrey_con
        txtMuniFreX.Text = !munifrex_con
        txtMuniFreY.Text = !munifrey_con
        txtQtdFreX.Text = !qtdfrex_con
        txtQtdFreY.Text = !qtdfrey_con
        txtEspFreX.Text = !espfrex_con
        txtEspFreY.Text = !espfrey_con
        txtNumFreX.Text = !numfrex_con
        txtNumFreY.Text = !numfrey_con
        txtDadosAddX.Text = !dadosaddx_con
        txtDadosAddY.Text = !dadosaddy_con
        txtPbFreX.Text = !pbrutofrex_con
        txtPBFreY.Text = !pbrutofrey_con
        txtPliqFreX.Text = !pliqfrex_con
        txtPliqFreY.Text = !pliqfrey_con
        
        
        
        
    End With
    
    
End Sub
