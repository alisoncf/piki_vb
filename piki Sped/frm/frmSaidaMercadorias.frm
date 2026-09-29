VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSaidaMercadorias 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Venda de Produtos"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10590
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Totais:"
      Height          =   1335
      Left            =   80
      TabIndex        =   47
      Top             =   5445
      Width           =   10455
      Begin VB.TextBox txtbxObservacoes_Sai 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   735
         Left            =   4560
         TabIndex        =   65
         ToolTipText     =   "Entre com as informações complementares a nota fiscal."
         Top             =   405
         Width           =   3975
      End
      Begin VB.CommandButton cmndbtnRegistraNota 
         Caption         =   "&Concluir"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Registra  a Nota Fiscal."
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnFechar 
         Caption         =   "Sai&r"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Fecha o formulário atual."
         Top             =   840
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   4440
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Label lblObservacao_Sai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informações Complementares:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   66
         Top             =   195
         Width           =   2130
      End
      Begin VB.Label lblDesconto_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         ToolTipText     =   "Informa o valor Total do desconto."
         Top             =   940
         Width           =   900
      End
      Begin VB.Label lblDesconto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   740
         Width           =   735
      End
      Begin VB.Label lblTotal_Sai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tot da Nota:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   61
         Top             =   740
         Width           =   900
      End
      Begin VB.Label lblIPI_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   60
         ToolTipText     =   "Informa o valor total do IPI lançado."
         Top             =   940
         Width           =   900
      End
      Begin VB.Label lblTotalProdutos_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   59
         ToolTipText     =   "Informa o valor total dos produtos."
         Top             =   405
         Width           =   900
      End
      Begin VB.Label lblTotalNota_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   58
         ToolTipText     =   "Informa o Valor Total lançado."
         Top             =   940
         Width           =   900
      End
      Begin VB.Label lblICMSSubst_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   57
         ToolTipText     =   "Informa o valor total lançado no ICMS Substuição."
         Top             =   940
         Width           =   900
      End
      Begin VB.Label lblBaseCalculoICMSSubst_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         ToolTipText     =   "Informa o valor total sobre a Base de Cálculo do ICMS Substuição."
         Top             =   400
         Width           =   900
      End
      Begin VB.Label lblValorICMS_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         ToolTipText     =   "Informa o valor total de ICMS sobre a Base de Cálculo."
         Top             =   405
         Width           =   900
      End
      Begin VB.Label lblBaseCalculo_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   54
         ToolTipText     =   "Informa o Valor total da Base de Cálculo."
         Top             =   400
         Width           =   900
      End
      Begin VB.Label lblIPI 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor IPI:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   53
         Top             =   735
         Width           =   645
      End
      Begin VB.Label lblValorProduto 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.dos Prod:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   52
         Top             =   195
         Width           =   960
      End
      Begin VB.Label lblICMSSubst 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS Subst:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   51
         Top             =   740
         Width           =   885
      End
      Begin VB.Label lblBaseCalculoICMSSubst 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "B.C.Icm Sub:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   50
         Top             =   200
         Width           =   930
      End
      Begin VB.Label lblValorIcms 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor ICMS:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   49
         Top             =   200
         Width           =   840
      End
      Begin VB.Label lblBaseCalculo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Base Cálculo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   80
      TabIndex        =   46
      Top             =   3840
      Width           =   10455
      Begin VB.TextBox txtbxNome_Mer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Informe o nome da mercadoria seguido de <ENTER> para localizá-lo."
         Top             =   360
         Width           =   3615
      End
      Begin VB.OptionButton OptnMercadoria 
         Caption         =   "&Código"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   68
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton OptnMercadoria 
         Caption         =   "&Mercadoria"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtbxUnitario_Its 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   26
         Text            =   "0,00"
         ToolTipText     =   "Informe o valor unitário."
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxDesconto_Its 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   39
         Text            =   "0,00"
         ToolTipText     =   "Informe o valor do desconto sobre a mercadoria."
         Top             =   960
         Width           =   900
      End
      Begin VB.TextBox txtbxICMS_Its 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         MaxLength       =   2
         TabIndex        =   32
         Text            =   "00"
         ToolTipText     =   "Informe o ICMS Sobre a Mercadoria."
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtbxIPI_Its 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   30
         Text            =   "0,00"
         ToolTipText     =   "Informe o valor do IPI sobre a mercadoria."
         Top             =   360
         Width           =   900
      End
      Begin VB.CommandButton cmndbtnExcluiItem 
         Caption         =   "E&xcluir"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Exclui a mercadoria selecionada."
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnIncluiItem 
         Caption         =   "&Adcionar"
         Height          =   375
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Inclui a mercadoria especificada à nota fiscal."
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbxBaseCalculo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSaidaMercadorias.frx":0000
         Left            =   3840
         List            =   "frmSaidaMercadorias.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   37
         ToolTipText     =   "Permite selecionar o tipo de Base de Cálculo aplicada sobre o produto."
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cmbxSituacaoTri 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   35
         ToolTipText     =   "Permite selecionar o tipo da situação tributária."
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtbxQuantidade_Its 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   24
         Text            =   "0,00"
         ToolTipText     =   "Informe a quantidade (use vírgula para separar os decimais, não ponto)."
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblDesconto_Its 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   38
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblICMS_Its 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7560
         TabIndex        =   31
         Top             =   720
         Width           =   390
      End
      Begin VB.Label lblPerc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   7920
         TabIndex        =   33
         Top             =   960
         Width           =   120
      End
      Begin VB.Label lblTotal_Its 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6480
         TabIndex        =   28
         ToolTipText     =   "Indica o valor total do produto."
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblMarca_Mer 
         Appearance      =   0  'Flat
         Caption         =   "Marca"
         ForeColor       =   &H80000007&
         Height          =   225
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Indica a Marca da mercadoria selecionada."
         Top             =   360
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Label lblSituacao2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Cálculo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   36
         Top             =   750
         Width           =   1200
      End
      Begin VB.Label lblIPI_Its 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IPI:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7560
         TabIndex        =   29
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblSituacao 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Situação Tributária:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   750
         Width           =   1380
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   27
         Top             =   120
         Width           =   405
      End
      Begin VB.Label lblUnit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   25
         Top             =   120
         Width           =   585
      End
      Begin VB.Label lblQuantidade_Its 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4440
         TabIndex        =   23
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblUnidade_Its 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         TabIndex        =   22
         ToolTipText     =   "Fornece a unidade de medida da mercadoria."
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblUn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un.:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   21
         Top             =   120
         Width           =   300
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrdItensSaida 
      Height          =   2295
      Left            =   0
      TabIndex        =   45
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      FixedCols       =   0
      BackColor       =   -2147483624
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   80
      TabIndex        =   44
      Top             =   0
      Width           =   10455
      Begin VB.OptionButton optnbtnCupomFiscal 
         Caption         =   "Cupom Fiscal"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optnbtnNotaFiscal 
         Caption         =   "Nota Fiscal"
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxNome_Cli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Informe o nome do cliente seguido de <Enter> para localizá-lo."
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtbxNota_Sai 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         MaxLength       =   8
         TabIndex        =   7
         ToolTipText     =   "Informe o número da nota fiscal."
         Top             =   360
         Width           =   1140
      End
      Begin VB.TextBox txtbxSerie_Sai 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         MaxLength       =   4
         TabIndex        =   9
         ToolTipText     =   "Informe a série da nota fiscal."
         Top             =   360
         Width           =   550
      End
      Begin VB.Label lblNome_Cli 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lblNota_Sai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Número:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6960
         TabIndex        =   6
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblSerie_Sai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Série:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8280
         TabIndex        =   8
         Top             =   120
         Width           =   405
      End
      Begin VB.Label lblUF 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblUF_Cli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frmeCabecalho2 
      Height          =   735
      Left            =   80
      TabIndex        =   43
      Top             =   720
      Width           =   10455
      Begin VB.CommandButton cmndbtnVinculaOrcamento 
         Height          =   495
         Left            =   7200
         Picture         =   "frmSaidaMercadorias.frx":0052
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vincula nota fiscal de acordo com orçamento."
         Top             =   170
         Width           =   495
      End
      Begin VB.ComboBox cmbxPercICMS_Sai 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSaidaMercadorias.frx":091C
         Left            =   4920
         List            =   "frmSaidaMercadorias.frx":091E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Selecione o Percentual do ICMS."
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxDespesas_Sai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   17
         Text            =   "0,00"
         ToolTipText     =   $"frmSaidaMercadorias.frx":0920
         Top             =   360
         Width           =   900
      End
      Begin VB.TextBox txtbxEmissao_Sai 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   13
         ToolTipText     =   "Entre com a data de emissão."
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cmbxOperacao_Sai 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Selecione o tipo de operação da nota fiscal."
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblPercIcms 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&ICMS:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4920
         TabIndex        =   14
         Top             =   120
         Width           =   435
      End
      Begin VB.Label lblDespesaNormal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Despesa:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5880
         TabIndex        =   16
         Top             =   165
         Width           =   675
      End
      Begin VB.Label lblEmissao_Ent 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Emissão:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblOperacao_Sai 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Operação:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSaidaMercadorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variáveis do Cliente
    Dim rcrdstClientes As Recordset
'Variáveis de Controle Externo
    'Indica que a tributação da mercadoria é normal
    Dim blnFlag As Boolean
    Dim blnMercadTrib As Boolean
Private Sub P_CarregaControles()
    '   Este procedimento carrega os controles conforme o orçamento
    'especificado
    
    'on error goto trataerro
    Dim rcrdstPesq As Recordset
    Dim rcrdstPesquisa As Recordset
    Dim strSQL As String
    Dim dblConstante As Double
    Dim crcQuant As Currency
    
    'Gerando SQL para consulta dos itens do orçamento
    strSQL = "Select TI.QUANTIDADE_ITO, TI.UNITARIO_ITO, TI.DESCRICAO_ITO, "
    strSQL = strSQL & "TCL.NOME_CLI, TCL.COD_CLI, TCL.UF_CLI, TOR.COD_PLA, "
    strSQL = strSQL & "TI.COD_MER, TI.MARCA_MER, TM.TIPO_MER,  "
    strSQL = strSQL & "TM.UNIDADE_MER, TC.ESTOQUE_MER, TC.SITUACAO_MER, "
    strSQL = strSQL & "TC.VALORIVA_MER, TC.NOTA_ENT "
    strSQL = strSQL & "From  as TI, "
    strSQL = strSQL & "TAB_MERCADORIAS As TM, TAB_CMPLT01 AS TC, "
    strSQL = strSQL & " as TOR, TAB_CLIENTES AS TCL "
    strSQL = strSQL & "Where TI.ORCAMENTO_ORC = '" & Me.Tag & "' And "
    strSQL = strSQL & "TOR.ORCAMENTO_ORC = '" & Me.Tag & "' And "
    strSQL = strSQL & "TCL.COD_CLI = TOR.COD_CLI And "
    strSQL = strSQL & "TI.COD_MER = TM.COD_MER And "
    strSQL = strSQL & "TI.MARCA_MER = TM.MARCA_MER And "
    strSQL = strSQL & "TI.COD_MER = TC.COD_MER And "
    strSQL = strSQL & "TI.MARCA_MER = TC.MARCA_MER "
    
    Set rcrdstPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    If rcrdstPesq.RecordCount < 1 Then
        'Não encontrou itens referente ao SQL pedido
        rcrdstPesq.Close
        MsgBox "O Orçamento " & Me.Tag & " não dispõe de itens!", vbExclamation
        Me.Tag = ""
        Exit Sub
    End If
    
    'Fazendo pesquisa para encontrar o multiplicador do plano de pagamento
    strSQL = "Select Constante_Pla From TAB_PLANOPAGAMENTOS Where Cod_Pla = '"
    strSQL = strSQL & rcrdstPesq!COD_PLA & "'"
    
    Set rcrdstPesquisa = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If rcrdstPesquisa.RecordCount < 1 Then
       'Não encontrou então elemento neutro (1)
       dblConstante = 1
    Else
       dblConstante = rcrdstPesquisa!Constante_Pla
    End If
    rcrdstPesquisa.Close
    
    
    
    With rcrdstPesq
        'Preenchendo os dados do cliente
        txtbxNome_Cli.Text = !nome_cli
        lblUF_Cli.caption = !UF_Cli
        txtbxNome_Cli.Tag = !cod_cli
        'Ajustando o ICMS
        optnbtnNotaFiscal_Click
    End With
    With msflxgrdItensSaida
        'Iniciando o preenchimento do grid
        strSQL = "" 'Utilizado para indicar as mercadorias não disponíveis
        .Rows = 1
        
        Do While Not rcrdstPesq.EOF
            .AddItem rcrdstPesq!cod_Mer
            .Row = .Rows - 1 'Incluindo o item
            crcQuant = 0
            If rcrdstPesq!Estoque_Mer > 0 Then
                .Col = 1 'Descrição do item
                .Text = rcrdstPesq!Descricao_ITO
                .Col = 2 'Situação tributária
                If rcrdstPesq!Situacao_MER <> "T" And lblUF_Cli = "GO" Then
                    'Não é Tributada e é ESTADO de Goiás
                    .Text = "01" 'Situação tributária
                    .Col = 11 'Base de Cálculo
                    .Text = "01"
                    .Col = 7 'ICMS
                    .Text = "00"
                Else
                    'É tributada ou não é estado de Goiás
                    .Text = "00" 'Situação tributária
                    .Col = 11
                    .Text = "00"
                    .Col = 7 'ICMS
                    .Text = Left(cmbxPercICMS_Sai.Text, 2)
                End If
                
                .Col = 3 'Unidade da mercadoria
                .Text = rcrdstPesq!Unidade_Mer
                .Col = 4 'Quantidade da mercadoria
                crcQuant = 0
                If rcrdstPesq!Estoque_Mer <= 0 And rcrdstPesq!Quantidade_ITo > 0 Then
                   'A mercadoria não existe em estoque
                    
                    'Adiciona na mensagem de mercadoria faltante
                    strSQL = strSQL & Chr(13) & rcrdstPesq!Descricao_ITO & " => "
                    strSQL = strSQL & Format(rcrdstPesq!Quantidade_ITo, "###,##0.00")
                    
                ElseIf rcrdstPesq!Quantidade_ITo > rcrdstPesq!Estoque_Mer Then
                    '  A quantidade do orçamento é maior que a quantidade
                    'em estoque
                    If rcrdstPesq!Tipo_MEr = "M" Or rcrdstPesq!Tipo_MEr = "I" Then
                        'Colocando todo o estoque na nota
                        .Text = F_ValToText(rcrdstPesq!Estoque_Mer)
                        crcQuant = rcrdstPesq!Estoque_Mer
                    End If
                    
                    'Coloca na mensagem de mercadoria faltante
                    strSQL = strSQL & Chr(13) & rcrdstPesq!Descricao_ITO & " => "
                    strSQL = strSQL & Format(rcrdstPesq!Quantidade_ITo - rcrdstPesq!Estoque_Mer, "###,##0.00")
                    
                Else
                    If rcrdstPesq!Tipo_MEr = "M" Or rcrdstPesq!Tipo_MEr = "I" Then
                        'Colocando todo o estoque na nota
                        .Text = F_ValToText(rcrdstPesq!Quantidade_ITo)
                        crcQuant = rcrdstPesq!Quantidade_ITo
                    End If
                End If
                .Col = 5  'Valor unitário
                .Text = F_ValToText(F_Arredonda(dblConstante * rcrdstPesq!unitario_ITo))
                .Col = 6 'Valor total
                .Text = F_ValToText(F_Arredonda(dblConstante * rcrdstPesq!unitario_ITo) * crcQuant)
                .Col = 8 '% IPI
                .Text = "00"
                .Col = 9 'Valor do IPI
                .Text = "0,00"
                .Col = 10 'Marca
                .Text = rcrdstPesq!marca_mer
                .Col = 12 'Desconto
                .Text = "0,00"
                .Col = 13 'Situação da mercadoria
                .Text = rcrdstPesq!Situacao_MER
                .Col = 14 'Estorno do IVA
                .Text = F_ValToText(rcrdstPesq!ValorIva_Mer)
                .Col = 15 'Número da nota de entrada
                .Text = rcrdstPesq!Nota_Ent
            End If
            If crcQuant = 0 Then
                'Excluindo a última linha
                .Rows = .Rows - 1
            End If
            rcrdstPesq.MoveNext
        Loop
        If Not F_Vazio(strSQL) Then
            strSQL = "Mercadorias não disponíveis: " & strSQL
            MsgBox strSQL, vbInformation
        End If
    End With
    rcrdstPesq.Close
    P_AjustaValores
    Exit Sub
TrataErro:
    P_MostraErro
    
End Sub

Private Function F_EmiteCupomFiscal(strNome_Cli As String)
    'Esta função emite o cupom fiscal da nota

    Dim blnResp As Boolean
    Dim intCont As Integer, strST As String, strInf As String
    Dim strCod As String, strDescr As String, strUnidade As String
    Dim crcQuantidade As Currency, crcUnitario As Currency
    Dim strFormaP As String
    Dim crcDesconto As Currency


    strInf = strInf & F_Centra(strNome_Cli, _
    48, "-") & Chr(13) & Chr(10)

    strInf = strInf & F_Centra("Tenha " & _
    IIf(Left(F_Cumprimenta(), 3) = "Bom", _
    "um ", "uma ") & F_Cumprimenta(), 48, "-") & Chr(13) & Chr(10)

    'Carrega o form de formas de pagamento
    'Passando o valor total da nota
    frmFormas_Pagamento.lblRestante = lblTotalNota_Sai.caption
    frmFormas_Pagamento.txtbxValor.Text = lblTotalNota_Sai.caption
    frmFormas_Pagamento.Show 1
    strFormaP = frmFormas_Pagamento.Tag
    Unload frmFormas_Pagamento
    
    If F_Vazio(strFormaP) Then
       MsgBox "Não foi informado corretamente a forma de pagamento!", vbExclamation
       Exit Function
    End If
    
    blnResp = F_AbrirCupomFscl()
    If Not blnResp Then
        Exit Function
    End If

    With msflxgrdItensSaida
        For intCont = 1 To .Rows - 1
            .Row = intCont
            .Col = 0
            strCod = .Text
            .Col = 1
            strDescr = Trim(.Text)
            .Col = 2 'Situação Tributária
            If .Text = "00" Then
                'Mercadoria Tributada
                strST = Format(cmbxPercICMS_Sai.ListIndex + 1, "0#")
            Else
                'Mercadoria com ICMS Subst
                strST = "FF"
            End If
            .Col = 3
            strUnidade = .Text
            .Col = 4
            crcQuantidade = F_TextToVal(.Text)
            .Col = 5
            crcUnitario = F_TextToVal(.Text)
            .Col = 12
            crcDesconto = F_TextToVal(.Text)
            
            blnResp = F_VendaItem(strCod, strDescr, strST, _
            crcQuantidade, crcUnitario, strUnidade, crcDesconto)
        Next
    End With
    If Not F_IniFechaCupomFsl("a", _
    F_TextToVal(txtbxDespesas_Sai.Text)) Then
        Exit Function
    End If
    'frmFormas_Pagamento.Show 1 'Chama as formas de pagamento
    
    'Fazer enquanto houver forma de pagamento
    Do While Len(strFormaP) > 0
       
       'Pegando a expressão passada pelo form até o caracter 13
       strCod = Left(strFormaP, InStr(1, strFormaP, Chr(13)) - 1)
       
       'Executa o comando da forma de pagamento
       F_AbrePortaECF
       F_ComandoECF (conAbre_Protocolo_Fiscal & "72|" & strCod & _
       conFecha_Protocolo_Fiscal)
       F_FechaPortaECF
       
       'Tirando a expressão que foi pega anteriormente
       strFormaP = Right(strFormaP, Len(strFormaP) - (Len(strCod) + 1))
       
    Loop
    
    
    blnResp = F_TermFechamentoCupom(strInf)

    F_EmiteCupomFiscal = True
End Function
Private Sub P_AjustaValores()
    'Este procedimento soma os valores totais da nota
    Dim intCont, intBC As Integer
    Dim crcICMS   As Currency
    Dim crcValTotal As Currency
    Dim strST As String
    
    'Percentual de ICMS em decimal
    crcICMS = CCur(Format(cmbxPercICMS_Sai.Text))
    With msflxgrdItensSaida
      'Zera os valores para calcular os já digitados
      lblBaseCalculo_Sai.caption = "0,00"
      lblBaseCalculoICMSSubst_Sai.caption = "0,00"
      lblIPI_Sai.caption = "0,00"
      lblTotalProdutos_Sai.caption = "0,00"
      lblDesconto_Sai.caption = "0,00"
      
      'Processa todos os itens já lançados
      For intCont = 1 To .Rows - 1
          .Row = intCont
          
          .Col = 6 'Valor Total do Produto
          crcValTotal = F_TextToVal(.Text)
          
          .Col = 11 'Indica o tipo da base de cálculo
          intBC = val(.Text)
          
          .Col = 12 'Desconto
          lblDesconto_Sai.caption = F_ValToText( _
          F_TextToVal(.Text) + F_TextToVal(lblDesconto_Sai.caption))
          
          If optnbtnCupomFiscal.Value Then
              If Not blnMercadTrib Then
                  .Col = 13
                  strST = .Text
                  .Col = 2
                  .Text = IIf(strST = "R" Or strST = "I", "01", "00")
                  .Col = 7
                  .Text = IIf(strST = "T", Left(cmbxPercICMS_Sai.Text, 2), "00")
              Else
                  .Col = 7
                  .Text = Left(cmbxPercICMS_Sai.Text, 2)
                  .Col = 2
                  .Text = "00"
              End If
          End If
          
          .Col = 7 'Indica se há ICMS
          If CCur(.Text) <> 1 Then
             
             .Col = 2 'Mudando para tributação do ICMS
             If intBC = 0 And .Text = "00" Then
                
                'ICMS Normal
                lblBaseCalculo_Sai.caption = _
                F_ValToText(F_TextToVal _
                (lblBaseCalculo_Sai.caption) + crcValTotal)
                
                'Debitando desconto se houver
                .Col = 12
                With lblBaseCalculo_Sai
                     .caption = F_ValToText(F_TextToVal(.caption) - F_TextToVal(msflxgrdItensSaida.Text))
                End With
                
             ElseIf intBC = 2 Then
                'ICMS SUBST
                lblBaseCalculoICMSSubst_Sai = _
                F_ValToText(F_TextToVal _
                (lblBaseCalculoICMSSubst_Sai.caption) + crcValTotal)
             
             End If
          End If
          .Col = 9 'IPI
          lblIPI_Sai.caption = F_ValToText(F_TextToVal _
          (lblIPI_Sai.caption) + F_TextToVal(.Text))
          
          'Soma no total dos produtos
          lblTotalProdutos_Sai.caption = F_ValToText(F_TextToVal _
          (lblTotalProdutos_Sai.caption) + crcValTotal)
          
      Next
      
      If F_TextToVal(lblBaseCalculo_Sai.caption) > 0 Then
         'Calcula despesas sobre a Base de Cálculo
         lblBaseCalculo_Sai.caption = F_ValToText(F_TextToVal _
         (lblBaseCalculo_Sai.caption) + F_TextToVal( _
         txtbxDespesas_Sai.Text))
      End If
      
      If F_TextToVal(lblBaseCalculoICMSSubst_Sai.caption) > 0 Then
         'Calcula desconto sobre a Base de Cálculo Substituição
         lblBaseCalculoICMSSubst_Sai.caption = F_ValToText _
         (F_TextToVal(lblBaseCalculoICMSSubst_Sai.caption) _
         + (txtbxDespesas_Sai.Text))
      End If
      
      If F_TextToVal(lblBaseCalculo_Sai.caption) > _
      F_TextToVal(lblDesconto_Sai.caption) Then
         'Calcula desconto sobre a Base de Cálculo
         lblBaseCalculo_Sai.caption = F_ValToText(F_TextToVal _
         (lblBaseCalculo_Sai.caption) - F_TextToVal( _
         lblDesconto_Sai.caption))
      End If
      
      If F_TextToVal(lblBaseCalculoICMSSubst_Sai.caption) _
      > F_TextToVal(lblDesconto_Sai.caption) Then
         'Calcula desconto sobre a Base de Cálculo Substituição
         lblBaseCalculoICMSSubst_Sai.caption = F_ValToText( _
         F_TextToVal(lblBaseCalculoICMSSubst_Sai.caption) - _
         (lblDesconto_Sai.caption))
      End If
      'Calculando o valor do ICMS
      'Trunca o valor
      lblValorICMS_Sai.caption = F_ValToText(Int(F_TextToVal _
      (lblBaseCalculo_Sai.caption) * crcICMS * 100) / 100)
      
      'Calculando o valor do ICMS Subst
      lblICMSSubst_Sai.caption = F_ValToText(F_TextToVal _
      (lblBaseCalculoICMSSubst_Sai) * crcICMS)
      
      'Calculando o valor total da nota
      lblTotalNota_Sai.caption = F_ValToText(F_TextToVal _
      (lblTotalProdutos_Sai.caption) + _
      F_TextToVal(lblIPI_Sai.caption) + _
      F_TextToVal(txtbxDespesas_Sai.Text) - _
      F_TextToVal(lblDesconto_Sai.caption))
    End With
End Sub
Private Sub P_LimpaItem()
    'Resetando controles para o padrão
    txtbxNome_Mer.Text = ""
    txtbxNome_Mer.Tag = ""
    lblUnidade_Its.caption = ""
    txtbxQuantidade_Its.Text = "0,00"
    txtbxQuantidade_Its.Tag = "0,00"
    txtbxDesconto_Its.Text = "0,00"
    txtbxUnitario_Its.Text = "0,00"
    lblTotal_Its.caption = "0,00"
    txtbxIPI_Its.Text = "0,00"
    txtbxICMS_Its.Text = cmbxPercICMS_Sai.Text
    cmbxBaseCalculo.ListIndex = 0
End Sub
Private Function F_Devolucao() As Boolean
    '  Retorna verdadeiro se a nota que está sendo lançada for uma nota
    'de devolução ou outras saídas
    Dim strOperacao As String
    strOperacao = Mid(cmbxOperacao_Sai.Text, 2, 3)
    If strOperacao = ".32" Or strOperacao = ".99" Then
        F_Devolucao = True
        msflxgrdItensSaida.ColWidth(12) = 0  'Não mostra desconto
        msflxgrdItensSaida.ColWidth(9) = 940 'Mostra IPI
    Else
        F_Devolucao = False
        msflxgrdItensSaida.ColWidth(12) = 940 'Mostra desconto
        msflxgrdItensSaida.ColWidth(9) = 0 'Não mostra IPI
    End If
End Function
Private Function F_VerificaCabecalho() As Boolean
    
    'Esta função verifica a consistência do cabeçalho
    F_VerificaCabecalho = False
    If F_Vazio(txtbxNome_Cli.Tag) Then
       'Tem que se informar o cliente
       MsgBox "É necessário entrar com um cliente!", vbExclamation
       txtbxNome_Cli.SetFocus
    ElseIf F_Vazio(txtbxNota_Sai.Text) Then
       'Tem que se informar a nota fiscal
       MsgBox "É necessário entrar com um número de nota fiscal!", _
       vbExclamation
       txtbxNota_Sai.SetFocus
    ElseIf F_Vazio(txtbxSerie_Sai.Text) Then
       'Tem que se informar a série da nota fiscal
       MsgBox "É necessário entrar com a série da nota fiscal!", _
       vbExclamation
       txtbxSerie_Sai.SetFocus
    ElseIf F_Vazio(cmbxOperacao_Sai.Text) Then
       'Tem que se informar o tipo de operação
       MsgBox "Selecione um Tipo de Operação!", vbExclamation
       cmbxOperacao_Sai.SetFocus
    ElseIf F_Vazio(F_ValidaData(txtbxEmissao_Sai.Text)) Then
       'Tem que se informado a emissão da nota
       MsgBox "Deve-se informar uma data válida!", vbExclamation
       txtbxEmissao_Sai.SetFocus
    'ElseIf Val(txtbxPercIcms_Sai.Text) = 0 And Not F_Devolucao Then
       'Tem que se informar o percentual do ICMS se ñ for devolução
     '  MsgBox "Deve-se informar um percentual para o ICMS!", _
     '  vbExclamation
     '  txtbxPercIcms_Sai.SetFocus
    Else
       'Tudo OK
       F_VerificaCabecalho = True
    End If
End Function
Private Function F_VerificaNota() As Boolean
   'Esta função retorna como resultado se foi encontrada a
   'nota de entrada
   ''On Error  GoTo TrataErro
   Dim rcrdstVerifica As Recordset
   Dim intRes  As Integer
   Dim strSQL As String
   
   F_VerificaNota = True
   
   'Montando instrução SQL
   strSQL = "Select count(*) as Total From TAB_SAIDAMERC01 Where "
   strSQL = strSQL & "NOTA_SAI = '" & txtbxNota_Sai.Text & "' And "
   strSQL = strSQL & "SERIE_SAI = '" & txtbxSerie_Sai.Text & "'"
   'Abrindo Recordset como Snapshot
   Set rcrdstVerifica = banco.OpenRecordset(strSQL, dbOpenSnapshot)
   intRes = rcrdstVerifica!Total
   rcrdstVerifica.Close
   
   If intRes > 0 Then
      'Encontrou uma nota fiscal já lançada
      MsgBox "Nota fiscal já lançada!", vbExclamation
      txtbxNota_Sai.SetFocus
      Exit Function
   End If
   F_VerificaNota = False
   Exit Function

TrataErro:
   P_MostraErro Me.caption
End Function

Private Sub cmbxOperacao_Sai_Click()
    If Not F_Devolucao Then
       cmbxBaseCalculo.ListIndex = 1
       cmbxBaseCalculo.Enabled = False
       cmbxSituacaoTri.ListIndex = 0
       cmbxSituacaoTri.Enabled = False
    Else
       cmbxBaseCalculo.Enabled = True
       cmbxSituacaoTri.Enabled = True
    End If
    
End Sub





Private Sub cmndbtnExcluiItem_Click()
    Dim strNome_Mer As String
    With msflxgrdItensSaida
        If .Row > 0 Then
          .Col = 1
          strNome_Mer = .Text
          .Col = 10
          If .Row = .Rows - 1 Then
             'Não se pode excluir o último item do Grid
'             MsgBox "Este item não pode ser eliminado, você tem " & _
'             "que alterá-lo!", vbExclamation
             .Rows = 1
             cmndbtnIncluiItem.Tag = ""
             Exit Sub
          End If
          If MsgBox("Você deseja excluir o item: " & Chr(13) _
            & Trim(strNome_Mer) & Chr(10) & "Marca: " & Trim(.Text), _
            vbQuestion + vbYesNo) = vbYes Then
            'Remove o item
             .RemoveItem .Row
          End If
       End If
    End With
End Sub

Private Sub cmndbtnFechar_Click()
    'Fecha o formulário
    Unload Me
End Sub

Private Sub cmndbtnIncluiItem_Click()
    If Not F_VerificaCabecalho Then
       Exit Sub
    End If
    
    If F_VerificaNota Then
       'Já existe uma nota fiscal
       Exit Sub
    End If
    
    If F_Vazio(txtbxNome_Mer.Tag) Then
       'Sem uma mercadoria não há como incluir o item
       MsgBox "Informe a mercadoria!", vbExclamation
       txtbxNome_Mer.SetFocus
       Exit Sub
    End If
    
'    If CCur(txtbxQuantidade_Its.Text) > _
'    F_TextToVal(txtbxQuantidade_Its.Tag) Then
'       'Não há a quantidade solicitada em estoque!!!
'       MsgBox "Não há no estoque a quantidade especificada!", _
'       vbExclamation
'       txtbxQuantidade_Its.SetFocus
'       Exit Sub
'    End If
    
    If Not F_Devolucao Then
       'Indica que é nota de venda
       If cmbxBaseCalculo.ListIndex = 2 Then
          ' Não é possível utilizar a base de cálculo de ICMS Subst para
          'venda
          MsgBox "Esta base de Cálculo não é pratica para venda!", _
          vbExclamation
          cmbxBaseCalculo.SetFocus
          Exit Sub
       
       ElseIf cmbxSituacaoTri.ListIndex > 1 Then
          '   Impede a entrada de outra situação tributária que não seja
          '00 ou 01
          MsgBox "A Situação tributária para venda é 00 ou 01"
          cmbxOperacao_Sai.SetFocus
          Exit Sub
       End If
    End If
    
    If CCur(F_AjustaValor(txtbxIPI_Its.Text)) > _
    CCur(F_AjustaValor(lblTotal_Its)) Then
       ' O valor do IPI está além do esperado
       MsgBox "O valor do IPI está muito alto!", vbExclamation
       txtbxIPI_Its.SetFocus
       Exit Sub
    
    ElseIf CCur(txtbxICMS_Its.Text) <> CCur(Left(cmbxPercICMS_Sai.Text, 2)) And _
    val(txtbxICMS_Its.Text) <> 0 Then
       'O ICMS está com valor fora da faixa
       MsgBox "O percentual do ICMS está fora do padrão!", _
       vbExclamation
       If txtbxICMS_Its.Enabled Then
          txtbxICMS_Its.SetFocus
       End If
       Exit Sub
    End If
    
    'Verificando Desconto
    If F_TextToVal(txtbxDesconto_Its.Text) >= _
    F_TextToVal(lblTotal_Its) Then
        'Valor do desconto maior ou igual o valor de venda
        
        MsgBox "Desconto muito alto!!", vbExclamation
        txtbxDesconto_Its.SetFocus 'Focalizando para entrada de desconto
        Exit Sub
    End If
        
    
    
'    If MsgBox("Deseja " & IIf(cmndbtnIncluiItem.Tag <> "", _
'    "alterar ", "incluir ") & "a mercadoria na nota!", _
'    vbYesNo + vbQuestion) = vbNo Then
'       'Ajusta os valores
'       P_AjustaValores
'       'Limpa os controles de inclusão de item
'       P_LimpaItem
'       Exit Sub
'    End If
'
        
    With msflxgrdItensSaida
         
         If val(cmndbtnIncluiItem.Tag) <= 0 Then
            'Incluindo
            .AddItem Format(txtbxNome_Mer.Tag, "0#\.##\.###\-#")
            .Row = .Rows - 1
         Else
            'Editando
            .Row = val(cmndbtnIncluiItem.Tag)
            .Col = 0
            .Text = Format(txtbxNome_Mer.Tag, "0#\.##\.###\-#")
            cmndbtnIncluiItem.Tag = ""
         End If
         'Nome Mercadoria
         .Col = 1: .Text = txtbxNome_Mer.Text
         'Situação Tributária
         .Col = 2: .Text = Left(cmbxSituacaoTri, 2)
         'Unidade da Mercadoria
         .Col = 3: .Text = lblUnidade_Its
         'Quantidade
         .Col = 4: .Text = txtbxQuantidade_Its.Text
         'Valor Unitário
         .Col = 5: .Text = txtbxUnitario_Its.Text
         'Valor Total
         .Col = 6: .Text = lblTotal_Its.caption
         'Valor do ICMS
         .Col = 7: .Text = txtbxICMS_Its.Text
         'Percentual do IPI
         .Col = 8: .Text = Format((CCur(txtbxIPI_Its.Text) * 100) _
         / CCur(lblTotal_Its.caption), "00")
         'IPI
         .Col = 9: .Text = txtbxIPI_Its.Text
         'Código da marca
         .Col = 10: .Text = lblMarca_Mer.Tag
         'Base de Cálculo
         .Col = 11: .Text = Trim(Str(cmbxBaseCalculo.ListIndex))
         'Desconto
         .Col = 12: .Text = txtbxDesconto_Its
         'Situação da mercadoria
         .Col = 13: .Text = cmbxSituacaoTri.Tag
         'Valor do estorno do IVA
         .Col = 14: .Text = txtbxDesconto_Its.Tag
         'Número da nota fiscal de Entrada
         .Col = 15: .Text = lblDesconto_Its.Tag
     End With
     P_LimpaItem       'Limpa os controles de inclusão dos itens
     P_AjustaValores   'Ajusta os valores
End Sub

Private Sub cmndbtnRegistraNota_Click()
    Dim strSQL         As String 'Instrução SQL
    Dim intCont        As Integer
    Dim strIsenta      As String
    Dim strOutras      As String
    Dim strCod_Mer     As String
    Dim strMarca_Mer   As String
    Dim strValor       As String
    Dim strQuant       As String
    Dim crcEstornoICMS As Currency  'Acumula o Estorno do ICMS
    Dim strNota_Ent    As String
    Dim rcrdstPesquisa As Recordset
    Dim strCusto       As String
    On Error GoTo TrataErro
    If Not F_VerificaCabecalho Then
       Exit Sub
    End If
    
    If F_VerificaNota Then
       Exit Sub
    End If
    P_AjustaValores
    If F_TextToVal(lblTotalNota_Sai) <= 0 Then
       MsgBox "Não há valor para fechamento!"
       Exit Sub
    End If
    
    If MsgBox("Confirma Fechamento da Venda?", vbQuestion + vbYesNo) = vbNo Then
       'Não confirmou fechamento da venda
       Exit Sub
    End If
    
    If optnbtnCupomFiscal.Value Then
    'Emitindo Cupom Fiscal
       If Not F_EmiteCupomFiscal(txtbxNome_Cli.Text) Then
          Exit Sub
       End If
    End If
    If Not optnbtnCupomFiscal.Value And Not optnbtnNotaFiscal.Value Then
        MsgBox "Deve-se escolher uma forma de impressão de nota fiscal!", _
        vbExclamation
        optnbtnCupomFiscal.SetFocus
        Exit Sub
     End If
    'Calculando o valor de produtos isentos
    strIsenta = F_ValToText(F_TextToVal(lblTotalNota_Sai.caption) - _
      F_TextToVal(lblIPI_Sai.caption) - _
      F_TextToVal(lblBaseCalculo_Sai.caption) - _
      F_TextToVal(lblBaseCalculoICMSSubst_Sai.caption))
    
    strOutras = F_ValToText(F_TextToVal(strIsenta) + F_TextToVal(lblIPI_Sai))
    'Passando instruções SQL
    'Gravando na Tabela de Cabeçalho
    strSQL = "Insert Into TAB_SAIDAMERC01 ("
    strSQL = strSQL & "COD_CLI, NOTA_SAI, SERIE_SAI, EMISSAO_SAI, "
    strSQL = strSQL & "VALORCONTABIL_SAI, BASECALCULO_SAI, ISENTA_SAI, "
    strSQL = strSQL & "OUTRAS_SAI, DESCONTO_SAI, ICMS_SAI, "
    strSQL = strSQL & "OPERACAO_SAI, BASECALCULOSUBST_SAI, "
    strSQL = strSQL & "ICMSSUBST_SAI, ITENS_SAI, UF_SAI, PERCICMS_SAI,"
    strSQL = strSQL & "NICK_USU) Values ( "
    strSQL = strSQL & "'" & txtbxNome_Cli.Tag & "', "
    strSQL = strSQL & "'" & txtbxNota_Sai.Text & "', "
    strSQL = strSQL & "'" & txtbxSerie_Sai.Text & "', "
    strSQL = strSQL & F_GravaData(txtbxEmissao_Sai.Text) & ", "
    strSQL = strSQL & F_AjustaValor(lblTotalNota_Sai.caption) & ", "
    strSQL = strSQL & F_AjustaValor(lblBaseCalculo_Sai.caption) & ", "
    strSQL = strSQL & "0, "
    strSQL = strSQL & F_AjustaValor(strOutras) & ", "
    strSQL = strSQL & F_AjustaValor(lblDesconto_Sai.caption) & ", "
    strSQL = strSQL & F_AjustaValor(lblValorICMS_Sai.caption) & ", "
    strSQL = strSQL & "'" & Left(cmbxOperacao_Sai.Text, 4) & "', "
    strSQL = strSQL & F_AjustaValor(lblBaseCalculoICMSSubst_Sai.caption) & ", "
    strSQL = strSQL & F_AjustaValor(lblICMSSubst_Sai.caption) & ", "
    strSQL = strSQL & Trim(Str(msflxgrdItensSaida.Rows - 1)) & ", "
    strSQL = strSQL & "'" & lblUF_Cli.caption & "', "
    strSQL = strSQL & F_AjustaValor(Left(cmbxPercICMS_Sai.Text, 2)) & ", "
    strSQL = strSQL & "'" & strNome_Usu & "') "
    
    banco.Execute strSQL
    'Gravando os itens da nota fiscal
    crcEstornoICMS = 0
    For intCont = 1 To msflxgrdItensSaida.Rows - 1
        With msflxgrdItensSaida
            .Row = intCont
            
            'Instrução SQL
            strSQL = "Insert Into TAB_ITENSSAIDA01 ("
            strSQL = strSQL & "NOTA_SAI, SERIE_SAI, COD_MER, "
            strSQL = strSQL & "MARCA_MER, "
            strSQL = strSQL & "QUANTIDADE_ITS, VALORTOTAL_ITS, "
            strSQL = strSQL & "SITUACAOTRIB_ITS, VALORIPI_ITS, "
            strSQL = strSQL & "VALORICMS_ITS, IVA_ITS, NOTA_ITS) Values ("
            strSQL = strSQL & "'" & txtbxNota_Sai.Text & "', "
            strSQL = strSQL & "'" & txtbxSerie_Sai.Text & "', "
            .Col = 0 'Código da mercadoria
            strCod_Mer = F_SoNumero(.Text)
            strSQL = strSQL & "'" & strCod_Mer & "', "
            .Col = 10 'Código da marca da mercadoria
            strSQL = strSQL & "'" & .Text & "', "
            strMarca_Mer = .Text
            .Col = 4 'Quantidade
            strSQL = strSQL & F_AjustaValor(.Text) & ", "
            strQuant = .Text
            
            .Col = 6 'Valor Total
            strSQL = strSQL & F_AjustaValor(.Text) & ", "
            strValor = .Text 'Guarda o valor total em variável
            
            .Col = 2 'Situação tributária
            strSQL = strSQL & "'" & .Text & "', "
            
            .Col = 9 'Valor do IPI
            strSQL = strSQL & F_AjustaValor(.Text) & ", "
            .Col = 7 'Percentual do ICMS
            strSQL = strSQL & F_AjustaValor(F_TextToVal(strValor) _
            * (F_TextToVal(.Text) / 100)) & ", "
            .Col = 13
            'Verificando a Situação Tributária da Mercadoria
            If .Text = "R" And lblUF_Cli <> "GO" Then
               .Col = 15
               If val(.Text) = 0 Then
                  'Não encontrou a nota fiscal
                  'Então calculará o estorno sobre o valor de venda
                  
                  Set rcrdstPesquisa = banco.OpenRecordset("Select " & _
                  "IVAINT_MER, ALIQUOTA_MER From TAB_CMPLT01 Where " & _
                  "COD_MER = '" & strCod_Mer & "' And " & _
                  "MARCA_MER = '" & strMarca_Mer & "'", dbOpenSnapshot)
                  
                  .Col = 5
                  strCusto = .Text
                  
                  'Calcula-se a metade do IVA sobre a venda
                  '(30 / 100) / 2 =  30 / 200
                  
                  strSQL = strSQL & Str( _
                  F_CalcICMSSubst(F_TextToVal(strCusto), 0, _
                  rcrdstPesquisa!IVAINT_MER / 200, _
                  rcrdstPesquisa!ALIQUOTA_MER / 100) - _
                  CCur(strCusto)) & ", "
                  
                  strSQL = strSQL & "'00000000' )"
                  
                  rcrdstPesquisa.Close
                  
               Else
                  'Se for Retida no Estado então, estorno de ICMS
                  'Aplicar a fórmula, E passar o número da nota
                  .Col = 14
                  strSQL = strSQL & F_AjustaValor(.Text) & ", "
                  crcEstornoICMS = crcEstornoICMS + F_TextToVal(.Text)
                  .Col = 15
                  strSQL = strSQL & "'" & .Text & "') "
               End If
            Else
               'Caso contrário passar valores nulos
               strSQL = strSQL & "0, "
               strSQL = strSQL & "'00000000' )"
            End If
            
            
            'Executando SQL
            banco.Execute strSQL
            
            'Ajustando Tabela Complemento do Estoque
            strSQL = "Update TAB_CMPLT01 Set ESTOQUE_MER = "
            
            strSQL = strSQL & "ESTOQUE_MER - " & F_AjustaValor(strQuant)
            
            strSQL = strSQL & "Where COD_MER = '" & strCod_Mer
            strSQL = strSQL & "' And MARCA_MER = '" & _
            strMarca_Mer & "'"
            
            banco.Execute strSQL
        End With
    Next
    If F_TextToVal(lblValorICMS_Sai) > 0 Then
        'Houve débito de ICMS
        'Gravando na Tabela de ICMS
        strSQL = "Insert into TAB_ICMS (COD_CLI, NOTA_SAI, "
        strSQL = strSQL & "SERIE_SAI , DATA_ICM, TIPO_ICM, "
        strSQL = strSQL & "VALOR_ICM, NICK_USU, VENCIMENTO_ICM)"
         
        strSQL = strSQL & " Values ("
        strSQL = strSQL & "'" & txtbxNome_Cli.Tag & "', "
        strSQL = strSQL & "'" & txtbxNota_Sai.Text & "', "
        strSQL = strSQL & "'" & txtbxSerie_Sai.Text & "', "
        strSQL = strSQL & F_GravaData(txtbxEmissao_Sai.Text) & ", "
        strSQL = strSQL & "'-', "
        strSQL = strSQL & F_AjustaValor(lblValorICMS_Sai.caption) & ", "
        strSQL = strSQL & "'" & strNome_Usu & "', "
        strSQL = strSQL & F_GravaData(Format(DateAdd("M", 1, _
        CDate(txtbxEmissao_Sai.Text)), "\0\9\/MM/YYYY")) & ")"
        
        'Gravando...
        banco.Execute strSQL
    End If
    
    If crcEstornoICMS > 0 Then
        'Houve estorno de iva,  grava crédito na Tabela de ICMS
        strSQL = "Insert into TAB_ICMS (COD_CLI, NOTA_SAI, "
        strSQL = strSQL & "SERIE_SAI , DATA_ICM, TIPO_ICM, "
        strSQL = strSQL & "VALOR_ICM, NICK_USU, VENCIMENTO_ICM)"
        strSQL = strSQL & " Values ("
        strSQL = strSQL & "'" & txtbxNome_Cli.Tag & "', "
        strSQL = strSQL & "'" & txtbxNota_Sai.Text & "', "
        strSQL = strSQL & "'" & txtbxSerie_Sai.Text & "', "
        strSQL = strSQL & F_GravaData(txtbxEmissao_Sai.Text) & ", "
        strSQL = strSQL & "'+', "
        strSQL = strSQL & F_AjustaValor(F_ValToText(crcEstornoICMS)) & ", "
        strSQL = strSQL & "'" & strNome_Usu & "', "
        strSQL = strSQL & F_GravaData(Format(DateAdd("M", 1, _
        CDate(txtbxEmissao_Sai.Text)), "\0\9\/MM/YYYY")) & ")"
        
        'Gravando ...
        banco.Execute strSQL
    End If
        
    If optnbtnCupomFiscal.Value = True Then
        If MsgBox("Deseja imprimir nota fiscal referente ao cupom?", vbYesNo + vbExclamation) = vbYes Then
            'P_CarregaNF
        End If
    End If
    If MsgBox("Nota Fiscal Registrada!!!" & Chr(13) & "Deseja Imprimir Planilha???", vbYesNo + vbQuestion, "Saída de Mercadorias") = vbYes Then
        'P_RelSaidaMercadoria txtbxNota_Sai.Text, txtbxSerie_Sai.Text
    End If
    'Unload Me
    P_LimpaItem
    txtbxNome_Cli.Text = ""
    txtbxNota_Sai.Text = ""
    txtbxNome_Cli.SetFocus
    Exit Sub
    
TrataErro:
    P_MostraErro Me.caption
    
End Sub
'Sub P_CarregaNF()
'    Dim iCount As Integer
'    Dim iCol As Integer
'    Dim strSQL As String
'    With frmSaidaNF
'        .Tag = txtbxNota_Sai
'        For iCount = 1 To msflxgrdItensSaida.Rows - 1
'            msflxgrdItensSaida.Row = iCount
'            strSQL = ""
'            For iCol = 0 To msflxgrdItensSaida.Cols - 1
'                msflxgrdItensSaida.Col = iCol
'                strSQL = strSQL & msflxgrdItensSaida.Text
'            Next
'            .msflxgrdItensSaida.AddItem strSQL
'        Next
'
'        .Show 1
'        If .Tag <> "" Then
'        End If
'    End With
'End Sub
'Private Sub cmndbtnVinculaOrcamento_Click()
'    '   Este procedimento irá vincular o orçamento juntamente com a nota
'    'fiscal de origem
'    If Me.Tag = "" Then
'       Me.Tag = F_EntraDados("Informe o número do orçamento", , "0#######")
'       If F_Vazio(Me.Tag) Then
'           'Cancelou a escolha do orçamento
'           Exit Sub
'       End If
'    End If
'    P_CarregaControles
'End Sub

Private Sub Form_Activate()
   
      
   If Me.Tag = "" And F_Vazio(txtbxNota_Sai.Text) Then
      txtbxEmissao_Sai.Text = Date
      'optnbtnCupomFiscal.Value = True
      'optnbtnCupomFiscal_Click
      optnbtnNotaFiscal.Value = True
      blnFlag = True
   End If
   
   'Carregando ComboBox
  
   
   'Configurando grid
   With msflxgrdItensSaida
       .Cols = 16
       .ColWidth(0) = 950  'Código da mercadoria
       .ColWidth(1) = 3000 'Descrição do produto
       .ColWidth(2) = 320  'Situação tributária
       .ColWidth(3) = 350  'Unidade
       .ColWidth(4) = 940  'Quantidade do item
       .ColWidth(5) = 940  'Valor unitário do item
       .ColWidth(6) = 940  'Valor total do item
       .ColWidth(7) = 350  'Percentual do ICMS
       .ColWidth(8) = 0    'Percetual IPI
       .ColWidth(9) = 0  'Valor do IPI
       .ColWidth(10) = 0   'Código da marca
       .ColWidth(11) = 0   'Tipo da Base de Cálculo
       .ColWidth(12) = 940   'Desconto sobre a mercadoria
       .ColWidth(13) = 0   'Situação da Mercadoria
       .ColWidth(14) = 0   'Estorno do IVA
       .ColWidth(15) = 0   'Número da nota fiscal de entrada
       .Row = 0
       
       .Col = 0: .Text = "Código:"
       .Col = 1: .Text = "Descrição do Produto:"
       .Col = 2: .Text = "ST"
       .Col = 3: .Text = "Un:"
       .Col = 4: .Text = "Quantidade:"
       .Col = 5: .Text = "Unitário:"
       .Col = 6: .Text = "Total:"
       .Col = 7: .Text = "Icms"
       .Col = 8: .Text = "Ipi"
       .Col = 9: .Text = "Valor IPI"
       .Col = 12: .Text = "Desconto:"
   End With
   If Not F_Vazio(Me.Tag) Then
        P_CarregaControles
   End If
   P_Status ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strMens As String
    If KeyCode = vbKeyF1 Then
        strMens = "Alt + I => Inclui o item a nota fiscal." & Chr(13)
        strMens = strMens & "Alt + E => Exclui o item da nota fiscal." & Chr(13)
        strMens = strMens & "Alt + G => Registra nota fiscal." & Chr(13)
        strMens = strMens & "Alt + V => Vincula orçamento à nota fiscal."
        'P_MostraTeclaAtalho strMens
    ElseIf KeyCode = vbKeyG And Shift = vbAltMask Then
'        If F_ControleAtivo(cmndbtnRegistraNota) Then
'            cmndbtnRegistraNota_Click
'        End If
'    ElseIf KeyCode = vbKeyV And Shift = vbAltMask Then
'        If F_ControleAtivo(cmndbtnVinculaOrcamento) Then
'            cmndbtnVinculaOrcamento_Click
'        End If
'    ElseIf KeyCode = vbKeyE And Shift = vbAltMask Then
'        If F_ControleAtivo(cmndbtnExcluiItem) Then
'            cmndbtnExcluiItem_Click
'        End If
'    ElseIf KeyCode = vbKeyI And Shift = vbAltMask Then
'        If F_ControleAtivo(cmndbtnIncluiItem) Then
'            cmndbtnIncluiItem_Click
'        End If
    End If
End Sub

Private Sub Form_Load()
   P_Status "Iniciando Lançamento"
   P_PreencheNotaSaida cmbxOperacao_Sai
   P_PreencheSituacaoTributaria cmbxSituacaoTri
   
   cmbxOperacao_Sai.ListIndex = 0
   cmbxSituacaoTri.ListIndex = 0
End Sub

Private Sub msflxgrdItensSaida_DblClick()
   'Permite editar a mercadoria
   With msflxgrdItensSaida
     If .Row > 0 Then
        cmndbtnIncluiItem.Tag = Str(.Row)
        .Col = 1
        txtbxNome_Mer.Text = Trim(.Text)
        'txtbxNome_Mer_DblClick
        Exit Sub
     End If
    End With
End Sub


Private Sub optnbtnCupomFiscal_Click()
    txtbxNota_Sai.Text = F_NumeroCupom()
    If txtbxNota_Sai.Text = "" Then
        MsgBox "Não é possível registrar cupom! " & Chr(13) & "A Impressora ECF Não foi Encontrada", vbCritical
        Unload Me
        Exit Sub
    End If
    P_PreencheECFTaxasICMS cmbxPercICMS_Sai
    txtbxNota_Sai.Enabled = False
    txtbxSerie_Sai.Text = "CUPM"
    txtbxSerie_Sai.Enabled = False
    cmbxPercICMS_Sai.ListIndex = IIf(blnMercadTrib, 0, 1)
    'cmbxPercICMS_Sai.Enabled = False
    cmbxOperacao_Sai.ListIndex = IIf(blnMercadTrib, 1, 0)
    txtbxICMS_Its.Enabled = False
End Sub

Private Sub optnbtnNotaFiscal_Click()
    txtbxNota_Sai.Text = ""
    txtbxSerie_Sai.Text = ""
    P_PreencheTaxasICMS cmbxPercICMS_Sai
    cmbxPercICMS_Sai.AddItem "00%"
    txtbxNota_Sai.Enabled = True
    txtbxSerie_Sai.Enabled = True
    cmbxPercICMS_Sai.ListIndex = IIf(blnMercadTrib, 2, 3)
    cmbxOperacao_Sai.ListIndex = IIf(blnMercadTrib, 1, 0)
    txtbxICMS_Its.Enabled = F_Devolucao
End Sub



Private Sub txtbxDesconto_Its_GotFocus()
    'Selecionando texto digitado
    P_ValorGotFocus txtbxDesconto_Its
End Sub
Private Sub txtbxDesconto_Its_LostFocus()
    'Selecionando texto digitado
    P_ValorLostFocus txtbxDesconto_Its
End Sub

Private Sub txtbxDespesas_Sai_GotFocus()
    'Selecionando texto digitado
    P_ValorGotFocus txtbxDespesas_Sai
End Sub

Private Sub txtbxDespesas_Sai_LostFocus()
    'Ajustando valor
    P_ValorLostFocus txtbxDespesas_Sai
End Sub

Private Sub txtbxEmissao_Sai_GotFocus()
   'Selecionando texto digitado anteriormente
   P_SelecTexto txtbxEmissao_Sai
End Sub

Private Sub txtbxEmissao_Sai_LostFocus()
   'Ajustando data digitada
   P_AjustaData txtbxEmissao_Sai
End Sub

Private Sub txtbxNome_Cli_DblClick()
    
'    Dim psCliente As Pessoa
'    Dim strSQL As String
'    Dim strNome_Cli As String
'
'    'Ajustando texto a fim de evitar erros
'    P_AjustaTexto txtbxNome_Cli
'    If F_Vazio(txtbxNome_Cli.Text) Then
'       MsgBox "É necessário entrar com um nome primeiro!", _
'       vbExclamation
'       txtbxNome_Cli.SetFocus
'       Exit Sub
'    End If
'    strNome_Cli = txtbxNome_Cli.Text
'
'    psCliente = F_LocalizaTerceiro(txtbxNome_Cli.Text)
'    txtbxNome_Cli.Text = psCliente.Nome  'Nome do cliente
'    txtbxNome_Cli.Tag = psCliente.codigo 'Código do cliente
'    lblUF_Cli.caption = psCliente.UF     'UF do cliente
'
'
'
'
'
'
'    If F_Vazio(txtbxNome_Cli.Tag) Then
'       txtbxNome_Cli.Tag = ""
'       strSQL = "Nenhum cliente selecionado!"
'       strSQL = strSQL & "Você deseja selecionar outro cliente?"
'       If MsgBox(strSQL, vbQuestion + vbYesNo) = vbYes Then
'          txtbxNome_Cli.Text = strNome_Cli
'          txtbxNome_Cli_DblClick
'       Else
'          MsgBox "Sem selecionar cliente não é possível emitir Nota Fiscal!", vbExclamation
'          Unload Me
'       End If
'
'    End If
'
'    'Pesquisando cliente para verificar o tipo de pessoa:
'    'física ou jurídica
'
'    strSQL = "Select TIPOPESSOA_TER From TAB_TERCEIROS Where COD_TER = '"
'    strSQL = strSQL & psCliente.codigo & "'"
'
'    Set rcrdstClientes = banco.OpenRecordset(strSQL, dbOpenSnapshot)
'
'    If rcrdstClientes.RecordCount < 1 Then
'       Unload Me
'       Exit Sub
'    End If
'
'
'    'Indica a situação da tributação da mercadoria
'    blnMercadTrib = (rcrdstClientes!TipoPessoa_Ter = "J" _
'    And lblUF_Cli.caption <> "GO")
'
'
'
'    If optnbtnCupomFiscal.Value Then
'       optnbtnCupomFiscal_Click  'Ajusta o ICMS
'    Else
'       optnbtnNotaFiscal_Click   'Ajusta o ICMS
'    End If
'    P_AjustaValores
'    'SendKeys Chr(9) 'Vai para o próximo controle
'
End Sub

Private Sub txtbxNome_Cli_GotFocus()
   P_SelecTexto txtbxNome_Cli
End Sub

Private Sub txtbxNome_Cli_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then txtbxNome_Cli_DblClick
End Sub

Private Sub txtbxNome_Cli_LostFocus()
   P_AjustaTexto txtbxNome_Cli
End Sub
'Private Sub txtbxNome_Mer_DblClick()
'   ''On Error  GoTo TrataErro
'   Dim mrcdrNome As Mercadoria
'   Dim rcrdstMercadoria As Recordset
'   Dim strSQL As String
'   Dim intCont As Integer
'   Dim strNome_Mer As String
'   If OptnMercadoria(2).Value = True Then
'        P_ValorLostFocus txtbxNome_Mer
'   ElseIf OptnMercadoria(1).Value = True Then
'        txtbxNome_Mer = F_SoNumero(txtbxNome_Mer.Text)
'   Else
'        P_AjustaTexto txtbxNome_Mer
'   End If
'   strNome_Mer = txtbxNome_Mer.Text
'   If OptnMercadoria(2).Value Then
'        mrcdrNome = F_Mercadoria_Localiza("PRECO_MER", txtbxNome_Mer.Text)
'   ElseIf OptnMercadoria(1).Value Then
'        mrcdrNome = F_Mercadoria_Localiza("Cod_Mer", txtbxNome_Mer.Text)
'   Else
'        mrcdrNome = F_Mercadoria_Localiza("Nome_Mer", txtbxNome_Mer.Text)
'   End If
'   P_LimpaItem
'   txtbxNome_Mer.Tag = mrcdrNome.codigo     'Código da mercadoria
'   txtbxNome_Mer.Text = mrcdrNome.Nome      'Nome da mercadoria
'   lblMarca_Mer.caption = mrcdrNome.NMarca  'Nome da marca
'   lblMarca_Mer.Tag = mrcdrNome.CMarca      'Código da marca
'
'   If F_Vazio(txtbxNome_Mer.Tag) Then
'      'Se retornou vazio então chame novamente o procedimento
'      strSQL = "Nenhuma mercadoria foi selecionada!"
'      strSQL = strSQL & Chr(13) & "Deseja selecionar mercadoria?"
'      If MsgBox(strSQL, vbQuestion + vbYesNo) = vbYes Then
'         txtbxNome_Mer.Text = strNome_Mer
'         txtbxNome_Mer_DblClick
'      Else
'         Exit Sub
'      End If
'   Else
'     'Verificando o tipo da mercadoria
'     If mrcdrNome.Tipo <> "M" And mrcdrNome.Tipo <> "I" Then
'        MsgBox "Você deve selecionar uma mercadoria válida!"
'        txtbxNome_Mer = strNome_Mer
'        txtbxNome_Mer_DblClick
'     End If
'
'     cmbxSituacaoTri.Tag = mrcdrNome.Situacao
'     'Verificando a Situação Tributária da Mercadoria
'
'     If mrcdrNome.Situacao <> "T" And lblUF_Cli = "GO" Then
'        'Não é Tributada e é ESTADO de Goiás
'        cmbxSituacaoTri.ListIndex = 1
'        cmbxBaseCalculo.ListIndex = 1
'        txtbxICMS_Its.Text = "00"
'     Else
'        'É tributada ou não é estado de Goiás
'        cmbxSituacaoTri.ListIndex = 0
'        cmbxBaseCalculo.ListIndex = 0
'        txtbxICMS_Its.Text = Left(cmbxPercICMS_Sai.Text, 2)
'     End If
'
'     'Usando SQL para pesquisar valor unitário,  a unidade da mercadoria
'     'e o valor do Estorno do IVA, NOTA DO ESTORNO
'     strSQL = "Select TM.UNIDADE_MER, TR.PRECO_MER, TC.ESTOQUE_MER, "
'     strSQL = strSQL & "TC.VALORIVA_MER, TC.NOTA_ENT From "
'     strSQL = strSQL & "TAB_MERCADORIAS AS TM, TAB_REAL01 As TR, "
'     strSQL = strSQL & "TAB_CMPLT01 As TC Where "
'     strSQL = strSQL & "TM.COD_MER = '" & txtbxNome_Mer.Tag & "' "
'     strSQL = strSQL & "And TM.MARCA_MER = '" & lblMarca_Mer.Tag & "' "
'     strSQL = strSQL & "And TM.COD_MER = TR.COD_MER And "
'     strSQL = strSQL & "TM.COD_MER = TC.COD_MER "
'     strSQL = strSQL & "And TM.MARCA_MER = TC.MARCA_MER "
'     strSQL = strSQL & "And TM.MARCA_MER = TR.MARCA_MER"
'
'     'Faz um recordset de pesquisa do item
'     Set rcrdstMercadoria = banco.OpenRecordset(strSQL, dbOpenSnapshot)
'
'     txtbxUnitario_Its.Text = rcrdstMercadoria!Preco_Mer
'
'     'Pegando o valor unitário do item
'     lblUnidade_Its.caption = rcrdstMercadoria!Unidade_Mer
'
'     'Pegando a quantidade máxima disponível
'     txtbxQuantidade_Its.Tag = rcrdstMercadoria!Estoque_Mer
'
'     'Pegando o último preço do estorno do IVA
'     txtbxDesconto_Its.Tag = rcrdstMercadoria!ValorIva_Mer
'     lblDesconto_Its.Tag = rcrdstMercadoria!Nota_Ent
'
'     rcrdstMercadoria.Close
'
'     If CCur(txtbxUnitario_Its) = 0 Then
'        'Informa que a mercadoria está com preço de venda zerado
'        MsgBox "A mercadoria está com preço de venda zerado!", _
'        vbCritical
'        txtbxNome_Mer.Text = strNome_Mer
'        txtbxNome_Mer_DblClick
'        Exit Sub
'
''     ElseIf CCur(txtbxQuantidade_Its.Tag) = 0 Then
''        'Não há a mercadoria em estoque
''         MsgBox "Não há estoque desta mercadoria!", vbExclamation
''         txtbxNome_Mer.Text = Trim(strNome_Mer)
''         P_SelecTexto txtbxNome_Mer
''         txtbxNome_Mer.SetFocus
''         Exit Sub
'
'     End If
'
'     If msflxgrdItensSaida.Rows > 1 Then
'        For intCont = 1 To msflxgrdItensSaida.Rows - 1
'            With msflxgrdItensSaida
'                 .Row = intCont
'                 .Col = 10
'                 'Igualando com a marca que está no grid
'                 mrcdrNome.CMarca = .Text
'                 .Col = 0
'
'                 If F_SoNumero(.Text) = txtbxNome_Mer.Tag And _
'                 mrcdrNome.CMarca = lblMarca_Mer.Tag Then
'                   'Marca e código igual: Mercadoria já está registrada
'
'                   If MsgBox("Este item já está lançado, alterar?", _
'                   vbQuestion + vbYesNo, conSigla_Sistema) = vbYes Then
'                    'Igualando valores com os já registrados
'
'                     If Not optnbtnCupomFiscal.Value Then
'                        .Col = 2 'Situação da mercadoria
'                        cmbxSituacaoTri.ListIndex = _
'                        CCur(IIf(.Text = "09", "08", .Text))
'                     End If
'
'                     .Col = 3 'Unidade
'                     lblUnidade_Its.caption = .Text
'                     .Col = 4 'Quantidade
'                     txtbxQuantidade_Its.Text = .Text
'                     .Col = 5 'Unitário
'                     txtbxUnitario_Its.Text = .Text
'                     .Col = 6 'Total
'                     lblTotal_Its.caption = .Text
'                     '.Col = 7 'ICMS
'                     'txtbxICMS_Its.Text = .Text
'
'                     cmndbtnIncluiItem.Tag = Trim(Str(intCont))
'
'                     .Col = 9 'IPI
'                     txtbxIPI_Its.Text = .Text
'                     .Col = 11 'Base de Cálculo
'                     cmbxBaseCalculo.ListIndex = CCur(.Text)
'                     .Col = 12 'Desconto
'                     txtbxDesconto_Its.Text = .Text
'                     Exit For
'                 Else
'
'                     Exit Sub
'                 End If
'              End If
'            End With
'        Next
'
'     End If
'  End If
'  txtbxIPI_Its.Enabled = F_Devolucao()
'  txtbxQuantidade_Its.SetFocus
'  Exit Sub
'
'TrataErro:
'  P_MostraErro Me.caption
'
'End Sub

Private Sub txtbxNome_Mer_GotFocus()
   'Seleciona o texto
   P_SelecTexto txtbxNome_Mer
End Sub

Private Sub txtbxNome_Mer_KeyPress(KeyAscii As Integer)
   'Se for pressionada enter execute o procedimento de duplo clique
   If KeyAscii = 13 Then
      'txtbxNome_Mer_DblClick
   End If
End Sub

Private Sub txtbxNome_Mer_LostFocus()
   'Ajusta o texto
   P_AjustaTexto txtbxNome_Mer
End Sub

Private Sub txtbxNota_Sai_GotFocus()
    'Selecionando texto
    P_SelecTexto txtbxNota_Sai
End Sub

Private Sub txtbxNota_Sai_LostFocus()
    'Ajustanto número da nota
    P_AjustaNota txtbxNota_Sai
End Sub

Private Sub txtbxQuantidade_Its_GotFocus()
   'Seleciona o valor digitado
   P_ValorGotFocus txtbxQuantidade_Its
End Sub

Private Sub txtbxQuantidade_Its_LostFocus()
   'Ajusta o valor digitado e multiplica para achar valor total
   
   P_ValorLostFocus txtbxQuantidade_Its
   If CCur(txtbxQuantidade_Its.Text) <= 0 And _
   Not F_Vazio(txtbxNome_Mer.Tag) And F_TextToVal(txtbxQuantidade_Its.Tag) > 0 Then
      'Tem que informar uma quantidade maior que zero
      MsgBox "Deve-se entrar com valor para quantidade", vbExclamation
      'txtbxQuantidade_Its.SetFocus
   
 '  ElseIf CCur(txtbxQuantidade_Its.Text) > _
'   F_TextToVal(txtbxQuantidade_Its.Tag) Then
'      'A quantidade digitada é maior que a quantidade em estoque
'      MsgBox "Não há a quantidade especificada em Estoque", _
'      vbExclamation
'      txtbxQuantidade_Its.Text = F_ValToText(Val _
'      (txtbxQuantidade_Its.Tag))
'      'txtbxQuantidade_Its.SetFocus
   End If
   'Atualiza o valor total do item truncando o valor
   lblTotal_Its.caption = Format(Int(CCur(txtbxQuantidade_Its.Text) * _
   CCur(txtbxUnitario_Its) * 100) / 100, "###,##0.00")
    
End Sub

Private Sub txtbxSerie_Sai_GotFocus()
   'Seleciona o texto digitado anteriormente
   P_SelecTexto txtbxSerie_Sai
End Sub

Private Sub txtbxSerie_Sai_LostFocus()
   'Ajustando a série
    P_AjustaTexto txtbxSerie_Sai
    If txtbxSerie_Sai.Text = "####" Then
    'Esta série é reservada ao cupom fiscal
       txtbxSerie_Sai.Text = ""
       txtbxSerie_Sai.SetFocus
    End If
End Sub

Private Sub txtbxUnitario_Its_GotFocus()
    'Selecionando o texto digitado anteriormente
    P_ValorGotFocus txtbxUnitario_Its
End Sub

Private Sub txtbxUnitario_Its_LostFocus()
    'Ajustando o número digitado
    P_ValorLostFocus txtbxUnitario_Its
    txtbxQuantidade_Its_LostFocus
End Sub
