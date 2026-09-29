VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSaidaCupom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda de Produtos"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frControle 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   6960
         TabIndex        =   8
         Top             =   120
         Width           =   4455
         Begin VB.ComboBox cmbxOperacao_Sai 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmSaidaCupom.frx":0000
            Left            =   120
            List            =   "frmSaidaCupom.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Selecione o tipo de operação da nota fiscal."
            Top             =   360
            Width           =   4215
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
            TabIndex        =   9
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2175
         Left            =   9240
         TabIndex        =   50
         Top             =   4845
         Width           =   2175
         Begin VB.CommandButton cmndbtnFechar 
            Cancel          =   -1  'True
            Caption         =   "Sai&r "
            Height          =   855
            Left            =   120
            Picture         =   "frmSaidaCupom.frx":0004
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Fecha o formulário atual."
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnRegistraNota 
            Caption         =   "&Emitir Cupom Fiscal"
            Height          =   855
            Left            =   120
            Picture         =   "frmSaidaCupom.frx":014E
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Registra  a Nota Fiscal."
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6735
         Begin VB.ComboBox cmbxUF 
            Height          =   315
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   630
         End
         Begin VB.TextBox txtbxNome_Cli 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Tag             =   "00000"
            ToolTipText     =   "Informe o nome do cliente seguido de <Enter> para localizá-lo."
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtbxCNPJ 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            MaxLength       =   50
            TabIndex        =   7
            ToolTipText     =   "Informe o nome do cliente seguido de <Enter> para localizá-lo."
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblUF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&UF:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3840
            TabIndex        =   4
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblNome_Cli 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nome do Cliente /Razão Social:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   2280
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CP&F/CNPJ:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4560
            TabIndex        =   6
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.Frame frCupom 
         Caption         =   "Cupom Fiscal"
         Enabled         =   0   'False
         Height          =   975
         Left            =   9240
         TabIndex        =   11
         Top             =   840
         Width           =   2175
         Begin VB.TextBox txtbxNota_Sai 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   13
            ToolTipText     =   "Informe o número da nota fiscal."
            Top             =   480
            Width           =   1260
         End
         Begin VB.TextBox txtbxSerie_Sai 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   15
            ToolTipText     =   "Informe a série da nota fiscal."
            Top             =   480
            Width           =   550
         End
         Begin VB.Label lblSerie_Sai 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Série:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblNota_Sai 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Número:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   75
         TabIndex        =   18
         Top             =   4440
         Width           =   9015
         Begin VB.TextBox txtbxQuantidade_Its 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Height          =   285
            Left            =   4440
            MaxLength       =   10
            TabIndex        =   24
            Text            =   "0,00"
            ToolTipText     =   "Informe a quantidade (use vírgula para separar os decimais, não ponto)."
            Top             =   360
            Width           =   900
         End
         Begin VB.ComboBox cmbxSituacaoTri 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   32
            ToolTipText     =   "Permite selecionar o tipo da situação tributária."
            Top             =   960
            Width           =   3615
         End
         Begin VB.ComboBox cmbxBaseCalculo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmSaidaCupom.frx":06D8
            Left            =   3840
            List            =   "frmSaidaCupom.frx":06E5
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Permite selecionar o tipo de Base de Cálculo aplicada sobre o produto."
            Top             =   960
            Width           =   2535
         End
         Begin VB.CommandButton cmndbtnIncluiItem 
            Caption         =   "&Adcionar"
            Height          =   375
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Inclui a mercadoria especificada à nota fiscal."
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmndbtnExcluiItem 
            Caption         =   "E&xcluir"
            Height          =   375
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Exclui a mercadoria selecionada."
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtbxDesconto_Its 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            TabIndex        =   28
            Text            =   "0,00"
            ToolTipText     =   "Informe o valor do desconto sobre a mercadoria."
            Top             =   360
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtbxUnitario_Its 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            Height          =   285
            Left            =   5520
            TabIndex        =   26
            Text            =   "0,00"
            ToolTipText     =   "Informe o valor unitário."
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxNome_Pro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   20
            ToolTipText     =   "Informe o nome da mercadoria seguido de <ENTER> para localizá-lo."
            Top             =   360
            Width           =   3615
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
         Begin VB.Label lblQuantidade_Its 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Quantidade:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4440
            TabIndex        =   23
            Top             =   120
            Width           =   870
         End
         Begin VB.Label lblUnit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "U&nitário:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   25
            Top             =   120
            Width           =   585
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
            TabIndex        =   53
            Top             =   720
            Width           =   405
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
            TabIndex        =   31
            Top             =   750
            Width           =   1380
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
            TabIndex        =   33
            Top             =   750
            Width           =   1200
         End
         Begin VB.Label lblMarca_Mer 
            Appearance      =   0  'Flat
            Caption         =   "Marca"
            ForeColor       =   &H80000007&
            Height          =   225
            Left            =   120
            TabIndex        =   70
            ToolTipText     =   "Indica a Marca da mercadoria selecionada."
            Top             =   360
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.Label lblTotal_Its 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6480
            TabIndex        =   35
            ToolTipText     =   "Indica o valor total do produto."
            Top             =   960
            Width           =   900
         End
         Begin VB.Label lblDesconto_Its 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "De&sconto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6480
            TabIndex        =   27
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Totais:"
         Height          =   1215
         Left            =   75
         TabIndex        =   54
         Top             =   5805
         Width           =   4575
         Begin VB.Label lblItens 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1200
            TabIndex        =   72
            Tag             =   "!"
            ToolTipText     =   "Informa o valor total de ICMS sobre a Base de Cálculo."
            Top             =   930
            Width           =   900
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Itens Lanç:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1245
            TabIndex        =   71
            Top             =   720
            Width           =   795
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
            TabIndex        =   69
            Top             =   200
            Width           =   975
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
            TabIndex        =   68
            Top             =   200
            Width           =   840
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
            TabIndex        =   67
            Top             =   200
            Width           =   930
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
            TabIndex        =   66
            Top             =   735
            Width           =   885
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
            TabIndex        =   65
            Top             =   195
            Width           =   960
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
            TabIndex        =   64
            Tag             =   "!"
            ToolTipText     =   "Informa o Valor total da Base de Cálculo."
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
            TabIndex        =   63
            Tag             =   "!"
            ToolTipText     =   "Informa o valor total de ICMS sobre a Base de Cálculo."
            Top             =   405
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
            TabIndex        =   62
            Tag             =   "!"
            ToolTipText     =   "Informa o valor total sobre a Base de Cálculo do ICMS Substuição."
            Top             =   400
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
            TabIndex        =   61
            Tag             =   "!"
            ToolTipText     =   "Informa o valor total lançado no ICMS Substuição."
            Top             =   945
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
            TabIndex        =   60
            Tag             =   "!"
            ToolTipText     =   "Informa o Valor Total lançado."
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
            TabIndex        =   58
            Tag             =   "!"
            ToolTipText     =   "Informa o valor total dos produtos."
            Top             =   405
            Width           =   900
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
            TabIndex        =   57
            Top             =   740
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
            TabIndex        =   56
            Top             =   740
            Width           =   735
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
            TabIndex        =   55
            Tag             =   "!"
            ToolTipText     =   "Informa o valor Total do desconto."
            Top             =   940
            Width           =   900
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            X1              =   120
            X2              =   4440
            Y1              =   675
            Y2              =   675
         End
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrdItensSaida 
         Height          =   3135
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Informações Complementares:"
         Height          =   1215
         Left            =   4680
         TabIndex        =   36
         Top             =   5805
         Width           =   4455
         Begin VB.TextBox txtbxObservacoes_Sai 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   855
            Left            =   120
            TabIndex        =   37
            ToolTipText     =   "Entre com as informações complementares a nota fiscal."
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   9240
         TabIndex        =   43
         Top             =   2640
         Width           =   2175
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   $"frmSaidaCupom.frx":072A
            Top             =   340
            Width           =   900
         End
         Begin VB.TextBox txtbxDespesas_Sai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   45
            Text            =   "0,00"
            ToolTipText     =   $"frmSaidaCupom.frx":07B6
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descon&to:"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1080
            TabIndex        =   59
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblDespesaNormal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Despesa:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   675
         End
      End
      Begin VB.Frame frmeCabecalho2 
         Height          =   735
         Left            =   9240
         TabIndex        =   38
         Top             =   1800
         Width           =   2175
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
            Left            =   120
            TabIndex        =   40
            ToolTipText     =   "Entre com a data de emissão."
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbxPercICMS_Sai 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmSaidaCupom.frx":0842
            Left            =   1200
            List            =   "frmSaidaCupom.frx":0844
            Style           =   2  'Dropdown List
            TabIndex        =   42
            ToolTipText     =   "Selecione o Percentual do ICMS."
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblEmissao_Ent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "E&missão:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lblPercIcms 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&ICMS:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   41
            Top             =   120
            Width           =   435
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   9240
         TabIndex        =   47
         Top             =   3480
         Width           =   2175
         Begin VB.CommandButton cmndbtnBuscaVEnda 
            Caption         =   "&Buscar Venda"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Registra  a Nota Fiscal."
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton optnProd2 
            Caption         =   "&2 - Editar Produtos"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.OptionButton optnProd 
            Caption         =   "&1 - Salvar Produtos"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produtos &Vendidos:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   9015
      End
   End
End
Attribute VB_Name = "frmSaidaCupom"
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

Private Function F_EmiteCupomFiscal(strNome_Cli As String, Optional strCNPJ As String = "")
    'Esta função emite o cupom fiscal da nota

    Dim blnResp As Boolean
    Dim intCont As Integer, strST As String, strInf As String
    Dim strCod As String, strDescr As String, strUnidade As String
    Dim crcQuantidade As Currency, crcUnitario As Currency
    Dim strFormaP As String
    Dim crcDesconto As Currency
    Dim strCPF As String
    Dim sMsgECF As String
    Dim sTipoMsg As String
    sTipoMsg = F_BuscaUmCampo("tipomsgecf_par", "tab_Parametros", "")
    
    strInf = strInf & F_Centra(strNome_Cli, _
    48, "-") & Chr(13) & Chr(10)
    If Len(strCNPJ) = 18 Then
        strCPF = "CNPJ: " & strCNPJ
    ElseIf Len(strCNPJ) = 14 Then
        strCPF = "CPF: " & strCNPJ
    End If
    
    strInf = strInf & Chr(13) & F_Centra(strCPF, 48, "*")
    
    If sTipoMsg = "1" Then
        strInf = strInf & F_Centra("Tenha " & _
        IIf(Left(F_Cumprimenta(), 3) = "Bom", _
        "um ", "uma ") & F_Cumprimenta(), 48, "-") & Chr(13) & Chr(10)
    Else
        sMsgECF = F_BuscaUmCampo("msgecf_Par", "tab_parametros", "")
        strInf = strInf & sMsgECF
    End If

    'Carrega o form de formas de pagamento
    'Passando o valor total da nota
    frmFormas_Pagamento.lblRestante = lblTotalNota_Sai.caption
    frmFormas_Pagamento.txtbxValor.Text = lblTotalNota_Sai.caption
    frmFormas_Pagamento.Show 1
    strFormaP = frmFormas_Pagamento.Tag
    Unload frmFormas_Pagamento
    P_Status "Emitindo Cupom Fiscal"
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
      'lblIPI_Sai.caption = "0,00"
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
          
          'If optnbtnCupomFiscal.Value Then
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
          'End If
          
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
          'lblIPI_Sai.caption = F_ValToText(F_TextToVal _
          (0) + F_TextToVal(.Text))
          
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
      0 + _
      F_TextToVal(txtbxDespesas_Sai.Text) - _
      F_TextToVal(lblDesconto_Sai.caption))
      
      lblItens.caption = msflxgrdItensSaida.Rows - 1
      
    End With
End Sub
Private Sub P_LimpaItem()
    'Resetando controles para o padrão
    txtbxNome_Pro.Text = ""
    txtbxNome_Pro.Tag = ""
    lblUnidade_Its.caption = ""
    txtbxQuantidade_Its.Text = "0,00"
    txtbxQuantidade_Its.Tag = "0,00"
    txtbxDesconto_Its.Text = "0,00"
    txtbxUnitario_Its.Text = "0,00"
    lblTotal_Its.caption = "0,00"
    'txtbxIPI_Its.Text = "0,00"
    'txtbxICMS_Its.Text = cmbxPercICMS_Sai.Text
    
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
   strSQL = "Select count(*) as Total From TAB_SAIDAS Where "
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

Private Sub cmbxBaseCalculo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOperacao_Sai_Click()
    If Not F_Devolucao Then
       cmbxBaseCalculo.ListIndex = 1
       cmbxBaseCalculo.Enabled = False
       cmbxSituacaoTri.ListIndex = 0
       'cmbxSituacaoTri.Enabled = False
    Else
       cmbxBaseCalculo.Enabled = True
       cmbxSituacaoTri.Enabled = True
    End If
    cmbxBaseCalculo.Enabled = True
       cmbxSituacaoTri.Enabled = True
End Sub





Private Sub cmbxOperacao_Sai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacaoTri_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnExcluiItem_Click()
    Dim strNome_Mer As String
    With msflxgrdItensSaida
        If .Rows > 1 Then
            .Col = 1
            strNome_Mer = .Text
            If MsgBox("Deseja realmente excluir o item: " & strNome_Mer & " ?", _
                vbYesNo + vbCritical) = vbYes Then
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem .Row
                End If
            End If
        End If
        P_AjustaValores
    End With
End Sub

Private Sub cmndbtnFechar_Click()
    'Fecha o formulário
    If MsgBox("Deseja Realmente Sair?", vbYesNo + vbQuestion) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmndbtnIncluiItem_Click()
    If Not F_VerificaCabecalho Then
       Exit Sub
    End If
    
    If F_VerificaNota Then
       'Já existe uma nota fiscal
       Exit Sub
    End If
    
    If F_Vazio(txtbxNome_Pro.Tag) Then
       'Sem uma mercadoria não há como incluir o item
       MsgBox "Informe o Nome do Produto e Tecle <ENTER>!", vbExclamation
       txtbxNome_Pro.SetFocus
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
    
    
    'Verificando Desconto
'    If F_TextToVal(txtbxDesconto_Its.Text) >= _
'    F_TextToVal(lblTotal_Its) Then
'        'Valor do desconto maior ou igual o valor de venda
'
'        MsgBox "Desconto muito alto!!", vbExclamation
'       'txtbxDesconto_Its.SetFocus 'Focalizando para entrada de desconto
'        Exit Sub
'    End If
        
    If CCur(lblTotal_Its.caption) <= 0 Then
        MsgBox "Informe um valor correto!", vbExclamation
        txtbxQuantidade_Its.SetFocus
        Exit Sub
    End If
        
    With msflxgrdItensSaida
         
         If val(cmndbtnIncluiItem.Tag) <= 0 Then
            'Incluindo
            .AddItem Format(txtbxNome_Pro.Tag, "0##\.###\-#")
            .Row = .Rows - 1
         Else
            'Editando
            .Row = val(cmndbtnIncluiItem.Tag)
            .Col = 0
            .Text = Format(txtbxNome_Pro.Tag, "0##\.###\-#")
            cmndbtnIncluiItem.Tag = ""
         End If
         'Nome Mercadoria
         .Col = 1: .Text = txtbxNome_Pro.Text
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
         .Col = 7: .Text = cmbxPercICMS_Sai.Text   'txtbxICMS_Its.Text
         'Percentual do IPI
         .Col = 8: .Text = "0" 'Format((CCur(txtbxIPI_Its.Text) * 100) _
         / CCur(lblTotal_Its.caption), "00")
         'IPI
         .Col = 9: .Text = "0" 'txtbxIPI_Its.Text
         'Código da marca
         .Col = 10: .Text = "" 'lblMarca_Mer.Tag
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
     txtbxNome_Pro.SetFocus
End Sub
Sub grava()
    Dim strSQL As String
    Dim sCod As String
    Dim sUnit As String
    Dim sDesc As String
    Dim sQTD As String
    Dim iCount As Integer
    strSQL = "Insert Into TAB_saidas ("
    strSQL = strSQL & "COD_CLI, NOTA_SAI, SERIE_SAI, EMISSAO_SAI, "
    strSQL = strSQL & "VALOR_SAI,"
    strSQL = strSQL & "DESCONTO_SAI, ICMS_SAI, "
    strSQL = strSQL & "OPERACAO_SAI, "
    strSQL = strSQL & "UF_SAI,COD_VEN,COD_USU "
    strSQL = strSQL & ") Values ( "
    strSQL = strSQL & "'" & txtbxNome_Cli.Tag & "', "
    strSQL = strSQL & "'" & txtbxNota_Sai.Text & "', "
    strSQL = strSQL & "'" & txtbxSerie_Sai.Text & "', "
    strSQL = strSQL & F_GravaData(txtbxEmissao_Sai.Text) & ", "
    strSQL = strSQL & F_GravaDec(lblTotalNota_Sai.caption) & ", "
    strSQL = strSQL & F_GravaDec(lblDesconto_Sai.caption) & ", "
    strSQL = strSQL & F_GravaDec(Left(cmbxPercICMS_Sai.Text, 2)) & ", "
    strSQL = strSQL & "'" & Left(cmbxOperacao_Sai.Text, 4) & "', "
    strSQL = strSQL & "'" & cmbxUF.Text & "', "
    
    strSQL = strSQL & "'0000000','000') "
    
    banco.Execute strSQL
    For iCount = 1 To msflxgrdItensSaida.Rows - 1
        With msflxgrdItensSaida
            .Row = iCount
            .Col = 0
            sCod = F_SoNumero(.Text)
            .Col = 4
            sQTD = .Text
            .Col = 5
            sUnit = .Text
            .Col = 12
            sDesc = .Text
            strSQL = "insert into tab_itsai(cod_pro,nota_sai,serie_sai,unit_its,"
            strSQL = strSQL & "qtd_its,desconto_its,seq_its)values("
            strSQL = strSQL & "'" & sCod & "',"
            strSQL = strSQL & "'" & txtbxNota_Sai.Text & "',"
            strSQL = strSQL & "'" & txtbxSerie_Sai.Text & "',"
            strSQL = strSQL & F_GravaDec(sUnit) & ","
            strSQL = strSQL & F_GravaDec(sQTD) & ","
            strSQL = strSQL & F_GravaDec(sDesc) & ","
            
            strSQL = strSQL & iCount & ")"
            banco.Execute strSQL
        End With
    Next
    
End Sub

Private Sub cmndbtnRegistraNota_Click()
    Dim strSQL         As String 'Instrução SQL
    'On Error GoTo TrataErro
    P_Status "Emitindo Cupom Fiscal"
    If txtbxNome_Cli.Text = "" Then
        MsgBox "Informe o Nome do Cliente!", vbExclamation
        txtbxNome_Cli.SetFocus
        Exit Sub
    End If
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
    
       If Not F_EmiteCupomFiscal(txtbxNome_Cli.Text, txtbxCNPJ.Text) Then
          Exit Sub
       End If
    
    grava
    P_Status ""
    If MsgBox("Deseja imprimir outro Cupom?", vbYesNo + vbQuestion) = vbYes Then
        msflxgrdItensSaida.Rows = 1
        P_LimpaItem
        Cupom
        txtbxNome_Cli.SetFocus
        
        Exit Sub
    Else
        Unload Me
    End If
    Exit Sub
TrataErro:
    P_MostraErro Me.caption
    
End Sub


Private Sub Form_Activate()
    MudaAparencia Me
    Me.caption = "Emissão ECF " & strNome_Emp & " - CNPJ: " & strCNPJ_Emp & " - " & "IE: " & strIE_Emp
    
    txtbxEmissao_Sai.Text = Date
    txtbxEmissao_Sai.Enabled = False
    
   If Me.Tag = "" And F_Vazio(txtbxNota_Sai.Text) Then
      txtbxEmissao_Sai.Text = Date
      'optnbtnCupomFiscal.Value = True
      'optnbtnCupomFiscal_Click
      'optnbtnNotaFiscal.Value = True
      blnFlag = True
   End If
   
   'Carregando ComboBox
  
   
   'Configurando grid
   
   If Not F_Vazio(Me.Tag) Then
        
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
        ' P_MostraTeclaAtalho strMens
    End If
End Sub

Private Sub Form_Load()
    P_Status "Iniciando Lançamento"
    P_PreencheNotaSaida cmbxOperacao_Sai
    P_PreencheSituacaoTributaria cmbxSituacaoTri
    P_PreencheUF cmbxUF
    cmbxOperacao_Sai.ListIndex = 0
    cmbxSituacaoTri.ListIndex = 0
    Cupom
    With msflxgrdItensSaida
       .Rows = 1
       .Cols = 16
       .ColWidth(0) = 950  'Código da mercadoria
       .ColWidth(1) = 4500 'Descrição do produto
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
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub msflxgrdItensSaida_DblClick()
   'Permite editar a mercadoria
   With msflxgrdItensSaida
     If .Row > 0 Then
        cmndbtnIncluiItem.Tag = Str(.Row)
        .Col = 1
        txtbxNome_Pro.Text = Trim(.Text)
        txtbxNome_Pro_DblClick
        Exit Sub
     End If
    End With
End Sub


Private Sub Cupom()
    txtbxNota_Sai.Text = F_NumeroCupom
    If txtbxNota_Sai.Text = "" Then
        MsgBox "Não é possível registrar cupom! " & Chr(13) & "A Impressora ECF Não foi Encontrada", vbCritical
        P_Fim
    End If
    P_PreencheECFTaxasICMS cmbxPercICMS_Sai
    'txtbxNota_Sai.Enabled = False
    
    txtbxSerie_Sai.Text = "CUPM"
    'txtbxSerie_Sai.Enabled = False
    frCupom.Enabled = False
    'cmbxPercICMS_Sai.ListIndex = IIf(blnMercadTrib, 0, 1)
    cmbxPercICMS_Sai.ListIndex = 0
    cmbxPercICMS_Sai.Enabled = True
    cmbxOperacao_Sai.ListIndex = IIf(blnMercadTrib, 1, 0)
    'txtbxICMS_Its.Enabled = False
End Sub






Private Sub msflxgrdItensSaida_GotFocus()
'    msflxgrdItensSaida.FocusRect = flexFocusHeavy
    
End Sub

Private Sub msflxgrdItensSaida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmndbtnExcluiItem_Click
    End If
End Sub

Private Sub msflxgrdItensSaida_LostFocus()
'    msflxgrdItensSaida.FocusRect = flexFocusNone
End Sub

Private Sub optnProd_Click()
    If msflxgrdItensSaida.Rows > 1 Then
        If optnProd.Tag = "X" Then
            optnProd.Tag = ""
            Exit Sub
        End If
        If MsgBox("Isto irá limpar os Produtos já lançados!" & _
            Chr(13) & "Deseja Continuar?", vbYesNo + vbCritical) = vbYes Then
            msflxgrdItensSaida.Rows = 1
        Else
            optnProd2.Tag = "X"
            optnProd2.Value = True
            
        End If
    End If
End Sub

Private Sub optnProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub optnProd2_Click()
    If optnProd2.Tag = "X" Then
        optnProd2.Tag = ""
        Exit Sub
    End If
    If msflxgrdItensSaida.Rows > 1 Then
        If MsgBox("Isto irá limpar os Produtos já lançados!" & _
            Chr(13) & "Deseja Continuar?", vbYesNo + vbCritical) = vbYes Then
            msflxgrdItensSaida.Rows = 1
        Else
            optnProd.Tag = "X"
            optnProd.Value = True
        End If
    End If
End Sub

Private Sub optnProd2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_GotFocus()
    P_SelecTexto txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_LostFocus()
    txtbxCNPJ.Text = F_SoNumero(txtbxCNPJ.Text)
    If txtbxCNPJ.Text = "" Then
        
        Exit Sub
    End If
    If Len(txtbxCNPJ.Text) = "14" Then
        If Not F_CGC(txtbxCNPJ.Text) Then
            txtbxCNPJ.SetFocus
        End If
        P_FormataCGC txtbxCNPJ
    ElseIf Len(txtbxCNPJ.Text) = "11" Then
        If Not F_CPF(txtbxCNPJ.Text) Then
            txtbxCNPJ.SetFocus
        End If
        P_FormataCPF txtbxCNPJ
    Else
        MsgBox "Confira sua Digitação!", vbExclamation
        txtbxCNPJ.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtbxDesconto_Its_GotFocus()
    'Selecionando texto digitado
    P_ValorGotFocus txtbxDesconto_Its
End Sub

Private Sub txtbxDesconto_Its_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
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

Private Sub txtbxICMS_Its_Change()

End Sub

Private Sub txtbxIPI_Its_Change()

End Sub

Private Sub txtbxNome_Cli_DblClick()
  '  frmSelectCli.txtbxNome.Tag = "C"
'    frmSelectCli.txtbxNome.Text = txtbxNome_Cli.Text
'
'    On Error Resume Next
'
'    frmSelectCli.Show 1
'    txtbxNome_Cli.Text = frmSelectCli.dtBusca.Recordset!Nome
'    txtbxNome_Cli.Tag = frmSelectCli.dtBusca.Recordset!cod
'    cmbxUF.Text = frmSelectCli.dtBusca.Recordset!uf_Cli
'    txtbxCNPJ.Text = frmSelectCli.dtBusca.Recordset!CNPJ
'    'Indica a situação da tributação da mercadoria
'    blnMercadTrib = (frmSelectCli.dtBusca.Recordset!tipo_cli = "J" _
'    And cmbxUF.Text <> "GO")
'
'    Unload frmSelectCli
    ' txtbxNum_Not.SetFocus

    
    
    
    
    
    
    
    
    
    
   ' P_AjustaValores
    'SendKeys Chr(9) 'Vai para o próximo controle
    
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
Private Sub txtbxNome_Mer_DblClick()
    Dim p As Produto
    p = F_BuscaProd
     cmbxSituacaoTri.Tag = "" ' situacao trib da mercadoria
     'Verificando a Situação Tributária da Mercadoria
     
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
End Sub


Private Sub txtbxNome_Pro_DblClick()
    Dim p As Produto
    Dim iCount As Integer
    If txtbxNome_Cli.Text = "" Then
        MsgBox "Informe o Nome do Cliente Primeiro!", vbExclamation
        txtbxNome_Cli.SetFocus
        Exit Sub
    End If
    If cmbxUF.Text = "" Then
        MsgBox "Informe o Estado do Cliente primeiro!", vbExclamation
        cmbxUF.SetFocus
        Exit Sub
    End If
    p = F_BuscaProdSaida(txtbxNome_Pro.Text)
    With msflxgrdItensSaida
        For iCount = 1 To .Rows - 1
            .Row = iCount
            .Col = 0
            If F_SoNumero(.Text) = p.cod_pro Then
                If MsgBox("Este item já se econtra lançado!" & Chr(13) & _
                "Deseja Alterar sua Quantidade?", vbYesNo + vbQuestion) = vbYes Then
                cmndbtnIncluiItem.Tag = Str(.Row)
                Else
            
                    txtbxNome_Pro.SetFocus
                    Exit Sub
                End If
            End If
            
        Next
        
    End With
    
    txtbxNome_Pro.Text = p.nome_pro
    txtbxNome_Pro.Tag = p.cod_pro
    If optnProd.Value = True Then
        txtbxUnitario_Its.Text = FormatVal(F_BuscaUmCampo("preco_sai", "tab_estoque", _
            "cod_pro='" & p.cod_cat & "' and cod_loj='" & strCod_Loj & "'"))
    Else
        txtbxUnitario_Its.Text = FormatVal(p.preco)
    End If
    lblUnidade_Its.caption = F_BuscaUmCampo("UNIDADE_PRO", "TAB_PRODUTOS", _
                    "COD_PRO='" & p.cod_pro & "' ")
    txtbxQuantidade_Its.SetFocus
End Sub


Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxNome_Pro_DblClick
    End If
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

Private Sub txtbxQuantidade_Its_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQuantidade_Its_LostFocus()
   'Ajusta o valor digitado e multiplica para achar valor total
   
   P_ValorLostFocus txtbxQuantidade_Its
   If CCur(txtbxQuantidade_Its.Text) <= 0 And _
   Not F_Vazio(txtbxNome_Pro.Tag) And F_TextToVal(txtbxQuantidade_Its.Tag) > 0 Then
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

Private Sub txtbxUnitario_Its_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnitario_Its_LostFocus()
    'Ajustando o número digitado
    P_ValorLostFocus txtbxUnitario_Its
    txtbxQuantidade_Its_LostFocus
End Sub
