VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProdutoSaida 
   Caption         =   "Selecionar um Produto"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   17790
   ForeColor       =   &H00000000&
   Icon            =   "frmProdutoSaida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   17790
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcRight 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8835
      Left            =   12495
      ScaleHeight     =   8805
      ScaleWidth      =   5265
      TabIndex        =   30
      Top             =   495
      Width           =   5295
      Begin VB.PictureBox pcContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   5025
         TabIndex        =   52
         Top             =   5160
         Width           =   5055
         Begin VB.PictureBox pcImagem 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   240
            ScaleHeight     =   1695
            ScaleWidth      =   1575
            TabIndex        =   53
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid gridVeiculo 
         Bindings        =   "frmProdutoSaida.frx":014A
         Height          =   2055
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         Caption         =   "Veículos"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "veiculo"
            Caption         =   "Veiculo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "aplicacao"
            Caption         =   "Aplicação"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "fabricante"
            Caption         =   "Fabricante"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   1635,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2009,764
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridConversao 
         Bindings        =   "frmProdutoSaida.frx":015C
         Height          =   1935
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         Caption         =   "Conversão"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "nome_mar"
            Caption         =   "Marca"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "referencia"
            Caption         =   "Referência"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2129,953
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAplicacao 
         Caption         =   "Aplicação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   960
         TabIndex        =   38
         Top             =   4440
         Width           =   4185
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicação:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   4440
         Width           =   750
      End
      Begin VB.Label lblFabricante 
         AutoSize        =   -1  'True
         Caption         =   "Fabr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   36
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   35
         Top             =   4200
         Width           =   795
      End
      Begin VB.Label lblVeiculo 
         AutoSize        =   -1  'True
         Caption         =   "Veículo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   34
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Veiculo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   4200
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   17760
      TabIndex        =   0
      Top             =   0
      Width           =   17790
      Begin VB.PictureBox pcPreco 
         Appearance      =   0  'Flat
         BackColor       =   &H0094DEF3&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   10320
         ScaleHeight     =   495
         ScaleWidth      =   5055
         TabIndex        =   41
         Top             =   0
         Width           =   5055
         Begin VB.Label lblPreco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   2760
            TabIndex        =   45
            Top             =   0
            Width           =   630
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Preço ao Cliente R$:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1680
            TabIndex        =   44
            Top             =   0
            Width           =   945
         End
         Begin VB.Label lblQtd 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   480
            TabIndex        =   43
            Top             =   120
            Width           =   120
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         Picture         =   "frmProdutoSaida.frx":016E
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   150
         Width           =   240
      End
      Begin VB.Label lblProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione um Produto"
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
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmProdutoSaida.frx":04BE
      Height          =   6495
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      Caption         =   "Lista de Produtos"
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "cod_emp"
         Caption         =   "Empresa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ean_pro"
         Caption         =   "EAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cod_pro"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "codfab_pro"
         Caption         =   "Ref."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "nome_pro"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "nome_cat"
         Caption         =   "Categoria"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "nome_mar"
         Caption         =   "Marca"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "disponivel"
         Caption         =   "Qtde."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "multiplo_pro"
         Caption         =   "Múltiplo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "pv_pro"
         Caption         =   "Preço Unitário"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "preco"
         Caption         =   "Prç Atacado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "preco_cli"
         Caption         =   "Preço Final"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "qtd"
         Caption         =   "Emb. Atacado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "perc_imp"
         Caption         =   "Imposto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "ncm_pro"
         Caption         =   "NCM"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4754,835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1260,284
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column14 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   17790
      TabIndex        =   3
      Top             =   7035
      Width           =   17790
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   47
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDetalhe 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   7440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.PictureBox pcImg 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   6840
         Picture         =   "frmProdutoSaida.frx":04D0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   32
         Top             =   1560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.CheckBox chkCompleto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pesquisar texto completo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Mostrar Produtos Zerados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Tag             =   "!"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.PictureBox pcSimilares 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   2640
         Picture         =   "frmProdutoSaida.frx":0946
         ScaleHeight     =   975
         ScaleWidth      =   3870
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   3870
         Begin VB.CommandButton cmdListSim 
            Caption         =   "Listar similares (S)"
            Height          =   435
            Left            =   2040
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdAbrirSim 
            Caption         =   "Abrir"
            Height          =   435
            Left            =   1200
            TabIndex        =   24
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblEstoqueSim 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   600
            TabIndex        =   27
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos Similares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.PictureBox pcBotoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   8280
         Picture         =   "frmProdutoSaida.frx":C988
         ScaleHeight     =   1215
         ScaleWidth      =   3615
         TabIndex        =   20
         Top             =   960
         Width           =   3615
         Begin VB.CommandButton cmdOk 
            Caption         =   "Selecionar"
            Height          =   855
            Left            =   360
            Picture         =   "frmProdutoSaida.frx":1BCD6
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "Cancelar"
            Height          =   855
            Left            =   1920
            Picture         =   "frmProdutoSaida.frx":1C15D
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtVeiculo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   2775
      End
      Begin VB.PictureBox pcQtd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   5760
         ScaleHeight     =   825
         ScaleWidth      =   1545
         TabIndex        =   16
         Top             =   360
         Width           =   1575
         Begin VB.TextBox txtQtd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   17
            Text            =   "0,00"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. (F3)"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox txtbxBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         Picture         =   "frmProdutoSaida.frx":1C6EB
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         Picture         =   "frmProdutoSaida.frx":1CA96
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdVeiculo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         Picture         =   "frmProdutoSaida.frx":1CE41
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdPesquisar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         Picture         =   "frmProdutoSaida.frx":1D1EC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdEmpresa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         Picture         =   "frmProdutoSaida.frx":1D597
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblFab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4560
         TabIndex        =   50
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   4440
         TabIndex        =   49
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alt + L (limpa veículos)"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Label lblVeicCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Veículo(Alt + v):"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar por: (Ctrl + L)"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmProdutoSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public WithEvents adoVei As ADODB.Recordset
Attribute adoVei.VB_VarHelpID = -1
Public WithEvents adoConversao As ADODB.Recordset
Attribute adoConversao.VB_VarHelpID = -1
Public ignora_empresa_qtd_produto As Boolean
Public cod_cli As String
Dim desab As Boolean
Dim listaImagens As New Collection
Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If desab Then Exit Sub
    VerSim ado!cod_pro
    lblProduto.caption = ado!nome_pro
    lblPreco.caption = FormatVal(ado!preco_cli)
    lblQtd.caption = (ado!disponivel)
    buscaVei
    buscaConversao
    
    BuscaImg
    
     If TiraNull(ado!detalhe_pro) <> "" Then
            txtDetalhe.Text = ado!detalhe_pro
            txtDetalhe.Visible = True
        Else
            txtDetalhe.Text = ""
            txtDetalhe.Visible = False
        End If
    
End Sub


Sub VerSim(cod_pro As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If cod_pro = "" Then
        pcSimilares.Visible = False
    End If
    strSQL = "select count(*) as total from tab_similares where cod_pro='" & cod_pro & "' "
   
    abreObjetoAdo rPesq, strSQL
   
    If rPesq!total > 0 Then
        pcSimilares.Visible = True
        lblEstoqueSim.caption = BuscaEstoqueSim(cod_pro)
    Else
        pcSimilares.Visible = False
    End If
   
   
   
   
    
    
End Sub

Private Sub adoVei_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    lblVeiculo.caption = adoVei!Veiculo
    lblFabricante.caption = adoVei!fabricante
    lblAplicacao.caption = adoVei!aplicacao
    
    
    
End Sub

Private Sub chkTodas_Click()
    busca
End Sub

Private Sub cmdAbrirSim_Click()
    On Error Resume Next
    buscaSimilares ado!cod_pro
End Sub

Private Sub cmdCancel_Click()
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmdCat_Click()
    buscaCat
End Sub

Private Sub cmdEmpresa_Click()
    buscaEmpresa
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    
    e = sped_empresa.SelecionarSped_Empresa(txtEmpresa.Text)
    If e.codigo <> "" Then
        txtEmpresa.Text = e.codigo
        txtEmpresa.tag = e.codigo
        busca
    End If
        
    
    
End Sub

Private Sub cmdMarca_Click()
    BuscarMarca
End Sub

Private Sub cmdOk_Click()
    Dim m As Double
    On Error GoTo trata_Erro
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If txtQtd.Text = "" Then
            Exit Sub
        End If
        If !qtd > 0 Then
            txtQtd.Text = !qtd
        End If
        
        
        
        If CCur(txtQtd.Text) > !disponivel Then
            If !disponivel > 0 Then
             m = CCur(txtQtd.Text) Mod !disponivel
             If m > 0 Then
                 MsgBox "A quantidade informada deve ser múltipla de " & !qtd, vbExclamation
                 txtQtd.SetFocus
                 Exit Sub
            
             
             End If
            End If
        End If
        
        
        
        If txtbxBuscar.Text = !barra Then
            If !disponivel <> CCur(txtQtd.Text) Then
        
                    'MsgBox "Para este produto ser vendido em preço de atacado, a quantidade informada deve ser " & !QTD, vbExclamation
                    'Exit Sub
                'txtQtd.Text = !QTD
            End If
        End If
        Me.tag = !cod_pro
        Me.Hide
    End With
trata_Erro:
End Sub

Private Sub cmdPesquisar_Click()
    busca
End Sub

Private Sub cmdVeiculo_Click()
    selecionaVei
End Sub
Sub selecionaVei()
    Dim v As Veiculo
    v = Veiculos.BuscaVeiculo(txtVeiculo.Text, "")
    If v.id > 0 Then
        txtVeiculo.Text = v.descricao
        txtVeiculo.tag = v.id
        lblFab.caption = v.fabricante.descricao
        UltimoVeiculoProcurado = v
        busca
    
    End If
End Sub

Private Sub dtgrd_DblClick()
    cmdOk_Click
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
    If KeyCode = vbKeyN Or KeyCode = vbKeyP Or KeyCode = vbKeyM Then
        frmBusca.tag = "N"
        frmBusca.Show 1
        Me.tag = frmBusca.tag
        Unload frmBusca
        busca
    ElseIf KeyCode = vbKeyS Then
        busca "", True
    End If
End Sub


Private Sub Form_Activate()
    ' 'MudaAparencia Me
    
    If bUsaVeiculo = False Then
        pcRight.Visible = False
        pcRight.width = 100
        lblVeiculo.Visible = False
        txtVeiculo.Visible = False
        cmdVeiculo.Visible = False
        lblVeicCap.Visible = False
        Form_Resize
    End If
    
    lblCliente.caption = F_BuscaUmCampo("nome_cli", TAB_CLIENTES, "cod_cli='" & cod_cli & "'")
    
    busca
    
    On Error Resume Next
    dtgrd.SetFocus
    dtgrd.MarqueeStyle = 3
End Sub
Sub buscaConversao()
    Dim strSQL As String
    strSQL = SelectProdutoConversao(ado!cod_pro)
    abreObjetoAdo adoConversao, strSQL, gridConversao
End Sub

Sub AbreImagens(cod_pro As String)
    frmImagens.tag = cod_pro
    frmImagens.txtProduto.Text = cod_pro
    frmImagens.txtProduto.tag = cod_pro
    
    frmImagens.Show 1
    
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtQtd.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
        Me.tag = ""
        Me.Hide
    End If
    If KeyCode = vbKeyV And Shift = vbAltMask Then
        txtVeiculo.SetFocus
    End If
    
    If KeyCode = vbKeyL And Shift = vbAltMask Then
        limpaVeiculo
    End If
    If KeyCode = vbKeyL And Shift = vbCtrlMask Then
        txtbxBuscar.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    Me.Visible = False
    Set ado = New ADODB.Recordset
    Set adoVei = New ADODB.Recordset
    Set adoConversao = New ADODB.Recordset
    resizeForm Me
    Form_Resize
End Sub
Sub busca(Optional limit As String = "", Optional bSim As Boolean = False)
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim iCount As Integer
    If cod_cli = "" Then
        cod_cli = "0"
    End If
        If txtbxBuscar.Text = "" And txtVeiculo.tag = "" And txtMarca.tag = "" And txtCat.tag = "" Then
            MsgBox "Informe parte do nome ou descrição.", vbInformation
            txtbxBuscar.SetFocus
            Exit Sub
        End If
                
                
        strSQL = "select distinct tc.nome_Cat, tc.descmax_cat,tm.nome_mar, ti.perc_imp,"
        strSQL = strSQL & "tp.*,(select tp.pv_pro-(tp.pv_pro*desconto_cli/100) from " & TAB_CLIENTES
        strSQL = strSQL & " where cod_cli=" & cod_cli & ") as preco_cli,"
        strSQL = strSQL & " e.er-e.reservado as disponivel,e.reservado, "
        strSQL = strSQL & " e.cod_emp,e.er,e.ef,tp.er_pro , tpa.qtd,tpa.preco,tpa.barra,  "
        strSQL = strSQL & "     (SELECT GROUP_CONCAT(caminho SEPARATOR ',') AS caminho FROM imagem_produto AS img where img.cod_pro=tp.cod_pro ORDER BY id desc ) AS caminho"
        strSQL = strSQL & " from "
        strSQL = strSQL & " " & TAB_PRODUTOS & " as tp "
        strSQL = strSQL & " inner join estoque as e on tp.cod_pro=e.cod_pro "
        strSQL = strSQL & " left join produto_atacado as tpa on tp.cod_pro=tpa.codigo "
        'strSQL = strSQL & " left join tab_codbarras as tb on tb.cod_pro=tp.cod_pro "
        strSQL = strSQL & " left join " & TAB_MARCAS & " as tm on tm.cod_mar=tp.cod_mar "
        strSQL = strSQL & " left join " & TAB_CATEGORIAS & " as tc on tc.cod_cat=tp.cod_cat "
        strSQL = strSQL & " left join " & TAB_IMPOSTOS & " as ti on ti.cod_imp=tp.cod_imp"
        strSQL = strSQL & " left join produto_veiculo as pv on tp.cod_pro=pv.cod_pro"
        strSQL = strSQL & " left join veiculo as v on pv.id_vei=v.id "
        strSQL = strSQL & " left join produto_conversao as conv on tp.cod_pro=conv.cod_pro "
        
        
        If bSim Then
            strSQL = strSQL & " left join tab_similares as ts on tp.cod_pro=ts.cod_pro or tp.cod_pro=ts.cod_sim "
            strSQL = strSQL & " where ts.cod_pro='" & ado!cod_pro & "' or ts.cod_sim='" & ado!cod_pro & "' "
        Else
        strSQL = strSQL & " where 1>0 " 'tc.cod_cat=tp.cod_cat  and tm.cod_mar=tp.cod_mar "
       
        
                
            strSQL = strSQL & " and  (tp.codbarra_PRO = '"
            strSQL = strSQL & Me.tag & "' "
            strSQL = strSQL & " or  ean_pro = '"
            strSQL = strSQL & Me.tag & "' "
            
            strSQL = strSQL & " or  codfab_pro = '"
            strSQL = strSQL & Me.tag & "' "
            
            
            strSQL = strSQL & " or  tpa.barra = '"
            strSQL = strSQL & Me.tag & "' "
            
            strSQL = strSQL & " or v.descricao like '%" & txtbxBuscar.Text & "%' "
            strSQL = strSQL & " or conv.referencia like '" & txtbxBuscar.Text & "%' "
            
             
            If bUsaVendaSupermercado = False Then
            
                strSQL = strSQL & " or tp.COD_PRO = '"
                strSQL = strSQL & (Me.tag) & "' "
            
                strSQL = strSQL & " or tp.CODfab_PRO like '%"
                strSQL = strSQL & (Me.tag) & "%' "
            End If
            If bUsaVendaSupermercado = True Then
            
                strSQL = strSQL & " or (tp.CODbarra_PRO = '"
                strSQL = strSQL & Left(Me.tag, 7) & "' "
                strSQL = strSQL & " and tp.balanca_pro='S') "
                
                strSQL = strSQL & " or (ean_pro= '"
                strSQL = strSQL & Left(Me.tag, 7) & "' "
                strSQL = strSQL & " and tp.balanca_pro='S') "
                
                strSQL = strSQL & " or tp.COD_PRO = '"
                strSQL = strSQL & (Me.tag) & "' "
                
                strSQL = strSQL & " or tp.CODfab_PRO = '"
                strSQL = strSQL & (Me.tag) & "' "
                
            End If
            
            
            tipo_busca = "B"
        
            If bUsaVendaSupermercado = True Then
                If F_SoNumero(Me.tag) <> Me.tag Then
                    strSQL = strSQL & " or NOME_PRO like '"
                    strSQL = strSQL & Me.tag & "%'"
                End If
            
            Else
                strSQL = strSQL & " or NOME_PRO like '%"
                strSQL = strSQL & Me.tag & "%'"
            End If
            
            strSQL = strSQL & ")"
        End If
        
        
        strSQL = strSQL & " and tp.visivel_pro='S' "
        
        'strSQL = strSQL & " and tp.cod_pro=tb.cod_pro "
    '   strSQL = strSQL & " and tp.cod_pro=ts.cod_pro "
        'strSQL = strSQL & " and tc.cod_cat=tp.cod_Cat and tm.cod_mar=tp.cod_mar "
    
        If chkTodas.Value = 0 Then
    
    '        strSQL = strSQL & " and (ts.er_pro-ts.r_pro) > 0 "
        End If
    
    
        If txtVeiculo.tag <> "" And txtVeiculo.tag <> "0" Then
            strSQL = strSQL & " and v.id='" & txtVeiculo.tag & "' "
        ElseIf txtVeiculo.Text <> "" Then
            strSQL = strSQL & " and v.descricao like '%" & txtVeiculo.Text & "%' "
        End If
        If txtMarca.tag <> "" Then
            strSQL = strSQL & " and tp.cod_mar='" & txtMarca.tag & "' "
        End If
        If txtCat.tag <> "" Then
            strSQL = strSQL & " and tp.cod_cat='" & txtCat.tag & "' "
        End If
        
        If txtEmpresa.tag <> "" Then
            strSQL = strSQL & " and e.cod_emp='" & txtEmpresa.tag & "' "
        End If
        strSQL = strSQL & " group by tp.cod_pro,e.cod_emp "
        strSQL = strSQL & " order by tp.nome_pro,e.cod_emp"
        
        abreObjetoAdo ado, strSQL, dtgrd
        
        If ignora_empresa_qtd_produto = True Then
           dtgrd.Visible = False
           
            If mesmo_produto Then
                If ado!qtd <> Null Then
                
                    txtQtd.Text = ado!qtd
                Else
                End If
                
                Me.tag = "x"
                Me.Hide
                Exit Sub
            Else
                dtgrd.Visible = True
            End If
        End If
        
        
        If ado.RecordCount = 1 Then
           
        
            cmdOk_Click
        Else
            
            frmProdutoSaida.Visible = True
        End If
End Sub
Function mesmo_produto() As Boolean
    On Error GoTo trata_Erro:
    Dim i As Integer
    Dim cod_pro As String
    On Error GoTo trata_Erro
    desab = True
    
    
    With ado
        .MoveFirst
        cod_pro = !cod_pro
        Do While Not .EOF
            If !cod_pro <> cod_pro Then
                ado.MoveFirst
                mesmo_produto = False
                desab = False
                Exit Function
            End If
            
            .MoveNext
        Loop
    End With
    mesmo_produto = True
    ado.MoveFirst
    
    desab = False
    Exit Function
trata_Erro:
    mesmo_produto = False
    
    
End Function
Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtgrd, Me, pcBot
    posicionaADireita pcRight, Me
    
    PosicionaGradeLeft dtgrd, pcRight, 200
    PosicionaGradeTop dtgrd, pcBot
    
    
    
   ' posicionaADireita pcBotoes, pcBot
   ' posicionaADireita pcPreco, pcBot
    
End Sub
Sub BuscaImg()
    On erro GoTo trata_Erro
    Dim img As Picture
    
    pcImagem.Cls
    
    If Not IsNull(ado!caminho) Then
          
        Set listaImagens = F_SeparaTexto(ado!caminho, ",", False)
        pcImg.Visible = True
        pcImagem.Cls
        
        
        
        'mostraMiniatura  ado!caminho, pcRight, 2000, 600, pcRight.width, pcRight.height
        'Set img = LoadPicture(listaImagens(1))
        'pcImagem.PaintPicture LoadPicture(listaImagens(1)), 0, 0
        'pcImagem.tag = listaImagens(1)
        'pcImagem.Visible = True
        
        carregaImagemFit pcImagem, pcContainer, listaImagens(1)
        pcContainer.Visible = True
        'Me.SetFocus
        
        'fitImage pcImagem, pcContainerImg, listaImagens(1)
    Else
        pcImg.Visible = False
        pcImagem.Cls
        pcImagem.Visible = False
        pcContainer.Visible = False
    End If
trata_Erro:
End Sub

Sub buscaVei()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = SelectProduto_Veiculo(ado!cod_pro, txtVeiculo.Text)
    abreObjetoAdo adoVei, strSQL, gridVeiculo
    With adoVei
        Do While Not .EOF
            If txtVeiculo.tag = "" Then
                dtgrd.SetFocus
                Exit Sub
            End If
            If !id_vei = txtVeiculo.tag Then
                dtgrd.SetFocus
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
trata_Erro:
End Sub



Private Sub gridVeiculo_DblClick()
    dtgrd_DblClick
End Sub

Private Sub pcImagem_Click()
    Shell "explorer " & pcImagem.tag
End Sub

Private Sub pcImg_Click()
    On Error Resume Next
    AbreImagens ado!cod_pro
    
End Sub

Private Sub txtbxBuscar_GotFocus()
    P_SelecTexto txtbxBuscar
End Sub

Private Sub txtbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.tag = F_AjustaTexto(txtbxBuscar)
        busca
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub

Private Sub txtCat_Change()
    txtCat.tag = ""
    
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCat
    End If
End Sub
Sub buscaCat()
    Dim c As categoria.categoria
    c = BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        busca
    End If
End Sub
Sub BuscarMarca()
    Dim c As Marca.Marca
    c = buscaMarca(txtMarca.Text)
    If c.codigo <> "" Then
        txtMarca.Text = c.nome
        txtMarca.tag = c.codigo
        busca
    End If
End Sub

Private Sub txtMarca_Change()
    txtMarca.tag = ""
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        BuscarMarca
    End If
End Sub

Private Sub txtQtd_GotFocus()
    ganhaFoco txtQtd
End Sub

Private Sub txtQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey 13
End Sub

Private Sub txtQtd_LostFocus()
On Error Resume Next
    If CCur(txtQtd.Text) Mod ado!qtd > 0 Then
        txtQtd.BackColor = vbRed
    Else
        txtQtd.BackColor = vbGreen
    End If
    
    PerdeFoco txtQtd
End Sub

Private Sub txtVeiculo_Change()
    txtVeiculo.tag = ""
    UltimoVeiculoProcurado.id = 0
    UltimoVeiculoProcurado.descricao = ""
End Sub
Sub limpaVeiculo()
    txtVeiculo.tag = ""
    UltimoVeiculoProcurado.id = 0
    UltimoVeiculoProcurado.descricao = ""
    lblFab.caption = ""
End Sub
Private Sub txtVeiculo_GotFocus()
    P_SelecTexto txtVeiculo
End Sub

Private Sub txtVeiculo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        selecionaVei
    End If
End Sub
