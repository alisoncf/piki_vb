VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntradas 
   Caption         =   "Entrada de Produtos"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13920
   Icon            =   "frmEntradas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   13920
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4935
      Left            =   0
      TabIndex        =   112
      Top             =   1560
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8705
      _Version        =   393216
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "seq_ite"
         Caption         =   "..."
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
         DataField       =   "cod_Pro"
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
      BeginProperty Column02 
         DataField       =   "nome_pro"
         Caption         =   "Descrição"
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
         DataField       =   "unit_ite"
         Caption         =   "Preço"
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
         DataField       =   "precofuturo_ite"
         Caption         =   "Preço Futuro"
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
         DataField       =   "total"
         Caption         =   "Total"
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
         DataField       =   "ipi_ite"
         Caption         =   "Ipi"
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
         DataField       =   "custo_ite"
         Caption         =   "Custo"
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
      BeginProperty Column08 
         DataField       =   "sujerido_ite"
         Caption         =   "Sugerido"
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
         BeginProperty Column00 
            ColumnWidth     =   434,835
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4860,284
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   10440
      ScaleHeight     =   915
      ScaleWidth      =   2835
      TabIndex        =   110
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Modificar"
         Height          =   855
         Left            =   1440
         Picture         =   "frmEntradas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnAbreVenda 
         Caption         =   "&Confirmar"
         Height          =   855
         Left            =   0
         Picture         =   "frmEntradas.frx":12CC
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcProduto 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13890
      TabIndex        =   105
      Top             =   6585
      Width           =   13920
      Begin VB.Frame frOrd 
         BorderStyle     =   0  'None
         Caption         =   "Ordem:"
         Height          =   795
         Left            =   12240
         TabIndex        =   106
         Top             =   0
         Width           =   1215
         Begin VB.OptionButton optnDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optnCod 
            Caption         =   "Có&digo"
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optnSeq 
            Caption         =   "Seqüência"
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   39
         Text            =   "0,0000"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   37
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   41
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxPrecoFuturo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9840
         TabIndex        =   50
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxSujerido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8880
         TabIndex        =   48
         Text            =   "0,00"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtbxIpi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   43
         Text            =   "0,00"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmndbtnConfirma 
         Caption         =   "&Inserir"
         Height          =   615
         Left            =   10800
         Picture         =   "frmEntradas.frx":1774
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   30
         Width           =   1095
      End
      Begin VB.TextBox txtbxCodBarra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   33
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtbxIcms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   45
         Text            =   "0,0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtbxValorICMS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   46
         Text            =   "0,00"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cmbxEmb 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   35
         Top             =   345
         Width           =   735
      End
      Begin VB.Label lblCodPro 
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   111
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   6000
         TabIndex        =   40
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         Height          =   195
         Left            =   4320
         TabIndex        =   36
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário"
         Height          =   195
         Left            =   5040
         TabIndex        =   38
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pç. &Futuro:"
         Height          =   195
         Left            =   9840
         TabIndex        =   49
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pç. Suj.:"
         Height          =   195
         Left            =   8880
         TabIndex        =   47
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Tag             =   "B"
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lblIPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ipi(R$):"
         Height          =   195
         Left            =   6960
         TabIndex        =   42
         Top             =   120
         Width           =   510
      End
      Begin VB.Label lblIcmsCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I&CMS(%):(R$)"
         Height          =   195
         Left            =   7800
         TabIndex        =   44
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Emb.:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3480
         TabIndex        =   34
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C6AE86&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   13890
      TabIndex        =   82
      Top             =   8460
      Width           =   13920
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "F2 - Editar Item"
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
         Left            =   240
         TabIndex        =   88
         Top             =   80
         Width           =   1215
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "F3 - Excluir Item"
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
         Left            =   1920
         TabIndex        =   87
         Top             =   75
         Width           =   1290
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "F4 - Edição de Produtos"
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
         Left            =   3960
         TabIndex        =   86
         Top             =   75
         Width           =   1920
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "F5 - Novo Produto"
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
         Left            =   6600
         TabIndex        =   85
         Top             =   75
         Width           =   1440
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "F7 - Conclui N.F"
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
         Left            =   8760
         TabIndex        =   84
         Top             =   75
         Width           =   1230
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "F12 - Sair"
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
         Left            =   10680
         TabIndex        =   83
         Top             =   75
         Width           =   735
      End
   End
   Begin VB.PictureBox pcInfo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   13890
      TabIndex        =   54
      Top             =   7320
      Width           =   13920
      Begin VB.PictureBox pcResu 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1DDB0&
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   6720
         ScaleHeight     =   1110
         ScaleWidth      =   5145
         TabIndex        =   89
         Top             =   0
         Width           =   5175
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICMS:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4080
            TabIndex        =   104
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblPercIms 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   210
            Left            =   4080
            TabIndex        =   103
            Top             =   480
            Width           =   315
         End
         Begin VB.Label lblTotalProd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   210
            Left            =   840
            TabIndex        =   102
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Prod:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   840
            TabIndex        =   101
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Itens:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4680
            TabIndex        =   100
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblItens 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   210
            Left            =   4680
            TabIndex        =   99
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lançado:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblValoLan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   210
            Left            =   120
            TabIndex        =   97
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IPI Lanç:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2640
            TabIndex        =   96
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblIpi_Lan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   210
            Left            =   2640
            TabIndex        =   95
            Top             =   480
            Width           =   315
         End
         Begin VB.Label lblGastos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   210
            Left            =   3360
            TabIndex        =   94
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3360
            TabIndex        =   93
            Top             =   240
            Width           =   570
         End
         Begin VB.Label lblRest 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   210
            Left            =   1800
            TabIndex        =   92
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1800
            TabIndex        =   91
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resumo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2520
            TabIndex        =   90
            Top             =   0
            Width           =   690
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5040
            Y1              =   240
            Y2              =   240
         End
      End
      Begin VB.TextBox txtbxCustoSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "0,0000"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxLucroLiq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "0,0000"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxLucro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "0,0000"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxNovoCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxPreco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "0,0000"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxGastos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0,0000"
         Top             =   800
         Width           =   855
      End
      Begin VB.TextBox txtbxICMSSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "0,00"
         Top             =   800
         Width           =   495
      End
      Begin VB.TextBox txtbxTaxas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "0,00"
         Top             =   800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   59
         Text            =   "%"
         Top             =   800
         Width           =   375
      End
      Begin VB.TextBox txtbxLucroFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "0,0000"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxEr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "0,0000"
         Top             =   800
         Width           =   855
      End
      Begin VB.TextBox txtbxCM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "0,0000"
         Top             =   800
         Width           =   855
      End
      Begin VB.TextBox txtbxSujM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "0,0000"
         Top             =   800
         Width           =   855
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo Saída"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   81
         Top             =   0
         Width           =   870
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atual(%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4800
         TabIndex        =   80
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lucro(%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   79
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo Entr.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   78
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pç Venda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3000
         TabIndex        =   77
         Top             =   0
         Width           =   750
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo Ant.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   76
         Top             =   0
         Width           =   810
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2040
         TabIndex        =   75
         Top             =   555
         Width           =   525
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Icms Saída"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   74
         Top             =   555
         Width           =   780
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taxas(%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   73
         Top             =   555
         Width           =   720
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.Liq.:(%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5640
         TabIndex        =   72
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estoque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3000
         TabIndex        =   71
         Top             =   555
         Width           =   585
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Média"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   70
         Top             =   555
         Width           =   420
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suj. Méd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4920
         TabIndex        =   69
         Top             =   555
         Width           =   660
      End
   End
   Begin VB.PictureBox pcOperacao 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00B1DDB0&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   13890
      TabIndex        =   53
      Top             =   1215
      Width           =   13920
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   13890
      TabIndex        =   0
      Top             =   0
      Width           =   13920
      Begin VB.TextBox txtbxFrete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtbxEntrada_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxEmissao_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxValor_Not 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   9120
         MaxLength       =   12
         TabIndex        =   15
         Tag             =   "!"
         Text            =   "0,00"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxBC_Not 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxICMS_Not 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxIPI_Not 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8160
         MaxLength       =   12
         TabIndex        =   13
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxDepesas_Not 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         MaxLength       =   12
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cmbxTipo_Op 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmEntradas.frx":1CFE
         Left            =   120
         List            =   "frmEntradas.frx":1D00
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtbxNOme_Ter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtbxSerie_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   4
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxNum_Not 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8160
         MaxLength       =   12
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtbxValorProd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9120
         MaxLength       =   12
         TabIndex        =   29
         Tag             =   "!"
         Text            =   "0,00"
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         Picture         =   "frmEntradas.frx":1D02
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Frete:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         TabIndex        =   24
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblUF_Ter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "--"
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
         Left            =   3585
         TabIndex        =   52
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "E&ntrada:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   20
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Emissão: "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4080
         TabIndex        =   18
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor Contábil:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9120
         TabIndex        =   14
         Tag             =   "!"
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "B.Cálc.:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6240
         TabIndex        =   8
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&ICMS:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7200
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "I&PI:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8160
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Despesas:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6240
         TabIndex        =   22
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Operação:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornecedor:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Série:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº: &Doc.:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4080
         TabIndex        =   4
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desco&nto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8160
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Prod:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9120
         TabIndex        =   28
         Tag             =   "!"
         Top             =   600
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean 'para saber se o frm já foi aberto
Dim iTam As Integer 'tam do form
Dim sTableCab As String
Dim sTableItens As String
Public operacao As String

Private Function F_ValorContabilNota()
    F_ValorContabilNota = CCur(txtbxValorProd.text) + F_Despesas
End Function
Private Function F_Despesas()
    F_Despesas = CCur(txtbxDepesas_Not.text) + CCur(txtbxIPI_Not.text) + CCur(txtbxFrete.text) - CCur(txtbxDesconto.text)
End Function
Private Function F_GastosItem()
    If CCur(F_Qtd) > 0 Then
        F_GastosItem = (CCur(txtbxIpi.text) / F_Qtd) + CCur(txtbxFrete.text) - CCur(txtbxDesconto.text)
    End If
End Function
Private Function F_ValorProdutos()
    F_ValorProdutos = CCur(txtbxValorProd.text)
End Function

Private Sub cmbxBusca_Click()
    lblExpr.caption = "&" & cmbxBusca.text
    
    If cmbxBusca.ListIndex = 0 Then
        lblExpr.tag = "B"
    Else
        lblExpr.tag = Left(cmbxBusca.text, 1)
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxEmb_LostFocus()
    P_ValorLostFocus cmbxEmb, "###,##0"
    confEmb
    txtbxQtd.SetFocus
End Sub
Sub confEmb()
    Dim strSQL As String
    Dim iCount As Integer
    Dim ok As Boolean
    On Error GoTo trata_Erro
    ok = False
    If txtbxCodBarra.tag = "" Then
        Exit Sub
    End If
    If cmbxEmb.text = "" Then
        cmbxEmb.text = "1"
        Exit Sub
    End If
    
    If cmbxEmb.text = "0" Then
        cmbxEmb.text = "1"
        Exit Sub
    End If
    If cmbxEmb.text = "1" Then
        ok = True
    Else
        For iCount = 0 To cmbxEmb.ListCount - 1
            If cmbxEmb.text = cmbxEmb.List(iCount) Then
                ok = True
                Exit For
            End If
        Next
    End If

    If Not ok Then
        strSQL = "A embalagem " & cmbxEmb.text & " não foi localizada para o produto!"
        strSQL = strSQL & " Deseja gravá-la agora para o produto?"
        iCount = MsgBox(strSQL, vbYesNoCancel + vbQuestion)
        If iCount = vbCancel Then
            cmbxEmb.SetFocus
            Exit Sub
        ElseIf iCount = vbYes Then
            cmbxEmb.tag = cmbxEmb.text
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & txtbxCodBarra.tag & "',"
            strSQL = strSQL & cmbxEmb.text & ")"
            Conn.Execute strSQL
            P_PreencheEmbalagens txtbxCodBarra.tag, cmbxEmb, True
            cmbxEmb.text = cmbxEmb.tag
            
        End If
        
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Private Sub cmbxTipo_Op_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Sub P_Grava()
    Dim sTipo As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    sTipo = UCase(Left(lblTipo.caption, 1))
    If sTipo = "S" Then
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.text & _
                "' and serie_not ='" & txtbxSerie_Not.text & "' " _
                & " and tipo_not= 'S' and cod_Cli = '" & strCOD_CLI & "'") <> "" Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.text & _
                "' and serie_not ='" & txtbxSerie_Not.text & "' " _
                & " and cod_ter ='" & txtbxNOme_Ter.tag & "' " _
                & " and tipo_not= 'E' and cod_Cli = '" & strCOD_CLI & "'") <> "" Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "insert into tab_notasfiscais(Cod_Cli,Cod_Ter,Tipo_Not,Num_Not,Serie_Not, "
    strSQL = strSQL & "valor_not,emissao_not,bc_not,bcs_not,icms_Not,ipi_not,despesas_not,"
    strSQL = strSQL & "operacao_not,entrada_not,tipodoc_not,codval_not)Values("
    strSQL = strSQL & "'" & strCOD_CLI & "', "
    strSQL = strSQL & "'" & txtbxNOme_Ter.tag & "',"
    strSQL = strSQL & "'" & sTipo & "',"
    strSQL = strSQL & "'" & txtbxNum_Not & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.text & "',"
    strSQL = strSQL & F_GravaDec(txtbxValor_Not.text) & ","
    strSQL = strSQL & F_GravaData(txtbxEmissao_Not.text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxBC_Not.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBCS_Not.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxICMS_Not.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIPI_Not.text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxDepesas_Not.text) & ","

    strSQL = strSQL & "'" & Left(cmbxTipo_Op.text, 5) & "', "
    strSQL = strSQL & F_GravaData(txtbxEntrada_Not.text) & ","
    strSQL = strSQL & "'" & Left(cmbxtipo_Not.text, 1) & "', "
    strSQL = strSQL & "'" & Left(cmbxCodVal_Not.text, 1) & "') "
    Conn.Execute strSQL
    MsgBox "Nota incluída com sucesso!", vbInformation
    P_Clear
    txtbxNOme_Ter.SetFocus
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Clear()
    txtbxNOme_Ter.text = ""
    txtbxNOme_Ter.tag = ""
    txtbxNum_Not.text = ""
    txtbxSerie_Not.text = ""
    txtbxEmissao_Not.text = ""
    cmbxTipo_Op.ListIndex = -1
    txtbxBC_Not.text = "0,00"
    txtbxBCS_Not.text = "0,00"
    txtbxICMS_Not.text = "0,00"
    txtbxIPI_Not.text = "0,00"
    txtbxDepesas_Not.text = "0,00"
    txtbxEntrada_Not.text = ""
    cmbxCodVal_Not.ListIndex = -1
    cmbxtipo_Not.ListIndex = -1
    txtbxValor_Not.text = "0,00"
End Sub

Private Sub cmndbtnAbreVenda_Click()
    If Me.tag = "" Then
        abreNota
        HabCab False
    Else
        AlteraNota
        HabCab False
        HabItens True
    End If
End Sub
Sub abreNota()
    Dim strSQL As String
    Dim sSeq As String
    If txtbxNOme_Ter.tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtbxNOme_Ter.SetFocus
        Exit Sub
    End If
    If F_ValorContabilNota <> CCur(txtbxValor_Not.text) Then
        MsgBox "OS valores não batem!", vbCritical
        Exit Sub
    End If
    If txtbxValor_Not.text = "0,00" Then
        On Error Resume Next
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        txtbxValor_Not.SetFocus
        Exit Sub
    End If
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtbxNOme_Ter.tag & "' " & _
        " and num_not='" & txtbxNum_Not.text & "' and serie_not ='" & _
        txtbxSerie_Not.text & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Abertura da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    'frBotoes.Visible = True
    sSeq = MaxKeySQL(sTableCab, "seq_not", "", "000000#")
    strSQL = "insert into " & sTableCab & "(seq_not,Cod_For,Num_not,Serie_not,desconto_not,valorprod_not,Valor_not,Bc_Not,"
    strSQL = strSQL & "frete_not,icms_not,ipi_not,despesas_not,operacao_not,emissao_not,"
    strSQL = strSQL & "entrada_not,cod_usu,status_not,data_not)Values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & txtbxNOme_Ter.tag & "', "
    strSQL = strSQL & "'" & txtbxNum_Not.text & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.text & "',"
    strSQL = strSQL & F_GravaDec(txtbxDesconto.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValorProd.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValor_Not.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBC_Not.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxFrete.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxICMS_Not.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIPI_Not.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxDepesas_Not.text) & ","
    strSQL = strSQL & "'" & Left(cmbxTipo_Op.text, 4) & "',"
    strSQL = strSQL & F_GravaData(txtbxEmissao_Not.text) & ","
    strSQL = strSQL & F_GravaData(txtbxEntrada_Not.text) & ","
    strSQL = strSQL & "'" & strCod_Usu & "','1', " 'em aberto
    strSQL = strSQL & F_GravaDataTempo(Now) & ") "
    Conn.Execute strSQL
    Me.tag = sSeq
    lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.text) + _
                                         CCur(txtbxFrete.text))
    bActive = False
    Form_Activate
    'cmndbtnAddProd.SetFocus
    HabItens True
    txtbxCodBarra.SetFocus
End Sub
Sub AlteraNota()
    Dim strSQL As String
    Dim sSeq As String
    If txtbxNOme_Ter.tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtbxNOme_Ter.SetFocus
        Exit Sub
    End If
    If F_ValorContabilNota <> CCur(txtbxValor_Not.text) Then
        MsgBox "OS valores não batem!", vbCritical
        Exit Sub
    End If
    If txtbxValor_Not.text = "0,00" Then
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        txtbxValor_Not.SetFocus
        Exit Sub
    End If
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtbxNOme_Ter.tag & "' " & _
        " and num_not='" & txtbxNum_Not.text & "' and serie_not ='" & _
        txtbxSerie_Not.text & "' and seq_not <>'" & Me.tag & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Alteração da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    'frBotoes.Visible = True
    
    
    strSQL = "update " & sTableCab & " set "
    strSQL = strSQL & "cod_for='" & txtbxNOme_Ter.tag & "', "
    strSQL = strSQL & "num_not='" & txtbxNum_Not.text & "',"
    strSQL = strSQL & "serie_not='" & txtbxSerie_Not.text & "',"
    strSQL = strSQL & "desconto_not=" & F_GravaDec(txtbxDesconto.text) & ","
    strSQL = strSQL & "valorprod_not=" & F_GravaDec(txtbxValorProd.text) & ","
    strSQL = strSQL & "valor_not=" & F_GravaDec(txtbxValor_Not.text) & ","
    strSQL = strSQL & "bc_not=" & F_GravaDec(txtbxBC_Not.text) & ","
    strSQL = strSQL & "frete_not=" & F_GravaDec(txtbxFrete.text) & ","
    strSQL = strSQL & "icms_not=" & F_GravaDec(txtbxICMS_Not.text) & ","
    strSQL = strSQL & "ipi_not=" & F_GravaDec(txtbxIPI_Not.text) & ","
    strSQL = strSQL & "despesas_not=" & F_GravaDec(txtbxDepesas_Not.text) & ","
    strSQL = strSQL & "operacao_not='" & Left(cmbxTipo_Op.text, 4) & "',"
    strSQL = strSQL & "emissao_not=" & F_GravaData(txtbxEmissao_Not.text) & ","
    strSQL = strSQL & "entrada_not=" & F_GravaData(txtbxEntrada_Not.text) & ","
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "', "
    strSQL = strSQL & "data_not=" & F_GravaDataTempo(Now)
    
    strSQL = strSQL & " where seq_not='" & Me.tag & "' "
    Conn.Execute strSQL
    
    lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.text) + _
                                         CCur(txtbxFrete.text))
    bActive = False
    Form_Activate
    RecalculaItens
    
End Sub

Private Sub cmndbtnAddProd_Click()
    Dim p As produto


    
    p = F_BuscaProdSaida(txtbxCodBarra.text)
    
    If p.cod_pro = "" Then
        MsgBox "Produto não encontado!", vbExclamation
        Exit Sub
    End If
    
    If F_BuscaUmCampo("cod_pro", sTableItens, "num_not='" & txtbxNum_Not.text & _
        "' and serie_not='" & txtbxSerie_Not.text & "' and cod_for='" & _
            txtbxNOme_Ter.tag & "' and cod_pro='" & p.cod_pro & "'") <> "" Then
        MsgBox "Este produto já se encontra lançado!", vbExclamation
        Exit Sub
    End If
    
    txtbxCodBarra.text = p.nome_pro
    txtbxCodBarra.tag = p.cod_pro
    lblCodPro.caption = p.cod_pro
    P_PreencheEmbalagens p.cod_pro, cmbxEmb, True
    cmbxEmb.ListIndex = 0
    'lblCodBarra.caption = p.codbarra_pro
    txtbxCusto.text = FormatVal(p.cr_pro)
    txtbxPreco.text = FormatVal(p.pv_pro)
    txtbxLucro.text = FormatVal(p.lucro_pro)
    txtbxEr.text = p.est.er_pro
    txtbxQtd.SetFocus
End Sub
Sub AddProd()
    Dim cFator As Currency 'para saber qual a relacao do produto pelo total
    If p.ipi > (CCur(txtbxIPI_Not.text) - CCur(lblIpi_Lan.caption)) Then
        MsgBox "Esse valor para IPI não está correto!", vbExclamation
        Exit Sub
    ElseIf Round((p.qtd_pro * p.preco) + CCur(lblValoLan.caption), 2) > CCur(txtbxValor_Not.text) Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    ElseIf ((p.qtd_pro * p.preco)) = 0 Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    End If
    
    
    cFator = (p.qtd_pro * p.preco) / (CCur(txtbxValor_Not.text) - CCur(lblGastos.caption) - CCur(txtbxIPI_Not.text))
    
    P_InsereProd p.cod_pro, p.preco, p.qtd_pro, p.ipi, cFator
    P_buscaitens
    soma
    lblMarca.caption = p.nome_mar
End Sub
Sub P_InsereProd(cod_pro As String, unit As Currency, qtd As Currency, _
                ipi As Currency, cFator As Currency)
    Dim strSQL As String
    Dim sSeq As String
    Dim cCusto As Currency
    Dim cGastos As Currency
    sSeq = F_BuscaUmCampo("Max(Seq_ite)", sTableItens, "Num_Not='" & txtbxNum_Not.text & "'" & _
        " and cod_for='" & txtbxNOme_Ter.tag & "'")
    If sSeq = "" Then
        sSeq = "1"
    Else
        sSeq = Int(sSeq) + 1
    End If
    
    cGastos = CCur(lblGastos.caption) * cFator
    cCusto = unit + (cGastos / qtd) + (ipi / qtd)
    
    strSQL = "insert into " & sTableItens & "(num_not,cod_for,serie_not,cod_pro,"
    strSQL = strSQL & "unit_ite,qtd_ite,ipi_ite,custo_ite,gastos_ite,seq_ite)values("
    strSQL = strSQL & "'" & txtbxNum_Not.text & "',"
    strSQL = strSQL & "'" & txtbxNOme_Ter.tag & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.text & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit) & ","
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(ipi) & ","
    
    strSQL = strSQL & F_GravaDec(cCusto) & ","
    strSQL = strSQL & F_GravaDec(cGastos) & ","
    strSQL = strSQL & sSeq & ") "
    Conn.Execute strSQL
    
End Sub
Sub P_AlteraProd(cod_pro As String, unit As Currency, qtd As Currency, _
                ipi As Currency, cFator As Currency)
    Dim strSQL As String
    Dim cCusto As Currency
    Dim cGastos As Currency
    
    cGastos = CCur(lblGastos.caption) * cFator
    cCusto = unit + (cGastos / qtd) + (ipi / qtd)
    
    strSQL = "update " & sTableItens & " set "
    
    
    strSQL = strSQL & "unit_ite=" & F_GravaDec(unit) & ","
    strSQL = strSQL & "qtd_ite=" & F_GravaDec(qtd) & ","
    strSQL = strSQL & "ipi_ite=" & F_GravaDec(ipi) & ","
    strSQL = strSQL & "gastos_ite=" & F_GravaDec(cGastos) & ","
    strSQL = strSQL & "custo_ite =" & F_GravaDec(cCusto) & " "
    strSQL = strSQL & " where  "
    strSQL = strSQL & "num_not='" & txtbxNum_Not.text & "' "
    strSQL = strSQL & "and cod_for='" & txtbxNOme_Ter.tag & "' "
    strSQL = strSQL & "and serie_not='" & txtbxSerie_Not.text & "' "
    strSQL = strSQL & "and cod_pro= '" & cod_pro & "' "
    
    
    Conn.Execute strSQL
    P_buscaitens
End Sub

Private Sub cmndbtnConfirmar_Click()
    
    If Me.tag = "" Then
        If txtbxNOme_Ter.tag = "" Then
            MsgBox "Não há um cliente selecionado!", vbExclamation
            Exit Sub
        End If
        P_Grava
                
    Else
        P_Altera
    End If
End Sub
Sub P_Altera()
    Dim iPos As Integer
    Dim sTipo As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    sTipo = UCase(Left(lblTipo.caption, 1))
    If sTipo = "S" Then
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.text & _
                "' and serie_not ='" & txtbxSerie_Not.text & "' " _
                & " and tipo_not= 'S' and cod_Cli = '" & _
                strCOD_CLI) <> "" And _
                txtbxNum_Not.text <> txtbxNum_Not.tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.text & _
                "' and serie_not ='" & txtbxSerie_Not.text & "' " _
                & " and cod_ter ='" & txtbxNOme_Ter.tag & "' " _
                & " and tipo_not= 'E' and cod_Cli = '" & strCOD_CLI & "'") <> "" _
                And txtbxNum_Not.text <> txtbxNum_Not.tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "update tab_NotasFiscais set "
    
    strSQL = strSQL & "num_not='" & txtbxNum_Not & "',"
    strSQL = strSQL & "serie_not='" & txtbxSerie_Not.text & "',"
    strSQL = strSQL & "valor_Not=" & F_GravaDec(txtbxValor_Not.text) & ","
    strSQL = strSQL & "emissao_not=" & F_GravaData(txtbxEmissao_Not.text) & ", "
    strSQL = strSQL & "bc_not=" & F_GravaDec(txtbxBC_Not.text) & ","
    strSQL = strSQL & "bcs_not=" & F_GravaDec(txtbxBCS_Not.text) & ","
    strSQL = strSQL & "icms_not=" & F_GravaDec(txtbxICMS_Not.text) & ","
    strSQL = strSQL & "ipi_Not=" & F_GravaDec(txtbxIPI_Not.text) & ", "
    strSQL = strSQL & "despesas_not=" & F_GravaDec(txtbxDepesas_Not.text) & ","

    strSQL = strSQL & "operacao_not='" & Left(cmbxTipo_Op.text, 5) & "', "
    strSQL = strSQL & "entrada_not=" & F_GravaData(txtbxEntrada_Not.text) & ","
    strSQL = strSQL & "tipodoc_not='" & Left(cmbxtipo_Not.text, 1) & "', "
    strSQL = strSQL & "codval_not='" & Left(cmbxCodVal_Not.text, 1) & "' "
    strSQL = strSQL & " where cod_cli='" & strCOD_CLI & "' "
    strSQL = strSQL & " and num_not ='" & txtbxNum_Not.tag & "' "
    strSQL = strSQL & " and tipo_not='" & sTipo & "' "
    strSQL = strSQL & " and cod_ter='" & Me.tag & "' "
    strSQL = strSQL & " and serie_not='" & txtbxSerie_Not.tag & "'"
    Conn.Execute strSQL
    MsgBox "Nota Alterada com sucesso!", vbInformation
    If frmEditaMov.Visible = True Then
        iPos = frmEditaMov.ado.AbsolutePosition
        frmEditaMov.ado.Requery
        frmEditaMov.ado.Move iPos
    End If
    Unload Me
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

 Sub Cancelar()
    If MsgBox("Deseja cancelar?", vbYesNo + vbCritical) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub editaItem(Optional opcao As Integer = 0)
        Dim cValor_Max As Currency
        Dim cIpi_Max As Currency
        Dim strSQL As String
        Dim cFator As Currency
        Dim preco As Currency
        Dim qtd_pro As Currency
        'On Error Resume Next
        If grid.row = -1 Then
            Exit Sub
        End If
        With ado
            txtbxCusto.text = FormatVal(!cf_pro)
            txtbxPreco.text = FormatVal(!pv_pro)
            txtbxLucro.text = FormatVal(!lucro_pro)
            txtbxCodBarra.text = !nome_pro
            txtbxCodBarra.tag = !cod_pro
            txtbxQtd.text = !qtd_ite
            cmbxEmb.text = !qtd_emb
            txtbxUnit.text = !unit_ite
            txtbxTotal.text = !total
            cmndbtnConfirma.tag = !seq_ite
            txtbxIpi.text = !ipi_ite
            txtbxSujerido.text = !sujerido_ite
            
            txtbxIcms.text = !icms_ite
            txtbxQtd_LostFocus
            txtbxPrecoFuturo_LostFocus
            txtbxIcms_LostFocus
            txtbxPrecoFuturo.text = !precofuturo_ite
            
        End With
         
         
         
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If MsgBox("Confirma exclusão do item " & ado!nome_pro & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from " & sTableItens & " where SEQ_NOT='" & Me.tag & "' "

        strSQL = strSQL & " and seq_ite='" & ado!seq_ite & "' "
        Conn.Execute strSQL
        
        If Banco.RecordsAffected > 0 Then
            MsgBox "Item excuído! ", vbInformation
            soma
        Else
            MsgBox "Nenhum Registro Foi Afetado!", vbInformation
        End If
        
        
    End If
    P_buscaitens
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnForn_Click()

End Sub

Private Sub cmndbtnConfirma_Click()
    
    If txtbxCodBarra.tag = "" Then
        MsgBox "Informe o Nome do Produto e tecle <ENTER>!", vbCritical
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    If F_Qtd = 0 Then
        MsgBox "Informe o tipo de embalagem e quantidade", vbExclamation
        cmbxEmb.SetFocus
        Exit Sub
    End If
    If cmndbtnConfirma.tag = "" Then
        insere
    Else
        altera
    End If
    
End Sub

Private Sub cmndbtnEdit_Click()
    HabCab True
    HabItens False
    cmndbtnAbreVenda.Visible = True
End Sub
Sub HabItens(opcao As Boolean)
    grid.Enabled = opcao
'    frPr.Enabled = opcao
End Sub
Sub HabCab(opcao As Boolean)
'    frCab.Enabled = opcao
    txtbxNOme_Ter.Enabled = opcao
    cmbxTipo_Op.Enabled = opcao
    txtbxNum_Not.Enabled = opcao
    txtbxSerie_Not.Enabled = opcao
    txtbxBC_Not.Enabled = opcao
    txtbxICMS_Not.Enabled = opcao
    txtbxIPI_Not.Enabled = opcao
    txtbxValor_Not.Enabled = opcao
    txtbxEmissao_Not.Enabled = opcao
    txtbxDesconto.Enabled = opcao
    txtbxFrete.Enabled = opcao
    txtbxDepesas_Not.Enabled = opcao
    txtbxValorProd.Enabled = opcao
    
End Sub
Private Sub cmndbtnLocaliza_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxNOme_Ter.text)
    If f.cod_For <> "" Then
        txtbxNOme_Ter.text = f.nome_for
        txtbxNOme_Ter.tag = f.cod_For
    End If
End Sub

Private Sub cmndbtnNovoProd_Click()
    Dim fr As New frmCadPro
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
    If IsLoad(frmEdiEntr.Name) Then
        frmEdiEntr.ado.Requery
    End If
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL  As String
    Dim cNovoPreco As Currency
    Dim cLucro As Currency

    sDif = CCur(lblValoLan.caption) - CCur(txtbxValorProd.text)
    If sDif < 0 Then
        sDif = sDif * -1
    End If
    If sDif > 0.4 Then
        MsgBox "Os valores não estão completamente lançados!", vbExclamation
        Exit Sub
    ElseIf lblIpi_Lan.caption <> txtbxIPI_Not.text Then
        MsgBox "O valor do IPI não está completamente lançado!", vbExclamation
        Exit Sub
    ElseIf ado.RecordCount = 0 Then
        MsgBox "Não é possível concluir uma nota sem itens!", vbExclamation
        Exit Sub
    End If
    
    
    If lblIpi_Lan.caption <> txtbxIPI_Not.text Then
        MsgBox "O valor do IPI não está completamente lançado!", vbExclamation
        Exit Sub
    ElseIf ado.RecordCount = 0 Then
        MsgBox "Não é possível concluir uma nota sem itens!", vbExclamation
        Exit Sub
    End If
    
'    If MsgBox("Confirma efetivação da Nota Fiscal? ", vbYesNo + vbQuestion) = vbYes Then
        With ado
            .MoveFirst
            strSQL = "update " & sTableCab & " set status_not = '2' "
            strSQL = strSQL & " where Seq_not='" & Me.tag & "' "
            
            Conn.Execute strSQL
            
        End With
        MsgBox "Nota Fiscal concluída!", vbInformation
        If IsLoad(frmEdiEntr.Name) Then
            frmEdiEntr.ado.Requery
        End If
        Unload Me
 '   End If
    
    
    
End Sub


Private Sub grid_DblClick()
'    If cmndbtnEditProd.Enabled And cmndbtnEditProd.Visible = True Then
    
        editaItem
        On Error Resume Next
        txtbxQtd.SetFocus
 '   End If
End Sub

Private Sub grid_GotFocus()
    grid.MarqueeStyle = 3
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        editaItem
    End If
    If KeyCode = vbKeyDelete Then
        cmndbtnExc_Prod_Click
    End If
    
End Sub

Private Sub grid_LostFocus()
    grid.MarqueeStyle = 0
    
End Sub

Private Sub adc_Reposition()
    With ado
        If .RecordCount > 0 Then
            'cmndbtnEditProd.Enabled = True
            'cmndbtnExc_Prod.Enabled = True
            If Not .EOF Then

            End If
        Else
            'cmndbtnEditProd.Enabled = False
            'cmndbtnExc_Prod.Enabled = False
'            cmndbtnForn.Visible = False
        End If
    End With
End Sub

Sub soma()
    Dim iPos As Integer
    Dim iCount As Integer
    Dim cValor As Currency
    Dim cIPI As Currency
    Dim cFrete As Currency
    Dim cDesp As Currency
    With ado
    Err.Clear
    On Error GoTo trata_Erro
        
        If .RecordCount = 0 Then
            lblIpi_Lan.caption = "0,00"
            lblValoLan.caption = "0,00"
            lblItens.caption = "0"
            
            
            lblTotalProd.caption = "0,0"
            
            
            lblGastos.tag = "0,00"
            lblRest.caption = "0,00"
            Exit Sub
        End If
        iPos = .AbsolutePosition
        
        .MoveFirst
        Do While Not .EOF
            cValor = cValor + !total
            cIPI = cIPI + !ipi_ite
            iCount = iCount + !qtd_ite
            cDesp = cDesp + !gastos_ite
            
            .MoveNext
        Loop
            lblIpi_Lan.caption = FormatVal(cIPI)
            lblValoLan.caption = FormatVal(cValor)
            lblTotalProd.caption = FormatVal(cValor)
            'lblItens.caption = .RecordCount
            lblItens.caption = Format(iCount)
            lblGastos.tag = cDesp
            lblRest.caption = F_Restante
            .MoveFirst
            .Move iPos
    End With
    If CCur(txtbxBC_Not.text) > 0 Then
        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.text) / CCur(txtbxBC_Not.text) * 100, 2)) & " %"
    End If
trata_Erro:
End Sub

Sub insere()
    Dim strSQL As String
    Dim sSeq As String
    strSQL = F_BuscaUmCampo("seq_ite", sTableItens, "cod_pro='" & txtbxCodBarra.tag & "'  and seq_not='" & Me.tag & "' ")
    
    
    If strSQL <> "" Then
'        MsgBox "O Item já se encontra lançado como Item: " & strSQL, vbCritical
'        txtbxCodBarra.SetFocus
'        Exit Sub
    End If
    sSeq = MaxKeySQL(sTableItens, "seq_ite", "seq_not='" & Me.tag & "'", "")
    strSQL = "Insert into " & sTableItens & "(seq_not,seq_ite,cod_pro,unit_ite,qtd_ite,qtd_emb,ipi_ite,"
    strSQL = strSQL & "gastos_ite,custo_ite,sujerido_ite,precofuturo_ite,lucrosuj_ite,"
    strSQL = strSQL & "lucroliq_ite,custosai_ite,icms_ite)values("
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & txtbxCodBarra.tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxUnit.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxQtd.text) & ","
    strSQL = strSQL & F_GravaDec(cmbxEmb.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIpi.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxGastos.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxNovoCusto.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxSujerido.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxPrecoFuturo.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxLucro.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxLucroLiq.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxCustoSaida.text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIcms.text) & ")"
    Conn.Execute strSQL
    P_buscaitens
    P_LimpaItem
    
End Sub
Sub RecalculaItens()
    
    With ado
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            editaItem
            altera 1 'não dá refresh
            .MoveNext
        Loop
        P_buscaitens
    End With
End Sub
Sub P_LimpaItem()
    On Error Resume Next
    txtbxCodBarra.text = ""
    txtbxCodBarra.tag = ""
    txtbxCodBarra.SetFocus
    cmndbtnConfirma.tag = ""
End Sub
Sub altera(Optional opcao As Integer = 0)
    Dim strSQL As String
    Dim sSeq As String
    Dim i As Integer
    If opcao = 0 Then
        strSQL = F_BuscaUmCampo("seq_ite", sTableItens, "cod_pro='" & txtbxCodBarra.tag & "' and seq_ite <> '" & cmndbtnConfirma.tag & "' and seq_not='" & Me.tag & "'")
    
    
        If MsgBox("Confirma alteração do Item?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    
        If strSQL <> "" Then
            MsgBox "O Item já se encontra lançado como Item: " & strSQL, vbCritical
            txtbxCodBarra.SetFocus
            Exit Sub
        End If
    End If
    
    strSQL = "update " & sTableItens & " set "
    strSQL = strSQL & "cod_pro='" & txtbxCodBarra.tag & "',"
    strSQL = strSQL & "unit_ite=" & F_GravaDec(txtbxUnit.text) & ","
    strSQL = strSQL & "qtd_ite=" & F_GravaDec(txtbxQtd.text) & ","
    strSQL = strSQL & "qtd_emb=" & F_GravaDec(cmbxEmb.text) & ","
    strSQL = strSQL & "ipi_ite=" & F_GravaDec(txtbxIpi.text) & ","
    strSQL = strSQL & "gastos_ite=" & F_GravaDec(txtbxGastos.text) & ","
    strSQL = strSQL & "custo_ite=" & F_GravaDec(txtbxNovoCusto.text) & ","
    strSQL = strSQL & "sujerido_ite=" & F_GravaDec(txtbxSujerido.text) & ","
    strSQL = strSQL & "precofuturo_ite=" & F_GravaDec(txtbxPrecoFuturo.text) & ","
    strSQL = strSQL & "lucrosuj_ite=" & F_GravaDec(txtbxLucro.text) & ","
    strSQL = strSQL & "lucroliq_ite=" & F_GravaDec(txtbxLucroLiq.text) & ","
    strSQL = strSQL & "custosai_ite=" & F_GravaDec(txtbxCustoSaida.text) & ","
    strSQL = strSQL & "icms_ite=" & F_GravaDec(txtbxIcms.text) & " "
    strSQL = strSQL & " where seq_not='" & Me.tag & "' "
    strSQL = strSQL & " and seq_ite='" & cmndbtnConfirma.tag & "' "
    
    Conn.Execute strSQL
    cmndbtnConfirma.tag = ""
    If opcao = 0 Then
        i = ado.AbsolutePosition
    
        P_buscaitens
        ado.Move i
    
        P_LimpaItem
    End If
End Sub
Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    
    If operacao = "F" Then
        sTableCab = "tab_cmp"
        sTableItens = "tab_itcmp"
        pcOperacao.BackColor = &H80000018
    Else
        sTableCab = "tab_compras"
        sTableItens = "tab_itcompras"
        pcOperacao.BackColor = &HB1DDB0
    End If
    
    
    
    
    
    If Me.tag = "" Then
        iTam = Me.height
        
        
        
        
        
        
        
        
        
    Else
        
        cmndbtnAbreVenda.Visible = False
        'cmndbtnCancelar.Visible = False
        'dbgrdProd.Visible = True
        'frPr.Visible = True
        pcInfo.Visible = True
        frOrd.Visible = True
        pcResu.Visible = True
      
        
        cmndbtnAbreVenda.Visible = False
        'cmndbtnCancelar.Visible = False
        'frCab.Enabled = False
        P_buscaitens
        lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.text) + _
                             CCur(txtbxFrete.text))
                             
                             
    End If
    
    
    
    bActive = True
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyI And Shift = 1 Then
        P_InfPro ado!cod_pro
    
    End If
    If Shift = 2 And KeyCode = vbKeyC Then
        'optnC.Value = True
        'txtbxCodBarra.SetFocus
    End If
    If Shift = 2 And KeyCode = vbKeyN Then
        optnN.Value = True
        txtbxCodBarra.SetFocus
    End If
    If Shift = 2 And KeyCode = vbKeyB Then
        optnBar.Value = True
        txtbxCodBarra.SetFocus
    End If
    If KeyCode = vbKeyF2 Then
        editaItem
    ElseIf KeyCode = vbKeyF3 Then
        cmndbtnExc_Prod_Click
    ElseIf KeyCode = vbKeyF4 Then
        Me.WindowState = vbMinimized
            frmEdiPro.Show 0
        
        frmEdiPro.SetFocus
    ElseIf KeyCode = vbKeyF5 Then
        Me.WindowState = vbMinimized
        Dim fr As New frmCadPro
        fr.Show 0
        
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    End If

End Sub

Sub P_buscaitens()
    Dim strSQL As String
    
        
    strSQL = "select distinct ti.*,nome_sec,nome_sub,(ti.unit_ite*ti.qtd_ite*ti.qtd_emb) as total,ti.gastos_ite,tp.*,"
    
    strSQL = strSQL & " tf.nome_for,tp.cr_pro, tp.pv_pro "
    strSQL = strSQL & " from " & sTableItens & " as ti inner join " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & TAB_FORNECEDORES & " as tf on tf.cod_for=tp.cod_fab "
    strSQL = strSQL & " left join tab_subsec as ts on tp.cod_sub=ts.cod_sub "
    strSQL = strSQL & " LEFT Join tab_secoes as tss on tss.cod_sec=ts.cod_sec "
    strSQL = strSQL & " where ti.seq_not ='" & Me.tag & "' "
    
    
    If optnCod.Value = True Then
        strSQL = strSQL & " order by ti.cod_pro"
    
    End If
    If optnDesc.Value = True Then
        strSQL = strSQL & " order by tp.nome_pro"
    
    End If
    If optnSeq.Value = True Then
        strSQL = strSQL & " order by ti.seq_ite"
    
    End If
    abreObjetoAdo ado, strSQL, grid
    soma
    If CCur(txtbxIpi.text) = 0 Then
        lblIPI.Enabled = False
        txtbxIpi.Enabled = False
    End If
    If CCur(txtbxICMS_Not.text) = 0 Then
        lblIcmsCap.Enabled = False
        txtbxIcms.Enabled = False
        txtbxValorICMS.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    P_PreencheOperacoes cmbxTipo_Op, "E"
    'cmbxBusca.ListIndex = 0
End Sub

Private Sub Form_Resize()
    PosicionaGrade grid, Me, pcProduto
    
    
End Sub

Private Sub Option2_Click()
    
End Sub

Private Sub optnBar_Click()
    lblExpr.caption = "C&od. Barra:"
    lblExpr.tag = "B"
End Sub

Private Sub optnC_Click()
        lblExpr.caption = "C&od. Prod.:"
        lblExpr.tag = "C"
End Sub



Private Sub optnN_Click()
    lblExpr.caption = "N&ome Prod.:"
    lblExpr.tag = "N"
End Sub




Private Sub optnCod_Click()
    P_buscaitens
End Sub

Private Sub optnDesc_Click()
    P_buscaitens
End Sub

Private Sub optnSeq_Click()
    P_buscaitens
End Sub

Private Sub txtbxBC_Not_GotFocus()
    P_SelecTexto txtbxBC_Not
End Sub

Private Sub txtbxBC_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBC_Not_LostFocus()
    P_ValorLostFocus txtbxBC_Not
    If CCur(txtbxBC_Not.text) > 0 Then
        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.text) / CCur(txtbxBC_Not.text), 2) * 100) & " %"
    End If
End Sub


Private Sub txtbxBCS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub
Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
    If lblExpr.tag = "C" Then
        txtbxCodBarra.text = Format(txtbxCodBarra.text, "0000000#")
    End If
    
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnAddProd_Click
    End If
End Sub

Private Sub txtbxDepesas_Not_GotFocus()
    P_SelecTexto txtbxDepesas_Not
End Sub

Private Sub txtbxDepesas_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDepesas_Not_LostFocus()
    P_ValorLostFocus txtbxDepesas_Not
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    P_ValorLostFocus txtbxDesconto
End Sub

Private Sub txtbxEmissao_Not_GotFocus()
    P_SelecTexto txtbxEmissao_Not
End Sub

Private Sub txtbxEmissao_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEmissao_Not_LostFocus()
    P_AjustaData txtbxEmissao_Not
End Sub




Private Sub txtbxEntrada_Not_GotFocus()
    P_SelecTexto txtbxEntrada_Not
End Sub

Private Sub txtbxEntrada_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntrada_Not_LostFocus()
    On Error GoTo trata_Erro
    P_AjustaData txtbxEntrada_Not
    If txtbxEntrada_Not.text = "" Then Exit Sub
    If CDate(txtbxEntrada_Not.text) < CDate(txtbxEmissao_Not.text) Then
        MsgBox "A data de Entrada não pode ser anterior a data de Emissão", vbExclamation
        txtbxEntrada_Not.SetFocus
    End If
trata_Erro:
    
End Sub

Private Sub txtbxFrete_GotFocus()
    P_SelecTexto txtbxFrete
End Sub

Private Sub txtbxFrete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFrete_LostFocus()
    P_ValorLostFocus txtbxFrete
End Sub

Private Sub txtbxIcms_GotFocus()
    P_SelecTexto txtbxIcms
End Sub

Private Sub txtbxIcms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxIcms_LostFocus()
    P_ValorLostFocus txtbxIcms
    txtbxICMSSaida.text = F_IcmsProd(txtbxCodBarra.tag, 3)  'percentual
    P_CalcItem
End Sub

Private Sub txtbxICMS_Not_GotFocus()
    P_SelecTexto txtbxICMS_Not
End Sub

Private Sub txtbxICMS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxICMS_Not_LostFocus()
    P_ValorLostFocus txtbxICMS_Not
    If CCur(txtbxBC_Not.text) > 0 Then
        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.text) / CCur(txtbxBC_Not.text) * 100, 2)) & " %"
    End If
End Sub

Sub P_CalcItem()
    Dim strSQL As String
    Dim cCusto As Currency
    Dim dInd As Double
    Dim cSuj As Currency
    Dim cLucro As Currency
    On Error GoTo trata_Erro
    txtbxTaxas.text = F_TaxasEntradas
    txtbxValorICMS.text = Round(CCur(txtbxIcms.text) * (txtbxTotal.text) / 100, 4)
    If F_Qtd > 0 Then
        dInd = Round(CDbl(txtbxTotal.text) / CDbl(txtbxValor_Not.text) / F_Qtd, 8)
    End If
    
    cLucro = txtbxLucro.text
    cCusto = CCur(txtbxUnit.text) + (dInd * F_Despesas)
    txtbxNovoCusto.text = FormatVal(cCusto)
    cSuj = cCusto + (cCusto * cLucro / 100)
    txtbxSujerido.text = FormatVal(cSuj)
    txtbxPrecoFuturo.text = FormatVal(cSuj)
    P_calcLucroLiq
    txtbxGastos.text = FormatVal(dInd * F_Despesas, "###,###,##0.0000")
    
    
    'custo médio das ultimas duas
    cSuj = CCur(txtbxCM.text) + (CCur(txtbxCM.text) * cLucro / 100)
    txtbxSujM.text = FormatVal(cSuj)
    
    Exit Sub
trata_Erro:
End Sub
Sub P_calcLucroLiq()
    Dim cLucro As Currency
    Dim cCusto As Currency
    Dim cIcms As Currency
    On Error GoTo trata_Erro
    If CCur(txtbxSujerido.text) > 0 Then
    On Error GoTo trata_Erro
        'cLucro = CCur(txtbxLucro.Text) * CCur(txtbxPrecoFuturo.Text) / CCur(txtbxSujerido.Text)
        cLucro = (CCur(txtbxPrecoFuturo.text) / CCur(txtbxNovoCusto.text)) - 1
        cLucro = cLucro * 100
    End If
    txtbxLucroLiq.text = cLucro
    'cLucro = cLucro - (cLucro * CCur(txtbxTaxas.Text) / 100)
    cLucro = cLucro - CCur(txtbxTaxas.text)
    txtbxLucroFinal.text = cLucro
    
    If F_Qtd > 0 Then
        If txtbxICMSSaida.text = "" Then txtbxICMSSaida.text = "0,00"
    
        If CCur(txtbxICMSSaida.text) = 0 Then
            cIcms = (CCur(txtbxICMSSaida.text) * CCur(txtbxPrecoFuturo.text) / 100) - (CCur(txtbxValorICMS.text) / F_Qtd)
        Else
            cIcms = (CCur(txtbxICMSSaida.text) * CCur(txtbxPrecoFuturo.text) / 100) - (CCur(txtbxValorICMS.text) / F_Qtd)
        End If
    End If
    'If cIcms > 0 Then
        cCusto = CCur(txtbxNovoCusto.text) + cIcms
    'End If
    txtbxCustoSaida.text = cCusto
    
    txtbxCM.text = FormatVal((cCusto + CCur(txtbxCusto.text)) / 2)
    
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub
Sub P_Unit()
    If F_Qtd <> 0 Then
        txtbxUnit.text = FormatVal(CCur(txtbxTotal.text) / F_Qtd, "###,###,##0.0000")
    End If
End Sub
Private Sub txtbxIpi_GotFocus()
    If CCur(txtbxIpi.text) = 0 Then
        txtbxIpi.text = F_RestIpi
    End If
    P_SelecTexto txtbxIpi
End Sub

Function F_Restante() As Currency
    F_Restante = F_ValorProdutos - CCur(lblValoLan.caption)

End Function
Function F_RestIpi() As Currency
    F_RestIpi = CCur(txtbxIPI_Not.text) - CCur(lblIpi_Lan.caption)
End Function
Private Sub txtbxIpi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxIpi_LostFocus()
    
    P_ValorLostFocus txtbxIpi
    If CCur(txtbxIpi.text) > F_RestIpi Then
        MsgBox "O Ipi Está maior do que o restante", vbCritical
        txtbxIpi.text = F_RestIpi
    End If
    P_CalcItem
End Sub

Private Sub txtbxIPI_Not_GotFocus()
    P_SelecTexto txtbxIPI_Not
End Sub

Private Sub txtbxIPI_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxIPI_Not_LostFocus()
    P_ValorLostFocus txtbxIPI_Not
End Sub

Private Sub txtbxNOme_Ter_GotFocus()
    P_SelecTexto txtbxNOme_Ter
End Sub

Private Sub txtbxNOme_Ter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxNum_Not_GotFocus()
    P_SelecTexto txtbxNum_Not
End Sub

Private Sub txtbxNum_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxNum_Not_LostFocus()
     txtbxNum_Not.text = Format(F_SoNumero(txtbxNum_Not.text), "0#######")
End Sub


Private Sub txtbxPrecoFuturo_GotFocus()
    P_SelecTexto txtbxPrecoFuturo
End Sub

Private Sub txtbxPrecoFuturo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPrecoFuturo_LostFocus()
    P_ValorLostFocus txtbxPrecoFuturo
    P_calcLucroLiq
End Sub

Function F_Qtd() As Currency
    If cmbxEmb.text <> "" Then
        F_Qtd = CCur(cmbxEmb.text) * CCur(txtbxQtd.text)
    
        
    End If
End Function
Private Sub txtbxQtd_GotFocus()
        P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###,##0.000"
    If txtbxIPI_Not.text <> "0,00" Then
        txtbxIpi.Enabled = True
        lblIPI.Enabled = True
    Else
        txtbxIpi.Enabled = False
        lblIPI.Enabled = False
    End If
    
    P_Unit
    P_CalcItem
End Sub

Private Sub txtbxSerie_Not_GotFocus()
    P_SelecTexto txtbxSerie_Not
End Sub

Private Sub txtbxSerie_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxSerie_Not_LostFocus()
    P_AjustaTexto txtbxSerie_Not
End Sub

Private Sub txtbxTotal_GotFocus()
    P_SelecTexto txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    P_ValorLostFocus txtbxTotal
    P_Unit
    P_CalcItem
End Sub

Private Sub txtbxValor_Not_GotFocus()
    P_SelecTexto txtbxValor_Not
End Sub

Private Sub txtbxValor_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Not_LostFocus()
    P_ValorLostFocus txtbxValor_Not
End Sub

Private Sub txtbxValorProd_GotFocus()
    P_SelecTexto txtbxValorProd
End Sub

Private Sub txtbxValorProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorProd_LostFocus()
    P_ValorLostFocus txtbxValorProd
End Sub
