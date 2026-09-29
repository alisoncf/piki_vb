VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCadPro 
   Caption         =   "Cadastro de Produtos:"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13350
   Icon            =   "frmCadPro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7290
      Left            =   9375
      ScaleHeight     =   7290
      ScaleWidth      =   3975
      TabIndex        =   122
      Top             =   0
      Width           =   3975
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gradeAlteracoesPreco 
         Height          =   3855
         Left            =   120
         TabIndex        =   160
         Top             =   3240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   6800
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.PictureBox pcCR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   1800
         Picture         =   "frmCadPro.frx":0ECA
         ScaleHeight     =   765
         ScaleWidth      =   1725
         TabIndex        =   124
         Top             =   120
         Width           =   1725
         Begin VB.TextBox txtbxCusto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            TabIndex        =   125
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.PictureBox pcPV 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   1800
         Picture         =   "frmCadPro.frx":5460
         ScaleHeight     =   765
         ScaleWidth      =   1725
         TabIndex        =   127
         Top             =   960
         Width           =   1725
         Begin VB.TextBox txtbxPreco 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            TabIndex        =   128
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.PictureBox pcEr 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   1800
         Picture         =   "frmCadPro.frx":99F6
         ScaleHeight     =   765
         ScaleWidth      =   1725
         TabIndex        =   130
         Top             =   2040
         Width           =   1725
         Begin VB.TextBox txtbxEstoque 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            TabIndex        =   131
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label44 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Alterações de Preço:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1080
         TabIndex        =   161
         Top             =   2880
         Width           =   1725
      End
      Begin VB.Label lblMlPraticada 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "ML Praticada:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FCB683&
         Height          =   210
         Left            =   360
         TabIndex        =   159
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblestCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Em Estoque:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   129
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lblPrecoCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "&Preço de Venda:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   126
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblCusCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "C&usto Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   123
         Top             =   360
         Width           =   1035
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   6975
      Left            =   120
      TabIndex        =   101
      Top             =   120
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Dados Básicos"
      TabPicture(0)   =   "frmCadPro.frx":DF8C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SPED/NF-E"
      TabPicture(1)   =   "frmCadPro.frx":DFA8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pcCST"
      Tab(1).Control(1)=   "pcComb"
      Tab(1).Control(2)=   "chkComb"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Similares"
      TabPicture(2)   =   "frmCadPro.frx":DFC4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pcSim"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Atacado"
      TabPicture(3)   =   "frmCadPro.frx":DFE0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture8"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Veículos"
      TabPicture(4)   =   "frmCadPro.frx":DFFC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "pcVeiculos"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Conversão"
      TabPicture(5)   =   "frmCadPro.frx":E018
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture3"
      Tab(5).ControlCount=   1
      Begin VB.CheckBox chkComb 
         Caption         =   "Combustível"
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
         Left            =   -74880
         TabIndex        =   92
         Top             =   5880
         Width           =   1575
      End
      Begin VB.PictureBox pcComb 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -74880
         ScaleHeight     =   705
         ScaleWidth      =   8865
         TabIndex        =   93
         Top             =   6120
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox txtDescAnp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   97
            Tag             =   "00"
            Top             =   360
            Width           =   6255
         End
         Begin VB.TextBox txtcProdANP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   95
            Tag             =   "00"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição ANP:"
            Height          =   195
            Left            =   1800
            TabIndex        =   96
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código ANP:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5775
         Left            =   -74880
         ScaleHeight     =   5745
         ScaleWidth      =   8865
         TabIndex        =   147
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtReferenciaConversao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5520
            TabIndex        =   155
            Tag             =   "49"
            Top             =   5160
            Width           =   2055
         End
         Begin VB.PictureBox pcBuscaMarcaConversao 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4920
            Picture         =   "frmCadPro.frx":E034
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   153
            Top             =   5160
            Width           =   255
         End
         Begin VB.TextBox txtMarcaConversao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   152
            Tag             =   "49"
            Top             =   5160
            Width           =   4815
         End
         Begin VB.CommandButton cmdAddConversao 
            Caption         =   "Adicionar"
            Height          =   375
            Left            =   7680
            TabIndex        =   156
            Top             =   5040
            Width           =   1095
         End
         Begin MSDataGridLib.DataGrid gridConversao 
            Bindings        =   "frmCadPro.frx":E3DF
            Height          =   3975
            Left            =   120
            TabIndex        =   149
            Top             =   360
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
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
                  ColumnWidth     =   5999,812
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2204,788
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdExcluirConversao 
            Caption         =   "Excluir"
            Height          =   255
            Left            =   7680
            TabIndex        =   150
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Tabela de Conversão do Produto"
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
            Height          =   315
            Left            =   0
            TabIndex        =   148
            Top             =   0
            Width           =   8940
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca:"
            Height          =   195
            Left            =   120
            TabIndex        =   151
            Top             =   4920
            Width           =   495
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            Height          =   195
            Left            =   5520
            TabIndex        =   154
            Top             =   4920
            Width           =   825
         End
      End
      Begin VB.PictureBox pcVeiculos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   -74880
         ScaleHeight     =   6225
         ScaleWidth      =   8865
         TabIndex        =   136
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Adicionar"
            Height          =   375
            Left            =   120
            TabIndex        =   145
            Top             =   5520
            Width           =   1095
         End
         Begin VB.TextBox txtAplicacao 
            Appearance      =   0  'Flat
            Height          =   1005
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   144
            Tag             =   "49"
            Top             =   4440
            Width           =   8655
         End
         Begin VB.TextBox txtVeiculo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   141
            Tag             =   "49"
            Top             =   3840
            Width           =   8415
         End
         Begin VB.PictureBox pcVeic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8520
            Picture         =   "frmCadPro.frx":E3F1
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   142
            Top             =   3840
            Width           =   255
         End
         Begin MSDataGridLib.DataGrid gridVeiculos 
            Bindings        =   "frmCadPro.frx":E79C
            Height          =   3135
            Left            =   120
            TabIndex        =   138
            Top             =   360
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   5530
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
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
            ColumnCount     =   3
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "veiculo"
               Caption         =   "Veículo"
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
               DataField       =   "aplicacao"
               Caption         =   "Ano/Modelo (Aplicação)"
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
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2564,788
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   4484,977
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "Excluir"
            Height          =   255
            Left            =   7680
            TabIndex        =   139
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label lblVeiculoFabricante 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
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
            Left            =   945
            TabIndex        =   146
            Top             =   3600
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano/Modelo (Aplicação)"
            Height          =   195
            Left            =   120
            TabIndex        =   143
            Top             =   4200
            Width           =   1725
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Veículo:"
            Height          =   195
            Left            =   120
            TabIndex        =   140
            Top             =   3600
            Width           =   600
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Este produto é aplicável aos seguintes veículos"
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
            Height          =   315
            Left            =   0
            TabIndex        =   137
            Top             =   0
            Width           =   8940
         End
      End
      Begin VB.PictureBox pcSim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   -74880
         ScaleHeight     =   6345
         ScaleWidth      =   8865
         TabIndex        =   110
         Top             =   540
         Width           =   8895
         Begin MSDataGridLib.DataGrid gridSim 
            Height          =   5295
            Left            =   0
            TabIndex        =   133
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   9340
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdAddSimilar 
            Caption         =   "Adicionar"
            Height          =   495
            Left            =   120
            TabIndex        =   112
            Top             =   5640
            Width           =   1575
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Similaridades com Este Produto"
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
            TabIndex        =   115
            Top             =   0
            Width           =   2685
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   -74880
         ScaleHeight     =   6225
         ScaleWidth      =   8865
         TabIndex        =   103
         Top             =   540
         Width           =   8895
         Begin VB.TextBox txtQtde 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   106
            Text            =   "1"
            Top             =   5640
            Width           =   615
         End
         Begin VB.CommandButton cmdAddAtacado 
            Caption         =   "Adicionar"
            Height          =   375
            Left            =   5640
            TabIndex        =   114
            Top             =   5520
            Width           =   855
         End
         Begin VB.CommandButton cmdExclAtacado 
            Caption         =   "E&xcluir"
            Height          =   375
            Left            =   6600
            TabIndex        =   116
            Top             =   5520
            Width           =   855
         End
         Begin VB.TextBox txtPreco 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   109
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   5640
            Width           =   1095
         End
         Begin VB.TextBox txtBarraAtacado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   113
            Tag             =   "!"
            Top             =   5640
            Width           =   3375
         End
         Begin MSDataGridLib.DataGrid dtgrd 
            Bindings        =   "frmCadPro.frx":E7AE
            Height          =   4935
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   8705
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
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
            Caption         =   "Quantidades e preços"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   2369,764
               EndProperty
            EndProperty
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Vendas por Atacado"
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
            Height          =   315
            Left            =   0
            TabIndex        =   108
            Top             =   0
            Width           =   8820
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   5400
            Width           =   390
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Preço:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   107
            Top             =   5400
            Width           =   465
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Barra da Embalagem (Exclusiva)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   111
            Top             =   5400
            Width           =   2280
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   120
         ScaleHeight     =   6375
         ScaleWidth      =   8175
         TabIndex        =   43
         Top             =   360
         Width           =   8175
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   166
            Top             =   6000
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtMultiplo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   4
            Text            =   "1"
            ToolTipText     =   "Só vende qtd mútipla desse valor"
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox txtQconv 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   3360
            TabIndex        =   157
            Text            =   "1"
            Top             =   4080
            Width           =   855
         End
         Begin VB.CheckBox chkComercializavel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Comercializável"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4200
            TabIndex        =   134
            ToolTipText     =   "Desmarcando este produto ele não poderá ser comercializado diretamente."
            Top             =   6000
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtQtdEmb 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Text            =   "1"
            Top             =   4080
            Width           =   855
         End
         Begin VB.CheckBox chkControlaEstoque 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Controla Estoque"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2400
            TabIndex        =   132
            Top             =   6000
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox txtCEST 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4920
            TabIndex        =   40
            Top             =   4680
            Width           =   3135
         End
         Begin VB.PictureBox pcMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7800
            Picture         =   "frmCadPro.frx":E7C0
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   17
            Top             =   2880
            Width           =   255
         End
         Begin VB.TextBox txtMarca 
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
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   2880
            Width           =   7695
         End
         Begin VB.TextBox txtSubCat 
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
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   7695
         End
         Begin VB.PictureBox pcSubCat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7800
            Picture         =   "frmCadPro.frx":EB6B
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   14
            Top             =   2280
            Width           =   255
         End
         Begin VB.PictureBox pcCat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7800
            Picture         =   "frmCadPro.frx":EF16
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   11
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox txtCat 
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
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   7695
         End
         Begin VB.TextBox txtDetalhe 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   5280
            Width           =   7935
         End
         Begin VB.PictureBox pcForn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6720
            Picture         =   "frmCadPro.frx":F2C1
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   20
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txtFornecedor 
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
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   3480
            Width           =   6615
         End
         Begin VB.TextBox txtBarra 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            Height          =   285
            Left            =   2400
            TabIndex        =   38
            Top             =   4680
            Width           =   2415
         End
         Begin VB.TextBox txtSubSec 
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
            Height          =   285
            Left            =   2040
            TabIndex        =   6
            Top             =   1080
            Width           =   5775
         End
         Begin VB.PictureBox pcSubSec 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7800
            Picture         =   "frmCadPro.frx":F66C
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   7
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox txtbxEstMin 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   4320
            TabIndex        =   30
            Text            =   "1"
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox txtbxNome_Pro 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   7935
         End
         Begin VB.TextBox txtbxUnidade 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7080
            TabIndex        =   22
            Text            =   "UN"
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox txtbxCodFab 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtbxLocal 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   4080
            Width           =   1215
         End
         Begin VB.TextBox txtbxMargem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5280
            TabIndex        =   32
            Text            =   "0"
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox txtNCM 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6240
            TabIndex        =   34
            Top             =   4080
            Width           =   1815
         End
         Begin VB.CheckBox chkProd 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Balança"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5880
            TabIndex        =   45
            Top             =   6000
            Width           =   975
         End
         Begin VB.CheckBox chkVis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Visível"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7080
            TabIndex        =   44
            Top             =   6000
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.PictureBox pcUnidade 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   7800
            Picture         =   "frmCadPro.frx":FA17
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   23
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txtEan 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   4680
            Width           =   2175
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Múltiplo:"
            Height          =   195
            Left            =   2400
            TabIndex        =   28
            Top             =   3840
            Width           =   585
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Q.Conv."
            Height          =   195
            Left            =   3360
            TabIndex        =   158
            Top             =   3840
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Embalagem."
            Height          =   195
            Left            =   1440
            TabIndex        =   26
            Top             =   3840
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&CEST"
            Height          =   195
            Left            =   4920
            TabIndex        =   39
            Top             =   4440
            Width           =   420
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Marca:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label SubCategoria 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subcategoria"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Categoria"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Outros detalhes do produto"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   5040
            Width           =   1920
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor/Fabricante:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   3240
            Width           =   1680
         End
         Begin VB.Label lblSecao 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Seção"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3360
            TabIndex        =   8
            Top             =   840
            Width           =   4695
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&EAN 2 (Embalagem):"
            Height          =   195
            Left            =   2400
            TabIndex        =   37
            Top             =   4440
            Width           =   1470
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subseção:"
            Height          =   195
            Left            =   2040
            TabIndex        =   5
            Top             =   840
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. &Mín."
            Height          =   195
            Left            =   4320
            TabIndex        =   29
            Top             =   3840
            Width           =   645
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Nome do Produto:"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Unidade:"
            Height          =   195
            Left            =   7080
            TabIndex        =   21
            Top             =   3240
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Referência.:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Localização:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   3840
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M.&Luc(%)"
            Height          =   195
            Left            =   5280
            TabIndex        =   31
            Top             =   3840
            Width           =   660
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&NCM:"
            Height          =   195
            Left            =   6240
            TabIndex        =   33
            Top             =   3840
            Width           =   405
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&EAN:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   4440
            Width           =   375
         End
      End
      Begin VB.PictureBox pcCST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   -74880
         ScaleHeight     =   5265
         ScaleWidth      =   8865
         TabIndex        =   46
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtPCofins 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8160
            TabIndex        =   61
            Tag             =   "17"
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtPPis 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7440
            TabIndex        =   59
            Tag             =   "17"
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtPredIcms 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6720
            TabIndex        =   57
            Tag             =   "17"
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtPIcms 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            TabIndex        =   55
            Tag             =   "17"
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton cmdAlterarEstoque 
            Caption         =   "Alterar"
            Height          =   375
            Left            =   6720
            TabIndex        =   84
            Top             =   4200
            Width           =   975
         End
         Begin VB.TextBox txtEmpresa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Left            =   4440
            TabIndex        =   52
            Top             =   480
            Width           =   1095
         End
         Begin VB.PictureBox pcEmpresa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5520
            Picture         =   "frmCadPro.frx":FDC2
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   53
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton cmdAddEstoque 
            Caption         =   "Adicionar"
            Height          =   495
            Left            =   6720
            TabIndex        =   83
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CommandButton cmdExcluirEstoque 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   7800
            TabIndex        =   85
            Top             =   4200
            Width           =   975
         End
         Begin VB.PictureBox pcCfopRevenda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4080
            Picture         =   "frmCadPro.frx":1016D
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   79
            Top             =   1680
            Width           =   255
         End
         Begin VB.PictureBox pcCfopConsumidor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            Picture         =   "frmCadPro.frx":10518
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   76
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox txtCFOP_Revenda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   78
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtCFOP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   75
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtCstRevenda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   66
            Top             =   1080
            Width           =   1815
         End
         Begin VB.PictureBox pcCstRevenda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4080
            Picture         =   "frmCadPro.frx":108C3
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   67
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox txtEanCopia 
            Appearance      =   0  'Flat
            BackColor       =   &H0073E1F4&
            Height          =   285
            Left            =   2280
            TabIndex        =   50
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtNCMCopia 
            Appearance      =   0  'Flat
            BackColor       =   &H0073E1F4&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   2055
         End
         Begin VB.PictureBox pcOrigem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4920
            Picture         =   "frmCadPro.frx":10C6E
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   91
            Top             =   4920
            Width           =   255
         End
         Begin VB.TextBox txtOrigem 
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
            Height          =   285
            Left            =   120
            TabIndex        =   90
            Top             =   4920
            Width           =   4815
         End
         Begin VB.TextBox txtGenero 
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
            Height          =   285
            Left            =   5280
            TabIndex        =   87
            Top             =   4920
            Width           =   3255
         End
         Begin VB.PictureBox pcGenero 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8520
            Picture         =   "frmCadPro.frx":11019
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   88
            Top             =   4920
            Width           =   255
         End
         Begin VB.PictureBox pcCstCofins 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8520
            Picture         =   "frmCadPro.frx":113C4
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   73
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox txtCst_Cofins 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6720
            TabIndex        =   72
            Top             =   1080
            Width           =   1815
         End
         Begin VB.PictureBox pcCstPis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6360
            Picture         =   "frmCadPro.frx":1176F
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   70
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox txtCst_Pis 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   69
            Top             =   1080
            Width           =   1935
         End
         Begin VB.PictureBox pcCstIpi 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6360
            Picture         =   "frmCadPro.frx":11B1A
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   82
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox txtCst_Ipi 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   81
            Top             =   1680
            Width           =   1935
         End
         Begin VB.PictureBox pcCstIcms 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            Picture         =   "frmCadPro.frx":11EC5
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   64
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox txtCst_Icms 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   1815
         End
         Begin MSDataGridLib.DataGrid gridestoque 
            Height          =   1815
            Left            =   120
            TabIndex        =   162
            Top             =   2280
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   3201
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
            Caption         =   "Estoque"
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "cod_emp"
               Caption         =   "Empr."
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
               DataField       =   "er"
               Caption         =   "ER."
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
               DataField       =   "ef"
               Caption         =   "EF"
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
               DataField       =   "reservado"
               Caption         =   "Res."
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
               DataField       =   "cfop_Revenda"
               Caption         =   "Cfop Rev."
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
               DataField       =   "Cfop_consumidor"
               Caption         =   "Cfop Cons."
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
               DataField       =   "cst_icms_revenda"
               Caption         =   "Cst Rev."
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
               DataField       =   "Cst_icms_consumidor"
               Caption         =   "Cst Cons."
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
               DataField       =   "cst_pis"
               Caption         =   "Pis"
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
               DataField       =   "cst_cofins"
               Caption         =   "Cofins"
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
            BeginProperty Column10 
               DataField       =   "picms"
               Caption         =   "P.Icms"
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
            BeginProperty Column11 
               DataField       =   "ppis"
               Caption         =   "PPis"
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
            BeginProperty Column12 
               DataField       =   "pcofins"
               Caption         =   "PCofins"
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
                  ColumnWidth     =   854,929
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   854,929
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   824,882
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   884,976
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   959,811
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   824,882
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   780,095
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   629,858
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   599,811
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   629,858
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column12 
               EndProperty
            EndProperty
         End
         Begin VB.Label lblReservado 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   1080
            TabIndex        =   165
            Top             =   4320
            Width           =   90
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Res. Geral:"
            Height          =   195
            Left            =   120
            TabIndex        =   164
            Top             =   4320
            Width           =   795
         End
         Begin VB.Label Label49 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "P.Cofins"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8160
            TabIndex        =   60
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label48 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "P.Pis"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7440
            TabIndex        =   58
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label47 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "P .Red."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6720
            TabIndex        =   56
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label46 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "P. ICMS"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6000
            TabIndex        =   54
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   195
            Left            =   4440
            TabIndex        =   51
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CFOP Revenda:"
            Height          =   195
            Left            =   2280
            TabIndex        =   77
            Top             =   1440
            Width           =   1170
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CFOP Consumidor Final:"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   1710
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst ICMS Revenda:"
            Height          =   195
            Left            =   2280
            TabIndex        =   65
            Top             =   840
            Width           =   1410
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&EAN:"
            Height          =   195
            Left            =   2280
            TabIndex        =   49
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&NCM:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblIcms 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Situação Trib."
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
            Left            =   6720
            TabIndex        =   121
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Origem:"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   4680
            Width           =   540
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gênero: (Apenas para produtos primários ex.: insumos)"
            Height          =   195
            Left            =   5280
            TabIndex        =   86
            Top             =   4680
            Width           =   3840
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Códigos de Situação tributária em operações de saída (vendas):"
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
            Left            =   120
            TabIndex        =   102
            Top             =   0
            Width           =   5265
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst COFINS:"
            Height          =   195
            Left            =   6720
            TabIndex        =   71
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst PIS:"
            Height          =   195
            Left            =   4440
            TabIndex        =   68
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst IPI:"
            Height          =   195
            Left            =   4440
            TabIndex        =   80
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst ICMS Consumidor Final:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   1950
         End
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   13350
      TabIndex        =   117
      Top             =   7290
      Width           =   13350
      Begin VB.CommandButton cmdSim 
         Caption         =   "Similares"
         Enabled         =   0   'False
         Height          =   735
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdImagens 
         Caption         =   "Imagens"
         Height          =   735
         Left            =   6360
         Picture         =   "frmCadPro.frx":12270
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdHist 
         Caption         =   "&Histórico"
         Height          =   735
         Left            =   3240
         Picture         =   "frmCadPro.frx":12364
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdVisivel 
         Caption         =   "Tornar Invisível"
         Height          =   735
         Left            =   1680
         Picture         =   "frmCadPro.frx":124AE
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnDelPro 
         Caption         =   "E&xcluir Produto"
         Height          =   735
         Left            =   120
         Picture         =   "frmCadPro.frx":125F8
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Gravar (F7)"
         Height          =   855
         Left            =   9480
         Picture         =   "frmCadPro.frx":12742
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar (ESC)"
         Height          =   855
         Left            =   11520
         Picture         =   "frmCadPro.frx":1288C
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnPro 
         Caption         =   "Pro&moção"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4800
         Picture         =   "frmCadPro.frx":12DDA
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCadPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoestoque As ADODB.Recordset
Attribute adoestoque.VB_VarHelpID = -1
Dim bActive As Boolean
Dim cPreco As String 'para carregar o preço atual no caso de mudança de preço
Dim adoAtac As ADODB.Recordset
Dim adoSim As ADODB.Recordset
Dim adoVei As ADODB.Recordset
Dim adoConversao As ADODB.Recordset
Dim visivel As String

Private Sub chkComb_Click()
    If chkComb.Value = 0 Then
        pcComb.Visible = False
    Else
        pcComb.Visible = True
        txtcProdANP.SetFocus
    End If
    
End Sub

Private Sub chkControlaEstoque_Click()
    If chkControlaEstoque.Value = 1 Then
        txtbxEstoque.Visible = True
    Else
        txtbxEstoque.Visible = False
    End If
End Sub

Private Sub chkProd_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Sub buscaEstoque()
    Set adoestoque = New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select * from estoque where cod_pro='" & Me.tag & "' order by cod_emp"
    abreObjetoAdo adoestoque, strSQL, gridestoque
    
    
    
End Sub
Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyF5 Then
        
    End If
    
End Sub

Private Sub chkVis_Click()
    If chkVis.Value = 1 Then
        visivel = "S"
    Else
        visivel = "N"
    End If
End Sub



Private Sub cmbxICMS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub




Private Sub cmbxSt_Click()
    If cmbxSt.ListIndex <> 0 Then
        cmbxICMS.ListIndex = 0
        cmbxICMS.Enabled = False
        'P_LocalizaOperacao "I05", cmbxICMS
    Else
        cmbxICMS.Enabled = True
    End If
End Sub

Private Sub cmbxSt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmdAdd_Click()
    If Me.tag = "" Then
        MsgBox "O produto precisa estar gravado para efetuar essa operação.", vbExclamation
        Exit Sub
    End If
    If txtVeiculo.tag = "" Then
        MsgBox "Informe qual veículo.", vbExclamation
        txtVeiculo.SetFocus
        Exit Sub
    End If
    
    If Veiculos.GravaProduto_Veiculo(Me.tag, txtVeiculo.tag, txtAplicacao.Text) Then
         buscaVei
    End If
    
End Sub
Sub buscaVei()
    Dim strSQL As String
    strSQL = Veiculos.SelectProduto_Veiculo(Me.tag)
    abreObjetoAdo adoVei, strSQL, gridVeiculos
End Sub
Sub buscaConversao()
    Dim strSQL As String
    strSQL = SelectProdutoConversao(Me.tag)
    abreObjetoAdo adoConversao, strSQL, gridConversao
End Sub

Private Sub cmdAddAtacado_Click()
    If Me.tag = "" Then
        MsgBox "Primeiro grave o produto!", vbExclamation
        Exit Sub
    End If
    If CCur(txtQtde.Text) <= 1 Then
        MsgBox "Quantidade inválida! Deve ser maior que 1.", vbExclamation
        txtQtde.SetFocus
        Exit Sub
    End If
'    If CCur(txtPreco.Text) <= 0 Then
'        MsgBox "Preço inválido!", vbExclamation
'        txtPreco.SetFocus
'        Exit Sub
'    End If
    If strAltpreco_usu <> "S" And strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
    If cmdAddAtacado.tag = "" Then
        GravaProdutoAtacado Me.tag, txtQtde.Text, txtPreco, txtBarraAtacado.Text
        BuscaProdutoAtacado
    Else
        AlteraProdutoAtacado cmdAddAtacado.tag, txtQtde.Text, txtPreco.Text, txtBarraAtacado.Text
        cmdAddAtacado.tag = ""
        cmdAddAtacado.caption = "Adicionar"
        BuscaProdutoAtacado
    End If
    
End Sub
Sub BuscaProdutoAtacado()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "select * from produto_atacado where codigo='" & Me.tag & "' order by qtd "
    abreObjetoAdo adoAtac, strSQL, dtgrd
End Sub

Private Sub cmdAddConversao_Click()
    Dim p As ProdutoConversao
    If Me.tag = "" Then
        MsgBox "É necessário que o produto esteja gravado!", vbExclamation
        Exit Sub
    End If
    If txtMarcaConversao.tag = "" Then
        MsgBox "Informe a marca!", vbExclamation
        txtMarcaConversao.SetFocus
        Exit Sub
    End If
    If Trim(txtReferenciaConversao.Text) = "" Then
        MsgBox "Informe a referência!", vbExclamation
        txtReferenciaConversao.SetFocus
        Exit Sub
    End If
    p = Produtos.BuscaProdutoConversao(Me.tag, txtMarcaConversao.tag)
    If p.id <> "" Then
        MsgBox "Já existe uma conversão para este produto e marca com a referência " & p.referencia, vbExclamation
        Exit Sub
    End If
    
    Produtos.GravaProdutoConversao Me.tag, txtMarcaConversao.tag, txtReferenciaConversao.Text
    buscaConversao
    
    
End Sub
Function verifica_tag()
    



    verifica_tag = True
End Function
Private Sub cmdAddEstoque_Click()
    Dim emp As sped_empresa.sped_empresa
    If Me.tag = "" Then
        MsgBox "Primeiro grave o produto para adicionar estoque!", vbExclamation
        Exit Sub
    End If
    If strAltprod_usu <> "S" Then
            MsgBox "Você não tem permissão pra essa operação", vbExclamation
            Exit Sub
    End If
    
    If txtEmpresa.tag = "" Then
        MsgBox "Informe a empresa", vbExclamation
        txtEmpresa.SetFocus
        Exit Sub
    End If
    
    If Not verifica_tag Then
        Exit Sub
    End If
    
    
    
    If cmdAlterarEstoque.tag = "" Then
    
    
        If F_BuscaUmCampo("cod_emp", "estoque", "cod_pro='" & Me.tag & "' and cod_emp='" & emp.codigo & "'") <> "" Then
            MsgBox "Já existe estoque desta empresa para este produto!", vbExclamation
            Exit Sub
        End If
        
        If AdicionaEstoque(txtEmpresa.tag, Me.tag, "0", "0", txtCFOP_Revenda.tag, txtCFOP.tag, txtCstRevenda.tag, _
            txtCst_Icms.tag, txtCst_Pis.tag, txtCst_Cofins.tag, txtCst_Ipi.tag, txtPIcms.tag, txtPredIcms.Text, _
            txtPPis.Text, txtPCofins.Text) Then
            
            buscaEstoque
            cmdAddEstoque.caption = "Adicionar"
            cmdAlterarEstoque.tag = ""
            cmdAlterarEstoque.caption = "Alterar"
            txtEmpresa.Enabled = True
            pcEmpresa.Enabled = True
        End If
    Else
        If AlteraDadosEstoque(txtEmpresa.tag, Me.tag, txtCFOP_Revenda.tag, txtCFOP.tag, txtCstRevenda.tag, _
            txtCst_Icms.tag, txtCst_Pis.tag, txtCst_Cofins.tag, txtCst_Ipi.tag, txtPIcms.Text, txtPredIcms.Text, _
            txtPPis.Text, txtPCofins.Text) Then
            
            buscaEstoque
            cmdAlterarEstoque_Click
        End If
    End If
    
End Sub

Private Sub cmdAddSimilar_Click()
    frmCadSim.txtPrinc.Text = txtbxNome_Pro.Text
    frmCadSim.txtPrinc.tag = Me.tag
    frmCadSim.Show 1
End Sub

Private Sub cmdAlterarEstoque_Click()
    
    If Me.tag = "" Then
        MsgBox "Não há o que alterar", vbExclamation
        Exit Sub
    End If
    
    If cmdAlterarEstoque.tag = "" Then
        cmdAlterarEstoque.caption = "Cancelar"
        cmdAlterarEstoque.tag = adoestoque!cod_emp
        cmdAddEstoque.caption = "Confirmar"
        
        With adoestoque
            txtCst_Cofins.Text = TiraNull(!cst_cofins)
            txtCst_Icms.Text = TiraNull(!cst_icms_consumidor)
            txtCstRevenda.Text = TiraNull(!cst_icms_revenda)
            txtCst_Ipi.Text = TiraNull(!CST_IPI)
            txtCst_Pis.Text = TiraNull(!cst_pis)
            txtCFOP_Revenda.Text = TiraNull(!cfop_revenda)
            txtCFOP.Text = TiraNull(!cfop_consumidor)
            txtPIcms.Text = !pICMS
            txtPCofins.Text = !pCOFINS
            txtPPis.Text = !pPIS
            txtEmpresa.Text = TiraNull(!cod_emp)
            txtEmpresa.tag = !cod_emp
            txtEmpresa.Enabled = False
            pcEmpresa.Enabled = False
        End With
        
        
        
    Else
        cmdAlterarEstoque.caption = "Alterar"
        cmdAlterarEstoque.tag = ""
        txtEmpresa.Enabled = True
            pcEmpresa.Enabled = True
        cmdAddEstoque.caption = "Adicionar"
    End If
    
    
End Sub

Private Sub cmdExclAtacado_Click()
    On Error GoTo trata_Erro
    If MsgBox("Confirma exclusão do item?", vbYesNo + vbQuestion) = vbYes Then
        ExcluiProdutoAtacado adoAtac!id
        BuscaProdutoAtacado
    End If
    Exit Sub
trata_Erro:
    MostraErro
End Sub




Private Sub cmndbtnExclBarcode_Click()
    If lstCodBarra.ListIndex = -1 Then
        MsgBox "É necessário que se escolha um código de barras!", vbCritical
        lstCodBarra.SetFocus
        Exit Sub
    End If
    If lstCodBarra.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstCodBarra.ListCount = 1 Then
        MsgBox "Este código de barras não pode ser excluído por ser o único!", vbCritical
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o código de barras?", vbYesNo + vbCritical) = vbYes Then
        lstCodBarra.RemoveItem lstCodBarra.ListIndex
    End If
End Sub

Private Sub cmndbtnMarcas_Click()
    buscaMarca
End Sub

Private Sub cmndbtnRefresh_Click()
    
End Sub

Private Sub cmndbtnExclEmb_Click()
    If lstEmb.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstEmb.ListIndex = -1 Then
        MsgBox "É necessário que se escolha m item!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir a embalagem?", vbYesNo + vbCritical) = vbYes Then
        lstEmb.RemoveItem lstEmb.ListIndex
    End If
End Sub


Private Sub cmndbtnFoto_click()
    frmFoto.caption = txtbxNome_Pro.Text
    frmFoto.tag = Me.tag
    frmFoto.Show 1
End Sub



Private Sub cmdExcluir_Click()
    deletaVei
End Sub

Private Sub cmdExcluirConversao_Click()
    On Error GoTo trata_Erro
    If MsgBox("Confirma exclusão da conversão " & adoConversao!referencia & "?", vbYesNo + vbCritical) = vbYes Then
        Produtos.ExcluiProdutoConversao adoConversao!id
        buscaConversao
    End If
    
    
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdExcluirEstoque_Click()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    If strCadprod_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If

    
    
    
    strSQL = "select num_ven from " & TAB_ITVENDAS & " where cod_pro='" & Me.tag & "' and cod_emp='" & adoestoque!cod_emp & "'"
    
    abreObjetoAdo rPesq, strSQL
   
    If rPesq.RecordCount > 0 Then
        MsgBox "Existem vendas para esse produto e essa empresa!", vbExclamation
        Exit Sub
    End If
   
    strSQL = "select num_ven from sped_nf as n inner join sped_nf_itens as i "
    strSQL = strSQL & " on n.id=i.id_nf "
    strSQL = strSQL & " where cod_item='" & Me.tag & "' and cod_emp='" & adoestoque!cod_emp & "'"
    
    abreObjetoAdo rPesq, strSQL
   
    If rPesq.RecordCount > 0 Then
        MsgBox "Existem compras para esse produto e essa empresa!", vbExclamation
        Exit Sub
    End If
    
    If adoestoque.RecordCount = 1 Then
        MsgBox "Você não pode excluir o único estoque existente!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja mesmo excluir esse produto para essa empresa?", vbYesNo + vbCritical) = vbNo Then
        
        Exit Sub
    End If
    
    
    
    strSQL = "delete from estoque where cod_pro='" & Me.tag & "'"
    strSQL = strSQL & " and cod_emp='" & adoestoque!cod_emp & "' "
    conn.Execute strSQL
    buscaEstoque
End Sub


Private Sub cmdHist_Click()
    P_HistProd Me.tag
End Sub

Private Sub cmdImagens_Click()
    AbreImagens Me.tag
End Sub

Private Sub cmdSim_Click()
    AbrirSimilares Me.tag
End Sub

Private Sub cmdVisivel_Click()
    Dim strSQL As String
    If strAltprod_usu <> "S" Then
            MsgBox "Você não tem permissão pra essa operação", vbExclamation
            Exit Sub
    End If
    If Me.tag = "" Then
        MsgBox "Não há um produto cadastrado!", vbExclamation
        Exit Sub
    End If
    If F_BuscaUmCampo("visivel_pro", TAB_PRODUTOS, "cod_pro='" & Me.tag & "'") = "S" Then
        strSQL = "update " & TAB_PRODUTOS & " set visivel_pro='N' where cod_pro='" & Me.tag & "' "
        conn.Execute strSQL
    Else
        strSQL = "update " & TAB_PRODUTOS & " set visivel_pro='S' where cod_pro='" & Me.tag & "' "
    conn.Execute strSQL
    End If
    '
    
    
End Sub

Private Sub cmndbtnDelPro_Click()
    If strAltpreco_usu <> "S" And strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão pra essa operação", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        MsgBox "Não há um produto para excluir!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja mesmo excluir este produto?", vbYesNo + vbCritical) = vbYes Then
        If Produtos.ExcluirProduto(Me.tag) = True Then
            MsgBox "Produto excluído!", vbExclamation
            Unload Me
            Exit Sub
        End If
        
    End If
    
    
End Sub

Private Sub cmndbtnPro_Click()
    frmProdProm.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    If Me.tag = "" Then
        If strCadprod_usu <> "S" Then
            MsgBox "Você não tem permissão pra essa operação", vbExclamation
            Exit Sub
        End If
    Else
        If strAltprod_usu <> "S" Then
            MsgBox "Você não tem permissão pra essa operação", vbExclamation
            Exit Sub
        End If
    End If
    
    If txtCat.tag = "" Then
        MsgBox "Você deve informar a Categoria do produto!", vbExclamation
        txtCat.SetFocus
        Exit Sub
    End If
    'If txtMarca.tag = "" Then
    '    MsgBox "Você deve informar a marca do produto!", vbExclamation
    '    txtMarca.SetFocus
    '    Exit Sub
    'End If
    
    
   
    If Me.tag = "" Then
        strSQL = F_BuscaUmCampo("nome_pro", TAB_PRODUTOS, "NOME_PRO='" & _
            txtbxNome_Pro.Text & "'")
        If strSQL <> "" Then
            MsgBox "Já existe um produto com este nome!", vbCritical
            txtbxNome_Pro.SetFocus
            Exit Sub
        End If
        If txtbxCodFab.Text <> "" Then
        
            strSQL = F_BuscaUmCampo("nome_pro", TAB_PRODUTOS, "codfab_pro='" & _
            txtbxCodFab.Text & "'")
            
        End If
            If strSQL <> "" Then
                If MsgBox("Já existe um produto com esta referência!" & _
                     Chr(13) & strSQL & Chr(13) & "Deseja continuar?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
                End If
            End If
        
        If MsgBox("confirma inclusão do Produto?", vbYesNo + vbQuestion) = vbYes Then
            P_Grava
            
        End If
        
    Else
        If MsgBox("Confirma alteração do Produto?", vbYesNo + vbQuestion) = vbYes Then
            P_Altera
        End If
    End If
    
    
        
End Sub
Sub P_Altera()
    On Error GoTo trata_Erro
    Dim strSQL As String
    
    
   If AlteraProduto(Me.tag, txtBarra.Text, txtbxNome_Pro.Text, txtbxUnidade.Text, txtCat.tag, _
        txtMarca.tag, txtSubCat.tag, txtSubSec.tag, txtbxCodFab.Text, txtbxMargem.Text, txtbxLocal.Text, _
        txtFornecedor.tag, txtbxEstMin.Text, txtDetalhe.Text, txtbxLocal.Text, "1", txtNCM.Text, _
         "0", "0", txtbxCusto.Text, txtbxCusto.Text, txtbxPreco.Text, txtbxPreco.Text, IIf(chkProd.Value = 1, "S", "N"), _
        "", "S", txtGenero.tag, txtOrigem.tag, txtEan.Text, txtPIcms.Text, txtCst_Icms.tag, txtCstRevenda.tag, _
        txtCFOP.tag, txtCFOP_Revenda.tag, txtCst_Ipi.tag, _
        txtCst_Pis.tag, txtCst_Cofins.tag, txtCEST.Text, IIf(chkControlaEstoque.Value = 1, "S", "N"), txtQtdEmb.Text, txtQconv.Text, txtMultiplo.Text) Then
        
        If CCur(txtbxPreco.Text) <> CCur(txtbxPreco.tag) Then
            P_AltPreco Me.tag, txtbxPreco.Text, txtbxPreco.tag
        End If
        
        If chkComb.Value = 1 Then
            If Not Produtos.GravaProdutoCombustivel(Me.tag, txtcProdANP.Text, txtDescAnp.Text) Then
                MsgBox "Não gravou o produto-combustível", vbExclamation
            End If
        Else
            If Not Produtos.RemoveProdutoCombustivel(Me.tag) Then
                
            End If
        End If
        
        MsgBox "Dados gravados.", vbInformation
    Else
        MsgBox "Erro", vbInformation
    End If
    
    
   Exit Sub
    
    
   ' strSQL = F_GeraLog(strSQL, tab_produtos)
    MsgBox "Produto Alterado com Sucesso!", vbInformation
    Unload Me
    On Error Resume Next
    If IsLoad("frmEdiPro") Then
        frmEdiPro.ado.Requery
    End If
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub buscaAltPreco()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select date_format(Data_alt,'%d/%m/%y %T') AS data,antigo_pro,novo_pro,cod_usu from tab_altpreco where cod_pro='" & Me.tag & "' "
    strSQL = strSQL & " order by data_alt desc "
    abreObjetoAdo rPesq, strSQL
    
    GridUtil.AjustaGridData gradeAlteracoesPreco, 3, 1, 0, "Data,Antigo,Novo", "data,Antigo_pro,novo_pro", "1400,1000,1000", rPesq, "", True
    
    rPesq.Close
    
End Sub
Sub gravaEmb()
    Dim strSQL As String
    strSQL = "delete from tab_embalagens where cod_pro='" & Me.tag & "' "
    conn.Execute strSQL
    With lstEmb
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & Me.tag & "',"
            strSQL = strSQL & .Text & ")"
            conn.Execute strSQL
        Next
    End With
End Sub
Sub P_Grava()
    Dim sCod As String
    Dim strSQL As String
    
    Dim sMar As String
    Dim iCount As Integer
    Dim rPesq As New ADODB.Recordset
    Dim empresa As sped_empresa.sped_empresa

On Error GoTo trata_Erro


    strSQL = GravaProduto(txtBarra.Text, txtbxNome_Pro.Text, txtbxUnidade.Text, txtCat.tag, _
        txtMarca.tag, txtSubCat.tag, txtSubSec.tag, txtbxCodFab.Text, txtbxMargem.Text, txtbxLocal.Text, _
        txtFornecedor.tag, txtbxEstMin.Text, txtDetalhe.Text, txtbxLocal.Text, "1", txtNCM.Text, _
         "0", "0", txtbxCusto.Text, txtbxCusto.Text, txtbxPreco.Text, txtbxPreco.Text, IIf(chkProd.Value = 1, "S", "N"), _
         "", "S", txtGenero.tag, txtOrigem.tag, txtEan.Text, txtPIcms.Text, txtCst_Icms.tag, txtCstRevenda.tag, _
        txtCFOP.tag, txtCFOP_Revenda.tag, txtCst_Ipi.tag, _
        txtCst_Pis.tag, txtCst_Cofins.tag, txtCEST.Text, IIf(chkControlaEstoque.Value = 1, "S", "N"), "", txtQconv.Text, txtMultiplo.Text)
        
        
    If strSQL <> "" Then
        
        empresa = sped_empresa.SelecionarSped_Empresa("")
        
        sCod = empresa.codigo
        If sCod = "" Then
            sCod = InputBox("Informe o código da empresa que deseja adicionar o estoque.", "Empresa", "01")
        End If
        If F_BuscaUmCampo("codigo", "sped_empresa", "codigo='" & sCod & "'") = "" Then
            
            MsgBox "Não foi informada uma empresa válida!", vbExclamation
            Exit Sub
        End If
        If chkComb.Value = 1 Then
            If Not Produtos.GravaProdutoCombustivel(strSQL, txtcProdANP.Text, txtDescAnp.Text) Then
                MsgBox "Não gravou o produto-combustível", vbExclamation
            End If
        End If
            
        
        If AdicionaEstoque(sCod, strSQL, "0", "0", txtCFOP_Revenda.Text, txtCFOP.Text, txtCstRevenda.Text, txtCst_Icms.Text, _
             txtCst_Pis.Text, txtCst_Cofins.Text, txtCst_Ipi.Text, txtPIcms.Text, txtPredIcms.Text, txtPPis.Text, txtPCofins.Text) Then
        
            Me.tag = strSQL
            buscaEstoque
            MsgBox "Produto gravado com sucesso!", vbInformation
        Else
            MsgBox "erro ao gravar estoque!", vbInformation
        End If
    Else
        MsgBox "Erro na gravação do produto!", vbExclamation
    End If



Exit Sub


    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub




Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select max(data_not) as data, "
End Sub



 Private Sub dtgrd_DblClick()
    On Error GoTo trata_Erro
    cmdAddAtacado.tag = adoAtac!id
    cmdAddAtacado.caption = "Alterar"
    txtQtde.Text = adoAtac!qtd
    txtPreco.Text = adoAtac!preco
    txtBarraAtacado.Text = Format(adoAtac!barra)
    txtQtde.SetFocus
    Exit Sub
trata_Erro:
End Sub
Sub buscaSim()
        Dim strSQL As String
    strSQL = "select distinct tp.cod_pro as Codigo,tp.nome_pro as Nome,tp.codfab_pro as Ref,tp.er_pro as Er, tp.pv_pro as Pv,tp.cr_pro as cr,"
    strSQL = strSQL & " tm.nome_mar as marca,tc.nome_cat as categoria,ts.nome_sub as subcat, if(sim.cod_PRO=tp.cod_pro,'Sim','Não') AS Principal "
    strSQL = strSQL & " from tab_similares as sim, " & TAB_PRODUTOS & " as tp left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " left join tab_categorias as tc on tp.cod_cat=tc.cod_Cat "
    strSQL = strSQL & " left join tab_subcat as ts on tp.cod_sub=ts.cod_sub "
    strSQL = strSQL & " where (sim.cod_pro='" & Me.tag & "' or sim.cod_sim='" & Me.tag & "') "
    
    strSQL = strSQL & " and (tp.cod_pro=sim.cod_pro or tp.cod_pro=sim.cod_sim) "
    
    abreObjetoAdo adoSim, strSQL, gridSim
    

End Sub
Sub calcMLPraticada()
    On Error Resume Next
    If txtbxCusto.Text = "0" Then
        lblMlPraticada.caption = ""
    Else
        lblMlPraticada.caption = "M.L. Praticada" & Chr(13) & FormatVal(((CCur(txtbxPreco.Text) / CCur(txtbxCusto.Text)) - 1) * 100) & "%"
    End If
    
    
    
End Sub
Private Sub Form_Activate()
'    On Error GoTo trata_erro
    
    If bActive Then Exit Sub
    bActive = True
    cPreco = txtbxPreco.Text
    
    If bUsaVeiculo = False Then
        sstb.TabVisible(2) = False
        'sstb.TabVisible(3) = False
        sstb.TabVisible(4) = False
        sstb.TabVisible(5) = False
        
    
    End If
    
    If Me.tag <> "" Then
        Carrega
        
    End If
    
    If strAltpreco_usu = "S" Then
        lblPrecoCap.Enabled = True
        lblCusCap.Enabled = True
        txtbxPreco.Enabled = True
        txtbxCusto.Enabled = True
        
    Else
        lblPrecoCap.Enabled = False
        lblCusCap.Enabled = False
        txtbxPreco.Enabled = False
        txtbxCusto.Enabled = False
        'txtbxCustoFiscal.Enabled = False
    End If
    calcMLPraticada
    'P_PreencheCateg cmbxCat
    'P_PreencheMarca cmbxMarca
    Exit Sub
    
trata_Erro:
MostraErro
End Sub
Sub buscaProdutoComb()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = "select * from produtos_comb where cod_pro='" & Me.tag & "'"
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount = 0 Then
            chkComb.Value = 0
        Else
            chkComb.Value = 1
            txtcProdANP.Text = TiraNull(!cProdAnp)
            txtDescAnp.Text = TiraNull(!descAnp)
        End If
        
        
        .Close
    End With
    
    
End Sub
Sub Carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = "select ti.cod_imp,pis.descricao as pis, icms.descricao as icms, cofins.descricao as cofins, ipi.descricao as ipi, ti.desc_imp,tp.*,ss.nome_sub as nome_ss,cod_sec, tsc.nome_sub,tm.nome_mar,"
    strSQL = strSQL & "tc.nome_cat,gen.descricao as genero from " & TAB_PRODUTOS & " as tp left join sped_genero as gen"
    strSQL = strSQL & " on tp.cod_gen=gen.codigo "
    strSQL = strSQL & " left join  " & TAB_MARCAS & "  as tm on tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " left join " & TAB_CATEGORIAS & " as tc on tp.cod_cat=tc.cod_cat "
    strSQL = strSQL & " left join " & TAB_SUBCAT & " as tsc on tp.cod_sub=tsc.cod_sub "
    strSQL = strSQL & " left join " & TAB_SUBSEC & " as ss on tp.subsecao=ss.cod_sub "
    strSQL = strSQL & " left join " & TAB_IMPOSTOS & " as ti on tp.cod_imp=ti.cod_imp"
    strSQL = strSQL & " left join sped_cst_pis as pis on tp.cst_pis=pis.codigo"
    strSQL = strSQL & " left join sped_cst_cofins as cofins on tp.cst_cofins=cofins.codigo"
    strSQL = strSQL & " left join sped_cst_ipi as ipi on tp.cst_ipi=ipi.codigo"
    strSQL = strSQL & " left join nfe_tipo_icms as icms on tp.cst_icms=icms.codigo"
    strSQL = strSQL & " where tp.cod_pro='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Me.tag = !cod_pro
        Me.caption = !cod_pro & " - " & !nome_pro & " - Ref: " & !codfab_pro
        txtMarca.Text = !nome_mar
        txtMarca.tag = !cod_mar
        txtCat.Text = !nome_cat
        txtCat.tag = !cod_cat
        txtSubCat.tag = !cod_sub
        txtSubCat.Text = !nome_sub
        txtFornecedor.Text = F_BuscaUmCampo("nome_for", TAB_FORNECEDORES, "cod_for='" & !cod_For & "'")
        txtFornecedor.tag = !cod_For
        txtSubSec.Text = !nome_ss
        txtSubSec.tag = !s
        lblSecao.caption = F_BuscaUmCampo("nome_sec", TAB_SECOES, "cod_sec='" & !cod_sec & "'")
        txtQtdEmb.Text = !qtdembcompra_pro
        txtCEST.Text = TiraNull(!cest_pro)
        If !visivel_pro = "N" Then
            chkVis.Value = 0
        End If
        Me.txtbxEstMin.Text = !estmin_pro
        If TiraNull(!controlaestoque_pro) = "N" Then
            chkControlaEstoque.Value = 0
        Else
            chkControlaEstoque.Value = 1
        End If
        
        If !balanca_pro = "S" Then
            Me.chkProd.Value = 1
        Else
            Me.chkProd.Value = 0
        End If
        Me.txtbxCodFab.Text = !codfab_pro
        'me.txtbxCustoFiscal.Text = FormatVal(!c1_pro)
        
        'me.txtbxEstoque.Text = FormatVal(!estoque_pro)
        Me.txtbxNome_Pro.Text = !nome_pro
        Me.txtbxPreco.Text = FormatVal(!pv_pro)
        txtbxPreco.tag = FormatVal(!pv_pro)
        txtGenero.Text = !genero
        txtGenero.tag = !id_gen
        txtOrigem.Text = F_BuscaUmCampo("descricao", "sped_origem", "codigo='" & !cod_ori & "'")
        txtOrigem.tag = !cod_ori
        
        txtQconv.Text = !qConv
        txtDetalhe.Text = TiraNull(!detalhe_pro)
        Me.txtbxCusto.Text = FormatVal(!cr_pro)
        txtMultiplo.Text = !multiplo_pro
        txtEan.Text = !ean_pro
        txtBarra.Text = !codbarra_pro
        
        'me.txtbxCodBar = !codbarra_pro
        Me.txtbxUnidade.Text = !unidade_pro
        If Not IsNull(!ncm_pro) Then
            Me.txtNCM.Text = !ncm_pro
        End If
        Me.txtbxMargem.Text = FormatVal(!lucro_pro)
        Me.txtbxLocal = !endereco_pro
        
        
        txtCst_Icms.Text = Trim(!cst_icms)
        txtCst_Icms.tag = !cst_icms
        'txtCst_Ipi.Text = Format(!ipi)
        txtCst_Ipi.Text = Format(!cst_ip)
        
        txtCFOP.Text = TiraNull(!cfop_pro)
        txtCFOP.tag = TiraNull(!cfop_pro)
        txtCFOP_Revenda.Text = TiraNull(!cfop_revenda)
        txtCFOP_Revenda.tag = TiraNull(!cfop_revenda)
        txtCstRevenda.Text = TiraNull(!cst_revenda)
        txtCstRevenda.tag = TiraNull(!cst_revenda)
        
        txtCst_Ipi.tag = !CST_IPI
        'txtCst_Pis.Text = Trim(!cst_pis & " " & Format(!pis))
        txtCst_Pis.Text = Trim(!cst_pis)
        txtCst_Pis.tag = !cst_pis
        txtCst_Cofins.Text = Trim(!cst_cofins)
        txtCst_Cofins.tag = !cst_cofins
        chkControlaEstoque_Click
        xtbxEstoque.Text = FormatVal(!er_pro)
        lblReservado.caption = !reservado_pro
        buscaEstoque
        buscaAltPreco
        buscaProdutoComb
        
        'P_LocalizaOperacao !cod_gen, cmbxGenero
        
        BuscaProdutoAtacado
        buscaSim
        buscaVei
        buscaConversao
        txtCodigo.Text = Me.tag
        txtCodigo.Visible = True
        
        .Close
    End With
End Sub
Sub busca_cst()
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
        
    End If
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    End If
    
End Sub
Private Sub Form_Load()
    sstb.Tab = 0
    bActive = False
    
    visivel = "S"
    'P_PreencheImpostos cmbxICMS
    'P_PreencheNFE_Cst cmbxNFECst
    'P_PreencheGenero cmbxGenero
    'cmbxSt.ListIndex = 1
    Set adoAtac = New ADODB.Recordset
    Set adoSim = New ADODB.Recordset
    Set adoVei = New ADODB.Recordset
    Set adoConversao = New ADODB.Recordset
    txtCstRevenda.Text = F_SoNumero(strTipo_IcmsNFe)
    txtCst_Icms.Text = txtCstRevenda.Text
End Sub

Private Sub Form_Resize()
'     P_CentraObejct frControle, Me
End Sub

Private Sub frControle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub


Private Sub lblInfoEnt_Click()
    pcEnt_Click
End Sub

Private Sub lblInfoEnt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = True
    lblInfoEnt.FontUnderline = True
    pcEnt.BorderStyle = 1
End Sub

Private Sub lblInfoSai_Click()
    pcSai_Click
End Sub

Private Sub lblInfoSai_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoSai.FontBold = True
    lblInfoSai.FontUnderline = True
    pcSai.BorderStyle = 1
End Sub

Private Sub lstCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub lstEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub gridestoque_DblClick()
    cmdAlterarEstoque_Click
End Sub

Private Sub gridVeiculos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        deletaVei
    End If
End Sub
Sub deletaVei()
    If Me.tag = "" Then
        MsgBox "O produto precisa estar gravado para efetuar essa operação.", vbExclamation
        Exit Sub
    End If
    'If strAltpreco_usu <> "" Then
    '    MsgBox "Você não tem permissão para esse acesso!", vbExclamation
    '
    'End If
    
    If Veiculos.DeletaProduto_Veiculo(Me.tag, adoVei!id_vei) Then
         buscaVei
    Else
        MsgBox "Erro!", vbExclamation
        
    End If
End Sub

Private Sub pcBuscaMarcaConversao_Click()
    buscaMarcaConversao
End Sub
Sub buscaMarcaConversao()
    Dim m As Marca.Marca
    m = Marca.buscaMarca(txtMarcaConversao.Text)
    If m.codigo <> "" Then
        txtMarcaConversao.Text = m.nome
        txtMarcaConversao.tag = m.codigo
        txtReferenciaConversao.SetFocus
    End If
End Sub
Private Sub pcCat_Click()
    BuscaCategoria
End Sub
Sub BuscaCategoria()
    Dim c As categoria.categoria
    c = categoria.BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        txtSubCat.SetFocus
    End If
End Sub
Sub buscaMarca()
    Dim c As Marca.Marca
    c = Marca.buscaMarca(txtMarca.Text)
    If c.codigo <> "" Then
        txtMarca.Text = c.nome
        txtMarca.tag = c.codigo
        txtbxNome_Pro.SetFocus
    End If
End Sub

Private Sub pcCfopConsumidor_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCFOP.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_CFOP(expr)
    If p.codigo <> "" Then
        txtCFOP.Text = p.codigo & " - " & p.nome
        txtCFOP.tag = p.codigo
        
        
    End If
    
End Sub

Private Sub pcCfopRevenda_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCFOP_Revenda.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_CFOP(expr)
    If p.codigo <> "" Then
        txtCFOP_Revenda.Text = p.codigo & " - " & p.nome
        txtCFOP_Revenda.tag = p.codigo
        
        
    End If
End Sub

Private Sub pcCstCofins_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Cofins.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Cofins(expr)
    If p.codigo <> "" Then
        txtCst_Cofins.Text = p.codigo
        txtCst_Cofins.tag = p.codigo
        
    End If
End Sub

Private Sub pcCstIcms_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Icms.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Icms(expr)
    If p.codigo <> "" Then
        txtCst_Icms.Text = p.codigo & " - " & p.nome
        txtCst_Icms.tag = p.codigo
        txtCstRevenda.SetFocus
        
    End If
    
    
End Sub

Private Sub pcCstIpi_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Ipi.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_IPI(expr)
    If p.codigo <> "" Then
        txtCst_Ipi.Text = p.codigo & " - " & p.nome
        txtCst_Ipi.tag = p.codigo
        txtCst_Pis.SetFocus
    End If
End Sub

Private Sub pcCstPis_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Pis.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Pis(expr)
    If p.codigo <> "" Then
        txtCst_Pis.Text = p.codigo
        txtCst_Pis.tag = p.codigo
        txtCst_Cofins.SetFocus
    End If
End Sub

Private Sub pcEnt_Click()
    If Me.tag <> "" Then
        P_InfoEntradasPro Me.tag
    Else
        MsgBox "Não há produto gravado!", vbCritical
    End If
End Sub

Private Sub pcEnt_GotFocus()
    lblInfoEnt_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcEnt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcEnt_Click
    End If
End Sub

Private Sub pcEnt_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcCstRevenda_Click()
        Dim p As Padrao1
    Dim expr As String
    
    expr = txtCstRevenda.Text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Icms(expr)
    If p.codigo <> "" Then
        txtCstRevenda.Text = p.codigo & "-" & p.nome
        txtCstRevenda.tag = p.codigo
        txtCFOP.SetFocus
    End If

End Sub

Private Sub pcEmpresa_Click()
    buscaEmpresa
End Sub

Private Sub pcForn_Click()
    buscaFor
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    
    f = fornecedor.F_BuscaFornecedor(txtFornecedor.Text)
    If f.cod_For <> "" Then
        txtFornecedor.Text = f.nome_For
        txtFornecedor.tag = f.cod_For
        txtbxLocal.SetFocus
    End If
End Sub
Sub buscaEmpresa()
    Dim f As sped_empresa.sped_empresa
    
    f = sped_empresa.BuscaSpedEmpresa(txtEmpresa.Text)
    If f.codigo <> "" Then
        txtEmpresa.Text = f.apelido
        txtEmpresa.tag = f.codigo
        
    End If
End Sub


Private Sub pcGenero_Click()
    buscaGenero
End Sub

Private Sub pcInfo_Click()
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub pcInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub



Private Sub pcMarca_Click()
    buscaMarca
End Sub

Private Sub pcSai_Click()
    If Me.tag = "" Then Exit Sub
    P_InfoSaidasPro Me.tag
End Sub

Private Sub pcSai_GotFocus()
    lblInfoSai_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcSai_Click
    End If
End Sub

Private Sub pcSai_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub



Private Sub pcOrigem_Click()
    buscaOrigem
End Sub

Private Sub pcSubCat_Click()
    BuscaSubCateg
End Sub

Private Sub pcSubSec_Click()
    buscaSubSecao
End Sub
Sub buscaSubSecao()
    Dim s As Produtos.subsecao
    s = Produtos.F_BuscaSubSec(txtSubSec.Text)
    If s.id <> "" Then
        txtSubSec.Text = s.nome
        txtSubSec.tag = s.id
        lblSecao.caption = s.nome_secao
        
        txtGenero.SetFocus
    End If
    
    
End Sub
Sub BuscaSubCateg()
    Dim s As SubCategoria
    s = BuscaSubCategoria(txtSubCat.Text)
    If s.codigo <> "" Then
        txtSubCat.Text = s.nome
        txtSubCat.tag = s.codigo
        
        
        
    End If
    
    
End Sub

Sub buscaGenero()
    Dim p As Padrao1
    p = Sped_Genero(txtGenero.Text)
    If p.codigo <> "" Then
        txtGenero.Text = p.nome
        txtGenero.tag = p.codigo
        txtOrigem.SetFocus
    End If
End Sub
Sub buscaOrigem()
    Dim p As Padrao1
    p = Sped_Origem(txtOrigem.Text)
    If p.codigo <> "" Then
        txtOrigem.Text = p.nome
        txtOrigem.tag = p.codigo
        txtbxUnidade.SetFocus
    End If
End Sub

Private Sub pcUnidade_Click()
    buscaUnidade
End Sub
Sub buscaUnidade()
    Dim p As Padrao1
    p = SelectUnidade(txtbxUnidade.Text)
    If p.codigo <> "" Then
        txtbxUnidade.Text = p.codigo
        txtbxLocal.SetFocus
    End If
    
End Sub



Private Sub pcVeic_Click()
    Dim v As Veiculo
    v = Veiculos.BuscaVeiculo(txtVeiculo.Text, "")
    If v.id > 0 Then
        txtVeiculo.Text = v.descricao
        txtVeiculo.tag = v.id
        lblVeiculoFabricante.caption = v.fabricante.descricao
        txtAplicacao.SetFocus
    End If
End Sub



Private Sub txtAplicacao_GotFocus()
    P_SelecTexto txtAplicacao
End Sub

Private Sub txtAplicacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAplicacao_LostFocus()
    P_AjustaTexto txtAplicacao, nada_muda
End Sub

Private Sub txtBarra_GotFocus()
    ganhaFoco txtBarra
End Sub

Private Sub txtBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBarra_LostFocus()
    AjustaTexto txtBarra, 0
End Sub

Private Sub txtBarraAtacado_GotFocus()
    ganhaFoco txtBarraAtacado
End Sub

Private Sub txtBarraAtacado_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub




Private Sub txtbxPrecoMin_GotFocus()
    
End Sub

Private Sub txtbxCodFab_GotFocus()
    ganhaFoco txtbxCodFab
End Sub

Private Sub txtbxCodFab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCodFab_LostFocus()
    AjustaTexto txtbxCodFab
End Sub

Private Sub txtbxDescricao_GotFocus()
    'GanhaFoco txtbxDescricao
End Sub

Private Sub txtbxDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDescricao_LostFocus()
    'AjustaTexto txtbxDescricao
End Sub

Private Sub txtbxEstMin_GotFocus()
    ganhaFoco txtbxEstMin
End Sub

Private Sub txtbxEstMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxEstMin_LostFocus()
    PerdeFoco txtbxEstMin
End Sub

Private Sub txtbxCusto_GotFocus()
    ganhaFoco txtbxCusto
End Sub

Private Sub txtbxCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCusto_LostFocus()
    PerdeFoco txtbxCusto
End Sub




Private Sub txtbxEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub


Private Sub txtbxEstoque_GotFocus()
    ganhaFoco txtbxEstoque
End Sub

Private Sub txtbxEstoque_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
     End If
End Sub

Private Sub txtbxEstoque_LostFocus()
    PerdeFoco txtbxEstoque
End Sub



Private Sub txtbxLocal_GotFocus()
    ganhaFoco txtbxLocal
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxLocal_LostFocus()
    AjustaTexto txtbxLocal
End Sub

Private Sub txtbxMargem_GotFocus()
    ganhaFoco txtbxMargem
End Sub

Private Sub txtbxMargem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxMargem_LostFocus()
    PerdeFoco txtbxMargem
End Sub

Private Sub txtbxNome_Pro_GotFocus()
    ganhaFoco txtbxNome_Pro
End Sub

Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNome_Pro_LostFocus()
    P_AjustaTexto txtbxNome_Pro
End Sub

Private Sub txtbxPei_GotFocus()
    
End Sub


Private Sub txtbxPreco_GotFocus()
    ganhaFoco txtbxPreco
End Sub

Private Sub txtbxPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxPreco_LostFocus()
    PerdeFoco txtbxPreco
    calcMLPraticada
End Sub

Private Sub txtbxPrecoMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxPrecoMin_LostFocus()
    
End Sub


Private Sub txtbxUltComp_Change()

End Sub

Private Sub txtbxUnidade_GotFocus()
    ganhaFoco txtbxUnidade
End Sub

Private Sub txtbxUnidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaUnidade
    End If
End Sub

Private Sub txtbxUnidade_LostFocus()
    AjustaTexto txtbxUnidade
End Sub

Private Sub txtCat_GotFocus()
    ganhaFoco txtCat
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCat_Click
    End If
End Sub

Private Sub txtCEST_GotFocus()
    P_SelecTexto txtCEST
End Sub

Private Sub txtCEST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCEST_LostFocus()
    P_AjustaTexto txtCEST
End Sub

Private Sub txtcodANP_Change()

End Sub

Private Sub txtCFOP_GotFocus()
    P_SelecTexto txtCFOP
End Sub

Private Sub txtCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        pcCfopConsumidor_Click
    
    End If
End Sub

Private Sub txtCFOP_Revenda_GotFocus()
    P_SelecTexto txtCFOP_Revenda
End Sub

Private Sub txtCFOP_Revenda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCfopRevenda_Click
    End If
End Sub

Private Sub txtcProdANP_GotFocus()
    P_SelecTexto txtcProdANP
End Sub

Private Sub txtcProdANP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtcProdANP_LostFocus()
    P_AjustaTexto txtcProdANP
End Sub

Private Sub txtCst_Cofins_GotFocus()
    P_SelecTexto txtCst_Cofins
End Sub

Private Sub txtCst_Cofins_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCstCofins_Click
    End If
    
End Sub

Private Sub txtCst_Icms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCstIcms_Click
    End If
End Sub

Private Sub txtCst_Icms_LostFocus()
    pcCstIcms_Click
End Sub

Private Sub txtCst_Ipi_GotFocus()
    P_SelecTexto txtCst_Ipi
End Sub

Private Sub txtCst_Ipi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCstIpi_Click
    End If
End Sub

Private Sub txtCst_Pis_GotFocus()
    P_SelecTexto txtCst_Pis
End Sub

Private Sub txtCst_Pis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCstPis_Click
    End If
End Sub





Private Sub txtCstRevenda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCstRevenda_Click
    End If
End Sub

Private Sub txtCstRevenda_LostFocus()
    pcCstRevenda_Click
End Sub

Private Sub txtDescANP_GotFocus()
    P_SelecTexto txtDescAnp
End Sub

Private Sub txtDescANP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtDescANP_LostFocus()
    P_AjustaTexto txtDescAnp
End Sub
Private Sub txtEan_Change()
    txtEanCopia.Text = txtEan.Text
End Sub

Private Sub txtEanCopia_Change()
    txtEan.Text = txtEanCopia.Text
End Sub

Private Sub txtFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcForn_Click
    End If
End Sub

Private Sub txtGenero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaGenero
    End If
End Sub

Private Sub txtMarca_GotFocus()
    ganhaFoco txtMarca
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcMarca_Click
    End If
End Sub

Private Sub txtMarcaConversao_Change()
    txtMarcaConversao.tag = ""
End Sub

Private Sub txtMarcaConversao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaMarcaConversao
    End If
End Sub

Private Sub txtNCM_Change()
    txtNCMCopia.Text = txtNCM.Text
End Sub

Private Sub txtNCM_GotFocus()
    ganhaFoco txtNCM
End Sub

Private Sub txtNCM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtNCM_LostFocus()
    AjustaTexto txtNCM
End Sub

Private Sub txtNCMCopia_Change()
    txtNCM.Text = txtNCMCopia.Text
End Sub

Private Sub txtPIcms_GotFocus()
    P_SelecTexto txtPIcms
End Sub

Private Sub txtPIcms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtPIcms_LostFocus()
    P_ValorLostFocus txtPIcms
End Sub
Private Sub txtPPis_GotFocus()
    P_SelecTexto txtPPis
End Sub

Private Sub txtPPis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtPPis_LostFocus()
    P_ValorLostFocus txtPPis
End Sub
Private Sub txtPredIcms_GotFocus()
    P_SelecTexto txtPredIcms
End Sub

Private Sub txtPredIcms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtPredIcms_LostFocus()
    P_ValorLostFocus txtPredIcms
End Sub
Private Sub txtPCofins_GotFocus()
    P_SelecTexto txtPCofins
End Sub

Private Sub txtPCofins_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtPCofins_LostFocus()
    P_ValorLostFocus txtPCofins
End Sub
Private Sub txtPreco_GotFocus()
    ganhaFoco txtPreco
End Sub

Private Sub txtPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtPreco_LostFocus()
    PerdeFoco txtPreco
End Sub

Private Sub txtQtde_GotFocus()
    ganhaFoco txtQtde
End Sub

Private Sub txtQtde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtQtde_LostFocus()
    PerdeFoco txtQtde
End Sub

Private Sub txtQtdEmb_GotFocus()
    P_SelecTexto txtQtdEmb
End Sub

Private Sub txtQtdEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQtdEmb_LostFocus()
    P_ValorLostFocus txtQtdEmb
End Sub



Private Sub txtMultiplo_GotFocus()
    P_SelecTexto txtMultiplo
End Sub

Private Sub txtMultiplo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtMultiplo_LostFocus()
    P_ValorLostFocus txtMultiplo, "###"
End Sub


Private Sub txtReferenciaConversao_GotFocus()
    P_SelecTexto txtReferenciaConversao
End Sub

Private Sub txtReferenciaConversao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtReferenciaConversao_LostFocus()
    P_AjustaTexto txtReferenciaConversao
End Sub

Private Sub txtSubCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcSubCat_Click
    End If
End Sub

Private Sub txtSubSec_GotFocus()
    ganhaFoco txtSubSec
End Sub

Private Sub txtSubSec_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        pcSubSec_Click
    End If
End Sub

Private Sub txtVeiculo_Change()
    txtVeiculo.tag = ""
End Sub

Private Sub txtVeiculo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcVeic_Click
    End If
End Sub
