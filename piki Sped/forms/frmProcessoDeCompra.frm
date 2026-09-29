VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProcessoDeCompra 
   Caption         =   "Processo Eletrônico de Compras"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15270
   Icon            =   "frmProcessoDeCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   15270
   Begin VB.Frame Frame1 
      Caption         =   "&Fornecedor:"
      Height          =   615
      Left            =   480
      TabIndex        =   44
      Top             =   2880
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox cmbxFor 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.PictureBox pcFiltro 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   15240
      TabIndex        =   36
      Top             =   1050
      Width           =   15270
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9240
         TabIndex        =   46
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtCategoria 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   41
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtFiltrar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   37
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdForn 
         Height          =   285
         Left            =   5760
         Picture         =   "frmProcessoDeCompra.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Similares"
         Height          =   195
         Left            =   12720
         TabIndex        =   49
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblSimilar 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   13080
         TabIndex        =   48
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar Marca"
         Height          =   195
         Left            =   9240
         TabIndex        =   47
         Top             =   0
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar Categoria"
         Height          =   195
         Left            =   6120
         TabIndex        =   42
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar produtos:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar fornecedor"
         Height          =   195
         Left            =   3000
         TabIndex        =   39
         Top             =   0
         Width           =   1185
      End
   End
   Begin MSDataGridLib.DataGrid grade 
      Height          =   2655
      Left            =   0
      TabIndex        =   35
      Top             =   1800
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   4683
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "sts_ipc"
         Caption         =   "M"
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
      BeginProperty Column02 
         DataField       =   "codfab_pro"
         Caption         =   "Cód.Fab."
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "origem"
         Caption         =   "Origem"
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
         DataField       =   "er_pro"
         Caption         =   "Estoque"
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
         DataField       =   "reservado_pro"
         Caption         =   "Res"
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
         DataField       =   "mmd_pro"
         Caption         =   "MMD"
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
         DataField       =   "qtd"
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
      BeginProperty Column11 
         DataField       =   "qtdembcompra_pro"
         Caption         =   "Emb."
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
         DataField       =   "qfa_ipc"
         Caption         =   "A Comprar"
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
      BeginProperty Column13 
         DataField       =   "upc_itc"
         Caption         =   "Un. Previsto"
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
         DataField       =   "total"
         Caption         =   "Total Prev."
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
      BeginProperty Column15 
         DataField       =   "nome_for"
         Caption         =   "Fornecedor"
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
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1635,024
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column15 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pctop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   1050
      ScaleWidth      =   15270
      TabIndex        =   13
      Top             =   0
      Width           =   15270
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   595
         Left            =   11640
         Picture         =   "frmProcessoDeCompra.frx":0694
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFecha 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   13080
         Picture         =   "frmProcessoDeCompra.frx":0B85
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Fecha formulário de Compra"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar Itens"
         Height          =   615
         Left            =   4440
         Picture         =   "frmProcessoDeCompra.frx":0DCF
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnList 
         Caption         =   "&Buscar Produtos"
         Height          =   615
         Left            =   4440
         Picture         =   "frmProcessoDeCompra.frx":1159
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Listar Mercadorias"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnInicia 
         Caption         =   "&Iniciar Processo"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5880
         Picture         =   "frmProcessoDeCompra.frx":1B03
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Iniciar Processo de Compra"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinaliza 
         Caption         =   "&Concluir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7320
         Picture         =   "frmProcessoDeCompra.frx":208D
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Finalizar Processo de Compra"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCancela 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8760
         Picture         =   "frmProcessoDeCompra.frx":21D7
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cancelar Processo de Compra"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprList 
         Caption         =   "Imprimir Lista"
         Enabled         =   0   'False
         Height          =   615
         Left            =   10200
         Picture         =   "frmProcessoDeCompra.frx":2321
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.PictureBox frPeriod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   825
         ScaleWidth      =   4185
         TabIndex        =   14
         Top             =   0
         Width           =   4215
         Begin VB.TextBox txtbxTotalDias 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
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
            Height          =   285
            Left            =   3000
            TabIndex        =   23
            Text            =   "10"
            ToolTipText     =   "Tempo de duração desse processo..."
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtbxDuracao 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   2160
            TabIndex        =   20
            Text            =   "1"
            ToolTipText     =   "Tempo de duração desse processo..."
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtbxEntrega 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Text            =   "2"
            ToolTipText     =   "Tempo previsto para entrega"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtbxSustent 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Text            =   "7"
            ToolTipText     =   "Refere-se ao período de sustentação(Tempo em que o estoque deve durar)"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Total de Dias:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3000
            TabIndex        =   22
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "&Duração:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2160
            TabIndex        =   21
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "&Prazo Entrega:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1080
            TabIndex        =   18
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "&Suprimento"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Comprar para quantos dias?"
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
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   2385
         End
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   6960
      Width           =   15270
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Incluir Produto &Não Sugerido:"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9120
         TabIndex        =   31
         Top             =   0
         Width           =   4455
         Begin VB.CommandButton cmndbtnInc_Mer 
            Caption         =   "..."
            Height          =   285
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtbxInc_Mer 
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   32
            ToolTipText     =   "Informe o nome da mercadoria seguido de <ENTER> para localizá-la."
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label lblMarca_mer 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            ForeColor       =   &H8000000C&
            Height          =   195
            Left            =   1560
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Itens:"
         Height          =   735
         Left            =   3720
         TabIndex        =   6
         Top             =   0
         Width           =   5415
         Begin VB.Label lblMax 
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
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   3840
            TabIndex        =   51
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label12 
            Caption         =   "Projeção Máxima:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   3360
            TabIndex        =   50
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   2280
            TabIndex        =   12
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblListados 
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
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   1560
            TabIndex        =   11
            Top             =   480
            Width           =   90
         End
         Begin VB.Label lblTotalPrev 
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
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   2280
            TabIndex        =   10
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Listados:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Selecionados:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label lblSelec 
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
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   480
            TabIndex        =   7
            Top             =   480
            Width           =   90
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ordem:"
         Height          =   735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3615
         Begin VB.OptionButton optnrOrderBy 
            Caption         =   "&Quantidade"
            Height          =   195
            Index           =   2
            Left            =   2280
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optnrOrderBy 
            Caption         =   "Descrição"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optnrOrderBy 
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblCompra 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "000000"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   540
         End
      End
   End
   Begin VB.Menu mnOpt 
      Caption         =   "Opções"
      Begin VB.Menu mnAjust 
         Caption         =   "Ajustar quantidades"
      End
      Begin VB.Menu mnEditProd 
         Caption         =   "Editar produto"
      End
   End
End
Attribute VB_Name = "frmProcessoDeCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmbxCat_Click()
    cmbxCat.tag = F_BuscaUmCampo("cod_cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
End Sub

Private Sub cmbxLoc_Ipc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
    
End Sub

Private Sub cmbxMarca_Click()
    cmbxMarca.tag = F_BuscaUmCampo("Cod_mar", TAB_MARCAS, "nome_mar='" & cmbxMarca.Text & "'")
    P_SalvaComp
    P_BuscaItensCompra lblCompra.caption
End Sub


Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub


Private Sub cmndbtnAddFor_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If dtForn.Recordset.RecordCount = 0 Then
'        MsgBox "Não há fornecedores listados!", vbCritical
        Exit Sub
    End If
    With dtForn.Recordset
        strSQL = "update tab_fornpcp set sts_fpc='S' "
        strSQL = strSQL & " where cod_ter = '" & !cod_For & "' "
        strSQL = strSQL & " and seq_pcp ='" & lblCompra.caption & "' "
        conn.Execute strSQL
    End With
    P_buscaForn
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If bActive = False Then Exit Sub
    lblSimilar.caption = BuscaEstoqueSim(ado!cod_pro)
    grade.caption = ado.AbsolutePosition & " de " & ado.RecordCount & " registros listados "
End Sub

Private Sub cmbxFor_Click()
    cmbxFor.tag = F_BuscaUmCampo("cod_for", TAB_FORNECEDORES, "nome_For='" & cmbxFor.Text & "'")
    Carrega
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV grade, ado
End Sub

Private Sub cmdForn_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtForn.Text)
    If f.cod_For <> "" Then
        txtForn.Text = f.nome_For
        txtForn.tag = f.cod_For
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    bActive = False
    P_BuscaItensCompra lblCompra.caption
    bActive = True
End Sub



Private Sub cmndbtnCancela_Click()
    On Error GoTo trata_Erro
    Dim strSQL As String
    If MsgBox("Deseja realmente excluir o Processo de Compra?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "Delete from tab_procomp where seq_pcp='" & lblCompra.caption & "' "
        conn.Execute strSQL
        strSQL = "Delete from tab_itenspcp where seq_pcp='" & lblCompra.caption & "' "
        conn.Execute strSQL
        MsgBox "Processo Cancelado!", vbCritical
        Unload Me
        
    End If
        
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnFecha_Click()
    Unload Me
End Sub

Private Sub cmndbtnFinaliza_Click()
    Dim strSQL As String
    If MsgBox("Confirma Finalização do Processo de Compra: " & lblCompra.caption & " ?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = F_BuscaUmCampo("count(cod_pro)", "TAB_ITENSPCP", "STS_IPC='S' and SEQ_PCP='" & lblCompra.caption & "'")
        If strSQL = "0" Then
            MsgBox "Nenhum item foi selecionado!", vbCritical
            Exit Sub
        End If
        strSQL = F_BuscaUmCampo("count(COD_ter)", "TAB_FORNPCP", " SEQ_PCP='" & lblCompra.caption & "'")
        If strSQL = "0" Then
            strSQL = "Não foi informado nenhum fornecedor!"
            strSQL = strSQL & Chr(13) & "Deseja Continuar mesmo assim?"
            If MsgBox(strSQL, vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
            
        End If
        strSQL = "Update tab_procomp set sts_pcp= '2',COD_usu='" & strCod_Usu & "' "
        strSQL = strSQL & " where seq_pcp='" & lblCompra.caption & "' "
        
        conn.Execute strSQL
        strSQL = "Update tab_itenspcp set sts_ipc = 'C' where sts_ipc = 'N' and seq_pcp='" & lblCompra.caption & "' "
        conn.Execute strSQL  'passa de N -não pra C -cancelado
        If MsgBox("Processo Finalizado! Deseja imprimir?", vbYesNo + vbQuestion) = vbYes Then
            ImprimePCPHtml lblCompra.caption
        End If
        
        
        
        Unload Me
    End If
End Sub

Private Sub cmndbtnForn_Click()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select "
    strSQL = strSQL & " distinct tt.*"
    strSQL = strSQL & " From "
    strSQL = strSQL & " " & TAB_FORNECEDORES & " as tt,"
    strSQL = strSQL & " tab_itcmp as ti,"
    strSQL = strSQL & " tab_cmp as te,"
    strSQL = strSQL & " tab_itenspcp As tc"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " ti.seq_not=te.seq_not "
    strSQL = strSQL & " and ti.cod_pro = tc.cod_pro"
    
    strSQL = strSQL & " and te.cod_for = tt.cod_for"
    
    strSQL = strSQL & " and seq_pcp = '" & lblCompra.caption & "'"
    rPesq.Open strSQL, conn_client
    With rPesq
        strSQL = "Delete from tab_fornpcp where seq_pcp = '" & lblCompra.caption & "' "
        conn.Execute strSQL
        Do While Not .EOF
            strSQL = "insert into tab_fornpcp(Cod_Ter,Seq_Pcp,Sts_Fpc)Values("
            strSQL = strSQL & "'" & !cod_For & "', "
            strSQL = strSQL & "'" & lblCompra.caption & "',"
            strSQL = strSQL & "'N')"
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    P_buscaForn
    'stbCompras.Tab = 1
End Sub
Sub P_buscaForn()
'    strSQL = "select tt.* from tab_fornecedores as tt, tab_fornpcp as tf "
'    strSQL = strSQL & " where tt.cod_for=tf.cod_ter"
'    strSQL = strSQL & " and tf.seq_pcp='" & lblCompra.caption & "' "
'    strSQL = strSQL & " and tf.sts_fpc='N' "
'    strSQL = strSQL & " Order by tt.nome_for "
'    dtForn.RecordSource = strSQL
'    dtForn.Refresh
'    If dtForn.Recordset.RecordCount = 0 Then
'        MsgBox "Não há fornecedores a serem listados!", vbCritical
'    Else
'       cmndbtnFinaliza.Enabled = True
'    End If
'    strSQL = "select tt.* from tab_fornecedores as tt, tab_fornpcp as tf "
'    strSQL = strSQL & " where tt.cod_for=tf.cod_ter "
'    strSQL = strSQL & " and tf.seq_pcp='" & lblCompra.caption & "' "
'    strSQL = strSQL & " and tf.sts_fpc='S' "
'    strSQL = strSQL & " Order by tt.nome_for "
'    dtSelect.RecordSource = strSQL
'    dtSelect.Refresh
End Sub

Private Sub cmndbtnImportar_Click()
    
    
End Sub

Private Sub cmndbtnImprCot_Click()
    Dim iQtd_pag As Integer
    P_ImprPcp lblCompra.caption, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprPcp lblCompra.caption, Printer
        Printer.EndDoc
    End If
    Unload frmVisualiza
    
    Exit Sub
    
    
    
    
End Sub

Private Sub cmndbtnImprList_Click()
    ApontarPrinterToPDF "procomp" & Me.tag & ".pdf"
    
    imprimeListagem Printer
    Printer.EndDoc
    
End Sub
Sub imprimeListagem(obj As Object)
    If ado.RecordCount = 0 Then
        MsgBox "Não há mercadorias para serem listadas!", vbExclamation
        Exit Sub
    End If
    Dim i As Integer
    grade.Visible = False
    
    P_CabRelTimbr "Listagem de Mercadorias a Adquirir", obj
    With ado
        .MoveFirst
        P_ImprX obj, "Código", 0
        P_ImprX obj, "Marca", 20
        P_ImprX obj, "Descrição", 60
        'P_ImprX "L", 175
        P_ImprX obj, "Emb.", 150
        P_ImprX obj, "Estoque", 160
        P_ImprX obj, "Similares", 175
        P_ImprX obj, "A Comprar", 190, False, True
        Do While Not .EOF
            If TypeName(obj) = "Printer" Then
                If Printer.CurrentY >= Printer.ScaleHeight - 30 Then
                    If TypeName(obj) = "Printer" Then
                        P_Rodape strEmp_Nome & "-" & strEmp_Cidade & "/" & strEmp_UF
                    Else
                        P_RodapeVis obj, strEmp_Nome & "-" & strEmp_Cidade & "/" & strEmp_UF
                    End If
                End If
            End If
            If Left(sDataBase, 7) = "estoque" Then
                P_ImprX obj, !codfab_pro, 0
            Else
                P_ImprX obj, !cod_pro, 0
            End If
            
            
            P_ImprX obj, TiraNull(!nome_mar), 20
            
            P_ImprX obj, Left(StrConv(!nome_pro, vbProperCase), 50), 60
            P_ImprX obj, !qtdembcompra_pro, 150
            P_ImprX obj, FormatVal(!er_pro), 160, True, True
            P_ImprX obj, FormatVal(lblSimilar.caption), 180, True, True
            P_ImprX obj, FormatVal(!qfa_ipc), 195, False, True
            
            
            
            .MoveNext
            i = i + 1
            If i > 30 Then
                i = 0
                frmVisualizadorImpressao.addNewPage
            End If
            
            
            
        Loop
    End With
    
    If TypeName(obj) = "Printer" Then
        P_FimRel obj, strEmp_Nome & "-" & strEmp_Cidade & "/" & strEmp_UF
        Printer.EndDoc
    End If
    grade.Visible = True
End Sub
Private Sub cmndbtnInc_Mer_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If txtbxInc_Mer.tag = "" Then
        txtbxInc_Mer_DblClick
    End If
    If MsgBox("Confirma inclusão da mercadoria: " & Trim(txtbxInc_Mer.Text) & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    strSQL = "Insert into tab_itenspcp(seq_pcp,cod_pro,qfa_ipc,"
    strSQL = strSQL & "sts_ipc,upc_ipc)Values("
    strSQL = strSQL & "'" & lblCompra.caption & "',"
    strSQL = strSQL & "'" & txtbxInc_Mer.tag & "',"
    strSQL = strSQL & "0,"
    strSQL = strSQL & "'N'," 'não selecionado
    strSQL = strSQL & str(F_UltPreco_Mer(txtbxInc_Mer.tag, lblMarca_mer.tag)) & ")"
    
    conn.Execute strSQL
    txtbxInc_Mer.Text = ""
    txtbxInc_Mer.tag = ""
    lblMarca_mer.caption = ""
    lblMarca_mer.tag = ""
    P_BuscaItensCompra lblCompra.caption
Exit Sub

trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnInicia_Click()
     
        grade.Visible = False
        P_InsereProc
        
        cmndbtnInicia.Enabled = False
        cmndbtnList.Visible = False
        cmndbtnBuscar.Visible = True
        cmndbtnList.Visible = False
        cmndbtnFinaliza.Enabled = True
        txtbxInc_Mer.Enabled = True
        cmndbtnInc_Mer.Enabled = True
        cmndbtnImprList.Enabled = True
        cmndbtnCancela.Enabled = True
        'cmndbtnImportar.Enabled = True
        'cmndbtnExport.Enabled = True
        grade.Visible = True
     
End Sub

Private Sub cmndbtnList_Click()
    Dim strSQL As String
    Dim iNDC As Integer 'numero de dias para consumir
    
'    If cmbxFor.tag = "" Then
 '       MsgBox "Informe primeiro o fornecedor!", vbCritical
'        cmbxFor.SetFocus
  '      Exit Sub
 '   End If
    
    
    iNDC = CInt(txtbxDuracao.Text) + Int(txtbxEntrega.Text) + Int(txtbxSustent.Text)
    'frmMediaConsumo_Mer.cmbxForn.ListIndex = cmbxFor.ListIndex
    frmMediaConsumo_Mer.Show 1
    strSQL = "Select tp.*,tc.nome_cat,tm.nome_mar,"
    strSQL = strSQL & " ((tp.MMD_PRO *" & iNDC & ")-tp.er_pro) as qtd from "
    strSQL = strSQL & " " & TAB_PRODUTOS & " as tp inner join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar  "
    strSQL = strSQL & "  inner join tab_categorias as tc on tc.cod_cat=tp.cod_cat "
    'If cmbxCat.ListIndex <> -1 Then
    '    strSQL = strSQL & " and tm.cod_cat='" & cmbxCat.Tag & "' "
    'End If
    strSQL = strSQL & " where tp.er_pro-tp.er_pro-(tp.MMD_PRO *" & iNDC & ") < 0"
    If MsgBox("Deseja listar os itens em que o estoque seja maior que a sugestão?", _
        vbYesNo + vbQuestion) = vbNo Then
        strSQL = strSQL & " having qtd > 0 "
    End If
    abreObjetoAdo ado, strSQL, grade
    
    If ado.RecordCount > 0 Then
        cmndbtnInicia.Enabled = True
    Else
        MsgBox "Não foram encontrados itens com condições de compra!" & Chr(13) & _
                "Reinicie o processo de cálculo de média das mercadorias!", vbExclamation
        cmndbtnInicia.Enabled = False
    End If
End Sub



Private Sub cmndbtnRemFor_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If dtSelect.Recordset.RecordCount = 0 Then
        MsgBox "Não há fornecedores Selecionados!", vbCritical
        Exit Sub
    End If
    With dtSelect.Recordset
        strSQL = "update tab_fornpcp set sts_fpc='N' "
        strSQL = strSQL & " where cod_ter = '" & !cod_ter & "' "
        strSQL = strSQL & " and seq_pcp ='" & lblCompra.caption & "' "
        conn.Execute strSQL
    End With
    P_buscaForn
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub



Private Sub Grade_DblClick()
    If cmndbtnList.Visible = True Then
        MsgBox "O processo ainda não foi inicializado!", vbCritical
        Exit Sub
    End If
 '   frmAltProPCP.cmbxForn.ListIndex = cmbxFor.ListIndex
    frmAltProPCP.frAltMer.Visible = True
    With ado
        frmAltProPCP.caption = "Incluir/Alterar " & !nome_pro & " - Processo: " & Me.tag
        frmAltProPCP.tag = Me.tag
        frmAltProPCP.lblNome_Mer.tag = !cod_pro
        frmAltProPCP.txtbxQFA_Ipc.Text = FormatVal(!qfa_ipc)
        frmAltProPCP.txtbxTotal.Text = FormatVal(!total)
        frmAltProPCP.lblNome_Mer.caption = !nome_pro
        frmAltProPCP.id = !id
        frmAltProPCP.Show 1
        buscaFor
    End With
    Unload frmAltProPCP
    grade.Visible = False
    RefreshAdoPosition ado
    soma
    grade.Visible = True
    grade.SetFocus
End Sub
Sub soma()
    
    Dim cTotal As Currency
    Dim iLocal As Integer
    Dim iItem As Integer
    With ado
        grade.Visible = False
        iLocal = .AbsolutePosition
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            If !sts_ipc = "S" Then
                cTotal = cTotal + !total
                iItem = iItem + 1
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
            .Move iLocal
        End If
        lblTotalPrev.caption = FormatVal(cTotal)
        lblListados.caption = .RecordCount
        lblSelec.caption = iItem
    End With
    grade.Visible = True
End Sub
Private Sub dbgrdCompras_GotFocus()
    
    dbgrdCompras.MarqueeStyle = 3
End Sub

Sub P_SalvaComp()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "Update tab_procomp set cod_cat ='" & cmbxCat.tag & "', "
    
    strSQL = strSQL & " cod_mar='" & cmbxMarca.tag & "', "
    strSQL = strSQL & " nome_pro='" & txtbxNome_Mer.Text & "' "
    strSQL = strSQL & " where seq_pcp='" & lblCompra.caption & "'"
    conn.Execute strSQL
    strSQL = "Update tab_ItensPcp set sts_ipc='N' cod_cat <> '"
    strSQL = strSQL & cmbxCat.tag & "' "
    
    strSQL = strSQL & " and cod_mar <> '" & cmbxMarca.tag & "' "
    'conn.execute strSQL rever isto
    
End Sub

Private Sub grade_HeadClick(ByVal ColIndex As Integer)
    Dim s As String
    s = grade.Columns(ColIndex).DataField
          ado.Sort = s
        
    End Sub

Private Sub grade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Grade_DblClick
    End If
End Sub



Private Sub dbgrdCompras_LostFocus()
    dbgrdCompras.MarqueeStyle = 0
End Sub


Private Sub Form_Activate()
    'MudaAparencia Me
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        Carrega
    End If
    bActive = True
End Sub
Sub verificaProduto(cod_pro As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ti.*,tp.sts_pcp,tp.dti_pcp from tab_itenspcp as ti,tab_procomp as tp "
    strSQL = strSQL & " where tp.seq_pcp=ti.seq_pcp "
    strSQL = strSQL & " and cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " and sts_ipc='S'"
    strSQL = strSQL & " and sts_pcp<3 "
    strSQL = strSQL & " and tp.seq_pcp<>'" & Me.tag & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            strSQL = "Este produto encontra-se marcado como já pedido " & Chr(13)
            strSQL = strSQL & "Processo Nº: " & !seq_pcp & " de: " & !dti_pcp & Chr(13)
            strSQL = strSQL & "Quantidade: " & !qfa_ipc & " Situacao do Processo: " & !sts_pcp
            MsgBox strSQL, vbExclamation
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Ajuda
    End If
End Sub
Sub Ajuda()
    Dim strSQL As String
    strSQL = "Sujestão de Compra é o processo que tem o objetivo de ajudar o Empresário a "
    strSQL = strSQL & "definir quais os produtos devem ser comprados para que se tenha "
    strSQL = strSQL & "um melhor giro, e conseqüentemente mais lucratividade."
    strSQL = strSQL & Chr(13) & Chr(13)
    strSQL = strSQL & "É composto de três etapas: " & Chr(13)
    strSQL = strSQL & " 1 - cálculo do MMD(Média de Movimentação Diária)" & Chr(13)
    strSQL = strSQL & " 2 - preenchimento periodicidade" & Chr(13)
    strSQL = strSQL & " 3 - filtragem dos produtos " & Chr(13)
    strSQL = strSQL & " 4 - seleção e quantização dos itens a serem comprados" & Chr(13)
    strSQL = strSQL & " 5 - impressão da cotação/processo" & Chr(13)
    strSQL = strSQL & " 6 - finalização " & Chr(13)
    strSQL = strSQL & " 7 - entrada de produtos " & Chr(13)
    MsgBox strSQL, vbInformation, "PCIP - Processo de Compra de Produtos Informatizada"
End Sub
Private Sub Form_Load()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    bActive = False
    
    Set ado = New ADODB.Recordset

    resizeForm Me
        
    
    
    ajustaEstoque

End Sub
Sub buscaFor()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select distinct nome_for from " & TAB_FORNECEDORES & " as tf, tab_itenspcp as ti "
    strSQL = strSQL & " where ti.cod_for=tf.cod_for "
    strSQL = strSQL & " and ti.seq_pcp='" & Me.tag & "' "
    strSQL = strSQL & " order by nome_for "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
    cmbxFor.Clear
        cmbxFor.AddItem "(Todos)"
        Do While Not .EOF
            cmbxFor.AddItem !nome_For
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Sub Carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select tp.*   from tab_procomp as tp "
    
    strSQL = strSQL & " where seq_pcp='" & Me.tag & "'  "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            P_Status "Buscando Compra em aberto..."
            Me.tag = !seq_pcp
            Me.caption = "Processo de Compra Nº: " & Me.tag
            lblCompra.caption = !seq_pcp
            txtbxSustent.Text = !pse_pcp
            txtbxEntrega.Text = !pec_pcp
            txtbxDuracao.Text = !ppc_pcp
            'P_LocalizaOperacao !nome_for, cmbxFor
            P_BuscaItensCompra !seq_pcp
            
            cmndbtnList.Visible = False
            cmndbtnBuscar.Visible = True
            If ado.RecordCount > 0 Then
                cmndbtnInicia.Enabled = False
            End If
            cmndbtnFinaliza.Enabled = True
            txtbxInc_Mer.Enabled = True
            cmndbtnInc_Mer.Enabled = True
            cmndbtnImprList.Enabled = True
            cmndbtnCancela.Enabled = True
        Else
            cmndbtnFinaliza.Enabled = False
            txtbxInc_Mer.Enabled = False
            cmndbtnInc_Mer.Enabled = False
            .Close
        End If
    End With
    buscaFor
    P_Status ""
End Sub
Sub P_InsereProc()
    Dim strSQL As String
    
    Dim sSeq_Pcp As String
    sSeq_Pcp = F_BuscaUmCampo("max(SEQ_PCP)", "TAB_PROCOMP", "")
    If sSeq_Pcp = "" Then
        sSeq_Pcp = "000001"
    Else
        sSeq_Pcp = Format(sSeq_Pcp + 1, "0#####")
    End If
    strSQL = "insert tab_procomp(SEQ_PCP,DTI_PCP,DTF_PCP,STS_PCP,"
    strSQL = strSQL & "PSE_PCP,PEC_PCP,PPC_PCP,cod_For,COD_USU)Values("
    strSQL = strSQL & "'" & sSeq_Pcp & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & F_GravaData(DateAdd("D", CDbl(txtbxDuracao.Text), Date)) & "," ' data prevista para término
    strSQL = strSQL & "'1',"
    strSQL = strSQL & str(txtbxSustent.Text) & ","
    strSQL = strSQL & str(txtbxEntrega.Text) & ","
    strSQL = strSQL & str(txtbxDuracao.Text) & ", "
    strSQL = strSQL & "'" & txtForn.tag & "',"
    strSQL = strSQL & "'" & strCod_Usu & "')"
    conn.Execute strSQL
    P_InsereItens sSeq_Pcp
    lblCompra.caption = sSeq_Pcp
    Me.tag = "Processo de Compra - Processo Nº" & sSeq_Pcp
    Me.tag = sSeq_Pcp
    
End Sub

Sub P_InsereItens(sSeq_Pcp As String)
    Dim strSQL As String
    With ado
        grade.Visible = False
        Do While Not .EOF
            strSQL = "Insert into tab_itenspcp(seq_pcp,cod_pro,qfa_ipc,"
            strSQL = strSQL & "sts_ipc,suj_ipc,upc_ipc)Values("
            strSQL = strSQL & "'" & sSeq_Pcp & "',"
            strSQL = strSQL & "'" & !cod_pro & "',"
            strSQL = strSQL & str(Round(!qtd)) & ","
            strSQL = strSQL & "'N'," 'não selecionado
            strSQL = strSQL & F_GravaDec(!qtd) & ","
            strSQL = strSQL & str(!cr_pro) & ")"

            conn.Execute strSQL
            
            P_Status "Inserindo Itens de Compra: " & Round((.AbsolutePosition + 1) * 100 / .RecordCount, 2) & " %"
            .MoveNext
        Loop
        P_BuscaItensCompra sSeq_Pcp
    End With
    P_Status ""
    grade.Visible = True
End Sub
Sub P_BuscaItensCompra(sSeq_Pcp As String, Optional sOrderBy As String)
        Dim strSQL As String
        Dim cTotal As Currency
        Dim cMax As Currency
        Dim cSelec As Currency
        On Error GoTo trata_Erro
        strSQL = "select sum(p.er_pro) as est,ori.tipo as origem, p.reservado_pro,tp.codfab_pro,tp.qtdembcompra_pro,"
        strSQL = strSQL & " tp.nome_pro,tf.nome_for,tp.mmd_pro,tm.nome_mar,tp.endereco_pro,tp.pv_pro,tt.nome_cat,"
        strSQL = strSQL & " tp.er_pro,ti.*,ti.suj_ipc as qtd,"
        strSQL = strSQL & "(ti.upc_ipc*ti.qfa_ipc) as total "
        
        strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
        strSQL = strSQL & " left join tab_categorias as tt on tt.cod_cat=tp.cod_cat   "
        strSQL = strSQL & " left join sped_origem as ori on ori.codigo=tp.cod_ori"
        strSQL = strSQL & " inner join tab_itenspcp as ti on ti.cod_pro=tp.cod_pro "
        strSQL = strSQL & " left join " & TAB_FORNECEDORES & " as tf on ti.cod_for=tf.cod_for"
        strSQL = strSQL & " left join tab_similares as ts on tp.cod_pro=ts.cod_pro or tp.cod_pro=ts.cod_sim "
        strSQL = strSQL & " inner join produtos as p on ti.cod_pro=p.cod_pro "
        
        
        strSQL = strSQL & " where 1>0 "
        
        
        If txtForn.Text <> "" Then
            strSQL = strSQL & " and nome_for like '" & txtForn.Text & "%' "
        End If
                
        If txtFiltrar.Text <> "" Then
            strSQL = strSQL & " and nome_pro like '" & txtFiltrar.Text & "%' "
        End If
        If txtCategoria.Text <> "" Then
            strSQL = strSQL & " and nome_cat like '" & txtCat.Text & "%' "
        End If
        If txtMarca.Text <> "" Then
            strSQL = strSQL & " and nome_mar like '" & txtMarca.Text & "%' "
        End If
                
        strSQL = strSQL & " and ti.seq_pcp = '" & sSeq_Pcp & "' "
        strSQL = strSQL & " group by tp.cod_pro "
        'If cmbxCat.ListIndex > -1 Then
        '    strSQL = strSQL & " And tp.cod_cat='" & cmbxCat.Tag & "' "
        'End If
        
        'If cmbxMarca.ListIndex > -1 Then
        '    strSQL = strSQL & " And tp.cod_mar='" & cmbxMarca.Tag & "' "
        'End If
        'strSQL = strSQL & " and tp.nome_pro like '" & txtbxNome_Mer.Text & "%' "
        strSQL = strSQL & " order by "
        
        
        If sOrderBy = "" Then
        
            If optnrOrderBy(0).Value = True Then
                strSQL = strSQL & " tp.codfab_pro,tp.Cod_pro"
            ElseIf optnrOrderBy(2).Value = True Then
                strSQL = strSQL & "suj_ipc desc,tp.nome_pro"
            Else
                strSQL = strSQL & "tp.nome_pro,tp.Cod_pro"
            End If
        Else
            strSQL = strSQL & " order by " & sOrderBy
        End If
        abreObjetoAdo ado, strSQL, grade
        
        cTotal = 0
        With ado
            Do While Not .EOF
            grade.Visible = False
                If !sts_ipc = "S" Then
                    cTotal = cTotal + !total
                        
                    cSelec = cSelec + 1
                End If
                cMax = cMax + !total
                
                .MoveNext
            Loop
            lblTotalPrev.caption = FormatVal(cTotal)
            lblMax.caption = FormatVal(cMax)
            lblListados.caption = .RecordCount
            lblSelec.caption = cSelec
            If .RecordCount > 0 Then
                .MoveFirst
                grade.Visible = True
            End If
            grade.caption = .RecordCount & " Itens listados"
        End With
        
trata_Erro:
End Sub
Function F_UltPreco_Mer(sCod_Mer As String, sMarca_Mer As String) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sData As String
    Dim cValor As Currency
    Dim r2 As New ADODB.Recordset
    On Error GoTo trata_Erro
    
        strSQL = "Select ti.unit_ite as ult "
        strSQL = strSQL & " from tab_cmp as te, "
        strSQL = strSQL & " tab_itcmp as ti where ti.Seq_not=te.Seq_not and "
        strSQL = strSQL & " ti.cod_pro='" & sCod_Mer & "' "
        
        strSQL = strSQL & " ORDER by entrada_not desc limit 1"
        abreObjetoAdo r2, strSQL
    
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        
        F_UltPreco_Mer = r2!ult
    Else
        F_UltPreco_Mer = 0
    End If
    
    
    
    rPesq.Close
    Exit Function
trata_Erro:

    F_UltPreco_Mer = 0
End Function

Private Sub frAltMer_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If State <> 1 Then
    'frAltMer.Top = Y
    'frAltMer.Left = X
    End If
End Sub

Private Sub Form_Resize()
    PosicionaGrade grade, Me, pcBot
End Sub

Private Sub grade_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
        
    End If
End Sub

Private Sub mnAjust_Click()
    Grade_DblClick
End Sub

Private Sub mnEditProd_Click()
    On Error GoTo trata_Erro
    frmCadPro.tag = ado!cod_pro
    frmCadPro.Show 1
    Exit Sub
trata_Erro:
    
End Sub

Private Sub optnrOrderBy_Click(Index As Integer)
    If lblCompra.caption <> "" Then
        P_Status "Buscando itens..."
        P_BuscaItensCompra lblCompra.caption
        P_Status ""
    End If
End Sub

Private Sub stbCompras_Click(PreviousTab As Integer)
'    If stbCompras.Tab = 0 Then
'        cmndbtnImprCot.Visible = False
'        frPeriod.Visible = True
'    Else
'        cmndbtnImprCot.Visible = True
'        frPeriod.Visible = False
'    End If
End Sub

Private Sub stbCompras_KeyDown(KeyCode As Integer, Shift As Integer)
    frAltMer.Visible = False
    
End Sub

Private Sub txtbxDuracao_GotFocus()
    ganhaFoco txtbxDuracao
End Sub

Private Sub txtbxDuracao_LostFocus()
        PerdeFoco txtbxDuracao, "#0"
        txtbxTotalDias.Text = Int(txtbxSustent.Text) + Int(txtbxDuracao.Text) + Int(txtbxEntrega.Text)
End Sub

Private Sub txtbxEntrega_GotFocus()
    ganhaFoco txtbxEntrega
End Sub

Private Sub txtbxEntrega_LostFocus()
    PerdeFoco txtbxEntrega, "#0"
    txtbxTotalDias.Text = Int(txtbxSustent.Text) + Int(txtbxDuracao.Text) + Int(txtbxEntrega.Text)
End Sub

Private Sub txtbxInc_Mer_Change()
    txtbxInc_Mer.tag = ""
    lblMarca_mer.caption = ""
End Sub

Private Sub txtbxInc_Mer_DblClick()
    Dim p As Produto
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    If txtbxInc_Mer.Text = "" Then
        MsgBox "Informe um nome para o produto que deseja incluir!", vbExclamation
        Exit Sub
        Exit Sub
    End If
    p = F_BuscaProdSaida(txtbxInc_Mer.Text)
    
    If p.cod_pro = "" Then
        Exit Sub
    End If
    txtbxInc_Mer.Text = p.nome_pro
    txtbxInc_Mer.tag = p.cod_pro
    lblMarca_mer.tag = p.cod_mar
    lblMarca_mer.caption = p.nome_mar
    strSQL = "select * from tab_itenspcp where seq_pcp ='" & lblCompra.caption & "' "
    strSQL = strSQL & " and cod_pro = '" & p.cod_pro & "' "
    
    rPesq.Open strSQL, conn_client
    If rPesq.RecordCount > 0 Then
        MsgBox "Este Produto já está relacionado nessa compra!", vbExclamation
        rPesq.Close
'        txtbxInc_Mer.tag = ""
 '       txtbxInc_Mer.SetFocus
    Else
    
    
        strSQL = "Insert into tab_itenspcp(seq_pcp,cod_pro,qfa_ipc,"
            strSQL = strSQL & "sts_ipc,suj_ipc,upc_ipc)Values("
            strSQL = strSQL & "'" & lblCompra.caption & "',"
            strSQL = strSQL & "'" & p.cod_pro & "',"
            strSQL = strSQL & "0,"
            strSQL = strSQL & "'N'," 'não selecionado
            strSQL = strSQL & "0,"
            strSQL = strSQL & F_GravaDec(F_BuscaUltimoPrecoCompra(p.cod_pro)) & ")"

            conn.Execute strSQL
        
        rPesq.Close
        cmndbtnBuscar_Click
    End If
  '  cmndbtnInc_Mer_Click
End Sub

Private Sub txtbxInc_Mer_GotFocus()
    ganhaFoco txtbxInc_Mer
End Sub

Private Sub txtbxInc_Mer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxInc_Mer_DblClick
    End If
End Sub

Private Sub txtbxInc_Mer_LostFocus()
    AjustaTexto txtbxInc_Mer
End Sub

Private Sub txtbxNome_Mer_GotFocus()
    
End Sub

Private Sub txtbxNome_Mer_LostFocus()
    
    P_SalvaComp
    P_BuscaItensCompra lblCompra.caption
End Sub


Private Sub txtbxSustent_GotFocus()
    ganhaFoco txtbxSustent
End Sub

Private Sub txtbxSustent_LostFocus()
    PerdeFoco txtbxSustent, "#0"
    txtbxTotalDias.Text = Int(txtbxSustent.Text) + Int(txtbxDuracao.Text) + Int(txtbxEntrega.Text)
End Sub

Private Sub txtCategoria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        cmndbtnBuscar_Click
    End If
    
End Sub

Private Sub txtFiltrar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AjustaTexto txtFiltrar
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
End Sub
