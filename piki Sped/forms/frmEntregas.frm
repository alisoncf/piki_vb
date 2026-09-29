VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEntregas 
   Caption         =   "Entregas"
   ClientHeight    =   9405
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15165
   Icon            =   "frmEntregas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   15165
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   15135
      TabIndex        =   4
      Top             =   8430
      Width           =   15165
      Begin VB.PictureBox pcControles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   12945
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   12975
         Begin VB.CommandButton cmdMudaEntregador 
            Caption         =   "Mudar Entregador"
            Height          =   375
            Left            =   6000
            TabIndex        =   24
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdCancelarEntrega 
            Caption         =   "Cancelar Entrega"
            Height          =   375
            Left            =   4440
            TabIndex        =   23
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdImprimirVenda 
            Caption         =   "Imprimir Venda"
            Height          =   375
            Left            =   7560
            TabIndex        =   21
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdGerarEntrega 
            Caption         =   "Gerar Entrega"
            Height          =   375
            Left            =   1200
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdConcluirEntrega 
            Caption         =   "Concluir Entrega"
            Height          =   375
            Left            =   2760
            TabIndex        =   19
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdPreparar 
            Caption         =   "Preparar"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   13320
         Picture         =   "frmEntregas.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   8145
      ScaleWidth      =   14505
      TabIndex        =   2
      Top             =   0
      Width           =   14535
      Begin VB.PictureBox pcDadosCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   8745
         TabIndex        =   38
         Top             =   7200
         Width           =   8775
         Begin VB.Label lblEndereco 
            Caption         =   "..."
            Height          =   735
            Left            =   1080
            TabIndex        =   40
            Top             =   0
            Width           =   5055
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Endereço:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2640
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
         Begin VB.TextBox txtEntregador 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdEntregador 
            Height          =   285
            Left            =   1200
            Picture         =   "frmEntregas.frx":1458
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entregador"
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
            Left            =   120
            TabIndex        =   35
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4320
         ScaleHeight     =   705
         ScaleWidth      =   1545
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
         Begin VB.TextBox txtPesquisar 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
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
            Left            =   120
            TabIndex        =   31
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.PictureBox pcLote 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   8880
         ScaleHeight     =   1695
         ScaleWidth      =   5415
         TabIndex        =   27
         Top             =   5760
         Visible         =   0   'False
         Width           =   5415
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridLote 
            Height          =   1335
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   2355
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblLote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega em lote"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.PictureBox pcPesq 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   3945
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   3975
         Begin VB.OptionButton optnVenda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Venda"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1080
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optnEntregador 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Entregador"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2280
            TabIndex        =   25
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordem:"
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
            Left            =   120
            TabIndex        =   1
            Top             =   0
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid gridAPreparar 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   5760
         Visible         =   0   'False
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "Pedidos a Preparar"
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "num_ven"
            Caption         =   "Venda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "nomecli_ven"
            Caption         =   "Cliente"
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
            DataField       =   "obs_ven"
            Caption         =   "Obs"
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
            DataField       =   "Vendedor"
            Caption         =   "Vendedor"
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
            DataField       =   "guerra_fun"
            Caption         =   "Preparado por"
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
            DataField       =   "guerra_fun"
            Caption         =   "Entregador"
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
            DataField       =   "data_ven"
            Caption         =   "Hora Pedido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "hora_enc"
            Caption         =   "Hora Envio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "hora_conf"
            Caption         =   "Retorno"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "tempo"
            Caption         =   "Tempo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "obs_entrega"
            Caption         =   "Obs Entrega"
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
            DataField       =   "dataefet_ven"
            Caption         =   "Data Efet"
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
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1574,929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1244,976
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdConcluidas 
         BackColor       =   &H00FFC0C0&
         Caption         =   "F5 - Concluídas"
         Height          =   735
         Left            =   12000
         Picture         =   "frmEntregas.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid gridProdutos 
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   5760
         Visible         =   0   'False
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Produtos do Pedido Selecionado"
         ColumnCount     =   6
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "codfab_pro"
            Caption         =   "Cód.Fábrica"
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
         BeginProperty Column03 
            DataField       =   "desc_ite"
            Caption         =   "Produto"
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
            DataField       =   "endereco_pro"
            Caption         =   "Endereço"
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
            DataField       =   "qtd_ite"
            Caption         =   "Qtd"
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
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2880
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   4410,142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   824,882
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdSemEntrega 
         Caption         =   "F6 - Sem Entrega"
         Height          =   735
         Left            =   13800
         Picture         =   "frmEntregas.frx":1CFE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
      Begin VB.PictureBox pcReferencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   705
         ScaleWidth      =   2385
         TabIndex        =   9
         Top             =   120
         Width           =   2415
         Begin VB.TextBox txtAte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtDe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Até:"
            Height          =   195
            Left            =   1200
            TabIndex        =   12
            Top             =   0
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdEntregaOk 
         BackColor       =   &H00B0EAA4&
         Caption         =   "F4 - Em Entrega"
         Height          =   735
         Left            =   10200
         Picture         =   "frmEntregas.frx":21D3
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEntregar 
         BackColor       =   &H00FCB683&
         Caption         =   "F3 - A Entregar"
         Height          =   735
         Left            =   8400
         Picture         =   "frmEntregas.frx":26BF
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdPrep 
         BackColor       =   &H00EDF8FE&
         Caption         =   "F2 - A Preparar"
         Height          =   735
         Left            =   6600
         MaskColor       =   &H00FFFFC0&
         Picture         =   "frmEntregas.frx":44B5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
         Height          =   4455
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         FocusRect       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblTituloMsh 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Registros encontrados"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   14175
      End
   End
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   6720
      Top             =   720
   End
   Begin VB.Menu mnMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnPreparar 
         Caption         =   "Preparar"
      End
      Begin VB.Menu mnEntregar 
         Caption         =   "Gerar Entrega (Envio)"
      End
      Begin VB.Menu mnRetornoEntrega 
         Caption         =   "Concluir Entrega (Retorno)"
      End
      Begin VB.Menu mnCancelarEntrega 
         Caption         =   "&Cancelar Entrega"
      End
      Begin VB.Menu mnMudarEntrega 
         Caption         =   "Mudar Entregador"
      End
      Begin VB.Menu mnExport 
         Caption         =   "Exportar Grade"
      End
      Begin VB.Menu mnExcluirEntregaDoLote 
         Caption         =   "Excluir Entrega do Lote"
      End
      Begin VB.Menu mnSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmEntregas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim WithEvents adoProd As ADODB.Recordset
Attribute adoProd.VB_VarHelpID = -1
Dim bActive As Boolean
Dim status_ped As String
Sub buscaProd(venda As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select ti.*,codfab_pro,endereco_pro,nome_mar from " & TAB_ITVENDAS & "  as ti inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & TAB_MARCAS & " as tm on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " where num_ven='" & venda & "' "
    strSQL = strSQL & " order by ti.seq_ite "
    abreObjetoAdo adoProd, strSQL, gridProdutos
    gridProdutos.Visible = True
    Exit Sub
    
    
trata_Erro:
   ' P_MostraErro
    
    
End Sub

Private Sub adoProd_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    HabControle
End Sub
Sub HabControle()
    If ado.RecordCount > 0 Then
        pcControles.Visible = True
    Else
        pcControles.Visible = False
    End If
End Sub

Private Sub cmdCancelarEntrega_Click()
    cancelarentrega
    cmdEntregaOk_Click
End Sub

Private Sub cmdConcluidas_Click()
    buscaPedidos 4
    pcTop.BackColor = cmdConcluidas.BackColor
    msh.BackColor = cmdConcluidas.BackColor
End Sub

Private Sub cmdConcluirEntrega_Click()
    concluientrega
End Sub
Sub concluientrega()
        On Error GoTo trata_Erro
    Dim lote As String
    Dim venda As String
    msh.col = 13
    If msh.Text <> "3" Then
        MsgBox "Esta venda não está apta para ser concluída!", vbExclamation
        Exit Sub
    End If
    
    msh.col = 15
    lote = msh.Text
    msh.col = 0
    venda = msh.Text
    
    
    If lote <> "" Then
    
        frmEntregaEmLote.tag = lote
        frmEntregaEmLote.txtLote.Text = lote
        frmEntregaEmLote.status = "3"
        msh.col = 7
        frmEntregaEmLote.txtHora.Text = Format(msh.Text, "hh:mm:ss")
        frmEntregaEmLote.txtHora.Enabled = False
        msh.col = 5
        frmEntregaEmLote.txtEntregador.Text = TiraNull(msh.Text)
        msh.col = 13
        frmEntregaEmLote.txtEntregador.tag = TiraNull(msh.Text)
        frmEntregaEmLote.txtRetorno.Text = Time
        
        frmEntregaEmLote.Show 1
    Else
        frmEntrega.tag = venda
        frmEntrega.txtVenda.Text = venda
        frmEntrega.status = "3"
        msh.col = 7
        frmEntrega.txtHora.Text = Format(msh.Text, "hh:mm:ss")
        frmEntrega.txtHora.Enabled = False
        msh.col = 6
        frmEntrega.txtEntregador.Text = TiraNull(msh.Text)
        msh.col = 13
        frmEntrega.txtEntregador.tag = TiraNull(msh.Text)
        frmEntrega.txtRetorno.Text = Time
        
        frmEntrega.Show 1
    End If
    cmdEntregaOk_Click
'    RefreshAdoPosition ado
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro

End Sub
Sub cancelarentrega()
        On Error GoTo trata_Erro
    Dim strSQL As String
    
    msh.col = 13
    If msh.Text = "4" Then
        MsgBox "Esta venda não está apta para ser cancelada!", vbExclamation
        Exit Sub
    End If
    Dim venda As String
    msh.col = 0
    venda = msh.Text
    
    If MsgBox("Confirma cancelamento da entrega?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    strSQL = "Update " & TAB_VENDAS & " set entrega_ven=0 where num_ven='" & venda & "'"
    conn.Execute strSQL
    
'    RefreshAdoPosition ado
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro

End Sub

Private Sub cmdEntregador_Click()
     Dim f As funcionario.funcionario
    On Error Resume Next
    f = funcionario.F_SelectFuncionario(txtEntregador.Text)
    If f.cod_fun <> "" Then
        txtEntregador.Text = f.guerra_fun
        txtEntregador.tag = f.cod_fun
        
        txtHora.SetFocus
    End If
    
End Sub

Private Sub cmdEntregaOk_Click()
    buscaPedidos 3
    pcTop.BackColor = cmdEntregaOk.BackColor
    msh.BackColor = pcTop.BackColor
End Sub

Private Sub cmdEntregar_Click()
    buscaPedidos "2"
    pcTop.BackColor = cmdEntregar.BackColor
    msh.BackColor = pcTop.BackColor
    
    
End Sub

Private Sub cmdGerarEntrega_Click()
    GeraEntrega
    cmdEntregar_Click
End Sub

Private Sub cmdGerarEntregaEmLote_Click()
    GeraEntregaEmLote
    cmdEntregar_Click
End Sub

Private Sub cmdImprimirVenda_Click()
    On Error GoTo trata_Erro
    
    ImprVenda ado!venda
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdMudaEntregador_Click()
    Dim f As funcionario.funcionario
    Dim venda As String
    Dim strSQL As String
    msh.col = 13
    If msh.Text = "4" Then
        MsgBox "Esta venda não está apta para ser modificada!", vbExclamation
        Exit Sub
    End If
    
    msh.col = 0
    venda = msh.Text
    
    
    
    f = f_BuscaFun("")
    If f.cod_fun <> "" Then
        If MsgBox("Confirma alteração do entregador?", vbYesNo + vbQuestion) = vbYes Then
            strSQL = "update historico_entrega set entregador='" & f.cod_fun & "' "
            strSQL = strSQL & " where num_ven='" & venda & "'"
            conn.Execute strSQL
            buscaPedidos 3
        End If
    End If
    
    
End Sub
Sub RastreiaTempo()
    Dim col As Integer
    Dim row As Integer
    Dim i As Integer
    Dim j As Integer
    With msh
        col = .col
        row = .row
        
        
        
        
    
        .col = col
        .row = row
    End With
End Sub

Private Sub cmdPrep_Click()
    buscaPedidos
    pcTop.BackColor = cmdPrep.BackColor
    msh.BackColor = pcTop.BackColor
    
End Sub
Sub BuscaVenda()
    
End Sub
Sub GeraEntrega()
    On Error GoTo trata_Erro
      Dim entrega_ven As String
    Dim venda As String
    msh.col = 0
    venda = msh.Text
    msh.col = 13
    entrega_ven = msh.Text
    If entrega_ven <> "2" Then
        MsgBox "Esta venda não está apta para ser entregue!", vbExclamation
        Exit Sub
    End If
    
    frmEntrega.tag = venda
    frmEntrega.txtVenda.Text = venda
    frmEntrega.status = "2"
    
    frmEntrega.Show 1
    ado.Requery
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub GeraEntregaEmLote()
    On Error GoTo trata_Erro
      Dim entrega_ven As String
    Dim venda As String
    With msh
    End With
    
    
    
    msh.col = 0
    venda = msh.Text
    msh.col = 13
    entrega_ven = msh.Text
    If entrega_ven <> "2" Then
        MsgBox "Esta venda não está apta para ser entregue!", vbExclamation
        Exit Sub
    End If
    
    frmEntrega.tag = venda
    frmEntrega.txtVenda.Text = venda
    frmEntrega.status = "2"
    
    frmEntrega.Show 1
    ado.Requery
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdPreparar_Click()
    Dim venda As String
    msh.col = 0
    venda = msh.Text
    frmPreparacaoConferencia.tag = venda
    frmPreparacaoConferencia.Show 1
    If frmPreparacaoConferencia.tag = "1" Then
        
        prepara
    Else
        
    End If
    Unload frmPreparacaoConferencia
End Sub

Sub prepara()
    On Error GoTo trata_Erro
    Dim entrega_ven As String
    Dim venda As String
    msh.col = 0
    venda = msh.Text
    msh.col = 13
    entrega_ven = msh.Text
    
    If entrega_ven <> "1" And entrega_ven <> "0" And entrega_ven <> "" Then
        MsgBox "Esta venda não está apta para ter o pedido preparado.", vbExclamation
        Exit Sub
    End If
    
    If PreparaEntrega(venda) Then
        MsgBox "Preparação Ok!", vbInformation
        cmdPrep_Click
        
    Else
    
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub cmdSemEntrega_Click()
    buscaPedidos 0
    pcTop.BackColor = cmdSemEntrega.BackColor
    msh.BackColor = pcTop.BackColor
End Sub

Sub ajustaGrid()
    With msh
        .Clear
        .rows = 2
        .Cols = 18
        .FixedCols = 2
        .FixedRows = 1
       ' .rows = 1
        .row = 0:
        
        .col = 0: .Text = "Venda"
        .col = 1: .Text = "Cliente"
        .col = 2: .Text = "Obs."
        .col = 3: .Text = "Vendedor"
        .col = 4: .Text = "Preparado Por"
        .col = 5: .Text = "Entregador"
        .col = 6: .Text = "Hr Pedido"
        .col = 7: .Text = "Hr Envio"
        .col = 8: .Text = "Hr Retorno"
        .col = 9: .Text = "Tp Saída"
        .col = 10: .Text = "Tp Total"
        .col = 11: .Text = "Obs Entrega"
        .col = 12: .Text = "Data Efet."
        .col = 13: .Text = "S"
        .col = 14: .Text = "Cód.Entregador"
        .col = 15: .Text = "Lote"
        .col = 16: .Text = "Fantasia"
        .col = 17: .Text = "Endereço"
        .ColWidth(0) = 900
        .ColWidth(1) = 4000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1100
        .ColWidth(9) = 1200
        .ColWidth(10) = 900
        .ColWidth(11) = 2000
        .ColWidth(12) = 1000
        .ColWidth(13) = 500
        .ColWidth(14) = 0
        .ColWidth(15) = 500
        .ColWidth(16) = 4000
        .ColWidth(17) = 4000
        .ColAlignment(3) = 2
        .ColAlignment(4) = 2
        .ColAlignment(5) = 2
        .ColAlignment(6) = 3
        .ColAlignment(7) = 3
        .ColAlignment(8) = 3
        
        .ColAlignment(9) = 3
        .ColAlignment(10) = 3
        .AllowUserResizing = flexResizeColumns
        
    End With
End Sub
Private Sub Form_Activate()
    If bActive Then Exit Sub
    cmdPrep_Click
    bActive = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyDown Then
        gridAPreparar.SetFocus
    End If
    If KeyCode = vbKeyF2 Then
        cmdPrep_Click
    ElseIf KeyCode = vbKeyF3 Then
        cmdEntregar_Click
    ElseIf KeyCode = vbKeyF4 Then
        cmdEntregaOk_Click
    ElseIf KeyCode = vbKeyF5 Then
        cmdConcluidas_Click
    ElseIf KeyCode = vbKeyF6 Then
        cmdSemEntrega_Click
    End If
    
End Sub

Private Sub Form_Load()
    bActive = False
    Set ado = New ADODB.Recordset
    Set adoProd = New ADODB.Recordset
    resizeForm Me
    ajustaGrid
    txtDe.Text = Date
    txtAte.Text = Date
End Sub

Sub buscaPedidos(Optional status As String = "1")
    Dim strSQL As String
    On Error GoTo trata_Erro
    Dim i As Integer
    Dim min As Single
    status_ped = status
    msh.Visible = False
    
    Dim TempoEsperaEntregar As Integer
    TempoEsperaEntregar = 30
    
    gridProdutos.Visible = False
    P_Status "aguarde..."
    strSQL = "select tv.num_ven as venda,endereco_ven,bairro_ven,fone_ven,dataefet_ven,right(data_ven,5) as hora_ven,tv.data_ven,tv.obs_ven,tf4.guerra_fun as vendedor,tc.nome_cli as cliente, tc.fantasia_cli as fantasia,"
    strSQL = strSQL & " tf.guerra_fun as entregador_nome,he.*,tf1.guerra_fun as prepara, tv.endereco_ven,bairro_ven,fone_ven, "
    strSQL = strSQL & " tf2.guerra_fun as encaminha, tf3.guerra_fun as confirma, left(timediff(time(hora_conf),time(hora_enc)),8) as tempo ,"
    strSQL = strSQL & " left(timediff(time(hora_prep),time(data_ven)),5) as tempot, tv.entrega_ven "
    strSQL = strSQL & " from " & TAB_VENDAS & " as tv inner join " & TAB_CLIENTES & " as tc on tv.cod_cli=tc.cod_cli  "
    strSQL = strSQL & " left join historico_entrega as he on tv.num_ven=he.num_ven "
    strSQL = strSQL & " left join tab_funcionarios as tf on tf.cod_fun=entregador "
    strSQL = strSQL & " left join tab_funcionarios as tf1 on tf1.cod_fun=preparado_por "
    strSQL = strSQL & " left join tab_funcionarios as tf2 on tf2.cod_fun=encaminhado_por"
    strSQL = strSQL & " left join tab_funcionarios as tf3 on tf3.cod_fun=confirmado_por "
    strSQL = strSQL & " left join tab_funcionarios as tf4 on tf4.cod_fun=abrevenda_usu "
    strSQL = strSQL & " where situacao_ven>1 "
    If status <> "" Then
        strSQL = strSQL & " and tv.entrega_ven='" & status & "' "
    End If
    If txtDe.Text <> "" Then
        strSQL = strSQL & " and data_ven>=" & F_GravaDataTempo(txtDe.Text & " 00:00") & " "
    End If
    If txtAte.Text <> "" Then
        strSQL = strSQL & " and data_ven<=" & F_GravaDataTempo(txtAte.Text & " 23:59") & " "
    End If
    If txtPesquisar.Text <> "" Then
        strSQL = strSQL & " and (nomecli_ven like '" & txtPesquisar.Text & "%' "
        strSQL = strSQL & " or tv.num_ven = '" & txtPesquisar.Text & "'"
        strSQL = strSQL & " or tc.fantasia_cli like '%" & txtPesquisar.Text & "%'"
        strSQL = strSQL & " or obs_entrega like '" & txtPesquisar.Text & "%'"
        strSQL = strSQL & ") "
    End If
    If txtEntregador.tag <> "" And status > 2 Then
        strSQL = strSQL & " and entregador='" & txtEntregador.tag & "' "
    End If
    If optnVenda.Value Then
        strSQL = strSQL & " order by data_ven desc "
    Else
        strSQL = strSQL & " order by entregador, data_ven desc "
    End If
    
    abreObjetoAdo ado, strSQL ', gridAPreparar
  '  abreObjetoAdo ado, strSQL, msh
    
    ajustaGrid
    With ado
        Do While Not .EOF
            strSQL = !venda & Chr(9) & !cliente & IIf(TiraNull(!fantasia) <> "", " (" & !fantasia & ")", "") & Chr(9) & !obs_ven & Chr(9) & !vendedor & Chr(9) & _
            !prepara & Chr(9) & !entregador_nome & Chr(9) & Format(!data_ven, "hh:mm") & Chr(9)
            strSQL = strSQL & Format(!hora_enc, "hh:mm") & Chr(9)
            strSQL = strSQL & Format(!hora_conf, "hh:mm") & Chr(9)
            
            If IsNull(!Tempo) Then
                If status = "3" Then
                    strSQL = strSQL & Format(DateDiff("n", !hora_enc, HoraServidor), "") & Chr(9)
                    strSQL = strSQL & "" & Chr(9)
                Else
                    strSQL = strSQL & Format(!Tempot, "") & " Min" & Chr(9)
                    strSQL = strSQL & Format(!Tempo, "") & " Min" & Chr(9)
                    
                
                End If
                
            Else
                strSQL = strSQL & DateDiff("n", !data_ven, !hora_prep) & " min." & Chr(9)
                strSQL = strSQL & DateDiff("n", !hora_enc, !hora_conf) & " min." & Chr(9)
            End If
            strSQL = strSQL & TiraNull(!obs_entrega) & Chr(9)
            strSQL = strSQL & !dataefet_ven & Chr(9)
            strSQL = strSQL & !entrega_ven & Chr(9) & !entregador & Chr(9) & TiraNull(!lote) & Chr(9) & TiraNull(!fantasia) & Chr(9) & TiraNull(!endereco_ven) & " " & TiraNull(!bairro_ven) & " " & TiraNull(!fone_ven)
            msh.AddItem strSQL, msh.rows - 1
            msh.row = msh.rows - 1
            If status = "2" Then
                msh.row = 0
                msh.col = 9
                msh.Text = "Tp Espera"
                msh.row = msh.rows - 1
                
                
                min = DateDiff("N", !data_ven, HoraServidor)
                msh.row = msh.rows - 2
                msh.Text = min & " Min."
                
                If min > TempoEsperaEntregar Then
                    msh.row = msh.rows - 2
                    For i = 0 To msh.Cols - 1
                        msh.col = i
                        msh.CellFontBold = True
                        msh.CellForeColor = vbRed
                        msh.CellFontItalic = True
                        msh.CellFontName = "Arial"
                    Next
                End If
                
                        
            ElseIf status = "3" Then
                
                msh.row = 0
                msh.col = 9
                msh.Text = "Tp Saída"
                msh.row = msh.rows - 1
                If Not IsNull(!hora_enc) Then
                    min = DateDiff("N", !data_ven, !hora_enc)
                Else
                    min = 0
                End If
                msh.row = msh.rows - 2
                msh.Text = min & " Min."
                
                
                 If min > TempoEsperaEntregar Then
                    msh.row = msh.rows - 2
                    For i = 0 To msh.Cols - 1
                        msh.col = i
                        msh.CellFontBold = True
                        msh.CellForeColor = vbRed
                        msh.CellFontItalic = True
                        msh.CellFontName = "Arial"
                    Next
                End If
            
            ElseIf status = "4" Then
            
            End If
            
            .MoveNext
        Loop
    End With
    If msh.rows > 2 Then
        msh.rows = msh.rows - 1
    End If

msh.GridLines = flexGridDots
msh.GridColor = &HD7E6D8
msh.GridLinesFixed = flexGridFlat
msh.SelectionMode = flexSelectionByRow

    ZebraGrid msh, 1
    msh.Visible = True
    P_Status ""
    lblTituloMsh.caption = ado.RecordCount & " Registros encontrados"
    
    Exit Sub
trata_Erro:
    P_MostraErro
    msh.Visible = True
    P_Status ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    'PosicionaGrade gridProdutos, Me, pcTop
    'PosicionaGrade gridAPreparar, Me, gridProdutos
    pcTop.width = Me.width - pcTop.Left - 200
    pcTop.height = pcBot.Top - pcTop.Top - 100
    gridAPreparar.width = pcTop.width - gridAPreparar.Left - 100
    msh.width = pcTop.width - msh.Left - 100
    posicionaADireita cmdSair, pcBot
    
    posicionaADireita pcLote, Me
     
    lblTituloMsh.width = msh.width
End Sub

Private Sub gridAPreparar_DblClick()
    With ado
        If !entrega_ven = "1" Then
            prepara
        ElseIf !entrega_ven = "2" Then
            GeraEntrega
        
        ElseIf !entrega_ven = "3" Then
        
            concluientrega
        End If
    End With
End Sub

Private Sub gridAPreparar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        gridAPreparar_DblClick
    End If
End Sub

Private Sub gridAPreparar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub gridLote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        MsgBox "Deletar venda do lote"
    End If
End Sub

Private Sub mnEntregar_Click()
    GeraEntrega
End Sub

Private Sub mnExcluirEntregaDoLote_Click()
    With msh
        .col = 14
        If .Text = "" Then
            MsgBox "Essa entrega não faz parte de nenhum lote!", vbExclamation
            Exit Sub
        End If
        .col = 13
        If .Text = "3" Then
            If MsgBox("Deseja realmente excluir essa entrega do lote?", vbYesNo + vbCritical) = vbYes Then
                .col = 0
                RetiraVendaEntregaLote .Text
            End If
            
        End If
        
        
    End With
End Sub

Private Sub mnExport_Click()
    ExportaMSFLXGridToCSV msh
End Sub

Private Sub mnPreparar_Click()
    prepara
End Sub

Private Sub mnRetornoEntrega_Click()
    concluientrega
End Sub

Private Sub mnSair_Click()
    cmdSair_Click
End Sub



Private Sub msh_DblClick()
    With msh
        .col = 13
        If .Text = "2" Then
            enviaLote
        ElseIf .Text = "3" Then
            concluientrega
        
        End If
    
       End With
End Sub
Sub enviaLote()
    With msh
        .col = 13
        If .Text <> "2" Then
            Exit Sub
        End If
        
        If Not IsLoad("frmEntregaEmLote") Then
            frmEntregaEmLote.Show 0, frmEntregas
            posicionaADireita frmEntregaEmLote, mdiFrm
        End If
        .col = 0
        frmEntregaEmLote.status = "2"
        frmEntregaEmLote.AdicionaVenda .Text
        
        
    End With
End Sub

Private Sub msh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        Clipboard.SetText msh.Text
    End If
    If KeyCode = vbKeyReturn Then
        msh_DblClick
    End If
End Sub

Private Sub msh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub msh_RowColChange()
    Dim st As String
    msh.col = 0
    buscaProd msh.Text
    
    msh.col = 17
    lblEndereco.caption = msh.Text
    
    buscaLote
    BuscaEndereco
    msh.col = 13
    
    st = msh.Text
    cmdPreparar.Enabled = False
    cmdGerarEntrega.Enabled = False
    cmdConcluirEntrega.Enabled = False
    
    If st = "1" Then
        cmdPreparar.Enabled = True
        
    ElseIf st = "2" Then
        cmdGerarEntrega.Enabled = True
    ElseIf st = "3" Then
        cmdConcluirEntrega.Enabled = True
    ElseIf st = "4" Then
    
    Else
    
    End If
    
    cmdPreparar.Enabled = True
    msh.col = 0
End Sub
Sub BuscaEndereco()

End Sub
Sub buscaLote()
    msh.col = 15
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    If msh.Text = "" Or msh.Text = "0" Then
        pcLote.Visible = False
        gridLote.Clear
    Else
        pcLote.Visible = True
        strSQL = "select tv.* from " & TAB_VENDAS & " as tv left join historico_entrega AS h on tv.num_ven=h.num_Ven "
        strSQL = strSQL & " where 1 > 0 "
    
    
    
        strSQL = strSQL & " and lote='" & msh.Text & "'"
        abreObjetoAdo rPesq, strSQL
        lblLote.caption = "Entrega em Lote Nº " & msh.Text
        AjustaGridData gridLote, 2, 1, 0, "Venda,Cliente", "Num_Ven,nomecli_ven", "800,3000", rPesq, "", True
        msh.col = 0
        msh.SetFocus
    End If
End Sub

Private Sub txtAte_GotFocus()
    P_SelecTexto txtAte
End Sub

Private Sub txtAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAte_LostFocus()
    P_AjustaData txtAte
End Sub
Private Sub txtDe_GotFocus()
    P_SelecTexto txtDe
End Sub

Private Sub txtDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDe_LostFocus()
    P_AjustaData txtDe
End Sub

Private Sub txtEntregador_GotFocus()
    P_SelecTexto txtEntregador
End Sub

Private Sub txtEntregador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdEntregador_Click
    End If
End Sub

Private Sub txtPesquisar_GotFocus()
    P_SelecTexto txtPesquisar
End Sub

Private Sub txtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaPedidos status_ped
    End If
    
End Sub
