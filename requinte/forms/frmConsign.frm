VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsign 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   Caption         =   "Consignação:"
   ClientHeight    =   9765
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   16785
   Icon            =   "frmConsign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   16785
   Begin VB.PictureBox pcFotoBag 
      Height          =   2295
      Left            =   15480
      ScaleHeight     =   2235
      ScaleWidth      =   1875
      TabIndex        =   81
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   2745
         TabIndex        =   84
         Tag             =   "!"
         Top             =   0
         Width           =   2775
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            ScaleHeight     =   180
            ScaleWidth      =   210
            TabIndex        =   85
            Top             =   0
            Width           =   210
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fotografia:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   360
            TabIndex        =   86
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.PictureBox pc3X4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2655
         Index           =   0
         Left            =   0
         ScaleHeight     =   2625
         ScaleWidth      =   2745
         TabIndex        =   82
         Top             =   240
         Width           =   2775
         Begin VB.PictureBox p 
            BorderStyle     =   0  'None
            Height          =   1695
            Index           =   0
            Left            =   480
            ScaleHeight     =   1695
            ScaleWidth      =   1575
            TabIndex        =   83
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox pcAtalhos 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   16755
      TabIndex        =   53
      Top             =   8370
      Width           =   16785
      Begin VB.CommandButton cmdNFedev 
         Caption         =   "NFe Devolução"
         Height          =   735
         Left            =   1560
         Picture         =   "frmConsign.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnTransfAberto 
         Caption         =   "F3 - Transf. Aberto"
         Height          =   735
         Left            =   3000
         Picture         =   "frmConsign.frx":0F3E
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEstoque 
         Caption         =   "NFe Manifesto"
         Height          =   735
         Left            =   120
         Picture         =   "frmConsign.frx":10B5
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "F4 - Nova Tranf."
         Height          =   735
         Left            =   4560
         Picture         =   "frmConsign.frx":1329
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnAcerto 
         Caption         =   "F5 - Acerto"
         Height          =   735
         Left            =   6120
         Picture         =   "frmConsign.frx":148D
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnDevolver 
         Caption         =   "F6 -Devolução"
         Height          =   735
         Left            =   7680
         Picture         =   "frmConsign.frx":1817
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "F12 - Sair"
         Height          =   735
         Left            =   15480
         Picture         =   "frmConsign.frx":1DA1
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnListaDev 
         Caption         =   "Listar Devoluções"
         Height          =   735
         Left            =   9240
         Picture         =   "frmConsign.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnConsign 
         Caption         =   "Impr. Consignação"
         Height          =   735
         Left            =   10800
         Picture         =   "frmConsign.frx":2496
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnAtualizarPrecos 
         Caption         =   "Atualizar Preços"
         Height          =   735
         Left            =   12360
         Picture         =   "frmConsign.frx":2592
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprEst 
         Caption         =   "Impr. Estoque"
         Height          =   735
         Left            =   13920
         Picture         =   "frmConsign.frx":325C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   8295
      Left            =   5520
      TabIndex        =   21
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "F2 - Posição Atual de Estoque Consignado"
      TabPicture(0)   =   "frmConsign.frx":3358
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgProd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pcOpt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "F3 - Posição das Transferências em Aberto"
      TabPicture(1)   =   "frmConsign.frx":3374
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtgTrans"
      Tab(1).Control(1)=   "adcTrans"
      Tab(1).Control(2)=   "Picture1"
      Tab(1).Control(3)=   "cmndbtnAcertos"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmndbtnAcertos 
         Caption         =   "Acertos"
         Height          =   255
         Left            =   -66480
         TabIndex        =   76
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox pcOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   5385
         TabIndex        =   71
         Top             =   480
         Width           =   5415
         Begin VB.OptionButton optnPorSub 
            Caption         =   "Por Subseção"
            Height          =   195
            Left            =   3720
            TabIndex        =   75
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optnPorSec 
            Caption         =   "Por Seção"
            Height          =   195
            Left            =   2520
            TabIndex        =   74
            Top             =   0
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optnPorCod 
            Caption         =   "Por Código"
            Height          =   195
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optnPorDesc 
            Caption         =   "Por Descrição"
            Height          =   195
            Left            =   1200
            TabIndex        =   72
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         ScaleHeight     =   225
         ScaleWidth      =   8265
         TabIndex        =   65
         Top             =   480
         Width           =   8295
         Begin VB.CheckBox chkPcDif 
            Caption         =   "Pç Diferente"
            Height          =   195
            Left            =   6360
            TabIndex        =   79
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optnPorSubTra 
            Caption         =   "Por Subseção"
            Height          =   195
            Left            =   3240
            TabIndex        =   70
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optnPorSecTra 
            Caption         =   "Por Seção"
            Height          =   195
            Left            =   2160
            TabIndex        =   69
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optnPorData 
            Caption         =   "Por Data"
            Height          =   195
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optnPorProduto 
            Caption         =   "Por Produto"
            Height          =   195
            Left            =   960
            TabIndex        =   67
            Top             =   0
            Width           =   1215
         End
         Begin VB.CheckBox chkAberta 
            Caption         =   "Com remanescente"
            Height          =   195
            Left            =   4680
            TabIndex        =   66
            Top             =   0
            Value           =   1  'Checked
            Width           =   1695
         End
      End
      Begin MSAdodcLib.Adodc adcTrans 
         Height          =   330
         Left            =   -74640
         Top             =   15360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dtgTrans 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   77
         Top             =   840
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11668
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "item"
            Caption         =   "N."
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
            DataField       =   "ref_pro"
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
         BeginProperty Column03 
            DataField       =   "data_con"
            Caption         =   "Data"
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
            DataField       =   "qtd_con"
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
         BeginProperty Column06 
            DataField       =   "reman_con"
            Caption         =   "Restam"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "unit_con"
            Caption         =   "Unitário"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "pv_pro"
            Caption         =   "Unit. atual"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$ ""#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "total"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "atual"
            Caption         =   "Preço Atual"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "nome_sec"
            Caption         =   "Seção"
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
            DataField       =   "nome_sub"
            Caption         =   "Subseção"
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
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200,189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2445,166
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3555,213
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1709,858
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1860,095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgProd 
         Height          =   6975
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   12303
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "item"
            Caption         =   "N"
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
            DataField       =   "reman"
            Caption         =   "estoque"
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
            DataField       =   "pv_pro"
            Caption         =   "Preço Atual"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "total"
            Caption         =   "Total Transferido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """R$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   464,882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   5309,858
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1305,071
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8370
      Left            =   0
      ScaleHeight     =   8340
      ScaleWidth      =   5385
      TabIndex        =   13
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   0
         ScaleHeight     =   1305
         ScaleWidth      =   5385
         TabIndex        =   14
         Top             =   6120
         Width           =   5415
         Begin VB.Label lblTotalPrecoAtual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2925
            TabIndex        =   52
            Tag             =   "!"
            Top             =   360
            Width           =   2505
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Estoque X Preço Atual:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   0
            TabIndex        =   51
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Transferências:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lblProd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1800
            TabIndex        =   49
            Tag             =   "!"
            Top             =   0
            Width           =   945
         End
         Begin VB.Label lblItens 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4080
            TabIndex        =   48
            Tag             =   "!"
            Top             =   0
            Width           =   1305
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Itens:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2640
            TabIndex        =   47
            Top             =   0
            Width           =   2070
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Total X Preço Tansferido:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   0
            TabIndex        =   46
            Tag             =   "!"
            Top             =   720
            Width           =   3030
         End
         Begin VB.Label lblValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   2925
            TabIndex        =   15
            Tag             =   "!"
            Top             =   720
            Width           =   2505
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   5175
         TabIndex        =   32
         Top             =   3960
         Width           =   5175
         Begin VB.TextBox txtbxFones 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            MaxLength       =   20
            TabIndex        =   44
            Top             =   1440
            Width           =   4335
         End
         Begin VB.TextBox txtbxUf 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   4680
            MaxLength       =   20
            TabIndex        =   43
            Text            =   "GO"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtbxCNPJ 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   840
            MaxLength       =   18
            TabIndex        =   37
            Top             =   0
            Width           =   2295
         End
         Begin VB.TextBox txtbxEndereco 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            MaxLength       =   40
            TabIndex        =   36
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox txtbxBairro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            MaxLength       =   20
            TabIndex        =   35
            Top             =   720
            Width           =   4335
         End
         Begin VB.TextBox txtbxCidade 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   840
            MaxLength       =   20
            TabIndex        =   34
            Text            =   "Anápolis"
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox txtbxRg 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   3720
            MaxLength       =   18
            TabIndex        =   33
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Fones:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   45
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblCpf 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&CPF:"
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
            Height          =   285
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&Endereço:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   41
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&Bairro:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   40
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Ci&dade:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   39
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "&RG:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3240
            TabIndex        =   38
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.PictureBox pcLimite 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   5175
         TabIndex        =   23
         Top             =   960
         Width           =   5175
         Begin VB.Label lblLimiteRest 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Width           =   390
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   690
         End
         Begin VB.Label lblEmUso 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000342C5&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   390
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label lblLimite 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   405
         End
         Begin VB.Label lblStatus_CLi 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Liberado"
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
            TabIndex        =   25
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox pcCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00EF9C7C&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   825
         ScaleWidth      =   9105
         TabIndex        =   22
         Top             =   0
         Width           =   9135
         Begin VB.CommandButton cmdLembrete 
            Caption         =   "Lembrete"
            Height          =   255
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtbxNomeCli 
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
            Top             =   240
            Width           =   4575
         End
         Begin VB.CommandButton cmndbtnCliente 
            Height          =   330
            Left            =   4680
            Picture         =   "frmConsign.frx":3390
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Representante:"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label lblPExtenso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Left            =   2640
         TabIndex        =   16
         Top             =   6120
         Width           =   150
      End
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   16725
      TabIndex        =   3
      Top             =   9270
      Width           =   16785
      Begin VB.CommandButton cmndbtnBuscaFun 
         Height          =   315
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtbxVendedor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   17
         Text            =   "Nome do Funcionario"
         Top             =   15
         Width           =   3255
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4680
         TabIndex        =   20
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   60
         Width           =   420
      End
      Begin VB.Label lblTur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9000
         TabIndex        =   12
         Top             =   0
         Width           =   270
      End
      Begin VB.Label lblBloqueado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9960
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8400
         TabIndex        =   10
         Top             =   60
         Width           =   465
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5880
         TabIndex        =   9
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7560
         TabIndex        =   8
         Top             =   60
         Width           =   435
      End
      Begin VB.Label lblNomeEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11520
         TabIndex        =   7
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   60
         Width           =   630
      End
      Begin VB.Label lblnum_ven 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6480
         TabIndex        =   5
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atendente:"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   15
         Width           =   840
      End
   End
   Begin VB.CommandButton cmndbtnLocaliza 
      Caption         =   "..."
      Height          =   255
      Left            =   14760
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmConsign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim prod As ProdutoSaida
Option Explicit
Dim bDirecao As Integer
Dim bActive As Boolean
Dim preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer

Public Atacado As Boolean

Private Type prodPromo
    nome As String
    preco As String
End Type
Dim pPromo() As prodPromo
Dim QtdPromo As Integer
Dim iIndPromo As Integer
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim cFile As New Collection
Dim ai As New ADODB.Stream
Dim Indice As Integer

Public WithEvents adoprod As ADODB.Recordset
Attribute adoprod.VB_VarHelpID = -1
Public WithEvents AdoTrans As ADODB.Recordset
Attribute AdoTrans.VB_VarHelpID = -1

Sub buscaFoto()
    Dim sFile As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    Set cFile = Nothing
    ai.type = adTypeBinary
    strSQL = "select seq_fot,foto_fot from tab_fotorep where cod_rep='" & txtbxNomeCli.tag & "' order by seq_fot desc"
    
    'Set r = c.Execute(strSQL)
    r.Open strSQL, Conn_Client
    If Dir(sFoto_Temp, vbDirectory) = "" Then
        fileSystem.MkDir sFoto_Temp
    End If
    
    With r
        Do While Not .EOF
            sFile = sFoto_Temp & "\F" & !seq_fot & ".tmp"
            ai.Open
            ai.Write r!foto_fot
            ai.SaveToFile (sFile)
            ai.Close
            cFile.Add sFile
            .MoveNext
        Loop
        
        .Close
    
    End With
trata_Erro:
        
    pcDataFoto_click
End Sub

Private Sub pcDataFoto_click()
    Dim sFile As String
    Dim wid As Long
    Dim hei As Long
    Dim w As Long
    Dim h As Long
    Dim iCount As Integer
    
    For iCount = 0 To cFile.Count - 1
    
        'If pc3X4(iCount).Tag = "" Then
        '    If cFile.count > 0 Then
        '        pc3X4(iCount).Tag = 1
        '    Else
        '        Exit Sub
        '    End If
        'End If
            
        If iCount >= pc3X4.Count Then
            Load pc3X4(iCount)
            Load p(iCount)
            pc3X4(iCount).Visible = True
            pc3X4(iCount).top = (pc3X4(iCount - 1).height) + pc3X4(iCount - 1).top + 100
            pc3X4(iCount).Left = (pc3X4(iCount - 1).Left)
        End If
            
        'sFile = cFile(Int(pc3X4(iCount).Tag))
        sFile = cFile(iCount + 1)
        
        'lblFoto.caption = pc3X4(iCount).Tag & " de " & cFile.count
        'lblFoto.Tag = sFile
        On Error GoTo trata_Erro
                pc3X4(iCount).ScaleMode = 1
                p(iCount).ScaleMode = 1
                p(iCount).AutoRedraw = True
                p(iCount).AutoSize = True
                p(iCount).Picture = LoadPicture(sFile)
                wid = p(iCount).image.width
                hei = p(iCount).image.height
                w = pc3X4(iCount).width
                h = pc3X4(iCount).height
                Do Until hei < h And wid < w
                    hei = hei * 0.995
                    wid = wid * 0.995
                Loop
                pc3X4(iCount).Picture = Nothing
                pc3X4(iCount).PaintPicture LoadPicture(sFile), 0, 0, ScaleX(wid, pc3X4(iCount).ScaleMode, vbTwips), ScaleY(hei, pc3X4(iCount).ScaleMode, vbTwips)
                pc3X4(iCount).AutoSize = True
                pc3X4(iCount).Refresh
                'lblFoto.caption = .AbsolutePosition & " de " & .RecordCount
               ' Kill sFile
        Next
        
        
        
        
        
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub


Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Currency
    Dim cDesconto As Currency
    Dim cFinal As Currency
    Dim bDif As Boolean
    On Error GoTo trata_Erro:
    Dim cTotPrecoAtual As Currency
    P_Status "Atualizando Valores..."
    
    iPos = AdoTrans.AbsolutePosition
    AdoTrans.MoveFirst
    
    dtgProd.Visible = False
    dtgTrans.Visible = False
    With AdoTrans
        
        
        Do While Not .EOF
            
            If !pv_pro <> !unit_con Then
                bDif = True
            End If
            cTotal = cTotal + !total
            cFinal = cFinal + !total
            'cDesconto = cDesconto + !Desconto
            cItens = cItens + CCur(!reman_con)
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cFinal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = AdoTrans.RecordCount
        'lblTotal.caption = FormatVal(cTotal)
        'lblDesconto.caption = FormatVal(cDesconto)
        lblPExtenso.caption = Extenso(CDbl(lblValor.caption))
    End With
        
    
    
    If AdoTrans.RecordCount > 0 Then
        AdoTrans.MoveFirst
    End If
    If iPos > 0 Then
        AdoTrans.Move iPos
    End If
    

    cTotal = 0
    iPos = adoprod.AbsolutePosition
    If adoprod.RecordCount > 0 Then
        adoprod.MoveFirst
    End If
    With adoprod
        
        
        Do While Not .EOF
            
            cTotal = cTotal + !totalatual
            
            
            
            .MoveNext
        Loop
        lblTotalPrecoAtual.caption = FormatVal(cTotal, "###,##0.000")
        
    End With
        
    If bDif = True And bActive = False Then
'        MsgBox "Há diferenças de preços atuais e preços de mercadorias consignadas para este representante!", vbExclamation
        
    End If
    dtgProd.Visible = True
    dtgTrans.Visible = True
    
    
    P_Status ""
    Exit Sub
trata_Erro:
dtgProd.Visible = False
    dtgTrans.Visible = False
    lblTotalPrecoAtual.caption = "0"
    lblValor.caption = "0"
End Sub

Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    Dim sOrderBy As String
    
    If optnPorCod.Value Then
        sOrderBy = "tp.cod_pro"
    ElseIf optnPorDesc.Value Then
       sOrderBy = "tp.nome_pro,tp.cod_pro"
    ElseIf optnPorSec.Value Then
        sOrderBy = "nome_sec,nome_sub,tp.cod_pro"
    ElseIf optnPorSub.Value Then
        sOrderBy = "nome_sub,nome_sec,tp.cod_pro"
    End If

    strSQL = "select row_number() over(order by " & sOrderBy & ") as item, sum(tc.reman_con) as reman,(unit_con) as ultimo, estoque_pro*pv_pro as totalatual,tp.pa_pro,"
    strSQL = strSQL & " tp.pv_pro,tp.ref_pro,tp.pv_pro*(sum(reman_con)) as total, "
    strSQL = strSQL & "tp.nome_pro,te.* "
    strSQL = strSQL & " from  tab_produtos as tp, tab_secoes as ts, tab_subsec as tu,"
    strSQL = strSQL & " tab_estcons as te,tab_consignacoes as tc  "
    strSQL = strSQL & " where te.cod_pro=tp.cod_pro and te.cod_rep='" & txtbxNomeCli.tag & "' "
    
    strSQL = strSQL & " and tu.cod_sub=tp.cod_sub "
    strSQL = strSQL & " and tu.cod_sec=ts.cod_sec"
    strSQL = strSQL & " and tc.cod_pro=te.cod_pro "
    strSQL = strSQL & " and tc.cod_rep=te.cod_rep "
    
    
    strSQL = strSQL & " group by te.cod_pro "
    
    strSQL = strSQL & " having sum(reman_con) >0 "
    strSQL = strSQL & " order by " & sOrderBy
    
    'strSQL = "select * from tab_consignacoes"
    'strSQL = "select * from tab_produtos"
       
    'adcProd.RecordSource = strSQL
    'adcProd.Refresh
    P_Status "aguarde..."
    
    abreObjetoAdo adoprod, strSQL, dtgProd
    
    
    If optnPorData.Value Then
        sOrderBy = " data_con desc"
    ElseIf optnPorSecTra.Value Then
        sOrderBy = "nome_sec,nome_sub,data_con desc "
    ElseIf optnPorSubTra.Value Then
        sOrderBy = "nome_sub,nome_sec,data_con desc "
    Else
        sOrderBy = "tp.cod_pro,data_con desc "
    
    End If
    
    strSQL = "select row_number() over( order by " & sOrderBy & ") as item,nome_sub,nome_sec,tp.pv_pro,tp.pa_pro,tp.nome_pro,tp.ref_pro,tc.*, unit_con*reman_con as total,pv_pro*reman_con as atual"
    strSQL = strSQL & " from tab_produtos as tp, tab_consignacoes as tc,tab_subsec as tu,tab_secoes as ts "
    strSQL = strSQL & " where tc.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and tu.cod_sub=tp.cod_sub "
    strSQL = strSQL & " and tu.cod_sec=ts.cod_sec"
    If chkAberta.Value = 1 Then
        strSQL = strSQL & " and tc.reman_con>0"
    End If
    strSQL = strSQL & " and tc.cod_rep='" & txtbxNomeCli.tag & "' "
     
     
    If chkPcDif.Value = 1 Then
        strSQL = strSQL & " and tp.pv_pro<>unit_con "
    End If
     
    
    strSQL = strSQL & " order by " & sOrderBy
    
'    strSQL = "select * from tab_estcons"
    
    'adcTrans.RecordSource = strSQL
    'adcTrans.Refresh
    
    
    abreObjetoAdo AdoTrans, strSQL, dtgTrans
    
    
    
    soma
    
    
    P_Status
    Exit Sub
    P_Status
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub






Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub








Private Sub chkAberta_Click()
    P_buscaitens
End Sub

Private Sub chkPcDif_Click()
    P_buscaitens
End Sub

Private Sub cmdLembrete_Click()
    buscaLembrete txtbxNomeCli.tag, True
End Sub

Private Sub cmdNFedev_Click()
    If MsgBox("Confirma geração do manifesto de retorno de remessa?", vbYesNo + vbQuestion) = vbYes Then
        gerarManifesto tipo_operacao_venda.retorno_remessa
    End If
End Sub

Private Sub cmndbtnAcerto_Click()
    If txtbxNomeCli.tag = "" Then
        MsgBox "Escolha primeiro o Representante!", vbInformation
        cmndbtnCliente_Click
        Exit Sub
    End If
    frmAcerto.txtbxRepresentante.tag = txtbxNomeCli.tag
    frmAcerto.caption = "Acerto de: " & txtbxNomeCli.Text
    frmAcerto.txtbxRepresentante.Text = txtbxNomeCli.Text
    frmAcerto.lblCpf.caption = lblCpf.caption
    frmAcerto.Show 0
    
    
End Sub

Private Sub cmndbtnAcertos_Click()
    On Error Resume Next
    P_HistProdCons txtbxNomeCli.tag, AdoTrans!cod_pro
End Sub

Private Sub cmndbtnAtualizarPrecos_Click()
    Dim strSQL As String
    If strAlterarComRepr_Usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso! (alterarComRepr)", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja Realmente mudar os preços dos produtos do Representante para os preços atuais?", vbYesNo + vbCritical) = vbYes Then
        
            
        With AdoTrans
            .MoveFirst
            Do While Not .EOF
                strSQL = "update tab_consignacoes set unit_con=" & F_GravaDec(!pv_pro) & " where seq_con='" & !seq_con & "' and cod_pro='" & !cod_pro & "'"
                Conn.Execute strSQL
                
                .MoveNext
            Loop
        End With
        P_buscaitens
            
        'strSQL = "update tab_consignacoes set unit_con=" & F_GravaDec(
    
    
    End If
End Sub

Private Sub cmndbtnBuscaFun_Click()
    Dim f As funcionario.funcionario
    On Error GoTo trata_Erro
    
    
    f = funcionario.F_SelectVendedor("")
    
    
    
    If f.cod_fun <> "" Then
        If F_senhaReturnUsu(True, False, f.guerra_fun, False) Then
        
        
            txtbxVendedor.Text = f.nome_fun
            txtbxVendedor.tag = f.cod_fun
            lblDesc.caption = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & f.cod_fun & "'")
            
        End If
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Private Sub cmndbtnCliente_Click()
    Dim c As representante.representante
    
    c = representante.F_BuscaRep(txtbxNomeCli.Text)
    If c.cod_Rep <> "" Then
        txtbxNomeCli.Text = c.nome_rep
        txtbxNomeCli.tag = c.cod_Rep
        lblLimite.caption = FormatVal(c.limite_rep)
        lblEmUso.caption = FormatVal(F_BuscaLimiteEmUsoRep(c.cod_Rep))
        lblLimiteRest.caption = FormatVal(c.limite_rep - CCur(lblEmUso.caption))
        lblStatus_CLi.tag = c.status_rep
        txtbxCNPJ.Text = c.cpf_rep
        txtbxEndereco.Text = c.endereco_rep
        txtbxBairro.Text = c.bairro_rep
        txtbxCidade.Text = c.cidade_rep
        txtbxUf.Text = c.uf_rep
        txtbxFones.Text = c.telefone_rep
        txtbxRg.Text = c.rg_rep
        If c.status_rep = "1" Then
            lblStatus_CLi.caption = "Liberado"
            lblStatus_CLi.ForeColor = vbGreen
        ElseIf c.status_rep = "2" Then
            lblStatus_CLi.caption = "Cheques"
            lblStatus_CLi.ForeColor = vbYellow
        Else
            lblStatus_CLi.caption = "Bloqueado"
            lblStatus_CLi.ForeColor = vbRed
        End If
        If c.status_rep = "2" Then
            MsgBox "O Representante encontra-se negativado, portanto, só é possível venda a vista!", vbExclamation
            
            
        End If
        
        
        
        
        
        buscaFoto
        bActive = False
        
        
        
                    
        buscaLembrete txtbxNomeCli.tag
            
            
        
        
        pcLimite.Visible = True
        
        If RepresentanteBloqueado(c.cod_Rep) Then
            MsgBox "Este representante encontra-se bloqueado, não é possível carregar seu estoque!", vbExclamation
            pcOpt.Visible = False
            Exit Sub
        End If
        
        
        P_buscaitens

        'carrega
        
    Else
        
    End If
     
            
     
    
End Sub

Private Sub cmbxFormaPag_Change()

End Sub

Private Sub cmbxFormaPag_DropDown()

End Sub










Private Sub cmndbtnConsign_Click()
    On Error GoTo trata_Erro
    imprimeConsign AdoTrans!seq_con
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnDevolver_Click()
    abreDevolucao
End Sub

Private Sub cmndbtnEstoque_Click()
    If MsgBox("Confirma geração do manifesto?", vbYesNo + vbQuestion) = vbYes Then
        gerarManifesto
    End If
End Sub
Sub gerarManifesto(Optional tipo As tipo_operacao_venda = tipo_operacao_venda.Remessa)
    Dim NumVen As String
    Dim item As String
    Dim p As Produto
    With adoprod
        If .RecordCount <= 0 Then
            Exit Sub
        End If
        NumVen = F_GeraNovaVenda("00000", tipo)
        dtgProd.Visible = False
        dtgTrans.Visible = False
        
        .MoveFirst
        Do While Not .EOF
            p.cod_pro = !cod_pro
            p.pv_pro = !pv_pro
            p.nome_pro = !nome_pro
            item = Vendas.F_AdicionaItem(NumVen, p, !reman, "0", !total)
            
            
            
            .MoveNext
        Loop
        .MoveFirst
    End With
    dtgProd.Visible = True
    dtgTrans.Visible = True
    Dim fr As New frmVenda
    fr.tag = NumVen
    fr.Show 1
    
    
End Sub
Private Sub cmndbtnImprEst_Click()
    Dim strSQL As String
    Dim sEnd As String
    If F_BuscaUmCampo("bloqimpr", "tab_representantes", "cod_rep='" & txtbxNomeCli.tag & "'") = "S" Then
        MsgBox "O representante está com a impressão de estoque bloqueada!", vbExclamation
        Exit Sub
    End If
    
    
    
    sEnd = txtbxEndereco.Text & " " & txtbxBairro.Text & " " & txtbxCidade.Text & "/ " & txtbxUf.Text & " Fone: " & txtbxFones.Text
    strSQL = "Data: " & Date & " -  Valor Total: " & FormatVal(lblTotalPrecoAtual.caption)
    P_DbgrdToHtml txtbxNomeCli.tag & "-" & txtbxNomeCli.Text, dtgProd, adoprod, strSQL, "5", sEnd
    
End Sub

Private Sub cmndbtnListaDev_Click()
    P_ListaDevConsign txtbxNomeCli.tag
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    If adoprod.RecordCount > 0 Then
        
            If MsgBox("Confirma Saída da Consignação?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
    End If
    
    
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro "cmndbtnSair"
    Unload Me
End Sub






Private Sub cmndbtnTransf_Click()
    frmTransfConsign.txtbxNomeCli.tag = txtbxNomeCli.tag
    frmTransfConsign.txtbxNomeCli.Text = txtbxNomeCli.Text
    frmTransfConsign.Show 0
End Sub


Private Sub dbgrdProd_DblClick()
    'cmndbtnEditProd_Click
End Sub

Private Sub dbgrdProd_GotFocus()
    
    
    
End Sub






Private Sub Form_Activate()
    On Error GoTo trata_Erro
    
    Form_Resize
    lblnum_ven.caption = Me.tag
    
    lblNomeEmp.caption = strFantasia_Emp
    lblCaixa = strCod_Cai
    lblTur.caption = strCod_Tur
    If bActive = True Then Exit Sub
    
    
    
    
    P_Status "Buscando Itens..."
    
    P_buscaitens
    
    P_Status ""
    
    bActive = True
    If strCod_Cai = "" Then
        MsgBox "Não existe caixa para este computador, o caixa ficará para o '00' ", vbExclamation
        strCod_Cai = "00"
    End If
    
    If txtbxNomeCli.tag = "" Then
        cmndbtnCliente_Click
    End If
        
    Exit Sub
trata_Erro:
'    P_MostraErro "Form_Activate"
    
End Sub
Sub playGif()
    'WebBrowser1.Navigate "c:\market\gifs\maebela2.gif"
    'WebBrowser1.Navigate "about:<html><body scroll='no'><img src='c:\maebela.gif'></img></body></html>"
    
End Sub

    


Sub abreDevolucao()
    
    frmDevolucaoConsign.txtbxRepresentante.Text = txtbxNomeCli.Text
    frmDevolucaoConsign.txtbxRepresentante.tag = txtbxNomeCli.tag
    frmDevolucaoConsign.Show 0
    
    
    
    
    
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    bDescancoTela = False
    iDescancoTela = 0
    
    
    
    If KeyCode = vbKeyF3 Then
        
    ElseIf KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    
    ElseIf KeyCode = vbKeyF4 Or (KeyCode = vbKeyT And Shift = vbCtrlMask) Then
    
         cmndbtnTransf_Click
    ElseIf KeyCode = vbKeyF5 Or (KeyCode = vbKeyA And Shift = vbCtrlMask) Then
        cmndbtnAcerto_Click
    ElseIf KeyCode = vbKeyF6 Then
        abreDevolucao
    ElseIf KeyCode = vbKeyF7 Then
        
        'cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF2 Then
        
    ElseIf KeyCode = vbKeyF8 Then
        
        
        operacao = "C"
        mudamodo
    ElseIf KeyCode = vbKeyF9 Then
        'If operacao = "C" And adc.Recordset.RecordCount > 0 Then
        '    MsgBox "Operacação Impossível!", vbCritical
        '    Exit Sub
        'End If
        If lblBloqueado.caption = "S" Then
            Exit Sub
        End If
        
        operacao = "B"
        mudamodo
    ElseIf KeyCode = vbKeyF10 Then
        If lblBloqueado.caption = "S" Then
            Exit Sub
        End If
        
        
        
    
        mudamodo
    ElseIf KeyCode = vbKeyF11 Then
        
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        frmECF.Show 1
    End If
    
    If KeyCode = vbKeyC And (Shift = vbAltMask Or Shift = vbCtrlMask) Then
        txtbxNomeCli.SetFocus
    ElseIf KeyCode = vbKeyP And Shift = vbCtrlMask Then
        
    ElseIf KeyCode = vbKeyE And Shift = vbAltMask Then
        
    End If
    
End Sub
Sub mudamodo()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set modo_cai='" & operacao & "' "
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
    Conn.Execute strSQL
End Sub
Private Sub Form_Load()
    resizeForm Me
    bActive = False
    
    'adcProd.ConnectionString = Conn.ConnectionString
    'adcTrans.ConnectionString = Conn.ConnectionString
    
    Set AdoTrans = New ADODB.Recordset
    Set adoprod = New ADODB.Recordset

    
    
    lblNomeEmp.caption = strNome_Emp
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", "tab_parametros", "")
    
    
    'P_PreencheFormasPag cmbxFormaspag
    
    
'    cmbxBuscar.ListIndex = 0
    
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
    txtbxVendedor.tag = strCod_Usu
    txtbxVendedor.Text = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & strCod_Usu & "'")
    lblDesc.caption = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & strCod_Usu & "'")
    'ajustaGridAcerto
    If Me.tag = "t" Then
    End If
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDescancoTela = False
    iDescancoTela = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Resize()
    ajustaForm
    
End Sub
Sub ajustaForm()
    Dim iCount As Integer
    On Error Resume Next
    
    
    
    
    
    
    sstb.width = Me.width - sstb.Left - 300
    sstb.height = pcAtalhos.top - sstb.top - 100
    
    lblNomeEmp.Left = Me.width - lblNomeEmp.width - 200
    
    'pcProd.Width = sstb.Width - 200
    'msflxgrdSaidas.Width = pcProd.Width
    dtgProd.height = sstb.height - dtgProd.top - 100
    dtgProd.width = sstb.width - 400
    dtgTrans.width = dtgProd.width
    dtgTrans.height = dtgProd.height
    'pcAtalhos.Width = sstb.Width
    
    
    
    'For iCount = 0 To pcBotao.count - 1
    '    pcBotao(iCount).Width = pcAtalhos.Width / (pcBotao.count)
    '    If iCount > 0 Then
    '        pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).Width + 50
    '    End If
        
    'Next
 
    
    
    
    
End Sub


Private Sub optnCod_Click()
    
End Sub


Private Sub optnSeq_Click()
    P_buscaitens
End Sub



Private Sub Form_Terminate()
    On Error Resume Next
    
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFrm.Enabled = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub frmCli_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub lblMsg2_Click()

End Sub



Private Sub optnCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtbxNomeCli.SetFocus
    End If

End Sub


Private Sub optnPorCod_Click()
    P_buscaitens
End Sub

Private Sub optnPorData_Click()
    P_buscaitens
End Sub

Private Sub optnPorDesc_Click()
    P_buscaitens
End Sub

Private Sub optnPorProduto_Click()
    P_buscaitens
End Sub

Private Sub optnPorSecTra_Click()
    P_buscaitens
End Sub

Private Sub optnPorSub_Click()
    P_buscaitens
End Sub

Private Sub optnPorSubTra_Click()
    P_buscaitens
End Sub

Private Sub pcMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub pcPromo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub




Private Sub pcProd_GotFocus()
    Sendkeys Chr(9)
End Sub



Private Sub pcProdAcerto_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub sstb_Click(PreviousTab As Integer)
    If sstb.Tab = 0 Then
        
    ElseIf sstb.Tab = 1 Then
    ElseIf sstb.Tab = 2 Then
    ElseIf sstb.Tab = 3 Then
        'txtbxClienteAcerto.SetFocus
    ElseIf sstb.Tab = 4 Then
    End If
End Sub

Private Sub txtbxBairro_GotFocus()
    'P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub




Private Sub txtbxEndereco_GotFocus()
 '   P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub





Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
    If KeyCode = vbKeyDown Then
        
    End If
End Sub


Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub

Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli, 1
    
End Sub




















Private Sub txtbxVendedor_GotFocus()
    P_SelecTexto txtbxVendedor
End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaFun_Click
    End If
End Sub



