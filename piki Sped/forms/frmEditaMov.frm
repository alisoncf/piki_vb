VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEditaMov 
   Caption         =   "Documentos a Receber"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   17205
   Icon            =   "frmEditaMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   17205
   Begin VB.CommandButton cmndbtnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "Pesquisar (F5)"
      Height          =   600
      Left            =   14160
      Picture         =   "frmEditaMov.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   480
      Width           =   1575
   End
   Begin TabDlg.SSTab sstb 
      Height          =   1260
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   2223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Informações Gerais"
      TabPicture(0)   =   "frmEditaMov.frx":0D0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Busca Avançada"
      TabPicture(1)   =   "frmEditaMov.frx":0D28
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture4"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -74880
         ScaleHeight     =   735
         ScaleWidth      =   12255
         TabIndex        =   26
         Top             =   420
         Width           =   12255
         Begin VB.TextBox txtNFe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11040
            TabIndex        =   31
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtbxHist 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   60
            TabIndex        =   49
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtbxSeqDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            MaxLength       =   60
            TabIndex        =   44
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxSeqItd 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5400
            MaxLength       =   60
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxValor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtBoleto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtbxContaCob 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8760
            TabIndex        =   28
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtbxTipoCob 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6480
            TabIndex        =   27
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnTipoCob 
            Height          =   285
            Left            =   8400
            Picture         =   "frmEditaMov.frx":0D44
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnCob 
            Height          =   285
            Left            =   10680
            Picture         =   "frmEditaMov.frx":10CE
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NFe"
            Height          =   195
            Left            =   11040
            TabIndex        =   30
            Top             =   0
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Histórico:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   0
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id. Doc:"
            Height          =   195
            Left            =   4440
            TabIndex        =   48
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Id. Parc:"
            Height          =   195
            Left            =   5400
            TabIndex        =   47
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   3480
            TabIndex        =   46
            Tag             =   "!"
            Top             =   0
            Width           =   405
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Boleto:"
            Height          =   195
            Left            =   2280
            TabIndex        =   45
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cobrança:"
            Height          =   195
            Left            =   6480
            TabIndex        =   34
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label lblTipoCob 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta Cobrança:"
            Height          =   195
            Left            =   8760
            TabIndex        =   33
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   15615
         TabIndex        =   18
         Top             =   420
         Width           =   15615
         Begin VB.CheckBox chkVinc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Documento Vinculador"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   8880
            TabIndex        =   54
            Top             =   120
            Width           =   1935
         End
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmEditaMov.frx":1458
            Left            =   7320
            List            =   "frmEditaMov.frx":1468
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtbxTerceiro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3720
            TabIndex        =   35
            Top             =   240
            Width           =   3255
         End
         Begin VB.PictureBox pcLista 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            Picture         =   "frmEditaMov.frx":1494
            ScaleHeight     =   195
            ScaleWidth      =   180
            TabIndex        =   22
            Top             =   0
            Width           =   180
         End
         Begin VB.TextBox txtbxNumDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            MaxLength       =   60
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnTerc 
            Height          =   285
            Left            =   6960
            Picture         =   "frmEditaMov.frx":158D
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenar Por:"
            Height          =   195
            Left            =   7320
            TabIndex        =   40
            Tag             =   "!"
            Top             =   0
            Width           =   900
         End
         Begin VB.Label lblTerCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Participante:"
            Height          =   195
            Left            =   3720
            TabIndex        =   38
            Top             =   0
            Width           =   885
         End
         Begin VB.Label lblCPF 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ"
            Height          =   195
            Left            =   4815
            TabIndex        =   37
            Top             =   0
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Doc."
            Height          =   195
            Left            =   2520
            TabIndex        =   25
            Tag             =   "!"
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1320
            TabIndex        =   24
            Tag             =   "!"
            Top             =   0
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Tag             =   "!"
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Height          =   5055
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   8916
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
      Caption         =   "Lista de Documentos Encontrados"
      ColumnCount     =   25
      BeginProperty Column00 
         DataField       =   "num_doc"
         Caption         =   "Número"
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
         DataField       =   "boleto_itd"
         Caption         =   "Boleto"
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
         DataField       =   "nnf"
         Caption         =   "NFe"
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
         DataField       =   "serie"
         Caption         =   "Serie"
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
         DataField       =   "cod_ter"
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
      BeginProperty Column05 
         DataField       =   "nome"
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
      BeginProperty Column06 
         DataField       =   "data_doc"
         Caption         =   "Data Doc."
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
         DataField       =   "reman_itd"
         Caption         =   "Restante"
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
         DataField       =   "venc_itd"
         Caption         =   "Vencimento"
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
         DataField       =   "vencini_itd"
         Caption         =   "Inicial"
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
         DataField       =   "valor_itd"
         Caption         =   "Valor"
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
         DataField       =   "historico_doc"
         Caption         =   "Histórico"
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
         DataField       =   "parcela_itd"
         Caption         =   "Parc."
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
         DataField       =   "parcelas_doc"
         Caption         =   "De"
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
         DataField       =   "status"
         Caption         =   "Status Cliente"
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
         DataField       =   "nome_tip"
         Caption         =   "Tipo de Cobrança"
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
      BeginProperty Column16 
         DataField       =   "nome_cob"
         Caption         =   "Cobrança"
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
      BeginProperty Column17 
         DataField       =   "esp_doc"
         Caption         =   "E"
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
      BeginProperty Column18 
         DataField       =   "seq_doc"
         Caption         =   "Id.Doc"
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
      BeginProperty Column19 
         DataField       =   "seq_itd"
         Caption         =   "Id.Parc."
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
      BeginProperty Column20 
         DataField       =   "dias"
         Caption         =   "Dias"
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
      BeginProperty Column21 
         DataField       =   "nome_cid"
         Caption         =   "Cidade"
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
      BeginProperty Column22 
         DataField       =   "boleto_itd"
         Caption         =   "Boleto"
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
      BeginProperty Column23 
         DataField       =   "nome_set"
         Caption         =   "setor"
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
      BeginProperty Column24 
         DataField       =   "cod_usu"
         Caption         =   "Usuário"
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
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   510,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3614,74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   3030,236
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   929,764
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   17145
      TabIndex        =   3
      Top             =   8490
      Width           =   17205
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   595
         Left            =   8760
         Picture         =   "frmEditaMov.frx":1917
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdVinc 
         Caption         =   "&Vincular"
         Height          =   595
         Left            =   4440
         Picture         =   "frmEditaMov.frx":1E08
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbrnPrejuizo 
         Caption         =   "Prejuízo"
         Height          =   595
         Left            =   11400
         Picture         =   "frmEditaMov.frx":2287
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEditer 
         Caption         =   "&Editar Terceiro"
         Height          =   595
         Left            =   14280
         Picture         =   "frmEditaMov.frx":2811
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnContrato 
         Caption         =   "Gerar PDF"
         Height          =   595
         Left            =   10080
         Picture         =   "frmEditaMov.frx":296D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnCarne 
         Caption         =   "Emitir &Doc"
         Height          =   595
         Left            =   7320
         Picture         =   "frmEditaMov.frx":2A69
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   595
         Left            =   2880
         Picture         =   "frmEditaMov.frx":2B65
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   595
         Left            =   12840
         Picture         =   "frmEditaMov.frx":2CAF
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   595
         Left            =   1440
         Picture         =   "frmEditaMov.frx":2DF9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   595
         Left            =   0
         Picture         =   "frmEditaMov.frx":3183
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   595
         Left            =   15600
         Picture         =   "frmEditaMov.frx":32CD
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   595
         Left            =   5880
         Picture         =   "frmEditaMov.frx":3517
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EF9C7C&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   17175
      TabIndex        =   0
      Top             =   8115
      Width           =   17205
      Begin VB.CommandButton cmdSoma 
         Caption         =   "Soma:"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   0
         Width           =   615
      End
      Begin VB.CheckBox chkPrej 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Prejuízo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkAberto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Parcelas em Aberto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.Menu mnAcoes 
      Caption         =   "&Ações"
      Begin VB.Menu mnSelec 
         Caption         =   "&Selecionar"
      End
      Begin VB.Menu Div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSPC 
         Caption         =   "&SPC"
      End
      Begin VB.Menu dov4 
         Caption         =   "-"
      End
      Begin VB.Menu mnEdiPar 
         Caption         =   "&Editar Parcela"
      End
      Begin VB.Menu mnNova 
         Caption         =   "&Nova Parcela"
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnBaixar 
         Caption         =   "&Baixar Selecionadas"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnParcela 
      Caption         =   "&Parcela"
      Begin VB.Menu mnMostraVinc 
         Caption         =   "Mostrar Vinculações"
      End
      Begin VB.Menu mnSelecPar 
         Caption         =   "Selecionar"
      End
      Begin VB.Menu mnVisDoc 
         Caption         =   "Buscar Parcelas Deste Documento"
      End
      Begin VB.Menu mnVisAcertoVenda 
         Caption         =   "&Visualizar Acerto/Venda"
      End
      Begin VB.Menu mnExclDoc 
         Caption         =   "E&xcluir Documento"
      End
      Begin VB.Menu mnBaixas 
         Caption         =   "&Baixas"
      End
   End
   Begin VB.Menu mnImprimir 
      Caption         =   "&Imprimir"
      Begin VB.Menu mnImpr 
         Caption         =   "&Imprimir Grade (Gera HTML)"
      End
      Begin VB.Menu mnImprCarne 
         Caption         =   "Emitir Carnê"
      End
      Begin VB.Menu mnImprContr 
         Caption         =   "Emitir Contrato"
      End
      Begin VB.Menu mnRecibo 
         Caption         =   "Imprimir Recibo"
      End
   End
End
Attribute VB_Name = "frmEditaMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim sListaDoc As String
Dim bActive As Boolean
Dim bDesativaBusca As Boolean


Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If bActive Then
        moveData
    End If
End Sub
Private Sub buscaTel()
    Dim strSQL As String
    
    On Error Resume Next
    With ado
        
        
        'lblTel.caption = F_BuscaTel(!cod_ter, !tipo_ter)
        
        
        
    
    End With
End Sub

Private Sub ado_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If bActive Then
        moveData
    End If
End Sub
Sub moveData()
    If bDesativaBusca Then Exit Sub
    buscaBaixas
    buscaOrg
    buscaTel
End Sub
Private Sub chkAberto_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub chkProrr_Click()
    cmndbtnBuscar_Click
End Sub


Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub chkRepExc_Click()

End Sub

Private Sub cmbxOrder_Click()
    If Not bActive Then Exit Sub
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub cmndbtnBaixar_Click2()
    Dim cJuros As Currency
    Dim iCount As Integer
    Dim jCount As Integer
    On Error GoTo trata_Erro
    With ado
        cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
        
        frmBaixaDocReq.tag = !cod_ter
        
                
        frmBaixaDocReq.txtbxCliente.text = !nome
        frmBaixaDocReq.txtbxCliente.tag = !codigo
        frmBaixaDocReq.txtbxJuros.text = FormatVal(cJuros)
        frmBaixaDocReq.txtbxReman.text = frmListaDoc.lblTotal.caption
        frmBaixaDocReq.txtbxValor_Par.text = frmListaDoc.lblTotal.caption
        'frmBaixaDocReq.lblParc.caption = "Parcelas: " & frmListaDoc.msflxgrd.rows
        frmBaixaDocReq.txtbxTotalBaixa.text = "0,00"
        frmBaixaDocReq.cmndbtnConfirma.tag = "R"
        frmBaixaDocReq.lblCPF.caption = lblCPF.caption
        frmBaixaDocReq.msflxParcelas.Cols = frmListaDoc.msflxgrd.Cols
        frmBaixaDocReq.msflxParcelas.rows = frmListaDoc.msflxgrd.rows + 1
        For iCount = 0 To frmListaDoc.msflxgrd.rows - 1
            frmBaixaDocReq.msflxParcelas.row = iCount + 1
            frmListaDoc.msflxgrd.row = iCount
            For jCount = 0 To frmListaDoc.msflxgrd.Cols - 1
                frmBaixaDocReq.msflxParcelas.col = jCount
                frmListaDoc.msflxgrd.col = jCount
                frmBaixaDocReq.msflxParcelas.text = frmListaDoc.msflxgrd.text
            Next
            
        Next
        frmBaixaDocReq.Show 1
        Unload frmBaixaDocReq
        Unload frmListaDoc
        ado.Requery
    End With
trata_Erro:
    P_MostraErro
    Exit Sub
End Sub



Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrd, ado
End Sub

Private Sub cmdSoma_Click()
    soma
End Sub

Private Sub cmdVinc_Click()
    frmVinculaDoc.txtbxExpr.text = ado!nome
    frmVinculaDoc.txtbxExpr.tag = ado!cod_ter
    frmVinculaDoc.Show 1
    
End Sub

Private Sub cmndbrnPrejuizo_Click()
    On Error GoTo trata_Erro
    Dim strSQL As String
    With ado
        If Not F_senha() Then
            Exit Sub
        End If
        If strExcldoc_usu <> "S" Then
            MsgBox "Você não tem permissão para esta operação!", vbExclamation
            Exit Sub
        End If
        If MsgBox("Confirma Inclusão do Carnê nos Prejuízos?", vbYesNo + vbQuestion) = vbYes Then
            If !prejuizo_itd = "S" Then
                strSQL = "Update  " & TAB_ITDOC & "  set prejuizo_itd='N' where seq_itd='" & !seq_itd & "'"
                conn.Execute strSQL
            Else
                strSQL = "Update  " & TAB_ITDOC & "  set prejuizo_itd='S' where seq_itd='" & !seq_itd & "'"
                conn.Execute strSQL
            End If
        End If
        cmndbtnBuscar_Click
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnBaixar_Click()
    On Error GoTo trata_Erro
    If strBaixadoc_usu <> "S" Then
        MsgBox "Não tem permissão pra baixar documento!", vbExclamation
        Exit Sub
    End If
    With ado
        frmBaixaDocReq.tag = !cod_ter
        frmBaixaDocReq.txtbxCliente.text = !nome
        
        frmBaixaDocReq.txtbxCliente.tag = !codigo
        frmBaixaDocReq.lblCPF.caption = Format(!cpf)
        frmBaixaDocReq.txtbxSeq.text = !seq_doc
        frmBaixaDocReq.Show 1
        
        Unload frmBaixaDocReq
    End With
    Exit Sub
trata_Erro:
P_MostraErro
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    Dim sCampo As String
    On Error GoTo trata_Erro
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    
    
    bDesativaBusca = True
    
    P_Status "Aguarde... buscando dados..."
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "  venc_itd,nome,parcela_itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "  nome,venc_itd,parcela_itd"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "  num_doc,venc_itd,parcela_itd"
    Else
        sOrderBy = "  valor_itd desc,venc_itd"
    End If
    
    
    If txtbxdataFin.text = "" And txtbxDataIni.text = "" Then
        If txtNFe.text = "" And txtBoleto.text = "" And txtbxNumDoc.text = "" And _
            txtbxSeqDoc.text = "" And txtbxTerceiro.text = "" And txtbxTerceiro.tag = "" And _
            txtbxSeqItd.text = "" And txtbxHist.text = "" Then
            MsgBox "Informe algum critério para busca!", vbExclamation
            Exit Sub
        End If
    End If
    
    
    
    
        strSQL = doc.F_SelectDocsReceber("C", txtbxDataIni.text, txtbxdataFin.text, _
            txtBoleto.text, txtbxTipoCob.tag, txtbxContaCob.tag, txtbxNumDoc.text, txtbxHist.text, txtbxSeqItd.text, txtbxSeqDoc.text, _
            "", "", txtbxTerceiro.tag, chkAberto.Value, sOrderBy, IIf(chkPrej.Value = 1, "S", ""), , , "", sListaDoc, chkVinc.Value = 1, _
            txtNFe.text)
        
    
    
    
    
    abreObjetoAdo ado, strSQL, dtgrd
    
    dtgrd.Visible = True
    
    lblTotal.caption = "clique"
    P_Status ""
    
    bDesativaBusca = False
    moveData
    bDesativaBusca = False
    
    Exit Sub
    
trata_Erro:
    bDesativaBusca = False
    P_Status
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnCarne_Click()
    imprimeCarne ado!seq_doc
End Sub

Private Sub cmndbtnCob_Click()
    buscaContaCob
End Sub
Sub buscaContaCob()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtbxContaCob.text)
    If c.cod_cob <> "" Then
        txtbxContaCob.text = c.nome_cob
        txtbxContaCob.tag = c.cod_cob
    End If
    
    
End Sub

Private Sub cmndbtnContrato_Click()
    'imprimeContratoHTML ado!seq_doc
'    imprimeContrato ado!seq_doc
    ApontarPrinterToPDF "Relatorio" & F_SoNumero(Now) & ".pdf"
    imprime Printer
    Printer.EndDoc
    
End Sub
Sub imprime(obj As Object)
    On Error GoTo trata_Erro
    dtgrd.Visible = False
    With ado
        .MoveFirst
        P_CabRelTimbr "RELAÇÃO DE DOCUMENTOS A RECEBER", obj, True
        P_ImprX obj, "Vencimento inicial: " & txtbxDataIni.text, 2
        
        P_ImprX obj, "Vencimento final: " & txtbxdataFin.text, 60, False
        P_ImprX obj, "Participante: " & txtbxTerceiro.text, 2, False
        obj.Print
cabecalho:
        P_line1 obj
        P_ImprX obj, "Número", 2, True, esquerda, 0, , 1
        P_ImprX obj, "Boleto", 20, True, esquerda, 0, , 0
        P_ImprX obj, "Data", 50, True, esquerda, 0, , 0
        P_ImprX obj, "Vencimento", 70, True, esquerda, 0, , 0
        P_ImprX obj, "Histórico", 90, True, esquerda, 0, , 0
        
        P_ImprX obj, "Valor", 170, True, direita, 0, , 0
        P_ImprX obj, "Restam", 192, True, direita, 0
        P_ImprX obj, "Parcela", 195, False, esquerda, 0
        P_line1 obj
        
        
        
        
        Do While Not .EOF
            
            
            
            
            If obj.CurrentY >= obj.ScaleHeight - 200 Then
                GoTo cabecalho
            End If
            
            P_ImprX obj, !num_doc, 2, True, esquerda
            P_ImprX obj, !boleto_itd, 20, True, esquerda
            P_ImprX obj, !data_doc, 50, True, esquerda
            P_ImprX obj, !venc_itd, 70, True, esquerda
            P_ImprX obj, !historico_doc, 90, True, esquerda
        
            P_ImprX obj, FormatVal(!valor_itd), 170, True, direita
            P_ImprX obj, FormatVal(!reman_itd), 192, True, direita
            P_ImprX obj, !parcela_itd & "/" & !parcelas_doc, 195, False, esquerda
            
            
            
            .MoveNext
        Loop
        
        .MoveFirst
    End With
    P_FimRel obj
    dtgrd.Visible = True
    Exit Sub
trata_Erro:
dtgrd.Visible = True
    P_MostraErro
End Sub
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEdiLan
    On Error GoTo trata_Erro
    
    
    With ado
        If .RecordCount < 0 Then
            Exit Sub
        End If
        fr.tag = !seq_itd
        
        fr.Show 1
        
        
        
        RefreshAdoPosition ado
        
        Exit Sub
        
        fr.lblData.caption = !data_doc
        fr.lblNUM_Doc.caption = !num_doc
        fr.lblParc.caption = !parcela_itd & "/" & !parcelas_doc
        fr.lblSeq_Doc.caption = !seq_doc
        fr.lblSeq_Itd.caption = !seq_itd
        fr.lblValor.caption = FormatVal(!valor_itd)
        fr.txtbxHist.text = !historico_doc
        fr.txtbxNome.text = !nome
        fr.txtbxNome.tag = !codigo
        fr.txtbxVenc.text = !venc_itd
        fr.txtbxValor.text = FormatVal(!reman_itd)
        fr.txtbxBoleto.text = !boleto_itd
        fr.txtbxClass.tag = !cod_pla
        fr.txtbxClass.text = !nome_pla
        fr.txtbxContaCob.text = !nome_cob
        fr.txtbxContaCob.tag = !cod_cob
        fr.txtbxTipoCob.text = !nome_tip
        
        fr.txtbxTipoCob.tag = !cod_tip
        fr.lblTipo.caption = "Crédito"
        fr.Show 1
        'adc.Refresh
        ado.Requery
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    If strExcldoc_usu <> "S" Then
        MsgBox "Você não pode excluir este documento!", vbExclamation
        Exit Sub
    End If
    
    With ado
        If !num_ven = "" Then
            If F_BuscaUmCampo("reman_itd", " " & TAB_ITDOC & " ", "reman_itd<valor_itd and seq_doc='" & !seq_doc & "'") = "" Then
                If MsgBox("Deseja realmente excluir o documento e suas parcelas?", vbYesNo + vbCritical) = vbYes Then
                    strSQL = "delete from " & TAB_DOC & " where seq_doc='" & !seq_doc & "' "
                    conn.Execute strSQL
                    strSQL = "delete from " & TAB_ITDOC & " where seq_doc='" & !seq_doc & "' "
                    conn.Execute strSQL
                    strSQL = "delete from  " & TAB_CHEQUES & " where seq_doc='" & !seq_doc & "' "
                    conn.Execute strSQL
                    ado.Requery
                End If
            Else
                MsgBox "Este documento não pode ser excluído por que já teve uma parcela baixada!", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Este documento não pode ser excluído pois se refere a uma venda!", vbExclamation
            
        End If
    End With
    
    
    
    
    
    
End Sub


Private Sub cmndbtnImpr_Click()
    P_DtgrdToHtml "Contas a Receber de " & txtbxDataIni.text & " Até " & txtbxdataFin.text, dtgrd, ado, _
    "Impresso: " & Now & " por " & strNome_Usu, "3"
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub



Private Sub cmndbtnSair_Click()
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub DBGrid1_GotFocus()
    'DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    'DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    
    
        
            buscaCli
        
    
    If txtbxTerceiro.tag <> "" Then
        cmndbtnBuscar_Click
    End If
End Sub
Sub buscaFun()
    Dim c As funcionario.funcionario
    c = funcionario.F_SelectFuncionario(txtbxTerceiro.text)
        If c.cod_fun <> "" Then
            txtbxTerceiro.text = c.nome_fun
            txtbxTerceiro.tag = c.cod_fun
            lblCPF.caption = c.cpf_fun
        End If
End Sub

Sub buscaCli()
    Dim c As Clientes.cliente
    c = F_BuscaCliente(txtbxTerceiro.text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.text = c.nome_Cli
            txtbxTerceiro.tag = c.cod_cli
            lblCPF.caption = c.cnpj_Cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.text)
        If f.cod_For <> "" Then
            txtbxTerceiro.text = f.nome_For
            txtbxTerceiro.tag = f.cod_For
            lblCPF.caption = f.cnpj_for
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
            lblCPF.caption = r.cpf_rep
        End If
    
End Sub


Private Sub cmndbtnTerceiro_Click()
    
End Sub

Private Sub cmndbtnTipoCob_Click()
    buscaTipoCob
End Sub
Sub buscaTipoCob()
    Dim t As tipocob
    t = F_BuscaTipoCob(txtbxTipoCob.text)
    If t.cod_tip <> "" Then
        txtbxTipoCob.text = t.nome_tip
        txtbxTipoCob.tag = t.cod_tip
    End If
End Sub


'Private Sub dbgrd_DblClick1()
'        Dim cJuros As Currency
'        With ado
'        Dim sJur As String
'        If !reman_itd <= 0 Then
'            Exit Sub
'        End If
'        On Error Resume Next
'        cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
'        If cJuros < 0 Then cJuros = 0
'
'        If IsLoad("frmListaDoc") Then
'            If frmListaDoc.tag <> Me.Name And frmListaDoc.msflxGrd.rows > 0 Then
'                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
'                Unload frmListaDoc
'                Exit Sub
'
'            End If
'            If frmListaDoc.lblCliente.tag <> "" And _
'                frmListaDoc.lblCliente.tag <> !cod_ter Then
'                MsgBox "Você só pode selecionar parcelas do mesmo Terceiro!", vbExclamation
'                Exit Sub
'            End If
'            frmListaDoc.lblCliente.tag = !cod_ter
'            frmListaDoc.lblCliente.caption = !Nome
'
'            If F_InFlxgrd(frmListaDoc.msflxGrd, !seq_itd, 0) = -1 Then
'                frmListaDoc.msflxGrd.AddItem !seq_itd & Chr(9) & FormatVal(!reman_itd) & Chr(9) & FormatVal(cJuros) _
'                & Chr(9) & !num_doc & Chr(9) & !cod_pla & Chr(9) & FormatVal(!reman_itd) & Chr(9) & !seq_doc
'            End If
'            frmListaDoc.soma
'        Else
'
'            frmListaDoc.tag = Me.Name
'            frmListaDoc.lblCliente.tag = !cod_ter
'            frmListaDoc.lblCliente.caption = !Nome
'            frmListaDoc.caption = "Documentos"
'            frmListaDoc.msflxGrd.Cols = 7
'            frmListaDoc.msflxGrd.rows = 0
'            frmListaDoc.msflxGrd.AddItem !seq_itd & Chr(9) & FormatVal(!reman_itd) & Chr(9) & FormatVal(cJuros) _
'                & Chr(9) & !num_doc & Chr(9) & !cod_pla & Chr(9) & FormatVal(!reman_itd) & Chr(9) & !seq_doc
'            frmListaDoc.Show 0, mdiFrm
'            frmListaDoc.soma
'        End If
'
'    End With
'
'End Sub

Private Sub dbgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        mnSelec_Click
    End If
End Sub



'Private Sub adc_Reposition()
'    Dim strSQL As String
'    On Error Resume Next
'    If bActive = False Then Exit Sub
'    With ado
'
'        If .RecordCount > 0 Then
'            If !reman_itd = 0 Then
'                cmndbtnEdit.Enabled = False
'                cmndbtnBaixar.Enabled = False
'            Else
'                cmndbtnEdit.Enabled = True
'                cmndbtnBaixar.Enabled = True
'            End If
'
'        End If
'        If !tipo_doc = "R" Then
'            cmndbtnCarne.Enabled = True
'        Else
'            cmndbtnCarne.Enabled = False
'        End If
'        lblTel.caption = F_BuscaTel(!cod_ter, !tipo_ter)
'        buscaBaixas
'        buscaOrg
'        buscaVias
'    End With
'End Sub
Sub buscaVias()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    'lblVias.caption = F_BuscaUmCampo("count(*)", "docimpress", "seq_doc='" & ado!seq_doc & "'")
    
    
    
End Sub
Sub buscaOrg()
'    Dim strSQL As String
'    Dim rPesq As New ADODB.Recordset
'    On Error GoTo trata_Erro:
'    strSQL = orgao.F_SelectOrgaoCli(ado!cod_ter)
'    rPesq.Open strSQL, conn
'    With rPesq
'        If .RecordCount > 0 Then
'            lblNomeOrg.caption = !nome_org
'            lblEnd.caption = !endereco_org & " " & !nome_set & " " & !nome_cid & "/" & !uf_cid & "(" & !fone_org & ")"
'
'        Else
'            lblNomeOrg.caption = "Orgão não localizado"
'            lblEnd.caption = ""
'        End If
'        .Close
'    End With
'
'    Exit Sub
'trata_Erro:
End Sub
Sub buscaBaixas()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    On Error GoTo trata_Erro
    Dim sMsg As String
    
    'FechaObjeto rPesq
    
   ' With ado
   '     sMsg = "Parcelas" & Chr(13)
   '
    '    txtbxBaixas.Text = sMsg
   '   strSQL = "select * from " & TAB_BAIXAFIN & " where seq_itd='" & !seq_itd & "' "
   '  dtBaixas.RecordSource = strSQL
   ' dtBaixas.Refresh
    'End With
    Exit Sub
trata_Erro:
End Sub

Private Sub dtgrd_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo trata_Erro
    
    
    Dim s As String
    s = dtgrd.Columns(ColIndex).DataField
          ado.Sort = s
        
 
    
    
trata_Erro:
    
End Sub

Private Sub dtgrd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnParcela
    End If
End Sub

Private Sub Form_Activate()
    
    If bActive = True Then
        P_Status
        Exit Sub
    End If
    
    'cmndbtnBuscar_Click
    bActive = True
    
    If Me.tag <> "" Then
        cmndbtnBuscar_Click
    End If
    
    sstb.Tab = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyF5 Then
        cmndbtnBuscar_Click
    End If
End Sub
Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    On Error GoTo trata_Erro
    
    
    Dim rPesq As New ADODB.Recordset
    bActive = False
    
    P_Status "Aguarde"
    strSQL = ado.Source
    abreObjetoAdo rPesq, strSQL
    
    
    dtgrd.Visible = False
    With rPesq
        'dbgrd.caption = .RecordCount & " Documentos Listados"
        
        
        Do While Not .EOF
            cTotal = cTotal + !reman_itd
            On Error GoTo trata_Erro
                
            
            .MoveNext
            P_Status "Somando " & .AbsolutePosition
        Loop
        .Close
        
        
        
    End With
    P_Status
    lblTotal.caption = FormatVal(cTotal)
    bActive = True
    dtgrd.Visible = True
    Exit Sub
trata_Erro:
    P_Status
    lblTotal.caption = FormatVal(cTotal) & "Err"
    dtgrd.Visible = True
    P_Status
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    On Error Resume Next
    bActive = False
    resizeForm Me
    txtbxDataIni.text = DateAdd("D", -30, Date)
    
    txtbxdataFin.text = Date
    'adc.ConnectionString = Conn.ConnectionString
    'cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    'adc.ConnectionString = Conn.ConnectionString
    'dtBaixas.Connect = Banco.Connect
    'cmbxBuscar.ListIndex = 1
    cmbxOrder.ListIndex = 0
    'cmbxTipo.ListIndex = 0
    'cmbxTipoTer.ListIndex = 0
    'cmndbtnBuscar_Click
    Set ado = New ADODB.Recordset
        
End Sub


Private Sub Form_Resize()
    'On Error Resume Next
    dtgrd.width = Me.width - dtgrd.Left - 300
    'dtgrdChe.Top = pcTotal.Top - dbgrdBaixas.Height - 100
    dtgrd.height = pcTotal.Top - dtgrd.Top - 100
    'dbgrdChe.Width = dbgrd.Width
    PosicionaGrade dtgrd, Me, pcTotal
    posicionaADireita cmndbtnSair, pcBot
End Sub

Private Sub lblTotal_Click()
    soma
End Sub

Private Sub mnBaixar_Click()
    cmndbtnBaixar_Click
End Sub

Private Sub mnBaixas_Click()
    'On Error Resume Next
    frmRelCarnesRecebidos.txtbxSeqItd.text = ado!seq_itd
    frmRelCarnesRecebidos.tag = "x"
    frmRelCarnesRecebidos.Show 1
    

End Sub

Private Sub mnEdiPar_Click()
    cmndbtnEdit_Click
End Sub

Private Sub mnExclDoc_Click()
    cmndbtnExcl_Click
End Sub

Private Sub mnImpr_Click()
    cmndbtnImpr_Click
End Sub

Private Sub mnImprCarne_Click()
    cmndbtnCarne_Click
End Sub

Private Sub mnImprContr_Click()
    cmndbtnContrato_Click
End Sub

Private Sub mnMostraVinc_Click()
    On Error Resume Next
    P_DocsVinculados ado!seq_doc
End Sub

Private Sub mnNova_Click()
    mnNova_Click
End Sub

Private Sub mnRecibo_Click()
    On Error Resume Next
    frmRelCarnesRecebidos.txtbxNumDoc.text = ado!num_doc
    frmRelCarnesRecebidos.Show 1
End Sub

Private Sub mnSair_Click()
    Unload Me
End Sub

Private Sub mnSelec_Click()
    If Me.tag <> "" Then
        Me.tag = "X"
        Me.Hide
    End If
End Sub

Private Sub mnSelecPar_Click()
    dbgrd_DblClick
End Sub

Private Sub mnSPC_Click()
    On Error Resume Next
    P_ParcelaSpc ado!cod_ter, ado!seq_itd, ado!reman_itd
End Sub

Private Sub mnVisAcertoVenda_Click()
    On Error GoTo trata_Erro:
    With ado
        If !esp_doc = "A" Then
            P_VisualizaAcerto !num_ven
        ElseIf !esp_doc = "V" Then
            ImprVenda !num_ven
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub mnVisDoc_Click()
    On Error GoTo trata_Erro
    With ado
        txtbxSeqDoc.text = !seq_doc
        
        txtbxSeqItd.text = ""
        'txtbx.Text = ""
        txtbxdataFin.text = ""
        txtbxDataIni.text = ""
        cmndbtnBuscar_Click
        txtbxSeqDoc.text = ""
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub pcLista_Click()
    sListaDoc = BuscaListaDoc
    
End Sub

Private Sub pcTop_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub txtBoleto_GotFocus()
    P_SelecTexto txtBoleto
End Sub

Private Sub txtBoleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBoleto_LostFocus()
    P_AjustaTexto txtBoleto
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
    If txtbxdataFin.text <> "" Then
        
    End If
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()

End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        cmndbtnTerc_Click
    End If
End Sub




Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub
Private Sub txtbxNumDoc_GotFocus()
    P_SelecTexto txtbxNumDoc
End Sub

Private Sub txtbxNumDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNumDoc_LostFocus()
    P_AjustaTexto txtbxNumDoc
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    If txtbxValor.text <> "" Then
        P_ValorLostFocus txtbxValor
    End If
End Sub


Private Sub txtbxSeqDoc_GotFocus()
    P_SelecTexto txtbxSeqDoc
End Sub

Private Sub txtbxSeqDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
        cmndbtnBuscar_Click
    End If
End Sub

Private Sub txtbxSeqDoc_LostFocus()
    txtbxSeqDoc.text = Format(txtbxSeqDoc.text, "000000#")
End Sub

Private Sub txtbxSeqItd_GotFocus()
    P_SelecTexto txtbxSeqItd
End Sub

Private Sub txtbxSeqItd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxSeqItd_LostFocus()
    txtbxSeqItd.text = Format(txtbxSeqItd.text, "000000#")
End Sub

Private Sub txtbxTerceiro_Change()
    txtbxTerceiro.tag = ""
    lblCPF.caption = ""
    lblCPF.tag = ""

End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.tag = ""
End Sub




Private Sub txtNFe_GotFocus()
    P_SelecTexto txtNFe
End Sub

Private Sub txtNFe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
    End If
    
End Sub
