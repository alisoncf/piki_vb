VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEditaMov 
   Caption         =   "Movimentações"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   17685
   Icon            =   "frmEditaMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   17685
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dtgrd 
      Height          =   4575
      Left            =   0
      TabIndex        =   65
      Top             =   1440
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   8070
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
      Caption         =   "Lista de Documentos Encontrados"
      ColumnCount     =   22
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
         DataField       =   "valor_doc"
         Caption         =   "Valor Original"
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
      BeginProperty Column14 
         DataField       =   "diferenca_ace"
         Caption         =   "Bol."
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
         DataField       =   "nome_org"
         Caption         =   "Órgão"
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
      BeginProperty Column17 
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
      BeginProperty Column18 
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
      BeginProperty Column19 
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
      BeginProperty Column20 
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
      BeginProperty Column21 
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
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3614,74
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3030,236
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
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
            ColumnWidth     =   929,764
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   17655
      TabIndex        =   47
      Top             =   7635
      Width           =   17685
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   720
         Left            =   12600
         Picture         =   "frmEditaMov.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbrnPrejuizo 
         Caption         =   "Prejuízo"
         Height          =   720
         Left            =   9480
         Picture         =   "frmEditaMov.frx":0406
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdContrato 
         Caption         =   "Subir &Contrato"
         Height          =   720
         Left            =   14160
         Picture         =   "frmEditaMov.frx":0990
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnContrato 
         Caption         =   "Emitir Contrato"
         Height          =   720
         Left            =   7920
         Picture         =   "frmEditaMov.frx":0AEC
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnCarne 
         Caption         =   "Emitir &Carnê"
         Height          =   720
         Left            =   6360
         Picture         =   "frmEditaMov.frx":0BE8
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   720
         Left            =   3240
         Picture         =   "frmEditaMov.frx":0CE4
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   720
         Left            =   11040
         Picture         =   "frmEditaMov.frx":0E2E
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   720
         Left            =   1680
         Picture         =   "frmEditaMov.frx":0F78
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   720
         Left            =   120
         Picture         =   "frmEditaMov.frx":1302
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   720
         Left            =   15720
         Picture         =   "frmEditaMov.frx":144C
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   720
         Left            =   4800
         Picture         =   "frmEditaMov.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FCB683&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      ScaleHeight     =   1440
      ScaleWidth      =   17685
      TabIndex        =   46
      Top             =   0
      Width           =   17685
      Begin VB.CheckBox chkInicial 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inicial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cmbxCreditoSuspenso 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":1DDF
         Left            =   14040
         List            =   "frmEditaMov.frx":1DEC
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkCliOrgaos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Somente vinculados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9960
         TabIndex        =   70
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox cmbxListar 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":1DFC
         Left            =   7680
         List            =   "frmEditaMov.frx":1E09
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtBoleto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9360
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12120
         TabIndex        =   40
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox pcLista 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         Picture         =   "frmEditaMov.frx":1E2B
         ScaleHeight     =   195
         ScaleWidth      =   180
         TabIndex        =   66
         Top             =   0
         Width           =   180
      End
      Begin VB.TextBox txtbxOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9360
         TabIndex        =   36
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtbxGrupo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8760
         TabIndex        =   33
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtbxValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   60
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxHist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxRepr 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   12120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxEspDoc 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmEditaMov.frx":1F24
         Left            =   10560
         List            =   "frmEditaMov.frx":1F34
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxSeqItd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   60
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxSeqDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   60
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxTipoCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   30
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   7320
         Picture         =   "frmEditaMov.frx":1F55
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnCob 
         Height          =   285
         Left            =   9000
         Picture         =   "frmEditaMov.frx":22DF
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":2669
         Left            =   2400
         List            =   "frmEditaMov.frx":2676
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":26A8
         Left            =   14040
         List            =   "frmEditaMov.frx":26B8
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtbxNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   60
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Pesquisar (F5)"
         Height          =   600
         Left            =   15480
         Picture         =   "frmEditaMov.frx":26E4
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   25
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnTipoCob 
         Height          =   285
         Left            =   7320
         Picture         =   "frmEditaMov.frx":30E6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnRep 
         Enabled         =   0   'False
         Height          =   285
         Left            =   13680
         Picture         =   "frmEditaMov.frx":3470
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnGru 
         Height          =   285
         Left            =   10080
         Picture         =   "frmEditaMov.frx":37FA
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnOrg 
         Height          =   285
         Left            =   11760
         Picture         =   "frmEditaMov.frx":3B84
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdCidade 
         Height          =   285
         Left            =   13680
         Picture         =   "frmEditaMov.frx":3F0E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito Suspenso:"
         Height          =   195
         Left            =   14040
         TabIndex        =   72
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listar/Status:"
         Height          =   195
         Left            =   7680
         TabIndex        =   68
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblTerCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   195
         Left            =   4320
         TabIndex        =   29
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Boleto:"
         Height          =   195
         Left            =   9360
         TabIndex        =   15
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   12120
         TabIndex        =   39
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblOrgao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Órgão:"
         Height          =   195
         Left            =   9360
         TabIndex        =   35
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblGru 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Clientes:"
         Height          =   195
         Left            =   6000
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   3360
         TabIndex        =   3
         Tag             =   "!"
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblRepr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representante:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   12120
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie de Doc:"
         Height          =   195
         Left            =   10560
         TabIndex        =   17
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         Caption         =   "CPF"
         Height          =   195
         Left            =   4920
         TabIndex        =   56
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Parc:"
         Height          =   195
         Left            =   5160
         TabIndex        =   7
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id. Doc:"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   0
         Width           =   570
      End
      Begin VB.Label lblTipoCob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança:"
         Height          =   195
         Left            =   7680
         TabIndex        =   12
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cobrança:"
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   2400
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar Por:"
         Height          =   195
         Left            =   14040
         TabIndex        =   38
         Tag             =   "!"
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Doc."
         Height          =   195
         Left            =   2160
         TabIndex        =   27
         Tag             =   "!"
         Top             =   0
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   24
         Tag             =   "!"
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Tag             =   "!"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   660
      End
   End
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   17655
      TabIndex        =   44
      Top             =   6300
      Width           =   17685
      Begin VB.PictureBox pcMaximiza 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   89
         Top             =   0
         Width           =   135
      End
      Begin VB.PictureBox pcAnot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   10560
         ScaleHeight     =   1335
         ScaleWidth      =   5295
         TabIndex        =   84
         Top             =   0
         Width           =   5295
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridAnot 
            Height          =   975
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1720
            _Version        =   393216
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Anotações do cliente"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   2040
         ScaleHeight     =   1335
         ScaleWidth      =   1575
         TabIndex        =   78
         Top             =   0
         Width           =   1575
         Begin VB.CheckBox chkRepExc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Rep Excl"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkPrej 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Prejuízo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkProrr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Prorrogadas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkAberto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Em Aberto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Filtrar parcelas:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.PictureBox pcBaixas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   15840
         ScaleHeight     =   1335
         ScaleWidth      =   4095
         TabIndex        =   76
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
         Begin MSDataGridLib.DataGrid gridBai 
            Height          =   1215
            Left            =   120
            TabIndex        =   77
            Top             =   0
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2143
            _Version        =   393216
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
            Caption         =   "Baixa"
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "data_bai"
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
            BeginProperty Column01 
               DataField       =   "cod_cob"
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
            BeginProperty Column02 
               DataField       =   "valor_bai"
               Caption         =   "Valor"
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
            BeginProperty Column04 
               DataField       =   "cod_cai"
               Caption         =   "Caixa"
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
                  ColumnWidth     =   1094,74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1289,764
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   975,118
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   870,236
               EndProperty
               BeginProperty Column04 
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton cmdSoma 
         Caption         =   "Soma:"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   8160
         ScaleHeight     =   1335
         ScaleWidth      =   2415
         TabIndex        =   62
         Top             =   0
         Width           =   2415
         Begin VB.TextBox txtFones 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   600
            MultiLine       =   -1  'True
            TabIndex        =   75
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fones:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.PictureBox pcTrab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   3600
         ScaleHeight     =   1335
         ScaleWidth      =   4575
         TabIndex        =   58
         Top             =   0
         Width           =   4575
         Begin VB.Label lblNomeOrg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orgão"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   840
            TabIndex        =   60
            Top             =   30
            Width           =   525
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Trabalho:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   45
            Width           =   675
         End
         Begin VB.Label lblEnd 
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   4020
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vias:"
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   840
         Width           =   345
      End
      Begin VB.Label lblVias 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   480
         TabIndex        =   87
         Top             =   840
         Width           =   120
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
         Left            =   120
         TabIndex        =   45
         Top             =   480
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
      Begin VB.Menu mnExportXML 
         Caption         =   "Exportar para XML"
      End
      Begin VB.Menu mnExportCSV 
         Caption         =   "Exportar para CSV"
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
      Begin VB.Menu mnSelecPar 
         Caption         =   "Selecionar"
      End
      Begin VB.Menu mnProrrogTodas 
         Caption         =   "Prorrogar todas as parcelas vencidas deste doc para o dia 30"
         Shortcut        =   ^T
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
      Begin VB.Menu mnCliente 
         Caption         =   "&Edita Cliente"
      End
      Begin VB.Menu abrirAnotacoes 
         Caption         =   "Abrir Anotações"
      End
      Begin VB.Menu mnPesqInternet 
         Caption         =   "Pesquisar na internet"
      End
      Begin VB.Menu mnEnviarDoc 
         Caption         =   "Enviar este doc para Internet"
      End
      Begin VB.Menu mnEnviarTodos 
         Caption         =   "Enviar todos desta lista para Internet"
      End
      Begin VB.Menu enviarFila 
         Caption         =   "Enviar todos para fila de transferência"
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
   End
   Begin VB.Menu mnCons 
      Caption         =   "&Consultas "
      Begin VB.Menu mnConsOutra 
         Caption         =   "Consultar em Outra loja"
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
Public WithEvents adoBaixa As ADODB.Recordset
Attribute adoBaixa.VB_VarHelpID = -1
Dim sListaDoc As String
Dim bActive As Boolean
Dim bDesativaBusca As Boolean

Private Sub abrirAnotacoes_Click()
    If ado!tipo_ter = "C" Then AbreAnotacoesCli ado!cod_ter
End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If bActive Then
        moveData
    End If
End Sub
Private Sub buscaTel()
    Dim strSQL As String
    
    On Error Resume Next
    With ado
        
        
        txtFones.text = F_BuscaTel(!cod_ter, !tipo_ter, CrLf)
        
        
        
    
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
    buscaAnot
End Sub
Private Sub chkAberto_Click()
    cmndbtnBuscar_Click
    If chkAberto.Value = 1 Then
        pcBaixas.Visible = False
    Else
        pcBaixas.Visible = True
    End If
End Sub


Private Sub chkProrr_Click()
    cmndbtnBuscar_Click
End Sub


Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxEspDoc_Click()
    If Left(cmbxEspDoc.text, 1) = "A" Then
        cmndbtnRep.Enabled = True
        txtbxRepr.Enabled = True
        lblRepr.Enabled = True
        txtbxRepr.SetFocus
    Else
        cmndbtnRep.Enabled = False
        txtbxRepr.Enabled = False
        lblRepr.Enabled = False
    End If
End Sub

Private Sub cmbxListar_Change()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxListar_Click()
    cmndbtnBuscar_Click
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
        frmBaixaDocReq.lblcpf.caption = lblcpf.caption
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

Private Sub cmbxTipoTer_Click()
    lblTerCap.caption = "&" & cmbxTipoTer.text
    If Left(cmbxTipoTer.text, 1) = "C" Then
        txtbxGrupo.Enabled = True
        txtbxOrgao.Enabled = True
        lblOrgao.Enabled = True
        lblGru.Enabled = True
        txtbxGrupo.Enabled = True
        txtbxOrgao.Enabled = True
    Else
        txtbxGrupo.Enabled = False
        txtbxOrgao.Enabled = False
        lblOrgao.Enabled = False
        lblGru.Enabled = False
        txtbxGrupo.Enabled = False
        txtbxOrgao.Enabled = False
    
    End If
End Sub

Private Sub cmdCidade_Click()
    Dim c As Cidade
    c = Cidades.F_BuscaCidades(txtCidade.text)
    txtCidade.text = c.nome_cid
    txtCidade.tag = c.cod_cid
    
    
    
End Sub

Private Sub cmdContrato_Click()
    On Error GoTo trata_Erro:
    frmContratoUpload.tag = ado!seq_doc
    frmContratoUpload.Show 1
trata_Erro:
End Sub

Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrd, ado
End Sub

Private Sub cmdSoma_Click()
    soma
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
                strSQL = "Update tab_itdoc set prejuizo_itd='N' where seq_itd='" & !seq_itd & "'"
                Conn.Execute strSQL
            Else
                strSQL = "Update tab_itdoc set prejuizo_itd='S' where seq_itd='" & !seq_itd & "'"
                Conn.Execute strSQL
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
        
        frmBaixaDocReq.txtbxCliente.text = !nome
        frmBaixaDocReq.txtbxCliente.tag = !codigo
        
        frmBaixaDocReq.lblcpf.caption = !cpf
        frmBaixaDocReq.txtbxSeq.text = !seq_doc
        frmBaixaDocReq.busca
        frmBaixaDocReq.Show 1
        
        Unload frmBaixaDocReq
        
        .Requery
        
        
    End With
trata_Erro:
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
    Dim tipoDat As tipo_vencimento
    
    bDesativaBusca = True
    
    P_Status "Aguarde... buscando dados..."
    
    If chkInicial.Value = 1 Then
        tipoDat = vencimentoInicial
    Else
        tipoDat = vencimento_da_parcela
    End If
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "  venc_itd,nome,parcela_itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "  nome,venc_itd,parcela_itd"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "  num_doc,venc_itd,parcela_itd"
    Else
        sOrderBy = "  valor_itd desc,venc_itd"
    End If
    
    
    
    If Left(cmbxEspDoc.text, 1) = "A" Then
        strSQL = doc.F_SelectDocsReceberAcerto(Left(cmbxTipoTer.text, 1), txtbxDataIni.text, txtbxdataFin.text, _
         txtBoleto.text, txtbxTipoCob.tag, txtbxContaCob.tag, txtbxNumDoc.text, txtbxHist.text, txtbxSeqItd.text, txtbxSeqDoc.text, _
        txtbxGrupo.tag, "", txtbxTerceiro.tag, chkAberto.Value, "", txtbxRepr.tag, sOrderBy, IIf(chkPrej.Value = 1, "S", ""), tipoDat, "", sListaDoc, txtCidade.tag, txtbxValor.text, cmbxListar.ListIndex)
    Else
        If txtbxGrupo.tag <> "" Or txtbxOrgao.tag <> "" Then
            strSQL = doc.F_SelectDocsReceberGruposOrgaos(Left(cmbxTipoTer.text, 1), txtbxDataIni.text, txtbxdataFin.text, _
            txtBoleto.text, txtbxTipoCob.tag, txtbxContaCob.tag, txtbxNumDoc.text, txtbxHist.text, txtbxSeqItd.text, txtbxSeqDoc.text, _
             txtbxGrupo.tag, txtbxOrgao.tag, "", txtbxTerceiro.tag, chkAberto.Value, sOrderBy, IIf(chkPrej.Value = 1, "S", ""), tipoDat, "", sListaDoc, txtCidade.tag, txtbxValor.text, cmbxListar.ListIndex)
        Else
            strSQL = doc.F_SelectDocsReceber(Left(cmbxTipoTer.text, 1), txtbxDataIni.text, txtbxdataFin.text, _
            txtBoleto.text, txtbxTipoCob.tag, txtbxContaCob.tag, txtbxNumDoc.text, txtbxHist.text, txtbxSeqItd.text, txtbxSeqDoc.text, _
            txtbxGrupo.tag, "", txtbxTerceiro.tag, chkAberto.Value, sOrderBy, IIf(chkPrej.Value = 1, "S", ""), tipoDat, , Left(cmbxEspDoc.text, 1), sListaDoc, txtCidade.tag, txtbxValor.text, _
                cmbxListar.ListIndex, chkCliOrgaos.Value = 1, "", Left(cmbxCreditoSuspenso.text, 1))
        End If
    End If
    
    
    
    abreObjetoAdo ado, strSQL, dtgrd
    dtgrd.caption = ado.RecordCount & " Registros encontrados"
    dtgrd.Visible = True
    
    lblTotal.caption = "clique"
    P_Status ""
    
    bDesativaBusca = False
    moveData
    bDesativaBusca = False
    P_Status ""
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
    imprimeContrato ado!seq_doc
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
        'fr.lblTipo.caption = "Crédito"
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
            If F_BuscaUmCampo("reman_itd", "tab_itdoc", "reman_itd<valor_itd and seq_doc='" & !seq_doc & "'") = "" Then
                If MsgBox("Deseja realmente excluir o documento e suas parcelas?", vbYesNo + vbCritical) = vbYes Then
                    strSQL = "delete from tab_doc where seq_doc='" & !seq_doc & "' "
                    Conn.Execute strSQL
                    strSQL = "delete from tab_itdoc where seq_doc='" & !seq_doc & "' "
                    Conn.Execute strSQL
                    strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "' "
                    Conn.Execute strSQL
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

Private Sub cmndbtnGru_Click()
    Dim g As grupoCli
    g = F_BuscaGrupoCli
    txtbxGrupo.text = g.nome_gru
    txtbxGrupo.tag = g.cod_gru
    
End Sub

Private Sub cmndbtnImpr_Click()
    P_DbgrdToHtml "Contas a Receber de " & txtbxDataIni.text & " Até " & txtbxdataFin.text, dtgrd, ado, _
    "Impresso: " & Now & " por " & strNome_Usu, "3"
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 1
End Sub

Private Sub cmndbtnOrg_Click()
    Dim o As orgao.orgao
    o = F_BuscaOrgao(txtbxOrgao.text, txtbxGrupo.tag)
    
    txtbxOrgao.text = o.nome_org
    txtbxOrgao.tag = o.cod_org
    
    
    
End Sub

Private Sub cmndbtnRep_Click()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxRepr.text)
        If r.cod_Rep <> "" Then
            txtbxRepr.text = r.nome_rep
            txtbxRepr.tag = r.cod_Rep
            
        End If
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
    
    
        If cmbxTipoTer.ListIndex = 0 Then
            buscaCli
        ElseIf cmbxTipoTer.ListIndex = 2 Then
            buscaFun
        Else
            buscaRep
        End If
    
    
    cmndbtnBuscar_Click
End Sub
Sub buscaFun()
    Dim c As funcionario.funcionario
    c = funcionario.F_SelectFuncionario(txtbxTerceiro.text)
        If c.cod_fun <> "" Then
            txtbxTerceiro.text = c.nome_fun
            txtbxTerceiro.tag = c.cod_fun
            lblcpf.caption = c.cpf_fun
        End If
End Sub

Sub buscaCli()
    Dim c As cliente
    c = F_BuscaCliente(txtbxTerceiro.text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.text = c.nome_cli
            txtbxTerceiro.tag = c.cod_cli
            lblcpf.caption = c.cnpj_Cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.text)
        If f.cod_For <> "" Then
            txtbxTerceiro.text = f.nome_for
            txtbxTerceiro.tag = f.cod_For
            lblcpf.caption = f.cnpj_for
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.text)
        If r.cod_Rep <> "" Then
            txtbxTerceiro.text = r.nome_rep
            txtbxTerceiro.tag = r.cod_Rep
            lblcpf.caption = r.cpf_rep
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
    lblVias.caption = F_BuscaUmCampo("count(*)", "tab_docimpress", "seq_doc='" & ado!seq_doc & "'")
    
    
    
End Sub
Sub buscaOrg()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro:
    strSQL = orgao.F_SelectOrgaoCli(ado!cod_ter)
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblNomeOrg.caption = !nome_org
            lblEnd.caption = !endereco_org & " " & !nome_set & " " & !nome_cid & "/" & !uf_cid & "(" & !fone_org & ")"
            
        Else
            lblNomeOrg.caption = "Orgão não localizado"
            lblEnd.caption = ""
        End If
        .Close
    End With
    
    Exit Sub
trata_Erro:
End Sub
Sub buscaBaixas()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    On Error GoTo trata_Erro
    Dim sMsg As String
    
    FechaObjeto rPesq
    
    With ado
        sMsg = "Parcelas" & Chr(13)
        
    '    txtbxBaixas.Text = sMsg
        strSQL = "select * from tab_baixaitd2 where seq_itd='" & !seq_itd & "' "
     abreObjetoAdo adoBaixa, strSQL, gridBai
     If adoBaixa.RecordCount > 0 Then
        pcBaixas.Visible = True
     Else
        pcBaixas.Visible = False
     End If
    End With
    Exit Sub
trata_Erro:
End Sub

Private Sub dtgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnParcela
    End If
End Sub

Private Sub enviarFila_Click()
         On Error GoTo trata_Erro
    dtgrd.Visible = False
    With ado
        .MoveFirst
        Do While Not .EOF
            
            insere_fila !seq_doc, 2
            .MoveNext
        Loop
        .MoveFirst
    End With
    P_Status
    dtgrd.Visible = True
trata_Erro:

End Sub

Private Sub Form_Activate()
Form_Resize
    If bActive = True Then
        P_Status
        Exit Sub
    End If
    
    'cmndbtnBuscar_Click
    bActive = True
    
    If Me.tag <> "" Then
        cmndbtnBuscar_Click
    End If
    
    
End Sub
Sub buscaAnot()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    If ado!tipo_ter = "C" Then
        strSQL = "select * from tab_clientesanot where cod_cli='" & ado!cod_ter & "' and excluido='N'  order by id desc "
        abreObjetoAdo rPesq, strSQL
    
        GridUtil.AjustaGridData gridAnot, 4, 1, 0, "id,anotacao,data,usuario", "id,anotacao,datahora,cod_usu", "0,4000,1800,1000", rPesq, "", False, vbActiveTitleBar, False
        rPesq.Close
    Else
        gridAnot.rows = 1
    End If
    Exit Sub
trata_Erro:
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
    Dim i As Integer
    
    Dim rPesq As New ADODB.Recordset
    bActive = False
    
    P_Status "Aguarde"
    strSQL = ado.Source
    'abreObjetoAdo rPesq, strSQL
    
    
    dtgrd.Visible = False
    With ado
        'dbgrd.caption = .RecordCount & " Documentos Listados"
        
        i = ado.AbsolutePosition
        ado.MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !reman_itd
            On Error GoTo trata_Erro
                
            
            .MoveNext
            P_Status "Somando " & .AbsolutePosition
        Loop
        
        ado.MoveFirst
        ado.Move i - 1
        
        
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
    txtbxDataIni.text = Date
    
    txtbxdataFin.text = Date
    'adc.ConnectionString = Conn.ConnectionString
    'cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    'adc.ConnectionString = Conn.ConnectionString
    
    'cmbxBuscar.ListIndex = 1
    cmbxOrder.ListIndex = 0
    'cmbxTipo.ListIndex = 0
    cmbxTipoTer.ListIndex = 0
    'cmndbtnBuscar_Click
    Set ado = New ADODB.Recordset
    Set adoBaixa = New ADODB.Recordset
End Sub


Private Sub Form_Resize()
   PosicionaGradeWidthVal dtgrd, Me.width, 300
   PosicionaGradeHeigthVal dtgrd, pcTotal.top, 100
    posicionaADireita cmndbtnSair, pcBot
   
End Sub

Private Sub gridAnot_DblClick()
    abrirAnotacoes_Click
End Sub

Private Sub gridBai_DblClick()
    mnBaixas_Click
End Sub

Private Sub Label2_dblClick()
    txtbxdataFin.text = ""
    txtbxDataIni.text = ""
End Sub

Private Sub Label3_DblClick()
    txtbxdataFin.text = ""
    txtbxDataIni.text = ""
End Sub

Private Sub lblEnd_DblClick()
    mnCliente_Click
End Sub

Private Sub lblTotal_Click()
    soma
End Sub

Private Sub mnBaixar_Click()
    cmndbtnBaixar_Click
End Sub

Private Sub mnBaixas_Click()
    On Error Resume Next
    frmRelCarnesRecebidos.txtbxSeqItd.text = adoBaixa!seq_itd
    frmRelCarnesRecebidos.txtbxDataIni.text = ""
    frmRelCarnesRecebidos.txtbxdataFin.text = ""
    
    frmRelCarnesRecebidos.tag = "x"
    frmRelCarnesRecebidos.Show 1
    
End Sub

Private Sub mnCliente_Click()
    On Error GoTo trata_Erro
    With ado
        If !tipo_ter = "C" Then
            frmCadCli.tag = !cod_ter
            frmCadCli.Show 0
        End If
    End With
trata_Erro:
    
End Sub

Private Sub mnEdiPar_Click()
    cmndbtnEdit_Click
End Sub

Private Sub mnEnviarDoc_Click()
    On Error GoTo trata_Erro
    API.EnviarDocInternet ado!seq_doc
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub mnEnviarTodos_Click()
    On Error GoTo trata_Erro
    With ado
        .MoveFirst
        
        Do While Not .EOF
            API.EnviarDocInternet ado!seq_doc, False
            P_Status "enviando " & .AbsolutePosition & "/" & .RecordCount
            .MoveNext
        Loop
        .MoveFirst
    End With
    P_Status
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub mnExclDoc_Click()
    cmndbtnExcl_Click
End Sub

Private Sub mnExportXML_Click()
    Dim Parametros As String
    Parametros = "cnpj=" & aspas & F_SoNumero(strCNPJ_Emp) & aspas & " qtd=" & aspas & ado.RecordCount & aspas
    Parametros = Parametros & " dth=" & aspas & Now & aspas
    ExportaGridToXML dtgrd, ado, Parametros
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

Private Sub mnNova_Click()
    mnNova_Click
End Sub



Private Sub mnPesqInternet_Click()
    Dim ret As String
    Dim rows As String
    Dim lista As String
    Dim msg As String
    On Error GoTo trata_Erro
    
    ret = API.APIBuscarDoc(strCod_Emp, ado!seq_doc)
    ret = AspasSimplesParaAspasDuplas(ret)
    rows = GetValorJson("countRows", ret)
    If rows = "0" Or rows = """0""" Then
        MsgBox "Este registro não foi localizando na Internet", vbExclamation
        Exit Sub
    Else
        frmDetalheDoc.json = ret
        frmDetalheDoc.Show 1
    End If
    
    
    
        
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub mnProrrogTodas_Click()
    Dim novoVenc As String
    If strProrrogarPar_Usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação.", vbExclamation
        Exit Sub
    End If
    novoVenc = "30/" & Format(Date, "mm/yyyy")
    
    If MsgBox("Confirma prorrogação de todas as parcelas vencidas  do DOC: " & ado!num_doc & " para " & novoVenc & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    
    P_GravaItDocPre ado!seq_itd, ado!venc_itd, novoVenc, "U:" & strNome_Usu & " - " & Format(ado!venc_itd, "dd/mm/yyyy") & "->" & Format(novoVenc, "dd/mm/yy") & "->" & Format(Now, "dd/mm/yy HH:mm")
    
   
  
    P_AlteraDocsAnteriores ado!seq_doc, Date, novoVenc
   
    
    cmndbtnBuscar_Click
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

Private Sub pcMaximiza_Click()
    If pcTotal.height > pcMaximiza.height Then
        pcTotal.height = pcMaximiza.height
        
    Else
        pcTotal.height = 1335
    End If
    
End Sub

Private Sub pcTop_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub pcTrab_DblClick()
    mnCliente_Click
End Sub

Private Sub Picture1_DblClick()
    mnCliente_Click
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


Private Sub txtbxGrupo_Change()
    txtbxGrupo.tag = ""
End Sub

Private Sub txtbxGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnGru_Click
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
        txtbxNumDoc_GotFocus
        txtbxNumDoc.SetFocus
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
Private Sub txtbxRepr_Change()
    txtbxRepr.tag = ""
End Sub

Private Sub txtbxRepr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnRep_Click
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
    lblcpf.caption = ""
    lblcpf.tag = ""

End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.tag = ""
End Sub

Private Sub txtCidade_Change()
    txtCidade.tag = ""
End Sub

Private Sub txtCidade_GotFocus()
    P_SelecTexto txtCidade
End Sub

Private Sub txtCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCidade_Click
    End If
End Sub
