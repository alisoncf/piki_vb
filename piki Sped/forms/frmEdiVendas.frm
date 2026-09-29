VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVendas 
   Caption         =   "Edição de Vendas"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   18960
   Icon            =   "frmEdiVendas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   18960
   Begin VB.PictureBox pcVenda 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   18930
      TabIndex        =   31
      Top             =   6660
      Width           =   18960
      Begin VB.PictureBox pcNfe 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   14760
         ScaleHeight     =   855
         ScaleWidth      =   4455
         TabIndex        =   43
         Top             =   0
         Width           =   4455
         Begin VB.CommandButton cmdCFCe 
            Appearance      =   0  'Flat
            Caption         =   "Gerar NFC-e"
            Height          =   615
            Left            =   3000
            MaskColor       =   &H0095FAFF&
            Picture         =   "frmEdiVendas.frx":0CCA
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdEdiNFE 
            Appearance      =   0  'Flat
            Caption         =   "Abrir NFEs"
            Height          =   615
            Left            =   120
            MaskColor       =   &H0095FAFF&
            Picture         =   "frmEdiVendas.frx":1384
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdGerarNFe 
            Appearance      =   0  'Flat
            Caption         =   "Gerar DANFE"
            Height          =   615
            Left            =   1560
            MaskColor       =   &H0095FAFF&
            Picture         =   "frmEdiVendas.frx":170E
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   120
            Width           =   1335
         End
      End
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
         Left            =   840
         TabIndex        =   39
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEmpresa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         Picture         =   "frmEdiVendas.frx":1DC8
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblSoma 
         Caption         =   "Soma:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   660
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   38
         Top             =   480
         Width           =   525
      End
      Begin VB.Label lblCrediario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   3240
         TabIndex        =   33
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblVendaAberta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   3240
         TabIndex        =   32
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.PictureBox pcBotoesFinais 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   18930
      TabIndex        =   20
      Top             =   8250
      Width           =   18960
      Begin VB.CommandButton cmdLista 
         Caption         =   "Gerar CSV"
         Height          =   705
         Left            =   13080
         Picture         =   "frmEdiVendas.frx":2173
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   705
         Left            =   14520
         Picture         =   "frmEdiVendas.frx":2580
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDestrava 
         Caption         =   "&Destravar Venda"
         Height          =   705
         Left            =   11640
         Picture         =   "frmEdiVendas.frx":2ACE
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReabre 
         Caption         =   "&Reabrir Venda"
         Height          =   705
         Left            =   10200
         Picture         =   "frmEdiVendas.frx":2E58
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReem 
         Caption         =   "&Reemitir CUPM"
         Height          =   705
         Left            =   7320
         Picture         =   "frmEdiVendas.frx":2FA2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   705
         Left            =   8760
         Picture         =   "frmEdiVendas.frx":30EC
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizarVenda 
         Caption         =   "&Efetivar Venda"
         Height          =   705
         Left            =   4440
         Picture         =   "frmEdiVendas.frx":3236
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Nova"
         Height          =   705
         Left            =   120
         Picture         =   "frmEdiVendas.frx":35C0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   705
         Left            =   3000
         Picture         =   "frmEdiVendas.frx":394A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   705
         Left            =   1560
         Picture         =   "frmEdiVendas.frx":3A94
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdDevolver 
         Caption         =   "Devolver"
         Height          =   705
         Left            =   5880
         Picture         =   "frmEdiVendas.frx":3E1E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   2640
      Top             =   4200
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FCB683&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   18930
      TabIndex        =   19
      Top             =   7515
      Width           =   18960
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         ItemData        =   "frmEdiVendas.frx":4218
         Left            =   9960
         List            =   "frmEdiVendas.frx":422B
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         MaxLength       =   40
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtbxVendedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         MaxLength       =   40
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxSituacao_ven 
         Height          =   315
         ItemData        =   "frmEdiVendas.frx":4278
         Left            =   120
         List            =   "frmEdiVendas.frx":428B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiVendas.frx":42D8
         Left            =   11400
         List            =   "frmEdiVendas.frx":42E8
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnBuscar 
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
         Height          =   645
         Left            =   13080
         Picture         =   "frmEdiVendas.frx":4313
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   1335
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   9600
         Picture         =   "frmEdiVendas.frx":440C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdBuscarFun 
         Height          =   285
         Left            =   7440
         Picture         =   "frmEdiVendas.frx":4505
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cai&xa"
         Height          =   195
         Left            =   9960
         TabIndex        =   14
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   7800
         TabIndex        =   11
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblVenFec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formas Pag:"
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
         Left            =   15480
         TabIndex        =   35
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento:"
         Height          =   195
         Left            =   15480
         TabIndex        =   34
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   6120
         TabIndex        =   8
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação da Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   11400
         TabIndex        =   16
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número/Cliente"
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   0
         Width           =   1110
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   3960
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   5040
         TabIndex        =   6
         Top             =   0
         Width           =   285
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmEdiVendas.frx":45FE
      Height          =   5295
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   9340
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
      ColumnCount     =   19
      BeginProperty Column00 
         DataField       =   "num_ven"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "fantasia_cli"
         Caption         =   "Fantasia"
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
         DataField       =   "telefonecli_ven"
         Caption         =   "Telefone"
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
         DataField       =   "data_ven"
         Caption         =   "data"
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
         DataField       =   "dataefet_ven"
         Caption         =   "Dt. Efet."
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
         DataField       =   "valor_ven"
         Caption         =   "Valor"
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
      BeginProperty Column08 
         DataField       =   "desc_perc"
         Caption         =   "Desconto  %"
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
      BeginProperty Column09 
         DataField       =   "desconto_ven"
         Caption         =   "Desconto"
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
         DataField       =   "total"
         Caption         =   "Vlr Total"
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
         DataField       =   "devolvido_Ven"
         Caption         =   "Devolvido"
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
         DataField       =   "situacao_ven"
         Caption         =   "S"
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
         DataField       =   "guerra_fun"
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
      BeginProperty Column14 
         DataField       =   "fechavenda_usu"
         Caption         =   "Concluído Por"
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
         Caption         =   "Forma Pag."
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
         DataField       =   "autorizadesconto_ven"
         Caption         =   "Autoriza Desconto."
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
         DataField       =   "cancelavenda_usu"
         Caption         =   "Cancelamento"
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
         DataField       =   "cpfcli_ven"
         Caption         =   "CPF"
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
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3374,929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2324,977
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1604,976
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   299,906
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column18 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpt 
      Caption         =   "&Opções"
      Begin VB.Menu mnImprVenda 
         Caption         =   "Imprimir Venda"
      End
      Begin VB.Menu mnGerarPDF 
         Caption         =   "Gerar PDF"
      End
      Begin VB.Menu mnEnviarEmail 
         Caption         =   "Enviar Email"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frmEdiVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim sTab As String
Dim sCod As String
Dim sNome As String
Dim WithEvents AdoVendas As ADODB.Recordset
Attribute AdoVendas.VB_VarHelpID = -1


'Private Sub adc_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

'End Sub

Private Sub cmbxBusca_Click()
    P_Busca
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub AdoVendas_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    If bActive = False Then Exit Sub


    Dim strSQL As String
    If AdoVendas.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        cmndbtnImpr.Enabled = True
        cmndbtnEdit.Enabled = True
    End If
    If AdoVendas!situacao_ven > 1 Then
        cmndbtnEdit.Enabled = False

    Else
        cmndbtnEdit.Enabled = True

    End If

    strSQL = F_BuscaUmCampo("ip_ven", "vendaaberta", "num_ven='" & AdoVendas!num_ven & "'")
    If strSQL = "" Then
        lblVendaAberta.caption = "..."
    Else
        lblVendaAberta.caption = "Esta venda está sendo utilizada pelo Terminal " & strSQL
    End If
    buscaCrediario
    buscaVenFec
End Sub

Private Sub cmbxOrd_Click()
   If Not bActive Then Exit Sub
    P_Busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_ven_Click()
   habilita_visao
   
   If Not bActive Then Exit Sub
   cmndbtnBuscar_Click
End Sub

Private Sub cmdBuscarFun_Click()
    buscaFun
End Sub

Private Sub cmdCFCe_Click()
    Dim id As String
    If AdoVendas!devolvido_ven >= AdoVendas!valor_ven Then
        MsgBox "Essa venda foi devolvida!", vbExclamation
         Exit Sub
    End If
    
    id = NFe_DAO.GravaNFCeVenda(AdoVendas!num_ven, strTipo_IcmsNFe, "Venda de Produtos", NFCe_65, False)
    If id <> "" Then
        If MsgBox("Geração concluída. Deseja transmitir NFC-e?", vbYesNo + vbQuestion) = vbYes Then
            TransmitirNFcE id, ""
        Else
             frmNFE_Detalhe.tag = id
            frmNFE_Detalhe.Show 1
    
        End If
        
    End If
End Sub

Private Sub cmdDev_Click()
    On Error GoTo trata_Erro
    frmDevolucao.txtVenda.text = AdoVendas!num_ven
    frmDevolucao.tag = AdoVendas!num_ven
    frmDevolucao.Show 1
trata_Erro:
End Sub

Private Sub cmdDevolver_Click()
    On Error Resume Next
    frmDevolucao.txtVenda.text = AdoVendas!num_ven
    frmDevolucao.Show 1
End Sub

Private Sub cmdEdiNFE_Click()
    frmEdiNFe.Show 1
End Sub

Private Sub cmdEmpresa_Click()
    buscaEmpresa
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    
    e = sped_empresa.SelecionarSped_Empresa(txtEmpresa.text)
    If e.codigo <> "" Then
        txtEmpresa.text = e.nome
        txtEmpresa.tag = e.codigo
        P_Busca
    End If
        
    
    
End Sub
Private Sub cmdGerarNFe_Click()
    Dim id As String
  '  On Error GoTo trata_Erro
    
    If AdoVendas!devolvido_ven >= AdoVendas!valor_ven Then
        MsgBox "Essa venda foi devolvida!", vbExclamation
         Exit Sub
    End If
    
    
    id = NFe_DAO.GravaNFCeVenda(AdoVendas!num_ven, strTipo_IcmsNFe, "Venda de Produtos", modelo_Nfe.NFe_55)
    If id <> "" Then
        If MsgBox("Geração concluída. Deseja transmitir NFC-e?", vbYesNo + vbQuestion) = vbYes Then
            TransmitirNFcE id, ""
        Else
                
            frmNFE_Detalhe.tag = id
            frmNFE_Detalhe.Show 1
    
            


        End If
        
    End If
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmdLista_Click()
    'P_DbgrdToHtml "Lista de Vendas", dtg, AdoVendas
    bActive = False
    ExportaGridToCSV dtg, AdoVendas
    bActive = True
End Sub

Private Sub cmdProd_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produtos.Produto
    
    p = Produtos.F_BuscaProdSaida(txtProd.text)
    If p.cod_pro <> "" Then
        txtProd.text = p.nome_pro
        txtProd.tag = p.cod_pro
        If Not bActive Then Exit Sub
        P_Busca
    End If
    
End Sub


Private Sub cmndbtnBuscar_Click()
    P_Status "Aguarde..."
    P_Busca
    P_Status ""
End Sub

Private Sub cmndbtnCancel_Click()
    Dim i As Integer
    
    If AdoVendas!situacao_ven = "1" Then
        i = CancelarVenda(AdoVendas!num_ven)
        If i <> 0 Then
            MsgBox "Não foi possível cancelar a venda.", vbExclamation
        Else
            AdoVendas.Requery
        End If
        
    Else
    
        If strExclVenda_usu <> "S" Then
            MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        End If
            If MsgBox("Deseja realmente cancelar todos os lancamentos dessa venda?", vbYesNo + vbQuestion) = vbYes Then
                If AdoVendas!situacao_ven = 2 Then
                    
                    MsgBox "A venda está fechada. Não pode ser cancelada nessa etapa.", vbExclamation
                    
                ElseIf AdoVendas!situacao_ven = 3 Then
                    MsgBox "A venda está finalizda. Não pode ser cancelada nessa etapa.", vbExclamation
                    
                Else
                    
                    
                    If AdoVendas!devolvido_ven > 0 Then
                        MsgBox "A venda não pode ser cancelada por aqui! Há devoluções", vbCritical
                        Exit Sub
                    End If
                    If Not F_senha Then
                        Exit Sub
                    End If
                        
                    If F_CancelaVendaAposCupom(AdoVendas!num_ven) Then
                        MsgBox "Venda cancelada!", vbExclamation
                        AdoVendas.Requery
                    End If
                End If
            End If
        End If
        
End Sub

Private Sub cmndbtnDestrava_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    
    If Not F_senha(True) Then
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("ip_ven", "vendaaberta", "num_ven='" & AdoVendas!num_ven & "'")
    If strSQL = "" Then
        MsgBox "A venda não está marcada como sendo utilizada em nenhum computador!", vbExclamation
        Unload Me
    End If
    strSQL = "A Venda " & AdoVendas!num_ven & " está marcada como sendo ulitazada pelo computador " & strSQL & Chr(13)
    strSQL = strSQL & "Tenha certeza que há um problema de visualização da venda, "
    strSQL = strSQL & "e que esta não se encontra aberta em nenhum computador."
    strSQL = strSQL & " Caso contrário, poderá haver problemas de valores e anomalias nesta Venda. " & Chr(13)
    strSQL = strSQL & "Tem certeza que deseja forçar a abertura desta venda?"
    If MsgBox(strSQL, vbCritical + vbYesNo) = vbYes Then
        P_DeletaVendaAberta AdoVendas!num_ven
    End If
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEdit_Click()
    Dim host As String
    If Me.tag = "U" Then
        'editaVendaFunc
    Else
        
    End If
    If AdoVendas!situacao_ven <> "1" Then
        MsgBox "Essa venda não pode ser editada!", vbExclamation
        Exit Sub
    End If
    
    host = F_VerificaVendaAberta(AdoVendas!num_ven)
    If host <> "" Then
        MsgBox "Esta venda está aberta em outro computador! (" & host & ")", vbExclamation
        Exit Sub
    End If
        
        
        
    editaVenda AdoVendas!num_ven
    P_Busca
End Sub
Sub editaVendaFunc()
'    Dim fr As New frmVendaFunc
'    P_Status "Montando venda..."
'    With AdoVendas
'        fr.tag = !num_ven
'        'fr.lblAbert.caption = !abertura_ven
'        'frmVenda.lblNome_Cli.caption = !nome_cli
'        'frmVenda.lblNome_Cli.Tag = !cod_cli
'        'frmVenda.lblHora.caption = !data_ven
'        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
'        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
'        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", tab_funcionarios, _
'            "cod_fun='" & !cod_usu & "'")
'        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
'        fr.Atacado = True
'        fr.lblValor.caption = FormatVal(!valor_ven)
'        fr.caption = "Venda " & !num_ven & "-" & !nome_Cli
'
'        'frmVenda.txtbxObs.Text = Format(!obs_ven)
'        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
'        fr.Show 0, mdiFrm
'        'fr.lblAtacado.tag = "X" 'indica que a venda foi editada
'    End With
'    P_Status ""
'    dtVendas.Refresh
End Sub
Private Sub cmndbtnFinalizar_Click()
    Dim s As String
    Dim sStatus As String
    On Error GoTo trata_Erro
    With AdoVendas
        If !situacao_ven <> "3" Then
            MsgBox "A venda não está apta a ser finalizada!", vbExclamation
            Exit Sub
        End If
        sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
        If sStatus = "F" Then
            MsgBox "Não é possível efetuar o recebimento pois este caixa está fechado!", vbExclamation
            Exit Sub
        End If
        frmLanParcelasVen.tag = !num_ven
        frmLanParcelasVen.Show 1
        
        
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnFinalizarVenda_Click()
    Dim s As String
    Dim sStatus As String
    
    On Error GoTo trata_Erro
    With AdoVendas
        
        If strEfetivarVenda_Usu = "N" Then
            MsgBox "Você não tem permissão para efetuar venda " & strNome_Usu, vbExclamation
            Exit Sub
        End If
        
        s = F_BuscaUmCampo("status_cre", "tab_crediarioven", "num_ven='" & !num_ven & "'")
        If s = "1" Then
            MsgBox "Esta venda está sendo apreciada pelo crediário!", vbExclamation
            Exit Sub
        End If
        If s = "3" Then
            MsgBox "Esta venda foi  reprovada pelo crediário!", vbExclamation
            Exit Sub
        End If
        If strEfetivarVenda_Usu <> "S" Then
            MsgBox "Este usuário não tem permissão para efetivar venda!", vbExclamation
            Exit Sub
        End If
        
        If !situacao_ven <> "2" Then
            MsgBox "A venda não está apta a ser finalizada!", vbExclamation
            Exit Sub
        End If
            
        If Not F_IsCaixa Then
            MsgBox "Este computador não é um caixa, portanto,não pode efetivar venda!", vbExclamation
            Exit Sub
        End If
            
        sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
        If sStatus = "F" Then
            MsgBox "Não é possível efetuar o recebimento pois este caixa está fechado!", vbExclamation
            Exit Sub
        End If
        s = F_Troco(!num_ven)
        
        P_Busca
                
    End With
    Exit Sub
trata_Erro:
    MostraErro
End Sub

Private Sub cmndbtnImpr_Click()

On Error GoTo trata_Erro
    If bUsaConfeccao = True Then
        imprimeVendaHTML AdoVendas!num_ven
    Else
    
        If AdoVendas!situacao_ven = "1" Then
            MsgBox "Essa venda não pode ser impressa pois encontra-se em aberto.", vbExclamation
            Exit Sub
        End If
        ImprVenda AdoVendas!num_ven
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
   NovaVenda
End Sub



Private Sub cmndbtnReabre_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If AdoVendas!situacao_ven <> "2" Then
        MsgBox "A Venda não pode ser reaberta!", vbExclamation
        Exit Sub
    End If
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para Reabrir Vendas/Excluir Vendas!", vbExclamation
        Exit Sub
    
    
    End If
    
    If MsgBox("Confirma Reabertura da Venda?", vbYesNo + vbQuestion) = vbYes Then
        If vendas.F_AlteraStatusVenda(AdoVendas!num_ven, "1") Then
            DeletaVenFec AdoVendas!num_ven
            MsgBox "Venda reaberta com sucesso!", vbInformation
            
            
            
            AdoVendas.Requery
            
            
            
        
        
        End If
        
    End If
    Exit Sub
    
trata_Erro:
P_MostraErro
End Sub

Private Sub cmndbtnReem_Click()
    'frmReemissao.Show 1
    On Error GoTo trata_Erro
    ImprimeComprovanteNaoFiscal Printer, AdoVendas!num_ven
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    AdoVendas.Close
    Unload Me
End Sub

Private Sub cmndbtnSepd_Click()
    On Error GoTo trata_Erro
    P_notaSEPD AdoVendas!num_ven, ""
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnVendasEfet_Click()
    frmEfetivaVen.Show 0
End Sub


Private Sub DBGrid1_DblClick()
    On Error GoTo trata_Erro
    If cmndbtnEdit.Enabled = True Then
        cmndbtnEdit_Click
    End If
    If AdoVendas!situacao_ven = "2" Then
        cmndbtnFinalizarVenda_Click
    ElseIf AdoVendas!situacao_ven = "3" Then
        cmndbtnFinalizar_Click
    
    End If
    Exit Sub
trata_Erro:
    
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DBGrid1_DblClick
    End If
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub dtVendas_Reposition()
    On Error Resume Next
    Dim strSQL As String
    If AdoVendas.RecordCount = 0 Then
        cmndbtnEdit.Enabled = False
        cmndbtnImpr.Enabled = False
        Exit Sub
    Else
        cmndbtnImpr.Enabled = True
        cmndbtnEdit.Enabled = True
    End If
    If AdoVendas!situacao_ven > 1 Then
        cmndbtnEdit.Enabled = False
        
    Else
        cmndbtnEdit.Enabled = True
        
    End If
    
    strSQL = F_BuscaUmCampo("ip_ven", "vendaaberta", "num_ven='" & AdoVendas!num_ven & "'")
    If strSQL = "" Then
        lblVendaAberta.caption = "..."
    Else
        lblVendaAberta.caption = "Esta venda está sendo utilizada pelo Terminal " & strSQL
    End If
    buscaCrediario
    buscaVenFec
    
End Sub
Sub buscaVenFec()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro:
    lblVenFec.caption = ""
    strSQL = "select * from " & TAB_VENFEC & " where num_ven='" & AdoVendas!num_ven & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & !esp_cai & ": " & FormatVal(!valor_ven) & " "
            .MoveNext
        Loop
        .Close
    End With
    lblVenFec.caption = strSQL
    Exit Sub
trata_Erro:
End Sub
Sub buscaCrediario()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
   ' On Error Resume Next
    On Error GoTo trata_Erro
    strSQL = "select * from tab_Crediarioven where num_ven='" & AdoVendas!num_ven & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            lblCrediario.caption = "Crediário. Status=" & !status_cre & " " & !avaliacao_cre & "  Data: " & !data_cre & " Fechamento: " & !fechamento_Cre & " Usuário: " & F_BuscaUmCampo("guerra_fun", TAB_FUNCIONARIOS, "cod_fun='" & !cod_usu & "'")
        Else
        
            lblCrediario.caption = ""
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
    
End Sub


Private Sub dtg_DblClick()
    On Error GoTo trata_Erro
    With AdoVendas
        If !situacao_ven = "1" Then
            cmndbtnEdit_Click
        ElseIf !situacao_ven = "2" Then
            cmndbtnFinalizarVenda_Click
        ElseIf !situacao_ven = "3" Then
            cmndbtnFinalizar_Click
        
        Else
            'MsgBox "Esta venda não pode ser editada!", vbExclamation
            cmndbtnEdit_Click
            
        End If
    End With
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub habilita_visao()

    If strVisualizarvenda_Usu = "S" Then
        
    Else
        If Left(cmbxSituacao_ven.text, 1) <> "1" Then
            txtbxVendedor.text = strNome_Usu
            txtbxVendedor.tag = strCod_Usu
            
            txtbxVendedor.Enabled = False
            
            
            
            
            
        Else
            txtbxVendedor.tag = ""
            txtbxVendedor.text = ""
            txtbxVendedor.Enabled = False
        End If
        
    End If
    
End Sub

Private Sub dtg_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo trata_Erro
    
    
    Dim s As String
    s = dtg.Columns(ColIndex).DataField
          AdoVendas.Sort = s
        
 
    
    
trata_Erro:
    
End Sub

Private Sub dtg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub Form_Activate()
    
    
    If bActive Then Exit Sub
    
    buscaTabela
    
    
    If Me.tag <> "" Then
        P_Busca
        
    End If
    
    If strVisualizarvenda_Usu <> "S" Then
        cmbxSituacao_ven.Enabled = False
        txtbxVendedor.text = strNome_Usu
        txtbxVendedor.tag = strCod_Usu
        txtbxVendedor.Enabled = False
        cmdBuscarFun.Enabled = False
    End If
    
    
    bActive = True
    MudaCorEmpresa pcBotoesFinais
End Sub
Sub buscaTabela()
    If Me.tag = "U" Then
        sTab = TAB_FUNCIONARIOS
        sCod = "cod_fun"
        sNome = "nome_fun"
        Me.caption = "Edição de Vendas de Funcionários"
    Else
        sTab = TAB_CLIENTES
        sCod = "cod_cli"
        sNome = "nome_cli"
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        If cmbxSituacao_ven.ListCount <= 2 Then
            cmbxSituacao_ven.AddItem "3- Canceladas"
            
            cmbxSituacao_ven.AddItem "4- Efetivadas"
        End If
    End If
    If KeyCode = vbKeyF5 Then
        P_Busca
    End If
    If KeyCode = vbKeyF9 Then
        'If pcVenda.BackColor <> vbYellow Then
        '    pcVenda.BackColor = vbYellow
        'Else
        '    pcVenda.BackColor = vbWhite
        'End If
    End If
End Sub

Private Sub Form_Load()
    bActive = False
    resizeForm Me
    txtbxDe.text = Format(DateAdd("M", -1, Date), "dd/mm/yyyy")
    txtbxAte.text = Date
'    adc.ConnectionString = Conn.ConnectionString
    
    
    Set AdoVendas = New ADODB.Recordset
    
    
    cmbxSituacao_ven.ListIndex = 0
   ' If strVisualizarvenda_Usu <> "S" Then
   '     cmbxSituacao_ven.Enabled = False
   ' End If
    'cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 1
    
    P_PreencheCaixas cmbxCaixa, True
    
    P_LocalizaOperacao strCod_Cai, cmbxCaixa
    
    'pcVenda.BackColor = vbYellow
    
End Sub
Sub P_Busca()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    On Error GoTo trata_Erro
   'If Not bActive Then Exit Sub
    
    buscaTabela
    
    If cmbxOrd.ListIndex = 0 Then
        sOrd = "nomecli_ven"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrd = "tv.num_ven"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrd = "data_ven"
    Else
        sOrd = "dataefet_ven"
    End If
    
    
    If txtbxDe.text = "" And txtbxAte.text <> "" And Me.tag = "" Then
        MsgBox "Informe a data!", vbExclamation
        txtbxDe.text = ""
        Exit Sub
    End If
    
    
    
    strSQL = "select distinct tv.cod_emp,fantasia_cli,logradouro_cli,numero_cli,complemento_cli, telefonecli_ven,cpfcli_ven,tfo.nome_for,tc." & sNome & " as nome_cli,autorizadesconto_ven, "
    strSQL = strSQL & " tv.num_ven,tv.devolvido_ven,tv.nomecli_ven,tv.valor_ven,fechavenda_usu,cancelavenda_usu,"
    strSQL = strSQL & "tv.desconto_ven,(tv.desconto_ven/(tv.valor_ven+tv.desconto_ven)*100) as desc_perc, tv.situacao_ven,tv.data_ven,tv.dataefet_ven,"
    strSQL = strSQL & "tv.cod_cli,tv.nomecli_ven,dataefet_ven,data_ven,"
    strSQL = strSQL & " tv.abrevenda_usu,tv.valor_ven+tv.desconto_ven as total,"
    strSQL = strSQL & " tf.guerra_fun from "
    
    strSQL = strSQL & TAB_VENDAS & " as tv inner join " & sTab & " as tc on tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " inner join tab_funcionarios as tf on tv.abrevenda_usu=tf.cod_fun "
    strSQL = strSQL & " left join tab_formaspag as tfo  on tv.cod_For=tfo.cod_For  "
    If cmbxCaixa.text <> "" Then
        strSQL = strSQL & " inner join tab_caixas as tx on tv.cod_cai=tx.cod_cai "
    End If
    If txtProd.tag <> "" Then
        strSQL = strSQL & "left join " & TAB_ITVENDAS & " as ti on tv.num_ven=ti.num_ven "
                
    End If
    
    strSQL = strSQL & " where 1>0 "
    
    If txtEmpresa.tag <> "" Then
        strSQL = strSQL & " and tv.cod_emp='" & txtEmpresa.tag & "' "
    End If
    If (Left(cmbxSituacao_ven.text, 1) = "1" Or Left(cmbxSituacao_ven.text, 1) = "2") And Me.tag <> "" Then
        strSQL = strSQL & " and tv.situacao_ven in ('1','2')"
    Else
        strSQL = strSQL & " and tv.situacao_ven='" & Left(cmbxSituacao_ven.text, 1) & "' "
    
    
    End If
    
  
    If txtbxExpr.text <> "" Then
        strSQL = strSQL & " and (nomecli_ven like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & " or nome_cli like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & " or fantasia_cli like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & " or tv.num_ven like '" & txtbxExpr.text & "%' ) "
    
    End If
    
    If txtEmpresa.text <> "" Then
        strSQL = strSQL & " and tv.cod_emp='" & txtEmpresa.tag & "'"
    End If
    
    
    'If Left(cmbxSituacao_ven.Text, 1) = "4" And pcVenda.BackColor <> vbYellow Then
    '    strSQL = strSQL & " and tx.num_ven=tv.num_ven "
        
     '   If pcVenda.BackColor <> vbYellow Then
'            strSQL = strSQL & " and tx.visivel_cai='S' "
     '   End If
        
    'End If
    If cmbxCaixa.text <> "" Then
        strSQL = strSQL & " and tv.cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
    End If
    
    
    If txtProd.tag <> "" Then
        
        
        strSQL = strSQL & " and ti.cod_pro='" & txtProd.tag & "'"
    End If
    
    If txtEmpresa.tag <> "" Then
        strSQL = strSQL & " and tv.cod_emp='" & txtEmpresa.tag & "' "
    End If
    
    
    
    strSQL = strSQL & " and tc." & sCod & "=tv.cod_cli "
    If txtbxVendedor.tag <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & txtbxVendedor.tag & "' "
    End If
    
    'If UCase(App.Title) = "PIKIBOX_CX" Then
    '    strSQL = strSQL & " and tv.cod_cai='" & strCod_Cai & "' "
    'End If
    
    If txtbxDe.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" Then
            strSQL = strSQL & " and dataefet_ven >=" & F_GravaData(txtbxDe.text)
        Else
            strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(txtbxDe.text)
        End If
    End If
    If txtbxAte.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxAte.text)
        Else
            strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(txtbxAte.text & " 23:59:59")
        End If
    End If
    
    strSQL = strSQL & " order by " & sOrd & " desc "
    
    
    
    abreObjetoAdo AdoVendas, strSQL, dtg
    
    dtg.caption = AdoVendas.RecordCount & " Registros listados"
    lblTotal.caption = "0,00"
    
    
    P_Status ""
    
        Exit Sub

trata_Erro:
MostraErro


P_Status
End Sub
Sub soma()
    On Error GoTo trata_Erro
    P_Status "somando..."
    
    Dim cTotal As Currency
    bActive = False
    dtg.Visible = False
    With AdoVendas
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !valor_ven
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    lblTotal.caption = FormatVal(cTotal)
    dtg.Visible = True
    P_Status ""
    bActive = True
trata_Erro:
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    
    PosicionaGrade dtg, Me, pcVenda
    posicionaADireita cmndbtnSair, pcBot
    
    posicionaADireita pcNfe, pcVenda
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    AdoVendas.Close
End Sub


Sub gerarPDF()
    On Error GoTo trata_Erro
    bioPDF.ApontarPrinterToPDF "Venda" & AdoVendas!num_ven
    imprimeVenda Printer, AdoVendas!num_ven
    Printer.EndDoc
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub lblSoma_Click()
    soma
End Sub

Private Sub mnEnviarEmail_Click()
    enviarEmail
End Sub
Sub enviarEmail()
    EnviarEmailVenda AdoVendas!num_ven
End Sub
Private Sub mnGerarPDF_Click()
    gerarPDF
End Sub

Private Sub mnImprVenda_Click()
    cmndbtnImpr_Click
End Sub

Private Sub mnSair_Click()
    cmndbtnSair_Click
End Sub

Private Sub txtbxAte_GotFocus()
    ganhaFoco txtbxAte
End Sub

Private Sub txtbxAte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAte_LostFocus()
    P_AjustaData txtbxAte
End Sub

Private Sub txtbxDe_GotFocus()
    ganhaFoco txtbxDe
End Sub

Private Sub txtbxDe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDe_LostFocus()
    P_AjustaData txtbxDe
End Sub

Private Sub txtbxExpr_DblClick()
    P_Busca
End Sub

Private Sub txtbxExpr_GotFocus()
    ganhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        P_Busca
        
    End If
    If KeyCode = vbKeyDown Then
        
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    AjustaTexto txtbxExpr
End Sub

Private Sub txtbxVendedor_Change()
    txtbxVendedor.tag = ""
End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaFun
    End If
End Sub
Sub buscaFun()
    Dim f As funcionario.funcionario
    f = F_SelectVendedor(txtbxVendedor.text)
    txtbxVendedor.text = f.guerra_fun
    txtbxVendedor.tag = f.cod_fun
    If Not bActive Then Exit Sub
    P_Busca
End Sub

Private Sub txtEmpresa_GotFocus()
    P_SelecTexto txtEmpresa
End Sub

Private Sub txtEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtEmpresa_LostFocus()
    P_AjustaTexto txtEmpresa
End Sub

Private Sub txtProd_Change()
    txtProd.tag = ""
End Sub


