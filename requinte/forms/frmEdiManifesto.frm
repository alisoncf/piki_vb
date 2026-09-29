VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiManifesto 
   Caption         =   "Manifestos"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14190
   Icon            =   "frmEdiManifesto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14190
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0073E1F4&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   14160
      TabIndex        =   18
      Top             =   7560
      Width           =   14190
      Begin VB.TextBox txtProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         MaxLength       =   40
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtbxVendedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         MaxLength       =   40
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxSituacao_ven 
         Height          =   315
         ItemData        =   "frmEdiManifesto.frx":1084A
         Left            =   120
         List            =   "frmEdiManifesto.frx":1085D
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiManifesto.frx":108AA
         Left            =   10080
         List            =   "frmEdiManifesto.frx":108BA
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "Pesquisar"
         Height          =   525
         Left            =   11640
         Picture         =   "frmEdiManifesto.frx":108E5
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         MaxLength       =   40
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdProd 
         Height          =   285
         Left            =   9600
         Picture         =   "frmEdiManifesto.frx":109DE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdBuscarFun 
         Height          =   285
         Left            =   7440
         Picture         =   "frmEdiManifesto.frx":10AD7
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   7800
         TabIndex        =   37
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
         Left            =   13200
         TabIndex        =   36
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento:"
         Height          =   195
         Left            =   13200
         TabIndex        =   35
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   6120
         TabIndex        =   34
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação da Venda:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   10080
         TabIndex        =   32
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número/Cliente"
         Height          =   195
         Left            =   1800
         TabIndex        =   31
         Top             =   0
         Width           =   1110
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   3960
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   5040
         TabIndex        =   29
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   14160
      TabIndex        =   6
      Top             =   8175
      Width           =   14190
      Begin VB.CommandButton cmdLista 
         Caption         =   "Lista"
         Height          =   720
         Left            =   13080
         Picture         =   "frmEdiManifesto.frx":10BD0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   705
         Left            =   14280
         Picture         =   "frmEdiManifesto.frx":10FDD
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDestrava 
         Caption         =   "&Destravar Venda"
         Height          =   705
         Left            =   11640
         Picture         =   "frmEdiManifesto.frx":1152B
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReabre 
         Caption         =   "&Reabrir Venda"
         Height          =   705
         Left            =   10200
         Picture         =   "frmEdiManifesto.frx":118B5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReem 
         Caption         =   "&Reemitir TEF"
         Height          =   705
         Left            =   7320
         Picture         =   "frmEdiManifesto.frx":119FF
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   705
         Left            =   8760
         Picture         =   "frmEdiManifesto.frx":11B49
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizarVenda 
         Caption         =   "&Efetivar Venda"
         Height          =   705
         Left            =   4440
         Picture         =   "frmEdiManifesto.frx":11C93
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Nova"
         Height          =   705
         Left            =   120
         Picture         =   "frmEdiManifesto.frx":1201D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   705
         Left            =   3000
         Picture         =   "frmEdiManifesto.frx":123A7
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   705
         Left            =   1560
         Picture         =   "frmEdiManifesto.frx":124F1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizar 
         Caption         =   "&Finalizar"
         Height          =   705
         Left            =   5880
         Picture         =   "frmEdiManifesto.frx":1287B
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox pcVenda 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   14160
      TabIndex        =   0
      Top             =   7065
      Width           =   14190
      Begin VB.CommandButton cmdGerarNFe 
         Caption         =   "Gerar NFe (DANFE)"
         Height          =   375
         Left            =   12240
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdNFes 
         Caption         =   "Abrir NFe's"
         Height          =   375
         Left            =   9600
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdGerarNFcE 
         Caption         =   "Gerar NFCe"
         Height          =   375
         Left            =   10920
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblCrediario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda Disponível"
         Height          =   195
         Left            =   5400
         TabIndex        =   5
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblVendaAberta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda Disponível"
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1275
      End
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmEdiManifesto.frx":12C05
      Height          =   5295
      Left            =   120
      TabIndex        =   38
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
      ColumnCount     =   12
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
            ColumnWidth     =   3374,929
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   299,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiManifesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
