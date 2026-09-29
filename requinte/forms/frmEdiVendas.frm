VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiVendas 
   Caption         =   "Edição de Vendas"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   16545
   Icon            =   "frmEdiVendas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   16545
   Begin VB.PictureBox pcVenda 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   16515
      TabIndex        =   29
      Top             =   7140
      Width           =   16545
      Begin VB.CommandButton cmdGerarNFe 
         Caption         =   "Gerar NFe (DANFE)"
         Height          =   375
         Left            =   14160
         TabIndex        =   38
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdNFes 
         Caption         =   "Abrir NFe's"
         Height          =   375
         Left            =   11520
         TabIndex        =   37
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdGerarNFcE 
         Caption         =   "Gerar NFCe"
         Height          =   375
         Left            =   12840
         TabIndex        =   36
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblCrediario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crediário"
         Height          =   195
         Left            =   3720
         TabIndex        =   31
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblVendaAberta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   3720
         TabIndex        =   30
         Top             =   240
         Width           =   135
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
      ScaleWidth      =   16515
      TabIndex        =   18
      Top             =   8250
      Width           =   16545
      Begin VB.CommandButton cmdLista 
         Caption         =   "Lista"
         Height          =   720
         Left            =   13080
         Picture         =   "frmEdiVendas.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   705
         Left            =   14520
         Picture         =   "frmEdiVendas.frx":10D7
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDestrava 
         Caption         =   "&Destravar Venda"
         Height          =   705
         Left            =   11640
         Picture         =   "frmEdiVendas.frx":1625
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReabre 
         Caption         =   "&Reabrir Venda"
         Height          =   705
         Left            =   10200
         Picture         =   "frmEdiVendas.frx":19AF
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnReem 
         Caption         =   "&Reemitir TEF"
         Height          =   705
         Left            =   7320
         Picture         =   "frmEdiVendas.frx":1AF9
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCancel 
         Caption         =   "&Cancelar"
         Height          =   705
         Left            =   8760
         Picture         =   "frmEdiVendas.frx":1C43
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizarVenda 
         Caption         =   "&Efetivar Venda"
         Height          =   705
         Left            =   4440
         Picture         =   "frmEdiVendas.frx":1D8D
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Nova"
         Height          =   705
         Left            =   120
         Picture         =   "frmEdiVendas.frx":2117
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   705
         Left            =   3000
         Picture         =   "frmEdiVendas.frx":24A1
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   705
         Left            =   1560
         Picture         =   "frmEdiVendas.frx":25EB
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnFinalizar 
         Caption         =   "&Finalizar"
         Height          =   705
         Left            =   5880
         Picture         =   "frmEdiVendas.frx":2975
         Style           =   1  'Graphical
         TabIndex        =   19
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
      BackColor       =   &H00EEDFD5&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   16515
      TabIndex        =   17
      Top             =   7635
      Width           =   16545
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
         ItemData        =   "frmEdiVendas.frx":2CFF
         Left            =   120
         List            =   "frmEdiVendas.frx":2D12
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiVendas.frx":2D5F
         Left            =   10080
         List            =   "frmEdiVendas.frx":2D6F
         Style           =   2  'Dropdown List
         TabIndex        =   15
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
         Height          =   525
         Left            =   11640
         Picture         =   "frmEdiVendas.frx":2D9A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
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
         Picture         =   "frmEdiVendas.frx":2E93
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdBuscarFun 
         Height          =   285
         Left            =   7440
         Picture         =   "frmEdiVendas.frx":2F8C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   255
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
         Left            =   13200
         TabIndex        =   33
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento:"
         Height          =   195
         Left            =   13200
         TabIndex        =   32
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   6120
         TabIndex        =   8
         Top             =   0
         Width           =   735
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
         Left            =   10080
         TabIndex        =   14
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
      Bindings        =   "frmEdiVendas.frx":3085
      Height          =   6255
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   11033
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
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "num_ven"
         Caption         =   "Número"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """R$""#.##0,00"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "total_produtos"
         Caption         =   "Total Produtos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "desconto_ven"
         Caption         =   "Desconto"
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
      BeginProperty Column07 
         DataField       =   "valorentrega_ven"
         Caption         =   "Tx. Entrega"
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
      BeginProperty Column08 
         DataField       =   "total"
         Caption         =   "Valor Final"
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
      BeginProperty Column10 
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
      BeginProperty Column11 
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
      BeginProperty Column12 
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
      BeginProperty Column13 
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
      BeginProperty Column14 
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
      BeginProperty Column15 
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
      BeginProperty Column16 
         DataField       =   "email_ven"
         Caption         =   "Email_ven"
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
         DataField       =   "aniversario_ven"
         Caption         =   "Aniv."
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
         DataField       =   "enderecoentrega_ven"
         Caption         =   "Entrega"
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
         DataField       =   "operacao"
         Caption         =   "Operacao"
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
            ColumnWidth     =   1604,976
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1305,071
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   299,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2684,977
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   3869,858
         EndProperty
         BeginProperty Column19 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnMenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnImprimir 
         Caption         =   "Imprimir Venda"
      End
      Begin VB.Menu mnImprTef 
         Caption         =   "Imprimir 2a. Via TEF"
      End
   End
   Begin VB.Menu mnOpt 
      Caption         =   "Opções"
      Begin VB.Menu mnDevolução 
         Caption         =   "&Devoluções"
      End
      Begin VB.Menu mnPrint 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnTroca 
         Caption         =   "&Troca"
      End
      Begin VB.Menu mnParcelas 
         Caption         =   "&Parcelas"
      End
      Begin VB.Menu mnLancarProdutos 
         Caption         =   "&Lançar mais produtos"
      End
      Begin VB.Menu mnMudar 
         Caption         =   "Mudar para Remessa/Devolução"
      End
      Begin VB.Menu mnAbrirContasReceber 
         Caption         =   "Abrir Contas a Receber"
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
Public Tipo_Operacao As tipo_operacao_venda



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

    Dim strSQL As String
    
    If bActive = False Then
        Exit Sub
    End If
    
    
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

    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & AdoVendas!num_ven & "'")
    If strSQL = "" Then
        lblVendaAberta.caption = "..."
    Else
        lblVendaAberta.caption = "Esta venda está sendo utilizada pelo Terminal " & strSQL
    End If
    buscaCrediario
    buscaVenFec
End Sub

Private Sub cmbxOrd_Click()
    P_Busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_ven_Click()
    If bActive = False Then
        Exit Sub
    End If
   
   
   habilita_visao
   
   
   cmndbtnBuscar_Click
End Sub

Private Sub cmdBuscarFun_Click()
    buscaFun
End Sub

Private Sub cmdGerarNFcE_Click()
    Dim id As String
    On Error GoTo trata_Erro
    Dim ncm As String
    id = F_BuscaUmCampo("id", "nfe", "num_Ven='" & AdoVendas!num_ven & "'")
    If id <> "" Then
        MsgBox "Já existe uma  NFe gerada para esta venda. ID: " & id, vbExclamation
        Exit Sub
    End If
    
    
    P_Status "Gerando NF-E..."
    
    ncm = VerificaNCMVenda(AdoVendas!num_ven)
    If ncm = "" Then
        id = NFe_DAO.GravaNFCeVenda(AdoVendas!num_ven, strTipo_IcmsNFe, "Venda de Produtos", NFCe_65)
    Else
        MsgBox "Não é possível gerar a nota, há produtos sem NCM ou com ncm inválido." & Chr(13) & ncm, vbExclamation
    End If
    P_Status ""
    If id <> "" Then
        If MsgBox("Deseja transmitir agora?", vbYesNo + vbQuestion) = vbYes Then
            TransmitirNFcE id, ""
        End If
    End If
    Exit Sub
trata_Erro:
    
    
End Sub

Private Sub cmdGerarNFe_Click()
    Dim id As String
    On Error GoTo trata_Erro
    
    id = NFe_DAO.GravaNFCeVenda(AdoVendas!num_ven, strTipo_IcmsNFe, "Venda de Produtos", modelo_Nfe.NFe_55)
    If id <> "" Then
        If MsgBox("Geração concluída. Deseja transmitir NFC-e?", vbYesNo + vbQuestion) = vbYes Then
            TransmitirNFcE id, ""
        Else
                
            frmNFE_Detalhe.tag = id
            frmNFE_Detalhe.Show 1
    
            


        End If
        
    End If
    P_Status ""
Exit Sub
trata_Erro:
    P_MostraErro

End Sub

Private Sub cmdLista_Click()
    P_DbgrdToHtml "Lista de Vendas", dtg, AdoVendas
End Sub

Private Sub cmdNFes_Click()
    frmEdiNFe.Show 1
End Sub

Private Sub cmdProd_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As produtos.produto
    
    p = produtos.F_BuscaProdSaida(txtProd.text)
    If p.cod_pro <> "" Then
        txtProd.text = p.nome_pro
        txtProd.tag = p.cod_pro
        P_Busca
    End If
    
End Sub
Private Sub cmndbtnBuscar_Click()
     
    
    P_Status "Aguarde..."
    
    P_Busca
    
    P_Status ""
End Sub
Sub P_Busca()
    bActive = False
    If txtbxAte.text = "" And txtbxDe.text = "" Then
        MsgBox "Informe datas", vbExclamation
        txtbxDe.text = Date
        txtbxDe.SetFocus
    End If
    
    If cmbxSituacao_ven.ListIndex < 3 Then
    
        BuscaVendaAbertaFechada
    Else
        BuscaVendaCompleta
    End If
    
End Sub
Private Sub cmndbtnCancel_Click()
    If strExclVenda_usu = "S" Then
        If MsgBox("Deseja realmente cancelar todos os lancamentos dessa venda?", vbYesNo + vbQuestion) = vbYes Then
            If AdoVendas!situacao_ven < 3 Then
                MsgBox "A venda não pode ser cancelada por aqui!", vbCritical
                Exit Sub
            End If
            If AdoVendas!situacao_ven = "4" And AdoVendas!cod_cai <> strCod_Cai Then
                MsgBox "Esta venda só pode ser cancelada no caixa " & strCod_Cai, vbExclamation
                Exit Sub
            End If
            If F_CancelaVendaAposCupom(AdoVendas!num_ven) Then
                AdoVendas.Requery
            End If
        End If
    Else
        MsgBox "Você não tem permissão para isso!", vbCritical
    End If
End Sub

Private Sub cmndbtnDestrava_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If strNivel_Usu < "2" And strCod_Car = "004" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    
    If Not F_senha(True) Then
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & AdoVendas!num_ven & "'")
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
    If Me.tag = "U" Then
        editaVendaFunc
    Else
        editaVenda
    End If
End Sub
Sub editaVenda()
    Dim fr As New frmVenda
    P_Status "Montando venda..."
    
    
    With AdoVendas
        If !situacao_ven <> 1 Then
    '            MsgBox "Esta venda não pode ser editada. Ela não está na situação 'em aberto'", vbExclamation
            
            P_Status
            Vendas.AlteraCabvenda !num_ven
            
            Exit Sub
        End If
        fr.tag = !num_ven
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        fr.Atacado = True
        fr.lblValor.caption = FormatVal(!valor_ven)
        fr.caption = "Venda " & !num_ven & "-" & !nome_cli
        
        'frmVenda.txtbxObs.Text = Format(!obs_ven)
        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
        fr.Show 0, mdiFrm
        'fr.lblAtacado.tag = "X" 'indica que a venda foi editada
    End With
    P_Status ""
    'dtVendas.Refresh
End Sub
Sub editaVendaFunc()
    Dim fr As New frmVendaFunc
    P_Status "Montando venda..."
    With AdoVendas
        fr.tag = !num_ven
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        fr.Atacado = True
        fr.lblValor.caption = FormatVal(!valor_ven)
        fr.caption = "Venda " & !num_ven & "-" & !nome_cli
        
        'frmVenda.txtbxObs.Text = Format(!obs_ven)
        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
        fr.Show 0, mdiFrm
        'fr.lblAtacado.tag = "X" 'indica que a venda foi editada
    End With
    P_Status ""
    dtVendas.Refresh
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
        
        s = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & !num_ven & "'")
        If s <> "" Then
            MsgBox "Essa venda está sendo manipulada no computador de IP: " & s, vbExclamation
            Exit Sub
        End If
        
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
        
        
        
        
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnImpr_Click()
'On Error GoTo trata_erro
    ImprVenda AdoVendas!num_ven
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnNew_Click()
    On Error Resume Next
   Dim strSQL As String
    
Dim s As String
    
    
    
    
    
    If Me.tag = "U" Then
        s = F_GeraNovaVendaFuncionario()
    
    
    
    
        If s <> "" Then
            frmVendaFunc.tag = s
            frmVendaFunc.Show 0
        End If
    
    Else
        If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    strSQL = F_GeraNovaVenda("00000", Tipo_Operacao)
    
    
    
    frmVenda.caption = "Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    frmVenda.Show 0, mdiFrm
    
    
    End If
    
End Sub



Private Sub cmndbtnReabre_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If AdoVendas!situacao_ven <> "2" Then
        MsgBox "A Venda não pode ser reaberta!", vbExclamation
        Exit Sub
    End If
    If strDesbloquearVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para Reabrir Vendas!", vbExclamation
        Exit Sub
    
    
    End If
    
    If MsgBox("Confirma Reabertura da Venda?", vbYesNo + vbQuestion) = vbYes Then
        If Not F_senha Then
        
            Exit Sub
        End If
        
        If Vendas.F_AlteraStatusVenda(AdoVendas!num_ven, "1") Then
            MsgBox "Venda reaberta com sucesso!", vbInformation
            AdoVendas.Requery
        End If
        
    End If
    Exit Sub
    
trata_Erro:
P_MostraErro
End Sub

Private Sub cmndbtnReem_Click()
    On Error GoTo trata_Erro
    ImprimeSegundaViaCartao AdoVendas!num_ven
    Exit Sub
trata_Erro:
P_MostraErro
    
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    AdoVendas.Close
    Unload Me
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
    
    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & AdoVendas!num_ven & "'")
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
    strSQL = "select * from tab_venfec where num_ven='" & AdoVendas!num_ven & "' "
    abreObjetoAdo rPesq, strSQL
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
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            lblCrediario.caption = "Crediário. Status=" & !status_cre & " " & !avaliacao_cre & "  Data: " & !data_cre & " Fechamento: " & !fechamento_Cre & " Usuário: " & F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & !cod_usu & "'")
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
            editaVenda
            
        End If
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub habilita_visao()

    If strVisualizarvenda_Usu = "S" Or strEfetivarVenda_Usu = "S" Then
        txtbxVendedor.Enabled = True
    Else
        If Left(cmbxSituacao_ven.text, 1) <> "1" Then
            txtbxVendedor.text = strNome_Usu
            txtbxVendedor.tag = strCod_Usu
            
            txtbxVendedor.Enabled = False
            
        Else
            txtbxVendedor.tag = ""
            txtbxVendedor.text = ""
            txtbxVendedor.Enabled = True
        End If
        
    End If
    
End Sub

Private Sub dtg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpt
    End If
    
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    carrega
    bActive = True
End Sub
Sub carrega()
    buscaTabela
    
    
    If Tipo_Operacao <> tipo_operacao_venda.venda Then
         pcBot.BackColor = &HBCDF15
        Me.caption = "Lista de Remessas e Devoluções"
        mnMudar.caption = "Mudar para Vendas"
    Else
       Me.caption = "Edição de Vendas"
        pcBot.BackColor = &HEEDFD5
        mnMudar.caption = "Mudar para Remessa/Devolução"
    End If
    
    
    
    P_Busca
End Sub
Sub muda_operacao()
    If Tipo_Operacao = venda Then
        If strAbrenota_usu <> "S" Then
            MsgBox "Você não tem permissão para essa operação!", vbExclamation
            Exit Sub
        End If
        
        Tipo_Operacao = Remessa
    Else
        Tipo_Operacao = venda
    End If
    
    carrega
    
End Sub
Sub buscaTabela()
    If Me.tag = "U" Then
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "nome_fun"
        Me.caption = "Edição de Vendas de Funcionários"
    Else
        sTab = "tab_clientes"
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
        If pcVenda.BackColor <> vbYellow Then
            pcVenda.BackColor = vbYellow
        Else
            pcVenda.BackColor = vbWhite
        End If
    End If
End Sub

Private Sub Form_Load()
    resizeForm Me
    Tipo_Operacao = venda
    txtbxDe.text = Date
    txtbxAte.text = Date
'    adc.ConnectionString = Conn.ConnectionString
    
    
    Set AdoVendas = New ADODB.Recordset
    
    
    cmbxSituacao_ven.ListIndex = 0
   ' If strVisualizarvenda_Usu <> "S" Then
   '     cmbxSituacao_ven.Enabled = False
   ' End If
    'cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 1
    
    
    bActive = False
End Sub
Sub BuscaVendaCompleta()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    On Error GoTo trata_Erro
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
    bActive = False
    
    strSQL = "select distinct (valorentrega_ven),logradouro_cli,numero_cli,complemento_cli, telefonecli_ven,cpfcli_ven,"
    strSQL = strSQL & " tfo.nome_for,tc." & sNome & " as nome_cli,autorizadesconto_ven, "
    strSQL = strSQL & " tv.num_ven,tv.nomecli_ven,tv.valor_ven,fechavenda_usu,cancelavenda_usu,"
    strSQL = strSQL & "tv.desconto_ven,tv.situacao_ven,tv.data_ven,tv.dataefet_ven,"
    strSQL = strSQL & "tv.cod_cli,tv.nomecli_ven,dataefet_ven,data_ven,if(tipo_operacao=1,'Venda',if(tipo_operacao=2,'Troca','Remessa')) as operacao,"
    strSQL = strSQL & " tv.abrevenda_usu,(tv.valor_ven+desconto_ven) as total_produtos,tv.valor_ven+tv.valorentrega_ven as total,email_ven,aniversario_ven, "
    strSQL = strSQL & " tf.guerra_fun, tv.cod_cai "
    strSQL = strSQL & " from " & sTab & " as tc inner join  " & TAB_VENDAS & " as tv on  tc." & sCod & "=tv.cod_cli "
    strSQL = strSQL & " inner join tab_funcionarios as tf "
    strSQL = strSQL & " on tf.cod_fun=tv.abrevenda_usu "

    strSQL = strSQL & " inner join tab_formaspag as tfo  "
    
    strSQL = strSQL & " on tfo.cod_for=tv.cod_for "
    
        
    strSQL = strSQL & "inner join tab_caixa as tx on tv.cod_cai=tx.cod_cai "
    
    
    
        strSQL = strSQL & " inner join tab_itvendas as ti on ti.num_ven=tv.num_ven "
                
        
    
    strSQL = strSQL & " where tf.cod_fun=tv.abrevenda_usu "
    
    If txtProd.tag <> "" Then
        
        
        strSQL = strSQL & " and ti.cod_pro='" & txtProd.tag & "'"
    End If
    If txtbxExpr.text <> "" Then
        strSQL = strSQL & " and (nomecli_ven like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & " or tv.num_ven like '" & txtbxExpr.text & "%' ) "
    
    End If
    If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor <> vbYellow Then
        strSQL = strSQL & " and tx.num_ven=tv.num_ven "
        
        If pcVenda.BackColor <> vbYellow Then
       '     strSQL = strSQL & " and tx.visivel_cai='S' "
        End If
        
    End If
    

    strSQL = strSQL & " and tv.situacao_ven='" & Left(cmbxSituacao_ven.text, 1) & "' "
    
    If txtbxVendedor.tag <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & txtbxVendedor.tag & "' "
    End If
    
    If txtbxDe.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor = vbYellow Then
            strSQL = strSQL & " and dataefet_ven >=" & F_GravaData(txtbxDe.text)
        Else
            strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(txtbxDe.text)
        End If
    End If
    If txtbxAte.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor = vbYellow Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxAte.text)
        Else
            strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(txtbxAte.text & " 23:59:59")
        End If
    End If
    
   If Tipo_Operacao <> venda Then
        strSQL = strSQL & " and tipo_operacao>1 "
    Else
        strSQL = strSQL & " and tipo_operacao=1"
    End If
    strSQL = strSQL & " order by " & sOrd & " desc "
    
    
    
    abreObjetoAdo AdoVendas, strSQL, dtg
    
    
    
    P_Status ""
    bActive = True
        Exit Sub
trata_Erro:
P_MostraErro
bActive = True

P_Status
End Sub
Sub BuscaVendaAbertaFechada()
    Dim strSQL As String
    Dim sBusca As String
    Dim sOrd As String
    On Error GoTo trata_Erro
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
    
    
    strSQL = "select distinct logradouro_cli,(valorentrega_ven),numero_cli,complemento_cli, telefonecli_ven,cpfcli_ven,tfo.nome_for,tc." & sNome & " as nome_cli,"
    strSQL = strSQL & " autorizadesconto_ven, "
    strSQL = strSQL & " tv.num_ven,tv.nomecli_ven,tv.valor_ven,fechavenda_usu,cancelavenda_usu,"
    strSQL = strSQL & "tv.desconto_ven,tv.situacao_ven,tv.data_ven,tv.dataefet_ven,"
    strSQL = strSQL & "tv.cod_cli,tv.nomecli_ven,dataefet_ven,data_ven,if(tipo_operacao=1,'Venda',if(tipo_operacao=2,'Troca','Remessa')) as operacao,"
    strSQL = strSQL & " tv.abrevenda_usu,(tv.valor_ven+desconto_ven) as total_produtos,tv.valor_ven+ valorentrega_ven as total,email_ven,aniversario_ven, "
    strSQL = strSQL & " tf.guerra_fun "
    strSQL = strSQL & " from " & sTab & " as tc inner join  " & TAB_VENDAS & " as tv on  tc." & sCod & "=tv.cod_cli "
    strSQL = strSQL & " inner join tab_funcionarios as tf "
    strSQL = strSQL & " on tf.cod_fun=tv.abrevenda_usu "

    strSQL = strSQL & " inner join tab_formaspag as tfo  "
    
    strSQL = strSQL & " on tfo.cod_for=tv.cod_for "
    
    
    
    
                
        
    
    strSQL = strSQL & " where tf.cod_fun=tv.abrevenda_usu "
    
    If txtProd.tag <> "" Then
        
        
        strSQL = strSQL & " and ti.cod_pro='" & txtProd.tag & "'"
    End If
    If txtbxExpr.text <> "" Then
        strSQL = strSQL & " and (nomecli_ven like '" & txtbxExpr.text & "%' "
        strSQL = strSQL & " or tv.num_ven like '" & txtbxExpr.text & "%' ) "
    
    End If
    If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor <> vbYellow Then
        strSQL = strSQL & " and tx.num_ven=tv.num_ven "
        
        If pcVenda.BackColor <> vbYellow Then
            'strSQL = strSQL & " and tx.visivel_cai='S' "
        End If
        
    End If
    

    strSQL = strSQL & " and tv.situacao_ven='" & Left(cmbxSituacao_ven.text, 1) & "' "
    
    If txtbxVendedor.tag <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & txtbxVendedor.tag & "' "
    End If
    
    If txtbxDe.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor = vbYellow Then
            strSQL = strSQL & " and dataefet_ven >=" & F_GravaData(txtbxDe.text)
        Else
            strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(txtbxDe.text)
        End If
    End If
    If txtbxAte.text <> "" Then
        If Left(cmbxSituacao_ven.text, 1) = "4" And pcVenda.BackColor = vbYellow Then
            strSQL = strSQL & " and dataefet_ven<=" & F_GravaData(txtbxAte.text)
        Else
            strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(txtbxAte.text & " 23:59:59")
        End If
    End If
    If Tipo_Operacao <> venda Then
        strSQL = strSQL & " and tipo_operacao>1 "
    Else
        strSQL = strSQL & " and tipo_operacao=1"
    End If
    strSQL = strSQL & " order by " & sOrd & " desc "
    
    
    
    
    
    
    abreObjetoAdo AdoVendas, strSQL, dtg
    
    
    
    P_Status ""
        
        Exit Sub
trata_Erro:
P_MostraErro


P_Status
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtg, Me, pcVenda
    posicionaADireita cmndbtnSair, pcBot
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    AdoVendas.Close
End Sub

Private Sub mnAbrirContasReceber_Click()
    On Error Resume Next
    If AdoVendas!cod_cli <> 0 Then
        AbrirContasReceberCliente AdoVendas!cod_cli
    End If
    
End Sub

Private Sub mnDevolução_Click()
    Tipo_Operacao = tipo_operacao_venda.Devolucao
    cmndbtnBuscar_Click
    
End Sub

Private Sub mnLancarProdutos_Click()
    Dim usuario As String
    With AdoVendas
    
        If !situacao_ven = "1" Then
            cmndbtnEdit_Click
        ElseIf !situacao_ven = "2" Then
            usuario = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & !abrevenda_usu & "'")
            If !abrevenda_usu <> strCod_Usu Then
                MsgBox "Apenas o usuários " & usuario & " pode adicionar produtos nessa venda!", vbExclamation
                Exit Sub
            Else
            
                AdicionaProdutoEmVendaFechada
            End If
        Else
            MsgBox "Não é possível lançar mais produtos nessa venda!", vbExclamation
        End If
    End With
End Sub
Sub AdicionaProdutoEmVendaFechada()
    Dim fr As New frmVenda
    P_Status "Montando venda..."
    
    
    With AdoVendas
        
        fr.tag = !num_ven
        'fr.lblAbert.caption = !abertura_ven
        'frmVenda.lblNome_Cli.caption = !nome_cli
        'frmVenda.lblNome_Cli.Tag = !cod_cli
        'frmVenda.lblHora.caption = !data_ven
        'frmVenda.lblCpf.caption = F_BuscaUmCampo("CPF_Cli", "tab_clientes", "cod_cli='" & !cod_cli & "'")
        'frmVenda.lblCPFCap.caption = "CPF/CNPJ:"
        'frmVenda.lblNome_Ven.caption = F_BuscaUmCampo("Guerra_Fun", "tab_Funcionarios", _
            "cod_fun='" & !cod_usu & "'")
        'frmVenda.txtbxEndereco_ven.Text = Format(!endereco_ven)
        fr.Atacado = True
        fr.lblValor.caption = FormatVal(!valor_ven)
        fr.caption = "Venda " & !num_ven & "-" & !nome_cli
        fr.status = "2"
        'frmVenda.txtbxObs.Text = Format(!obs_ven)
        'fr.Frame1.caption = Format(!nomecli_ven) & " Atendente[" & strNome_Usu & "]"
        fr.Show 0, mdiFrm
        'fr.lblAtacado.tag = "X" 'indica que a venda foi editada
    End With
    P_Status ""
    'dtVendas.Refresh
End Sub

Private Sub mnMudar_Click()
    muda_operacao
End Sub

Private Sub mnParcelas_Click()
    If strVisaoFinanceiro_usu <> "S" Or strVisaoCaixa_usu <> "S" Or strVisaoGeral_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação!", vbExclamation
        Exit Sub
    End If
    frmEditaMov.txtbxNumDoc.text = AdoVendas!num_ven
    frmEditaMov.txtbxDataIni.text = ""
    frmEditaMov.txtbxdataFin.text = ""
    frmEditaMov.Show 0
End Sub

Private Sub mnPrint_Click()
    cmndbtnImpr_Click
End Sub

Private Sub mnTroca_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para essa operação.", vbExclamation
        Exit Sub
    End If
    frmTroca.txtbxNum_Ven.text = AdoVendas!num_ven
    frmTroca.Show 1
End Sub

Private Sub tmr_Timer()
    If cmbxSituacao_ven.ListIndex = 0 Then
        'RefreshDataPosition adc
    End If
End Sub

Private Sub txtbxAte_GotFocus()
    P_SelecTexto txtbxAte
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
    P_SelecTexto txtbxDe
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
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        P_Busca
        
    End If
    If KeyCode = vbKeyDown Then
        
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
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
    P_Busca
End Sub

Private Sub txtProd_Change()
    txtProd.tag = ""
End Sub
