VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParametros 
   Caption         =   "Parâmetros da Empresa"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   10710
   Icon            =   "frmParametros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   10710
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   10680
      TabIndex        =   45
      Top             =   5400
      Width           =   10710
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   9120
         Picture         =   "frmParametros.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   735
         Left            =   7680
         Picture         =   "frmParametros.frx":0E58
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   120
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Dados da Empresa"
      TabPicture(0)   =   "frmParametros.frx":12DF
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Taxas"
      TabPicture(1)   =   "frmParametros.frx":12FB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Plano de Contas"
      TabPicture(2)   =   "frmParametros.frx":1317
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label28"
      Tab(2).Control(1)=   "Picture3"
      Tab(2).Control(2)=   "txtMensagemVenda"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "SPED"
      TabPicture(3)   =   "frmParametros.frx":1333
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "NFe"
      TabPicture(4)   =   "frmParametros.frx":134F
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Configurações"
      TabPicture(5)   =   "frmParametros.frx":136B
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Picture6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4065
         ScaleWidth      =   10305
         TabIndex        =   96
         Top             =   480
         Width           =   10335
         Begin VB.CheckBox chkTef 
            Caption         =   "Usa TEF"
            Height          =   315
            Left            =   120
            TabIndex        =   104
            Top             =   2400
            Width           =   3855
         End
         Begin VB.TextBox txtCaminhoExec 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   102
            Top             =   960
            Width           =   6975
         End
         Begin VB.CheckBox chkUsaCaixaSeparado_Emp 
            Caption         =   "Usa Caixa Separado das Vendas"
            Height          =   315
            Left            =   120
            TabIndex        =   101
            Top             =   2040
            Width           =   3855
         End
         Begin VB.CheckBox chkUsaImpedirProdutoSemEstoque 
            Caption         =   "Usa Impedir Venda de Produto sem Estoque"
            Height          =   315
            Left            =   120
            TabIndex        =   100
            Top             =   1680
            Width           =   3855
         End
         Begin VB.CheckBox chkUsaVendaSupermercado 
            Caption         =   "Usa Venda em Modo Supermercado"
            Height          =   315
            Left            =   120
            TabIndex        =   99
            Top             =   1320
            Width           =   3855
         End
         Begin VB.TextBox txtFormaPagamentoPadrao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   97
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho para procurar atualizações:"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   720
            Width           =   2580
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma Pgto Padrão"
            Height          =   195
            Left            =   120
            TabIndex        =   98
            Top             =   120
            Width           =   1365
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -74880
         ScaleHeight     =   3585
         ScaleWidth      =   10185
         TabIndex        =   83
         Top             =   780
         Width           =   10215
         Begin VB.ComboBox cmbAmb 
            Height          =   315
            ItemData        =   "frmParametros.frx":1387
            Left            =   1320
            List            =   "frmParametros.frx":1391
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtInfAddNf55 
            Appearance      =   0  'Flat
            Height          =   1125
            Left            =   120
            MaxLength       =   12
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   2160
            Width           =   9135
         End
         Begin VB.TextBox txtTipoICMS_Nfe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   12
            TabIndex        =   90
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtCaminhoCompiaXML 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   88
            Top             =   1560
            Width           =   9135
         End
         Begin VB.TextBox txtCRT 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   86
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtVersaoNFe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   84
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ambiente"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1320
            TabIndex        =   94
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informação Adicional (NF-e 55)"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   1920
            Width           =   2190
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de ICMS"
            Height          =   195
            Left            =   1320
            TabIndex        =   91
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho das Cópias de XML NFe"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   1320
            Width           =   2385
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CRT"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   330
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Versão"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -74880
         ScaleHeight     =   4065
         ScaleWidth      =   10185
         TabIndex        =   66
         Top             =   780
         Width           =   10215
         Begin VB.TextBox txtDescontoSugerido 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   5280
            MaxLength       =   4
            TabIndex        =   80
            Text            =   "0,0"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtDescontoMaximo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   5280
            MaxLength       =   4
            TabIndex        =   79
            Text            =   "0,0"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtComissVenda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   5280
            MaxLength       =   4
            TabIndex        =   77
            Text            =   "0,0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtTx_Federal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   71
            Text            =   "3,0"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtTx_Adm 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   70
            Text            =   "10,0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtMargemSeg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   69
            Text            =   "0,0"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtTx_Estadual 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   68
            Text            =   "0,0"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtJuros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   5280
            MaxLength       =   4
            TabIndex        =   67
            Text            =   "0,0"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Sugerido"
            Height          =   195
            Left            =   3690
            TabIndex        =   82
            Top             =   1680
            Width           =   1365
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Máximo"
            Height          =   195
            Left            =   3780
            TabIndex        =   81
            Top             =   1200
            Width           =   1275
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comissão s/ Vendas"
            Height          =   195
            Left            =   3660
            TabIndex        =   78
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tx. Federal"
            Height          =   195
            Left            =   960
            TabIndex        =   76
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tx. Adm."
            Height          =   195
            Left            =   1110
            TabIndex        =   75
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marg. Seg."
            Height          =   195
            Left            =   975
            TabIndex        =   74
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tx. Estadual"
            Height          =   195
            Left            =   870
            TabIndex        =   73
            Top             =   1680
            Width           =   885
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Juros s/ Atrazo"
            Height          =   195
            Left            =   4050
            TabIndex        =   72
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74760
         ScaleHeight     =   825
         ScaleWidth      =   10065
         TabIndex        =   59
         Top             =   900
         Width           =   10095
         Begin VB.ComboBox cmbxPerfil 
            Height          =   315
            ItemData        =   "frmParametros.frx":13B3
            Left            =   120
            List            =   "frmParametros.frx":13C0
            Sorted          =   -1  'True
            TabIndex        =   62
            Text            =   "C - Perfil C"
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cmbInd_Atv 
            Height          =   315
            ItemData        =   "frmParametros.frx":13EE
            Left            =   2280
            List            =   "frmParametros.frx":13F8
            Sorted          =   -1  'True
            TabIndex        =   61
            Text            =   "1 - Outros"
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtSuframa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   60
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label42 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Perfil Sped:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label43 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Indicador de Atividade:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2280
            TabIndex        =   64
            Top             =   120
            Width           =   1635
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suframa:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4440
            TabIndex        =   63
            Top             =   120
            Width           =   630
         End
      End
      Begin VB.PictureBox pc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74880
         ScaleHeight     =   825
         ScaleWidth      =   10185
         TabIndex        =   48
         Top             =   4020
         Width           =   10215
         Begin VB.TextBox txtCpf 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            MaxLength       =   14
            TabIndex        =   51
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtRg 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6600
            MaxLength       =   14
            TabIndex        =   50
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtResponsavel 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   49
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&PF:"
            Height          =   195
            Left            =   4440
            TabIndex        =   54
            Top             =   120
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&RG:"
            Height          =   195
            Left            =   6600
            TabIndex        =   53
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Responsável:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox txtMensagemVenda 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   -74880
         MaxLength       =   200
         TabIndex        =   36
         Top             =   3420
         Width           =   7215
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   -74880
         ScaleHeight     =   1665
         ScaleWidth      =   7545
         TabIndex        =   23
         Top             =   780
         Width           =   7575
         Begin VB.TextBox txtPlaVendas 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   31
            Top             =   120
            Width           =   4215
         End
         Begin VB.CommandButton cmndbtnBuscaCCVendas 
            Caption         =   "..."
            Height          =   255
            Left            =   6360
            TabIndex        =   30
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtplaCompras 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   29
            Top             =   480
            Width           =   4215
         End
         Begin VB.CommandButton cmndbtnCCDDuplPag 
            Caption         =   "..."
            Height          =   255
            Left            =   6360
            TabIndex        =   28
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox txtPLaDev 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   27
            Top             =   840
            Width           =   4215
         End
         Begin VB.CommandButton cmndbtnCCDuplRec 
            Caption         =   "..."
            Height          =   255
            Left            =   6360
            TabIndex        =   26
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox txtPlaNC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   25
            Top             =   1200
            Width           =   4215
         End
         Begin VB.CommandButton cmndbtnCCNaoClassC 
            Caption         =   "..."
            Height          =   255
            Left            =   6360
            TabIndex        =   24
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compras"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Devoluções"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não classificados"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   1245
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   -74880
         ScaleHeight     =   3225
         ScaleWidth      =   10185
         TabIndex        =   1
         Top             =   660
         Width           =   10215
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8640
            MaxLength       =   40
            TabIndex        =   58
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   16
            TabIndex        =   55
            Top             =   2880
            Width           =   9615
         End
         Begin VB.TextBox txtIm 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8640
            MaxLength       =   12
            TabIndex        =   43
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox txtBairro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   40
            Top             =   1680
            Width           =   4215
         End
         Begin VB.TextBox txtNumero 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            MaxLength       =   40
            TabIndex        =   38
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtFantasia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            MaxLength       =   40
            TabIndex        =   11
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtNome 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   10
            Top             =   360
            Width           =   5535
         End
         Begin VB.TextBox txtLogradouro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   9
            Top             =   1080
            Width           =   4215
         End
         Begin VB.TextBox txtComplemento 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            MaxLength       =   40
            TabIndex        =   8
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtCidade 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            MaxLength       =   40
            TabIndex        =   7
            Top             =   1680
            Width           =   4095
         End
         Begin VB.TextBox txtCep 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8640
            MaxLength       =   12
            TabIndex        =   6
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtCnpj 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            MaxLength       =   18
            TabIndex        =   5
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txtFax 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            MaxLength       =   16
            TabIndex        =   4
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txtFone 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   16
            TabIndex        =   3
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txtIe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6600
            MaxLength       =   12
            TabIndex        =   2
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmndbtnBuscar 
            Height          =   285
            Left            =   9720
            Picture         =   "frmParametros.frx":1433
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   2640
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IM:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   8640
            TabIndex        =   44
            Top             =   2040
            Width           =   225
         End
         Begin VB.Label lblUF 
            Caption         =   "UF"
            Height          =   255
            Left            =   8040
            TabIndex        =   42
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4440
            TabIndex        =   39
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Fantasia:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5760
            TabIndex        =   22
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço(Logradouro):"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5760
            TabIndex        =   19
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Cidade:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4440
            TabIndex        =   18
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CE&P:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   8640
            TabIndex        =   17
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&NPJ:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4440
            TabIndex        =   16
            Top             =   2040
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Fone:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F&ax:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   2280
            TabIndex        =   14
            Top             =   2040
            Width           =   300
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&IE:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   6600
            TabIndex        =   13
            Top             =   2040
            Width           =   195
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Código:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   8640
            TabIndex        =   12
            Top             =   1440
            Width           =   540
         End
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Mensagem Promocional da Venda:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -74880
         TabIndex        =   37
         Top             =   3180
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim bPrimeiroAcesso As Boolean

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim strSQL As String
    If bPrimeiroAcesso = False Then
        If Not F_senha Then
        
            Exit Sub
        End If
        
        If strFechames_usu = "N" Then
            MsgBox "Você não tem permissão para esse acesso!", vbExclamation
            Exit Sub
        End If
    End If
    strSQL = "update tab_parametros SET NOME='" & txtNome.Text & "',"
    strSQL = strSQL & "logradouro='" & txtLogradouro.Text & "',"
    strSQL = strSQL & " BAIRRO='" & txtBairro & "',"
    strSQL = strSQL & " CIDADE='" & txtCidade.Text & "',"
    strSQL = strSQL & "Numero='" & txtNumero.Text & "',"
   
    strSQL = strSQL & "UF='" & lblUF.caption & "',"
    strSQL = strSQL & "CEP='" & txtCep.Text & "',"
    strSQL = strSQL & "FONE='" & txtFone.Text & "',"
    strSQL = strSQL & "FAX='" & txtFax.Text & "',"
    strSQL = strSQL & "CNPJ='" & txtCnpj.Text & "',"
    strSQL = strSQL & "IE='" & txtIe.Text & "',"
    strSQL = strSQL & "responsavel='" & txtResponsavel.Text & "',"
    strSQL = strSQL & "cpf_responsavel='" & txtCpf.Text & "',"
    strSQL = strSQL & "rg_responsavel='" & txtRg.Text & "',"
    strSQL = strSQL & "tx_federal=" & F_GravaDec(txtTx_Federal.Text) & ","
    strSQL = strSQL & "tx_adm=" & F_GravaDec(txtTx_Adm.Text) & ","
    strSQL = strSQL & "tx_Estadual=" & F_GravaDec(txtTx_Federal.Text) & ","
    strSQL = strSQL & "MargemSeg=" & F_GravaDec(txtMargemSeg.Text) & ","
    strSQL = strSQL & "MensagemVenda='" & txtMensagemVenda.Text & "',"
    strSQL = strSQL & "comissven_par=" & F_GravaDec(txtComissVenda.Text) & ","
    strSQL = strSQL & "juros_par=" & F_GravaDec(txtJuros.Text) & ","
    strSQL = strSQL & "data_alteracao=" & F_GravaData(Date) & ","
    strSQL = strSQL & "DescontoMaximo=" & F_GravaDec(txtDescontoMaximo.Text) & ","
   ' strSQL = strSQL & "DescontoSugerido=" & F_GravaDec(txtDescontoSugerido.Text) & ","
    strSQL = strSQL & "caminhoexec_par='" & txtCaminhoExec.Text & "',"
    
    strSQL = strSQL & "complemento='" & txtComplemento.Text & "',"
    strSQL = strSQL & "fantasia='" & txtFantasia.Text & "',"
    strSQL = strSQL & "cod_cid='" & txtCodigo.Text & "',"
    strSQL = strSQL & "CRT='" & txtCRT.Text & "',"
    strSQL = strSQL & "UsaCaixaSeparado='" & IIf(chkUsaCaixaSeparado_Emp.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "UsaTef='" & IIf(chkTef.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "PlanoContasVenda='" & txtPlaVendas.Text & "',"
    strSQL = strSQL & "PlanoContasCompra='" & txtplaCompras.Text & "',"
    strSQL = strSQL & "usacaixaseparado_emp='" & IIf(chkUsaCaixaSeparado_Emp.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "ambiente_nfe='" & Left(cmbAmb.Text, 1) & "',"
    strSQL = strSQL & "tipoicms_nfe='" & txtTipoICMS_Nfe.Text & "',"
    strSQL = strSQL & "usavendasupermercado='" & IIf(chkUsaVendaSupermercado.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "formapagamentopadrao='" & txtFormaPagamentoPadrao.tag & "',"
    strSQL = strSQL & "CaminhoCopiaXML='" & txtCaminhoCompiaXML.Text & "',"
    strSQL = strSQL & "UsaImpedirVendaDeProdutoSemEstoque='" & IIf(chkUsaImpedirProdutoSemEstoque.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "versao_nfe='" & txtVersaoNFe.Text & "',"
    strSQL = strSQL & "infaddnfe55='" & txtInfAddNf55.Text & "' "
 
    
    conn.Execute strSQL
    
            
   
    MsgBox "Dados alterados..."
    
End Sub

Private Sub Form_Activate()
    If bActive = True Then
        Exit Sub
    End If
    
        Carrega
    
    bActive = True
End Sub

Sub Carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim c As Cidade
    strSQL = "select * from " & TAB_PARAMETROS
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            bPrimeiroAcesso = True
            Exit Sub
        End If
        txtNome.Text = !nome
        txtLogradouro.Text = !Logradouro
        txtNumero.Text = !numero
        txtFantasia.Text = !fantasia
        txtBairro.Text = !bairro
        txtNumero.Text = !numero
        txtComplemento.Text = !complemento
        txtCep.Text = !cep
        txtCodigo.Text = !cod_cid
        c = F_BuscaCidadeCodigo(!cod_cid)
        txtCidade.Text = c.nome_cid
        txtCidade.tag = c.cod_cid
        lblUF.caption = c.uf_cid
        txtFone.Text = !fone
        txtFax.Text = !fax
        txtCnpj.Text = !CNPJ
        txtIe.Text = !IE
        'txtIm.Text = !IM
        
        
        txtResponsavel.Text = !responsavel
        txtCpf.Text = !cpf_responsavel
        txtRg.Text = !rg_responsavel
        
        
        txtTx_Adm.Text = !tx_adm
        txtTx_Estadual.Text = !tx_Estadual
        txtTx_Federal.Text = !tx_federal
        txtMargemSeg.Text = !MargemSeg
        txtDescontoMaximo.Text = !descontomaximo
       ' txtDescontoSugerido.Text = !descontosugerido
        txtComissVenda.Text = !comissven_par
        txtJuros.Text = !juros_par
        txtplaCompras.Text = !planocontascompra
        txtplaCompras.Text = !planocontasvenda
        txtMensagemVenda.Text = TiraNull(!mensagemvenda)
        txtVersaoNFe.Text = !versao_nfe
        P_LocalizaOperacao !ambiente_nfe, cmbAmb
        txtCRT.Text = !CRT
        txtCaminhoCompiaXML.Text = !caminhocopiaxml
        txtInfAddNf55.Text = TiraNull(!infaddnfe55)
        txtTipoICMS_Nfe.Text = !tipoicms_nfe
        txtFormaPagamentoPadrao.Text = F_BuscaFormasPag(!formapagamentopadrao, "").nome_For
        txtFormaPagamentoPadrao.tag = (!formapagamentopadrao)
        txtCaminhoExec.Text = !caminhoexec_par
        If !usatef = "S" Then
            chkTef.Value = 1
        Else
            chkTef.Value = 0
        End If
        If !usacaixaseparado = "S" Then
            chkUsaCaixaSeparado_Emp.Value = 1
        Else
            chkUsaCaixaSeparado_Emp.Value = 0
        End If
        If !usaimpedirvendadeprodutosemestoque = "S" Then
            chkUsaImpedirProdutoSemEstoque.Value = 1
        Else
            chkUsaImpedirProdutoSemEstoque.Value = 0
        End If
        If !usavendasupermercado = "S" Then
            chkUsaVendaSupermercado.Value = 1
        Else
            chkUsaVendaSupermercado.Value = 1
        End If
    End With
    
    sstb.tag = 0
End Sub

Private Sub Form_Load()
    bActive = False
    
End Sub

Private Sub txtcodigo_GotFocus()
    P_SelecTexto txtCodigo
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCid
        txtFone.SetFocus
    End If
End Sub
Sub buscaCid()
    Dim c As Cidade
    c = F_BuscaCidades(txtCodigo.Text)
    If c.cod_cid <> "" Then
        txtCodigo.Text = c.cod_cid
        txtCidade.Text = c.nome_cid
        txtCidade.tag = c.cod_cid
        lblUF.caption = c.uf_cid
        
    End If
    
End Sub
