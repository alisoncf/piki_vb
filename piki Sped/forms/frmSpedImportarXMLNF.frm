VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form frmSpedImportarXMLNF 
   Caption         =   "Importar Nota Fiscal Eletrônica de Entrada"
   ClientHeight    =   7500
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16710
   Icon            =   "frmSpedImportarXMLNF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   16710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   11280
      TabIndex        =   62
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   16650
      TabIndex        =   60
      Top             =   7245
      Width           =   16710
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Height          =   495
      Left            =   7920
      TabIndex        =   52
      Top             =   0
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Produtos"
      TabPicture(0)   =   "frmSpedImportarXMLNF.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtQtd"
      Tab(0).Control(1)=   "txtNome"
      Tab(0).Control(2)=   "pcProdutoVinculado"
      Tab(0).Control(3)=   "txtGrid"
      Tab(0).Control(4)=   "cmdProduto"
      Tab(0).Control(5)=   "grid"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Imposto"
      TabPicture(1)   =   "frmSpedImportarXMLNF.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pcImpostos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Duplicatas/Frete"
      TabPicture(2)   =   "frmSpedImportarXMLNF.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "gridDup"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdLancarContaAPagar"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdLancarContaAPagar 
         Caption         =   "Lançar Conta a Pagar"
         Height          =   495
         Left            =   13320
         TabIndex        =   66
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtQtd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   61
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72360
         TabIndex        =   59
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox pcProdutoVinculado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -73320
         ScaleHeight     =   375
         ScaleWidth      =   11535
         TabIndex        =   57
         Top             =   360
         Width           =   11535
         Begin VB.Label lblProdutoVinculado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto  Vinculado:"
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
            TabIndex        =   58
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.TextBox txtGrid 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73680
         TabIndex        =   56
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdProduto 
         Caption         =   "Produto <Enter>"
         Height          =   375
         Left            =   -74880
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox pcImpostos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3960
         Left            =   -74880
         ScaleHeight     =   3930
         ScaleWidth      =   13065
         TabIndex        =   19
         Top             =   480
         Width           =   13095
         Begin VB.CommandButton cmdRecalc 
            Caption         =   "Recalcular"
            Height          =   495
            Left            =   9960
            TabIndex        =   65
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdFreteExterno 
            Caption         =   "Frete"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtValorProd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   4800
            MaxLength       =   12
            TabIndex        =   34
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   33
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtOut 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   32
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtIpi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   31
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtValorTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   405
            Left            =   6480
            MaxLength       =   12
            TabIndex        =   30
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtFrete 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   29
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtBCicms 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   28
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtVlIcms 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   27
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtBcIcmsSt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   26
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtVlIcmsSt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   25
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSeguro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            MaxLength       =   12
            TabIndex        =   24
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtPis 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5520
            MaxLength       =   12
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtCofins 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6600
            MaxLength       =   12
            TabIndex        =   22
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtPisSt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7680
            MaxLength       =   12
            TabIndex        =   21
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtCofinsSt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8760
            MaxLength       =   12
            TabIndex        =   20
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label34 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Produtos:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   50
            Tag             =   "!"
            Top             =   0
            Width           =   1080
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desco&nto:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2280
            TabIndex        =   49
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Outras Desp.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3360
            TabIndex        =   48
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "I&PI:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4440
            TabIndex        =   47
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor Total"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6480
            TabIndex        =   46
            Tag             =   "!"
            Top             =   0
            Width           =   765
         End
         Begin VB.Label lblUF_Ter 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "--"
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
            Left            =   11745
            TabIndex        =   45
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Frete:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   405
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BC ICMS:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   690
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "VL &ICMS:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   42
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label53 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BC ICMS ST:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2280
            TabIndex        =   41
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label54 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "VL &ICMS ST:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3360
            TabIndex        =   40
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label55 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Seguro:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   39
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label56 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5520
            TabIndex        =   38
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label57 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "COFINS"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6600
            TabIndex        =   37
            Top             =   720
            Width           =   585
         End
         Begin VB.Label Label58 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS ST:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7680
            TabIndex        =   36
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label59 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "COFINS ST:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8760
            TabIndex        =   35
            Top             =   720
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   6376
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid gridDup 
         Height          =   3975
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7011
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.PictureBox pcNota 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   16305
      TabIndex        =   4
      Top             =   720
      Width           =   16335
      Begin VB.CommandButton cmdCadFor 
         Caption         =   "Cadastrar"
         Height          =   255
         Left            =   1200
         TabIndex        =   54
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtEmitente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCNPJ 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtEmissao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtChaveNfe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9000
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   4440
         TabIndex        =   63
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lblFornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Emitente:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Número:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6240
         TabIndex        =   15
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Série:"
         Height          =   195
         Left            =   7320
         TabIndex        =   14
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "CNPJ:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4440
         TabIndex        =   13
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Emissão: "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         TabIndex        =   12
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chave NFE:"
         Height          =   195
         Left            =   9000
         TabIndex        =   11
         Top             =   120
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdCarrega 
      Caption         =   "Carregar XML"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   480
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCaminho 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdArquivo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      Picture         =   "frmSpedImportarXMLNF.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho do Arquivo:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1470
   End
   Begin VB.Menu mnMenu 
      Caption         =   "menu"
      Begin VB.Menu mnCop 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu mnNovoProd 
         Caption         =   "&Novo Produto"
      End
      Begin VB.Menu mnEdit 
         Caption         =   "&Editar este produto"
      End
      Begin VB.Menu mnBuscarProdCodBarra 
         Caption         =   "&Buscar Produto por Código de Barras"
      End
      Begin VB.Menu mnCodFab 
         Caption         =   "Buscar Produto pela Referência"
      End
      Begin VB.Menu mnBuscaProdDesc 
         Caption         =   "&Buscar produto pela descrição"
      End
      Begin VB.Menu mnNaoEEsse 
         Caption         =   "Não é esse?"
      End
      Begin VB.Menu mnVincularProd 
         Caption         =   "Vincular Todos os Produtos"
      End
      Begin VB.Menu mnLimpVinc 
         Caption         =   "Limpar Vinculação"
      End
   End
End
Attribute VB_Name = "frmSpedImportarXMLNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As XMLSped
Dim qtd_prod_sped As Integer
Private Sub cmdArquivo_Click()
    cmd.filter = "*.*"
    cmd.FileName = "*.xml"
    cmd.ShowOpen
    
    txtCaminho.Text = cmd.FileName
    cmdCarrega_Click
    
    
    
End Sub

Private Sub cmdCadFor_Click()
    frmCadParticipante.txtBairro.Text = X.cab.e.xBairro
    frmCadParticipante.txtCNPJ.Text = X.cab.e.CNPJ
    frmCadParticipante.txtIE.Text = X.cab.e.IE
    frmCadParticipante.txtCodMun.Text = X.cab.e.cMun
    frmCadParticipante.txtCodPais.Text = X.cab.e.cPais
    frmCadParticipante.txtcomplemento.Text = ""
    frmCadParticipante.txtEndereco.Text = X.cab.e.xLgr
    frmCadParticipante.txtIE.Text = X.cab.e.IE
    frmCadParticipante.txtNome.Text = X.cab.e.xNome
    frmCadParticipante.txtNumero.Text = X.cab.e.nro
    frmCadParticipante.txtsuframa.Text = ""
    
    frmCadParticipante.Show 1
    
    
End Sub
Sub BuscaSpedEmpresa(codigo As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from sped_empresa where cnpj='" & codigo & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            lblEmpresa.caption = !apelido & "-" & !nome & "(" & !codigo & ") - CNPJ: " & !CNPJ
            lblEmpresa.tag = !codigo
            
        Else
            lblEmpresa.caption = X.cab.d.xNome & "(Não encontrado)"
            MsgBox "Esta empresa não está cadastrada como SPED_EMPRESA", vbExclamation
            Exit Sub
        End If
        .Close
    End With
    
    
    
    
End Sub
Private Sub cmdCarrega_Click()
    On Error GoTo trata_Erro
    
    P_Status "Carregando dados..."
    
    X = Le_XML(txtCaminho.Text)
    txtEmitente.Text = X.cab.e.xNome
    txtCNPJ.Text = X.cab.e.CNPJ
    txtNumero.Text = X.cab.cNF
    txtSerie.Text = X.cab.serie
    txtEmissao.Text = Format(X.cab.dEmi, "dd/mm/yyyy")
    txtChaveNfe.Text = X.CH_NFE
    txtValorProd.Text = ConverteSpedValor(X.tot.vProd)
    txtValorTotal.Text = ConverteSpedValor(X.tot.vNF)
    txtBCicms.Text = ConverteSpedValor(X.tot.vBC)
    txtVlIcms.Text = ConverteSpedValor(X.tot.vICMS)
    txtBcIcmsSt.Text = ConverteSpedValor(X.tot.vBCST)
    txtVlIcms.Text = ConverteSpedValor(X.tot.vST)
    txtFrete.Text = ConverteSpedValor(X.tot.vFrete)
    txtSeguro.Text = ConverteSpedValor(X.tot.vSeg)
    txtOut.Text = ConverteSpedValor(X.tot.vOutro)
    txtDesconto.Text = ConverteSpedValor(X.tot.vDesc)
    txtIpi.Text = ConverteSpedValor(X.tot.vIPI)
    txtPis.Text = ConverteSpedValor(X.tot.vPIS)
    txtCofins.Text = ConverteSpedValor(X.tot.vCOFINS)
    BuscaSpedEmpresa X.cab.d.CNPJ
    
    
    
    
    carregaGrid
    
    
    
    VerificaFornecedorCadastrado
    
    P_Status
    Exit Sub
trata_Erro:
    P_Status
    MostraErro
End Sub
Sub VerificaFornecedorCadastrado()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    If X.cab.e.CNPJ <> "" Then
        strSQL = "select * from sped_participante where cnpj='" & X.cab.e.CNPJ & "'"
        abreObjetoAdo rPesq, strSQL
        With rPesq
            If .RecordCount > 0 Then
                lblFornecedor.caption = "Participante " & !codigo
                lblFornecedor.tag = !cod_for
            Else
                lblFornecedor.caption = "Participante não encontrado"
            End If
        End With
        
        
        
    End If
    
End Sub

Sub carregaGrid()
    
    Dim i As Integer
    Dim strSQL As String
    grid.Visible = False
    With grid
        .rows = 1
        .Cols = 25
        .col = 0: .Text = "..."
        .col = 1: .Text = "cEAN"
        .col = 2: .Text = "Código"
        .col = 3: .Text = "Produto"
        .col = 4: .Text = "CFOP"
        .col = 5: .Text = "cProd"
        .col = 6: .Text = "icms_"
        .col = 7: .Text = "IcmsST"
        .col = 8: .Text = "ipi_"
        .col = 9: .Text = "NCM"
        .col = 10: .Text = "pis_"
        .col = 11: .Text = "qCom"
        .col = 12: .Text = "qTrib"
        .col = 13: .Text = "Qtd.Convert."
        .col = 14: .Text = "uCom"
        .col = 15: .Text = "vProd"
        .col = 16: .Text = "vUnCom"
        .col = 17: .Text = "vUnTrib"
        .col = 18: .Text = "xProd"
        .col = 19: .Text = "VFrete"
        .col = 20: .Text = "VSeguro"
        .col = 21: .Text = "VOutro"
        .col = 22: .Text = "VDesc"
        .col = 23: .Text = "pICMS"
        .col = 24: .Text = "Custo"
        
        .ColWidth(3) = 2000
        .ColWidth(1) = 1500
        .ColWidth(13) = 1200
        .ColWidth(14) = 1000
        .ColWidth(18) = 3000
        
        For i = 0 To UBound(X.p) - 1
            strSQL = i + 1 & Chr(9)
            strSQL = strSQL & X.p(i).cEAN & Chr(9)
            strSQL = strSQL & "" & Chr(9) & "" & Chr(9) 'código + descrição
            strSQL = strSQL & ConverteCfop(X.p(i).CFOP) & Chr(9)
            strSQL = strSQL & X.p(i).cProd & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).icms_.vICMS) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).icms_.vICMSST) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).ipi_.vIPI) & Chr(9)
            
            
            strSQL = strSQL & X.p(i).ncm & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).pis_.vPIS) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).qCom) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).qTrib) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).qTrib) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).uCom) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).vProd) & Chr(9)
            strSQL = strSQL & FormatVal(ConverteSpedValor(X.p(i).vUnCom)) & Chr(9)
            strSQL = strSQL & FormatVal(ConverteSpedValor(X.p(i).vUnTrib)) & Chr(9)
            strSQL = strSQL & X.p(i).xProd & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).vFrete) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).vSeguro) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).vOutro) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).vDesc) & Chr(9)
            strSQL = strSQL & ConverteSpedValor(X.p(i).icms_.pICMS)
            
            
            .AddItem strSQL
            grid.row = grid.rows - 1
            grid.col = 13
            grid.CellFontBold = True
            grid.CellBackColor = vbYellow
            grid.col = 2
            grid.CellFontBold = True
            grid.CellBackColor = vbYellow
        Next
    End With
    
    
    
    With gridDup
        .rows = 1
        .Cols = 8
        .col = 0: .Text = "..."
        .col = 1: .Text = "Fatura"
        .col = 2: .Text = "Vlr Orig."
        .col = 3: .Text = "Vlr Desc"
        
        .col = 4: .Text = "Valor Final"
        .col = 5: .Text = "nDup"
        .col = 6: .Text = "dVenc"
        .col = 7: .Text = "vDup"
        
           
        
        For i = 0 To UBound(X.c.d) - 1
            strSQL = i + 1 & Chr(9)
            strSQL = strSQL & X.c.fat.nFat & Chr(9)
            strSQL = strSQL & FormatValNFe(X.c.fat.vOrig) & Chr(9)
            strSQL = strSQL & FormatValNFe(X.c.fat.vDesc) & Chr(9)
            strSQL = strSQL & FormatValNFe(X.c.fat.vLiq) & Chr(9)
            strSQL = strSQL & X.c.d(i).nDup & Chr(9)
            strSQL = strSQL & Format(X.c.d(i).dVenc, "dd/mm/yyyy") & Chr(9)
            strSQL = strSQL & FormatValNFe(X.c.d(i).vDup)
            If X.c.d(i).nDup <> "" Then
                .AddItem strSQL
            End If
        Next
    End With
    ZebraGrid gridDup, 1
    grid.Visible = True
    
End Sub

Private Sub cmdFreteExterno_Click()
    calcCusto
End Sub

Private Sub cmdImportar_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim custo As Currency
    Dim qtd As Currency
    Dim idNf As String
    Dim CodPart As String
    Dim i As Integer
    
    
    If grid.rows = 1 Then
        MsgBox "Não há produtos!", vbExclamation
        Exit Sub
    End If
    
    
    strSQL = "select * from sped_nf where chv_nfe='" & X.CH_NFE & "'"
    
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        MsgBox "Esta nota já se encontra gravada no banco de dados!", vbExclamation
        
        Exit Sub
    End If
    
    If MsgBox("Confirma importação da NFe?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    
    strSQL = "select * from sped_participante where cnpj='" & X.cab.e.CNPJ & "'"
    abreObjetoAdo rPesq, strSQL
    
    
    If rPesq.RecordCount > 0 Then
        CodPart = rPesq!codigo
        
    Else
        If MsgBox("O Participante com CNPJ " & X.cab.e.CNPJ & " não foi encontrado. Deseja cadastrá-lo?", vbYesNo + vbQuestion) = vbYes Then
            If GravaParticipante(X.cab.e.xNome, X.cab.e.cPais, X.cab.e.CNPJ, "", X.cab.e.IE, X.cab.e.cMun, _
            "", X.cab.e.xLgr, X.cab.e.nro, "", X.cab.e.xBairro, "", "", X.cab.e.cep, X.cab.e.UF) Then
                CodPart = F_BuscaUmCampo("codigo", "sped_participante", "cnpj='" & X.cab.e.CNPJ & "'")
            End If
            
        End If
    End If
    
    idNf = SpedNF.GravaSpedNf("0", "0", CodPart, X.cab.mod, "00", X.cab.serie, X.cab.nNF, X.CH_NFE, X.cab.dEmi, _
        Format(Date, "dd/mm/yyyy"), ConverteSpedValor(X.tot.vNF), X.cab.indPag, ConverteSpedValor(X.tot.vDesc), _
        "0", ConverteSpedValor(X.tot.vProd), X.t.modFrete, ConverteSpedValor(X.tot.vFrete), ConverteSpedValor(X.tot.vSeg), _
        ConverteSpedValor(X.tot.vOutro), ConverteSpedValor(X.tot.vBC), ConverteSpedValor(X.tot.vICMS), _
        ConverteSpedValor(X.tot.vBCST), ConverteSpedValor(X.tot.vST), ConverteSpedValor(X.tot.vIPI), _
        ConverteSpedValor(X.tot.vPIS), ConverteSpedValor(X.tot.vCOFINS), "0", "0", "1", CStr(Date), lblEmpresa.tag)
    
    For i = 1 To UBound(X.p)
        grid.row = i
        grid.col = 13
        X.p(i - 1).qtd_convertida = grid.Text
        qtd = CCur(grid.Text)
        grid.col = 2
        X.p(i - 1).codigo = grid.Text
        grid.col = 3
        X.p(i - 1).nome_pro = grid.Text
        
        grid.col = 24
        custo = CCur(grid.Text) / qtd
        
        
        If SpedNF.InsereItemSpedNF(idNf, CStr(i), X.p(i - 1).codigo, X.p(i - 1).cProd, X.p(i - 1).xProd, ConverteSpedValor(X.p(i - 1).qCom), _
            X.p(i - 1).uCom, ConverteSpedValor(X.p(i - 1).qCom), X.p(i - 1).uCom, ConverteSpedValor(X.p(i - 1).vProd), ConverteSpedValor(X.p(i - 1).vUnCom), ConverteSpedValor(X.p(i - 1).vDesc), "0", _
            X.p(i - 1).icms_.cst, X.p(i - 1).CFOP, _
            "", ConverteSpedValor(X.p(i - 1).icms_.vBC), ConverteSpedValor(X.p(i - 1).icms_.pICMS), ConverteSpedValor(X.p(i - 1).icms_.vICMS), _
            X.p(i - 1).icms_.cst, ConverteSpedValor(X.p(i - 1).icms_.pICMSST), _
            ConverteSpedValor(X.p(i - 1).icms_.vICMSST), "0", X.p(i - 1).ipi_.cst, X.p(i - 1).ipi_.cEnq, ConverteSpedValor(X.p(i - 1).ipi_.vBC), _
            ConverteSpedValor(X.p(i - 1).ipi_.pIPI), ConverteSpedValor(X.p(i - 1).ipi_.vIPI), ConverteSpedValor(X.p(i - 1).pis_.cst), _
            ConverteSpedValor(X.p(i - 1).pis_.vBC), ConverteSpedValor(X.p(i - 1).pis_.pPIS), ConverteSpedValor(X.p(i - 1).qCom), "0", _
            ConverteSpedValor(X.p(i - 1).pis_.vPIS), X.p(i - 1).cof_.cst, ConverteSpedValor(X.p(i - 1).cof_.vBC), _
            ConverteSpedValor(X.p(i - 1).cof_.pCOFINS), ConverteSpedValor(X.p(i - 1).qCom), "0", _
            ConverteSpedValor(X.p(i - 1).cof_.vCOFINS), "", X.p(i - 1).cEAN, X.p(i - 1).ncm, X.p(i - 1).qtd_convertida, CStr(custo), "0", "N") Then
        End If
    Next
    
    If MsgBox("Importação Concluída, deseja abrir a NF?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = F_BuscaUmCampo("max(id)", "sped_nf", "ind_oper=0 and status=1 and chv_nfe='" & X.CH_NFE & "'")
        frmEntradaSped.tag = strSQL
        frmEntradaSped.Show 1
    End If
    
End Sub

Private Sub cmdLancarContaAPagar_Click()
    LancarContaPagar
End Sub
Sub LancarContaPagar()
    Dim i As Integer
    Dim s As String
    Dim vFat As String
    Dim fat As String
    Dim sDup As String
    Dim sValor As String
    Dim sVen As String
    Dim sDoc As String
    Dim vDup As String
    Dim nDup As String
    If strLandoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    If gridDup.rows <= 1 Then
        MsgBox "Não há duplicata a lançar!", vbExclamation
        Exit Sub
    End If
    If lblFornecedor.tag = "" Then
        MsgBox "É necessário cadastrar o participante como fornecedor!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma lançamento da Duplicata?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With gridDup
        For i = 1 To .rows - 1
            .row = i
            .col = 1
            fat = .Text
            
            .col = 4
            sValor = .Text
            .col = 3
            
            .col = 5
            nDup = .Text
            .col = 6
            sVen = .Text
            .col = 7
            vDup = .Text
            
            
            s = ExisteDocSpedParticipante(fat, lblFornecedor.tag, "P", sDoc)
            If s <> "" Then
                MsgBox "Já existe uma fatura lançada com esse número para esse fornecedor", vbExclamation
                
            Else
                If sDoc = "" Then
                    sDoc = F_GravaDoc("COMPRA - NFE " & X.cab.nNF & " " & X.cab.e.xNome, fat, vFat, Date, lblFornecedor.tag, _
                    "P", "", gridDup.rows - 1, "P", "F", "24000")
                End If
                If F_GravaItDoc(sDoc, vDup, sVen, sVen, CStr(i), a_pagar_avulso, "", "1", "000", "") <> "" Then
                
                End If
                MsgBox "Fatura gravada com sucesso. Seq. Doc: " & sDoc, vbInformation
            End If
            
        Next
    
        
    
    
    End With
    
    
End Sub
Private Sub cmdProduto_Click()
    If grid.row = 0 Then
        MsgBox "Escolha uma linha válida!", vbExclamation
        Exit Sub
    End If
    
    
    
        
    
End Sub
Sub buscaProd()
    Dim p As Produtos.Produto
    
    p = Produtos.F_BuscaProdSaida(txtGrid.Text)
    If p.cod_pro <> "" Then
        grid.col = 2
        grid.Text = p.cod_pro
        txtGrid.Visible = False
        X.p(grid.row - 1).codigo = p.cod_pro
    End If
    
End Sub

Function confereProdutos() As Boolean
    
    
    
    
    
    
End Function

Private Sub cmdRecalc_Click()
    calcCusto
End Sub
Sub FreteExterno()
    Dim i As Integer
    Dim j As Integer
    Dim Indice As Double
    Dim valor_item As String
    Dim desconto_item As String
    Dim difal As Double
    Dim frete As String
    Dim vOutro As String
    Dim icms As String
    Dim custo  As Double
    On Error GoTo trata_Erro
    With grid
        For i = 1 To .rows - 1
            .row = i
            .col = 15
            
            Indice = CCur(.Text) / CCur(txtValorProd.Text)
             
            .col = 19
        
            .Text = Indice * CCur(txtFrete.Text)
             
            
        
        
        Next
    End With
    
Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub calcCusto()
    Dim i As Integer
    Dim j As Integer
    Dim difal As Currency
    Dim Indice As Double
    Dim valor_item As Currency
    Dim ICMSST As Currency
    Dim desconto_item As Currency
    Dim vSeguro As Currency
    Dim ipi As Currency
    Dim frete As Currency
    Dim vOutro As Currency
    Dim icms As Currency
    Dim custo  As Currency
    On Error GoTo trata_Erro
    
    FreteExterno
    
    P_Status "Calculando custos..."
    With grid
        For i = 1 To .rows - 1
            .row = i
            .col = 7
            ICMSST = CCur(FormatVal(.Text))
            
            .col = 8
            ipi = CCur(FormatVal(.Text))
            .col = 15
            valor_item = CCur(.Text)
            Indice = CCur(.Text) / CCur(txtValorProd.Text)
            
            .col = 19
            frete = CCur(FormatVal(.Text))
            .col = 20
            vSeguro = CCur(FormatVal(.Text))
            .col = 21
            vOutro = CCur(FormatVal(.Text))
            .col = 22
            desconto_item = CCur(FormatVal(.Text))
            .col = 23
            icms = CCur(FormatVal(.Text))
        
            difal = RetornaDifal(X.cab.e.UF, CStr(icms), CStr(valor_item), CStr(desconto_item), CStr(vSeguro), CStr(frete), CStr(vOutro))
        
            custo = valor_item + ICMSST + frete + vOutro + vSeguro - desconto_item + difal + ipi
        
            .col = 24
            .Text = FormatVal(custo, "###,##0.0000")
        Next
    End With
    P_Status ""
Exit Sub
trata_Erro:
P_Status
    P_MostraErro
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    resizeForm Me
    SSTab1.Tab = 0
End Sub

Private Sub Form_Resize()
    PosicionaGrade SSTab1, Me, pcBot
    'PosicionaGradeWidth grid,  SSTab1
    'grid.Height = SSTab1.Height - grid.Top - 500
    
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If grid.col = 2 Then
            edita_codigo
        ElseIf grid.col = 13 Then
            edita_qtd
        Else
            BuscarProdutoDescricao grid.row
        End If
    End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    Else
        If grid.col = 2 Then
            If grid.row <= 0 Then Exit Sub
        Else
        End If
        With grid
            If .col = 2 Then
                
                edita_codigo
            ElseIf .col = 13 Then
                edita_qtd
            End If
        End With
        End If
    
    
End Sub
Sub edita_codigo()
    grid.col = 2
    txtGrid.Text = grid.Text
    txtGrid.Top = grid.CellTop + grid.Top
    txtGrid.Left = grid.CellLeft + grid.Left
    txtGrid.Visible = True
    txtGrid.SetFocus
End Sub
Sub edita_qtd()
    grid.col = 13
    txtQtd.Text = grid.Text
    txtQtd.width = grid.CellWidth
    txtQtd.Height = grid.CellHeight
    txtQtd.Top = grid.CellTop + grid.Top
    txtQtd.Left = grid.CellLeft + grid.Left
    txtQtd.Visible = True
    txtQtd.SetFocus
End Sub
Sub BuscarProdutoDescricao(Linha As Integer)
    Dim descricao As String
    Dim p As Produto
    descricao = X.p(Linha - 1).xProd
    
    descricao = Replace(descricao, " ", "%")
    
    descricao = InputBox("Buscar pela descrição", "Descrição", descricao)
    p = F_BuscaProdSaida(descricao)
    If p.cod_pro <> "" Then
        grid.row = Linha
        grid.col = 2
        grid.Text = p.cod_pro
        grid.col = 3
        grid.Text = p.nome_pro
    End If
    
    
    
End Sub
Sub BuscarProdCodBarra(Linha As Integer)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If X.p(Linha - 1).cEAN = "" Then
        Exit Sub
    End If
    
    strSQL = "select distinct tp.cod_pro,nome_pro from " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " left join sped_nf_itens as ti on tp.cod_pro=ti.cod_Pro and ti.cean='" & X.p(Linha - 1).cEAN & "' "
    
    strSQL = strSQL & " where tp.codbarra_pro='" & X.p(Linha - 1).cEAN & "' "
    strSQL = strSQL & " or tp.ean_pro='" & X.p(Linha - 1).cEAN & "' "
    strSQL = strSQL & " or tp.ean_pro='" & X.p(Linha - 1).cEANTrib & "' "
    strSQL = strSQL & " or tp.codbarra_pro='" & X.p(Linha - 1).cEANTrib & "' "
    strSQL = strSQL & " or tp.codfab_pro='" & X.p(Linha - 1).cEAN & "' "
    strSQL = strSQL & " or tp.codfab_pro='" & X.p(Linha - 1).cEANTrib & "' "
    strSQL = strSQL & " or tp.codfab_pro='" & X.p(Linha - 1).cProd & "' "
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            grid.row = Linha
            grid.col = 2
            grid.Text = !cod_pro
            grid.col = 3
            grid.Text = !nome_pro
            ColoreLinhaGrid grid, grid.row, "", False, vbBlue
        Else
            grid.row = Linha
            grid.col = 2
            grid.Text = ""
            grid.col = 3
            grid.Text = ""
            ColoreLinhaGrid grid, grid.row, "", False, vbBlack
        End If
        .Close
    End With
    
    
End Sub
Sub BuscarProdCodFab(Linha As Integer)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If X.p(Linha - 1).cProd = "" Then
        Exit Sub
    End If
    
    strSQL = "select distinct tp.cod_pro,nome_pro from " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " where tp.codfab_pro='" & X.p(Linha - 1).cProd & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            grid.row = Linha
            grid.col = 2
            grid.Text = !cod_pro
            grid.col = 3
            grid.Text = !nome_pro
            ColoreLinhaGrid grid, grid.row, "", False, "vbBlue"
        Else
            grid.row = Linha
            grid.col = 2
            grid.Text = ""
            grid.col = 3
            grid.Text = ""
            ColoreLinhaGrid grid, grid.row, "", False, "vbBlack"
        End If
        .Close
    End With
    
    
End Sub

Private Sub mnBuscaProdDesc_Click()
    If grid.row <= 0 Then
        MsgBox "Escolha uma linha com produto", vbExclamation
        Exit Sub
    End If
    grid.col = 2
    
     BuscarProdutoDescricao grid.row
    
    
End Sub

Private Sub mnBuscarProdCodBarra_Click()
    Dim cod As String
    grid.col = 1
    If grid.Text = "" Then Exit Sub
    
    cod = grid.Text
    grid.col = 2
    If cod = "" Then
       BuscarProdutoDescricao grid.row
    Else
        BuscarProdCodBarra grid.row
    End If
End Sub

Private Sub mnCodFab_Click()
     Dim cod As String
    grid.col = 5
    If grid.Text = "" Then Exit Sub
    
    cod = grid.Text
    grid.col = 2
    
    BuscarProdCodFab grid.row
    
End Sub

Private Sub mnCop_Click()
    If grid.row > 0 Then
        Clipboard.SetText (grid.Text)
    End If
End Sub

Private Sub mnEdit_Click()
    If grid.row > 0 Then
        grid.col = 2
        If grid.Text <> "" Then
            frmCadPro.tag = grid.Text
            frmCadPro.Show 1
        End If
    End If
End Sub

Private Sub mnLimpVinc_Click()
    Dim i As Integer
    P_Status "Aguarde..."
    With grid
        For i = 1 To .rows - 1
            .row = i
            .col = 2
            .Text = ""
        Next
    End With
    P_Status ""
End Sub

Private Sub mnNaoEEsse_Click()
    Dim i As Integer
    P_Status "Aguarde..."
    With grid
        
            
            .col = 2
            .Text = ""
        
    End With
    P_Status ""
End Sub

Private Sub mnNovoProd_Click()
    Dim i As Integer
    i = grid.row - 1
    If grid.row <= 0 Then
        frmCadPro.Show 1
    Else
        frmCadPro.tag = ""
        frmCadPro.txtBarra.Text = X.p(i).cEAN
        frmCadPro.txtEan.Text = X.p(i).cEAN
        frmCadPro.txtNCM.Text = X.p(i).ncm
        frmCadPro.txtbxCodFab.Text = X.p(i).cProd
        frmCadPro.txtbxCusto.Text = X.p(i).vUnTrib
        frmCadPro.txtbxNome_pro.Text = X.p(i).xProd
        frmCadPro.txtbxUnidade.Text = X.p(i).uTrib
        frmCadPro.txtNCM.Text = X.p(i).ncm
        frmCadPro.txtbxCusto.Text = FormatValNFe(X.p(i).vUnCom)
        'P_LocalizaOperacao x.p(i).icms_.cSt, frmCadPro.cmbxNFECst
        frmCadPro.Show 1
        
    End If
    
    
    
End Sub

Private Sub mnVincularProd_Click()
    Dim i As Integer
    P_Status "Aguarde..."
    grid.Visible = False
    With grid
        For i = 1 To .rows - 1
            
            BuscarProdCodBarra i
        Next
    End With
    P_Status ""
    
    calcCusto
    grid.Visible = True
End Sub

Private Sub txtFrete_GotFocus()
    P_SelecTexto txtFrete
End Sub

Private Sub txtFrete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtFrete_LostFocus()
    P_ValorLostFocus txtFrete
    calcCusto
End Sub

Private Sub txtGrid_GotFocus()
    ganhaFoco txtGrid
End Sub

Private Sub txtGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub

Private Sub txtGrid_LostFocus()
    
    txtGrid.Visible = False
End Sub

Private Sub txtQtd_GotFocus()
    ganhaFoco txtQtd
End Sub

Private Sub txtQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        PerdeFoco txtQtd
        grid.col = 13
        grid.Text = txtQtd.Text
        txtQtd.Visible = False
        On Error Resume Next
        grid.row = grid.row + 1
        grid.SetFocus
    End If
    
End Sub

Private Sub txtQtd_LostFocus()
    txtQtd.Visible = False
End Sub
