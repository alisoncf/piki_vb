VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadPro 
   Caption         =   "Cadastro de Produtos:"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13515
   Icon            =   "frmCadPro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   13515
   Begin TabDlg.SSTab sstb 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Dados Básicos"
      TabPicture(0)   =   "frmCadPro.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(1)=   "Picture8"
      Tab(0).Control(2)=   "Picture9"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Outros Dados"
      TabPicture(1)   =   "frmCadPro.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture11"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Similares"
      TabPicture(2)   =   "frmCadPro.frx":0F02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture10"
      Tab(2).Control(1)=   "Picture7"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "SPED/NFe"
      TabPicture(3)   =   "frmCadPro.frx":0F1E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "pcCST"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.PictureBox pcCST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   120
         ScaleHeight     =   5145
         ScaleWidth      =   7545
         TabIndex        =   85
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtCEST 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            TabIndex        =   111
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtCst_Icms 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   100
            Tag             =   "102"
            Text            =   "102 - Tributada pelo Simples Nacional sem permissão de crédito"
            Top             =   1080
            Width           =   6735
         End
         Begin VB.PictureBox pcCstIcms 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6840
            Picture         =   "frmCadPro.frx":0F3A
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   99
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox txtCst_Ipi 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   98
            Top             =   3480
            Width           =   6735
         End
         Begin VB.PictureBox pcCstIpi 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6840
            Picture         =   "frmCadPro.frx":12E5
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   97
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txtCst_Pis 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   96
            Tag             =   "49"
            Text            =   "49 - Outras Operações de Saída"
            Top             =   2280
            Width           =   6735
         End
         Begin VB.PictureBox pcCstPis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6840
            Picture         =   "frmCadPro.frx":1690
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   95
            Top             =   2280
            Width           =   255
         End
         Begin VB.TextBox txtCst_Cofins 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   94
            Tag             =   "49"
            Text            =   "49 - Outras Operações de Saída"
            Top             =   2880
            Width           =   6735
         End
         Begin VB.PictureBox pcCstCofins 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6840
            Picture         =   "frmCadPro.frx":1A3B
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   93
            Top             =   2880
            Width           =   255
         End
         Begin VB.TextBox txtbxICMS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   92
            Tag             =   "102"
            Text            =   "102"
            Top             =   1680
            Width           =   615
         End
         Begin VB.PictureBox pcGenero 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6840
            Picture         =   "frmCadPro.frx":1DE6
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   91
            Top             =   4080
            Width           =   255
         End
         Begin VB.TextBox txtGenero 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   90
            Top             =   4080
            Width           =   6735
         End
         Begin VB.TextBox txtOrigem 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   89
            Top             =   4680
            Width           =   6735
         End
         Begin VB.PictureBox pcOrigem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   6840
            Picture         =   "frmCadPro.frx":2191
            ScaleHeight     =   255
            ScaleWidth      =   225
            TabIndex        =   88
            Top             =   4680
            Width           =   255
         End
         Begin VB.TextBox txtNCMCopia 
            Appearance      =   0  'Flat
            BackColor       =   &H0073E1F4&
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
            Left            =   120
            TabIndex        =   87
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtEanCopia 
            Appearance      =   0  'Flat
            BackColor       =   &H0073E1F4&
            Height          =   285
            Left            =   1440
            TabIndex        =   86
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&CEST"
            Height          =   195
            Left            =   6120
            TabIndex        =   112
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst ICMS:"
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst IPI:"
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   3240
            Width           =   510
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst PIS:"
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   2040
            Width           =   570
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cst COFINS:"
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   2640
            Width           =   900
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Códigos de Situação tributária em operações de saída (vendas):"
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
            Left            =   120
            TabIndex        =   106
            Top             =   0
            Width           =   5265
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "&ICMS:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gênero: (Apenas para produtos primários ex.: insumos)"
            Height          =   195
            Left            =   120
            TabIndex        =   104
            Top             =   3840
            Width           =   3840
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Origem:"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   4440
            Width           =   540
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&NCM:"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&EAN/GTIN:"
            Height          =   195
            Left            =   1440
            TabIndex        =   101
            Top             =   240
            Width           =   840
         End
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   -74880
         ScaleHeight     =   1905
         ScaleWidth      =   6825
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   6855
         Begin VB.ListBox lstCodBarra 
            Appearance      =   0  'Flat
            Height          =   810
            ItemData        =   "frmCadPro.frx":253C
            Left            =   120
            List            =   "frmCadPro.frx":253E
            TabIndex        =   78
            Top             =   360
            Width           =   5415
         End
         Begin VB.TextBox txtbxCodBarra 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   77
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmndbtnExclBarcode 
            Height          =   315
            Left            =   3120
            Picture         =   "frmCadPro.frx":2540
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   1440
            Width           =   315
         End
         Begin VB.CommandButton cmndbtnAddBarcode 
            Height          =   315
            Left            =   2760
            Picture         =   "frmCadPro.frx":2ACA
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   1440
            Width           =   315
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adicionar Barcode:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1200
            Width           =   1350
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Códigos de Barra:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   79
            Top             =   0
            Width           =   2100
         End
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   -74880
         ScaleHeight     =   2145
         ScaleWidth      =   6825
         TabIndex        =   72
         Top             =   1800
         Width           =   6855
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos Similares"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   73
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74880
         ScaleHeight     =   825
         ScaleWidth      =   6825
         TabIndex        =   68
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtbxPrincipal 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   69
            Text            =   "<NENHUM>"
            Top             =   360
            Width           =   5775
         End
         Begin VB.CommandButton cmndbtnPrincipal 
            Height          =   315
            Left            =   5880
            Picture         =   "frmCadPro.frx":3054
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Prod Principal:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   105
            TabIndex        =   71
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74880
         ScaleHeight     =   825
         ScaleWidth      =   7545
         TabIndex        =   64
         Top             =   5640
         Width           =   7575
         Begin VB.TextBox txtbxFabricante 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   65
            Text            =   "<NENHUM>"
            Top             =   360
            Width           =   5775
         End
         Begin VB.CommandButton cmndbtnBuscaFab 
            Height          =   315
            Left            =   5880
            Picture         =   "frmCadPro.frx":314D
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   120
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -74880
         ScaleHeight     =   1305
         ScaleWidth      =   7545
         TabIndex        =   57
         Top             =   4200
         Width           =   7575
         Begin VB.TextBox txtbxSecao 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   5775
         End
         Begin VB.TextBox txtbxSubsec 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEDFD5&
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
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   5775
         End
         Begin VB.CommandButton cmndbtnSec 
            Height          =   315
            Left            =   5880
            Picture         =   "frmCadPro.frx":3246
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnSub 
            Height          =   315
            Left            =   5880
            Picture         =   "frmCadPro.frx":333F
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Seção:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   0
            Width           =   525
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "&Subseção:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3345
         ScaleWidth      =   7545
         TabIndex        =   1
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtBarra 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2880
            TabIndex        =   83
            Text            =   "SEM GTIN"
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox txtbxCodFab 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            TabIndex        =   81
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtEan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Text            =   "SEM GTIN"
            Top             =   2280
            Width           =   2655
         End
         Begin VB.CheckBox chkFracionado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Fracionado"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox txtbxQtdEmb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   840
            TabIndex        =   15
            Text            =   "1"
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtbxPFab 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5280
            TabIndex        =   11
            Text            =   "0,00"
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtbxNome_Pro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   7215
         End
         Begin VB.TextBox txtbxRef 
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtbxLocal 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtbxMargem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   9
            Text            =   "30"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtbxUnidade 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Tag             =   "01"
            Text            =   "UN"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtNCM 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            TabIndex        =   17
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Código de Barras Interno:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2880
            TabIndex        =   84
            Top             =   2040
            Width           =   1800
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Código de Fábrica"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3360
            TabIndex        =   82
            Top             =   1320
            Width           =   1290
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "EAN/GTIN:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   840
         End
         Begin VB.Label lblProvavel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(Provável)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   1440
            TabIndex        =   56
            Top             =   720
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Embalagem"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   14
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Pç. Fab:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nome:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   525
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "&Referência.:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "&Localização:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2880
            TabIndex        =   6
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "M.&Luc(%):"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4440
            TabIndex        =   8
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "&Unidade:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "NCM:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1800
            TabIndex        =   16
            Top             =   1320
            Width           =   405
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ST:"
            Height          =   195
            Left            =   5880
            TabIndex        =   18
            Top             =   2880
            Width           =   255
         End
         Begin VB.Label lblSt 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "T"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   6240
            TabIndex        =   55
            Top             =   2880
            Width           =   120
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   8160
      ScaleHeight     =   2505
      ScaleWidth      =   5265
      TabIndex        =   37
      Top             =   480
      Width           =   5295
      Begin VB.TextBox txtbxCusto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3360
         TabIndex        =   33
         Text            =   "0,00"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtbxQueima 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H007DBFDB&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3360
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtbxPa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   240
         TabIndex        =   29
         Text            =   "0,00"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtbxPv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtbxCr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         TabIndex        =   31
         Text            =   "0,00"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkQueima 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Queima"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo Bruto:"
         Height          =   195
         Left            =   3360
         TabIndex        =   32
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label lblEr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Esto&que:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lblPrecoCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Preço Varejo:"
         Height          =   195
         Left            =   1800
         TabIndex        =   24
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblestCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço &Atacado:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblCusCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&usto:"
         Height          =   195
         Left            =   1800
         TabIndex        =   30
         Top             =   1320
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   13515
      TabIndex        =   43
      Top             =   7365
      Width           =   13515
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   2280
         TabIndex        =   54
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblNumNot 
         AutoSize        =   -1  'True
         Caption         =   "00000000"
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
         Left            =   3000
         TabIndex        =   53
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblEmb 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   3240
         TabIndex        =   52
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embalagem:"
         Height          =   195
         Left            =   2280
         TabIndex        =   51
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   5040
         TabIndex        =   50
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Unitário:"
         Height          =   195
         Left            =   3840
         TabIndex        =   49
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   1320
         TabIndex        =   48
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd Comprada:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblNomeFor 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Fornecedor"
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
         Left            =   3840
         TabIndex        =   46
         Top             =   0
         Width           =   1740
      End
      Begin VB.Label lblUltComp 
         AutoSize        =   -1  'True
         Caption         =   "00/00/0000"
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
         Left            =   1320
         TabIndex        =   45
         Top             =   0
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Última Compra:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   13515
      TabIndex        =   39
      Top             =   0
      Width           =   13515
      Begin VB.Label lblNomePro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Novo Produto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         TabIndex        =   40
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox pcPrecos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   8160
      ScaleHeight     =   4185
      ScaleWidth      =   5265
      TabIndex        =   41
      Top             =   3120
      Width           =   5295
      Begin MSFlexGridLib.MSFlexGrid msflxgrdPrecos 
         Height          =   3975
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7011
         _Version        =   393216
         BackColor       =   -2147483637
         BackColorFixed  =   7594484
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         GridLinesFixed  =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   13515
      TabIndex        =   38
      Top             =   7860
      Width           =   13515
      Begin VB.CommandButton cmdHist 
         Caption         =   "Histórico do Produto"
         Height          =   855
         Left            =   120
         Picture         =   "frmCadPro.frx":3438
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   855
         Left            =   8280
         Picture         =   "frmCadPro.frx":38DA
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Salvar F7"
         Height          =   855
         Left            =   9960
         Picture         =   "frmCadPro.frx":39CE
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar (ESC)"
         Height          =   855
         Left            =   11640
         Picture         =   "frmCadPro.frx":3B2A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCadPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Dim cPreco As String 'para carregar o preço atual no caso de mudança de preço


Private Sub chkProd_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub



Private Sub chkBonif_Click()
    If chkBonif.Value = 1 Then
        txtbxBonif.Enabled = True
        txtbxBonif.SetFocus
    Else
        txtbxBonif.Enabled = False
    End If
    
End Sub

Private Sub chkEstoque_Click()
    If Me.tag <> "" Then
        
    End If
End Sub

Private Sub cmbxControle_Change()

End Sub

Private Sub cmbxControle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub chkFracionado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxSecao.SetFocus
    End If
End Sub

Private Sub chkQueima_Click()
    If chkQueima.Value = 1 Then
        txtbxQueima.Enabled = True
        txtbxQueima.SetFocus
    Else
        txtbxQueima.Enabled = False
        txtbxQueima.text = "0,00"
    End If
End Sub

Private Sub cmbxICMS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxMarca_Click()
    lblMar.caption = F_BuscaUmCampo("Cod_MAR", "tab_marcas", "nome_mar='" & cmbxMarca.text & "'")
End Sub

Private Sub cmbxEsp_Click()
    lblEsp.caption = F_BuscaUmCampo("Cod_Esp", "tab_Especies", "nome_Esp='" & cmbxEsp.text & "'")
End Sub

Private Sub cmbxCat_Click()
    lblCat.caption = F_BuscaUmCampo("Cod_Cat", "tab_Categorias", "nome_Cat='" & cmbxCat.text & "'")
    
End Sub


Private Sub cmbxList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxPrincipio_Click()
    cmbxPrincipio.tag = F_BuscaUmCampo("cod_pri", "tab_principios", "nome_pri='" & cmbxPrincipio.text & "'")
    
End Sub

Private Sub cmbxPrincipio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
    End If
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub



Private Sub cmbxSt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSt_LostFocus()
    On Error Resume Next
    
End Sub


Private Sub cmndbtnCat_Click()
    buscaCat
End Sub


Private Sub cmdHist_Click()
    P_HistProd Me.tag
End Sub

Private Sub cmndbtnAddBarcode_Click()
    If txtbxCodBarra.text = "" Then
        MsgBox "Informe os dígitos do código de barras!", vbExclamation
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    
    If Me.tag = "" Then
        lstCodBarra.AddItem txtbxCodBarra.text
    Else
        If produtos.F_AddCodBar(Me.tag, txtbxCodBarra.text) Then
            P_PreencheBarCode Me.tag, lstCodBarra
        End If
    End If
End Sub

    
    



Private Sub cmndbtnExclBarcode_Click()
    Dim strSQL As String
    If lstCodBarra.ListIndex = -1 Then
        MsgBox "É necessário que se escolha um código de barras!", vbCritical
        lstCodBarra.SetFocus
        Exit Sub
    End If
    If lstCodBarra.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstCodBarra.ListCount = 1 Then
        MsgBox "Este código de barras não pode ser excluído por ser o único!", vbCritical
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o código de barras?", vbYesNo + vbCritical) = vbYes Then
        If Me.tag = "" Then
            lstCodBarra.RemoveItem lstCodBarra.ListIndex
        Else
            strSQL = "delete from tab_codbarras where cod_pro='" & Me.tag & "' "
            strSQL = strSQL & " and cod_bar='" & lstCodBarra.text & "' "
            Conn.Execute strSQL
            buscaCodBar
        End If
    End If
End Sub

Private Sub cmndbtnMarcas_Click()
    buscaMarca
End Sub

Private Sub cmndbtnRefresh_Click()
    RefresheComp
End Sub

Private Sub cmndbtnExclEmb_Click()
    If lstEmb.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstEmb.ListIndex = -1 Then
        MsgBox "É necessário que se escolha m item!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir a embalagem?", vbYesNo + vbCritical) = vbYes Then
        lstEmb.RemoveItem lstEmb.ListIndex
    End If
End Sub


Private Sub cmndbtnFoto_click()
    frmFoto.caption = txtbxNome_Pro.text
    frmFoto.tag = Me.tag
    frmFoto.Show 1
End Sub

Private Sub cmbxSubSec_Click()
    cmbxSubSec.tag = F_BuscaUmCampo("cod_sub", "tab_subsec", "nome_sub='" & cmbxSubSec.text & "' and cod_sec='" & cmbxSecao.tag & "'")
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSubSec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxFabricante.SetFocus
    End If
End Sub



Private Sub cmndbtnBuscaFab_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxFabricante.text)
    If f.cod_For <> "" Then
        txtbxFabricante.text = f.nome_for
        txtbxFabricante.tag = f.cod_For
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscarControl_Click()
    Dim c As controle
    c = produtos.F_BuscaProdControle(txtbxControle.text)
    If c.cod_con <> "" Then
        txtbxControle.text = c.nome_con
        txtbxControle.tag = c.cod_con
    End If
    
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para isto!", vbExclamation
        Exit Sub
    End If
    With dtPro.Recordset
        If MsgBox("Confirma exclusão do lote? ", vbYesNo + vbCritical) = vbYes Then
            strSQL = F_BuscaUmCampo("num_ven", "tab_itvendas as ti, tab_vendas as tv", _
                "tv.num_ven=ti.nuM_ven and ti.seq_pro='" & !seq_pro & "' and situacao_ven<>'3' and cod_pro='" & !cod_pro & "'")
            If strSQL = "" Then
                strSQL = "delete from tab_estoque where cod_pro='" & !cod_pro & "' and seq_pro=" & !seq_pro
                Conn.Execute strSQL
                
            Else
                MsgBox "O produto não pode ser excluído!", vbExclamation
                Exit Sub
            End If
                
        End If
    End With
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnMatriz_Click()
    Dim p As produtos.produto
    If txtbxMatriz.text = "" Then
        MsgBox "Informe o nome do produto!", vbExclamation
        Exit Sub
    End If
    p = F_BuscaProdCabecalho(txtbxMatriz.text)
    If p.cod_pro <> "" Then
        txtbxMatriz.tag = p.cod_pro
        txtbxMatriz.text = p.nome_pro
        lblPVMatriz.caption = FormatVal(p.pv_pro)
        lblCustoMatriz.caption = FormatVal(p.cr_pro)
        'cmbxPrincipio.Text = p.prin.nome_pri
        'cmbxSecao.Text = p.sub.sec.nome_sec
        'cmbxSubSec.Text = p.sub.nome_sub
    End If
    
End Sub


Private Sub cmndbtnNovo_Click()
    Dim fr As New frmCadPro
    fr.tag = ""
    fr.Show 0
    Unload Me
End Sub

Private Sub cmndbtnPro_Click()
    frmProdProm.Show 0
End Sub

Private Sub cmndbtnQueima_Click()
    
End Sub

Private Sub cmndbtnPrincipal_Click()
    Dim p As produto
    If txtbxPrincipal.text = "<NENHUM>" Then
        txtbxPrincipal.text = ""
    End If
    p = produtos.F_BuscaProdSaida(txtbxPrincipal.text)
    If p.cod_pro = "" Then
        txtbxPrincipal.text = "<NENHUM>"
        txtbxPrincipal.tag = ""
        
    Else
        If p.cod_pro = Me.tag Then
            MsgBox "Um produto não pode ser apontado como similar dele mesmo!", vbExclamation
            Exit Sub
        End If
        txtbxPrincipal.text = p.nome_pro
        txtbxPrincipal.tag = p.cod_pro
        buscaSim
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
On Error GoTo trata_Erro

    If txtbxICMS.tag = "" Then
        MsgBox "Informe o ICMS!", vbExclamation
        Exit Sub
    End If
    If Trim(txtbxNome_Pro.text) = "" Then
        MsgBox "Informe um nome para o produto!", vbExclamation
        Exit Sub
    End If
    If txtbxSubsec.tag = "" Then
        MsgBox "Informe a subseção!", vbExclamation
        txtbxSecao.SetFocus
        Exit Sub
    End If
    If txtbxSecao.tag = "" Then
        MsgBox "Informe a seção!", vbExclamation
        txtbxSecao.SetFocus
        Exit Sub
    End If
    If strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso.", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        strSQL = F_BuscaUmCampo("nome_pro", "tab_produtos", "NOME_PRO='" & _
            txtbxNome_Pro.text & "'")
        If strSQL <> "" Then
            MsgBox "Já existe um produto com este nome!", vbCritical
            txtbxNome_Pro.SetFocus
            Exit Sub
        End If
        If txtbxRef.text <> "" Then
        
            strSQL = F_BuscaUmCampo("nome_pro", "tab_produtos", "codfab_pro='" & _
            txtbxRef.text & "'")
            
        End If
            If strSQL <> "" Then
                If MsgBox("Já existe um produto com esta referência!" & _
                     Chr(13) & strSQL & Chr(13) & "Deseja continuar?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
                End If
            End If
            
        
        If MsgBox("confirma inclusão do Produto?", vbYesNo + vbQuestion) = vbYes Then
            
            Me.tag = produtos.F_GravaProduto(txtbxNome_Pro.text, txtbxSubsec.tag, _
                  txtbxRef.text, txtbxCr.text, txtbxCr.text, txtbxPv.text, txtbxPa.text, txtbxUnidade.text, _
                    txtbxPFab.text, txtbxFabricante.tag, txtbxLocal.text, lblSt.caption, txtbxICMS.tag, _
                     txtbxMargem.text, txtbxCusto.text, txtbxQueima.text, _
                     txtbxPrincipal.tag, IIf(chkFracionado.Value = 1, "S", "N"), txtNCM.text, txtEan.text, txtBarra.text)
            
            If Me.tag <> "" Then
                MsgBox "Produto Gravado com o código " & Me.tag, vbInformation
                
                For iCount = 0 To lstCodBarra.ListCount - 1
                    lstCodBarra.ListIndex = iCount
                    If produtos.F_AddCodBar(Me.tag, lstCodBarra.text) Then
                    End If
                Next
            End If
            APIMergeProduto Me.tag
            
        
        End If
    Else
        strSQL = F_BuscaUmCampo("nome_pro", "tab_produtos", "NOME_PRO='" & _
            txtbxNome_Pro.text & "' and cod_pro<>'" & Me.tag & "'")
        If strSQL <> "" Then
            MsgBox "Já existe um produto com este nome!", vbCritical
            txtbxNome_Pro.SetFocus
            Exit Sub
        End If
        
        
        
        If MsgBox("Confirma alteração do Produto?", vbYesNo + vbQuestion) = vbYes Then
            If produtos.F_EditaProduto(Me.tag, txtbxNome_Pro.text, txtbxSubsec.tag, txtbxRef.text, _
                txtbxCr.text, txtbxPv.text, txtbxPa.text, txtbxUnidade.text, _
                txtbxPFab.text, txtbxFabricante.tag, txtbxLocal.text, lblSt.caption, txtbxICMS.tag, txtbxCusto.text, _
                    txtbxQueima.text, IIf(chkFracionado.Value = 1, "S", "N"), txtbxPrincipal.tag, txtNCM.text, txtGenero.tag, _
                    txtOrigem.tag, txtEan.text, txtCst_Icms.tag, txtCst_Ipi.tag, _
                    txtCst_Pis.tag, txtCst_Cofins.tag, txtCEST.text, txtEan.text, txtBarra.text) Then
                MsgBox "Produto alterado com sucesso!", vbInformation
                P_AltPreco Me.tag, CCur(txtbxPv.text), CCur(txtbxPv.tag)
                
                APIMergeProduto Me.tag, True
            End If
        End If
    End If
        P_Status ""
        Exit Sub
trata_Erro:
        P_MostraErro "Cadastro de Produto"
End Sub
Sub P_Altera()
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "update tab_produtos set "
    strSQL = strSQL & "nome_pro='" & txtbxNome_Pro.text & "',"
    strSQL = strSQL & "unidade_pro='" & txtbxUnidade.text & "',"
    strSQL = strSQL & "codfab_pro='" & txtbxRef.text & "', "
    strSQL = strSQL & "balanca_pro='" & IIf(chkProd.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "cod_mar='" & lblMar.caption & "', "
    strSQL = strSQL & "cod_cat='" & lblCat.caption & "', "
    
    strSQL = strSQL & "endereco_pro='" & txtbxLocal.text & "',"
    strSQL = strSQL & "lucro_pro=" & F_GravaDec(txtbxMargem.text) & ",  "
    strSQL = strSQL & "pv_pro=" & F_GravaDec(txtbxPreco.text) & ","
    strSQL = strSQL & "cr_pro=" & F_GravaDec(txtbxCusto.text) & ","
   ' If strAltpreco_usu = "S" Then
  '      strSQL = strSQL & "er_pro=" & F_GravaDec(txtbxEstoque.Text) & ","
   ' End If
    strSQL = strSQL & "estmin_pro=" & F_GravaDec(txtbxEstMin.text) & ","
    strSQL = strSQL & "st_pro='" & Left(cmbxSt.text, 1) & "', "  ' situacao trib"
    strSQL = strSQL & "cod_imp='" & txtbxICMS.tag & "'  "
    strSQL = strSQL & " where cod_pro='" & Me.tag & "' "
    Conn.Execute strSQL
    P_AltPreco Me.tag, CCur(txtbxPreco.text), CCur(cPreco)
'    strSQL = "delete from tab_codbarras where cod_pro='" & Me.tag & "' "
'    Conn.Execute strSQL
'    With lstCodBarra
'        For iCount = 0 To .ListCount - 1
'            .ListIndex = iCount
'            strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values("
'            strSQL = strSQL & "'" & Me.tag & "',"
'            strSQL = strSQL & "'" & .Text & "')"
'            Conn.Execute strSQL
'        Next
'    End With
'
    
   ' strSQL = F_GeraLog(strSQL, "tab_produtos")
    MsgBox "Produto Alterado com Sucesso!", vbInformation
    Unload Me
    On Error Resume Next
    If IsLoad("frmEdiPro") Then
        'RefreshDataPosition frmEdiPro.adc
    End If
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub gravaEmb()
    Dim strSQL As String
    strSQL = "delete from tab_embalagens where cod_pro='" & Me.tag & "' "
    Conn.Execute strSQL
    With lstEmb
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & Me.tag & "',"
            strSQL = strSQL & .text & ")"
            Conn.Execute strSQL
        Next
    End With
End Sub

Private Sub cmnxdbtnBuscaFor_Click()
    frmSelectCli.txtbxNome.tag = "F"
    frmSelectCli.txtbxNome.text = txtbxFornecedor.text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxFornecedor.text = frmSelectCli.adoR!nome
    txtbxFornecedor.tag = frmSelectCli.adoR!cod
    Unload frmSelectCli
End Sub


Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select num_not,entrada_not,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " tab_fornecedores as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & Me.tag & "' "
    strSQL = strSQL & " order by entrada_not desc limit 1 "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            lblUltComp.caption = ""
            lblNomeFor.caption = ""
            lblNumNot.caption = ""
            lblQtd.caption = ""
            lblUnit.caption = ""
            lblEmb.caption = ""
        Else
            lblUltComp.caption = !entrada_not
            lblNomeFor.caption = !nome_for
            lblNumNot.caption = !Num_not
            lblQtd.caption = !qtd_ite
            lblUnit.caption = FormatVal(!unit_ite)
            lblEmb.caption = !qtd_emb
        End If
        .Close
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnSusp_Click()
    If Me.tag <> "" Then
        
    End If
End Sub



Private Sub cmndbtnSec_Click()
    Dim s As secao.secao
    s = F_BuscaSec(txtbxSecao.text)
    If s.cod_sec <> "" Then
        txtbxSecao.text = s.nome_sec
        txtbxSecao.tag = s.cod_sec
        txtbxSubsec.SetFocus
    End If
End Sub

Private Sub cmndbtnSub_Click()
    Dim s As subsecAO.subsec
   ' If txtbxSecao.tag = "" Then
    '    MsgBox "Informe a seção", vbExclamation
    '    txtbxSecao.SetFocus
    '    Exit Sub
    'End If
    
    s = F_BuscaSubSec(txtbxSecao.tag, txtbxSubsec.text)
    
    
    
    If s.cod_sub <> "" Then
        txtbxSubsec.text = s.nome_sub
        txtbxSubsec.tag = s.cod_sub
        txtbxSecao.text = s.sec.nome_sec
        txtbxSecao.tag = s.sec.cod_sec
        txtbxFabricante.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    
    If bActive Then Exit Sub
    bActive = True
   ' cPreco = txtbxPreco.Text
    buscaCodBar
    If strAltpreco_usu = "S" Then
        lblPrecoCap.Enabled = True
        lblCusCap.Enabled = True
      '  txtbxPreco.Enabled = True
'        txtbxCusto.Enabled = True
        
    Else
        lblPrecoCap.Enabled = False
        lblCusCap.Enabled = False
        'txtbxPreco.Enabled = False
        'txtbxCusto.Enabled = False
        'txtbxCustoFiscal.Enabled = False
    End If
    If Me.tag <> "" Then
        cmdHist.Visible = True
        carrega
        buscaPrecos
        buscaSim
        BuscaUltEntrada
        sstb.Tab = 0
        cmndbtnNovo.Visible = True
        pcPrecos.Visible = True
    Else
        cmdHist.Visible = False
        cmndbtnNovo.Visible = True
        pcPrecos.Visible = True
    End If
    sstb.TabVisible(1) = False
    sstb.TabVisible(2) = False
    chkQueima.caption = strNomedaPromocao
    
End Sub

Sub buscaSim()
    Dim strSQL As String
   ' On Error Resume Next

        strSQL = "select tp.* from tab_produtos as tp "
        strSQL = strSQL & " where "
        
        strSQL = strSQL & " ( cod_pro='" & txtbxPrincipal.tag & "' and cod_pro<>'')"
        strSQL = strSQL & " or (principal_pro ='" & Me.tag & "' )"
        
'        dt.RecordSource = strSQL
'        dt.Refresh
    
    
        
    
    
End Sub


Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = produtos.F_SelectProduto("cod_pro='" & Me.tag & "'")
    abreObjetoAdo rPesq, strSQL
    With rPesq
    
        Me.caption = !cod_pro & " - " & !nome_pro
        lblNomePro.caption = !nome_pro
        txtbxPrincipal.tag = !principal_pro
        If !principal_pro = "" Then
            txtbxPrincipal.text = "<NENHUM>"
        Else
            txtbxPrincipal.text = F_BuscaUmCampo("nome_pro", "tab_produtos", "cod_pro='" & !principal_pro & "'")
        End If
        
        txtbxFabricante.tag = !cod_fab
        txtbxFabricante.text = F_BuscaUmCampo("nome_for", "tab_fornecedores", "cod_for='" & !cod_fab & "'")
        txtbxSecao.tag = !cod_sec
        txtbxSecao.text = !nome_sec
        If Not IsNull(!ncm_pro) Then
            txtNCM.text = !ncm_pro
        End If
        txtbxSubsec.text = !nome_sub
        txtbxSubsec.tag = !cod_sub
        If !fracionado_pro = "S" Then
            chkFracionado.Value = 1
        Else
            chkFracionado.Value = 0
        End If
        
        txtbxPv.text = FormatVal(!pv_pro)
        txtbxPa.text = FormatVal(!pa_pro)
        txtbxPa.tag = FormatVal(!pa_pro)
        txtbxPv.tag = FormatVal(!pv_pro)
        txtbxRef.text = !ref_pro
        txtbxCr.text = FormatVal(!cr_pro)
        lblEr.caption = !er_pro
        txtbxQueima.text = FormatVal(!pq_pro)
        txtbxICMS.text = !icms_pro
        
        If strAltpreco_usu <> "S" Then
            txtbxPv.Enabled = False
            txtbxCusto.Enabled = False
            chkQueima.Enabled = False
            txtbxPa.Enabled = False
        End If
        txtEan.text = TiraNull(!ean_pro)
        txtBarra.text = TiraNull(!codbarra_pro)
        
        If !pq_pro > 0 Then
            chkQueima.Value = 1
            
        Else
        
        End If
        txtbxCusto.text = FormatVal(!cb_pro)
        txtbxNome_Pro.text = !nome_pro
        txtbxUnidade.text = !unidade_pro
        txtbxMargem.text = FormatVal(!lucro_pro)
        
        txtbxLocal.text = !endereco_pro
        
        lblSt.caption = !st_pro
        
        'P_LocalizaOperacao !st_pro, cmbxSt
        
        
        
        txtbxICMS.text = !desc_imp
        txtbxICMS.tag = !cod_imp
        txtCst_Icms.text = Trim(!cst_icms & " " & Format(!icms))
        txtCst_Icms.tag = !cst_icms
        txtCst_Ipi.text = Format(!ipi)
        txtCst_Ipi.tag = !CST_IPI
        txtCst_Pis.text = Trim(!cst_pis & " " & Format(!pis))
        txtCst_Pis.tag = !cst_pis
        txtCst_Cofins.text = Trim(!cst_cofins & " " & Format(!cofins))
        txtCst_Cofins.tag = !cst_cofins
        
        
        .Close
    End With
End Sub
Sub buscaPrecos()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim desc As Currency
    msflxgrdPrecos.rows = 1
    strSQL = "select distinct * from tab_formaspag where parcelas_for>0 and  visivel_for = 'S' order by parcelas_for,avista_for desc, tipo_cai"
    abreObjetoAdo rPesq, strSQL
    msflxgrdPrecos.ColAlignment(0) = 3
    msflxgrdPrecos.ColWidth(0) = 2000
    msflxgrdPrecos.ColAlignment(1) = 3
    
    With rPesq
        Do While Not .EOF
            desc = CCur(txtbxPv.text) * !desconto_for / 100
            strSQL = !nome_for & Chr(9) & FormatVal((CCur(txtbxPv.text) - desc) / !parcelas_for)
            msflxgrdPrecos.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msflxgrdPrecos, 1, &HB4DFE0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
    
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    End If
    
End Sub
Sub buscaCat()
    frmEdiCat.Show 0
    frmEdiCat.SetFocus
End Sub
Sub buscaMarca()
     frmEdiMar.Show 0
        frmEdiMar.SetFocus
End Sub
Sub RefresheComp()
    
End Sub
Private Sub Form_Load()
'    Me.Height = 5280
'    Me.Width = 4935
    bActive = False
'    resizeForm Me
    Set ado = New ADODB.Recordset
    With msflxgrdPrecos
        .FixedCols = 0
        .rows = 1
        .Cols = 2
        .col = 0: .text = "Plano"
        .col = 1: .text = "Preço"
        
        .ColWidth(0) = 1800
    End With
    'P_PreencheMarca cmbxMarca
    'P_PreencheCateg cmbxCat
    
    
    
            'cmbxSt.ListIndex = 0
    'cmbxSt_LostFocus
    
    resizeForm Me
    sstb.Tab = 0
    
    
End Sub
Sub buscaCodBar()
    P_PreencheBarCode Me.tag, lstCodBarra
End Sub
Private Sub Form_Resize()
'     P_CentraObejct frControle, Me
End Sub

Private Sub frControle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub lblInfoEnt_Click()
    pcEnt_Click
End Sub

Private Sub lblInfoEnt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = True
    lblInfoEnt.FontUnderline = True
    pcEnt.BorderStyle = 1
End Sub

Private Sub lblInfoSai_Click()
    pcSai_Click
End Sub

Private Sub lblInfoSai_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoSai.FontBold = True
    lblInfoSai.FontUnderline = True
    pcSai.BorderStyle = 1
End Sub

Private Sub lstCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub lstEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub pcCstCofins_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Cofins.text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Cofins(expr)
    If p.codigo <> "" Then
        txtCst_Cofins.text = p.codigo & " - " & p.nome
        txtCst_Cofins.tag = p.codigo
        
    End If
End Sub

Private Sub pcCstIcms_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Icms.text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Icms(expr)
    If p.codigo <> "" Then
        txtCst_Icms.text = p.codigo & " - " & p.nome
        txtCst_Icms.tag = p.codigo
        txtCst_Ipi.SetFocus
    End If
    
    
End Sub

Private Sub pcCstIpi_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Ipi.text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_IPI(expr)
    If p.codigo <> "" Then
        txtCst_Ipi.text = p.codigo & " - " & p.nome
        txtCst_Ipi.tag = p.codigo
        txtCst_Pis.SetFocus
    End If
End Sub

Private Sub pcCstPis_Click()
    Dim p As Padrao1
    Dim expr As String
    
    expr = txtCst_Pis.text
    If F_SoNumero(expr) <> "" Then
        expr = F_SoNumero(expr)
           
    End If
    
    p = Sped_Lista_Cst_Pis(expr)
    If p.codigo <> "" Then
        txtCst_Pis.text = p.codigo & " - " & p.nome
        txtCst_Pis.tag = p.codigo
        txtCst_Cofins.SetFocus
    End If
End Sub

Private Sub pcEnt_Click()
    If Me.tag <> "" Then
        P_InfoEntradasPro Me.tag
    Else
        MsgBox "Não há produto gravado!", vbCritical
    End If
End Sub

Private Sub pcEnt_GotFocus()
    lblInfoEnt_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcEnt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcEnt_Click
    End If
End Sub

Private Sub pcEnt_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcGenero_Click()
    buscaGenero
End Sub

Private Sub pcInfo_Click()
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub pcInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub


Private Sub pcOrigem_Click()
    buscaOrigem
End Sub
Sub buscaGenero()
    Dim p As Padrao1
    p = Sped_Genero(txtGenero.text)
    If p.codigo <> "" Then
        txtGenero.text = p.nome
        txtGenero.tag = p.codigo
        txtOrigem.SetFocus
    End If
End Sub
Sub buscaOrigem()
    Dim p As Padrao1
    p = Sped_Origem(txtOrigem.text)
    If p.codigo <> "" Then
        txtOrigem.text = p.nome
        txtOrigem.tag = p.codigo
        txtbxUnidade.SetFocus
    End If
End Sub

Private Sub pcUnidade_Click()
    buscaUnidade
End Sub
Sub buscaUnidade()
    Dim p As Padrao1
    p = SelectUnidade(txtbxUnidade.text)
    If p.codigo <> "" Then
        txtbxUnidade.text = p.codigo
        txtbxLocal.SetFocus
    End If
    
End Sub


Private Sub pcSai_Click()
    If Me.tag = "" Then Exit Sub
    P_InfoSaidasPro Me.tag
End Sub

Private Sub pcSai_GotFocus()
    lblInfoSai_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcSai_Click
    End If
End Sub

Private Sub pcSai_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub


Private Sub txtbxPrecoMin_GotFocus()
    
End Sub


Private Sub Picture9_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub sstb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
    
    
End Sub

Private Sub txtbxCusto_GotFocus()
    P_SelecTexto txtbxCusto
End Sub

Private Sub txtbxCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCusto_LostFocus()
    P_AjustaTexto txtbxCusto
End Sub

Private Sub txtbxCr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCr_LostFocus()
    P_ValorLostFocus txtbxCr
End Sub

Private Sub txtbxFabricante_GotFocus()
    If txtbxFabricante.text = "<NENHUM>" Then
        txtbxFabricante.text = ""
    End If
End Sub

Private Sub txtbxFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaFab_Click
    End If
End Sub


Private Sub txtbxFabricante_LostFocus()
    If txtbxFabricante.tag = "" Then
        txtbxFabricante.text = "<NENHUM>"
        
    End If
End Sub



Private Sub txtbxIcms_GotFocus()
    P_SelecTexto txtbxICMS
End Sub

Private Sub txtbxIcms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxPrincipal_GotFocus()
    If txtbxPrincipal.text = "<NENHUM>" Then
        txtbxPrincipal.text = ""
    End If
End Sub

Private Sub txtbxPrincipal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnPrincipal_Click
    End If
End Sub

Private Sub txtbxPrincipal_LostFocus()
    
    If txtbxPrincipal.tag = "" Then
        txtbxPrincipal.text = "<NENHUM>"
        
    End If
    
End Sub

Private Sub txtbxPv_Change()
    txtbxPa.text = txtbxPv.text
End Sub

Private Sub txtbxPv_GotFocus()
    P_SelecTexto txtbxPv
End Sub

Private Sub txtbxPv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPv_LostFocus()
    P_ValorLostFocus txtbxPv
End Sub


Private Sub txtbxQtdEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub txtbxQueima_GotFocus()
    P_SelecTexto txtbxQueima
End Sub

Private Sub txtbxQueima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQueima_LostFocus()
    P_ValorLostFocus txtbxQueima
End Sub

Private Sub txtbxRef_GotFocus()
    If txtbxRef.text = "" And Me.tag = "" Then
        txtbxRef.text = MaxKeySQL("tab_produtos", "cod_pro", "", "0000000#")
        lblProvavel.Visible = True
    End If
    P_SelecTexto txtbxRef
End Sub

Private Sub txtbxRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxRef_LostFocus()
    P_AjustaTexto txtbxRef
End Sub

Private Sub txtbxDescricao_GotFocus()
    'P_SelecTexto txtbxDescricao
End Sub

Private Sub txtbxDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescricao_LostFocus()
    'P_AjustaTexto txtbxDescricao
End Sub







Private Sub txtbxFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmnxdbtnBuscaFor_Click
    End If
End Sub

Private Sub txtbxLocal_GotFocus()
    P_SelecTexto txtbxLocal
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxLocal_LostFocus()
    P_AjustaTexto txtbxLocal
End Sub


Private Sub txtbxMargem_GotFocus()
    P_SelecTexto txtbxMargem
End Sub

Private Sub txtbxMargem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMargem_LostFocus()
    P_ValorLostFocus txtbxMargem
End Sub

Private Sub txtbxNome_Pro_GotFocus()
    P_SelecTexto txtbxNome_Pro
End Sub

Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Pro_LostFocus()
    P_AjustaTexto txtbxNome_Pro
End Sub


Private Sub txtbxSecao_GotFocus()
    P_SelecTexto txtbxSecao
End Sub

Private Sub txtbxSecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSec_Click
    End If
End Sub

Private Sub txtbxSubsec_GotFocus()
    P_SelecTexto txtbxSubsec
End Sub

Private Sub txtbxSubsec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSub_Click
    End If
End Sub

Private Sub txtbxUnidade_GotFocus()
    P_SelecTexto txtbxUnidade
End Sub

Private Sub txtbxUnidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxUnidade_LostFocus()
    P_AjustaTexto txtbxUnidade
End Sub

Private Sub txtEan_Change()
    txtEanCopia.text = txtEan.text
End Sub

Private Sub txtEanCopia_Change()
    txtEan.text = txtEanCopia.text
End Sub

Private Sub txtNCM_Change()
    txtNCMCopia.text = txtNCM.text
End Sub

Private Sub txtNCM_GotFocus()
    P_SelecTexto txtNCM
End Sub

Private Sub txtNCM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNCM_LostFocus()
    P_AjustaTexto txtNCM
End Sub

Private Sub txtNCMCopia_Change()
    txtNCM.text = txtNCMCopia.text
End Sub
