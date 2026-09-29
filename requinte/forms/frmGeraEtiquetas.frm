VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Begin VB.Form frmGeraEtiquetas 
   Caption         =   "Geração de Etiquetas"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   17115
   Icon            =   "frmGeraEtiquetas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   17115
   Begin VB.PictureBox pcEtiquetas 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6585
      ScaleWidth      =   6585
      TabIndex        =   30
      Top             =   2055
      Width           =   6615
      Begin VB.PictureBox pcBarra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   1905
         TabIndex        =   77
         Top             =   4080
         Width           =   1935
         Begin BARCODEXLib.BarcodeX brcdx 
            Height          =   735
            Left            =   120
            Top             =   120
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   1296
            _StockProps     =   13
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "0000000"
            Title           =   ""
            Scaling         =   -1
            BarcodeType     =   8
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   735
         Left            =   5160
         TabIndex        =   76
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   3360
         TabIndex        =   75
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opções da Grade:"
         Height          =   855
         Left            =   3360
         TabIndex        =   69
         Top             =   3120
         Width           =   3135
         Begin VB.CommandButton cmndbtnEdit 
            Height          =   375
            Left            =   1080
            Picture         =   "frmGeraEtiquetas.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmndbtnExcl 
            Height          =   375
            Left            =   600
            Picture         =   "frmGeraEtiquetas.frx":0294
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmndbtnLimpa 
            Height          =   375
            Left            =   120
            Picture         =   "frmGeraEtiquetas.frx":03DE
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmndbtnSalvarComo 
            Height          =   375
            Left            =   1560
            Picture         =   "frmGeraEtiquetas.frx":0528
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Salvar Modelo Como..."
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmndbtnAbrir 
            Height          =   375
            Left            =   2040
            Picture         =   "frmGeraEtiquetas.frx":0672
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Abrir um modelo pre-definido"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame frameOpt 
         Caption         =   "Imprimir"
         Height          =   855
         Left            =   120
         TabIndex        =   64
         Top             =   3120
         Width           =   3135
         Begin VB.CheckBox chkVista 
            Caption         =   "Preço A Vista"
            Height          =   195
            Left            =   1320
            TabIndex        =   68
            Top             =   480
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkPrazo 
            Caption         =   "Preço A Prazo"
            Height          =   195
            Left            =   1320
            TabIndex        =   67
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkBarra 
            Caption         =   "Barra"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   480
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkRecorte 
            Caption         =   "Recorte"
            Height          =   195
            Left            =   240
            TabIndex        =   65
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtbxIni 
         Appearance      =   0  'Flat
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
         Left            =   840
         TabIndex        =   61
         Text            =   "1"
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txtbxCopias 
         Appearance      =   0  'Flat
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
         TabIndex        =   60
         Text            =   "1"
         Top             =   2580
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H007DBFDB&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   8265
         TabIndex        =   47
         Top             =   0
         Width           =   8295
         Begin VB.Label lblDescEti 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Etiqueta"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.TextBox txtbxLarg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5160
         TabIndex        =   46
         Text            =   "45"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxAlt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4440
         TabIndex        =   45
         Text            =   "12,7"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxMargEsq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3720
         TabIndex        =   44
         Text            =   "14"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxMarS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3000
         TabIndex        =   43
         Text            =   "13"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxQtdCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   34
         Text            =   "4"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxDistH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   33
         Text            =   "3"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxDistV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         TabIndex        =   32
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxQtdLinha 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   840
         TabIndex        =   31
         Text            =   "20"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em:"
         Height          =   195
         Left            =   840
         TabIndex        =   63
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cópias:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label lblAltBar 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2520
         TabIndex        =   59
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Alt. Bar."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         TabIndex        =   58
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblTamBar 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   840
         TabIndex        =   57
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tam. Bar."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   840
         TabIndex        =   56
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Páginas:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblTotalEti 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1680
         TabIndex        =   53
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Etiquetas:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1680
         TabIndex        =   52
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label lblTamFonte 
         AutoSize        =   -1  'True
         Caption         =   "10"
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
         Left            =   720
         TabIndex        =   51
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblFonte 
         AutoSize        =   -1  'True
         Caption         =   "Fonte:"
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
         TabIndex        =   50
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Fonte:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Largura:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5175
         TabIndex        =   42
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Altura:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4470
         TabIndex        =   41
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mrg.Esq:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3720
         TabIndex        =   40
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Mrg.Sup:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2970
         TabIndex        =   39
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Colunas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Dist.Hor.:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   37
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Dist.Vert:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   36
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Linhas"
         Height          =   195
         Left            =   870
         TabIndex        =   35
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   17085
      TabIndex        =   1
      Top             =   0
      Width           =   17115
      Begin VB.OptionButton optnCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         TabIndex        =   78
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton OptProd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Busca Avançada:"
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   7800
         TabIndex        =   16
         Top             =   240
         Width           =   8295
         Begin VB.CommandButton cmndbtnBuscaFab 
            Height          =   315
            Left            =   3000
            Picture         =   "frmGeraEtiquetas.frx":07BC
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnConf 
            Caption         =   "Listar"
            Height          =   315
            Left            =   6930
            TabIndex        =   22
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtbxBusca 
            Height          =   315
            Left            =   4800
            TabIndex        =   21
            Top             =   1080
            Width           =   2025
         End
         Begin VB.ComboBox cmbxBusca 
            Height          =   315
            ItemData        =   "frmGeraEtiquetas.frx":08B5
            Left            =   3360
            List            =   "frmGeraEtiquetas.frx":08C2
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1080
            Width           =   1425
         End
         Begin VB.ComboBox cmbxSecao 
            Height          =   315
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Width           =   3105
         End
         Begin VB.ComboBox cmbxSubSec 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   480
            Width           =   3465
         End
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
            Left            =   240
            TabIndex        =   17
            Top             =   1080
            Width           =   2745
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Listar de:"
            Height          =   195
            Left            =   3360
            TabIndex        =   28
            Top             =   900
            Width           =   645
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Expressão:"
            Height          =   195
            Left            =   4800
            TabIndex        =   27
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Seção:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   26
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "&Sub:"
            Height          =   255
            Left            =   3360
            TabIndex        =   25
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "Fabrica"
            Height          =   255
            Left            =   195
            TabIndex        =   24
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Adicionar Produto Avulso:"
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         Begin VB.TextBox txtOrgao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3840
            TabIndex        =   79
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CommandButton cmndbtnAdd 
            Caption         =   "&Adicionar"
            Height          =   495
            Left            =   6360
            Picture         =   "frmGeraEtiquetas.frx":08E9
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3120
            MaxLength       =   40
            TabIndex        =   6
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtbxExpr 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            MaxLength       =   40
            TabIndex        =   4
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox txtbxCodBarra 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3840
            MaxLength       =   40
            TabIndex        =   8
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblOrgao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento"
            Height          =   195
            Left            =   3840
            TabIndex        =   80
            Top             =   840
            Width           =   960
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "R$:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2880
            TabIndex        =   15
            Tag             =   "!"
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Tag             =   "!"
            Top             =   1200
            Width           =   330
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   600
            TabIndex        =   13
            Tag             =   "!"
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label lblCodFab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   195
            Left            =   1200
            TabIndex        =   12
            Tag             =   "!"
            Top             =   840
            Width           =   195
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Tag             =   "!"
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&qtd:"
            Height          =   195
            Left            =   3120
            TabIndex        =   5
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&xpressão:"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lblBarraCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Código de Barras:"
            Height          =   195
            Left            =   3840
            TabIndex        =   7
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblPreco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3480
            TabIndex        =   10
            Tag             =   "!"
            Top             =   840
            Width           =   195
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msflxgrd 
      Height          =   8175
      Left            =   6720
      TabIndex        =   0
      Top             =   2160
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   14420
      _Version        =   393216
      AllowUserResizing=   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmGeraEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmbxBusca_Click()
    cmbxBusca.tag = Left(cmbxBusca.Text, 1)
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxCopia_Click()
    cmbxCopia.tag = F_BuscaUmCampo("tipo_cop", "tab_copias", "cod_cop='" & cmbxCopia.Text & "'")
End Sub

Private Sub cmbxSubSec_Click()
    cmbxSubSec.tag = F_BuscaUmCampo("cod_sub", "tab_subsec", "nome_sub='" & cmbxSubSec.Text & "'")
End Sub

Private Sub cmbxSubSec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAbrir_Click()
    Dim e As etiqueta.etiqueta
    e = F_BuscaEtiqueta("")
    If e.cod <> "" Then
        lblDescEti.caption = e.desc
        txtbxAlt.Text = e.Altura
        txtbxDistH.Text = e.distH
        txtbxDistV.Text = e.distV
        txtbxLarg.Text = e.Largura
        txtbxMargEsq.Text = e.mEsq
        txtbxMarS.Text = e.mSup
        txtbxQtdCol.Text = e.qtdCol
        txtbxQtdLinha.Text = e.qtdLinha
        lblFonte.caption = e.fonte
        lblTamFonte.caption = e.tamfonte
        lblTamBar.caption = e.tambar
    End If
End Sub

Private Sub cmndbtnAdd_Click()
    Dim strSQL As String
    If txtbxExpr.tag = "" Then
        MsgBox "Escolha primeiramente o Produto!", vbExclamation
        txtbxExpr.SetFocus
        Exit Sub
    End If
    If CCur(txtbxQtd.Text) = 0 Then
        MsgBox "Informe uma quantidade válida!", vbExclamation
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If cmndbtnAdd.tag = "" Then
       strSQL = txtbxQtd.Text & Chr(9)
       strSQL = strSQL & txtbxExpr.tag & Chr(9)
       strSQL = strSQL & txtbxExpr.Text & Chr(9)
       
       strSQL = strSQL & lblCodFab.caption & Chr(9)
       strSQL = strSQL & lblPreco.caption & Chr(9)
       strSQL = strSQL & lblEnd.caption & " " & txtOrgao.Text
       msflxgrd.AddItem strSQL
       lblTotalEti.caption = Int(lblTotalEti.caption) + Int(txtbxQtd.Text)
       Inf
    Else
        With msflxgrd
            .row = Int(cmndbtnAdd.tag)
            .col = 0: .Text = txtbxQtd.Text
            .col = 1: .Text = FormatCodigo(txtbxCodBarra.Text)
            .col = 2: .Text = txtbxExpr.Text
            .col = 3: .Text = lblCodFab.caption
            .col = 4: .Text = lblPreco.caption
            
            If optnCliente.Value Then
                .col = 5
                .Text = lblEnd.caption
            
            End If
            
            
            cmndbtnAdd.tag = ""
        End With
    End If
    txtbxExpr.SetFocus
End Sub
Sub Inf()
    On Error Resume Next
    Dim iTot As Single
    Dim iAtual As Single
    iAtual = CSng(lblTotalEti.caption)
    iTot = Int(txtbxQtdCol.Text) * Int(txtbxQtdLinha.Text)
    If iAtual = 0 Then
        lblPaginas.caption = "0"
        Exit Sub
    End If
    If iAtual <= iTot Then
        lblPaginas.caption = 1
        Exit Sub
    End If
    If iTot Mod iAtual = 0 Then
        lblPaginas.caption = iTot / iAtual
    Else
        lblPaginas.caption = Int(iAtual / iTot) + 1
    End If
End Sub

Private Sub cmndbtnBuscaFab_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxFabricante.Text)
    If f.cod_For <> "" Then
        txtbxFabricante.Text = f.nome_for
        txtbxFabricante.tag = f.cod_For
    End If

End Sub

Private Sub CmndbtnConf_Click()
    If MsgBox("Confirma busca?", vbYesNo + vbQuestion) = vbYes Then
        
        If cmbxBusca.ListIndex = 0 Then
            BuscaEntrada "E"
        ElseIf cmbxBusca.ListIndex = 1 Then
            BuscaEntrada "C"
        Else
            BuscaAvancada
        End If
    End If
End Sub

Private Sub cmndbtnEdit_Click()
    If msflxgrd.rows = 1 Then Exit Sub
    With msflxgrd
        .col = 0:  txtbxQtd.Text = .Text
        .col = 1: txtbxQtd.tag = F_SoNumero(.Text)
        txtbxCodBarra.Text = txtbxQtd.tag
        .col = 2: lblCodFab.caption = .Text
        .col = 3: txtbxExpr.Text = .Text
        .col = 4: lblMarca.caption = .Text
        .col = 5: lblEnd.caption = .Text
        .col = 6: lblPreco.caption = FormatVal(.Text)
        
        cmndbtnAdd.tag = msflxgrd.row
        txtbxExpr.SetFocus
    End With
End Sub

Private Sub cmndbtnExcl_Click()
    Dim iAux As Integer
    If msflxgrd.rows = 1 Then Exit Sub
    
    If MsgBox("Confirma exclusão da linha?", vbYesNo + vbQuestion) = vbYes Then
        msflxgrd.col = 0
        iAux = msflxgrd.Text
        
        If msflxgrd.rows = 2 Then
            msflxgrd.row = 1
        Else
            msflxgrd.RemoveItem msflxgrd.row
        End If
        lblTotalEti.caption = Int(lblTotalEti.caption) - iAux
    End If
End Sub

Private Sub cmndbtnImprimir_Click()
    ImprimeEti frmVisualiza.Pic, CSng(txtbxAlt.Text), CSng(txtbxLarg.Text), CSng(txtbxDistH.Text), _
            CSng(txtbxDistV.Text), CSng(txtbxQtdCol.Text)
    frmVisualiza.Show 1
    
    If frmVisualiza.tag <> "" Then
        ImprimeEti Printer, CSng(txtbxAlt.Text), CSng(txtbxLarg.Text), CSng(txtbxDistH.Text), _
            CSng(txtbxDistV.Text), CSng(txtbxQtdCol.Text)
        Printer.EndDoc
    End If
    Unload frmVisualiza
End Sub

Sub ImprimeEti(obj As Object, Optional sAlt As Single = 20, Optional sLarg As Single = 89, _
     Optional sDistH As Single = 2, Optional sDistV As Single = 2, Optional IQtdFolha As Integer = 2)
'    Dim iCount As Integer
    Dim iQtd As Integer
    Dim iLinha As Integer
    Dim iEti As Integer
    Dim sMar As String
    Dim sNome As String
    Dim sCod As String
    Dim sRef As String
    Dim sEnd As String
    Dim sPre As String
    Dim sPos As Single
    Dim bCont As Boolean
    Dim sYIni As Single
    Dim sInicioX As Single
    Dim sInicioY As Single
    Dim sAuxY As Single
    Dim iImprimiu As Integer
    Dim iPage As Integer
    Dim sCentro As Single, sQtdLinha As Single
    Dim sUltY As Single, iAux As Integer, iPos As Integer, iPrimEsp As Integer
    Dim iPulaLin As Integer
    Dim iPulaCol As Integer
    Dim iRest As Integer
    Dim sRecorteEtiX As Single
    Dim sRecorteEtiY As Single
    Dim sAv As String
    Dim sAp As String
    Dim sYImprPreco As Single
    Dim sTamBarra As Single
    Dim sParcelado As String
    Dim strSQL As String
 '   iPulaLin = Int(txtbxIni.Text) / Int(txtbxQtdCol.Text)
        
   
        
    
'        iPulaCol = (Int(txtbxIni.Text) Mod Int(txtbxQtdCol.Text)) - 1
        msflxgrd.AddItem CStr(Int(txtbxIni.Text) - 1), 1
    
        
    
        obj.ScaleMode = vbMillimeters
'        On Error GoTo trata_erro
        
        If TypeName(obj) = "Printer" Then
            sQtdLinha = CSng(txtbxQtdLinha.Text)
            'obj.PaperSize = vbPRPSUser
            'obj.ScaleHeight = sAlt * (sQtdLinha) + sDistV * (sQtdLinha) + CSng(txtbxMarS.Text)
            'obj.ScaleWidth = sLarg * (IQtdFolha) + sDistH * (IQtdFolha) + CSng(txtbxMargEsq.Text)
        End If
        sInicioX = Int(txtbxMargEsq.Text)
        sInicioY = Int(txtbxMarS.Text)
        iImprimiu = iPulaLin
        With msflxgrd
        On Error Resume Next
        obj.FontName = "arial narrow"
        obj.fontsize = 6
            iPos = iPulaCol
            For iLinha = 1 To .rows - 1
                .row = iLinha
                .col = 0
                iEti = CInt(.Text) 'quantidade de etiquetas
                
                    For iAux = 1 To iEti
                        .col = 3: sRef = .Text
                        .col = 1: sCod = F_SoNumero(.Text)
                        .col = 2: sNome = .Text
                        .col = 5: sEnd = .Text
                        sPre = ""
                        If chkPrazo.Value = 1 Then
                            .col = 4:
                            If optnCliente.Value = False Then
                                If .Text <> "" Then
                                    sPre = " A Prazo:" & FormatVal(.Text) & " 3X " & FormatVal(CCur(.Text) / 3) & ""
                                    sAp = "R$ " & FormatVal(.Text)
                                   ' sParcelado = "3X " & FormatVal(CCur(.Text) / 3) & ""
                                End If
                            End If
                        End If
                        If chkVista.Value = 1 Then
                            .col = 4: sPre = sPre & " ou " & FormatVal(CCur(.Text) - (CCur(.Text) * (crcDesconto_Avista / 100))) & " A Vista"
                           ' sAv = "A VISTA:" & FormatVal(CCur(.Text) - (CCur(.Text) * (crcDesconto_Avista / 100)))
                        End If
                        
                        sPos = sLarg * (iPos) + (sDistH * (iPos)) + sInicioX
                        
                        sAuxY = sInicioY + ((sAlt) * (iImprimiu)) + (sDistV * (iImprimiu))
                        
                        
                        
                        sRecorteEtiX = sPos - (sDistH * (iPos))
                        
                        
                        sRecorteEtiY = sInicioY + ((sAlt) * (iImprimiu)) + (0 * (iImprimiu))
                        
                        
                        
                        obj.CurrentY = sAuxY
                        obj.fontsize = lblTamFonte.caption
                        obj.FontName = lblFonte.caption
                        iPrimEsp = F_RetornaPrimEspaco(sNome, 20)
                        brcdx.caption = sCod
                        'sNome = sNome + sPre
                        
                            'sNome = sNome + sPre
                        
                        
                        P_ImprQuebraTurbo sNome, obj, Int(sPos), sPos + sLarg - 5
                        
                        
                        sYImprPreco = obj.CurrentY
                        If sCod <> "" Then
                            If chkBarra.Value = 1 Then
                                
                                obj.PaintPicture brcdx.Picture, sPos, obj.CurrentY
                            
                            Else
                                If optnCliente.Value Then
                                    obj.fontsize = 7
                                    'P_PrintXY obj, sEnd, sPos, obj.CurrentY, False
                                    P_ImprQuebraTurbo sEnd, obj, Int(sPos), sPos + sLarg - 5
                                Else
                                    obj.FontBold = True
                                    obj.fontsize = CInt(lblTamFonte.caption) + 2
                                    P_PrintXY obj, sCod, sPos, obj.CurrentY, False
                                End If
                                
                            
                            End If
                            'sTamBarra = obj.ScaleX(brcdx.Width)
                        End If
                        'obj.Print ""
                        obj.ScaleMode = Printer.ScaleMode
                        
                        'preços
                        obj.FontBold = True
                        strSQL = obj.FontName
                        obj.FontName = "Times New Roman"
                        obj.fontsize = CInt(lblTamFonte.caption) + 4
                        P_PrintXY obj, sAp, sPos + sLarg - 10, sYImprPreco, False, 3
                        P_PrintXY obj, "", sPos, obj.CurrentY, False, 1
                        P_PrintXY obj, sParcelado, sPos + sLarg - 10, obj.CurrentY, False, 3
                        obj.FontName = "Courier New"
                        obj.fontsize = 8
                        obj.FontBold = False
                        obj.FontName = strSQL
                          '  P_ImprX obj, sCod & " - " & "R$ " & sPre, CInt(sPos), True
                        
                        iPos = iPos + 1
                        
                        
                        
                        If chkRecorte.Value = 1 Then
                            'P_PrintXY obj, "|", sRecorteEtiX, sRecorteEtiY
                            P_Line obj, sRecorteEtiX, sRecorteEtiY, sRecorteEtiX, sRecorteEtiY + 1, vbSolid
                            P_Line obj, sRecorteEtiX, sRecorteEtiY, sRecorteEtiX + 1, sRecorteEtiY, vbSolid
                            
                            P_Line obj, sRecorteEtiX, sRecorteEtiY + sAlt, sRecorteEtiX, sRecorteEtiY - 1 + sAlt, vbSolid
                            P_Line obj, sRecorteEtiX, sRecorteEtiY + sAlt, sRecorteEtiX - 1, sRecorteEtiY + sAlt, vbSolid
                            
                            
                        End If
                        
                        
                        If iPos >= IQtdFolha Then
                            iPos = 0
                            iImprimiu = iImprimiu + 1
                            If iImprimiu >= Int(txtbxQtdLinha.Text) Then
                                iImprimiu = CInt(txtbxIni.Text) - 1
                                If TypeName(obj) = "Printer" Then
                                    obj.NewPage
                                Else
                                    P_AddPicVis obj
                                End If
                            End If
                        End If
                    
                Next
            
            Next
            P_AddPicVis obj
        End With
    msflxgrd.RemoveItem 1
Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLimpa_Click()
    If msflxgrd.rows = 1 Then Exit Sub
    If MsgBox("Confirma exclusão de todas as linhas?", vbYesNo + vbQuestion) = vbYes Then
        msflxgrd.rows = 1
        lblTotalEti.caption = "0"
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub




Private Sub cmndbtnSalvarComo_Click()
    frmCadEti.caption = "Novo Modelo"
    frmCadEti.txtbxAltura.Text = txtbxAlt.Text
    frmCadEti.txtbxDistH.Text = txtbxDistH.Text
    frmCadEti.txtbxDistV.Text = txtbxDistV.Text
    frmCadEti.txtbxLargura.Text = txtbxLarg.Text
    frmCadEti.txtbxMEsq.Text = txtbxMargEsq.Text
    frmCadEti.txtbxMSup.Text = txtbxMarS.Text
    frmCadEti.txtbxNome.Text = "nome da etiqueta"
    frmCadEti.txtbxQtdCol.Text = txtbxQtdCol.Text
    frmCadEti.txtbxQtdLinha.Text = txtbxQtdLinha.Text
    frmCadEti.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    If bActive Then Exit Sub
    buscaPadrao
    If Me.tag <> "" Then
        cmbxComp.Text = Me.tag
        BuscaEntrada "e"
    End If
    bActive = True
End Sub
Sub buscaPadrao()
    Dim e As etiqueta.etiqueta
    e = F_BuscaEtiquetaPadrao
    If e.cod <> "" Then
        lblDescEti.caption = e.desc
        txtbxAlt.Text = e.Altura
        txtbxDistH.Text = e.distH
        txtbxDistV.Text = e.distV
        txtbxLarg.Text = e.Largura
        txtbxMargEsq.Text = e.mEsq
        txtbxMarS.Text = e.mSup
        txtbxQtdCol.Text = e.qtdCol
        txtbxQtdLinha.Text = e.qtdLinha
        lblFonte.caption = e.fonte
        lblTamFonte.caption = e.tamfonte
        lblTamBar.caption = e.tambar
    End If
End Sub
Private Sub Form_Load()
    bActive = False
    resizeForm Me
    msflxgrd.Cols = 6
    msflxgrd.FixedCols = 0
    msflxgrd.FixedRows = 1
    msflxgrd.rows = 1
    msflxgrd.col = 0
    msflxgrd.Text = "Qtd"
    msflxgrd.col = 1
    msflxgrd.Text = "Código"
    msflxgrd.col = 2
    msflxgrd.Text = "Descrição:"
    msflxgrd.col = 3
    msflxgrd.Text = "Ref:"
    msflxgrd.col = 4
    msflxgrd.Text = "Preço:"
    msflxgrd.col = 5
    msflxgrd.Text = "Endereço:"
    
    msflxgrd.ColWidth(0) = 200
    msflxgrd.ColWidth(1) = 1000
    msflxgrd.ColWidth(2) = 4500
    P_PreencheSecao cmbxSecao
    
End Sub

Private Sub cmbxSecao_Click()
    cmbxSecao.tag = F_BuscaUmCampo("cod_sec", "tab_secoes", "nome_sec='" & cmbxSecao.Text & "'")
    
    P_PreencheSubSec cmbxSubSec, cmbxSecao.tag
End Sub

Private Sub cmbxSecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        P_PreencheSecao cmbxSecao
        
    End If
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub Form_Resize()
    msflxgrd.width = Me.width - msflxgrd.Left - 200
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If msflxgrd.rows = 2 Then
            msflxgrd.rows = 1
        ElseIf msflxgrd.rows > 2 Then
            msflxgrd.RemoveItem msflxgrd.row
        End If
    End If
End Sub

Private Sub Option4_Click()

End Sub

Private Sub optnEstoque_Click()
    cmbxComp.Enabled = False
End Sub

Private Sub optnEstoque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optn_Click()
    cmbxComp.Clear
End Sub

Private Sub optnGrupo_Click()
    'P_PreencheCateg cmbxComp
    cmbxComp.Enabled = True
End Sub

Private Sub optnGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnMarca_Click()
    
    cmbxComp.Enabled = True
End Sub

Private Sub optnMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnSub_Click()
    'P_PreencheSubCat cmbxComp
    cmbxComp.Enabled = True
End Sub

Private Sub optnSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnCliente_Click()
    chkBarra.Value = 0
    frameOpt.Visible = False
    lblBarraCaption.caption = "Nº Doc.:"
    txtbxQtd.Text = "1"
End Sub

Private Sub OptProd_Click()
    chkBarra.Value = 1
    frameOpt.Visible = True
    lblBarraCaption.caption = "Código de Barras:"
End Sub

Private Sub txtbxBusca_GotFocus()
    P_SelecTexto txtbxBusca
End Sub

Private Sub txtbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaComp
    End If
End Sub
Sub buscaComp()
    Dim strSQL As String
    
    If cmbxBusca.ListIndex = 0 Then
        buscaMarca
    End If
    
End Sub
Sub buscaMarca()
    
End Sub
Sub buscaCat()

End Sub
Private Sub txtbxBusca_LostFocus()
    P_AjustaTexto txtbxBusca
End Sub

Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If optnCliente.Value = True Then
            buscaCliente True
        Else
            Sendkeys Chr(9)
        End If
    End If
    
End Sub

Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If OptProd.Value Then
            buscaProd
        Else
            buscaCliente False
        End If
        
    End If
End Sub
Sub buscaProd()
    
    Dim p As Produto
    p = F_BuscaProdSaida(txtbxExpr.Text)
    If p.cod_pro <> "" Then
        txtbxExpr.Text = p.nome_pro
        txtbxExpr.tag = p.cod_pro
        lblCodFab.caption = p.ref_pro
        lblPreco.caption = FormatVal(p.pv_pro)
        txtbxCodBarra.Text = p.cod_pro
        
        txtbxQtd.SetFocus
        
    End If
    
    
    
End Sub
Sub buscaCliente(Optional porDoc As Boolean = False)
    Dim p As Clientes.cliente
    p = F_BuscaCliente(txtbxExpr.Text, porDoc)
    If p.cod_cli <> "" Then
        txtbxExpr.Text = p.nome_Cli
        txtbxExpr.tag = p.cod_cli
        lblCodFab.caption = p.cod_cli
        
        'lblPreco.caption = FormatVal(p.pv_pro)
        'txtbxCodBarra.Text = p.cod_cli
        lblEnd.caption = p.endereco_cli + " " + p.bairro_cli + " " + p.cidade_cli
        If p.org.cod_org <> "" And p.org.cod_org <> "0" Then
             txtOrgao.Text = p.org.nome_org & " fone: " & p.telefone_cli
        Else
             txtOrgao.Text = p.endtrab_cli + " " + p.setortrab.nome + " " + p.setortrab.nome + " - " & p.telefone_cli
        End If
       
        
        
        txtbxQtd.SetFocus
        
    End If
    
    
End Sub

Sub BuscaAvancada()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
'    On Error GoTo trata_erro
    P_Status "Buscando produtos..."
    Dim bEst As Boolean
    strSQL = "select tp.*  from tab_produtos as tp,  "
    strSQL = strSQL & " tab_secoes as ts, tab_subsec as tu ,tab_fabricantes as tf "
    strSQL = strSQL & " where tp.cod_sub=tu.cod_sub and tu.cod_sec=ts.cod_sec "
    strSQL = strSQL & " AND tp.cod_fab=tf.cod_fab "
    If cmbxSecao.ListIndex > -1 Then
        strSQL = strSQL & " and ts.cod_sec='" & cmbxSecao.tag & "' "
    End If
    If cmbxSubSec.ListIndex > -1 Then
        strSQL = strSQL & " and tp.cod_sub='" & cmbxSubSec.tag & "' "
    End If
    If txtbxFabricante.tag <> "" Then
        strSQL = strSQL & " and tf.cod_fab=tp.cod_fab "
    End If
    
    bEst = MsgBox("Deseja listar produtos sem estoque?", vbYesNo + vbQuestion) = vbYes
    If Not bEst Then
        strSQL = strSQL & " and er_pro>0 "
    End If
    strSQL = strSQL & " order by nome_pro "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        msflxgrd.rows = 1
        msflxgrd.ColAlignment(2) = 1
        lblTotalEti.caption = "0"
        Do While Not .EOF
            If CCur(txtbxQtd.Text) > 0 Then
                strSQL = Int(txtbxQtd.Text) & Chr(9)
            Else
                If !er_pro > 0 Then
                    strSQL = !er_pro & Chr(9)
                Else
                    strSQL = "0" & Chr(9)
                End If
            End If
            
            strSQL = strSQL & FormatCodigo(!cod_pro) & Chr(9)
            
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & !ref_pro & Chr(9)
            
            strSQL = strSQL & FormatVal(!pv_pro)
            msflxgrd.AddItem strSQL
            lblTotalEti.caption = Int(lblTotalEti.caption) + !er_pro
            .MoveNext
        Loop
        Inf
        .Close
        
    End With
    P_Status ""
Exit Sub
trata_erro:
    On Error Resume Next
    P_Status ""
    MsgBox Err.Description, vbCritical
End Sub
Sub BuscaEntrada(tipo As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
'    On Error GoTo trata_erro
    P_Status "Buscando produtos..."
    Dim sTab As String
    Dim sTabIte As String
    Dim sCampo As String
    
    If tipo = "E" Then
        sTab = "tab_cmp"
        sTabIte = "tab_itcmp"
        sCampo = "seq_not"
    Else
        sTab = "tab_contagens"
        sTabIte = "tab_itcont"
        sCampo = "seq_cnt"
    End If
    
    
    strSQL = "select distinct ti.*,tp.nome_pro,tm.nome_mar,tp.cod_pro,tp.ref_pro,tp.endereco_pro,"
    strSQL = strSQL & " tp.pv_pro,tp.pa_pro,tc.er_pro "
    strSQL = strSQL & " from tab_produtos as tp,  "
    strSQL = strSQL & " tab_marcas as tm,"
    strSQL = strSQL & " tab_cmp as te, tab_itentprv as ti,"
    strSQL = strSQL & " tab_itcmp as ti"
    strSQL = strSQL & " where tp.cod_pro=tc.cod_pro "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and ti.seq_ent=te.seq_ent "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    
    strSQL = strSQL & " and ti.seq_ent='" & cmbxComp.Text & "' "
    
    strSQL = strSQL & " order by nome_pro,tam_pro "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        msflxgrd.rows = 1
        msflxgrd.ColAlignment(2) = 1
        lblTotalEti.caption = "0"
        Do While Not .EOF
            strSQL = !qtd_ite & Chr(9)
            strSQL = strSQL & FormatCodigo(!cod_pro) & Chr(9)
            
            strSQL = strSQL & !nome_pro & Chr(9)
            strSQL = strSQL & !ref_pro & Chr(9)
            
            
            strSQL = strSQL & FormatVal(!pv_pro)
            
            msflxgrd.AddItem strSQL
            lblTotalEti.caption = Int(lblTotalEti.caption) + !qtd_ite
            .MoveNext
        Loop
        Inf
        .Close
        
    End With
    P_Status ""
Exit Sub
trata_erro:
    On Error Resume Next
    P_Status ""
    MsgBox Err.Description, vbCritical
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
End Sub

Private Sub txtbxFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaFab_Click
    End If
End Sub

Private Sub txtbxIni_GotFocus()
    P_SelecTexto txtbxIni
End Sub

Private Sub txtbxIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxIni_LostFocus()
    P_ValorLostFocus txtbxIni, "#"
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###,##0"
End Sub


Private Sub txtOrgao_GotFocus()
    P_SelecTexto txtOrgao
End Sub

Private Sub txtOrgao_KeyPress(KeyAscii As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

