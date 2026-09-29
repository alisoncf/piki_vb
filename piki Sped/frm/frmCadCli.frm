VERSION 5.00
Begin VB.Form frmCadCli 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   13695
   Icon            =   "frmCadCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   13665
      TabIndex        =   62
      Top             =   7395
      Width           =   13695
      Begin VB.CommandButton cmdHist 
         Caption         =   "Histórico"
         Height          =   375
         Left            =   6480
         TabIndex        =   91
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdVendas 
         Caption         =   "Pesquisar"
         Height          =   375
         Left            =   5280
         TabIndex        =   88
         Top             =   120
         Width           =   1095
      End
      Begin VB.PictureBox pcDetalhes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   5175
         TabIndex        =   75
         Top             =   0
         Width           =   5175
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "em aberto"
            Height          =   195
            Left            =   2520
            TabIndex        =   87
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "em parcelas"
            Height          =   195
            Left            =   2520
            TabIndex        =   86
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   195
            Left            =   2520
            TabIndex        =   85
            Top             =   0
            Width           =   360
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Primeira"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maior"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   480
            Width           =   390
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Última:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblParcAb 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   3480
            TabIndex        =   81
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblParc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   3480
            TabIndex        =   80
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblCompras 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   3480
            TabIndex        =   79
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblMaiorCompr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   720
            TabIndex        =   78
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblUltCompr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   720
            TabIndex        =   77
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblPrimCompr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   720
            TabIndex        =   76
            Top             =   0
            Width           =   390
         End
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "&Fecha&r"
         Height          =   735
         Left            =   11640
         Picture         =   "frmCadCli.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "Sal&var"
         Height          =   735
         Left            =   9960
         Picture         =   "frmCadCli.frx":0E35
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox pcOutrosDados 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6765
      Left            =   7200
      ScaleHeight     =   6735
      ScaleWidth      =   6345
      TabIndex        =   40
      Top             =   600
      Width           =   6375
      Begin VB.ComboBox cmbxCFOP 
         Height          =   315
         ItemData        =   "frmCadCli.frx":0F7F
         Left            =   120
         List            =   "frmCadCli.frx":0F81
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   4440
         Width           =   6015
      End
      Begin VB.ComboBox cmbxTipoICMS 
         Height          =   315
         ItemData        =   "frmCadCli.frx":0F83
         Left            =   120
         List            =   "frmCadCli.frx":0F85
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   3840
         Width           =   6015
      End
      Begin VB.ComboBox cmbxIndIEDest 
         Height          =   315
         ItemData        =   "frmCadCli.frx":0F87
         Left            =   120
         List            =   "frmCadCli.frx":0F94
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   3240
         Width           =   6030
      End
      Begin VB.CommandButton cmdDesconto 
         Caption         =   "Desconto"
         Height          =   375
         Left            =   5040
         TabIndex        =   72
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   71
         Text            =   "0,00"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtFormasPag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   66
         Top             =   2280
         Width           =   5775
      End
      Begin VB.PictureBox pcFormasPag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5880
         Picture         =   "frmCadCli.frx":1082
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   65
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtbxLimite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   44
         Text            =   "0,00"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cmbxSituacao 
         Height          =   315
         ItemData        =   "frmCadCli.frx":142D
         Left            =   120
         List            =   "frmCadCli.frx":143A
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox txtbxObs 
         Appearance      =   0  'Flat
         Height          =   1365
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   5160
         Width           =   6135
      End
      Begin VB.Label Label35 
         Caption         =   "CFOP"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblT 
         Caption         =   "Tipo de ICMS"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preferências de NFe"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   92
         Top             =   2640
         Width           =   2265
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&IndIE NFe:"
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   3000
         Width           =   765
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Height          =   255
         Left            =   5280
         TabIndex        =   73
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         Height          =   195
         Left            =   4680
         TabIndex        =   70
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pagamento Preferencial:"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   2040
         Width           =   2445
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restante:"
         Height          =   195
         Left            =   3480
         TabIndex        =   61
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lblRestante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   3120
         TabIndex        =   60
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Em Uso:"
         Height          =   195
         Left            =   1920
         TabIndex        =   59
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label lblUso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
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
         Height          =   360
         Left            =   1560
         TabIndex        =   58
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de Crédito:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   2010
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Situação do Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Observações/Anotações:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   4920
         Width           =   1830
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   7200
      ScaleHeight     =   375
      ScaleWidth      =   6345
      TabIndex        =   55
      Top             =   120
      Width           =   6375
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outros Dados:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   0
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   6900
      TabIndex        =   53
      Top             =   5400
      Width           =   6930
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   54
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   6900
      TabIndex        =   51
      Top             =   3120
      Width           =   6930
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   52
         Top             =   0
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   6900
      TabIndex        =   49
      Top             =   120
      Width           =   6930
      Begin VB.OptionButton optnPessoa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pessoa &Física"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   69
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optnPessoa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pessoa &Jurídica"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5160
         TabIndex        =   68
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Identificação:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   50
         Top             =   0
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   6900
      TabIndex        =   37
      Top             =   5760
      Width           =   6930
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   6615
      End
      Begin VB.TextBox txtbxFax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtbxTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxCel_cli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Email:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Left            =   3480
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Celular:"
         Height          =   195
         Left            =   1800
         TabIndex        =   32
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1905
      ScaleWidth      =   6900
      TabIndex        =   36
      Top             =   3480
      Width           =   6930
      Begin VB.TextBox txtLogradouro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   18
         Top             =   360
         Width           =   3735
      End
      Begin VB.PictureBox pcCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5160
         Picture         =   "frmCadCli.frx":146D
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   29
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         MaxLength       =   14
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtComplemento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox txtbxCEP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtbxBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label lblCodCid 
         Caption         =   "Código Cidade"
         Height          =   255
         Left            =   5520
         TabIndex        =   74
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblUf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   64
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   6120
         TabIndex        =   21
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   3960
         TabIndex        =   19
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4200
         TabIndex        =   63
         Top             =   120
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CE&P:"
         Height          =   195
         Left            =   5520
         TabIndex        =   25
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bairro: (Setor, Vila)"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Logradouro: (Rua, Av, etc)"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1890
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   6900
      TabIndex        =   0
      Top             =   480
      Width           =   6930
      Begin VB.TextBox txtFantasia 
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
         MaxLength       =   50
         TabIndex        =   4
         Top             =   960
         Width           =   6615
      End
      Begin VB.TextBox txtIe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         MaxLength       =   16
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtbxMae 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   4215
      End
      Begin VB.TextBox txtbxInscProdRural 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         MaxLength       =   16
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtRg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   16
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtbxCNPJ 
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
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtbxNasc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         MaxLength       =   18
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtbxRazao 
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
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label lblFantasia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome &Fantasia:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label lblIE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Estadual:"
         Height          =   195
         Left            =   2760
         TabIndex        =   7
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label lblMae 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da &Mãe"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Insc. Prod. Rural:"
         Height          =   195
         Left            =   5400
         TabIndex        =   9
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label lblRg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RG:"
         Height          =   195
         Left            =   4440
         TabIndex        =   13
         Top             =   1920
         Width           =   285
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CNPJ:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label lblNasc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Nasc:"
         Height          =   195
         Left            =   5400
         TabIndex        =   15
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCadCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Dim bActive As Boolean
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL(TAB_CLIENTES, "cod_cli", "", "0000#")
    
    
    
    
    strSQL = "insert into " & TAB_CLIENTES & "(cod_cli,nome_cli,fantasia_cli,"
    
    If optnPessoa(0).Value Then
        strSQL = strSQL & " cpf_cli,"
    Else
        strSQL = strSQL & " cnpj_cli, "
    End If
    
    strSQL = strSQL & "rg_cli, inscprodrur_cli, "
    
    
    strSQL = strSQL & "bairro_cli,cidade_Cli,uf_cli,cep_cli,telefone_Cli,fax_Cli,"
    strSQL = strSQL & "nasc_cli,mae_cli,"
    strSQL = strSQL & "status_cli,tipo_cli,cel_cli,obs_Cli,limite_cli,numero_cli,complemento_cli,ie_cli,"
    strSQL = strSQL & " cod_cid,logradouro_cli,indiedest_cli,cod_tip,cfop_cli,tipoicms_cli) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & "'" & txtfantasia.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "'" & txtRg.Text & "',"
    strSQL = strSQL & "'" & txtbxInscProdRural.Text & "',"
    strSQL = strSQL & "'" & txtbxBairro.Text & "',"
    strSQL = strSQL & "'" & txtCidade.Text & "',"
    strSQL = strSQL & "'" & lblUF.caption & "',"
    strSQL = strSQL & "'" & txtbxCEP.Text & "',"
    strSQL = strSQL & "'" & txtbxTel.Text & "',"
    strSQL = strSQL & "'" & txtbxFax.Text & "',"
    strSQL = strSQL & F_GravaData(txtbxNasc.Text) & ","
   strSQL = strSQL & "'" & txtbxMae.Text & "',"
 
    
    strSQL = strSQL & "'" & Left(cmbxSituacao.Text, 1) & "',"
    If optnPessoa(0).Value = True Then
        strSQL = strSQL & "'F',"
    Else
        strSQL = strSQL & "'J',"
    End If
    strSQL = strSQL & "'" & txtbxCel_cli.Text & "',"
    strSQL = strSQL & "'" & txtbxObs.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxLimite.Text) & ", "
    strSQL = strSQL & "'" & txtNumero.Text & "',"
    strSQL = strSQL & "'" & txtcomplemento.Text & "',"
    strSQL = strSQL & "'" & txtIE.Text & "',"
    strSQL = strSQL & "'" & txtCidade.tag & "',"
    strSQL = strSQL & "'" & txtLogradouro.Text & "',"
    strSQL = strSQL & "'" & Left(cmbxIndIEDest.Text, 1) & "',"
    strSQL = strSQL & "'" & txtFormasPag.tag & "',"
    strSQL = strSQL & "'" & PegaCodigo(cmbxCFOP.Text, "-") & "',"
    strSQL = strSQL & "'" & PegaCodigo(cmbxTipoICMS.Text, "-") & "')"
    conn.Execute strSQL
    
End Sub
Sub P_Altera()
    Dim strSQL As String
    strSQL = "update " & TAB_CLIENTES & " set "
    strSQL = strSQL & "nome_cli='" & txtbxRazao.Text & "',"
    strSQL = strSQL & "fantasia_cli='" & txtfantasia.Text & "',"
    If optnPessoa(0).Value = True Then
        strSQL = strSQL & "cpf_Cli='" & txtbxCNPJ.Text & "',"
    Else
        strSQL = strSQL & "cnpj_Cli='" & txtbxCNPJ.Text & "',"
    End If
    
    strSQL = strSQL & "rg_cli='" & txtRg.Text & "',"
    strSQL = strSQL & "ie_cli='" & txtIE.Text & "',"
    strSQL = strSQL & "inscprodrur_cli='" & txtbxInscProdRural.Text & "',"
    'strSQL = strSQL & "endereco_Cli='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "logradouro_Cli='" & txtLogradouro.Text & "',"
    strSQL = strSQL & "bairro_Cli='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "cidade_Cli='" & txtCidade.Text & "',"
'    strSQL = strSQL & "uf_Cli='" & cmbxUf.Text & "',"
    strSQL = strSQL & "cep_Cli='" & txtbxCEP.Text & "',"
    strSQL = strSQL & "telefone_Cli='" & txtbxTel.Text & "',  "
     strSQL = strSQL & "fax_Cli='" & txtbxFax.Text & "',  "
    strSQL = strSQL & "cel_Cli='" & txtbxCel_cli.Text & "',  "
    strSQL = strSQL & "obs_Cli='" & txtbxObs.Text & "',  "
    strSQL = strSQL & "NASC_CLI=" & F_GravaData(txtbxNasc.Text) & ","
    strSQL = strSQL & "limite_cli=" & F_GravaDec(txtbxLimite.Text) & ","
    strSQL = strSQL & "desconto_cli=" & F_GravaDec(txtDesconto.Text) & ","
    strSQL = strSQL & "mae_cli='" & txtbxMae.Text & "',"
    strSQL = strSQL & "numero_cli='" & txtNumero.Text & "',"
    strSQL = strSQL & "complemento_cli='" & txtcomplemento.Text & "',"
    strSQL = strSQL & "cod_tip='" & txtFormasPag.tag & "',"
     strSQL = strSQL & "email_cli='" & txtEmail.Text & "',"
    If optnPessoa(0).Value Then
        strSQL = strSQL & "tipo_cli='F',"
    Else
        strSQL = strSQL & "tipo_cli='J',"
    End If
    strSQL = strSQL & "cod_cid='" & txtCidade.tag & "',"
    strSQL = strSQL & "indiedest_Cli='" & Left(cmbxIndIEDest.Text, 1) & "',"
    strSQL = strSQL & "cfop_cli='" & PegaCodigo(cmbxCFOP.Text, "-") & "',"
    strSQL = strSQL & "tipoicms_cli='" & PegaCodigo(cmbxTipoICMS.Text, "-") & "',"
    strSQL = strSQL & "status_cli='" & Left(cmbxSituacao.Text, 1) & "' "
         
         
    strSQL = strSQL & " where cod_Cli = '" & Me.tag & "'"
    conn.Execute strSQL
End Sub

Private Sub cmbxEstadoCiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmdDesconto_Click()
    strSQL = F_senhaReturnUsu(False, True)
    If strSQL = "" Then
        MsgBox "não autenticado!", vbExclamation
        Exit Sub
    End If
    
    If strLiberarCli_usu <> "S" Then
        MsgBox "Você não te permissão para liberar/bloquear cliente!", vbExclamation
        Exit Sub
    End If
    
    txtDesconto.Enabled = True
    txtDesconto.SetFocus
    
End Sub

Private Sub cmdVendas_Click()
    P_StatusCli
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
            
    
    
    If txtbxRazao.Text = "" Then
        MsgBox "Informe um nome para o Cliente!", vbExclamation
        Exit Sub
        
    ElseIf txtbxCNPJ.Text = "" Then
        If optnPessoa(0).Value = True Then
            MsgBox "Informe o CPF do Cliente!", vbExclamation
            Exit Sub
        Else
            MsgBox "Informe o CNPJ do Cliente!", vbExclamation
            Exit Sub
        End If
    ElseIf bCNPJ = False Then
        If optnPessoa(0).Value = True Then
            MsgBox "Informe um CPF Válido para Cliente!", vbExclamation
            'Exit Sub
        Else
            MsgBox "Informe um CNPJ Válido para Cliente!", vbExclamation
            'Exit Sub
        End If
    ElseIf txtLogradouro.Text = "" Then
        MsgBox "Informe o endereço do Cliente!", vbExclamation
        Exit Sub
    ElseIf txtbxBairro.Text = "" Then
        MsgBox "informe o Bairro do Cliente!,vbExclamation  "
        Exit Sub
    ElseIf txtCidade.tag = "" Then
        MsgBox "Informe a Cidade do Cliente!", vbExclamation
        Exit Sub
    'ElseIf cmbxUf.Text = "" Then
'        MsgBox "Informe o Estado de Origem do Cliente!", vbExclamation
'        Exit Sub
    End If
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("cod_cli", TAB_CLIENTES, "cpf_cli='" & txtbxCNPJ.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este CPF/CNPJ!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_cli", TAB_CLIENTES, "nome_cli='" & txtbxRazao.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este Nome/Razão Social!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_cli", TAB_CLIENTES, "rg_cli='" & txtRg.Text & "'") <> "" Then
            If txtRg.Text <> "" Then
                MsgBox "Já existe um cliente com esta Inscrição Estadual /RG! ", vbCritical
                Exit Sub
            End If
        End If
    
        P_Salva
        MsgBox "Cliente Cadastrado com sucesso!", vbInformation
        P_Limpa

        txtbxRazao.Text = ""
        optnPessoa(0).SetFocus
    Else
        If MsgBox("Confirma alteração do cliente " & Me.tag & " ?", vbYesNo + vbQuestion) = vbYes Then
            P_Altera
            MsgBox "Cliente Alterado com sucesso!", vbInformation
            Unload Me
        End If
    End If
    
End Sub
Sub P_Limpa()
    Dim i As Integer
    For i = 0 To Me.Controls.Count - 1
        If LCase(Left(Me.Controls(i).Name, 3)) = "txt" Then
            Me.Controls(i).Text = ""
        End If
    Next
   
    
End Sub

Private Sub Combo1_Change()
    Dim reg As New RegExp
    
    
    
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Sub P_StatusCli()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    lblPrimCompr.caption = F_BuscaUmCampo("min(data_ven)", "" & TAB_VENDAS & "", "cod_cli='" & Me.tag & "'")
    lblUltCompr.caption = F_BuscaUmCampo("max(dataefet_ven)", "" & TAB_VENDAS & "", "cod_cli='" & Me.tag & "'")
    lblMaiorCompr.caption = "R$ " & FormatVal(F_BuscaUmCampo("max(valor_ven-desconto_ven)", "" & TAB_VENDAS & "", "cod_cli='" & Me.tag & "'"))
    lblCompras.caption = F_BuscaUmCampo("sum(valor_ven)", "" & TAB_VENDAS & "", "cod_cli='" & Me.tag & "'")
    lblParc.caption = F_BuscaUmCampo("count(*)", " " & TAB_ITDOC, "venc_itd<=" & F_GravaData(Date) & " and reman_itd>0 and cod_ter='" & Me.tag & "'")
    lblParcAb.caption = F_BuscaUmCampo("count(*)", " " & TAB_ITDOC, "reman_itd>0 and cod_ter='" & Me.tag & "'")
    
    
End Sub
Private Sub Form_Activate()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    
    'P_PreencheUF cmbxUfFirma


    
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        Carrega
    End If
    bActive = True
'    MsgBox strHora & "-" & Time
    
End Sub
Sub Carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select tc.*,tf.nome_for,cfop_cli,tipoicms_cli,tf.cod_for,ti.nome_cid,ti.uf_cid from " & TAB_CLIENTES & " as tc left join " & TAB_CIDADES & " as ti "
        strSQL = strSQL & " on tc.cod_cid=ti.cod_cid "
        strSQL = strSQL & " left join " & TAB_FORMASPAG & " as tf on tc.cod_tip=tf.cod_For "
        strSQL = strSQL & " where tc.cod_cli = '" & Me.tag & "' "
        abreObjetoAdo rPesq, strSQL
        With rPesq
            
                txtbxBairro.Text = Format(!bairro_cli)
                txtbxCEP.Text = Format(!cep_cli)
                txtCidade.Text = Format(!cidade_cli)
                
                If !tipo_cli = "F" Then
                    optnPessoa(0).Value = True
                    txtbxCNPJ.Text = Format(!cpf_cli)
                    optnPessoa_Click 0
                    
                Else
                    optnPessoa(1).Value = True
                    txtbxCNPJ.Text = Format(!cnpj_Cli)
                    optnPessoa_Click 1
                    
                End If
                
                
                txtfantasia.Text = Format(!fantasia_cli)
'                txtbxEndereco.Text = Format(!endereco_cli)
                txtRg.Text = Format(!rg_cli)
                txtbxRazao.Text = !nome_Cli
                txtbxTel.Text = Format(!telefone_cli)
                txtbxObs.Text = Format(!obs_cli)
                txtbxNasc.Text = Format(!nasc_cli)
                txtbxCel_cli.Text = Format(!cel_cli)
                txtbxLimite.Text = FormatVal(!limite_cli)
                lblUso.caption = FormatVal(F_RetornaLimiteUsoCli(Me.tag))
                txtDesconto.Text = FormatVal(!desconto_cli)
                lblRestante.caption = FormatVal(!limite_cli - CCur(lblUso.caption))
                If TiraNull(!uf_cid) <> "" Then lblUF.caption = TiraNull(!uf_cid)
                
                txtbxFax.Text = Format(!fax_cli)
                txtbxInscProdRural.Text = Format(!inscprodrur_cli)
                txtCidade.Text = Format(!nome_cid)
                txtCidade.tag = Format(!cod_cid)
                lblCodCid.caption = Format(!cod_cid)
                txtNumero.Text = Format(!NUMERO_cli, "")
                txtcomplemento.Text = Format(!COMPLEMENTO_cli, "")
                If Not IsNull(!ie_Cli) Then
                    txtIE.Text = !ie_Cli
                End If
                lblUF.caption = Format(!uf_cid)
                txtCidade.Text = Format(!nome_cid)
                
                'optnPessoa(0).Enabled = False
                'optnPessoa(1).Enabled = False
                txtFormasPag.Text = TiraNull(!nome_For)
                txtFormasPag.tag = TiraNull(!cod_for)
                txtbxMae.Text = Format(!mae_cli)
                txtbxNasc.Text = Format(!nasc_cli)
                
                txtLogradouro.Text = TiraNull(!Logradouro_cli)
                

                txtEmail.Text = TiraNull(!email_cli)
                
                
                
                P_LocalizaOperacao TiraNull(!indiedest_cli), cmbxIndIEDest
                
                P_LocalizaOperacao TiraNull(!status_cli), cmbxSituacao
                
                
                P_LocalizaOperacao TiraNull(!cfop_cli), cmbxCFOP
                P_LocalizaOperacao TiraNull(!tipoicms_cli), cmbxTipoICMS
                
                
                If strNivel_Usu <> "4" Then
                    txtbxLimite.Enabled = False
                End If
                
                .MoveNext
            
            .Close
        End With
        Exit Sub
trata_Erro:
        P_MostraErro
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
   
    If Me.tag = "" Then
        bCNPJ = True
    End If
    P_PreencheNFETipoICMS cmbxTipoICMS
    P_PreencheSpedCFOP cmbxCFOP
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If IsLoad(frmEdiClientes.Name) Then
        RefreshAdoPosition frmEdiClientes.ado
    End If
End Sub

Private Sub optnPessoa_Click(Index As Integer)
    If Index = 0 Then
        lblCPF.caption = "&CPF:"
        lblfantasia.caption = "&Chama-se:"
        
        txtbxCNPJ.MaxLength = 14
        txtbxNasc.Enabled = True
        txtbxMae.Enabled = True
        lblNasc.Enabled = True
        
        lblMae.Enabled = True
        lblRg.Enabled = True
        txtRg.Enabled = True
        txtIE.Enabled = False
        lblIE.Enabled = False
        txtbxMae.Enabled = True
        lblMae.Enabled = True
        
    Else
        lblCPF.caption = "&CNPJ:"
        lblfantasia.caption = "&Fantasia:"
        lblRg.Enabled = False
        txtRg.Enabled = False
        txtIE.Enabled = True
        lblIE.Enabled = True
        txtbxMae.Enabled = False
        lblMae.Enabled = False
        
        txtbxCNPJ.MaxLength = 18
        txtbxNasc.Enabled = False
        txtbxMae.Enabled = False
        lblNasc.Enabled = False
        lblMae.Enabled = False
        
    End If
    'txtbxCNPJ.Text = ""
End Sub
Private Sub pcCidade_Click()
    buscaCid
End Sub
Sub buscaCid()
    Dim c As Cidades.Cidade
    c = Cidades.F_BuscaCidades(txtCidade.Text)
    If c.cod_cid <> "" Then
        txtCidade.Text = c.nome_cid
        txtCidade.tag = c.cod_cid
        lblUF.caption = c.uf_cid
        lblCodCid.caption = c.cod_cid
        txtbxTel.SetFocus
        
    End If
End Sub

Private Sub pcFormasPag_Click()
    BuscaFormaPag
End Sub
Sub BuscaFormaPag()
    Dim p As FormasPag.FormasPag
    
    p = FormasPag.F_BuscaFormasPag("")
    If p.cod_for <> "" Then
        txtFormasPag.Text = p.nome_For
        txtFormasPag.tag = p.cod_for
    End If
    
End Sub
Private Sub txtbxCel_cli_GotFocus()
    ganhaFoco txtbxCel_cli
End Sub


Private Sub txtbxBairro_GotFocus()
    ganhaFoco txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    AjustaTexto txtbxBairro
End Sub


Private Sub txtbxCel_cli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCel_cli_LostFocus()
    txtbxCel_cli.Text = F_AjustaTelefone(txtbxCel_cli.Text)
End Sub

Private Sub txtbxCEganhafoco()
    ganhaFoco txtbxCEP
End Sub

Private Sub txtbxCEP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCEP_LostFocus()
    P_CEPLostFocus txtbxCEP
End Sub


Private Sub txtbxCidade_GotFocus()
    ganhaFoco txtCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    AjustaTexto txtCidade
End Sub

Private Sub txtbxCNPJ_GotFocus()
    ganhaFoco txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    
    
    If optnPessoa(0).Value Then
        If Not F_CPF(txtbxCNPJ.Text) Then
            MsgBox "CPF Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCPF txtbxCNPJ
    Else
        If Not F_CGC(txtbxCNPJ.Text) Then
            MsgBox "CNPJ Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCGC txtbxCNPJ
    End If
    
    
End Sub








Private Sub txtbxFax_GotFocus()
    ganhaFoco txtbxFax
End Sub

Private Sub txtbxFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    
End Sub

Private Sub txtbxFax_LostFocus()
    P_AjustaFone txtbxFax
End Sub

Private Sub txtDesconto_GotFocus()
    P_SelecTexto txtDesconto
End Sub

Private Sub txtDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtDesconto_LostFocus()
    P_ValorGotFocus txtDesconto
    If CCur(txtDesconto.Text) > 40 Then
        txtDesconto.SetFocus
        MsgBox "Desconto muito alto!", vbExclamation
    End If
    txtDesconto.Enabled = False
    
End Sub

Private Sub txtfantasia_GotFocus()
    P_SelecTexto txtfantasia
End Sub

Private Sub txtfantasia_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtfantasia_LostFocus()
    P_AjustaTexto txtfantasia
End Sub

Private Sub txtLogradouro_GotFocus()
    P_SelecTexto txtLogradouro
End Sub

Private Sub txtLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtLogradouro_LostFocus()
    P_AjustaTexto txtLogradouro
End Sub

Private Sub txtrg_GotFocus()
    ganhaFoco txtRg
End Sub

Private Sub txtrg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtrg_LostFocus()
    AjustaTexto txtRg
End Sub

Private Sub txtbxInscProdRural_GotFocus()
    ganhaFoco txtbxInscProdRural
End Sub

Private Sub txtbxInscProdRural_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxInscProdRural_LostFocus()
    AjustaTexto txtbxInscProdRural
End Sub

Private Sub txtbxLimite_GotFocus()
    ganhaFoco txtbxLimite
End Sub

Private Sub txtbxLimite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxLimite_LostFocus()
    PerdeFoco txtbxLimite
End Sub

Private Sub txtbxMae_GotFocus()
    ganhaFoco txtbxMae
End Sub

Private Sub txtbxMae_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        TrocaTabKey KeyCode
        
    End If
End Sub

Private Sub txtbxMae_LostFocus()
    AjustaTexto txtbxMae
    
End Sub

Private Sub txtbxNasc_GotFocus()
    ganhaFoco txtbxNasc

End Sub

Private Sub txtbxNasc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxNasc_LostFocus()
    P_AjustaData txtbxNasc
End Sub



Private Sub txtbxObs_GotFocus()
    ganhaFoco txtbxObs
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then
    '    trocatabkey keycode
    'End If
End Sub

Private Sub txtbxObs_LostFocus()
    AjustaTexto txtbxObs


End Sub

Private Sub txtbxRazao_GotFocus()
    ganhaFoco txtbxRazao
End Sub

Private Sub txtbxRazao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxRazao_LostFocus()
    AjustaTexto txtbxRazao
End Sub


Private Sub txtbxTel_GotFocus()
    ganhaFoco txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        
        TrocaTabKey KeyCode

        
    End If
    
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
    
End Sub

Private Sub txtCidade_GotFocus()
    ganhaFoco txtCidade
End Sub

Private Sub txtCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pcCidade_Click
    End If
End Sub

Private Sub txtIE_GotFocus()
    ganhaFoco txtIE
End Sub

Private Sub txtIE_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtIE_LostFocus()
    AjustaTexto txtIE
End Sub

Private Sub txtNumero_GotFocus()
    ganhaFoco txtNumero
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtNumero_LostFocus()
    AjustaTexto txtNumero
End Sub

Private Sub txtComplemento_GotFocus()
    ganhaFoco txtcomplemento
End Sub

Private Sub txtComplemento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtComplemento_LostFocus()
    AjustaTexto txtcomplemento
End Sub


