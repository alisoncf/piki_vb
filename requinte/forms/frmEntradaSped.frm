VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntradaSped 
   Caption         =   "Lançamento de Notas Fiscais"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   18795
   Icon            =   "frmEntradaSped.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   18795
   Begin VB.CommandButton cmndbtnEdit 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   17520
      TabIndex        =   82
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmndbtnAbreVenda 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   17520
      TabIndex        =   81
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdProduto 
      Caption         =   "Produto"
      Height          =   315
      Left            =   17520
      TabIndex        =   86
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox pcOperacao 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   18795
      TabIndex        =   83
      Top             =   0
      Width           =   18795
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1905
         ScaleWidth      =   4425
         TabIndex        =   170
         Top             =   0
         Width           =   4455
         Begin VB.PictureBox pcTop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            ScaleHeight     =   495
            ScaleWidth      =   4455
            TabIndex        =   175
            Top             =   1440
            Width           =   4455
            Begin VB.ComboBox cmbxTipo 
               Height          =   315
               ItemData        =   "frmEntradaSped.frx":08CA
               Left            =   120
               List            =   "frmEntradaSped.frx":08D4
               Style           =   2  'Dropdown List
               TabIndex        =   176
               Top             =   120
               Width           =   4230
            End
         End
         Begin VB.OptionButton optnEntrada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Entrada:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   174
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optnSaida 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Saída"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   173
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtEmpresa 
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
            TabIndex        =   172
            Tag             =   "01"
            Top             =   1080
            Width           =   3855
         End
         Begin VB.CommandButton cmdXml 
            Caption         =   "Importar XML"
            Height          =   735
            Left            =   1440
            Picture         =   "frmEntradaSped.frx":08FC
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   120
            Width           =   2895
         End
         Begin VB.CommandButton cmdLocalizaEmpresa 
            Height          =   285
            Left            =   3960
            Picture         =   "frmEntradaSped.frx":0D98
            Style           =   1  'Graphical
            TabIndex        =   177
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H00B3D7EA&
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa:"
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
            Height          =   285
            Left            =   120
            TabIndex        =   178
            Top             =   840
            Width           =   1065
         End
      End
      Begin TabDlg.SSTab stb 
         Height          =   1935
         Left            =   4680
         TabIndex        =   114
         Top             =   0
         Width           =   15405
         _ExtentX        =   27173
         _ExtentY        =   3413
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "F9 - Dados Gerais"
         TabPicture(0)   =   "frmEntradaSped.frx":1143
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "pcNota"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "F10 - Totais"
         TabPicture(1)   =   "frmEntradaSped.frx":115F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "pcImpostos"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Outras informações"
         TabPicture(2)   =   "frmEntradaSped.frx":117B
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.PictureBox pcImpostos 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1440
            Left            =   -74880
            ScaleHeight     =   1410
            ScaleWidth      =   12045
            TabIndex        =   138
            Top             =   360
            Width           =   12075
            Begin VB.TextBox txtCofinsSTNot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8760
               MaxLength       =   12
               TabIndex        =   153
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtPisStNot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7680
               MaxLength       =   12
               TabIndex        =   152
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtCofinsNot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6600
               MaxLength       =   12
               TabIndex        =   151
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtPisNot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5520
               MaxLength       =   12
               TabIndex        =   150
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtSeguro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Height          =   285
               Left            =   1200
               MaxLength       =   12
               TabIndex        =   149
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtVlIcmsStNot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3360
               MaxLength       =   12
               TabIndex        =   148
               Text            =   "0,00"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtBcIcmsStNot 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2280
               MaxLength       =   12
               TabIndex        =   147
               Text            =   "0,00"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtbxICMS_Not 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               MaxLength       =   12
               TabIndex        =   146
               Text            =   "0,00"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtbxBC_Not 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   12
               TabIndex        =   145
               Text            =   "0,00"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtbxFrete 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Height          =   285
               Left            =   120
               MaxLength       =   12
               TabIndex        =   144
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtbxValor_Not 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   10200
               MaxLength       =   12
               TabIndex        =   143
               Tag             =   "!"
               Text            =   "0,00"
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox txtbxIPI_Not 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4440
               MaxLength       =   12
               TabIndex        =   142
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtbxDepesas_Not 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Height          =   285
               Left            =   3360
               MaxLength       =   12
               TabIndex        =   141
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtbxDesconto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000C000&
               Height          =   285
               Left            =   2280
               MaxLength       =   12
               TabIndex        =   140
               Text            =   "0,00"
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtbxValorProd 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   10200
               MaxLength       =   12
               TabIndex        =   139
               Tag             =   "!"
               Text            =   "0,00"
               Top             =   240
               Width           =   1575
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
               TabIndex        =   169
               Top             =   720
               Width           =   885
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
               TabIndex        =   168
               Top             =   720
               Width           =   555
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
               TabIndex        =   167
               Top             =   720
               Width           =   585
            End
            Begin VB.Label Label56 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "PIS:"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5520
               TabIndex        =   166
               Top             =   720
               Width           =   375
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
               TabIndex        =   165
               Top             =   720
               Width           =   555
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
               TabIndex        =   164
               Top             =   120
               Width           =   930
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
               TabIndex        =   163
               Top             =   120
               Width           =   945
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
               TabIndex        =   162
               Top             =   120
               Width           =   675
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
               TabIndex        =   161
               Top             =   120
               Width           =   690
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
               TabIndex        =   160
               Top             =   720
               Width           =   405
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
               Left            =   9465
               TabIndex        =   159
               Top             =   0
               Width           =   135
            End
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "&Valor Total"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   10200
               TabIndex        =   158
               Tag             =   "!"
               Top             =   720
               Width           =   765
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "I&PI:"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4440
               TabIndex        =   157
               Top             =   720
               Width           =   375
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
               TabIndex        =   156
               Top             =   720
               Width           =   975
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
               TabIndex        =   155
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label34 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Produtos:"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   10200
               TabIndex        =   154
               Tag             =   "!"
               Top             =   0
               Width           =   1080
            End
         End
         Begin VB.PictureBox pcNota 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   120
            ScaleHeight     =   1425
            ScaleWidth      =   12585
            TabIndex        =   115
            Top             =   360
            Width           =   12615
            Begin VB.CommandButton cmndbtnLocaliza 
               Height          =   285
               Left            =   4320
               Picture         =   "frmEntradaSped.frx":1197
               Style           =   1  'Graphical
               TabIndex        =   126
               Top             =   360
               Width           =   255
            End
            Begin VB.TextBox txtForn 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   50
               TabIndex        =   125
               Top             =   360
               Width           =   4215
            End
            Begin VB.TextBox txtbxNum_Not 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4680
               MaxLength       =   12
               TabIndex        =   124
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtbxSerie_Not 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5520
               MaxLength       =   4
               TabIndex        =   123
               Top             =   360
               Width           =   495
            End
            Begin VB.ComboBox cmbxSituacao 
               Height          =   315
               ItemData        =   "frmEntradaSped.frx":1542
               Left            =   120
               List            =   "frmEntradaSped.frx":1544
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   975
               Width           =   3615
            End
            Begin VB.ComboBox cmbxModelo 
               Height          =   315
               ItemData        =   "frmEntradaSped.frx":1546
               Left            =   3840
               List            =   "frmEntradaSped.frx":1548
               Style           =   2  'Dropdown List
               TabIndex        =   121
               Top             =   960
               Width           =   1695
            End
            Begin VB.ComboBox cmbIndPag 
               Height          =   315
               ItemData        =   "frmEntradaSped.frx":154A
               Left            =   5640
               List            =   "frmEntradaSped.frx":1557
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   975
               Width           =   2535
            End
            Begin VB.ComboBox cmbIndFre 
               Height          =   315
               ItemData        =   "frmEntradaSped.frx":1581
               Left            =   8280
               List            =   "frmEntradaSped.frx":1591
               Style           =   2  'Dropdown List
               TabIndex        =   119
               Top             =   960
               Width           =   4095
            End
            Begin VB.TextBox txtbxEntrada_Not 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   118
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtbxEmissao_Not 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6120
               MaxLength       =   10
               TabIndex        =   117
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtChaveNfe 
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
               Left            =   8280
               TabIndex        =   116
               Top             =   360
               Width           =   4095
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Participante:"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Número:"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   4680
               TabIndex        =   136
               Top             =   120
               Width           =   600
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Série:"
               Height          =   195
               Left            =   5520
               TabIndex        =   135
               Top             =   120
               Width           =   405
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Situação NF:"
               Height          =   195
               Left            =   120
               TabIndex        =   134
               Top             =   735
               Width           =   930
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modelo NF:"
               Height          =   195
               Left            =   3840
               TabIndex        =   133
               Top             =   735
               Width           =   825
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Indicador Pgto"
               Height          =   195
               Left            =   5640
               TabIndex        =   132
               Top             =   735
               Width           =   1035
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Indicador Frete"
               Height          =   195
               Left            =   8280
               TabIndex        =   131
               Top             =   720
               Width           =   1065
            End
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "E&ntrada:"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   7200
               TabIndex        =   130
               Top             =   120
               Width           =   600
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "&Emissão: "
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6120
               TabIndex        =   129
               Top             =   120
               Width           =   675
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Chave NFE:"
               Height          =   195
               Left            =   8280
               TabIndex        =   128
               Top             =   120
               Width           =   870
            End
            Begin VB.Label lblUF 
               Caption         =   "UF"
               Height          =   255
               Left            =   4080
               TabIndex        =   127
               Top             =   120
               Width           =   495
            End
         End
      End
   End
   Begin VB.PictureBox pcProd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   18765
      TabIndex        =   1
      Top             =   6345
      Width           =   18795
      Begin VB.TextBox txtUnitConv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
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
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0,0000"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtOutDesp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   13080
         TabIndex        =   29
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.PictureBox pcVoper 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   10440
         ScaleHeight     =   825
         ScaleWidth      =   1545
         TabIndex        =   103
         Top             =   720
         Width           =   1575
         Begin VB.Label lblGastos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(0,0000)"
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
            TabIndex        =   109
            Top             =   600
            Width           =   720
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos"
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DIFAL:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   107
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblDifal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(0,0000)"
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
            TabIndex        =   106
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VOPER:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   105
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lblVoper 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(0,0000)"
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
            TabIndex        =   104
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.PictureBox pcCusto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6C1A6&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   12120
         ScaleHeight     =   1095
         ScaleWidth      =   5055
         TabIndex        =   87
         Top             =   720
         Width           =   5055
         Begin VB.CheckBox chkPrecoFuturo 
            Caption         =   "Preço Futuro: (F6)"
            Height          =   195
            Left            =   2520
            TabIndex        =   94
            Top             =   0
            Width           =   1575
         End
         Begin VB.TextBox txtCustoUnitario 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   93
            Text            =   "0,00"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtPreco 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2520
            TabIndex        =   95
            Text            =   "0,00"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtTx 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   91
            Text            =   "0,00"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtML 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   89
            Text            =   "0,00"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblVariacao 
            AutoSize        =   -1  'True
            Caption         =   "(0,0000)"
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
            Left            =   4200
            TabIndex        =   111
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo Atual"
            Height          =   195
            Index           =   3
            Left            =   1800
            TabIndex        =   102
            Top             =   600
            Width           =   810
         End
         Begin VB.Label lblCustoAtual 
            AutoSize        =   -1  'True
            Caption         =   "(0,0000)"
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
            Left            =   1800
            TabIndex        =   101
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblPcAtual 
            AutoSize        =   -1  'True
            Caption         =   "(0,0000)"
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
            TabIndex        =   100
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pç Atual"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   99
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sugerido"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   98
            Top             =   600
            Width           =   630
         End
         Begin VB.Label lblSug 
            AutoSize        =   -1  'True
            Caption         =   "(0,0000)"
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
            Left            =   960
            TabIndex        =   97
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblCusto 
            AutoSize        =   -1  'True
            Caption         =   "0,0000"
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
            Left            =   360
            TabIndex        =   96
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo:"
            Height          =   195
            Index           =   0
            Left            =   1560
            TabIndex        =   92
            Top             =   0
            Width           =   450
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TX(%):"
            Height          =   195
            Left            =   840
            TabIndex        =   90
            Top             =   0
            Width           =   465
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ML(%)"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.TextBox txtDescontoItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   12240
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtQtdConvert 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
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
         Left            =   9240
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtCean 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Editar"
         Height          =   375
         Left            =   16200
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtUndSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtQtdSaida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Height          =   285
         Left            =   5760
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtUnidEntr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox pcIpi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   8400
         ScaleHeight     =   825
         ScaleWidth      =   2025
         TabIndex        =   77
         Top             =   720
         Width           =   2055
         Begin VB.TextBox txtAliqIpiItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   49
            Text            =   "0,0"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtVlIpiItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   47
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtBcIpiItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label66 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Caption         =   "IPI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label65 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   225
            Left            =   1680
            TabIndex        =   48
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   840
            TabIndex        =   46
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox pcIcms 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   2145
         TabIndex        =   75
         Top             =   720
         Width           =   2175
         Begin VB.TextBox txtBcIcmsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtVlICMSItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   35
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtAliqIcmsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   37
            Text            =   "0,0"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   840
            TabIndex        =   34
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label76 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   225
            Left            =   1680
            TabIndex        =   36
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label75 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "ICMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.PictureBox pcCofins 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   6360
         ScaleHeight     =   825
         ScaleWidth      =   2025
         TabIndex        =   73
         Top             =   720
         Width           =   2055
         Begin VB.TextBox txtBCCofinsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtVlCofinsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   59
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtAliqCofinsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   61
            Text            =   "0,0"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            Caption         =   "COFINS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   74
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label64 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   225
            Left            =   1680
            TabIndex        =   60
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   840
            TabIndex        =   58
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox pcPis 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   4320
         ScaleHeight     =   825
         ScaleWidth      =   2025
         TabIndex        =   71
         Top             =   720
         Width           =   2055
         Begin VB.TextBox txtBcPisItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtVlPisItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   53
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtAliqPisItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   55
            Text            =   "0,0"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   840
            TabIndex        =   52
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label63 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   225
            Left            =   1680
            TabIndex        =   54
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "PIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.PictureBox pcIcmsSt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2280
         ScaleHeight     =   825
         ScaleWidth      =   2025
         TabIndex        =   69
         Top             =   720
         Width           =   2055
         Begin VB.TextBox txtAliqIcmsStItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   43
            Text            =   "0,0"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtVlIcmsStItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   41
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtBcIcmsStItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "ICMS ST"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   70
            Top             =   0
            Width           =   2175
         End
         Begin VB.Label Label60 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   225
            Left            =   1680
            TabIndex        =   42
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   960
            TabIndex        =   40
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox txtCst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtNCM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCfop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6960
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxCodBarra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdInserir 
         Caption         =   "F5 - &Inserir"
         Height          =   375
         Left            =   15120
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   285
         Left            =   13920
         TabIndex        =   31
         Text            =   "0,00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FEF4D8&
         Height          =   285
         Left            =   8400
         TabIndex        =   15
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AC67&
         Height          =   285
         Left            =   10080
         TabIndex        =   23
         Text            =   "0,0000"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   113
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Left            =   120
         TabIndex        =   112
         Tag             =   "B"
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit. Conv"
         Height          =   195
         Left            =   11160
         TabIndex        =   24
         Top             =   0
         Width           =   750
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Out. Desp"
         Height          =   195
         Left            =   13080
         TabIndex        =   28
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         Height          =   195
         Left            =   12240
         TabIndex        =   26
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Qtd Conv."
         Height          =   195
         Left            =   9240
         TabIndex        =   16
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM:"
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Und Saída"
         Height          =   195
         Left            =   4560
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Und."
         Height          =   195
         Left            =   7800
         TabIndex        =   12
         Top             =   0
         Width           =   345
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST:"
         Height          =   195
         Left            =   6240
         TabIndex        =   8
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EAN:"
         Height          =   195
         Left            =   3720
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   195
         Left            =   7080
         TabIndex        =   10
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Qtd Saída:"
         Height          =   195
         Left            =   5760
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Tag             =   "B"
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unitário"
         Height          =   195
         Left            =   10080
         TabIndex        =   22
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         Height          =   195
         Left            =   8400
         TabIndex        =   14
         Top             =   0
         Width           =   300
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   13920
         TabIndex        =   30
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   18765
      TabIndex        =   63
      Top             =   8280
      Width           =   18795
      Begin VB.CommandButton cmdPrecos 
         Caption         =   "F8 - Precificação"
         Height          =   615
         Left            =   12360
         TabIndex        =   110
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdVincular 
         Caption         =   "V - Vincular Produto"
         Height          =   615
         Left            =   4200
         TabIndex        =   85
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdExcluiItem 
         Caption         =   "F3 - Excluir Item"
         Height          =   615
         Left            =   8280
         TabIndex        =   80
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdEditItem 
         Caption         =   "F2 - Editar Item"
         Height          =   615
         Left            =   6240
         TabIndex        =   79
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "F12 - Sair"
         Height          =   615
         Left            =   14400
         TabIndex        =   67
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdConcluir 
         Caption         =   "F7 - Concluir NF"
         Height          =   615
         Left            =   10320
         TabIndex        =   66
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdNovoProd 
         Caption         =   "N - Novo Produto"
         Height          =   615
         Left            =   2160
         TabIndex        =   65
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdEditProd 
         Caption         =   "F4 - Edição de Produtos"
         Height          =   615
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   18765
      TabIndex        =   62
      Top             =   6330
      Width           =   18795
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEntradaSped.frx":1611
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   6165
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "codfab_pro"
         Caption         =   "Referência"
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
         DataField       =   "cod_Pro"
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
         DataField       =   "cEAN"
         Caption         =   "EAN"
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
         DataField       =   "descr_compl"
         Caption         =   "Descrição Original"
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
         DataField       =   "nome_pro"
         Caption         =   "Produto"
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
         DataField       =   "cr_pro"
         Caption         =   "C.A"
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
      BeginProperty Column06 
         DataField       =   "er_pro"
         Caption         =   "ER"
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
         DataField       =   "ef_pro"
         Caption         =   "EF"
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
         DataField       =   "pv_pro"
         Caption         =   "PV"
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
         DataField       =   "qtd"
         Caption         =   "Qtd"
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
         DataField       =   "qtd_convert"
         Caption         =   "Qtd.Conv"
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
         DataField       =   "Unit"
         Caption         =   "Unit"
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
         DataField       =   "vl_item"
         Caption         =   "Total Item"
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
      BeginProperty Column13 
         DataField       =   "custo"
         Caption         =   "Custo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "preco"
         Caption         =   "Preço"
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
      BeginProperty Column15 
         DataField       =   "alterapreco"
         Caption         =   "Altera Preço"
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
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3569,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3119,811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   375
      Left            =   2040
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnMenu 
      Caption         =   "menu"
      Begin VB.Menu mnBuscarProduto 
         Caption         =   "Buscar Produto"
      End
      Begin VB.Menu mnNovoProduto 
         Caption         =   "&Novo Produto"
      End
      Begin VB.Menu div 
         Caption         =   "-"
      End
      Begin VB.Menu mnEditPro 
         Caption         =   "Editar Produto"
      End
      Begin VB.Menu mnVinc 
         Caption         =   "&Vincular"
         Begin VB.Menu mnVincNome 
            Caption         =   "Pela Descrição"
         End
         Begin VB.Menu mnVincEan 
            Caption         =   "Pelo código de barras (EAN)"
         End
         Begin VB.Menu mnVincCodFab 
            Caption         =   "Pela referência (Código de Fábrica)"
         End
      End
      Begin VB.Menu mnDesvincular 
         Caption         =   "Excluir Vínculo"
      End
      Begin VB.Menu mnAtlz 
         Caption         =   "Atualizar"
         Begin VB.Menu mnAtlzNome 
            Caption         =   "&Nome"
         End
         Begin VB.Menu mnAtlzEan 
            Caption         =   "EAN (Código de Barras)"
         End
         Begin VB.Menu mnAtlzCodFab 
            Caption         =   "Referência/Cod. Fábrica"
         End
         Begin VB.Menu mnAtlzNcm 
            Caption         =   "NCM"
         End
      End
   End
End
Attribute VB_Name = "frmEntradaSped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean 'para saber se o frm já foi aberto
Public bVisualizar As Boolean ' para abrir a nota apenas para visualizar
Dim iTam As Integer 'tam do form
Dim sTableCab As String
Dim sTableItens As String
Public operacao As String
Public bNotaFechada As Boolean
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Function F_ValorContabilNota()
    F_ValorContabilNota = CCur(txtbxValorProd.Text) + F_Despesas
End Function
Private Function F_Despesas()
    F_Despesas = CCur(txtbxDepesas_Not.Text) + CCur(txtbxFrete.Text) - CCur(txtbxDesconto.Text)
End Function
Private Function F_GastosItem()
    If CCur(F_Qtd) > 0 Then
        F_GastosItem = (CCur(txtbxIpi.Text) / F_Qtd) + CCur(txtbxFrete.Text) - CCur(txtbxDesconto.Text)
    End If
End Function
Private Function F_ValorProdutos()
    F_ValorProdutos = CCur(txtbxValorProd.Text)
End Function
Function Voper() As Currency
    Dim indice As Double
    
    indice = (ado!vl_item) / ado!vProd
    
    
    If lblUF.caption <> "GO" Then
        Voper = (ado!vlitem - ado!vldesc) + (ado!VL_FRT * indice) + (ado!VL_OUT_DA * indice) + (ado!VL_SEG * indice)
    Else
        Voper = 0
    End If


End Function
Private Sub cmbxBusca_Click()
    lblExpr.caption = "&" & cmbxBusca.Text
    
    If cmbxBusca.ListIndex = 0 Then
        lblExpr.tag = "B"
    Else
        lblExpr.tag = Left(cmbxBusca.Text, 1)
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo trata_erro
        If IsNull(ado!cod_pro) Or ado!cod_pro = "" Then
            cmdVincular.caption = "V-Vincular Produto"
        Else
            cmdVincular.caption = "V-Excluir Vínculo"
        End If
    Exit Sub
trata_erro:
    
End Sub

Private Sub chkPrecoFuturo_Click()
    If txtPreco.Enabled = True Then
        txtPreco.Text = lblPcAtual.caption
        txtPreco.Enabled = False
    Else
        If lblSug.caption > lblPcAtual.caption Then
            txtPreco.Text = lblSug.caption
        Else
            txtPreco.Text = lblPcAtual.caption
        End If
        
        txtPreco.Enabled = True
        txtPreco.SetFocus
    End If
End Sub

Private Sub cmbIndFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbIndPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Sub confEmb()
    Dim strSQL As String
    Dim iCount As Integer
    Dim ok As Boolean
    On Error GoTo trata_erro
    ok = False
    If txtbxCodBarra.tag = "" Then
        Exit Sub
    End If
    If cmbxEmb.Text = "" Then
        cmbxEmb.Text = "1"
        Exit Sub
    End If
    
    If cmbxEmb.Text = "0" Then
        cmbxEmb.Text = "1"
        Exit Sub
    End If
    If cmbxEmb.Text = "1" Then
        ok = True
    Else
        For iCount = 0 To cmbxEmb.ListCount - 1
            If cmbxEmb.Text = cmbxEmb.List(iCount) Then
                ok = True
                Exit For
            End If
        Next
    End If

    If Not ok Then
        strSQL = "A embalagem " & cmbxEmb.Text & " não foi localizada para o produto!"
        strSQL = strSQL & " Deseja gravá-la agora para o produto?"
        iCount = MsgBox(strSQL, vbYesNoCancel + vbQuestion)
        If iCount = vbCancel Then
            cmbxEmb.SetFocus
            Exit Sub
        ElseIf iCount = vbYes Then
            cmbxEmb.tag = cmbxEmb.Text
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & txtbxCodBarra.tag & "',"
            strSQL = strSQL & cmbxEmb.Text & ")"
            Conn.Execute strSQL
            'P_PreencheEmbalagens txtbxCodBarra.Tag, cmbxEmb, True
            cmbxEmb.Text = cmbxEmb.tag
            
        End If
        
    End If
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmbxModelo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Op_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Sub P_Grava()
    Dim sTipo As String
    Dim strSQL As String
    On Error GoTo trata_erro
    sTipo = UCase(Left(lblTipo.caption, 1))
    If sTipo = "S" Then
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and tipo_not= 'S' and cod_Cli = '" & strCOD_CLI & "'") <> "" Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and cod_ter ='" & txtForn.tag & "' " _
                & " and tipo_not= 'E' and cod_Cli = '" & strCOD_CLI & "'") <> "" Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "insert into tab_notasfiscais(Cod_Cli,Cod_Ter,Tipo_Not,Num_Not,Serie_Not, "
    strSQL = strSQL & "valor_not,emissao_not,bc_not,bcs_not,icms_Not,ipi_not,despesas_not,"
    strSQL = strSQL & "operacao_not,entrada_not,tipodoc_not,codval_not)Values("
    strSQL = strSQL & "'" & strCOD_CLI & "', "
    strSQL = strSQL & "'" & txtForn.tag & "',"
    strSQL = strSQL & "'" & sTipo & "',"
    strSQL = strSQL & "'" & txtbxNum_Not & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxValor_Not.Text) & ","
    strSQL = strSQL & F_GravaData(txtbxEmissao_Not.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxBC_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBCS_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxICMS_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIPI_Not.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxDepesas_Not.Text) & ","

    strSQL = strSQL & "'" & Left(cmbxTipo_Op.Text, 5) & "', "
    strSQL = strSQL & F_GravaData(txtbxEntrada_Not.Text) & ","
    strSQL = strSQL & "'" & Left(cmbxtipo_Not.Text, 1) & "', "
    strSQL = strSQL & "'" & Left(cmbxCodVal_Not.Text, 1) & "') "
    Conn.Execute strSQL
    MsgBox "Nota incluída com sucesso!", vbInformation
    P_Clear
    txtForn.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Clear()
    txtForn.Text = ""
    txtForn.tag = ""
    txtbxNum_Not.Text = ""
    txtbxSerie_Not.Text = ""
    txtbxEmissao_Not.Text = ""
    cmbxTipo_Op.ListIndex = -1
    txtbxBC_Not.Text = "0,00"
    txtbxBCS_Not.Text = "0,00"
    txtbxICMS_Not.Text = "0,00"
    txtbxIPI_Not.Text = "0,00"
    txtbxDepesas_Not.Text = "0,00"
    txtbxEntrada_Not.Text = ""
    cmbxCodVal_Not.ListIndex = -1
    cmbxtipo_Not.ListIndex = -1
    txtbxValor_Not.Text = "0,00"
End Sub

Private Sub cmbxTipo_Op_LostFocus()
    stb.Tab = 1
    txtbxBC_Not.SetFocus
End Sub

Private Sub cmbxTipo_Click()
    If bNotaFechada = True Then
        
        Exit Sub
    End If
    
    
    SpedNF.AlteraTipoNota Me.tag, Left(cmbxTipo.Text, 2)
    
    
    
    
End Sub

Private Sub cmdConcluir_Click()
    
    If bNotaFechada Then
        MsgBox "Esta nota já está fechada!", vbExclamation
        Exit Sub
    
    End If
    If txtForn.tag = "" Then
        MsgBox "Informe o Participante!", vbExclamation
        Exit Sub
    End If
    
    If Not confereProdutos Then
        MsgBox "Os produtos não estão ok! Verifique se todos os produtos estão devidamente apontados.", vbExclamation
        Exit Sub
    End If
    
    If strEfetnota_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If txtEmpresa.tag = "" Then
        MsgBox "É necessário informar uma empresa!", vbExclamation
        Exit Sub
        
    End If
    
    If MsgBox("Confirma fechamento da NF?", vbYesNo + vbQuestion) = vbYes Then
        SpedNF.AlteraStatusNf Me.tag, "2"
        SpedNF.AlteraCodEmp Me.tag, txtEmpresa.tag
        If MsgBox("Nota fiscal fechada. Deseja realizar a efetivação da Nota?", vbYesNo + vbQuestion) = vbYes Then
            EfetivaSpedNf Me.tag
        End If
        
        
        Unload Me
    End If
    
    
End Sub
Function confereProdutos() As Boolean
    With ado
        If .RecordCount = 0 Then
            confereProdutos = False
            Exit Function
        End If
        
        .MoveFirst
        confereProdutos = True
        Do While Not .EOF
            If IsNull(!cod_pro) Then
                confereProdutos = False
            ElseIf !cod_pro = "" Then
                confereProdutos = False
            End If
            If !qtd_convert <= 0 Then
            
                MsgBox "O produto " & !Produto & " está com a quantidade convertida zerada. Corrija!", vbExclamation
                confereProdutos = False
            End If
            
            
            
            
            
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    
    
    
End Function

Private Sub cmdEdit_Click()
    If cmdInserir.tag <> "" Then
        cmdInserir.tag = ""
        cmdInserir.caption = "Inserir"
        cmdEdit.caption = "Editar"
    Else
        dtgrd_DblClick
    End If
End Sub

Private Sub cmdLocalizaEmpresa_Click()
    buscaEmpresa
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    
    If bNotaFechada Then Exit Sub
    
    e = sped_empresa.SelecionarSped_Empresa(txtEmpresa.Text)
    If e.codigo <> "" Then
        txtEmpresa.Text = e.nome
        txtEmpresa.tag = e.codigo
        AlteraEmpresa
    End If
    
End Sub
Sub AlteraEmpresa()
    SpedNF.AlteraCodEmp Me.tag, txtEmpresa.tag
        
    
    
    
End Sub

Private Sub cmdNovoProd_Click()
    If bNotaFechada Then Exit Sub
    
    NovoProd
    
End Sub

Private Sub cmdProduto_Click()
    On Error GoTo trata_erro
    Dim pos As Integer
    With ado
        
        If IsNull(!codigo) Or !codigo = "" Then
            pos = .AbsolutePosition - 1
            buscaProd !DESCR_COMPL
            P_buscaitens
            .Move pos
        Else
            pos = .AbsolutePosition - 1
            buscaProd ""
            P_buscaitens
            .Move pos
        End If
    End With
    Exit Sub
trata_erro:
    
End Sub
Sub buscaProd(expr As String)
    Dim p As Produto
    Dim i As Integer
    p = Produtos.F_BuscaProdSaida(expr)
    
    If p.cod_pro <> "" Then
        i = ado.AbsolutePosition
        AlteraCodProItemNf ado!id, p.cod_pro
        ado.Requery
        ado.Move i
    End If
    
    
    
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVincular_Click()
    If bNotaFechada Then Exit Sub
    Me.PopupMenu mnVinc
End Sub
Sub vincular(Descricao As String)
    Dim i As Integer
    With ado
        i = ado.AbsolutePosition
        If IsNull(!cod_pro) Or !cod_pro = "" Then
            buscaProd Descricao
        Else
            If MsgBox("Deseja excluir o vínculo deste produto?", vbYesNo + vbQuestion) = vbYes Then
            
                ZeraCodigoItemSpedNF !id
                P_buscaitens
                
            End If
            
            
            'buscaProd !cod_pro
        End If
        ado.Move i
    End With
End Sub
Sub Desvincular()
    Dim i As Integer
    With ado
        i = ado.AbsolutePosition
        
            If MsgBox("Deseja excluir o vínculo deste produto?", vbYesNo + vbQuestion) = vbYes Then
            
                ZeraCodigoItemSpedNF !id
                P_buscaitens
                
        
            
            
            'buscaProd !cod_pro
        End If
        ado.Move i
    End With
End Sub
    
    
    
Private Sub cmdXml_Click()
    On Error Resume Next
    frmSpedImportarXMLNF.Show 0, mdiFrm
    frmSpedImportarXMLNF.SetFocus
    
End Sub

Private Sub cmndbtnAbreVenda_Click()
    If bNotaFechada Then Exit Sub
    
    If Me.tag = "" Then
        If abreNota = True Then
            'HabCab False
        End If
    Else
        AlteraNota
        
    End If
End Sub
Function abreNota() As Boolean
    Dim strSQL As String
    Dim sSEQ As String
    If txtForn.tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtForn.SetFocus
        Exit Function
    End If
    If F_ValorContabilNota <> CCur(txtbxValor_Not.Text) Then
        MsgBox "OS valores não batem!", vbCritical
        'Exit Function
    End If
    If txtbxValor_Not.Text = "0,00" Then
        On Error Resume Next
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        stb.Tab = 1
        txtbxValor_Not.SetFocus
        Exit Function
    End If
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtForn.tag & "' " & _
        " and num_not='" & txtbxNum_Not.Text & "' and serie_not ='" & _
        txtbxSerie_Not.Text & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Function
    End If
    
    
    If MsgBox("Confirma Abertura da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Function
    End If
    
    sSEQ = SpedNF.GravaSpedNf("0", "1", txtForn.tag, Left(cmbxModelo.Text, 2), Left(cmbxSituacao.Text, 2), _
    txtbxSerie_Not.Text, txtbxNum_Not.Text, txtChaveNfe.Text, txtbxEmissao_Not.Text, txtbxEntrada_Not.Text, _
    txtbxValor_Not.Text, Left(cmbIndPag.Text, 1), txtbxDesconto.Text, "0", txtbxValorProd.Text, Left(cmbIndFre.Text, 1), _
    txtbxFrete.Text, txtSeguro.Text, txtbxDepesas_Not.Text, txtbxBC_Not.Text, txtbxICMS_Not.Text, _
     txtBcIcmsStNot.Text, txtVlIcmsStNot.Text, txtbxIPI_Not.Text, txtPisNot.Text, txtCofinsNot.Text, txtPisStNot.Text, _
     txtCofinsSTNot.Text, "1", Date, txtEmpresa.tag)
    
    
    
    
    Me.tag = sSEQ
    Form_Activate
    
    mostraControles
    
    
    
    
    End Function
Sub mostraControles()
    If CCur(txtbxICMS_Not.Text) > 0 Then
        pcIcms.Visible = True
    Else
        pcIcms.Visible = False
    End If
    If CCur(txtVlIcmsStNot.Text) > 0 Then
        pcIcmsSt.Visible = True
    Else
        pcIcmsSt.Visible = False
    End If
    If CCur(txtCofinsNot.Text) > 0 Then
        pcCofins.Visible = True
    Else
        pcCofins.Visible = False
    End If
    If CCur(txtbxIPI_Not.Text) > 0 Then
        pcIpi.Visible = True
    Else
        pcIpi.Visible = False
    End If
End Sub
Sub AlteraNota()
    Dim strSQL As String
    Dim sSEQ As String
    If txtForn.tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtForn.SetFocus
        Exit Sub
    End If
    'If F_ValorContabilNota <> CCur(txtbxValor_Not.Text) Then
    '    MsgBox "OS valores não batem!", vbCritical
    '    Exit Sub
    'End If
    If txtbxValor_Not.Text = "0,00" Then
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        txtbxValor_Not.SetFocus
        Exit Sub
    End If
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtForn.tag & "' " & _
        " and num_not='" & txtbxNum_Not.Text & "' and serie_not ='" & _
        txtbxSerie_Not.Text & "' and seq_not <>'" & Me.tag & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Sub
    End If
    
    
    SpedNF.AlteraSpedNf Me.tag, "0", "1", txtForn.tag, Left(cmbxModelo.Text, 2), Left(cmbxSituacao.Text, 2), _
    txtbxSerie_Not.Text, txtbxNum_Not.Text, txtChaveNfe.Text, txtbxEmissao_Not.Text, txtbxEntrada_Not.Text, _
    txtbxValor_Not.Text, Left(cmbIndPag.Text, 1), txtbxDesconto.Text, "0", txtbxValorProd.Text, Left(cmbIndFre.Text, 1), _
    txtbxFrete.Text, txtSeguro.Text, txtbxDepesas_Not.Text, txtbxBC_Not.Text, txtbxICMS_Not.Text, _
    txtBcIcmsStNot.Text, txtVlIcmsStNot.Text, txtbxIPI_Not.Text, txtPisNot.Text, txtCofinsNot.Text, txtPisStNot.Text, txtCofinsSTNot.Text, "1", Left(cmbxTipo.Text, 2)
    
    'frBotoes.Visible = True
    RecalculaCusto
    
    
End Sub
Sub RecalculaCusto()
On erro GoTo trata_erro
    With ado
        If .RecordCount > 0 Then
            
            .MoveFirst
        End If
        Do While Not .EOF
            .MoveNext
        Loop
        
    End With
trata_erro:
End Sub
Private Sub cmndbtnAddProd_Click()
    Dim p As Produto


    
    p = F_BuscaProdSaida(txtbxCodBarra.Text)
    
    If p.cod_pro = "" Then
        MsgBox "Produto não encontado!", vbExclamation
        Exit Sub
    End If
    
    If F_BuscaUmCampo("cod_pro", sTableItens, "num_not='" & txtbxNum_Not.Text & _
        "' and serie_not='" & txtbxSerie_Not.Text & "' and cod_for='" & _
            txtForn.tag & "' and cod_pro='" & p.cod_pro & "'") <> "" Then
        MsgBox "Este produto já se encontra lançado!", vbExclamation
        Exit Sub
    End If
    
    txtbxCodBarra.Text = p.nome_pro
    txtbxCodBarra.tag = p.cod_pro
    txtNCM.Text = p.ncm_pro
    txtNCM.tag = p.ncm_pro
    'P_PreencheEmbalagens p.cod_pro, cmbxEmb, True
    'cmbxEmb.ListIndex = 0
    'lblCodBarra.caption = p.codbarra_pro
'    txtbxCusto.Text = FormatVal(p.cr_pro)
'    txtbxPreco.Text = FormatVal(p.pv_pro)
'    txtbxLucro.Text = FormatVal(p.lucro_pro)
 '   txtbxER.Text = p.er_pro
    txtML.Text = p.lucro_pro
    txtCean.Text = p.codbarra_pro
    txtUnidEntr.Text = p.unidade_pro
    lblPcAtual.caption = FormatVal(p.pv_pro)
    lblCustoAtual.caption = FormatVal(p.cr_pro)
    mostraControles
    
    If txtNCM.Text = "" Then
        txtNCM.SetFocus
    Else
        txtCst.SetFocus
    End If
End Sub
Sub AddProd()
    Dim cFator As Currency 'para saber qual a relacao do produto pelo total
    If p.ipi > (CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption)) Then
        MsgBox "Esse valor para IPI não está correto!", vbExclamation
        Exit Sub
    ElseIf Round((p.qtd_pro * p.Preco) + CCur(lblValoLan.caption), 2) > CCur(txtbxValor_Not.Text) Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    ElseIf ((p.qtd_pro * p.Preco)) = 0 Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    End If
    
    
    cFator = (p.qtd_pro * p.Preco) / (CCur(txtbxValor_Not.Text) - CCur(lblGastos.caption) - CCur(txtbxIPI_Not.Text))
    
    P_InsereProd p.cod_pro, p.Preco, p.qtd_pro, p.ipi, cFator
    P_buscaitens
    soma
    lblMarca.caption = p.nome_mar
End Sub
Sub P_InsereProd(cod_pro As String, unit As Currency, qtd As Currency, _
                ipi As Currency, cFator As Currency)
    Dim strSQL As String
    Dim sSEQ As String
    Dim cCusto As Currency
    Dim cGastos As Currency
    sSEQ = F_BuscaUmCampo("Max(Seq_ite)", sTableItens, "Num_Not='" & txtbxNum_Not.Text & "'" & _
        " and cod_for='" & txtForn.tag & "'")
    If sSEQ = "" Then
        sSEQ = "1"
    Else
        sSEQ = Int(sSEQ) + 1
    End If
    
    cGastos = CCur(lblGastos.caption) * cFator
    cCusto = unit + (cGastos / qtd) + (ipi / qtd)
    
    strSQL = "insert into " & sTableItens & "(num_not,cod_for,serie_not,cod_pro,"
    strSQL = strSQL & "unit_ite,qtd_ite,ipi_ite,custo_ite,gastos_ite,seq_ite)values("
    strSQL = strSQL & "'" & txtbxNum_Not.Text & "',"
    strSQL = strSQL & "'" & txtForn.tag & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit) & ","
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(ipi) & ","
    
    strSQL = strSQL & F_GravaDec(cCusto) & ","
    strSQL = strSQL & F_GravaDec(cGastos) & ","
    strSQL = strSQL & sSEQ & ") "
    Conn.Execute strSQL
    
End Sub
Sub P_AlteraProd(cod_pro As String, unit As Currency, qtd As Currency, _
                ipi As Currency, cFator As Currency)
    Dim strSQL As String
    Dim cCusto As Currency
    Dim cGastos As Currency
    
    cGastos = CCur(lblGastos.caption) * cFator
    cCusto = unit + (cGastos / qtd) + (ipi / qtd)
    
    strSQL = "update " & sTableItens & " set "
    
    
    strSQL = strSQL & "unit_ite=" & F_GravaDec(unit) & ","
    strSQL = strSQL & "qtd_ite=" & F_GravaDec(qtd) & ","
    strSQL = strSQL & "ipi_ite=" & F_GravaDec(ipi) & ","
    strSQL = strSQL & "gastos_ite=" & F_GravaDec(cGastos) & ","
    strSQL = strSQL & "custo_ite =" & F_GravaDec(cCusto) & " "
    strSQL = strSQL & " where  "
    strSQL = strSQL & "num_not='" & txtbxNum_Not.Text & "' "
    strSQL = strSQL & "and cod_for='" & txtForn.tag & "' "
    strSQL = strSQL & "and serie_not='" & txtbxSerie_Not.Text & "' "
    strSQL = strSQL & "and cod_pro= '" & cod_pro & "' "
    
    
    Conn.Execute strSQL
    P_buscaitens
End Sub

Private Sub cmndbtnConfirmar_Click()
    
    If Me.tag = "" Then
        If txtForn.tag = "" Then
            MsgBox "Não há um cliente selecionado!", vbExclamation
            Exit Sub
        End If
        P_Grava
                
    Else
        P_Altera
    End If
End Sub
Sub P_Altera()
    Dim iPos As Integer
    Dim sTipo As String
    Dim strSQL As String
    On Error GoTo trata_erro
    sTipo = UCase(Left(lblTipo.caption, 1))
    If sTipo = "S" Then
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and tipo_not= 'S' and cod_Cli = '" & _
                strCOD_CLI) <> "" And _
                txtbxNum_Not.Text <> txtbxNum_Not.tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and cod_ter ='" & txtForn.tag & "' " _
                & " and tipo_not= 'E' and cod_Cli = '" & strCOD_CLI & "'") <> "" _
                And txtbxNum_Not.Text <> txtbxNum_Not.tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    End If
    strSQL = "update tab_NotasFiscais set "
    
    strSQL = strSQL & "num_not='" & txtbxNum_Not & "',"
    strSQL = strSQL & "serie_not='" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & "valor_Not=" & F_GravaDec(txtbxValor_Not.Text) & ","
    strSQL = strSQL & "emissao_not=" & F_GravaData(txtbxEmissao_Not.Text) & ", "
    strSQL = strSQL & "bc_not=" & F_GravaDec(txtbxBC_Not.Text) & ","
    strSQL = strSQL & "bcs_not=" & F_GravaDec(txtbxBCS_Not.Text) & ","
    strSQL = strSQL & "icms_not=" & F_GravaDec(txtbxICMS_Not.Text) & ","
    strSQL = strSQL & "ipi_Not=" & F_GravaDec(txtbxIPI_Not.Text) & ", "
    strSQL = strSQL & "despesas_not=" & F_GravaDec(txtbxDepesas_Not.Text) & ","

    strSQL = strSQL & "operacao_not='" & Left(cmbxTipo_Op.Text, 5) & "', "
    strSQL = strSQL & "entrada_not=" & F_GravaData(txtbxEntrada_Not.Text) & ","
    strSQL = strSQL & "tipodoc_not='" & Left(cmbxtipo_Not.Text, 1) & "', "
    strSQL = strSQL & "codval_not='" & Left(cmbxCodVal_Not.Text, 1) & "' "
    strSQL = strSQL & " where cod_cli='" & strCOD_CLI & "' "
    strSQL = strSQL & " and num_not ='" & txtbxNum_Not.tag & "' "
    strSQL = strSQL & " and tipo_not='" & sTipo & "' "
    strSQL = strSQL & " and cod_ter='" & Me.tag & "' "
    strSQL = strSQL & " and serie_not='" & txtbxSerie_Not.tag & "'"
    Conn.Execute strSQL
    MsgBox "Nota Alterada com sucesso!", vbInformation
    If frmEditaMov.Visible = True Then
        iPos = frmEditaMov.adorMov.AbsolutePosition
        frmEditaMov.adorMov.Requery
        frmEditaMov.adorMov.Move iPos
    End If
    Unload Me
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

 Sub Cancelar()
    If MsgBox("Deseja cancelar?", vbYesNo + vbCritical) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub editaItem(Optional opcao As Integer = 0)
        Dim cValor_Max As Currency
        Dim cIpi_Max As Currency
        Dim strSQL As String
        Dim cFator As Currency
        Dim Preco As Currency
        Dim qtd_pro As Currency
        'On Error Resume Next
        If dtgrd.row = -1 Then
            Exit Sub
        End If
        With ado
            
            
            lblDescricao.caption = !DESCR_COMPL
            txtbxCodBarra.Text = Format(!Produto)
            txtbxCodBarra.tag = !cod_pro
            txtbxQtd.Text = !qtd
            txtQtdConvert.Text = !qtd_convert
            txtbxUnit.Text = !vUnCom
            txtbxTotal.Text = FormatVal(!vl_item)
            txtDescontoItem.Text = !VL_DESC
            
            cmdInserir.tag = !id
            txtVlIpiItem.Text = !vl_pis
            'txtbxSujerido.Text = !sujerido_ite
            
            
            txtBcIcmsItem.Text = !VL_BC_ICMS
            txtVlICMSItem.Text = !vl_icms
            txtAliqIcmsItem.Text = !aliq_icms
            txtVlCofinsItem.Text = !vl_cofins
            txtVlIcmsStItem.Text = !VL_ICMS_ST
'            txtAliqCofinsItem.Text = !vl_aliq_cofins
            txtBCCofinsItem.Text = !vl_bc_cofins
            
            txtVlIpiItem.Text = !VL_IPI
            
            txtVlPisItem.Text = !vl_pis
            txtAliqPisItem.Text = !aliq_pis
            txtBcPisItem.Text = !vl_bc_pis
            
            txtCean.Text = !cEAN
            txtNCM.Text = Format(!ncm, "0#######")
            txtCst.Text = !cst_icms
            txtCfop.Text = !CFOP
            
            txtML.Text = Format(!lucro_pro)
            txtTx.Text = crcTxFederativa_Par + crcTxEsdadual_Par
            CalculaPreco
            cmdInserir.caption = "Alterar"
            cmdEditItem.caption = "Cancelar"
            txtCustoUnitario.Text = FormatVal(!custo)
            cmdEdit.tag = !NUM_ITEM
            
            txtPreco.Text = FormatVal(!Preco)
            
            lblDifal.caption = RetornaDifal(lblUF_Ter.caption, !aliq_icms, !vl_item, !VL_DESC, "0", "0", "0")
            
            lblPcAtual.caption = FormatVal(!pv_pro)
            lblCustoAtual.caption = FormatVal(!cr_pro)
            
            
            
            If CCur(txtPreco.Text) < CCur(lblPcAtual.caption) Then
                'txtPreco.Text = lblPcAtual.caption
            End If
            If !alterapreco = "S" Then
                chkPrecoFuturo.Value = 1
            Else
                chkPrecoFuturo.Value = 0
            End If
            
            P_Unit
            P_CalcItem
            
            txtbxCodBarra.SetFocus
        End With
         
         
         
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Dim strSQL As String
    'On Error GoTo trata_erro
    If MsgBox("Confirma exlusão do item " & ado!Produto & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from " & sTableItens & " where id='" & ado!id & "' "

        
        Conn.Execute strSQL
        
        
        
        
    End If
    P_buscaitens
    
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnForn_Click()

End Sub

Private Sub cmdInserir_Click()
    If bNotaFechada Then
        MsgBox "A nota não pode ser alterada!", vbExclamation
        Exit Sub
    End If
    If Me.tag = "" Then
        If MsgBox("A NF ainda não está aberta. Deseja abrir?", vbYesNo + vbQuestion) = vbYes Then
            If Not abreNota Then
                Exit Sub
            End If
        End If
    End If
    
    
    If txtbxCodBarra.tag = "" Then
        MsgBox "Informe o Nome do Produto e tecle <ENTER>!", vbCritical
        P_buscaitens
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
'    If F_Qtd = 0 Then
'        MsgBox "Informe o tipo de embalagem e quantidade", vbExclamation
'        cmbxEmb.SetFocus
'        Exit Sub
'    End If
    If cmdInserir.tag = "" Then
        insere
        P_buscaitens
    Else
        altera
    End If
    
End Sub

Private Sub cmndbtnEdit_Click()
    HabCab True
    HabItens False
    cmndbtnAbreVenda.Visible = True
End Sub
Sub HabItens(opcao As Boolean)
    
    
End Sub
Sub HabCab(opcao As Boolean)
    pcTop.Enabled = opcao
    txtForn.Enabled = opcao
    'cmbxTipo_Op.Enabled = opcao
    txtbxNum_Not.Enabled = opcao
    txtbxSerie_Not.Enabled = opcao
    txtbxBC_Not.Enabled = opcao
    txtbxICMS_Not.Enabled = opcao
    txtbxIPI_Not.Enabled = opcao
    txtbxValor_Not.Enabled = opcao
    txtbxEmissao_Not.Enabled = opcao
    txtbxDesconto.Enabled = opcao
    txtbxFrete.Enabled = opcao
    txtbxDepesas_Not.Enabled = opcao
    txtbxValorProd.Enabled = opcao
    
End Sub
Private Sub cmndbtnLocaliza_Click()
    Dim f As participante
    f = SpedParticipante.F_BuscaParticipante(txtForn.Text)
    If f.codigo <> "" Then
            txtForn.Text = f.nome
            lblUF.caption = f.UF
            If txtForn.tag <> f.codigo And Me.tag <> "" Then
                If MsgBox("Confirma alteração do participante?", vbYesNo + vbQuestion) = vbYes Then
                    SpedNF.AlteraParticipante Me.tag, f.codigo
                End If
            End If
            txtForn.tag = f.codigo
            
            txtbxNum_Not.SetFocus
    End If
    
    
End Sub

Private Sub cmndbtnNovoProd_Click()
    Dim fr As New frmCadPro
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    Unload Me
    If IsLoad(frmEdiEntr.Name) Then
        
    End If
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL  As String
    Dim cNovoPreco As Currency
    Dim cLucro As Currency

    sDif = CCur(lblValoLan.caption) - CCur(txtbxValorProd.Text)
    If sDif < 0 Then
        sDif = sDif * -1
    End If
    If sDif > 0.4 Then
        MsgBox "Os valores não estão completamente lançados!", vbExclamation
        Exit Sub
    ElseIf lblIpi_Lan.caption <> txtbxIPI_Not.Text Then
        MsgBox "O valor do IPI não está completamente lançado!", vbExclamation
        Exit Sub
    ElseIf adc.Recordset.RecordCount = 0 Then
        MsgBox "Não é possível concluir uma nota sem itens!", vbExclamation
        Exit Sub
    End If
    
    
    If lblIpi_Lan.caption <> txtbxIPI_Not.Text Then
        MsgBox "O valor do IPI não está completamente lançado!", vbExclamation
        Exit Sub
    ElseIf adc.Recordset.RecordCount = 0 Then
        MsgBox "Não é possível concluir uma nota sem itens!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma efetivação da Nota Fiscal? ", vbYesNo + vbQuestion) = vbYes Then
        With adc.Recordset
'            .MoveFirst
            strSQL = "update " & sTableCab & " set status_not = '2' "
            strSQL = strSQL & " where Seq_not='" & Me.tag & "' "
            
            Conn.Execute strSQL
            
        End With
        MsgBox "Nota Fiscal concluída!", vbInformation
'        If IsLoad(frmEdiEntr.Name) Then
'            frmEdiEntr.adc.Refresh
'        End If
        Unload Me
    End If
    
    
    
End Sub



Private Sub dtgrd_DblClick()
    On Error GoTo trata_erro
        cmdInserir.tag = ado!id
        editaItem
        On Error Resume Next
        If txtbxCodBarra.tag <> "" Then
            txtbxQtd.SetFocus
        Else
            txtbxCodBarra.SetFocus
        End If
        Exit Sub
trata_erro:
        
 '   End If
End Sub
Sub CalculaPreco()
    Dim Preco As Currency
    Dim custo As Currency
    Dim ml As Currency
    Dim tx As Currency
    On Error Resume Next
    custo = CCur(txtCustoUnitario.Text)
    ml = CCur(txtML.Text) / 100
    tx = CCur(txtTx.Text) / 100
    Preco = custo + (custo * (ml + tx))
    
    
    If Preco > CCur(lblPcAtual.caption) Then
    
       ' txtPreco.Text = FormatVal(Preco)
    End If
    
    lblSug.caption = FormatVal(Preco)
    
    
    lblVariacao.caption = FormatVal((CCur(txtPreco.Text) - custo) / custo * 100, "###,##0.0") & "%"
    
End Sub
Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        editaItem
    End If
    If KeyCode = vbKeyDelete Then
        cmndbtnExc_Prod_Click
    End If
    If KeyCode = vbKeyP Then
        ControlarPrecoItem ado!Id_nf, ado!id
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
    
End Sub

Private Sub adc_Reposition()
    With adc.Recordset
        If .RecordCount > 0 Then
            'cmndbtnEditProd.Enabled = True
            'cmndbtnExc_Prod.Enabled = True
            If Not .EOF Then

            End If
        Else
            'cmndbtnEditProd.Enabled = False
            'cmndbtnExc_Prod.Enabled = False
'            cmndbtnForn.Visible = False
        End If
    End With
End Sub

Sub soma()
    Dim iPos As Integer
    Dim iCount As Integer
    Dim cValor As Currency
    Dim cIPI As Currency
    Dim cFrete As Currency
    Dim cDesp As Currency
    Dim total As Currency
    With adc.Recordset
    Err.Clear
    On Error GoTo trata_erro
        
        If .RecordCount = 0 Then
            lblIpi_Lan.caption = "0,00"
            lblValoLan.caption = "0,00"
            lblItens.caption = "0"
            
            
            lblTotalProd.caption = "0,0"
            
            
            lblGastos.tag = "0,00"
            lblRest.caption = "0,00"
            Exit Sub
        End If
        iPos = .AbsolutePosition
        
        .MoveFirst
        Do While Not .EOF
            cValor = cValor + !total
            cIPI = cIPI + !ipi_ite
            iCount = iCount + !qtd_ite
            cDesp = cDesp + !gastos_ite
            total = total + !vl_item
            .MoveNext
        Loop
            lblIpi_Lan.caption = FormatVal(cIPI)
            lblValoLan.caption = FormatVal(cValor)
            lblTotalProd.caption = FormatVal(cValor)
            'lblItens.caption = .RecordCount
            lblItens.caption = Format(iCount)
            lblGastos.tag = cDesp
            lblRest.caption = F_Restante
            lblTotal.caption = FormatVal(total)
            .MoveFirst
            .Move iPos
    End With
    If CCur(txtbxBC_Not.Text) > 0 Then
        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.Text) / CCur(txtbxBC_Not.Text) * 100, 2)) & " %"
    End If
trata_erro:
End Sub

Sub insere()
    Dim CST_IPI As String
    Dim sItem As String
    Dim custo As Currency
    sItem = MaxKeySQL("sped_nf_itens", "num_item", "id=" & Me.tag, "")
    'ind_apur = 0 mensal 1 decendial
    'ind_mov = movimentação física do item/produto ==> 0 - sim 1 - não
    
    
    
    If Not SpedNF.InsereItemSpedNF(Me.tag, sItem, txtbxCodBarra.tag, txtbxCodBarra.tag, _
        txtbxCodBarra.Text, txtbxQtd.Text, txtUnidEntr.Text, txtQtdSaida.Text, txtUndSaida.Text, _
         CStr(CCur(txtbxUnit.Text) * CCur(txtQtdConvert.Text)), txtbxUnit.Text, txtDescontoItem.Text, "1", txtCst.Text, txtCfop.Text, "", txtBcIcmsItem.Text, _
         txtAliqIcmsItem.Text, txtVlICMSItem.Text, txtBcIcmsStItem.Text, txtAliqIcmsStItem.Text, txtVlIcmsStItem.Text, _
         "0", CST_IPI, "", txtBcIpiItem.Text, txtAliqIpiItem.Text, txtVlIpiItem.Text, "98", txtBcPisItem.Text, _
        txtAliqPisItem.Text, txtbxQtd.Text, "0", txtVlPisItem.Text, "98", txtBCCofinsItem.Text, _
        txtAliqCofinsItem.Text, txtbxQtd.Text, "0", txtVlCofinsItem.Text, "", txtCean.Text, txtNCM.Text, _
        txtQtdConvert.Text, lblCusto.caption, txtPreco.Text, IIf(chkPrecoFuturo.Value = 1, "S", "N")) Then
    End If
    
    
   Exit Sub
    
    
    
    
    
    
    
    
    Dim strSQL As String
    Dim sSEQ As String
    strSQL = F_BuscaUmCampo("seq_ite", sTableItens, "cod_pro='" & txtbxCodBarra.tag & "'  and seq_not='" & Me.tag & "' ")
    
    
    If strSQL <> "" Then
        MsgBox "O Item já se encontra lançado como Item: " & strSQL, vbCritical
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    sSEQ = MaxKeySQL(sTableItens, "id", "id_doc='" & Me.tag & "'", "")
    strSQL = "Insert into " & sTableItens & "(seq_not,seq_ite,cod_pro,unit_ite,qtd_ite,qtd_emb,ipi_ite,"
    strSQL = strSQL & "gastos_ite,custo_ite,sujerido_ite,precofuturo_ite,lucrosuj_ite,"
    strSQL = strSQL & "lucroliq_ite,custosai_ite,icms_ite)values("
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "'" & sSEQ & "',"
    strSQL = strSQL & "'" & txtbxCodBarra.tag & "',"
    strSQL = strSQL & F_GravaDec(txtbxUnit.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxQtd.Text) & ","
    strSQL = strSQL & F_GravaDec(cmbxEmb.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIpi.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxGastos.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxNovoCusto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxSujerido.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxPrecoFuturo.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxLucro.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxLucroLiq.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxCustoSaida.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIcms.Text) & ")"
    'conn.Execute strSQL
    P_buscaitens
    P_LimpaItem
    
End Sub
Sub RecalculaItens()
    
    With adc.Recordset
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            editaItem
            altera 1  'não dá refresh
            .MoveNext
        Loop
        P_buscaitens
    End With
End Sub
Sub P_LimpaItem()
    On Error Resume Next
    txtbxCodBarra.Text = ""
    txtbxCodBarra.tag = ""
    txtbxCodBarra.SetFocus
    cmndbtnConfirma.tag = ""
    cmdEdit.caption = "Editar"
    cmdEdit.tag = ""
    cmdInserir.tag = ""
    cmdInserir.caption = "Inserir"
End Sub
Sub altera(Optional opcao As Integer = 0)
    Dim strSQL As String
    Dim sSEQ As String
    Dim i As Integer
    Dim sItem As String
    SpedNF.AlteraItemSpedNF cmdInserir.tag, cmdEdit.tag, txtbxCodBarra.tag, "", lblDescricao.caption, txtbxQtd.Text, txtUnidEntr.Text, txtQtdSaida.Text, txtUndSaida.Text, _
    txtbxUnit.Text, txtDescontoItem.Text, "1", txtCst.Text, txtCfop.Text, "", txtBcIcmsItem.Text, _
    txtAliqIcmsItem.Text, txtVlICMSItem.Text, txtBcIcmsStItem.Text, txtAliqIcmsStItem.Text, txtVlIcmsStItem.Text, _
    "0", CStr(CST_IPI), "", txtBcIpiItem.Text, txtAliqIpiItem.Text, txtVlIpiItem.Text, "98", txtBcPisItem.Text, _
    txtAliqPisItem.Text, txtbxQtd.Text, "0", txtVlPisItem.Text, "98", txtBCCofinsItem.Text, _
    txtAliqCofinsItem.Text, txtbxQtd.Text, "0", txtVlCofinsItem.Text, txtVlCofinsItem.Text, _
    txtQtdConvert.Text, txtCustoUnitario.Text, txtPreco.Text, IIf(chkPrecoFuturo.Value = 1, "S", "N")
        
        RefreshAdoPosition ado
    
        P_LimpaItem
    
End Sub

Private Sub dtgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnMenu
    End If
End Sub

Private Sub Form_Activate()
    If bNotaFechada Then
        mnMenu.Enabled = False
        cmndbtnAbreVenda.Enabled = False
        cmndbtnLocaliza.Enabled = False
        cmdVincular.Enabled = False
    End If
    
    
    
    If bActive = True Then Exit Sub
    
    
    stb.Tab = 0
    
    
    pcTop.BackColor = vbGreen
        sTableCab = "sped_nf"
        sTableItens = "sped_nf_itens"
        
    
    If Me.tag <> "" Then
    
        carrega
        P_buscaitens
        
    End If
    
    If bVisualizar = True Then
        desabilitaForm
    End If
        
    
    bActive = True
    'MudaAparencia Me
End Sub
Sub desabilitaForm()
    pcOperacao.Enabled = False
    cmndbtnAbreVenda.Enabled = False
    cmndbtnEdit.Enabled = False
    pcProd.Visible = False
    dtgrd.BackColor = vbInactiveBorder
    cmdConcluir.Enabled = False
    cmdNovoProd.Enabled = False
    cmdEditProd.Enabled = False
    cmdEditItem.Enabled = False
    cmdVincular.Enabled = False
    cmdExcluiItem.Enabled = False
    pcImpostos.Enabled = False
    pcNota.Enabled = False
    Me.BackColor = vbInactiveCaptionText
    mnMenu.Enabled = False
    Me.caption = "NFe " & Me.tag & "(APENAS VISUALIZAÇÃO)"
    
End Sub
Sub carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select tf.nome,es.nome as nome_emp, tf.uf, es.apelido, "
    strSQL = strSQL & " te.* "
    strSQL = strSQL & " from sped_nf as te left join sped_participante as tf  "
    strSQL = strSQL & " on tf.codigo=te.COD_PART "
    strSQL = strSQL & " left join sped_empresa as es on te.cod_emp=es.codigo "
    strSQL = strSQL & " where id=" & Me.tag
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            
            txtForn.Text = Format(!nome)
            lblUF.caption = TiraNull(!UF)
            txtForn.tag = !cod_part
            txtbxNum_Not.Text = !num_doc
            txtbxSerie_Not.Text = !SER
            txtbxEmissao_Not.Text = Format(!dt_doc, "dd/mm/yyyy")
            txtbxEntrada_Not.Text = Format(!DT_E_S, "dd/mm/yyyy")
            txtChaveNfe.Text = !chv_nfe
            P_LocalizaOperacao !cod_sit, cmbxSituacao
            P_LocalizaOperacao !cod_mod, cmbxModelo
            P_LocalizaOperacao !IND_PGTO, cmbIndPag
            P_LocalizaOperacao !IND_FRT, cmbIndFre
            
            txtbxBC_Not.Text = FormatVal(!VL_BC_ICMS)
            txtbxBC_Not.Text = FormatVal(!vl_icms)
            txtBcIcmsStNot.Text = FormatVal(!VL_BC_ICMS_ST)
            txtVlIcmsStNot.Text = FormatVal(!VL_ICMS_ST)
            txtCofinsNot.Text = FormatVal(!vl_cofins)
            txtbxValor_Not.Text = FormatVal(!vl_doc)
            txtbxValorProd.Text = FormatVal(!VL_MERC)
            txtbxFrete.Text = FormatVal(!VL_FRT)
            txtSeguro.Text = FormatVal(!VL_SEG)
            txtbxDesconto.Text = FormatVal(!VL_DESC)
            txtbxDepesas_Not.Text = FormatVal(!VL_OUT_DA)
            txtbxIPI_Not.Text = FormatVal(!VL_IPI)
            txtPisNot.Text = FormatVal(!vl_pis)
            txtCofinsNot.Text = FormatVal(!vl_cofins)
            txtCofinsSTNot.Text = FormatVal(!VL_COFINS_ST)
            txtPisStNot.Text = FormatVal(!VL_PIS_ST)
            txtbxICMS_Not.Text = FormatVal(!vl_icms)
            
            txtEmpresa.tag = !cod_emp
            txtEmpresa.Text = Format(!apelido)
            If txtEmpresa.tag = "" Then
                txtEmpresa.Text = F_BuscaUmCampo("apelido", "sped_empresa", "codigo='01'")
                txtEmpresa.tag = "01"
            End If
            
            P_LocalizaOperacao !tipo_entrada, cmbxTipo
            
            
        End If
    End With
    
    
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyI And Shift = 1 Then
        P_InfPro adc.Recordset!cod_pro
    
    End If
    If Shift = 2 And KeyCode = vbKeyC Then
        'optnC.Value = True
        'txtbxCodBarra.SetFocus
    End If
    If Shift = 2 And KeyCode = vbKeyN Then
        optnN.Value = True
        txtbxCodBarra.SetFocus
    End If
    If Shift = 2 And KeyCode = vbKeyB Then
        optnBar.Value = True
        txtbxCodBarra.SetFocus
    End If
    If KeyCode = vbKeyF2 Then
        editaItem
    ElseIf KeyCode = vbKeyF3 Then
        cmndbtnExc_Prod_Click
    ElseIf KeyCode = vbKeyF4 Then
        Me.WindowState = vbMinimized
            On Error Resume Next
            frmEdiPro.Show 0, mdiFrm
        
        frmEdiPro.SetFocus
    ElseIf KeyCode = vbKeyF5 Then
        On Error Resume Next
        Me.WindowState = vbMinimized
        Dim fr As New frmCadPro
        fr.Show 0, mdiFrm
    ElseIf KeyCode = vbKeyF6 Then
        chkPrecoFuturo_Click
    ElseIf KeyCode = vbKeyF7 Then
        cmdConcluir_Click
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    End If

End Sub

Sub P_buscaitens()
    Dim strSQL As String
    Dim i As Integer
        
    strSQL = "select tp.cod_pro as codigo,nome_pro, tp.er_pro,tp.ef_pro,tp.cr_pro,tp.pv_pro, tp.codfab_pro,tp.nome_pro as produto,"
    strSQL = strSQL & "vl_item, vUnCom as unit, nome_mar as marca, nome_cat as categoria, tp.ncm_pro as ncm, ti.*,"
    
    strSQL = strSQL & " tp.cr_pro as cr, tp.pv_pro as pv,pv_pro,cr_pro,tp.lucro_pro "
    strSQL = strSQL & " from sped_nf_itens as ti left join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " left join " & TAB_MARCAS & "  as tm on tm.cod_mar=tp.cod_mar"
    strSQL = strSQL & " left join " & tab_categorias & " as tc on tc.cod_cat=tp.cod_cat "
    strSQL = strSQL & " left join " & TAB_SUBCAT & " as ss on tp.cod_sub=ss.cod_sub "
    strSQL = strSQL & " where ti.id_nf ='" & Me.tag & "' "
    
    
    
       
    
        strSQL = strSQL & " order by ti.id "
    
    
    abreObjetoAdo ado, strSQL, dtgrd
    
    
    
   
    
    
    soma
    
    
End Sub

Private Sub Form_Load()
    resizeForm Me
    txtbxEntrada_Not.Text = Date
    
    Set ado = New ADODB.Recordset
    
    
    
    stb.Tab = 0
    'P_PreencheOperacoes cmbxTipo_Op, "E"
    
    P_PreencheSituacaoDoc cmbxSituacao
    P_PreencheModeloNFTab411 cmbxModelo
    
bActive = False
    'cmbxBusca.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade dtgrd, Me, pcProd
    stb.width = dtgrd.width
End Sub

Private Sub Option2_Click()
    
End Sub

Private Sub optnBar_Click()
    lblExpr.caption = "C&od. Barra:"
    lblExpr.tag = "B"
End Sub

Private Sub optnC_Click()
        lblExpr.caption = "C&od. Prod.:"
        lblExpr.tag = "C"
End Sub



Private Sub optnN_Click()
    lblExpr.caption = "N&ome Prod.:"
    lblExpr.tag = "N"
End Sub




Private Sub lblConclui_Click()
    cmndbtnSalvar_Click
End Sub

Private Sub optnCod_Click()
    P_buscaitens
End Sub

Private Sub optnDesc_Click()
    P_buscaitens
End Sub

Private Sub optnSeq_Click()
    P_buscaitens
End Sub




Private Sub mnAtlzNome_Click()
    'Atualiza
End Sub

Private Sub mnBuscarProduto_Click()
    cmdProduto_Click
End Sub

Private Sub mnDesvincular_Click()
    Desvincular
End Sub

Private Sub mnEditPro_Click()
    With ado
        frmCadPro.tag = !cod_pro
        frmCadPro.Show 1
    End With
End Sub

Private Sub mnNovoProduto_Click()
    NovoProd
End Sub
Sub NovoProd()
        On Error GoTo trata_erro
    If bNotaFechada Then Exit Sub
    
    
    With ado
        If .RecordCount > 0 Then
            frmCadPro.txtbxCodBarra.Text = ""
            frmCadPro.txtbxRef.Text = !cod_item
            frmCadPro.txtbxCr.Text = !vl_item
            frmCadPro.txtbxNome_pro.Text = !DESCR_COMPL
            frmCadPro.txtbxUnidade.Text = !unid
            frmCadPro.txtEan.Text = !cEAN
            
            'P_LocalizaOperacao !cst_icms, frmCadPro.cmbxNFECst
            
        End If
            frmCadPro.Show 1
    End With
    Exit Sub
    
trata_erro:
    MostraErro

End Sub
Private Sub mnVincCodFab_Click()
    On Error Resume Next
    vincular ado!cod_item
End Sub

Private Sub mnVincEan_Click()
    On Error Resume Next
    vincular ado!cEAN
End Sub

Private Sub mnVincNome_Click()
    On Error Resume Next
    vincular ado!DESCR_COMPL
End Sub

Private Sub txtAliqCofinsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdInserir.SetFocus
    End If
    
End Sub

Private Sub txtAliqIcmsItem_GotFocus()
    GanhaFoco txtAliqIcmsItem
End Sub

Private Sub txtAliqIcmsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAliqIcmsItem_LostFocus()
    PerdeFoco txtAliqIcmsItem
    TotalAeqBmultC txtVlICMSItem, txtBcIcmsItem, txtAliqIcmsItem
    
    
End Sub


Private Sub txtAliqIcmsStItem_GotFocus()
    GanhaFoco txtAliqIcmsStItem
End Sub

Private Sub txtAliqIpiItem_GotFocus()
    GanhaFoco txtAliqIpiItem
End Sub

Private Sub txtAliqIpiItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAliqPisItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAliqPisItem_LostFocus()
    PerdeFoco txtAliqPisItem
End Sub


Private Sub txtAliqPisItem_GotFocus()
    GanhaFoco txtAliqPisItem
End Sub



Private Sub txtAliqIpiItem_LostFocus()
    PerdeFoco txtAliqIpiItem
End Sub



Private Sub txtAliqIcmsStItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAliqIcmsStItem_LostFocus()
    PerdeFoco txtAliqIcmsStItem
End Sub




Private Sub txtBcIcmsItem_GotFocus()
    GanhaFoco txtBcIcmsItem
End Sub

Private Sub txtBcIcmsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBcIcmsItem_LostFocus()
    PerdeFoco txtBcIcmsItem
End Sub



Private Sub txtBcIcmsStItem_GotFocus()
    GanhaFoco txtBcIcmsStItem
End Sub

Private Sub txtBcIcmsStItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBcIcmsStItem_LostFocus()
    PerdeFoco txtBcIcmsItem
End Sub


Private Sub txtBcIcmsStNot_GotFocus()
    GanhaFoco txtBcIcmsStNot
End Sub

Private Sub txtBcIcmsStNot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBcIcmsStNot_LostFocus()
    PerdeFoco txtBcIcmsStNot
End Sub

Private Sub txtBcIpiItem_GotFocus()
    GanhaFoco txtBcIpiItem
End Sub

Private Sub txtBcIpiItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBcIpiItem_LostFocus()
    PerdeFoco txtBcIpiItem
End Sub

Private Sub txtBcPisItem_GotFocus()
    GanhaFoco txtBcPisItem
End Sub

Private Sub txtBcPisItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBcPisItem_LostFocus()
    PerdeFoco txtBcPisItem
End Sub

Private Sub txtBcCofinsItem_GotFocus()
    GanhaFoco txtBCCofinsItem
End Sub

Private Sub txtBCCofinsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtBcCofinsItem_LostFocus()
    PerdeFoco txtBCCofinsItem
End Sub

Private Sub txtbxUnit_GotFocus()
    GanhaFoco txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxUnit_LostFocus()
    PerdeFoco txtbxUnit
    txtbxTotal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtDescontoItem.Text))
End Sub

Private Sub txtCean_GotFocus()
    GanhaFoco txtCean
End Sub

Private Sub txtCean_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCean_LostFocus()
    AjustaTexto txtCean
End Sub

Private Sub txtCustoUnitario_GotFocus()
    P_SelecTexto txtCustoUnitario
End Sub

Private Sub txtCustoUnitario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtCustoUnitario_LostFocus()
    P_ValorLostFocus txtCustoUnitario
End Sub

Private Sub txtDescontoItem_GotFocus()
    GanhaFoco txtDescontoItem
End Sub

Private Sub txtDescontoItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDescontoItem_LostFocus()
    PerdeFoco txtDescontoItem
    txtbxTotal.Text = FormatVal((CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text)) - CCur(txtDescontoItem.Text))
End Sub


Private Sub txtML_GotFocus()
    P_SelecTexto txtML
End Sub

Private Sub txtML_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtML_LostFocus()
    P_ValorLostFocus txtML
    calcCusto
End Sub

Private Sub txtOutDesp_GotFocus()
    P_SelecTexto txtOutDesp
End Sub

Private Sub txtOutDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxTotal.SetFocus
    End If
End Sub

Private Sub txtOutDesp_LostFocus()
    P_ValorLostFocus txtOutDesp
End Sub

Private Sub txtPreco_GotFocus()
    P_SelecTexto txtPreco
End Sub

Private Sub txtPreco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPreco_LostFocus
        cmdInserir.SetFocus
    End If
    
End Sub

Private Sub txtPreco_LostFocus()
    P_ValorLostFocus txtPreco
    CalculaPreco
End Sub

Private Sub txtQtdConvert_GotFocus()
    GanhaFoco txtQtdConvert
End Sub
Sub calcCusto()
    Dim cCusto As Currency
    Dim qtd As Currency
    Dim indice As Single
    Dim cGastos As Currency
    qtd = CCur(txtQtdConvert.Text)
    On Error Resume Next
    
    indice = CCur(txtbxTotal.Text) / CCur(txtbxValorProd.Text)
    
    lblGastos.caption = (CCur(txtbxFrete.Text) * indice) + _
        (CCur(txtbxDepesas_Not.Text) * indice)
    
    
    cCusto = (CCur(txtbxTotal.Text) / qtd) + (CCur(txtVlIpiItem.Text) / qtd) + _
    (CCur(txtVlIcmsStItem.Text) / qtd) + CCur(lblGastos.caption)
        
    
    lblCusto.caption = FormatVal(cCusto * qtd, "###,##0.0000")
    txtCustoUnitario.Text = FormatVal(cCusto, "###,##0.0000")
    
        
    CalculaPreco
        
    
    
    
End Sub
Private Sub txtQtdConvert_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQtdConvert_LostFocus()
    PerdeFoco txtQtdConvert
    P_Unit
    P_CalcItem
End Sub

Private Sub txtQtdSaida_GotFocus()
    GanhaFoco txtQtdSaida
End Sub

Private Sub txtQtdSaida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQtdSaida_LostFocus()
    PerdeFoco txtQtdSaida
End Sub


Private Sub txtTx_GotFocus()
    P_SelecTexto txtTx
End Sub

Private Sub txtTx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtTx_LostFocus()
    P_ValorLostFocus txtTx
End Sub

Private Sub txtUndSaida_GotFocus()
    GanhaFoco txtUndSaida
End Sub

Private Sub txtUndSaida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtUndSaida_LostFocus()
    AjustaTexto txtUndSaida
End Sub

Private Sub txtUnidEntr_GotFocus()
    GanhaFoco txtUnidEntr
End Sub

Private Sub txtUnidEntr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtUnidEntr_LostFocus()
    AjustaTexto txtUnidEntr
End Sub

Private Sub txtUnitConv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

'
Private Sub txtVlIcmsItem_GotFocus()
    GanhaFoco txtVlICMSItem
End Sub

Private Sub txtVlIcmsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVlIcmsItem_LostFocus()
    PerdeFoco txtVlICMSItem
    If CCur(txtBcIcmsItem.Text) = 0 Then
        
    End If
    
    
    
    TotalCeqBdivA txtBcIcmsItem, txtVlICMSItem, txtAliqIcmsItem
End Sub






Private Sub txtChaveNfe_GotFocus()
    GanhaFoco txtChaveNfe
End Sub

Private Sub txtChaveNfe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtChaveNfe_LostFocus()
    AjustaTexto txtChaveNfe, 0
    txtChaveNfe.Text = F_SoNumero(txtChaveNfe.Text)
End Sub

Private Sub txtPisNot_GotFocus()
    GanhaFoco txtPisNot
End Sub

Private Sub txtPisNot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtPisNot_LostFocus()
    PerdeFoco txtPisNot
End Sub

Private Sub txtCofinsNot_GotFocus()
    GanhaFoco txtCofinsNot
End Sub

Private Sub txtCofinsNot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCofinsNot_LostFocus()
    PerdeFoco txtCofinsNot
End Sub


'
Private Sub txtPisStNot_GotFocus()
    GanhaFoco txtPisStNot
End Sub

Private Sub txtPisStNot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtPisStNot_LostFocus()
    PerdeFoco txtPisStNot
End Sub
'
Private Sub txtCofinsStNot_GotFocus()
    GanhaFoco txtCofinsSTNot
End Sub

Private Sub txtCofinsStNot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCofinsStNot_LostFocus()
    PerdeFoco txtCofinsSTNot
End Sub
'fim
Private Sub txtSeguro_GotFocus()
    GanhaFoco txtSeguro
End Sub

Private Sub txtSeguro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtSeguro_LostFocus()
    PerdeFoco txtBcIcmsStNot
End Sub

'
Private Sub txtVlIcmsStNot_GotFocus()
    GanhaFoco txtVlIcmsStNot
End Sub

Private Sub txtVlIcmsStNot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVlIcmsStNot_LostFocus()
    PerdeFoco txtVlIcmsStNot
End Sub
'
Private Sub txtVlIcmsStItem_GotFocus()
    GanhaFoco txtVlIcmsStItem
End Sub

Private Sub txtVlIcmsStItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVlIcmsStItem_LostFocus()
    PerdeFoco txtVlICMSItem
    On Error GoTo trata_erro
    
    txtAliqIcmsStItem.Text = FormatVal((CCur(txtVlIcmsStItem.Text) + CCur(txtVlICMSItem.Text)) / CCur(txtBcIcmsStItem.Text) * 100)
    Exit Sub
trata_erro:
    
End Sub


Private Sub txtVlIpiItem_GotFocus()
    GanhaFoco txtVlIpiItem
End Sub

Private Sub txtVlIpiItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVlIpiItem_LostFocus()
    PerdeFoco txtVlIpiItem
End Sub


Private Sub txtVlPisItem_GotFocus()
    GanhaFoco txtVlPisItem
End Sub

Private Sub txtVlPisItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVlPisItem_LostFocus()
    PerdeFoco txtVlPisItem
End Sub

Private Sub txtVlCofinsItem_GotFocus()
    GanhaFoco txtVlCofinsItem
End Sub

Private Sub txtVlCofinsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtVlCofins_LostFocus()
    PerdeFoco txtVlCofinsItem
End Sub

'fim
Private Sub txtbxBC_Not_GotFocus()
    GanhaFoco txtbxBC_Not
End Sub

Private Sub txtbxBC_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBC_Not_LostFocus()
    PerdeFoco txtbxBC_Not
    If CCur(txtbxBC_Not.Text) > 0 Then
'        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.Text) / CCur(txtbxBC_Not.Text), 2) * 100) & " %"
    End If
End Sub


Private Sub txtbxBCS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub txtbxCodBarra_GotFocus()
    GanhaFoco txtbxCodBarra
End Sub
Private Sub txtbxCodBarra_LostFocus()
    AjustaTexto txtbxCodBarra
    
    
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnAddProd_Click
    End If
End Sub

Private Sub txtbxDepesas_Not_GotFocus()
    GanhaFoco txtbxDepesas_Not
End Sub

Private Sub txtbxDepesas_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDepesas_Not_LostFocus()
    PerdeFoco txtbxDepesas_Not
End Sub

Private Sub txtbxDesconto_GotFocus()
    GanhaFoco txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    PerdeFoco txtbxDesconto
End Sub

Private Sub txtbxEmissao_Not_GotFocus()
    GanhaFoco txtbxEmissao_Not
End Sub

Private Sub txtbxEmissao_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEmissao_Not_LostFocus()
    P_AjustaData txtbxEmissao_Not
End Sub




Private Sub txtbxEntrada_Not_GotFocus()
    GanhaFoco txtbxEntrada_Not
End Sub

Private Sub txtbxEntrada_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntrada_Not_LostFocus()
    On Error GoTo trata_erro
    P_AjustaData txtbxEntrada_Not
    If txtbxEntrada_Not.Text = "" Then Exit Sub
    If CDate(txtbxEntrada_Not.Text) < CDate(txtbxEmissao_Not.Text) Then
        MsgBox "A data de Entrada não pode ser anterior a data de Emissão", vbExclamation
        txtbxEntrada_Not.SetFocus
    End If
trata_erro:
    
End Sub



Private Sub txtbxFrete_GotFocus()
    GanhaFoco txtbxFrete
End Sub

Private Sub txtbxFrete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFrete_LostFocus()
    PerdeFoco txtbxFrete
    
End Sub

Private Sub txtbxIcms_LostFocus()
    'PerdeFoco txtbxIcms
    txtbxICMSSaida.Text = CCur(F_IcmsProd(txtbxCodBarra.tag, 3)) / 100 'percentual
    P_CalcItem
End Sub

Private Sub txtbxICMS_Not_GotFocus()
    GanhaFoco txtbxICMS_Not
End Sub

Private Sub txtbxICMS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxICMS_Not_LostFocus()
    PerdeFoco txtbxICMS_Not
    If CCur(txtbxBC_Not.Text) > 0 Then
    '        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.Text) / CCur(txtbxBC_Not.Text) * 100, 2)) & " %"
    End If
End Sub

Sub P_CalcItem()
    Dim strSQL As String
    Dim cCusto As Currency
    Dim dInd As Double
    Dim cSuj As Currency
    Dim cLucro As Currency
   
   
calcCusto

    Exit Sub
trata_erro:
End Sub
Sub P_calcLucroLiq()
    Dim cLucro As Currency
    Dim cCusto As Currency
    Dim cIcms As Currency
    If CCur(lblSug.caption) > 0 Then
    On Error GoTo trata_erro
        'cLucro = CCur(txtbxLucro.Text) * CCur(txtbxPrecoFuturo.Text) / CCur(txtbxSujerido.Text)
        cLucro = (CCur(txtbxPrecoFuturo.Text) / CCur(txtbxNovoCusto.Text)) - 1
        cLucro = cLucro * 100
    End If
    txtbxLucroLiq.Text = cLucro
    cLucro = cLucro - (cLucro * CCur(txtTx.Text) / 100)
    cLucro = cLucro - CCur(txtTx.Text)
    txtbxLucroFinal.Text = cLucro
    
    If F_Qtd > 0 Then
        If CCur(txtbxICMSSaida.Text) = 0 Then
            cIcms = (CCur(txtbxICMSSaida.Text) * CCur(txtbxPrecoFuturo.Text) / 100) - (CCur(txtbxValorICMS.Text) / F_Qtd)
        Else
            cIcms = (CCur(txtbxICMSSaida.Text) * CCur(txtbxPrecoFuturo.Text) / 100) - (CCur(txtbxValorICMS.Text) / F_Qtd)
        End If
    End If
    'If cIcms > 0 Then
        cCusto = CCur(txtbxNovoCusto.Text) + cIcms
    'End If
    txtbxCustoSaida.Text = cCusto
    
    txtbxCM.Text = FormatVal((cCusto + CCur(txtbxCusto.Text)) / 2)
    
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
    
End Sub
Sub P_Unit()
    If CCur(txtQtdConvert.Text) <= 0 Then
        txtQtdConvert.Text = txtbxQtd.Text
    End If
    If CCur(txtbxQtd.Text) > 0 Then
        txtbxUnit.Text = FormatVal((CCur(txtbxTotal.Text) + CCur(txtDescontoItem.Text)))
        txtUnitConv.Text = FormatVal((CCur(txtbxTotal.Text) + CCur(txtDescontoItem.Text)))
        
        
        
    End If
    
    
    
    
End Sub

Function F_Restante() As Currency
    F_Restante = F_ValorProdutos - CCur(lblValoLan.caption)

End Function
Function F_RestIpi() As Currency
    F_RestIpi = CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption)
End Function
Private Sub txtbxIpi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub txtbxIPI_Not_GotFocus()
    GanhaFoco txtbxIPI_Not
End Sub

Private Sub txtbxIPI_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxIPI_Not_LostFocus()
    PerdeFoco txtbxIPI_Not
End Sub


Private Sub txtForn_GotFocus()
    GanhaFoco txtForn
End Sub

Private Sub txtForn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxNum_Not_GotFocus()
    If txtbxNum_Not.Text = "" And operacao = "E" Then
        txtbxNum_Not.Text = MaxKeySQL("tab_cmp", "num_not", "cod_for='" & txtForn.tag & "'", "")
        
    End If
    GanhaFoco txtbxNum_Not
    
End Sub

Private Sub txtbxNum_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxNum_Not_LostFocus()
     txtbxNum_Not.Text = Format(F_SoNumero(txtbxNum_Not.Text), "0#######")
End Sub



Private Sub txtbxPrecoFuturo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub


Function F_Qtd() As Currency
'    If cmbxEmb.Text <> "" Then
'        F_Qtd = CCur(cmbxEmb.Text) * CCur(txtbxQtd.Text)
    
        
'    End If
End Function
Private Sub txtbxQtd_GotFocus()
        GanhaFoco txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    PerdeFoco txtbxQtd, "###,##0.000"
    If txtbxIPI_Not.Text <> "0,00" Then
        'txtbxIpi.Enabled = True
        'lblIPI.Enabled = True
    Else
        'txtbxIpi.Enabled = False
        'lblIPI.Enabled = False
    End If
    
    P_Unit
    P_CalcItem
End Sub

Private Sub txtbxSerie_Not_GotFocus()
    GanhaFoco txtbxSerie_Not
End Sub

Private Sub txtbxSerie_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxSerie_Not_LostFocus()
    AjustaTexto txtbxSerie_Not
End Sub

Sub total_item()
    
    
End Sub
Private Sub txtbxTotal_GotFocus()
    GanhaFoco txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdInserir.SetFocus
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    On Error Resume Next
    PerdeFoco txtbxTotal
    P_Unit
    calcCusto
End Sub

Private Sub txtbxValor_Not_GotFocus()
    GanhaFoco txtbxValor_Not
End Sub

Private Sub txtbxValor_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Not_LostFocus()
    PerdeFoco txtbxValor_Not
End Sub

Private Sub txtbxValorProd_GotFocus()
    GanhaFoco txtbxValorProd
End Sub

Private Sub txtbxValorProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorProd_LostFocus()
    PerdeFoco txtbxValorProd
End Sub

Private Sub txtPis_Change()

End Sub

Private Sub txtNCM_GotFocus()
    GanhaFoco txtNCM
End Sub

Private Sub txtNCM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNCM_LostFocus()
    AjustaTexto txtNCM
    If txtNCM.Text <> "" And txtNCM.Text <> txtNCM.tag Then
        alteraNCM txtbxCodBarra.tag, txtNCM.Text
    End If
End Sub


Private Sub txtCST_GotFocus()
    GanhaFoco txtCst
End Sub

Private Sub txtCST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCST_LostFocus()
    AjustaTexto txtCst
End Sub
Private Sub txtCFOganhafoco()
    GanhaFoco txtCfop
End Sub
Sub buscaCFOP()
    Dim c As Padrao1
    c = Sped_Util.Sped_Lista_CFOP(txtCfop.Text)
    If c.codigo <> "" Then
        txtCfop.tag = c.codigo
        txtCfop.Text = c.codigo
        txtCfop.BackColor = vbGreen
    Else
        txtCfop.BackColor = vbRed
    End If
    txtUnidEntr.SetFocus
    
End Sub
Private Sub txtCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCFOP
    End If
End Sub

Private Sub txtCFOP_LostFocus()
    AjustaTexto txtCfop
End Sub

