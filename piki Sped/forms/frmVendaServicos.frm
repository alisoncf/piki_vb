VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVendaServicos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "-Venda de Produtos:"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16980
   Icon            =   "frmVendaServicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   16980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcServicos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000080FF&
      Height          =   3255
      Left            =   5520
      ScaleHeight     =   3225
      ScaleWidth      =   11385
      TabIndex        =   134
      Top             =   600
      Width           =   11415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   155
         Top             =   1560
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2566
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   11025
         TabIndex        =   135
         Top             =   240
         Width           =   11055
         Begin VB.ComboBox cmbxStatusServico 
            Height          =   315
            ItemData        =   "frmVendaServicos.frx":2CFA
            Left            =   2760
            List            =   "frmVendaServicos.frx":2D07
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   152
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "excluir"
            Height          =   285
            Left            =   10200
            TabIndex        =   143
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Inserir"
            Height          =   285
            Left            =   9000
            TabIndex        =   142
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
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
            Left            =   7800
            TabIndex        =   141
            Text            =   "0,00"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6960
            TabIndex        =   140
            Text            =   "0,00"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            TabIndex        =   139
            Text            =   "0,0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5280
            TabIndex        =   138
            Text            =   "0,00"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   4680
            TabIndex        =   137
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   136
            Top             =   240
            Width           =   4455
         End
         Begin VB.CommandButton Command3 
            Height          =   285
            Left            =   2280
            Picture         =   "frmVendaServicos.frx":2D39
            Style           =   1  'Graphical
            TabIndex        =   154
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label45 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2760
            TabIndex        =   156
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label37 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   153
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "Promoção!!!"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   150
            Top             =   1200
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc.:(R$)"
            Height          =   195
            Left            =   6960
            TabIndex        =   149
            Top             =   0
            Width           =   765
         End
         Begin VB.Label Label42 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor Final:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7800
            TabIndex        =   148
            Top             =   0
            Width           =   780
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc.:(%)"
            Height          =   195
            Left            =   6120
            TabIndex        =   147
            Top             =   0
            Width           =   675
         End
         Begin VB.Label Label40 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Unitário:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   146
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label39 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Serviço"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   145
            Top             =   0
            Width           =   540
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4680
            TabIndex        =   144
            Top             =   0
            Width           =   300
         End
      End
      Begin VB.Label lblCapServ 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serviços"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   151
         Tag             =   "06"
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5520
      ScaleHeight     =   4425
      ScaleWidth      =   11385
      TabIndex        =   109
      Top             =   3960
      Width           =   11415
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1065
         ScaleWidth      =   11145
         TabIndex        =   110
         Top             =   120
         Width           =   11175
         Begin VB.TextBox txtbxProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   4455
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   4680
            TabIndex        =   118
            Text            =   "1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxUnit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            TabIndex        =   117
            Text            =   "0,00"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtbxDescPercent 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6480
            TabIndex        =   116
            Text            =   "0,0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtbxDescReais 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   7320
            TabIndex        =   115
            Text            =   "0,00"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtbxFinal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
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
            Left            =   8160
            TabIndex        =   114
            Text            =   "0,00"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtbxPQ_pro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   9360
            TabIndex        =   113
            Text            =   "0,00"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmndbtnInserir 
            Caption         =   "Inserir"
            Height          =   285
            Left            =   8160
            TabIndex        =   112
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "excluir"
            Height          =   285
            Left            =   9360
            TabIndex        =   111
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblEmpresa 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   195
            Left            =   3960
            TabIndex        =   132
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4680
            TabIndex        =   131
            Top             =   0
            Width           =   300
         End
         Begin VB.Label lblUltItem 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   130
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Unitário:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5640
            TabIndex        =   129
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc.:(%)"
            Height          =   195
            Left            =   6480
            TabIndex        =   128
            Top             =   0
            Width           =   675
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor Final:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8160
            TabIndex        =   127
            Top             =   0
            Width           =   780
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc.:(R$)"
            Height          =   195
            Left            =   7320
            TabIndex        =   126
            Top             =   0
            Width           =   765
         End
         Begin VB.Label lblQueima 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "Promoção!!!"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   125
            Top             =   1200
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.Label lblPq_pro 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prom."
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   9480
            TabIndex        =   124
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblEstoque 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "0 em estoque"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   123
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblReservado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Caption         =   "0 em reserva"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3480
            TabIndex        =   122
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblDisponivel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            Caption         =   "0 disponível"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5640
            TabIndex        =   121
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblRefPro 
            Alignment       =   2  'Center
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   120
            Top             =   600
            Width           =   1065
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1815
         Left            =   120
         TabIndex        =   158
         Top             =   1440
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   3201
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Veículo"
         Height          =   195
         Left            =   120
         TabIndex        =   159
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label lblVeiculo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         TabIndex        =   133
         Top             =   1200
         Width           =   195
      End
   End
   Begin VB.PictureBox pcDesconto 
      Appearance      =   0  'Flat
      BackColor       =   &H00B0EAA4&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5520
      ScaleHeight     =   1065
      ScaleWidth      =   5025
      TabIndex        =   89
      Top             =   8760
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmndbtnDesconto 
         Caption         =   "Confirmar Gerencial"
         Height          =   525
         Left            =   2160
         Picture         =   "frmVendaServicos.frx":3089
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxDescPerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   91
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   90
         Text            =   "0,00"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblFinalPrev 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   4200
         TabIndex        =   101
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final Previsto:"
         Height          =   195
         Left            =   4080
         TabIndex        =   100
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   1200
         TabIndex        =   99
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto Autorizado Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblAutorizaDesconto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   2160
         TabIndex        =   96
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto Gerencial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   95
         Top             =   0
         Width           =   1860
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Já Utilizado:"
         Height          =   195
         Left            =   2160
         TabIndex        =   94
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblDescontoGerencialUsado 
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
         Left            =   3120
         TabIndex        =   93
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   5385
      TabIndex        =   53
      Top             =   1320
      Width           =   5415
      Begin VB.TextBox txtComprador 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AC67&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   6
         Tag             =   "00000"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDescontoCli 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   4920
         TabIndex        =   8
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Tag             =   "00000"
         Top             =   840
         Width           =   1695
      End
      Begin VB.PictureBox pcLimite 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   5175
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
         Begin VB.Label lblLimiteRest 
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
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3120
            TabIndex        =   71
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   3120
            TabIndex        =   70
            Top             =   0
            Width           =   690
         End
         Begin VB.Label lblEmUso 
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
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2160
            TabIndex        =   69
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   2160
            TabIndex        =   68
            Top             =   0
            Width           =   600
         End
         Begin VB.Label lblLimite 
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
            Left            =   1200
            TabIndex        =   67
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1200
            TabIndex        =   66
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblStatus_CLi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liberado"
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
            Left            =   0
            TabIndex        =   65
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Tag             =   "00000"
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   330
         Left            =   5040
         Picture         =   "frmVendaServicos.frx":37F3
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprador/Contato"
         Height          =   195
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   5040
         TabIndex        =   7
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   675
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   0
      ScaleHeight     =   9885
      ScaleWidth      =   5385
      TabIndex        =   45
      Top             =   540
      Width           =   5415
      Begin VB.PictureBox pVeiculo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCB683&
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2505
         ScaleWidth      =   5385
         TabIndex        =   88
         Top             =   2640
         Width           =   5415
         Begin VB.TextBox txtObsServ 
            Appearance      =   0  'Flat
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   161
            Top             =   840
            Width           =   5175
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   1320
            TabIndex        =   105
            Top             =   240
            Width           =   3615
         End
         Begin VB.TextBox txtPlaca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   285
            Left            =   4920
            Picture         =   "frmVendaServicos.frx":38EC
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações:"
            Height          =   195
            Left            =   120
            TabIndex        =   160
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblFabricante 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4440
            TabIndex        =   107
            Top             =   0
            Width           =   750
         End
         Begin VB.Label Label36 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Veículo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1320
            TabIndex        =   106
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Placa:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   104
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         ScaleHeight     =   705
         ScaleWidth      =   5385
         TabIndex        =   80
         Top             =   0
         Width           =   5415
         Begin VB.CommandButton cmnbtnRecalc 
            Caption         =   "Recalcula Desconto"
            Height          =   435
            Left            =   3720
            Picture         =   "frmVendaServicos.frx":3C3C
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtbxFormaPag 
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
            TabIndex        =   82
            Tag             =   "01"
            Text            =   "Dinheiro à Vista"
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtbxDescForma 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   3120
            TabIndex        =   81
            Text            =   "0"
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmndbtnFormasPag 
            Height          =   315
            Left            =   2760
            Picture         =   "frmVendaServicos.frx":3FC6
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Forma de Pagamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   1560
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   3120
            TabIndex        =   86
            Top             =   120
            Width           =   120
         End
         Begin VB.Label lblTipoPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carteira"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1800
            TabIndex        =   85
            Tag             =   "06"
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.PictureBox pcMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   5175
         TabIndex        =   73
         Top             =   9480
         Width           =   5175
         Begin VB.Label lblMensagem 
            AutoSize        =   -1  'True
            Caption         =   "..."
            Height          =   195
            Left            =   2280
            TabIndex        =   74
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2385
         ScaleWidth      =   5385
         TabIndex        =   46
         Top             =   6840
         Width           =   5415
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   75
            Top             =   0
            Width           =   5415
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   3360
               TabIndex        =   79
               Top             =   0
               Width           =   930
            End
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   4440
               TabIndex        =   78
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   1320
               TabIndex        =   77
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Serviços:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   360
               TabIndex        =   76
               Top             =   0
               Width           =   870
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8AC67&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   54
            Top             =   1800
            Width           =   5415
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Produtos:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   360
               TabIndex        =   58
               Top             =   0
               Width           =   900
            End
            Begin VB.Label lblTotal 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   1320
               TabIndex        =   57
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label lblDesconto 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   4440
               TabIndex        =   56
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Desconto:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   3360
               TabIndex        =   55
               Top             =   0
               Width           =   930
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8AC67&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   48
            Top             =   2040
            Width           =   5415
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Itens:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   3840
               TabIndex        =   52
               Top             =   0
               Width           =   510
            End
            Begin VB.Label lblItens 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   4440
               TabIndex        =   51
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label lblProd 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   1320
               TabIndex        =   50
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Produtos:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   360
               TabIndex        =   49
               Top             =   0
               Width           =   900
            End
         End
         Begin VB.Label lblValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   63.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1440
            Left            =   1560
            TabIndex        =   47
            Tag             =   "!"
            Top             =   240
            Width           =   2250
         End
      End
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00F8AC67&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   16950
      TabIndex        =   35
      Top             =   10455
      Width           =   16980
      Begin VB.CommandButton cmndbtnBuscaFun 
         Height          =   315
         Left            =   3840
         Picture         =   "frmVendaServicos.frx":40BF
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtbxVendedor 
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
         Left            =   840
         TabIndex        =   59
         Text            =   "Nome do Funcionario"
         Top             =   15
         Width           =   3255
      End
      Begin VB.Label lblGerencial 
         AutoSize        =   -1  'True
         Caption         =   "Pressione F8 para gerencial"
         Height          =   195
         Left            =   14160
         TabIndex        =   102
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label lblDescontoMax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5160
         TabIndex        =   72
         Top             =   52
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4680
         TabIndex        =   62
         Top             =   0
         Width           =   165
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   61
         Top             =   60
         Width           =   420
      End
      Begin VB.Label lblTur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9960
         TabIndex        =   44
         Top             =   0
         Width           =   330
      End
      Begin VB.Label lblBloqueado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   10560
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   42
         Top             =   60
         Width           =   465
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8760
         TabIndex        =   41
         Top             =   0
         Width           =   330
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8160
         TabIndex        =   40
         Top             =   60
         Width           =   435
      End
      Begin VB.Label lblNomeEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   12180
         TabIndex        =   39
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5760
         TabIndex        =   38
         Top             =   60
         Width           =   630
      End
      Begin VB.Label lblnum_ven 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6600
         TabIndex        =   37
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atendente:"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   36
         Top             =   15
         Width           =   840
      End
   End
   Begin VB.PictureBox pcAtalhos 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   16950
      TabIndex        =   34
      Top             =   0
      Width           =   16980
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   12
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   1
            Left            =   0
            Picture         =   "frmVendaServicos.frx":41B8
            ScaleHeight     =   405
            ScaleWidth      =   405
            TabIndex        =   13
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F2 - Imprimir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   480
            TabIndex        =   14
            Top             =   120
            Width           =   1050
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   9
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "frmVendaServicos.frx":46A6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   10
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F1 - Forma Pag"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   480
            TabIndex        =   11
            Top             =   120
            Width           =   1200
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   3600
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   15
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   120
            Picture         =   "frmVendaServicos.frx":4A30
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   16
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F3 - Entrega"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   2
            Left            =   480
            TabIndex        =   17
            Top             =   120
            Width           =   960
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   5400
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   18
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   120
            Picture         =   "frmVendaServicos.frx":4CA4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   19
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F4 - Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   3
            Left            =   480
            TabIndex        =   20
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   7200
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   21
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   120
            Picture         =   "frmVendaServicos.frx":5846
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   22
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F5 - Navegar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   4
            Left            =   480
            TabIndex        =   23
            Top             =   120
            Width           =   990
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   9000
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   24
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   120
            Picture         =   "frmVendaServicos.frx":6248
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   25
            Top             =   105
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F6 - Cancelar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   5
            Left            =   480
            TabIndex        =   26
            Top             =   120
            Width           =   1050
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   10800
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   27
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmVendaServicos.frx":860F
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   28
            Top             =   105
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F7 - Concluir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   6
            Left            =   480
            TabIndex        =   29
            Top             =   120
            Width           =   1020
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   12600
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   30
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   120
            Picture         =   "frmVendaServicos.frx":8999
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   31
            Top             =   120
            Width           =   195
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F12 - Sair"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   7
            Left            =   480
            TabIndex        =   32
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmndbtnLocaliza 
      Caption         =   "..."
      Height          =   255
      Left            =   14760
      TabIndex        =   33
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmVendaServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemEstoque As Boolean
Dim WithEvents AdoVendas As ADODB.Recordset
Attribute AdoVendas.VB_VarHelpID = -1
Dim p As Produtos.Produto
Dim bBlock As Boolean
Dim bDirecao As Integer
Dim bActive As Boolean
Dim preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer
Dim FormaPagamento As FormasPag.FormasPag
Public Atacado As Boolean

Private Type prodPromo
    nome As String
    preco As String
End Type

Dim pPromo() As prodPromo
Dim QtdPromo As Integer
Dim iIndPromo As Integer
Sub AtualizaVendaEndereco()
    P_AtualizaVendaEndereco Me.tag, "", "", "", "", ", txtComprador.Text"
End Sub
Sub Carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    
    If Me.tag = "" Then Exit Sub
    If bActive = False Then
        strSQL = F_VerificaVendaAberta(Me.tag)
        If strSQL <> "" Then
            MsgBox "A venda se encontra aberta no computador " & strSQL, vbExclamation
            Unload Me
            Exit Sub
        End If
    End If
    strSQL = "select tv.num_ven,valor_ven,comprador_ven,telefonecli_ven,endereco_ven,bairro_ven,fone_ven,desconto_ven,nomecli_ven,entrega_ven,"
    strSQL = strSQL & " cod_cli,valor_ven,"
    strSQL = strSQL & "tf.*,obs_ven,tpp.nome_tip from " & TAB_VENDAS & " as tv, tab_formaspag as tf, tab_tipopag as tpp "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for "
    strSQL = strSQL & " and num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tpp.cod_tip=tf.cod_tip "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!nomecli_ven) Then
            txtbxNomeCli.Text = !nomecli_ven
        End If
        If Not IsNull(!telefonecli_ven) Then
            txtTel.Text = !telefonecli_ven
        End If
        txtbxNomeCli.tag = !cod_cli
        
        If UCase(txtbxNomeCli.Text) = "VENDA A CONSUMIDOR" Then
            txtbxNomeCli.Text = ""
        End If
        FormaPagamento = FormasPag.F_BuscaFormasPag(!cod_for)
        txtbxFormaPag.Text = !nome_For
        txtbxDescForma.Text = !desconto_for
        txtbxFormaPag.tag = !cod_for
        txtEnd.Text = Format(!endereco_ven)
        txtBairro.Text = Format(!bairro_ven)
        txtFone.Text = Format(!fone_ven)
        lblTipoPag.tag = !cod_tip
        lblTipoPag.caption = !nome_tip
        txtComprador.Text = Format(!comprador_ven)
        lblDescontoMax.caption = !desconto_for + CCur(lblDesc.caption)
        txtObsEntrega.Text = TiraNull(!obs_ven)
        txtDescontoCli.Text = F_BuscaUmCampo("desconto_cli", TAB_CLIENTES, "cod_cli='" & !cod_cli & "'")
        
        If CCur(txtDescontoCli.Text) > "0" Then
            txtbxDescForma.Text = txtDescontoCli.Text
        End If
        
        
        If !entrega_ven = "S" Then
            chkEntrega.Value = 1
            
        End If
        P_GravaVendaAberta Me.tag
        .Close
    End With
    Exit Sub
trata_Erro:
   ' MostraErro
End Sub



Sub p_totalItem()
    txtbxFinal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - (CCur(txtbxDescReais.Text)))
    
End Sub

Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    Dim cDesc As Currency
    On Error GoTo trata_Erro:
    Dim cPq As Currency
    P_Status "Atualizando Valores..."
    

    iPos = AdoVendas.AbsolutePosition
    AdoVendas.Requery
    If AdoVendas.RecordCount > 0 Then
        AdoVendas.MoveFirst
    End If
    With AdoVendas
        
        
        Do While Not .EOF
            cTotal = cTotal + !final
            cPq = cPq + !pq_itd * !qtd_ite
            cItens = cItens + !qtd_ite
            cDesc = cDesc + !Desconto_ite
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cTotal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = AdoVendas.RecordCount
        lblTotalQueima.caption = FormatVal(cPq)
        lblDesconto.caption = FormatVal(cDesc)
        lblTotal.caption = FormatVal(cTotal + cDesc)
    End With
        
    
    
    If AdoVendas.RecordCount > 0 Then
        AdoVendas.MoveFirst
        AdoVendas.Move iPos
    End If
    
    
    lblProd.caption = AdoVendas.RecordCount
    lblItens.caption = cItens

    strSQL = "update " & TAB_VENDAS & " set valor_ven =" & F_GravaDec(cTotal) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(cDesc) & ","
    strSQL = strSQL & "nomecli_ven= '" & txtbxNomeCli.Text & "', "
    strSQL = strSQL & "telefonecli_ven= '" & txtTel.Text & "', "
    strSQL = strSQL & "comprador_ven= '" & txtComprador.Text & "', "
    strSQL = strSQL & " cod_cli= '" & txtbxNomeCli.tag & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "'  "
    'strSQL = strSQL & " abrevenda_usu='" & txtbxVendedor.tag & "' "
    'strSQL = strSQL & " obs_VEN='" & txtbxObs.Text & "' "
    strSQL = strSQL & " where num_Ven='" & Me.tag & "' "
    conn.Execute strSQL
    
    
    P_Status ""
    Exit Sub
trata_Erro:
P_Status ""
'    mostraerro
End Sub
Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = vendas.F_BuscaItens(Me.tag)
    
    
    
    
    
    abreObjetoAdo AdoVendas, strSQL, dtgrd
    
    
    
            
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub cmndbtnCancel_Click()
    Dim strSQL  As String
    Dim i As Integer
    On Error Resume Next
    If AdoVendas.RecordCount > 0 Then
    
    If strExclVenda_usu <> "S" Then
        strSQL = F_senhaReturnUsu
        If F_BuscaUmCampo("exclvenda_usu", "" & TAB_DEFUSU & "", "cod_usu='" & strSQL & "'") <> "S" Then
            MsgBox "Usuário Não Autorizado!", vbCritical
            Exit Sub
        End If
    End If
    End If
    If F_BuscaUmCampo("num_ven", TAB_VENFEC, "num_ven='" & Me.tag & "'") <> "" Then
        MsgBox "Cancelamento impossível", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja cancelar esta venda?", vbYesNo + vbCritical) = vbYes Then
        
        With AdoVendas
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                Excluir False
                .MoveNext
            Loop
        End With
        strSQL = "insert into tab_cancelamentos(num_ven,data_can,valor_ven,tipo_ven,"
        strSQL = strSQL & "cod_cai,cod_tur)values("
        strSQL = strSQL & "'" & Me.tag & "',"
        strSQL = strSQL & F_GravaDataTempo(Now) & ","
        strSQL = strSQL & F_GravaDec(lblValor.caption) & ","
        strSQL = strSQL & "'" & operacao & "',"
        strSQL = strSQL & "'" & strCod_Cai & "',"
        strSQL = strSQL & "'" & strCod_Tur & "') "
        conn.Execute strSQL
        
        
        If F_BuscaUmCampo("num_ven", TAB_VENFEC, "num_ven='" & Me.tag & "'") <> "" Then
            MsgBox "Cancelamento impossível", vbExclamation
            Exit Sub
        End If
        
        strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        strSQL = "delete from " & TAB_ITVENDAS & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
'            RefreshDataPosition frmEdiVendas.adc
        End If
        
        If MsgBox("Deseja Gerar uma Nova Venda?" & Chr(13) & "Clique em não caso queira abandonar o modo de vendas!", vbYesNo + vbQuestion) = vbYes Then
            strSQL = F_GeraNovaVenda("000000")
            iItem = 0
            Me.tag = strSQL
            Form_Activate
        Else
            iItem = 0
            Unload Me
        End If
    End If
End Sub



Private Sub cmndbtnEditProd_Click()
    


trata_Erro:
MsgBox "Err: cmndbtnEditProd_Click() "
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Excluir True
End Sub
Sub Excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    With AdoVendas
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & !desc_ite & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        iPos = .AbsolutePosition
        strSQL = "delete from " & TAB_ITVENDAS & " where seq_ite=" & !seq_ite & " "
        
        strSQL = strSQL & " and num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        P_GravaItemCancelado
        
        
            strSQL = "update " & TAB_PRODUTOS & " set reservado_pro=reservado_pro - " & F_GravaDec(CStr(!qtd_ite))
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            conn.Execute strSQL
            strSQL = "update estoque set reservado=reservado - " & F_GravaDec(CStr(!qtd_ite))
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and cod_emp='" & !cod_emp & "'"
                        
            conn.Execute strSQL
        
        'AdoVendas.
        
        '    AdoVendas.Requery
        
        If bQuestion Then
            P_AtualizaVenda
        End If
        If AdoVendas.RecordCount = 0 Then
            txtbxQtd.Enabled = False
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
    
    
    If bQuestion Then
        P_buscaitens
    End If
    If AdoVendas.RecordCount = 0 Then
        lblDescontoGerencialUsado.caption = "0,00"
    End If
    Exit Sub
trata_Erro:
    MostraErro "Excluir()"
End Sub


Private Sub AdoVendas_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    With AdoVendas
    
    If Not .EOF Then
        buscaReservado
        
        
        
        
        End If
    End With
End Sub
Sub buscaReservado()
'    lblReservado.Visible = True
'    lblDisponivel.Visible = True
'    lblEstoque.Visible = True
'        lblReservado.caption = "Reservado: " & AdoVendas!reservado_pro
'        lblReservado.tag = AdoVendas!reservado_pro
'        lblDisponivel.caption = "Disponível: " & AdoVendas!er_pro - AdoVendas!reservado_pro
'        lblDisponivel.tag = AdoVendas!er_pro - AdoVendas!reservado_pro
'        lblEstoque.caption = "Em estoque: " & AdoVendas!er_pro
'        lblEstoque.tag = AdoVendas!er_pro
'
'        If AdoVendas!er_pro - AdoVendas!reservado_pro > 0 Then
'            lblDisponivel.BackColor = vbGreen
'        Else
'            lblDisponivel.BackColor = vbRed
'        End If
        
End Sub
Sub hab_entrega(op As Boolean)
    lblFone.Enabled = op
    lblEnd.Enabled = op
    lblFone.Enabled = op
    lblBairro.Enabled = op
   ' lblObsEntrega.Enabled = op
    txtEnd.Enabled = op
    txtFone.Enabled = op
    txtBairro.Enabled = op
    txtFone.Enabled = op
   ' txtObsEntrega.Enabled = op
    If txtEnd.Enabled = True Then
        txtEnd.SetFocus
    End If
    
    
End Sub

Private Sub chkEntrega_Click()
    hab_entrega chkEntrega.Value = 1
    AtualizaVendaEndereco
End Sub

Private Sub cmnbtnRecalc_Click()
    If MsgBox("Confirma cancelamento do desconto atual?", vbYesNo + vbQuestion) = vbYes Then
        P_CancelaDescontoVenda Me.tag, txtbxDescForma.Text
        P_buscaitens
        P_AtualizaVenda
    End If
End Sub

Private Sub cmndbtnCancelar_Click()

End Sub

Private Sub cmndbtnCliente_Click()
     Dim c As Clientes.cliente
     c = Clientes.F_BuscaCliente(txtbxNomeCli.Text)
     If c.cod_cli <> "" Then
        If c.status_cli = "2" Then
            MsgBox "Este cliente é baixado/inativo. Não é possível gerar vendas ou orçamentos.!", vbExclamation
            Exit Sub
        Else
            'Exit Sub
        End If
        
        
        txtbxNomeCli.Text = c.nome_Cli
        txtbxNomeCli.tag = c.cod_cli
        txtBairro.Text = c.bairro_cli
        txtEnd.Text = c.endereco_cli
        txtFone.Text = c.telefone_cli
        txtDescontoCli.Text = c.desconto_cli
        
        txtbxFormaPag.Text = c.formaPag.nome_For
        txtbxFormaPag.tag = c.formaPag.cod_for
        
        If c.status_cli = "3" Then
            
            MsgBox "Este cliente está com crédito bloqueado!", vbExclamation
        End If
        txtDescontoCli.Text = F_BuscaUmCampo("desconto_cli", TAB_CLIENTES, "cod_cli='" & c.cod_cli & "'")
        
        If CCur(txtDescontoCli.Text) >= "0" Then
            txtbxDescForma.Text = txtDescontoCli.Text
        End If
        txtTel.Text = c.telefone_cli
        P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text, txtTel.Text, ""
        P_AtualizaFormaPag Me.tag, txtbxFormaPag.tag
        AtualizaVendaEndereco
        AtualizaDesconto Me.tag, txtbxDescForma.Text
        txtComprador.SetFocus
     End If
    
End Sub

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    Dim sDes As String
    If txtbxDesconto.Text = 0 Then
        Exit Sub
    End If
    sUsu = F_SenhaCodBar
    
    If lblAutorizaDesconto.tag <> "" Then
        MsgBox "Já foi autorizado um desconto para esta venda!", vbExclamation
        Exit Sub
    End If
    
    If sUsu = "" Then
        MsgBox "Não autenticado!", vbExclamation
        Exit Sub
    End If
    
    sDes = F_BuscaUmCampo("desconto_usu", "" & TAB_DEFUSU & "", "cod_usu='" & sUsu & "'")
    If sDes = "" Then
        sDes = "0"
    End If
    If CCur(sDes) < CCur(txtbxDescPerc.Text) + CCur(lblDescontoGerencialUsado.caption) Then
        MsgBox "O Desconto possível para este usuário é de " & sDes & "%", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    
    
    If MsgBox("Confirma desconto de " & txtbxDesconto.Text & "?", vbYesNo + vbCritical) = vbYes Then
        
        vendas.P_RecalculaVendaDesconto lblFinalPrev.caption, txtbxDescPerc.Text, Me.tag, sUsu
        
        P_AtualizaVenda
    End If
    lblDescontoGerencialUsado.caption = CCur(lblDescontoGerencialUsado.caption) + CCur(txtbxDescPerc.Text)
    
End Sub

Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    Dim strSQL As String
    f = F_BuscaFormasPag(txtbxFormaPag.tag, "")
    FormaPagamento = f
    If f.cod_for <> "" Then
        If CCur(txtbxDescForma.Text) > f.desconto_for Then
            If AdoVendas.RecordCount > 0 Then
                strSQL = "Você optou por mudar a forma de pagamento de " & txtbxFormaPag.Text
                strSQL = strSQL & " para " & f.nome_For & Chr(13)
                strSQL = strSQL & "Continuando com esta operação, todo o desconto da venda será cancelado!"
                strSQL = strSQL & "Deseja continuar?"
                If MsgBox(strSQL, vbOKCancel) = vbCancel Then
                    Exit Sub
                Else
                    P_CancelaDescontoVenda Me.tag
                    P_buscaitens
                End If
            End If
        End If
        If f.avista_For = "N" And lblStatus_CLi <> "1" Then
            MsgBox "Este cliente está bloqueado!", vbCritical
            Exit Sub
        End If
        txtbxFormaPag.Text = f.nome_For
        txtbxFormaPag.tag = f.cod_for
        txtbxDescForma.Text = f.desconto_for
        txtbxDescForma.tag = f.avista_For
        lblDescontoMax.caption = f.desconto_for + CCur(lblDesc.caption)
        txtbxDescPercent.Text = txtbxDescForma.Text
        lblTipoPag.caption = f.nome_tip
        lblTipoPag.tag = f.cod_tip
        
        'P_recalculaVenda
        P_AtualizaFormaPag Me.tag, txtbxFormaPag.tag
        P_AtualizaVenda
        
        If Me.ActiveControl.Name = "txtbxNomeCli" Then
            
        Else
            txtbxNomeCli.SetFocus
        End If
        
        
        Exit Sub
    End If
End Sub

Private Sub cmndbtnInserir_Click()
    If TemEstoque Then
        insere
    Else
        MsgBox "Este item possui " & lblEstoque.tag & " em estoque e " & lblReservado.tag & " em vendas em fase de lançamento. ", vbExclamation
        
    End If
End Sub

Private Sub cmndbtnLocaliza_Click()
    
    buscaProd
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    P_AtualizaVenda
    If AdoVendas.RecordCount > 0 Then
        
            If MsgBox("Confirma Saída da Venda?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            If txtbxNomeCli.Text = "" Or txtbxNomeCli.Text = "" Then
                MsgBox "Informe o nome do cliente!", vbExclamation
                txtbxNomeCli.SetFocus
                Exit Sub
            End If
            
            
    End If
    If IsLoad(frmEdiVendas.Name) Then
        'frmEdiVendas.adc.Refresh
    End If
    If AdoVendas.RecordCount = 0 Then
        strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        strSQL = "delete from " & TAB_ITVENDAS & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        
    End If
    P_DeletaVendaAberta Me.tag
    Unload Me
    Exit Sub
trata_Erro:
    'MostraErro "cmndbtnSair"
    Unload Me

End Sub


Sub ConcluirVenda()
    cmndbtnSalvar_Click
End Sub
Private Sub cmndbtnSalvar_Click()
    Dim sTroco As String
    Dim strSQL As String
    Dim sVen As String
    Dim s As String
    Dim c As cliente
    If txtbxFormaPag.tag = "" Then
        MsgBox "Escolha uma forma de pagamento!", vbExclamation
        Exit Sub
    End If
    
    c = Clientes.F_BuscaClienteCodigo(txtbxNomeCli.tag)
    
    If AdoVendas.RecordCount = 0 Then
        MsgBox "Não há itens na sua venda!", vbExclamation
        Exit Sub
    End If
    
    If bUsaImpedirVendaDeProdutoSemEstoque Then
        If VerificaEstoqueProdutoMenorQueZero = True Then
            Exit Sub
        End If
    End If
    If VerificaEmpresas = False Then
        MsgBox "Existem produtos de empresas diferentes. É necessário gerar uma venda para cada empresa!", vbExclamation
        Exit Sub
    End If
    
    If LCase(Left(sDataBase, 7) = "estoque") Then
        If Trim(txtComprador.Text) = "" Then
            'MsgBox "Informe o nome do comprador", vbExclamation
            'txtComprador.SetFocus
            'Exit Sub
        End If
    End If
    
    
    strSQL = F_senhaReturnUsu(False, True)
    If strSQL = "" Then
        MsgBox "não autenticado!", vbExclamation
        Exit Sub
    End If
    If strSQL <> txtbxVendedor.tag Then
        txtbxVendedor.tag = strSQL
        txtbxVendedor.Text = F_BuscaUmCampo("nome_fun", TAB_FUNCIONARIOS, "cod_fun='" & strSQL & "'")
        vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text
        P_AtualizaVenda
    End If
    
    
    
        If txtbxNomeCli.Text = "" Then
            strSQL = InputBox("Informe um nome para o cliente", "Nome do Cliente", "Venda a consumidor")
            If strSQL = "" Then
                MsgBox "Informe um nome para o cliente!", vbExclamation
                txtbxNomeCli.SetFocus
                Exit Sub
            Else
                
                txtbxNomeCli.Text = strSQL
                vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, strSQL
            End If
            
            
            
        End If
        If bUsaCrediario Then
            If lblTipoPag.tag = "02" Then
                If F_GravaCrediarioVen(Me.tag) Then
                    MsgBox "Esta venda venda deve ser liberada pelo setor de Crediário para ser concretizada!" & Chr(13) & "Por favor, encaminhe o cliente ao Crediário.", vbInformation
                End If
            End If
        End If
        
        
        
        If FormaPagamento.avista_For = "N" And c.status_cli <> "1" Then
            
                MsgBox "Essa venda não pode ser concluída para esse cliente nessa forma de pagamento!", vbExclamation
                Exit Sub
            
        End If
        
        
        If vendas.F_AlteraStatusVenda(Me.tag, "2") Then
            
            If Not bUsaCaixaSeparado Then
                s = F_Troco(Me.tag)
                If s = "" Then
                    Exit Sub
                End If
            End If
            
            P_DeletaVendaAberta Me.tag
            Me.tag = ""
            Me.tag = F_GeraNovaVenda
            If Me.tag = "" Then
                MsgBox "Não foi possível gerar uma nova venda!", vbExclamation
                Unload Me
                Exit Sub
            End If
            bActive = False
            Form_Activate
        End If
    
    
    
End Sub


Private Sub dtgrd_DblClick()
    editaItem
End Sub

Private Sub dtgrd_GotFocus()
    
    dtgrd.MarqueeStyle = 3
    
    
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo trata_Erro

    '    On Error GoTo trata_erro

    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        
        
                
                    If dtgrd.row = -1 Then
                        MsgBox "Selecione o item que deseja excluir!", vbCritical
                        dtgrd.SetFocus
                        Exit Sub
                    End If
                    
            Excluir False
        
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxQtd.SetFocus
    End If
    P_AtualizaVenda
    Exit Sub
trata_Erro:
MostraErro

End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub


Private Sub dtgrd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub adc_Reposition()
    
        
End Sub


Sub buscaPromo()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
        
    strSQL = "select nome_pro,pv_pro from " & TAB_PRODUTOS & " as tp,"
    strSQL = strSQL & " tab_prodpromo as tpp where tpp.cod_Pro=tp.cod_pro"
    strSQL = strSQL & " and datafin_pro >=" & F_GravaData(Date)
    strSQL = strSQL & " order by nome_pro"
    
    rPesq.Open strSQL, conn
    
    With rPesq
        QtdPromo = .RecordCount
        If .RecordCount > 0 Then
            ReDim pPromo(.RecordCount) As prodPromo
        End If
        strSQL = ""
        Do While Not .EOF
            pPromo(.AbsolutePosition).nome = !nome_pro
            pPromo(.AbsolutePosition).preco = "R$ " & FormatVal(!pv_pro)
            
            .MoveNext
        Loop
        .Close
        
        
        
    
    End With
    
End Sub
Function VerificaEstoqueProdutoMenorQueZero() As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select tp.cod_pro,tp.codfab_pro, tp.nome_pro,er_pro from " & TAB_PRODUTOS & " as tp inner join " & TAB_ITVENDAS & "  as ti"
    strSQL = strSQL & " on ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " inner join estoque as te on tp.cod_pro=te.cod_pro and ti.cod_emp=te.cod_emp"
    strSQL = strSQL & " where ti.num_ven='" & Me.tag & "' "
    
    
    strSQL = strSQL & " and er_pro<=0"
    
    abreObjetoAdo rPesq, strSQL
    
    strSQL = ""
    With rPesq
        Do While Not .EOF
            strSQL = strSQL & !codfab_pro & " - " & !nome_pro & Chr(13) & Chr(10)
        
            .MoveNext
        Loop
    End With
    
    If strSQL <> "" Then
        VerificaEstoqueProdutoMenorQueZero = True
        MsgBox "Os seguintes produtos não possuem estoque disponível para a venda:" & Chr(13) & strSQL, vbExclamation
        
    End If
    
    
    
    
    
    
    
End Function
Function VerificaEmpresas() As Boolean
    Dim item As Integer
    Dim emp As String
    With AdoVendas
        VerificaEmpresas = True
        If .RecordCount > 0 Then
            item = AdoVendas.AbsolutePosition
            .MoveFirst
            emp = !cod_emp
        End If
        
        Do While Not .EOF
            If emp <> !cod_emp Then
            
                VerificaEmpresas = False
            End If
            emp = !cod_emp
            
            
            .MoveNext
        Loop
    
    
    
    
        If .RecordCount > 0 Then
            .MoveFirst
            .Move item
        End If
        
    
    End With
    
    
End Function



Private Sub Form_Activate()
'    On Error GoTo trata_erro
    
    If Not bUsaQueima Then
        txtbxPQ_pro.Visible = False
        lblPq_pro.Visible = False
        dtgrd.Columns(7).Visible = False
        lblTotQueimaCaption.Visible = False
        lblTotalQueima.Visible = False
    End If
    
    lblVeiculo.caption = UltimoVeiculoProcurado.descricao
    lblDesc.caption = crcDesconto_usu
    lblNomeEmp.caption = strFantasia_Emp
    lblnum_ven.caption = Me.tag
    txtbxVendedor.tag = strCod_Usu
    txtbxVendedor.Text = strNome_Usu
    lblNomeEmp.caption = strFantasia_Emp
    lblCaixa = strCod_Cai
    lblTur.caption = strCod_Tur
    
    
    
    
    
    Carrega
    
    
    
    
    
    P_Status "Buscando Itens..."
On Error Resume Next
    'txtbxProduto.SetFocus
    txtbxFormaPag_KeyDown 9, 0
    P_buscaitens
    P_AtualizaVenda
    P_Status ""
    
    If bActive = True Then
        Exit Sub
    Else
        txtbxProduto.SetFocus
    End If
    'buscaFormaAvista
    bActive = True
    buscaLogo
    
    Exit Sub
trata_Erro:
    MostraErro "Form_Activate"
    
End Sub
Sub playGif()
    'WebBrowser1.Navigate "c:\market\gifs\maebela2.gif"
    'WebBrowser1.Navigate "about:<html><body scroll='no'><img src='c:\maebela.gif'></img></body></html>"
    
End Sub
Sub buscaLogo()
    
End Sub
    


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    bDescancoTela = False
    iDescancoTela = 0
    
    If KeyCode = vbKeyF3 Then
        If chkEntrega.Value = 0 Then
            chkEntrega.Value = 1
        Else
            chkEntrega.Value = 0
        End If
    ElseIf KeyCode = vbKeyF5 Then
        
        frmEdiPro.Show 0
    ElseIf KeyCode = vbKeyF4 Then
         txtbxNomeCli.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF2 Then
        'ImprVenda Me.tag
    ElseIf KeyCode = vbKeyF8 Then
        pcDesconto.Visible = True
    ElseIf KeyCode = vbKeyF9 Then
        
    ElseIf KeyCode = vbKeyF10 Then
        
    ElseIf KeyCode = vbKeyF1 Then
        If Me.ActiveControl.Name = txtbxNomeCli.Name Then
            cmndbtnFormasPag.SetFocus
        End If
        cmndbtnFormasPag_Click
    ElseIf KeyCode = vbKeyF12 Or KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    
    End If
    
    
    If KeyCode = vbKeyP And (Shift = vbAltMask Or Shift = vbShiftMask) Then
        buscaProd
    End If
    
    If KeyCode = vbKeyL And Shift = vbAltMask Then
        limpaVeiculo
    End If
    
End Sub
Sub limpaVeiculo()
    lblVeiculo.tag = ""
    lblVeiculo.caption = ""
    UltimoVeiculoProcurado.id = 0
    UltimoVeiculoProcurado.descricao = ""
    
End Sub
Sub mudamodo()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set modo_cai='" & operacao & "' "
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
    conn.Execute strSQL
End Sub

Private Sub Form_Load()
'    buscaPromo
    bActive = False

    'adc.ConnectionString = Conn.ConnectionString

    lblNomeEmp.caption = strNome_Emp
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", tab_parametros, "")
    txtbxNomeCli.tag = "00000"
    txtbxNomeCli.Text = "VENDA A CONSUMIDOR"
        
        
        
        
    
    Set AdoVendas = New ADODB.Recordset
    
    
    
    
    
'    cmbxBuscar.ListIndex = 0
    'pcECF.BackColor = vbGreen
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
End Sub
Sub buscaFormaAvista()
    Dim f As FormasPag.FormasPag
    Dim s As String
    If strFormaPagamentoPadrao = "" Then
        strFormaPagamentoPadrao = "06"
    End If
    
    
    f = FormasPag.F_BuscaFormasPag("06")
    If f.cod_for <> "" Then
        txtbxDescForma.Text = f.desconto_for
        txtbxDescForma.tag = f.cod_for
        
        On Error Resume Next
        
        txtbxFormaPag.Text = f.nome_For
        txtbxFormaPag.tag = f.cod_for
        txtbxDescForma.Text = f.desconto_for
        txtbxDescForma.tag = f.avista_For
        lblDescontoMax.caption = f.desconto_for + CCur(lblDesc.caption)
        txtbxDescPercent.Text = txtbxDescForma.Text
        lblTipoPag.caption = f.nome_tip
        lblTipoPag.tag = f.cod_tip
        
        'P_recalculaVenda
        P_AtualizaFormaPag Me.tag, txtbxFormaPag.tag
        P_AtualizaVenda
        txtbxProduto.SetFocus
        
    End If
    
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDescancoTela = False
    iDescancoTela = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Resize()
    ajustaForm
    
End Sub
Sub ajustaForm()
        Dim iCount As Integer
    On Error Resume Next
    sstb.width = Me.width - sstb.Left - 200
    
    
    sstb.Height = pcStatusBar.Top - sstb.Top + 300
    lblNomeEmp.Left = Me.width - lblNomeEmp.width - 200
    'pcCli.Width = pcProd.Width
    
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).width = pcAtalhos.width / (pcBotao.Count - 1)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).width + 50
        End If
        
    Next
 
    pcProd.width = dtgrd.width
    'pcCli.Width = pcProd.Width
    
    PosicionaGrade dtgrd, sstb, pcStatusBar
    
    dtgrd.Height = sstb.Height - dtgrd.Top - 100
    
    pcProd.width = dtgrd.width
    
End Sub

Sub editaItem()
    'On Error GoTo trata_Erro
    Exit Sub 'para terminar depois
    With AdoVendas
        cmndbtnInserir.tag = !seq_ite
        txtbxProduto.Text = !desc_ite
        txtbxProduto.tag = !cod_pro
        txtbxUnit.Text = !unit_ite
        txtbxQtd.Text = !qtd_ite
        txtbxPQ_pro.Text = !pq_itd
        txtbxDescReais.Text = !Desconto_ite
        txtbxDescReais_LostFocus
        txtbxQtd.SetFocus
        
    End With
    
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub

Private Sub optnCod_Click()
    txtbxProduto.SetFocus
End Sub

Private Sub optnDesc_Click()
    txtbxProduto.SetFocus
End Sub

Private Sub optnEnd_Click()
    
End Sub

Private Sub optnSeq_Click()
    P_buscaitens
End Sub









Sub P_foco(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BorderStyle = 0
        
    Next
    pcBotao(Indice).BorderStyle = 1
'    pcbotao(Indice).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    AdoVendas.Close
End Sub

Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub lblReservado_Click()
    On Error Resume Next
    If txtbxProduto.tag <> "" Then
        P_BuscaEstoqueReservado (txtbxProduto.tag)
    Else
        P_BuscaEstoqueReservado (AdoVendas!cod_pro)
    End If
    
End Sub

Private Sub pcIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_GotFocus(Index As Integer)
    Sendkeys Chr(9)
End Sub



Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If Index = 0 Then
            pcBotao(pcBotao.Count - 1).SetFocus
        Else
            pcBotao(Index - 1).SetFocus
        End If
        
    ElseIf KeyCode = vbKeyDown Then
        If Index = pcBotao.Count - 1 Then
            pcBotao(0).SetFocus
        Else
            pcBotao(Index + 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcBotao_Click Index
    End If
    
     
    
End Sub
Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
End Sub


Private Sub txtbairro_GotFocus()
    ganhaFoco txtBairro
End Sub

Private Sub txtbairro_Keydown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtbairro_LostFocus()
    AjustaTexto txtBairro
    AtualizaVendaEndereco
End Sub

Private Sub txtbxDescForma_LostFocus()
    If CCur(txtbxDescForma.Text) > CCur(lblDescontoMax.caption) Then
        MsgBox "O Desconto máximo permitido para o usuário com esta forma de pagamento é de " & lblDescontoMax.caption, vbInformation
        txtbxDescForma.Text = lblDescontoMax.caption
        Exit Sub
    End If
End Sub

Private Sub txtbxDesconto_GotFocus()
    ganhaFoco txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    On Error Resume Next
    PerdeFoco txtbxDesconto
   txtbxDescPerc.Text = FormatVal(CCur(txtbxDesconto.Text) * 100 / CCur(lblValor.caption), "###,##0.0000")

    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.Text))

End Sub

Private Sub txtbxDescPerc_GotFocus()
    ganhaFoco txtbxDescPerc
End Sub

Private Sub txtbxDescPerc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub




Private Sub txtbxDescPerc_LostFocus()
    PerdeFoco txtbxDescPerc, "###,##0.0000"
    txtbxDesconto.Text = FormatVal(CCur(txtbxDescPerc.Text) / 100 * (CCur(lblValor.caption) - CCur(lblTotalQueima.caption)))
    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.Text))
    
End Sub


Private Sub txtbxDescPercent_GotFocus()
    ganhaFoco txtbxDescPercent
End Sub
Private Sub pcBotao_Click(Index As Integer)
    If Index = 7 Then
        cmndbtnSair_Click
    ElseIf Index = 6 Then
        cmndbtnSalvar_Click
    ElseIf Index = 5 Then
        cmndbtnCancel_Click
    ElseIf Index = 4 Then
        dtgrd.SetFocus
    ElseIf Index = 3 Then
        
    ElseIf Index = 2 Then
        If chkEntrega.Value = 0 Then
            chkEntrega.Value = 1
        Else
            chkEntrega.Value = 0
        End If
    ElseIf Index = 0 Then
        
        cmndbtnFormasPag_Click
    ElseIf Index = 0 Then

    ElseIf Index = 7 Then
        
    End If
End Sub

Private Sub txtbxDescPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescPercent_LostFocus()
    PerdeFoco txtbxDescPercent
    On Error GoTo trata_Erro
    txtbxDescReais.Text = FormatVal(CCur(txtbxDescPercent.Text) * CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) / 100)
    
    
    
    
    Exit Sub
trata_Erro:
    txtbxDescReais.Text = "0,00"
End Sub

Private Sub txtbxDescReais_GotFocus()
    ganhaFoco txtbxDescReais
End Sub

Private Sub txtbxDescReais_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescReais_LostFocus()
    On Error Resume Next
    PerdeFoco txtbxDescReais
    
    'txtbxDescPercent.Text = FormatVal((CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text))
    txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxQtd.Text) / CCur(txtbxUnit.Text) * 100)
    
    
    p_totalItem
    
    
End Sub

Private Sub txtbxEndereco_Change()

End Sub

Private Sub txtbxFinal_GotFocus()
    ganhaFoco txtbxFinal
End Sub

Private Sub txtbxFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxFinal_LostFocus
        
        insere
        txtbxFinal.Text = "0,00"
        txtbxProduto.SetFocus
    End If
End Sub

Private Sub txtbxFinal_LostFocus()
    On Error GoTo trata_Erro
'    On Error Resume Next
    'txtbxDescReais.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxFinal.Text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    
    
    
    txtbxDescReais.Text = FormatVal((CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text)) - CCur(txtbxFinal.Text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    txtbxDescReais_LostFocus
trata_Erro:
    
End Sub


Private Sub txtbxFormaPag_Change()
    txtbxFormaPag.tag = ""
End Sub

Private Sub txtbxFormaPag_GotFocus()
    ganhaFoco txtbxFormaPag
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxNomeCli_GotFocus()
    ganhaFoco txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    ElseIf KeyCode = vbKeyF2 Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNomeCli_LostFocus()
    AjustaTexto txtbxNomeCli, -1
    If txtbxNomeCli.tag <> "" Then
        P_AtualizaVenda
    End If
End Sub

Private Sub txtbxObs_GotFocus()
    
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub

Private Sub txtbxProduto_Change()
    cmndbtnInserir.tag = ""
    txtbxProduto.tag = ""
    
End Sub

Private Sub txtbxProduto_GotFocus()
    ganhaFoco txtbxProduto
End Sub
Sub buscaProd()
    
    p = Produtos.F_BuscaProdSaida(txtbxProduto.Text)
    If p.cod_pro <> "" Then
    
        txtbxProduto.Text = p.nome_pro
        txtbxProduto.tag = p.cod_pro
        txtbxUnit.Text = FormatVal(p.pv_pro)
        txtbxUnit.tag = FormatVal(p.pv_pro)
        txtbxUnit.Enabled = False
        lblEmpresa.caption = p.empresa
        lblEmpresa.tag = p.cod_emp
        
        If p.controlaestoque_pro = "S" Then
            lblEstoque.caption = p.er_pro & " em estoque"
            lblEstoque.tag = p.er_pro
            lblReservado.caption = p.r_pro & " em reserva"
            lblReservado.tag = p.r_pro
            lblDisponivel.caption = p.er_pro - p.r_pro & " disponível"
            lblDisponivel.tag = p.er_pro - p.r_pro
            lblRefPro.caption = p.codfab_pro
            lblEstoque.Visible = True
        lblDisponivel.Visible = True
        lblReservado.Visible = True
        Else
            lblEstoque.caption = "0"
            lblEstoque.tag = "0"
            lblReservado.caption = "0" & " em reserva"
            lblReservado.tag = p.r_pro
            lblDisponivel.caption = "0" & " disponível"
            lblDisponivel.tag = "0"
            lblRefPro.caption = p.codfab_pro
            lblEstoque.Visible = False
            lblDisponivel.Visible = False
            lblReservado.Visible = False
            txtbxUnit.Enabled = True
            
        End If
        If p.er_pro - p.r_pro > 0 Then
            lblDisponivel.BackColor = vbGreen
            lblDisponivel.ForeColor = vbBlack
            TemEstoque = True
        Else
            lblDisponivel.BackColor = vbRed
            lblDisponivel.ForeColor = vbWhite
            TemEstoque = False
        End If
        
        
                
        
        'lblFracionado.caption = p.fracionado_pro
        
        
        
        If p.pq_pro > 0 And lblTipoPag.tag = "01" Then
            lblQueima.Visible = True
            txtbxDescReais.Enabled = False
            txtbxDescPercent.Enabled = False
            txtbxPQ_pro.Visible = True
            lblPq_pro.Visible = True
            
            txtbxPQ_pro.Text = FormatVal(p.pq_pro)
            txtbxFinal.Locked = True
            txtbxDescReais.Text = FormatVal(p.pv_pro - p.pq_pro)
            txtbxPQ_pro.tag = FormatVal(p.pv_pro - p.pq_pro)
            txtbxDescReais_LostFocus
            
        Else
            lblQueima.Visible = False
            lblPq_pro.Visible = False
            txtbxPQ_pro.Text = "0,00"
            txtbxPQ_pro.Visible = False
            txtbxDescReais.Enabled = True
            txtbxDescPercent.Enabled = True
            txtbxFinal.Locked = False
            txtbxDescPercent.Text = CCur(txtbxDescForma.Text) + CCur(p.sub.descmax) + CCur(p.cat.descmax)
        End If
        
        
        
        txtbxQtd.Enabled = True
        txtbxQtd.SetFocus
        
    End If
    
    
End Sub
Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    Dim p As Produto
    Dim rPesq As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        buscaProd

    ElseIf KeyCode = vbKeyTab Then
        txtbxQtd.Enabled = True
    ElseIf KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        'cmndbtnLocaliza_Click
    ElseIf KeyCode = vbKeyF11 Then
        
        
    End If
    
End Sub

Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim e As EstadoImp
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoQtd As String
    If lblQueima.Visible = False Then
        If CCur(Round(txtbxDescPercent.Text)) > CCur(lblDesc.caption) + CCur(txtbxDescForma.Text) Then
            MsgBox "O desconto máximo permitido para o usuário é de " & CCur(lblDesc.caption) + CCur(txtbxDescForma.Text) & "%", vbExclamation
            Exit Sub
        End If
    Else
        
    End If
    
    
    
    If CCur(txtbxQtd.Text) <= 0 Then
        MsgBox "Quantidade não pode ser zero.", vbExclamation
        txtbxQtd.SetFocus
        Exit Sub
        
    End If
    If CCur(txtbxFinal.Text) <= 0 Then
        MsgBox "Valor do produto não pode ser zero.", vbExclamation
        txtbxFinal.SetFocus
        Exit Sub
        
    End If
    
    If bUsaImpedirVendaDeProdutoSemEstoque = True Then
        If CCur(lblDisponivel.tag) <= 0 Then
            MsgBox "Não é possível adicionar este produto. Estoque insuficiente!", vbExclamation
            Exit Sub
        End If
        
        If CCur(lblDisponivel.tag) - CCur(txtbxQtd.Text) < 0 Then
            MsgBox "Não é possível adicionar este produto. Estoque insuficiente!", vbExclamation
            Exit Sub
        End If
        
    End If
    
    If txtbxProduto.Text = "" Then
        MsgBox "Informe pelo menos parte do nome do produto para que seja localizado!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    On Error GoTo trata_Erro
                
        If AdoVendas.RecordCount >= 999 Then
            MsgBox "Não é possível adicionar mais itens nessa venda! (999)", vbExclamation
            Exit Sub
        End If
        
        If cmndbtnInserir.tag = "" Then
            iItem = iItem + 1
        Else
            
        End If
                
                p.pv_pro = CCur(txtbxUnit.Text)
                If vendas.F_AdicionaItem(Me.tag, p, txtbxQtd.Text, txtbxDescReais.Text, txtbxFinal.Text, txtbxPQ_pro.Text, lblEmpresa.tag) = "" Then
                    MsgBox "Erro na inserção do item!", vbExclamation
                    
                Else
                    'If Not Produtos.F_AlteraEstoquePro(p.cod_pro, "-" & txtbxQtd.Text) Then
                    '    MsgBox "Erro na baixa do estoque.", vbExclamation
                    'End If
                    
                    
                End If
                    
                P_buscaitens
                Me.Refresh
                txtbxProduto.tag = ""
                
                txtbxQtd.Text = "1"
                'AdoVendas.Move AdoVendas.RecordCount - 1
                txtbxUnit.Text = "0"
                txtbxDescPercent.Text = "0,0"
                txtbxDescReais.Text = "0,00"
                
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
                'lblEstoque.Visible = False
        
        P_AtualizaVenda
        On Error Resume Next

        Exit Sub
trata_Erro:
    MostraErro "Insere()"
End Sub
Private Sub txtbxQtd_GotFocus()
    
    ganhaFoco txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
Exit Sub
trata_Erro:
    MostraErro
End Sub

Private Sub txtbxQtd_LostFocus()
    On Error Resume Next
    'If lblFracionado = "S" Then
    '    PerdeFoco txtbxQtd, "###,##0.00"
    'Else
    '    PerdeFoco txtbxQtd, "###"
    'End If
    If txtbxPQ_pro.Visible = True Then
        txtbxDescReais.Text = FormatVal(CCur(txtbxPQ_pro.tag) * CCur(txtbxQtd.Text))
        txtbxDescReais_LostFocus
    End If
    p_totalItem
    
        
    
End Sub


Private Sub txtbxUnit_GotFocus()
    P_SelecTexto txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxUnit_LostFocus()
    PerdeFoco txtbxUnit, "###,##0.00"
    
    p_totalItem
End Sub
Sub P_GravaItemCancelado()
    On Error Resume Next
    Dim strSQL As String
    strSQL = "insert into tab_itvndcncl(seq_itv,num_ven,seq_ite,qtd_ite,unit_ite,cod_cai,cod_tur)values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_itvndcncl", "seq_itv", "", "000000#") & "',"
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "" & AdoVendas!seq_ite & ","
    strSQL = strSQL & F_GravaDec(AdoVendas!qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(AdoVendas!unit_ite) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "') "
    conn.Execute strSQL
    
    
    
End Sub




Private Sub txtComprador_GotFocus()
    P_SelecTexto txtComprador
End Sub

Private Sub txtComprador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_AtualizaVenda
        txtbxProduto.SetFocus
    End If
    
    
End Sub

Private Sub txtComprador_LostFocus()
    P_AjustaTexto txtComprador, nada_muda
    AtualizaVendaEndereco
End Sub

Private Sub txtEnd_GotFocus()
    ganhaFoco txtEnd
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtEnd_LostFocus()
    AjustaTexto txtEnd
    AtualizaVendaEndereco
End Sub

Private Sub txtfone_GotFocus()
    ganhaFoco txtFone
End Sub

Private Sub txtfone_Keydown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtfone_LostFocus()
    AjustaTexto txtFone
    AtualizaVendaEndereco
End Sub

Private Sub txtObsEntrega_GotFocus()
    ganhaFoco txtObsEntrega
End Sub

Private Sub txtObsEntrega_LostFocus()
    AjustaTexto txtObsEntrega, nada_muda
    AtualizaVendaEndereco
    
End Sub

Private Sub txtTel_GotFocus()
    ganhaFoco txtTel
End Sub

Private Sub txtTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        Sendkeys Chr(9)
        P_AtualizaVenda
    End If
    
End Sub

Private Sub txtTel_LostFocus()
    AjustaTexto txtTel
    
    If AdoVendas.RecordCount > 0 Then
    
        P_AtualizaVenda
    End If
End Sub
