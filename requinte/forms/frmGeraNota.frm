VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGeraNota 
   Caption         =   "Nota Fiscal de Saída"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13350
   Icon            =   "frmGeraNota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   13350
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13320
      TabIndex        =   119
      Top             =   0
      Width           =   13350
      Begin VB.TextBox txtbxNatOp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   127
         Text            =   "VENDAS MERC. ESTADUAL"
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtbxCFOP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   126
         Text            =   "5102"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtbxIESubTrib 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   125
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   10080
         ScaleHeight     =   705
         ScaleWidth      =   3105
         TabIndex        =   120
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtbxEmiss 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1560
            TabIndex        =   122
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtbxNota_Sai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   390
            Left            =   120
            MaxLength       =   8
            TabIndex        =   121
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblEmissao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Emissão:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   240
            Left            =   1560
            TabIndex        =   124
            Top             =   0
            Width           =   840
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Número Nota:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   240
            Left            =   120
            TabIndex        =   123
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmndbtnCfop 
         Height          =   285
         Left            =   720
         Picture         =   "frmGeraNota.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Natureza da Operação:"
         Height          =   195
         Left            =   1080
         TabIndex        =   131
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CFOP:"
         Height          =   195
         Left            =   120
         TabIndex        =   130
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.E. do Subst. Trib."
         Height          =   195
         Left            =   5160
         TabIndex        =   129
         Top             =   120
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13320
      TabIndex        =   110
      Top             =   9180
      Width           =   13350
      Begin VB.CommandButton cmndbtnFrete 
         Caption         =   "&Frete/Dados Add"
         Height          =   615
         Left            =   3480
         Picture         =   "frmGeraNota.frx":11CC
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEspelho 
         Caption         =   "&Espelho de Nota"
         Height          =   615
         Left            =   5160
         Picture         =   "frmGeraNota.frx":1316
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnImprNota 
         Caption         =   "&Imprimir Nota"
         Height          =   615
         Left            =   6840
         Picture         =   "frmGeraNota.frx":1460
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Gravar"
         Height          =   615
         Left            =   8520
         Picture         =   "frmGeraNota.frx":15AA
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   10200
         Picture         =   "frmGeraNota.frx":16F4
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox txtNfe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Exportar NFe"
         Height          =   615
         Left            =   1800
         Picture         =   "frmGeraNota.frx":183E
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "NFE:"
         Height          =   195
         Left            =   120
         TabIndex        =   118
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   1300
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   13320
      TabIndex        =   66
      Top             =   7875
      Width           =   13350
      Begin VB.CommandButton cmndbtnDesconto 
         Caption         =   "De&sconto:"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   97
         Top             =   625
         Width           =   1095
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Enabled         =   0   'False
         ForeColor       =   &H0000C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   96
         Text            =   "0,00"
         Top             =   915
         Width           =   1095
      End
      Begin VB.TextBox txtbxDesp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   5280
         TabIndex        =   95
         Text            =   "0,00"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxTotalNot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   6600
         TabIndex        =   94
         Text            =   "0,00"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxTotalProd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   6600
         TabIndex        =   93
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxVIcms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   2640
         TabIndex        =   92
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxBC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   1320
         TabIndex        =   90
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxBCSubs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   3960
         TabIndex        =   89
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxVlrSubs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   5280
         TabIndex        =   88
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtbxValorFre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   1320
         TabIndex        =   87
         Text            =   "0,00"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxValorIpi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   3960
         TabIndex        =   86
         Text            =   "0,00"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxValorSeg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   2640
         TabIndex        =   85
         Text            =   "0,00"
         Top             =   840
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   8040
         ScaleHeight     =   1185
         ScaleWidth      =   3705
         TabIndex        =   68
         Top             =   0
         Width           =   3735
         Begin VB.TextBox txtbxCupom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   70
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtbxVenda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmndbtnVenda 
            Height          =   285
            Left            =   1080
            Picture         =   "frmGeraNota.frx":1BC8
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnCupom 
            Height          =   285
            Left            =   1080
            Picture         =   "frmGeraNota.frx":1F52
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cupom:"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Venda:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   0
            Width           =   510
         End
         Begin VB.Label lblCUPOM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não"
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
            Left            =   2400
            TabIndex        =   82
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CUPOM:"
            Height          =   195
            Left            =   1560
            TabIndex        =   81
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não"
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
            Left            =   2400
            TabIndex        =   80
            Top             =   120
            Width           =   360
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
            Height          =   195
            Left            =   1560
            TabIndex        =   79
            Top             =   120
            Width           =   390
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não"
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
            Left            =   2040
            TabIndex        =   78
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caixa:"
            Height          =   195
            Left            =   1560
            TabIndex        =   77
            Top             =   840
            Width           =   435
         End
         Begin VB.Label lblTurno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não"
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
            Left            =   3000
            TabIndex        =   76
            Top             =   840
            Width           =   360
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Turno:"
            Height          =   195
            Left            =   2520
            TabIndex        =   75
            Top             =   840
            Width           =   465
         End
         Begin VB.Label lblAtendente 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Não"
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
            Left            =   2400
            TabIndex        =   74
            Top             =   600
            Width           =   360
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Atendente:"
            Height          =   195
            Left            =   1560
            TabIndex        =   73
            Top             =   600
            Width           =   780
         End
      End
      Begin VB.TextBox txtbxPicms 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   120
         TabIndex        =   91
         Text            =   "12"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnIcms 
         Height          =   360
         Left            =   480
         Picture         =   "frmGeraNota.frx":22DC
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Transforma todos os itens para o imposto selecionado"
         Top             =   240
         Width           =   355
      End
      Begin VB.CommandButton cmndbtnOriginal 
         Height          =   360
         Left            =   860
         Picture         =   "frmGeraNota.frx":2453
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Transforma todos os itens para o Imposto Original"
         Top             =   240
         Width           =   355
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Despesas Acess:"
         Height          =   195
         Left            =   5280
         TabIndex        =   109
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total da Nota:"
         Height          =   195
         Left            =   6600
         TabIndex        =   108
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Produtos"
         Height          =   195
         Left            =   6600
         TabIndex        =   107
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS(%)"
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Base Cálc Icms.:"
         Height          =   195
         Left            =   1320
         TabIndex        =   105
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Icms:"
         Height          =   195
         Left            =   2760
         TabIndex        =   104
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&B.Cálc Icms Sub:"
         Height          =   195
         Left            =   4080
         TabIndex        =   103
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Icms Subs:"
         Height          =   195
         Left            =   5280
         TabIndex        =   102
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do &Seguro:"
         Height          =   195
         Left            =   2640
         TabIndex        =   101
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do &Frete:"
         Height          =   195
         Left            =   1320
         TabIndex        =   100
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do &IPI:"
         Height          =   195
         Left            =   3960
         TabIndex        =   99
         Top             =   600
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture5 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0071CACE&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   13320
      TabIndex        =   0
      Top             =   735
      Width           =   13350
      Begin VB.TextBox txtbxCNPJDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtbxEndDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtbxBairroDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtbxCepDesp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10680
         TabIndex        =   8
         Text            =   "73.900-000"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtbxFoneDet 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8760
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxIeDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxUFDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10200
         TabIndex        =   5
         Text            =   "GO"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtbxNomeDest 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Tag             =   "00000"
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   14
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtComplemento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBuscaCli 
         Height          =   285
         Left            =   4800
         Picture         =   "frmGeraNota.frx":2530
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CNPJ/CPF:"
         Height          =   195
         Left            =   5160
         TabIndex        =   26
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/Razão Social"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro/Distrito:"
         Height          =   195
         Left            =   5160
         TabIndex        =   23
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP:"
         Height          =   195
         Left            =   10680
         TabIndex        =   22
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone/Fax:"
         Height          =   195
         Left            =   8760
         TabIndex        =   21
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc. Est:"
         Height          =   195
         Left            =   7080
         TabIndex        =   20
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Height          =   195
         Left            =   10200
         TabIndex        =   19
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc. Prod. Rural"
         Height          =   195
         Left            =   10200
         TabIndex        =   18
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label lblInscrProdRural 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
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
         Left            =   10200
         TabIndex        =   17
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblTipoPessoa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F"
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
         TabIndex        =   16
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         Height          =   195
         Left            =   8160
         TabIndex        =   13
         Top             =   600
         Width           =   540
      End
   End
   Begin TabDlg.SSTab sstb 
      Height          =   5295
      Left            =   0
      TabIndex        =   27
      Top             =   2280
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Produtos"
      TabPicture(0)   =   "frmGeraNota.frx":28BA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblItens"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label50"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label52"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label54"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label55"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdProduto"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "msflxGrd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtbxGrid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtFinal"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtUnit"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtQtd"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProduto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdOk"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdEdit"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdExcl"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Dados Adicionais/ Frete"
      TabPicture(1)   =   "frmGeraNota.frx":28D6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdExcl 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   11040
         TabIndex        =   143
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   10080
         TabIndex        =   142
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   8160
         TabIndex        =   141
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   135
         Top             =   4800
         Width           =   4815
      End
      Begin VB.TextBox txtQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5280
         TabIndex        =   134
         Text            =   "1"
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   133
         Text            =   "0,00"
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox txtFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
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
         Left            =   6840
         TabIndex        =   132
         Text            =   "0,00"
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtbxGrid 
         Height          =   405
         Left            =   1200
         TabIndex        =   63
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "&Transportadora / Volumes Transportados"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   11415
         Begin VB.TextBox txtbxNomeFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtbxPlacaFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   47
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtbxUfVeiFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5520
            TabIndex        =   46
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtbxCnpjFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            TabIndex        =   45
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtbxEnderecoFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtbxMuniFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   43
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txtbxIEFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6840
            TabIndex        =   42
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtbxUfFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6360
            TabIndex        =   41
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtbxQtdFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtbxEspFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   39
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtbxMarcaFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   38
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtbxNumFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            TabIndex        =   37
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtbxPBrutoFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5520
            TabIndex        =   36
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtbxPLiqFre 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6840
            TabIndex        =   35
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Nome/Razão Social"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Placa Veíc:"
            Height          =   195
            Left            =   4440
            TabIndex        =   61
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   5520
            TabIndex        =   60
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF:"
            Height          =   195
            Left            =   6000
            TabIndex        =   59
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Município"
            Height          =   195
            Left            =   3480
            TabIndex        =   57
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Est:"
            Height          =   195
            Left            =   6840
            TabIndex        =   56
            Top             =   840
            Width           =   660
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   6360
            TabIndex        =   55
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   1440
            Width           =   870
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
            Height          =   195
            Left            =   1200
            TabIndex        =   53
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
            Height          =   195
            Left            =   3000
            TabIndex        =   52
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   4440
            TabIndex        =   51
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "P.Bruto:"
            Height          =   195
            Left            =   5520
            TabIndex        =   50
            Top             =   1440
            Width           =   570
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "P.Líq:"
            Height          =   195
            Left            =   6840
            TabIndex        =   49
            Top             =   1440
            Width           =   435
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frete por Conta:"
         Height          =   1455
         Left            =   -66600
         TabIndex        =   30
         Top             =   2640
         Width           =   3135
         Begin VB.TextBox txtbxContaFre 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "1 - Emitente"
            Height          =   195
            Left            =   720
            TabIndex        =   33
            Top             =   285
            Width           =   840
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "2 - Destinatário"
            Height          =   195
            Left            =   1920
            TabIndex        =   32
            Top             =   285
            Width           =   1065
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados Adicionais:"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   28
         Top             =   2640
         Width           =   8175
         Begin VB.TextBox txtbxDadosAdd 
            Height          =   1125
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   7935
         End
      End
      Begin MSFlexGridLib.MSFlexGrid msflxGrd 
         Height          =   3855
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6800
         _Version        =   393216
         AllowUserResizing=   3
         Appearance      =   0
      End
      Begin VB.CommandButton cmdProduto 
         Height          =   285
         Left            =   4920
         Picture         =   "frmGeraNota.frx":28F2
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label55 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5280
         TabIndex        =   139
         Top             =   4560
         Width           =   300
      End
      Begin VB.Label Label54 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto (F2)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   138
         Top             =   4560
         Width           =   870
      End
      Begin VB.Label Label52 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5880
         TabIndex        =   137
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label Label50 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor Final:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   136
         Top             =   4560
         Width           =   780
      End
      Begin VB.Label lblItens 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Itens:"
         Height          =   195
         Left            =   11400
         TabIndex        =   65
         Top             =   4320
         Width           =   390
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnTrib 
         Caption         =   "&Tributar Todas"
      End
      Begin VB.Menu mnNaoTrib 
         Caption         =   "&Não Tributar Todas"
      End
   End
End
Attribute VB_Name = "frmGeraNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bBuscaCupom As Boolean
Dim bActive As Boolean
Dim seq_sai As String
Public Sub setSeq_Sai(valor As String)
    seq_sai = valor
End Sub
Private Sub chkClick_Click()

End Sub

Private Sub cmndbtnArred_Click()
    'P_valorlostfocus txtbxBC
    'P_valorlostfocus txtbxBCSubs
    'P_valorlostfocus txtbxVIcms
    'P_valorlostfocus txtbxVlrSubs
    'P_valorlostfocus txtbxTotalNot
    'P_valorlostfocus txtbxTotalProd
End Sub

Private Sub cmndbtnClose_Click()
    Me.tag = ""
    Me.Hide
    
End Sub

Private Sub cmndbtnConf_Click()
    If MsgBox("Confirma conclusão da NF?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    sNum = MaxKeySQL("tab_saidas", "nota_sai", "", "0000000#")
    txtbxNota_Sai.Text = sNum
    p_imprNF Printer, False
    If MsgBox("A Impressão saiu OK?", vbYesNo + vbQuestion) = vbNo Then
        If MsgBox("A Impressora está pronta para imprimir novamente?", vbOKCancel) = vbOK Then
            cmndbtnConf_Click
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    If cmndbtnConf.tag = "" Then
        grava
    Else
        altera
    End If
    Me.tag = "x"
    Me.Hide
    
    
    'Exit Sub
   
    
End Sub
Sub p_imprNF(obj As Object, Optional Entrada As Boolean = False)
    Dim iCount As Integer, sProd As String
    Dim sAltura As Single
    Dim sLargura As Single
    obj.fontsize = "10"
    obj.FontName = "Arial"
    Dim sYAj As Single
    sYAj = -15
    'On Error Resume Next
    sAltura = 210
    sLargura = 160 '17 cm
    obj.ScaleMode = vbMillimeters
'    Printer.PaperSize = 256
'    sAltura = obj.ScaleY(sAltura, vbMillimeters, vbTwips)
'    sLargura = obj.ScaleX(sLargura, vbMillimeters, vbTwips)
    
    Dim rPesq As New ADODB.Recordset
        strSQL = "select * from tab_confignf "
        rPesq.Open strSQL, Conn
        With rPesq
        obj.fontsize = 12
        'Printer.BackStyle = 0
       ' Printer.PaperSize = 11 'a 5
        P_PrintXY obj, txtbxNota_Sai.Text, !numx_con, !numy_con + sYAj
        P_PrintXY obj, "X", IIf(Not Entrada, !marcax_con, !marcax_con + 15), !marcay_con + sYAj
        obj.FontName = "arial"
        obj.fontsize = 9
        P_PrintXY obj, txtbxNatOp.Text, !natopx_con, !natopy_con + sYAj
        P_PrintXY obj, txtbxCFOP.Text, !cfopx_con, !cfopy_con + sYAj
        P_PrintXY obj, txtbxIESubTrib.Text, !iex_con, !iey_con + sYAj
        P_PrintXY obj, txtbxNomeDest.Text, !nomedestx_con, !nomedesty_con + sYAj
        P_PrintXY obj, txtbxCNPJDest.Text, !cnpjx_con, !cnpjy_con + sYAj
        P_PrintXY obj, txtbxEmiss.Text, !emissx_con, !emissy_con + sYAj
        P_PrintXY obj, txtbxEndDest.Text, !enddestx_con, !enddesty_con + sYAj
        P_PrintXY obj, txtbxBairroDest.Text, !baidestx_con, !baidesty_con + sYAj
        P_PrintXY obj, txtbxCepDesp.Text, !cepdestx_con, !cepdesty_con + sYAj
        'P_PrintXY obj, txtbxDataSai.Text, 185, 50
        P_PrintXY obj, txtbxMuniDest.Text, !munidestx_con, !munidesty_con + sYAj
        P_PrintXY obj, txtbxFoneDet.Text, !fonedestx_con, !fonedesty_con + sYAj
        P_PrintXY obj, txtbxUFDest.Text, !ufdestx_con, !ufdesty_con + sYAj
        P_PrintXY obj, txtbxIeDest.Text, !iedestx_con, !iedesty_con + sYAj
    '    P_PrintXY obj, txtbxHoraSai.Text, 185, 57
        With msflxgrd
            obj.CurrentY = rPesq!iniprody_con + sYAj
            For iCount = 1 To .rows - 1
                .row = iCount
                If iItem = 20 Then
                    iItem = 0
                    On Error Resume Next
                    obj.NewPage
                    
                    MsgBox "Clique em OK quando a impressora estiver pronta!", vbInformation
                    With rPesq
                        P_PrintXY obj, txtbxNota_Sai.Text, !numx_con, !numy_con + sYAj
                        P_PrintXY obj, "X", IIf(Not Entrada, !marcax_con, !marcax_con + 15), !marcay_con + sYAj
                        obj.FontName = "arial"
                        obj.fontsize = 9
                        P_PrintXY obj, txtbxNatOp.Text, !natopx_con, !natopy_con + sYAj
                        P_PrintXY obj, txtbxCFOP.Text, !cfopx_con, !cfopy_con + sYAj
                        P_PrintXY obj, txtbxIESubTrib.Text, !iex_con, !iey_con + sYAj
                        P_PrintXY obj, txtbxNomeDest.Text, !nomedestx_con, !nomedesty_con + sYAj
                        P_PrintXY obj, txtbxCNPJDest.Text, !cnpjx_con, !cnpjy_con + sYAj
                        P_PrintXY obj, txtbxEmiss.Text, !emissx_con, !emissy_con + sYAj
                        P_PrintXY obj, txtbxEndDest.Text, !enddestx_con, !enddesty_con + sYAj
                        P_PrintXY obj, txtbxBairroDest.Text, !baidestx_con, !baidesty_con + sYAj
                        P_PrintXY obj, txtbxCepDesp.Text, !cepdestx_con, !cepdesty_con + sYAj
                        'P_PrintXY obj, txtbxDataSai.Text, 185, 50
                        P_PrintXY obj, txtbxMuniDest.Text, !munidestx_con, !munidesty_con + sYAj
                        P_PrintXY obj, txtbxFoneDet.Text, !fonedestx_con, !fonedesty_con + sYAj
                        P_PrintXY obj, txtbxUFDest.Text, !ufdestx_con, !ufdesty_con + sYAj
                        P_PrintXY obj, txtbxIeDest.Text, !iedestx_con, !iedesty_con + sYAj
                    End With
                    
                    obj.CurrentY = rPesq!iniprody_con + sYAj
                    
                    
                End If
                sProd = ""
                .Col = 0
                P_ImprX obj, CStr(iCount), 0 '
                P_ImprX obj, .Text, rPesq!codprod_con - 2
                .Col = 1
                P_ImprX obj, Left(.Text, 35), rPesq!descprod_con + 1
                .Col = 2
                P_ImprX obj, .Text, rPesq!stprod_con
                .Col = 3
                P_ImprX obj, .Text, rPesq!unid_con
                .Col = 4
                P_ImprX obj, .Text, rPesq!qtdprod_con, True, True
                .Col = 5
                P_ImprX obj, .Text, rPesq!unitprod_con, True, True
                .Col = 6
                P_ImprX obj, .Text, rPesq!totalprod_con, True, True
                .Col = 7
                P_ImprX obj, .Text, rPesq!aliqprod_con, False
                'obj.CurrentY = obj.CurrentY - 0.7
                iItem = iItem + 1
            Next
        End With
        obj.CurrentY = 240
        
        P_PrintXY obj, txtbxBC.Text, !bcx_con, !bcy_con + sYAj
        P_PrintXY obj, txtbxVIcms.Text, !vlricmsx_con, !vlricmsy_con + sYAj
        P_PrintXY obj, txtbxBCSubs.Text, !bcsubsx_con, !bcsubsy_con + sYAj
        P_PrintXY obj, txtbxVlrSubs.Text, !vlrsubsx_con, !vlrsubsy_con + sYAj
        P_PrintXY obj, txtbxTotalProd.Text, !vlrprodx_con, !vlrprody_con + sYAj
        P_PrintXY obj, txtbxValorFre.Text, !vlrfretex_con, !vlrfretey_con + sYAj
        P_PrintXY obj, txtbxValorSeg.Text, !vlrsegx_con, !vlrsegy_Con + sYAj
        P_PrintXY obj, txtbxDesp.Text, !outrasdespx_con, !outrasdespy_con + sYAj
        P_PrintXY obj, txtbxValorIpi.Text, !vlripix_con, !vlripiy_con + sYAj
        P_PrintXY obj, txtbxTotalNot.Text, !vlrtotalx_con, !vlrtotaly_con + sYAj
        P_PrintXY obj, txtbxNota_Sai.Text, !numrodapex_con, !numrodapey_con + sYAj
        
        
            P_PrintXY obj, txtbxNomeFre.Text, !nomefrex_con, !nomefrey_con + sYAj
            P_PrintXY obj, txtbxContaFre.Text, !tipofrex_con, !tipofrey_con + sYAj
            P_PrintXY obj, txtbxPlacaFre.Text, !placafrex_con, !placafrey_con + sYAj
            P_PrintXY obj, txtbxUfVeiFre.Text, !ufveix_con, !ufveiy_con + sYAj
            P_PrintXY obj, txtbxCnpjFre.Text, !cnpjfrex_con, !cnpjfrey_con + sYAj
            P_PrintXY obj, txtbxEnderecoFre.Text, !endfrex_con, !endfrey_con + sYAj
            P_PrintXY obj, txtbxMuniFre.Text, !munifrex_con, !munifrey_con + sYAj
            P_PrintXY obj, txtbxUfFre.Text, !uffrex_con, !uffrey_con + sYAj
            P_PrintXY obj, txtbxIEFre.Text, !iefrex_con, !iefrey_con + sYAj
            P_PrintXY obj, txtbxQtdFre.Text, !qtdfrex_con, !qtdfrey_con + sYAj
            P_PrintXY obj, txtbxEspFre.Text, !espfrex_con, !espfrey_con + sYAj
            P_PrintXY obj, txtbxMarcaFre.Text, !marcafrex_con, !marcafrey_con + sYAj
            P_PrintXY obj, txtbxNumFre.Text, !numfrex_con, !numfrey_con + sYAj
            P_PrintXY obj, txtbxPBrutoFre.Text, !pbrutofrex_con, !pbrutofrey_con + sYAj
            P_PrintXY obj, txtbxPLiqFre.Text, !pliqfrex_con, !pliqfrey_con + sYAj
            P_PrintXY obj, txtbxDadosAdd.Text, !dadosaddx_con, !dadosaddy_con + sYAj
            If TypeName(obj) = "Printer" Then
                Printer.EndDoc
            End If
        
        .Close
    End With
End Sub
Sub p_imprNFLPT(Optional Entrada As Boolean = False)
    Dim iCount As Integer, sProd As String
    Dim sAltura As Single
    Dim sLargura As Single
    Dim obj As Object
    Dim sYAj As Single
    Dim Margem As Integer
    Dim sDesc As String
    Dim iItem As Integer
    Margem = 0
    Dim rPesq As New ADODB.Recordset
        strSQL = "select * from tab_confignf "
        rPesq.Open strSQL, Conn
        With rPesq
        If Not F_OpenLpt Then
            Exit Sub
        End If
        P_FeedLpt1
        P_FontLpt "M", "4"
        P_PrintLpt "X", IIf(Not Entrada, 43, 48), True
        P_PrintLpt txtbxNota_Sai.Text, 18, False
        P_FeedLpt1 2
        
        
        P_PrintLpt F_EsticaTexto(txtbxNatOp.Text, 20) & " " & F_EsticaTexto(txtbxCFOP, 5) & F_EsticaTexto(txtbxIESubTrib.Text, 13), 0, False
        P_FeedLpt1 1
        P_PrintLpt F_EsticaTexto(txtbxNomeDest.Text, 43) & " " & F_EsticaTexto(txtbxCNPJDest.Text, 16) & F_EsticaTexto(txtbxEmiss.Text, 13), 0, False
        P_FeedLpt1 1
        P_PrintLpt F_EsticaTexto(txtbxEndDest.Text, 31) & " " & F_EsticaTexto(txtbxBairroDest.Text, 18) & " " & F_EsticaTexto(txtbxCepDesp.Text, 10), 0, False
        
        'P_FeedLpt1 1
        sDesc = F_EsticaTexto(txtbxMuniDest.Text, 26) & " "
        sDesc = sDesc & F_EsticaTexto(txtbxFoneDet.Text, 12) & " "
        sDesc = sDesc & F_EsticaTexto(txtbxUFDest.Text, 2) & " "
        sDesc = sDesc & F_EsticaTexto(txtbxIeDest.Text, 12)
        P_PrintLpt sDesc, 0, False
        P_FeedLpt1 1
        With msflxgrd
            
            For iCount = 1 To .rows - 1
                .row = iCount
                If iItem = 15 Then
                    
                    iItem = 0
                    P_FeedLpt1 26
                    P_PrintLpt "X", IIf(Not Entrada, 43, 48), True
                    P_PrintLpt txtbxNota_Sai.Text, 18, False
                    P_FeedLpt1 2
        
        
                    P_PrintLpt F_EsticaTexto(txtbxNatOp.Text, 20) & " " & F_EsticaTexto(txtbxCFOP, 5) & F_EsticaTexto(txtbxIESubTrib.Text, 13), 0, False
                    P_FeedLpt1 1
                    P_PrintLpt F_EsticaTexto(txtbxNomeDest.Text, 43) & " " & F_EsticaTexto(txtbxCNPJDest.Text, 16) & F_EsticaTexto(txtbxEmiss.Text, 13), 0, False
        
                    P_PrintLpt F_EsticaTexto(txtbxEndDest.Text, 31) & " " & F_EsticaTexto(txtbxBairroDest.Text, 18) & " " & F_EsticaTexto(txtbxCepDesp.Text, 6), 0, False
        
                    P_FeedLpt1 1
                    sDesc = F_EsticaTexto(txtbxMuniDest.Text, 26) & " "
                    sDesc = sDesc & F_EsticaTexto(txtbxFoneDet.Text, 12) & " "
                    sDesc = sDesc & F_EsticaTexto(txtbxUFDest.Text, 2) & " "
                    sDesc = sDesc & F_EsticaTexto(txtbxIeDest.Text, 12)
                    P_PrintLpt sDesc, 0, False
                    P_PrintLpt "Continuação NF: " & txtbxNota_Sai.Text, 0, False
                    
                End If
                sProd = ""
                .Col = 0
                
                sDesc = F_EsticaTexto(.Text, 8) & " "
                .Col = 1
                sDesc = sDesc & F_EsticaTexto(.Text, 26) & " "
                .Col = 2
                sDesc = sDesc & F_EsticaTexto(.Text, 2) & " "
                .Col = 3
                sDesc = sDesc & F_EsticaTexto(.Text, 3) & " "
                .Col = 4
                sDesc = sDesc & F_EsticaTexto(.Text, 5, 2) & " "
                .Col = 5
                sDesc = sDesc & F_EsticaTexto(.Text, 7, 2) & " "
                .Col = 6
                sDesc = sDesc & F_EsticaTexto(.Text, 10, 2) & " "
                .Col = 7
                sDesc = sDesc & F_EsticaTexto(.Text, 4)
                P_PrintLpt sDesc, 0, False
                
                iItem = iItem + 1
            Next
        End With
        P_FeedLpt1 1
        P_FeedLpt1 16 - iItem
        sDesc = F_EsticaTexto(txtbxBC.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxVIcms.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxBCSubs.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxVlrSubs.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxTotalProd, 18, 2)
        P_PrintLpt sDesc, 0, False
        
        sDesc = F_EsticaTexto(txtbxValorFre.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxValorSeg.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxDesp.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxValorIpi.Text, 12, 2)
        sDesc = sDesc & F_EsticaTexto(txtbxTotalNot.Text, 18, 2)
        P_PrintLpt sDesc, 0, False
        
        
        
        
        If txtbxNomeFre.Text <> "" Or txtbxDadosAdd.Text <> "" Then
        
            sDesc = F_EsticaTexto(txtbxNomeFre.Text, 20) & Space(3)
            sDesc = sDesc & F_EsticaTexto(txtbxContaFre.Text, 2)
            sDesc = sDesc & F_EsticaTexto(txtbxPlacaFre.Text, 8)
            sDesc = sDesc & F_EsticaTexto(txtbxUfVeiFre.Text, 2)
            sDesc = sDesc & F_EsticaTexto(txtbxCnpjFre.Text, 16)
            P_PrintLpt sDesc, 0, False
            
            sDesc = F_EsticaTexto(txtbxEnderecoFre.Text, 20)
            sDesc = sDesc & F_EsticaTexto(txtbxMuniFre.Text, 14)
            sDesc = sDesc & F_EsticaTexto(txtbxUfFre.Text, 3)
            sDesc = sDesc & F_EsticaTexto(txtbxIEFre.Text, 12)
            P_PrintLpt sDesc, 0, False
            
            sDesc = F_EsticaTexto(txtbxQtdFre.Text, 20)
            sDesc = sDesc & F_EsticaTexto(txtbxEspFre.Text, 14)
            sDesc = sDesc & F_EsticaTexto(txtbxMarcaFre.Text, 3)
            sDesc = sDesc & F_EsticaTexto(txtbxNumFre.Text, 12)
            sDesc = sDesc & F_EsticaTexto(txtbxPBrutoFre.Text, 3)
            sDesc = sDesc & F_EsticaTexto(txtbxPLiqFre.Text, 12)
            P_PrintLpt sDesc, 0, False
            
            P_FeedLpt1
            P_PrintLpt F_Quebra(txtbxDadosAdd.Text, 14), 0, False
            P_FeedLpt1 4
        Else
            P_FeedLpt1 16
        End If
        .Close
        P_FeedLpt1 3
        P_PrintLpt F_EsticaTexto(txtbxNota_Sai.Text, 65, 2)
        
    End With
    
    If Not F_CloseLpt Then
            Exit Sub
        End If
End Sub

Sub imprimeEspelho(obj As Object)
    Dim iCount As Integer
    Dim iItem As Integer
    Dim sProd As String
    P_CabRelTimbr "Espelho de Nota Fiscal: " & txtbxNota_Sai & "-" & txtbxNomeDest.Text, obj
    P_ImprLinha obj
    P_ImprX obj, "Número: " & txtbxNota_Sai.Text, 0
    P_ImprX obj, "Base Cálc: " & txtbxBC.Text, 32
    P_ImprX obj, "ICMS(%): " & txtbxPicms.Text & "%", 62
    P_ImprX obj, "Vlr ICMS: " & txtbxVIcms.Text, 87
    P_ImprX obj, "Desp: " & txtbxDesp.Text, 115
    P_ImprX obj, "Produtos: " & txtbxTotalProd, 140
    P_ImprX obj, "Total: " & txtbxTotalProd, 175, False
    P_ImprLinha obj
    P_ImprX obj, "Cód", 0
    P_ImprX obj, "Descrição dos Produtos", 20
    P_ImprX obj, "Un.", 140
    P_ImprX obj, "Qtd.", 160, True, True
    P_ImprX obj, "Unit.", 175, True, True
    P_ImprX obj, "Total", 200, False, True
    P_ImprLinha obj
    With msflxgrd
        
        For iCount = 1 To .rows - 1
            .row = iCount
            If iItem = 45 Then
                iItem = 0
                If TypeName(obj) = "Printer" Then
                    P_Rodape strNome_Emp, True, "Continua..."
                End If
                P_ImprX obj, "Cód", 0
                P_ImprX obj, "Descrição dos Produtos", 20
                P_ImprX obj, "Un.", 140
                P_ImprX obj, "Qtd.", 160, True, True
                P_ImprX obj, "Unit.", 175, True, True
                P_ImprX obj, "Total", 200, False, True
                P_ImprLinha obj
                
            End If
            sProd = ""
            .Col = 0
            P_ImprX obj, .Text, 0
            .Col = 1
            P_ImprX obj, .Text, 20
            .Col = 3
            P_ImprX obj, .Text, 140
            .Col = 4
            P_ImprX obj, .Text, 160, True, True
            .Col = 5
            P_ImprX obj, .Text, 175, True, True
            .Col = 6
            P_ImprX obj, .Text, 200, False, True
            
            
            iItem = iItem + 1
        Next
        P_ImprLinha obj
        'P_ImprX obj, lblTotalArt.caption, 160, True, False
            
    End With
    
    
End Sub
Sub buscaTerceiro()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error Resume Next
    strSQL = "select * from tab_fornecedores where cod_for='" & txtbxNomeDest.tag & "' "
    rPesq.Open strSQL, Conn
    With rPesq
        If .RecordCount = 0 Then Exit Sub
        txtbxBairroDest.Text = !bairro_for
        txtbxCepDesp.Text = !cep_for
        txtbxCNPJDest.Text = !cnpj_for
        txtbxEndDest.Text = !endereco_for
        txtbxFoneDet.Text = !telefone_for
        txtbxIeDest.Text = !ie_for
        txtbxMuniDest.Text = !cidade_For
        txtbxNomeDest.Text = !nome_for
        txtbxUFDest.Text = !uf_for
        .Close
    End With
End Sub

Private Sub cmdExport_Click()
    If Me.tag <> "" Then
        'ExportaNfe Me.tag
    Else
        MsgBox "Grave a nota antes!", vbInformation
        
    End If
End Sub

Private Sub cmdProduto_Click()
    Dim p As Produto
    
    'p = Produtos.F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        
    End If
    
    
End Sub

Private Sub cmndbtnBuscaCli_Click()
    buscaCLi
End Sub

Private Sub cmndbtnCfop_Click()
    buscaCFOP
    If Me.tag <> "" Then
        altera
    End If
End Sub

Private Sub cmndbtnConfirmar_Click()
        
End Sub

Private Sub cmndbtnLimpar_Click()
    Dim iCount As Integer
    On Error Resume Next
    If MsgBox("Isto irá limpar todos os campos. Confirma?", vbYesNo + vbQuestion) = vbYes Then
        For iCount = 0 To Me.Controls.count - 1
            Me.Controls(iCount).Text = ""
            
        Next
    End If
    Unload Me
End Sub

Private Sub cmndbtnCupom_Click()
    txtbxCupom_LostFocus
    If bUsaTabelaCupons Then
        buscaCupom
    Else
        buscaCupomAntigo
    End If
End Sub

Private Sub cmndbtnIcms_Click()
    mudaIcms False
End Sub

Private Sub cmndbtnOriginal_Click()
    mudaIcms True
End Sub

Private Sub cmndbtnVenda_Click()
    txtbxVenda_LostFocus
        'Me.Tag = txtbxVenda.Text
        busca
End Sub

Private Sub mnEdiproduto_Click()
    On Error Resume Next
    msflxgrd.Col = 0
    frmCadPro.tag = msflxgrd.Text
    frmCadPro.Show 1
    If Me.tag = "" Then
        busca
    Else
        carrega
    End If
End Sub

Private Sub Label56_Click()

End Sub

Private Sub txtbxCnpjFre_GotFocus()
    P_SelecTexto txtbxCnpjFre
End Sub

Private Sub txtbxCnpjFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCnpjFre_LostFocus()
    P_AjustaTexto txtbxCnpjFre
End Sub

Private Sub txtbxEnderecoFre_GotFocus()
    P_SelecTexto txtbxEnderecoFre
End Sub

Private Sub txtbxEnderecoFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEnderecoFre_LostFocus()
    P_AjustaTexto txtbxEnderecoFre
End Sub

Private Sub txtbxEspFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIEFre_GotFocus()
    P_SelecTexto txtbxIEFre
End Sub

Private Sub txtbxIEFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIEFre_LostFocus()
    P_AjustaTexto txtbxIEFre
End Sub

Private Sub txtbxMarcaFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMuniFre_GotFocus()
    P_SelecTexto txtbxMuniFre
End Sub

Private Sub txtbxMuniFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMuniFre_LostFocus()
    P_AjustaTexto txtbxMuniFre
End Sub

Private Sub txtbxNomeFre_GotFocus()
    P_SelecTexto txtbxNomeFre
End Sub

Private Sub txtbxNomeFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNomeFre_LostFocus()
    P_AjustaTexto txtbxNomeFre
End Sub

Private Sub txtbxNumFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPBrutoFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPlacaFre_GotFocus()
    P_SelecTexto txtbxPlacaFre
End Sub

Private Sub txtbxPlacaFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPlacaFre_LostFocus()
    'P_AjustaPlaca txtbxPlacaFre
End Sub

Private Sub txtbxPLiqFre_Change()
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtdFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfFre_GotFocus()
    P_SelecTexto txtbxUfFre
End Sub

Private Sub txtbxUfFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfFre_LostFocus()
    P_AjustaTexto txtbxUfFre
End Sub

Private Sub txtbxUfVeiFre_GotFocus()
    P_SelecTexto txtbxUfFre
End Sub

Private Sub txtbxUfVeiFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUfVeiFre_LostFocus()
    P_AjustaTexto txtbxUfVeiFre
End Sub

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    sUsu = F_SenhaCodBar
    If txtbxTotalNot.Text = "0,00" Then Exit Sub
    If sUsu <> "" Then
        txtbxDesconto.Enabled = True
        cmndbtnDesconto.tag = sUsu
        txtbxDesconto.SetFocus
    End If
    
End Sub

Private Sub cmndbtnEspelho_Click()
    
        imprimeEspelho frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            imprimeEspelho Printer
            P_Rodape strNome_Emp, True, "(F)"
            Printer.EndDoc
        End If
        Unload frmVisualiza
        
        
        
    
End Sub

Private Sub cmndbtnFrete_Click()
    sstb.Tab = 1
End Sub

Private Sub cmndbtnImprNota_Click()
    Dim sAltura As Single
    Dim sLargura As Single
    
    callprinter
    If Me.tag = "" Then
        MsgBox "Você deve primeiramente gravar a NF!", vbInformation
        Exit Sub
    End If
    
    If chkLpt.Value = 1 Then
        p_imprNFLPT
        
        If MsgBox("A Impressão foi bem sucedida?", vbYesNo + vbQuestion) = vbYes Then
            cmndbtnSalvar_Click
            bActive = False
            msflxgrd.rows = 1
            Form_Activate
            txtbxVenda.Text = ""
            soma
            txtbxVenda.SetFocus
        End If
        
        Exit Sub
    End If
    
   
    
'    p_imprNF frmVisualiza.Pic
'    frmVisualiza.Show 1
'    Unload frmVisualiza
'    Exit Sub
    p_imprNF Printer
    Printer.EndDoc
    If MsgBox("A Impressão foi bem sucedida?", vbYesNo + vbQuestion) = vbYes Then
        cmndbtnSalvar_Click
        bActive = False
        msflxgrd.rows = 1
        Form_Activate
        txtbxVenda.Text = ""
        soma
        txtbxVenda.SetFocus
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Sub grava()
    Dim strSQL As String
    Dim sSeq As String
    Dim iCount As Integer
    If lblCUPOM.caption = "Não" Then
        lblCUPOM.caption = ""
    End If
    If lblCUPOM.caption = "" Then
        lblCUPOM.caption = "0"
    End If
    sSeq = MaxKeySQL("tab_saidasnf", "seq_sai", "", "")
    strSQL = "INSERT INTO tab_saidasNF (SEQ_SAI, COD_CLI, NOTA_SAI, SERIE_SAI, "
    strSQL = strSQL & "NATOP_SAI, CFOP_SAI, VALOR_SAI, BC_SAI, VICMS_SAI, PICMS_SAI, "
    strSQL = strSQL & "VALP_SAI, COD_USU, EMISSAO_SAI, situacao_sai, despesas_Sai,num_ven,nfe_sai,seq_cup) VALUES ("
    strSQL = strSQL & "" & sSeq & ","
    strSQL = strSQL & "'" & txtbxNomeDest.tag & "',"
    strSQL = strSQL & "'" & txtbxNota_Sai.Text & "',"
    strSQL = strSQL & "'" & "1" & "',"
    strSQL = strSQL & "'" & txtbxNatOp.Text & "',"
    strSQL = strSQL & "'" & txtbxCFOP.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxTotalNot.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBC.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxVIcms.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxPicms.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxTotalProd.Text) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & F_GravaData(txtbxEmiss.Text) & ","
    strSQL = strSQL & "'1',"
    strSQL = strSQL & F_GravaDec(txtbxDesp.Text) & ", "
    strSQL = strSQL & "'" & txtbxVenda.Text & "', "
    strSQL = strSQL & "'" & txtNfe.Text & "', "
    strSQL = strSQL & "'" & lblCUPOM.caption & "') "
    Conn.Execute strSQL
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            strSQL = "insert into tab_itsainf(seq_sai,seq_its,cod_pro,qtd_its,unit_its,icms_its)values("
            strSQL = strSQL & "" & sSeq & ","
            strSQL = strSQL & iCount & ","
            .Col = 0
            strSQL = strSQL & "'" & .Text & "', "
            .Col = 4
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .Col = 5
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .Col = 7
            strSQL = strSQL & F_GravaDec(.Text)
            strSQL = strSQL & ")"
            Conn.Execute strSQL
        Next
    End With
    strSQL = F_BuscaUmCampo("cancelado_cup", "tab_cupons", "num_ven='" & Me.tag & "'")
    If strSQL = "N" Or strSQL = "" Then
        baixaSAi
    End If
    
    Me.tag = sSeq
    P_GravaFrete
    bActive = False
    Form_Activate
End Sub
Sub baixaSAi()
    Dim strSQL As String
    Dim iCount As Integer
    Dim sCodPro As String
    Dim qtd As String
    On Error Resume Next
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 4
            strSQL = "update tab_produtos set ef_pro=ef_pro-" & F_GravaDec(.Text)
            .Col = 0
            strSQL = strSQL & " where cod_pro='" & .Text & "' "
            Conn.Execute strSQL

        Next
    End With
End Sub
Sub P_GravaFrete()
    If txtbxNomeFre.Text <> "" Then
        With Me
            strSQL = "INSERT INTO tab_fretes (seq_sai, nome_fre, plveic_fre, ufveic_fre, cnpj_fre, "
            strSQL = strSQL & "endereco_fre, municipio_fre, uf_fre, ie_fre, qdt_fre, esp_fre, "
            strSQL = strSQL & "marca_fre, num_fre, pb_fre, pl_fre, dadosadd_fre) VALUES ("
            strSQL = strSQL & "'" & Me.tag & "',"
            strSQL = strSQL & "'" & .txtbxNomeFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxPlacaFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxUfVeiFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxCnpjFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxEnderecoFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxMuniFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxUfFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxIEFre.Text & "',"
            strSQL = strSQL & F_GravaDec(.txtbxQtdFre.Text) & ","
            strSQL = strSQL & "'" & .txtbxEspFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxMarcaFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxNumFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxPBrutoFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxPLiqFre.Text & "',"
            strSQL = strSQL & "'" & .txtbxDadosAdd.Text & "')"
            Conn.Execute strSQL
        End With
        
    End If
    
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    Dim sNum As String
    If msflxgrd.rows = 1 Then
        MsgBox "Não há itens..."
        Exit Sub
    End If
    If txtbxNota_Sai.Text = "" Then
        MsgBox "Informe o número da Nota!", vbExclamation
        txtbxNota_Sai.SetFocus
        Exit Sub
    End If
    
    
    If seq_sai = "" Then
        strSQL = F_BuscaUmCampo("nota_sai", "tab_saidasnf", "num_ven='" & txtbxVenda.Text & "'")
        If strSQL <> "" Then
            If MsgBox("Já existe uma nota para esta venda: " & strSQL & "Deseja Continuar?", vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
        End If
        If MsgBox("Confirma Gravação da NF?", vbYesNo + vbQuestion) = vbYes Then
            grava
            
            cmdExport_Click
            
            
            sNum = MaxKeySQL("tab_saidasnf", "nota_sai", "", "0000000#")
            txtbxNota_Sai.Text = sNum
        End If
    Else
        altera
    End If
End Sub

Private Sub Form_Activate()
    
    'P_Status "Aguarde..."
    If bActive Then Exit Sub
    sNum = MaxKeySQL("tab_saidasnf", "nota_sai", "", "0000000#")
    txtNfe.Text = MaxKeySQL("tab_saidasnf", "nfe_sai", "", "")
    If seq_sai = "" Then
        busca
        sNum = MaxKeySQL("tab_saidasnf", "nota_sai", "", "0000000#")
    Else
        carrega
        soma
    End If
    
    
    
    
    txtbxNota_Sai.Text = sNum
    bActive = True
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sAux As String
    On Error Resume Next
    
    msflxgrd.rows = 1
    strSQL = "select ts.*,tc.* from tab_clientes as tc left join tab_saidasnf as ts "
    strSQL = strSQL & " on ts.cod_cli=tc.cod_cli "
    strSQL = strSQL & " where seq_sai='" & seq_sai & "'"
    rPesq.Open strSQL, Conn
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Não encontrado!", vbExclamation
        Else
            txtbxBairroDest.Text = Format(!bairro_cli)
            txtbxBC.Text = !bc_sai
            txtbxBCSubs.Text = "0,00" '!bcsubs_sai
            txtbxCepDesp.Text = Format(!cep_cli)
            txtbxCFOP.Text = !cfop_sai
            txtbxCNPJDest.Text = !cpf_cli
            lblCUPOM.caption = !seq_cup
            'txtbxDadosAdd.Text =
            txtbxDesconto.Text = FormatVal(!desconto_sai)
            txtbxDesp.Text = FormatVal(!despesas_sai)
            txtbxEndDest.Text = !logradouro_cli
            txtbxFoneDet.Text = !telefone_cli
            txtbxNatOp.Text = !natop_sai
            txtbxNomeDest.Text = !nome_Cli
            txtbxNomeDest.tag = !cod_cli
            txtbxNota_Sai.Text = !nota_sai
            txtbxPicms.Text = FormatVal(!picms_sai)
            txtbxVIcms.Text = FormatVal(!vicms_sai)
            txtbxTotalProd.Text = FormatVal(!valp_sai)
            txtbxTotalNot.Text = FormatVal(!valor_sai)
            txtbxValorFre.Text = FormatVal(!frete_sai)
            txtbxValorIpi.Text = FormatVal(!ipi_sai)
            txtbxValorSeg.Text = FormatVal(!seguro_sai)
            txtbxVenda.Text = !num_ven
                
            
        End If
        .Close
    End With
    
    strSQL = "select tp.*,ti.*,tm.perc_imp from tab_itsainf as ti,tab_produtos as tp,tab_impostos as tm "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_imp=tm.cod_imp"
    strSQL = strSQL & " and ti.seq_sai='" & seq_sai & "' "
    rPesq.Open strSQL, Conn
    With rPesq
        Do While Not .EOF
            sAux = ""
            sAux = sAux & !cod_pro & Chr(9)
            sAux = sAux & !nome_pro & Chr(9)
            sAux = sAux & !st_pro & Chr(9)
            sAux = sAux & !unidade_pro & Chr(9)
            sAux = sAux & Format(!qtd_its) & Chr(9)
            sAux = sAux & FormatVal(!unit_its, "###,##0.00") & Chr(9)
            sAux = sAux & FormatVal(!unit_its * !qtd_its, "###,##0.000") & Chr(9)
            If !st_pro = "T" Then
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & Format(!perc_imp, "0#")
                Else
                    If Len(F_SoNumero(txtbxCNPJDest.Text)) > 11 Then
                        sAux = sAux & "12"
                    Else
                        sAux = sAux & Format(!perc_imp, "0#")
                    End If
                End If
            Else
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & "0"
                Else
                    sAux = sAux & "0"
                End If
            End If
            sAux = sAux & Chr(9) & !ncm_pro
            If bIgual = False Then
                msflxgrd.AddItem sAux
            End If
            .MoveNext
        Loop
        .Close
    End With

    
End Sub
Sub somaItem()
    Dim cQtd As Currency
    Dim cUnit As Currency
    With msflxgrd
        .Col = 4: cQtd = .Text
        .Col = 5: cUnit = .Text
        .Col = 6: .Text = FormatVal(cQtd * cUnit)
        
    End With
End Sub
Sub soma()
    Dim iCount As Integer
    Dim cTotal As Currency
    Dim cIcms As Currency
    Dim cBC As Currency
    Dim cTotProd As Currency
    Dim iLinhas As Integer
    Dim iColula As Integer
    With msflxgrd
        iLinhas = .row
        iColula = .Col
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 6
            cTotProd = .Text
            cTotal = cTotal + CCur(.Text)
            .Col = 7
            cIcms = cIcms + (cTotProd * (CCur(.Text) / 100))
            If CCur(.Text) > 0 Then
                .Col = 6
                cBC = cBC + CCur(.Text)
            End If
            
            
        Next
        txtbxBC.Text = FormatVal(cBC)
        txtbxTotalNot.Text = FormatVal(cTotal)
        txtbxTotalProd.Text = FormatVal(cTotal)
        txtbxVIcms.Text = FormatVal(cIcms)
        .row = iLinhas
        .Col = iColula
    End With
End Sub
Sub gravaNota()
    On Error GoTo trata_erro
    Dim sNum As String
    sNum = MaxKeySQL("tab_saidas", "nota_sai", "", "0000000#")
    txtbxNota_Sai.Text = sNum
    strSQL = "delete from tab_notasai where seq_sai='" & Me.tag & "' "
    Conn.Execute strSQL
    strSQL = "insert into tab_notasai(seq_sai,nomedest_sai, cnpjdest_sai, enddest_sai,"
    strSQL = strSQL & "bairrodest_sai, cepdest_sai, munidest_sai, fonedest_sai,"
    strSQL = strSQL & "ufdest_sai,iedest_sai, emissao_sai, datahora_sai, BC_SAI,"
    strSQL = strSQL & "VICMS_SAI,PICMS_SAI,bcsubs_SAI,vsubs_SAI,vipi_sai,vfrete_sai,"
    strSQL = strSQL & "vseg_sai,totalprod_sai,totalnota_sai,nomefre_sai,contafre_sai,"
    strSQL = strSQL & "plveicfre_sai,ufveicfre_sai,cnpjfre_sai,endfre_sai,munifre_sai,"
    strSQL = strSQL & "uffre_sai,iefre_sai,qtdfre_sai,espfre_sai,numfre_sai,"
    strSQL = strSQL & "marcafre_sai,pbrutofre_sai, pliqfre_sai,dadosad_sai)values("
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "'" & txtbxNomeDest.Text & "',"
    strSQL = strSQL & "'" & txtbxCNPJDest.Text & "',"
    strSQL = strSQL & "'" & txtbxEndDest.Text & "',"
    strSQL = strSQL & "'" & txtbxBairroDest.Text & "',"
    strSQL = strSQL & "'" & txtbxCepDesp.Text & "',"
    strSQL = strSQL & "'" & txtbxMuniDest.Text & "',"
    strSQL = strSQL & "'" & txtbxFoneDet.Text & "',"
    strSQL = strSQL & "'" & txtbxUFDest.Text & "',"
    strSQL = strSQL & "'" & txtbxIeDest.Text & "',"
    strSQL = strSQL & F_GravaData(txtbxEmiss.Text) & ","
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & F_GravaDec(txtbxBC.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxVIcms.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxPicms.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxBCSubs.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxVlrSubs.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxValorIpi.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxValorFre.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxValorSeg.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxTotalProd.Text) & ", "
    strSQL = strSQL & F_GravaDec(txtbxTotalNot.Text) & ", "
    strSQL = strSQL & "'" & txtbxNomeFre.Text & "', "
    strSQL = strSQL & "'" & txtbxContaFre.Text & "', "
    strSQL = strSQL & "'" & txtbxPlacaFre.Text & "',"
    strSQL = strSQL & "'" & txtbxUfVeiFre.Text & "',"
    strSQL = strSQL & "'" & txtbxCnpjFre.Text & "',"
    strSQL = strSQL & "'" & txtbxEnderecoFre.Text & "',"
    strSQL = strSQL & "'" & txtbxMuniFre.Text & "',"
    strSQL = strSQL & "'" & txtbxUfFre.Text & "',"
    strSQL = strSQL & "'" & txtbxIEFre.Text & "',"
    strSQL = strSQL & "'" & txtbxQtdFre.Text & "',"
    strSQL = strSQL & "'" & txtbxEspFre.Text & "',"
    strSQL = strSQL & "'" & txtbxNumFre.Text & "',"
    strSQL = strSQL & "'" & txtbxMarcaFre.Text & "',"
    strSQL = strSQL & "'" & txtbxPBrutoFre.Text & "',"
    strSQL = strSQL & "'" & txtbxPLiqFre.Text & "',"
    strSQL = strSQL & "'" & txtbxDadosAd.Text & "')"
    Conn.Execute strSQL
    strSQL = "update tab_saidas set nota_sai='" & sNum & "' "
    strSQL = strSQL & " where seq_sai='" & Me.tag & "' "
    Conn.Execute strSQL
Exit Sub
trata_erro:
MsgBox Err.Description, vbCritical



End Sub
Sub altera()
    On Error GoTo trata_erro
    strSQL = "update tab_saidasnf set "
    
        
    
    
    strSQL = strSQL & "emissao_sai=" & F_GravaData(txtbxEmiss.Text) & ","
    'strSQL = strSQL & "datahora_sai=" & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "BC_SAI=" & F_GravaDec(txtbxBC.Text) & ", "
    strSQL = strSQL & "VICMS_SAI=" & F_GravaDec(txtbxVIcms.Text) & ", "
    strSQL = strSQL & "PICMS_SAI=" & F_GravaDec(txtbxPicms.Text) & ", "
   ' strSQL = strSQL & "bcsubs_SAI=" & F_GravaDec(txtbxBCSubs.Text) & ", "
    'strSQL = strSQL & "vsubs_SAI=" & F_GravaDec(txtbxVlrSubs.Text) & ", "
    strSQL = strSQL & "ipi_sai=" & F_GravaDec(txtbxValorIpi.Text) & ", "

    strSQL = strSQL & "frete_sai=" & F_GravaDec(txtbxValorFre.Text) & ", "
    strSQL = strSQL & "seguro_sai=" & F_GravaDec(txtbxValorSeg.Text) & ", "
    strSQL = strSQL & "valp_sai=" & F_GravaDec(txtbxTotalProd.Text) & ", "
    strSQL = strSQL & "valor_sai=" & F_GravaDec(txtbxTotalNot.Text) & ", "
    strSQL = strSQL & "cod_cli='" & txtbxNomeDest.tag & "',"
    strSQL = strSQL & "nota_sai='" & txtbxNota_Sai.Text & "',"
    strSQL = strSQL & "cfop_sai='" & txtbxCFOP.Text & "',"
    strSQL = strSQL & "natop_sai='" & txtbxNatOp.Text & "',"
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "',"
    strSQL = strSQL & "despesas_sai=" & F_GravaDec(txtbxDesp.Text) & ","
    strSQL = strSQL & "num_ven='" & txtbxVenda.Text & "',"
    If txtbxCupom.Text <> "" Then
        strSQL = strSQL & "seq_cup='" & txtbxCupom.Text & "',"
    End If
    If txtNfe.Text <> "" Then
        strSQL = strSQL & "nfe_sai=" & txtNfe.Text & ",     "
    End If
    strSQL = strSQL & "desconto_sai=" & F_GravaDec(txtbxDesconto.Text) & " "
    
    strSQL = strSQL & " where seq_sai='" & Me.tag & "' "
    
    Conn.Execute strSQL
'    MsgBox "Alterado!", vbInformation
Exit Sub
trata_erro:
MsgBox Err.Description, vbCritical
End Sub
Sub busca()
    If bBuscaCupom Then
        buscaSaidaNova
    Else
        buscaSaida
    End If
End Sub
Sub buscaSaidaNova()
    Dim rPesq As New ADODB.Recordset
    Dim rOrd As New ADODB.Recordset
    Dim sAux As String
    Dim sArt As String
    Dim cPreco As Currency
    Dim cQtd As Currency
    Dim bIgual As Boolean
    Dim sNum_ven As String
    msflxgrd.rows = 1
    
    
    
    
    
    
    
    strSQL = "select * from tab_cupons where num_ven='" & txtbxVenda.Text & "'"
    rPesq.Open strSQL, Conn
    
    
    With rPesq
        If .RecordCount > 0 Then
            lblCUPOM.caption = !num_cup
            txtbxCupom.Text = !num_cup
            lblCaixa.caption = !Cod_Cai
            lblTurno.caption = !cod_tur
            
        Else
            lblCUPOM.caption = "Não"
            txtbxCupom.Text = ""
            lblCaixa.caption = "--"
            lblTurno.caption = "--"

        End If
        .Close
    End With '
    
    strSQL = "select tv.* ,tc.* from tab_vendas as tv,tab_clientes as tc where num_Ven='" & txtbxVenda.Text & "' "
    strSQL = strSQL & " and tv.cod_cli=tc.cod_cli "
    rPesq.Open strSQL, Conn
    With rPesq
        sArt = ""
        Do While Not .EOF
            lblData.caption = Format(!data_ven, "dd/mm/yy")
            lblAtendente.caption = !cod_usu
            txtbxNomeDest.Text = !cod_cli
            lblCaixa.caption = !Cod_Cai
            'lblTurno.caption = !cod_tur
            buscaCLi
            'txtbxCNPJDest.Text = !cpf_cli
          '  txtbxIeDest.Text = Format(!rg_cli)
          '  txtbxEndDest.Text = Format(!logradouro_cli)
          '  txtbxBairroDest.Text = Format(!bairro_cli)
          '  txtbxFoneDet.Text = Format(!telefone_cli)
          '  txtbxEndDest.Text = Format(!logradouro_cli)
          '  txtbxBairroDest.Text = Format(!bairro_cli)
          '  txtbxCepDesp.Text = Format(!cep_cli)
          '  txtbxMuniDest.Text = Format(!cidade_cli)
          '  txtbxUFDest.Text = Format(!uf_cli)
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select tp.*,ti.*,tm.perc_imp from tab_itvendas as ti,tab_produtos as tp,tab_impostos as tm "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_imp=tm.cod_imp"
    strSQL = strSQL & " and ti.num_ven='" & txtbxVenda.Text & "' "
    rPesq.Open strSQL, Conn
    With rPesq
        Do While Not .EOF
            sAux = ""
            sAux = sAux & !cod_pro & Chr(9)
            sAux = sAux & !nome_pro & Chr(9)
            sAux = sAux & !st_pro & Chr(9)
            sAux = sAux & !unidade_pro & Chr(9)
            sAux = sAux & Format(!qtd_ite) & Chr(9)
            sAux = sAux & FormatVal(!unit_ite, "###,##0.00") & Chr(9)
            sAux = sAux & FormatVal(!unit_ite * !qtd_ite, "###,##0.000") & Chr(9)
            If !st_pro = "T" Then
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & Format(!perc_imp, "0#")
                Else
                    If Len(F_SoNumero(txtbxCNPJDest.Text)) > 11 Then
                        sAux = sAux & "12"
                    Else
                        sAux = sAux & Format(!perc_imp, "0#")
                    End If
                End If
            Else
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & "0"
                Else
                    sAux = sAux & "0"
                End If
            End If
            sAux = sAux & Chr(9) & !ncm_pro
            
            If bIgual = False Then
                msflxgrd.AddItem sAux
            End If
            .MoveNext
        Loop
        .Close
    End With
    msflxgrd.Visible = False
    soma
    ZebraGrid msflxgrd, 2
    sstb.Tab = 0
    lblItens.caption = msflxgrd.rows - 1
    msflxgrd.Visible = True
    
End Sub
Sub buscaSaida()
    Dim rPesq As New ADODB.Recordset
    Dim rOrd As New ADODB.Recordset
    Dim sAux As String
    Dim sArt As String
    Dim cPreco As Currency
    Dim cQtd As Currency
    Dim bIgual As Boolean
    Dim sNum_ven As String
    msflxgrd.rows = 1
    
    
    
    
    
    
    
    strSQL = "select * from tab_saidas where num_ven='" & txtbxVenda.Text & "' and serie_sai='CUPM' "
    rPesq.Open strSQL, Conn
    
    
    With rPesq
        If .RecordCount > 0 Then
            lblCUPOM.caption = !nota_sai
            txtbxCupom.Text = !nota_sai
            lblCaixa.caption = !Cod_Cai
            'lblTurno.caption = !cod_tur
            
        Else
            lblCUPOM.caption = "Não"
            txtbxCupom.Text = ""
            lblCaixa.caption = "--"
            lblTurno.caption = "--"

        End If
        .Close
    End With '
    
    strSQL = "select tv.* ,tc.* from tab_vendas as tv,tab_clientes as tc where num_Ven='" & txtbxVenda.Text & "' "
    strSQL = strSQL & " and tv.cod_cli=tc.cod_cli "
    rPesq.Open strSQL, Conn
    With rPesq
        sArt = ""
        Do While Not .EOF
            lblData.caption = Format(!data_ven, "dd/mm/yy")
            lblAtendente.caption = !cod_usu
            txtbxNomeDest.Text = !cod_cli
           ' lblCaixa.caption = !Cod_Cai
            'lblTurno.caption = !cod_tur
            buscaCLi
            'txtbxCNPJDest.Text = !cpf_cli
          '  txtbxIeDest.Text = Format(!rg_cli)
          '  txtbxEndDest.Text = Format(!logradouro_cli)
          '  txtbxBairroDest.Text = Format(!bairro_cli)
          '  txtbxFoneDet.Text = Format(!telefone_cli)
          '  txtbxEndDest.Text = Format(!logradouro_cli)
          '  txtbxBairroDest.Text = Format(!bairro_cli)
          '  txtbxCepDesp.Text = Format(!cep_cli)
          '  txtbxMuniDest.Text = Format(!cidade_cli)
          '  txtbxUFDest.Text = Format(!uf_cli)
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select tp.*,ti.seq_ite,unit_ite,qtd_ite,final_ite,tm.perc_imp from tab_itvendas as ti,tab_produtos as tp,tab_impostos as tm "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_imp=tm.cod_imp"
    strSQL = strSQL & " and ti.num_ven='" & txtbxVenda.Text & "' "
    rPesq.Open strSQL, Conn
    With rPesq
        Do While Not .EOF
            sAux = ""
            sAux = sAux & !cod_pro & Chr(9)
            sAux = sAux & !nome_pro & Chr(9)
            sAux = sAux & !st_pro & Chr(9)
            sAux = sAux & !unidade_pro & Chr(9)
            sAux = sAux & Format(!qtd_ite) & Chr(9)
            sAux = sAux & FormatVal(!unit_ite, "###,##0.00") & Chr(9)
            sAux = sAux & FormatVal(!unit_ite * !qtd_ite, "###,##0.000") & Chr(9)
            If !st_pro = "T" Then
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & Format(!perc_imp, "0#")
                Else
                    If Len(F_SoNumero(txtbxCNPJDest.Text)) > 11 Then
                        sAux = sAux & "12"
                    Else
                        sAux = sAux & Format(!perc_imp, "0#")
                    End If
                End If
            Else
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & "0"
                Else
                    sAux = sAux & "0"
                End If
            End If
            sAux = sAux & Chr(9) & !ncm_pro
            
            If bIgual = False Then
                msflxgrd.AddItem sAux
            End If
            .MoveNext
        Loop
        .Close
    End With
    msflxgrd.Visible = False
    soma
    ZebraGrid msflxgrd, 2
    sstb.Tab = 0
    lblItens.caption = msflxgrd.rows - 1
    msflxgrd.Visible = True
    
End Sub

Sub buscaCupom()
    Dim rPesq As New ADODB.Recordset
    Dim rOrd As New ADODB.Recordset
    Dim sAux As String
    Dim sArt As String
    Dim cPreco As Currency
    Dim cQtd As Currency
    Dim bIgual As Boolean
    Dim strSQL As String
    msflxgrd.rows = 1
    
    'strSQL = F_SelectCupom(txtbxCupom.Text)
    
    If strSQL <> "" Then
        txtbxVenda.Text = strSQL
        txtbxVenda_KeyDown 13, 0
    End If
    
    Exit Sub
    
    strSQL = "select tv.* ,tc.* from tab_vendas as tv,tab_clientes as tc,tab_cupons as tcu  "
    strSQL = strSQL & " where tcu.num_ven=tv.num_ven "
    
    strSQL = strSQL & " and tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " and tcu.num_cup='" & txtbxCupom.Text & "'"
    rPesq.Open strSQL, Conn
    
    
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Cupom não encontrado!", vbInformation
            .Close
            Exit Sub
        End If
        If .RecordCount > 1 Then
            
        End If
        sArt = ""
        Do While Not .EOF
            
            txtbxNomeDest.Text = !nomecli_ven
            buscaCLi
            txtbxCNPJDest.Text = !cpf_cli
            txtbxIeDest.Text = Format(!rg_cli)
            txtbxEndDest.Text = Format(!logradouro_cli)
            txtbxBairroDest.Text = Format(!bairro_cli)
            txtbxFoneDet.Text = Format(!telefone_cli)
            txtbxEndDest.Text = Format(!logradouro_cli)
            txtbxBairroDest.Text = Format(!bairro_cli)
            txtbxCepDesp.Text = Format(!cep_cli)
            txtbxMuniDest.Text = Format(!cidade_cli)
            txtbxUFDest.Text = Format(!uf_cli)
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select tp.*,ti.*,tm.perc_imp from tab_itvendas as ti,tab_produtos as tp,tab_impostos as tm "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_imp=tm.cod_imp"
    strSQL = strSQL & " and ti.num_ven='" & Me.tag & "' "
    rPesq.Open strSQL, Conn
    With rPesq
        Do While Not .EOF
            sAux = ""
            sAux = sAux & !cod_pro & Chr(9)
            sAux = sAux & !nome_pro & Chr(9)
            sAux = sAux & !st_pro & Chr(9)
            sAux = sAux & !unidade_pro & Chr(9)
            sAux = sAux & Format(!qtd_ite) & Chr(9)
            sAux = sAux & FormatVal(!final_ite / !qtd_ite, "###,##0.00") & Chr(9)
            sAux = sAux & FormatVal(!final_ite, "###,##0.000") & Chr(9)
            If !st_pro = "T" Then
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & Format(!perc_imp, "0#")
                Else
                    If Len(F_SoNumero(txtbxCNPJDest.Text)) > 11 Then
                        sAux = sAux & "12"
                    Else
                        sAux = sAux & Format(!perc_imp, "0#")
                    End If
                End If
            Else
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & "0"
                Else
                    sAux = sAux & "0"
                End If
            End If
            If bIgual = False Then
                msflxgrd.AddItem sAux
            End If
            .MoveNext
        Loop
        .Close
    End With
    msflxgrd.Visible = False
    soma
    ZebraGrid msflxgrd, 2
    lblItens.caption = msflxgrd.rows - 1
    msflxgrd.Visible = True
    
End Sub

Sub buscaCupomAntigo()
    Dim rPesq As New ADODB.Recordset
    Dim rOrd As New ADODB.Recordset
    Dim sAux As String
    Dim sArt As String
    Dim cPreco As Currency
    Dim cQtd As Currency
    Dim bIgual As Boolean
    Dim strSQL As String
    msflxgrd.rows = 1
    
    'strSQL = F_SelectVendaCupom(txtbxCupom.Text)
    
    If strSQL <> "" Then
        txtbxVenda.Text = strSQL
        txtbxVenda_KeyDown 13, 0
    End If
    
    Exit Sub
    
    strSQL = "select tv.* ,tc.* from tab_vendas as tv,tab_clientes as tc,tab_cupons as tcu  "
    strSQL = strSQL & " where tcu.num_ven=tv.num_ven "
    
    strSQL = strSQL & " and tv.cod_cli=tc.cod_cli "
    strSQL = strSQL & " and tcu.num_cup='" & txtbxCupom.Text & "'"
    rPesq.Open strSQL, Conn
    
    
    
    With rPesq
        If .RecordCount = 0 Then
            MsgBox "Cupom não encontrado!", vbInformation
            .Close
            Exit Sub
        End If
        If .RecordCount > 1 Then
            
        End If
        sArt = ""
        Do While Not .EOF
            
            txtbxNomeDest.Text = !nomecli_ven
            buscaCLi
            txtbxCNPJDest.Text = !cpf_cli
            txtbxIeDest.Text = Format(!rg_cli)
            txtbxEndDest.Text = Format(!logradouro_cli)
            txtbxBairroDest.Text = Format(!bairro_cli)
            txtbxFoneDet.Text = Format(!telefone_cli)
            txtbxEndDest.Text = Format(!logradouro_cli)
            txtbxBairroDest.Text = Format(!bairro_cli)
            txtbxCepDesp.Text = Format(!cep_cli)
            txtbxMuniDest.Text = Format(!cidade_cli)
            txtbxUFDest.Text = Format(!uf_cli)
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select tp.*,ti.*,tm.perc_imp from tab_itvendas as ti,tab_produtos as tp,tab_impostos as tm "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " and tp.cod_imp=tm.cod_imp"
    strSQL = strSQL & " and ti.num_ven='" & Me.tag & "' "
    rPesq.Open strSQL, Conn
    With rPesq
        Do While Not .EOF
            sAux = ""
            sAux = sAux & !cod_pro & Chr(9)
            sAux = sAux & !nome_pro & Chr(9)
            sAux = sAux & !st_pro & Chr(9)
            sAux = sAux & !unidade_pro & Chr(9)
            sAux = sAux & Format(!qtd_ite) & Chr(9)
            sAux = sAux & FormatVal(!final_ite / !qtd_ite, "###,##0.00") & Chr(9)
            sAux = sAux & FormatVal(!final_ite, "###,##0.000") & Chr(9)
            If !st_pro = "T" Then
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & Format(!perc_imp, "0#")
                Else
                    If Len(F_SoNumero(txtbxCNPJDest.Text)) > 11 Then
                        sAux = sAux & "12"
                    Else
                        sAux = sAux & Format(!perc_imp, "0#")
                    End If
                End If
            Else
                If txtbxUFDest.Text = "GO" Then
                    sAux = sAux & "0"
                Else
                    sAux = sAux & "0"
                End If
            End If
            If bIgual = False Then
                msflxgrd.AddItem sAux
            End If
            .MoveNext
        Loop
        .Close
    End With
    msflxgrd.Visible = False
    soma
    ZebraGrid msflxgrd, 2
    lblItens.caption = msflxgrd.rows - 1
    msflxgrd.Visible = True
    
End Sub



Private Sub Form_Load()
    bActive = False
    txtbxEmiss.Text = Date
    With msflxgrd
        .FixedCols = 0
        .FixedRows = 1
        .rows = 1
        .Cols = 9
        .Col = 0: .Text = "Cód"
        .Col = 1: .Text = "Descrição"
        .Col = 2: .Text = "St"
        .Col = 3: .Text = "Un."
        .Col = 4: .Text = "Qtd"
        .Col = 5: .Text = "Unit."
        .Col = 6: .Text = "Total"
        .Col = 7: .Text = "ICMS(%)"
        .Col = 8: .Text = "NCM"
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 3500
        .ColWidth(2) = 600
        .ColWidth(3) = 800
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 700
        .ColAlignment(0) = 0
        .ColAlignment(1) = 0
        .ColAlignment(2) = 0
    End With
    sstb.Tab = 0
End Sub

Private Sub mnNaoTrib_Click()
'
End Sub

Private Sub mnTrib_Click()
    '
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And msflxgrd.rows > 1 Then
        EditaLinha
    End If
End Sub
Sub EditaLinha()
    With msflxgrd
            If .Col = 4 Or .Col = 5 Or .Col = 7 Then
                txtbxGrid.Text = .Text
                txtbxGrid.Left = .CellLeft + .Left
                txtbxGrid.Top = .CellTop + .Top
                txtbxGrid.Width = .CellWidth
                txtbxGrid.height = .CellHeight
                txtbxGrid.Visible = True
                txtbxGrid.SetFocus
            End If
            
        End With
End Sub
Private Sub msflxgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With msflxgrd
        
           Me.PopupMenu mnMenu
    End With
End Sub

Private Sub txtbxBairroDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCepDesp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCFOP_GotFocus()
    P_SelecTexto txtbxCFOP
End Sub

Private Sub txtbxCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCFOP
    End If
End Sub

Private Sub txtbxCNPJDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub



Private Sub txtbxDataSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCupom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCupom_Click
    End If
End Sub

Private Sub txtbxCupom_LostFocus()
    P_AjustaNota txtbxCupom, 6
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    Dim cDesc As Currency
    On Error GoTo trata_erro
    'P_valorlostfocus txtbxDesconto
    cDesc = CCur(F_BuscaUmCampo("desconto_usu", "tab_defusu", _
        "cod_usu='" & cmndbtnDesconto.tag & "'"))
    
    If CCur(txtbxDesconto.Text) > CCur(txtbxTotalNot.Text) * cDesc / 100 Then
        MsgBox "Desconto muito alto!", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    
    End If
    lblValor.caption = FormatVal(lblValor.tag) - CCur(txtbxDesconto.Text)
    txtbxDesconto.Enabled = False
    Exit Sub
trata_erro:
    'MostraErro
    txtbxDesconto.Text = "0"
End Sub

Private Sub txtbxEmiss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEndDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub





Private Sub txtbxFoneDet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxHoraSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxGrid_GotFocus()
    P_SelecTexto txtbxGrid
End Sub

Private Sub txtbxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If msflxgrd.Col = 4 Then
            P_ValorLostFocus txtbxGrid, ""
        Else
            P_ValorLostFocus txtbxGrid
        End If
        msflxgrd.Text = txtbxGrid.Text
        txtbxGrid_LostFocus
        somaItem
        
        soma
        
    End If
End Sub

Private Sub txtbxGrid_LostFocus()
    txtbxGrid.Visible = False
    msflxgrd.SetFocus
End Sub

Private Sub txtbxIeDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub



Private Sub txtbxIESubTrib_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMarcaFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMuniDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub



Private Sub txtbxNatOp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCFOP
    End If
End Sub
Sub buscaCFOP()
    'Dim c As CFOP
    'c = F_BuscaCFOP(txtbxCFOP.Text)
    If c.cfop_cfo <> "" Then
        txtbxCFOP.Text = c.cfop_cfo
        txtbxNatOp.Text = c.nome_cfo
        txtbxEmiss.SetFocus
    End If
End Sub
Private Sub txtbxNomeDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCLi
    End If
End Sub
Sub buscaCLi()
    
End Sub


Private Sub txtbxNumFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxPBrutoFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub



Private Sub txtbxPLiqFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub



Private Sub txtbxPicms_GotFocus()
    P_SelecTexto txtbxPicms
End Sub

Private Sub txtbxPicms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mudaIcms
    End If
End Sub
Sub mudaIcms(Optional original As Boolean = False)
    Dim strSQL As String
    Dim iCount As Integer
    Dim sCodPro As String
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .Col = 0
            sCodPro = .Text
            .Col = 2
            If .Text = "T" Then
                .Col = 7
                If original = False Then
                    .Text = txtbxPicms
                Else
                    '.Text = F_IcmsProd(sCodPro, AliquotaNormal)
                End If
            End If
        Next
        soma
    End With
    
End Sub

Private Sub txtbxPicms_LostFocus()
    P_ValorLostFocus txtbxPicms, "###,##0"
End Sub

Private Sub txtbxUFDest_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub






Private Sub txtbxUFDest_LostFocus()
    buscaUF
End Sub
Sub buscaUF()
    Dim sICMS As String
    Dim sSt As String
    If txtbxUFDest.Text = "GO" Then
        With msflxgrd
            For iCount = 1 To .rows - 1
                .row = iCount
                .Col = 2
                If .Text = "T" Then
                    .Col = 7
                    .Text = "15"
                End If
            Next
        End With
    Else
        If Trim(lblInscrProdRural.caption) <> "" Then
            'é produtor rural
            txtbxCFOP.Text = "6102"
            'txtbxNatOp.Text = F_BuscaCFOPCodigo("6102").nome_cfo
            sICMS = "12"
        ElseIf lblTipoPessoa.caption = "J" Then
            If Trim(txtbxIeDest.Text) <> "" Then
                txtbxCFOP.Text = "6102"
                'txtbxNatOp.Text = F_BuscaCFOPCodigo("6102").nome_cfo
                sICMS = "12"
            Else
                txtbxCFOP.Text = "6929"
                'txtbxNatOp.Text = F_BuscaCFOPCodigo("6929").nome_cfo
                sICMS = "17"
            End If
        Else
            txtbxCFOP.Text = "6929"
            'txtbxNatOp.Text = F_BuscaCFOPCodigo("6929").nome_cfo
            sICMS = "17"
        End If
        
        With msflxgrd
            If sICMS = "" Then sICMS = "17"
            For iCount = 1 To .rows - 1
                .row = iCount
                    .Col = 2
                    sSt = .Text
                    .Col = 7
                    If sSt = "T" Then
                    
                        .Text = sICMS
                    Else
                        .Text = "0"
                    End If
            Next
        End With
    End If
End Sub
Private Sub txtbxVenda_GotFocus()
    P_SelecTexto txtbxVenda
End Sub

Private Sub txtbxVenda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnVenda_Click
        
    End If
End Sub

Private Sub txtbxVenda_LostFocus()
    txtbxVenda.Text = Format(txtbxVenda.Text, "000000#")
End Sub

Private Sub txtNfe_GotFocus()
    P_SelecTexto txtNfe
End Sub

Private Sub txtNfe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub txtQtd_GotFocus()
    P_SelecTexto txtQtd
End Sub

Private Sub txtQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtQtd_LostFocus()
    P_ValorLostFocus txtQtd
End Sub

Private Sub txtUnit_GotFocus()
    P_SelecTexto txtUnit
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtUnit_LostFocus()
    P_ValorLostFocus txtUnit
End Sub
