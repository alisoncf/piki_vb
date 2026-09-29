VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEntradas 
   Caption         =   "Entrada de Produtos"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmEntradas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame frControle 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin MSDBGrid.DBGrid dbgrdProd 
         Bindings        =   "frmEntradas.frx":08CA
         Height          =   4215
         Left            =   120
         OleObjectBlob   =   "frmEntradas.frx":08DF
         TabIndex        =   33
         Top             =   1560
         Width           =   11655
      End
      Begin VB.PictureBox pcOperacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1DDB0&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   120
         ScaleHeight     =   105
         ScaleWidth      =   11625
         TabIndex        =   105
         Top             =   1440
         Width           =   11655
      End
      Begin VB.PictureBox pcResu 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1DDB0&
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   6600
         ScaleHeight     =   1110
         ScaleWidth      =   5145
         TabIndex        =   82
         Top             =   6720
         Width           =   5175
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5040
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resumo"
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
            TabIndex        =   97
            Top             =   0
            Width           =   690
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1800
            TabIndex        =   96
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lblRest 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   1800
            TabIndex        =   95
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3360
            TabIndex        =   94
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblGastos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   3360
            TabIndex        =   93
            Top             =   480
            Width           =   315
         End
         Begin VB.Label lblIpi_Lan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2640
            TabIndex        =   92
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IPI Lanç:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2640
            TabIndex        =   91
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblValoLan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lançado:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblItens 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   4680
            TabIndex        =   88
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Itens:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4680
            TabIndex        =   87
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Prod:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   840
            TabIndex        =   86
            Top             =   240
            Width           =   810
         End
         Begin VB.Label lblTotalProd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   840
            TabIndex        =   85
            Top             =   480
            Width           =   315
         End
         Begin VB.Label lblPercIms 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   4080
            TabIndex        =   84
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICMS:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4080
            TabIndex        =   83
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.PictureBox pcInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   120
         ScaleHeight     =   1110
         ScaleWidth      =   6465
         TabIndex        =   69
         Top             =   6720
         Width           =   6495
         Begin VB.TextBox txtbxSujM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   112
            Text            =   "0,0000"
            Top             =   800
            Width           =   855
         End
         Begin VB.TextBox txtbxCM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   110
            Text            =   "0,0000"
            Top             =   800
            Width           =   855
         End
         Begin VB.TextBox txtbxEr 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   108
            Text            =   "0,0000"
            Top             =   800
            Width           =   855
         End
         Begin VB.TextBox txtbxLucroFinal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   106
            Text            =   "0,0000"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   104
            Text            =   "%"
            Top             =   800
            Width           =   375
         End
         Begin VB.TextBox txtbxTaxas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   102
            Text            =   "0,00"
            Top             =   800
            Width           =   855
         End
         Begin VB.TextBox txtbxICMSSaida 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   100
            Text            =   "0,00"
            Top             =   800
            Width           =   495
         End
         Begin VB.TextBox txtbxGastos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   98
            Text            =   "0,0000"
            Top             =   800
            Width           =   855
         End
         Begin VB.TextBox txtbxCusto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   75
            Text            =   "0,00"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxPreco 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   74
            Text            =   "0,0000"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxNovoCusto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   73
            Text            =   "0,00"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxLucro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   72
            Text            =   "0,0000"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtbxLucroLiq 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   71
            Text            =   "0,0000"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtbxCustoSaida 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   70
            Text            =   "0,0000"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suj. Méd."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   4920
            TabIndex        =   113
            Top             =   555
            Width           =   690
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Média"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3960
            TabIndex        =   111
            Top             =   555
            Width           =   450
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estoque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   3000
            TabIndex        =   109
            Top             =   555
            Width           =   600
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L.Liq.:(%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5640
            TabIndex        =   107
            Top             =   0
            Width           =   720
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Taxas(%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1080
            TabIndex        =   103
            Top             =   555
            Width           =   735
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icms Saída"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   101
            Top             =   560
            Width           =   825
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2040
            TabIndex        =   99
            Top             =   555
            Width           =   540
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo Ant.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   81
            Top             =   0
            Width           =   810
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pç Venda:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3000
            TabIndex        =   80
            Top             =   0
            Width           =   750
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo Entr.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1080
            TabIndex        =   79
            Top             =   0
            Width           =   870
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lucro(%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3960
            TabIndex        =   78
            Top             =   0
            Width           =   690
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Atual(%)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4800
            TabIndex        =   77
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custo Saída"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2040
            TabIndex        =   76
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   10440
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnAbreVenda 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   10440
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C6AE86&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   11625
         TabIndex        =   62
         Top             =   7920
         Width           =   11655
         Begin VB.Label Label38 
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
            Height          =   210
            Left            =   10680
            TabIndex        =   68
            Top             =   75
            Width           =   735
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F7 - Conclui N.F"
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
            Left            =   8760
            TabIndex        =   67
            Top             =   75
            Width           =   1230
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F5 - Novo Produto"
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
            Left            =   6600
            TabIndex        =   66
            Top             =   75
            Width           =   1440
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F4 - Edição de Produtos"
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
            Left            =   3960
            TabIndex        =   65
            Top             =   75
            Width           =   1920
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F3 - Excluir Item"
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
            Left            =   1920
            TabIndex        =   64
            Top             =   75
            Width           =   1290
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F2 - Editar Item"
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
            Left            =   240
            TabIndex        =   63
            Top             =   80
            Width           =   1215
         End
      End
      Begin VB.Frame frOrd 
         BorderStyle     =   0  'None
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   5760
         Width           =   4215
         Begin VB.OptionButton optnDesc 
            Caption         =   "Descrição"
            Height          =   195
            Left            =   1920
            TabIndex        =   59
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optnCod 
            Caption         =   "Có&digo"
            Height          =   195
            Left            =   960
            TabIndex        =   58
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optnSeq 
            Caption         =   "Seqüência"
            Height          =   195
            Left            =   3120
            TabIndex        =   57
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordem:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   510
         End
      End
      Begin VB.Data dtProd 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame frCab 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   11655
         Begin VB.TextBox txtbxValorProd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   9120
            MaxLength       =   12
            TabIndex        =   30
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8160
            MaxLength       =   12
            TabIndex        =   28
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxNum_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            MaxLength       =   12
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxSerie_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5160
            MaxLength       =   4
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxNOme_Ter 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   360
            Width           =   3615
         End
         Begin VB.ComboBox cmbxTipo_Op 
            Height          =   315
            ItemData        =   "frmEntradas.frx":28AC
            Left            =   120
            List            =   "frmEntradas.frx":28AE
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox txtbxDepesas_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            MaxLength       =   12
            TabIndex        =   24
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxIPI_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8160
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "0,00"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxICMS_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7200
            MaxLength       =   12
            TabIndex        =   12
            Text            =   "0,00"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxBC_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            MaxLength       =   12
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxValor_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   9120
            MaxLength       =   12
            TabIndex        =   16
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxEmissao_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            MaxLength       =   10
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxEntrada_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5160
            MaxLength       =   10
            TabIndex        =   22
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxFrete 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7200
            MaxLength       =   12
            TabIndex        =   26
            Text            =   "0,00"
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmndbtnLocaliza 
            Height          =   285
            Left            =   3720
            Picture         =   "frmEntradas.frx":28B0
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Valor Prod:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   9120
            TabIndex        =   29
            Tag             =   "!"
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Desco&nto:"
            Height          =   195
            Left            =   8160
            TabIndex        =   27
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº: &Doc.:"
            Height          =   195
            Left            =   4080
            TabIndex        =   5
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Série:"
            Height          =   195
            Left            =   5160
            TabIndex        =   7
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label12 
            Caption         =   "&Fornecedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Operação:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Despesas:"
            Height          =   195
            Left            =   6240
            TabIndex        =   23
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label8 
            Caption         =   "I&PI:"
            Height          =   255
            Left            =   8160
            TabIndex        =   13
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label7 
            Caption         =   "&ICMS:"
            Height          =   255
            Left            =   7200
            TabIndex        =   11
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "B.Cálc.:"
            Height          =   195
            Left            =   6240
            TabIndex        =   9
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "&Valor Contábil:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   9120
            TabIndex        =   15
            Tag             =   "!"
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Emissão: "
            Height          =   195
            Left            =   4080
            TabIndex        =   19
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "E&ntrada:"
            Height          =   195
            Left            =   5160
            TabIndex        =   21
            Top             =   720
            Width           =   600
         End
         Begin VB.Label lblUF_Ter 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   3585
            TabIndex        =   55
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Frete:"
            Height          =   195
            Left            =   7200
            TabIndex        =   25
            Top             =   720
            Width           =   405
         End
      End
      Begin VB.Frame frPr 
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   5880
         Width           =   11655
         Begin VB.ComboBox cmbxEmb 
            Height          =   315
            Left            =   3480
            TabIndex        =   38
            Top             =   345
            Width           =   735
         End
         Begin VB.TextBox txtbxValorICMS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   8160
            TabIndex        =   49
            Text            =   "0,00"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtbxIcms 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7680
            TabIndex        =   48
            Text            =   "0,0"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtbxCodBarra 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   36
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton cmndbtnConfirma 
            Caption         =   "&Inserir"
            Height          =   375
            Left            =   10800
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtbxIpi 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6960
            TabIndex        =   46
            Text            =   "0,00"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtbxSujerido 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   8880
            TabIndex        =   51
            Text            =   "0,00"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxPrecoFuturo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9840
            TabIndex        =   53
            Text            =   "0,00"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6000
            TabIndex        =   44
            Text            =   "0,00"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtbxQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   40
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtbxUnit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   5040
            TabIndex        =   42
            Text            =   "0,0000"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Emb.:"
            Height          =   195
            Left            =   3480
            TabIndex        =   37
            Top             =   120
            Width           =   405
         End
         Begin VB.Label lblIcmsCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I&CMS(%):(R$)"
            Height          =   195
            Left            =   7800
            TabIndex        =   47
            Top             =   120
            Width           =   945
         End
         Begin VB.Label lblIPI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Ipi(R$):"
            Height          =   195
            Left            =   6960
            TabIndex        =   45
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lblExpr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Tag             =   "B"
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pç. Suj.:"
            Height          =   195
            Left            =   8880
            TabIndex        =   50
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pç. &Futuro:"
            Height          =   195
            Left            =   9840
            TabIndex        =   52
            Top             =   120
            Width           =   780
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unitário"
            Height          =   195
            Left            =   5040
            TabIndex        =   41
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            Height          =   195
            Left            =   4320
            TabIndex        =   39
            Top             =   120
            Width           =   300
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   195
            Left            =   6000
            TabIndex        =   43
            Top             =   120
            Width           =   360
         End
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Para Abrir a Nota Fiscal, Informe os Dados de Cabeçalho e Clique em Confirmar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1793
         TabIndex        =   61
         Top             =   3600
         Width           =   8295
      End
   End
End
Attribute VB_Name = "frmEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean 'para saber se o frm já foi aberto
Dim iTam As Integer 'tam do form
Dim sTableCab As String
Dim sTableItens As String
Public operacao As String

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

Private Sub cmbxBusca_Click()
    lblExpr.caption = "&" & cmbxBusca.Text
    
    If cmbxBusca.ListIndex = 0 Then
        lblExpr.Tag = "B"
    Else
        lblExpr.Tag = Left(cmbxBusca.Text, 1)
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxEmb_LostFocus()
    P_ValorLostFocus cmbxEmb, "###,##0"
    confEmb
    txtbxQtd.SetFocus
End Sub
Sub confEmb()
    Dim strSQL As String
    Dim iCount As Integer
    Dim ok As Boolean
    On Error GoTo trata_Erro
    ok = False
    If txtbxCodBarra.Tag = "" Then
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
            cmbxEmb.Tag = cmbxEmb.Text
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & txtbxCodBarra.Tag & "',"
            strSQL = strSQL & cmbxEmb.Text & ")"
            banco.Execute strSQL
            P_PreencheEmbalagens txtbxCodBarra.Tag, cmbxEmb, True
            cmbxEmb.Text = cmbxEmb.Tag
            
        End If
        
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Private Sub cmbxTipo_Op_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Sub P_Grava()
    Dim sTipo As String
    Dim strSQL As String
    On Error GoTo trata_Erro
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
                & " and cod_ter ='" & txtbxNome_ter.Tag & "' " _
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
    strSQL = strSQL & "'" & txtbxNome_ter.Tag & "',"
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
    banco.Execute strSQL
    MsgBox "Nota incluída com sucesso!", vbInformation
    P_Clear
    txtbxNome_ter.SetFocus
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Clear()
    txtbxNome_ter.Text = ""
    txtbxNome_ter.Tag = ""
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

Private Sub cmndbtnAbreVenda_Click()
    If Me.Tag = "" Then
        abreNota
        HabCab False
    Else
        AlteraNota
        HabCab False
        HabItens True
    End If
End Sub
Sub abreNota()
    Dim strSQL As String
    Dim sSeq As String
    If txtbxNome_ter.Tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtbxNome_ter.SetFocus
        Exit Sub
    End If
    If F_ValorContabilNota <> CCur(txtbxValor_Not.Text) Then
        MsgBox "OS valores não batem!", vbCritical
        Exit Sub
    End If
    If txtbxValor_Not.Text = "0,00" Then
        On Error Resume Next
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        txtbxValor_Not.SetFocus
        Exit Sub
    End If
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtbxNome_ter.Tag & "' " & _
        " and num_not='" & txtbxNum_Not.Text & "' and serie_not ='" & _
        txtbxSerie_Not.Text & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Abertura da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    'frBotoes.Visible = True
    sSeq = MaxKeySQL(sTableCab, "seq_not", "", "000000#")
    strSQL = "insert into " & sTableCab & "(seq_not,Cod_For,Num_not,Serie_not,desconto_not,valorprod_not,Valor_not,Bc_Not,"
    strSQL = strSQL & "frete_not,icms_not,ipi_not,despesas_not,operacao_not,emissao_not,"
    strSQL = strSQL & "entrada_not,cod_usu,status_not,data_not)Values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & txtbxNome_ter.Tag & "', "
    strSQL = strSQL & "'" & txtbxNum_Not.Text & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxDesconto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValorProd.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxValor_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxBC_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxFrete.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxICMS_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxIPI_Not.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxDepesas_Not.Text) & ","
    strSQL = strSQL & "'" & Left(cmbxTipo_Op.Text, 4) & "',"
    strSQL = strSQL & F_GravaData(txtbxEmissao_Not.Text) & ","
    strSQL = strSQL & F_GravaData(txtbxEntrada_Not.Text) & ","
    strSQL = strSQL & "'" & strCod_Usu & "','1', " 'em aberto
    strSQL = strSQL & F_GravaDataTempo(Now) & ") "
    banco.Execute strSQL
    Me.Tag = sSeq
    lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.Text) + _
                                         CCur(txtbxFrete.Text))
    bActive = False
    Form_Activate
    'cmndbtnAddProd.SetFocus
    HabItens True
    txtbxCodBarra.SetFocus
End Sub
Sub AlteraNota()
    Dim strSQL As String
    Dim sSeq As String
    If txtbxNome_ter.Tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtbxNome_ter.SetFocus
        Exit Sub
    End If
    If F_ValorContabilNota <> CCur(txtbxValor_Not.Text) Then
        MsgBox "OS valores não batem!", vbCritical
        Exit Sub
    End If
    If txtbxValor_Not.Text = "0,00" Then
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        txtbxValor_Not.SetFocus
        Exit Sub
    End If
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtbxNome_ter.Tag & "' " & _
        " and num_not='" & txtbxNum_Not.Text & "' and serie_not ='" & _
        txtbxSerie_Not.Text & "' and seq_not <>'" & Me.Tag & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Alteração da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    'frBotoes.Visible = True
    
    
    strSQL = "update " & sTableCab & " set "
    strSQL = strSQL & "cod_for='" & txtbxNome_ter.Tag & "', "
    strSQL = strSQL & "num_not='" & txtbxNum_Not.Text & "',"
    strSQL = strSQL & "serie_not='" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & "desconto_not=" & F_GravaDec(txtbxDesconto.Text) & ","
    strSQL = strSQL & "valorprod_not=" & F_GravaDec(txtbxValorProd.Text) & ","
    strSQL = strSQL & "valor_not=" & F_GravaDec(txtbxValor_Not.Text) & ","
    strSQL = strSQL & "bc_not=" & F_GravaDec(txtbxBC_Not.Text) & ","
    strSQL = strSQL & "frete_not=" & F_GravaDec(txtbxFrete.Text) & ","
    strSQL = strSQL & "icms_not=" & F_GravaDec(txtbxICMS_Not.Text) & ","
    strSQL = strSQL & "ipi_not=" & F_GravaDec(txtbxIPI_Not.Text) & ","
    strSQL = strSQL & "despesas_not=" & F_GravaDec(txtbxDepesas_Not.Text) & ","
    strSQL = strSQL & "operacao_not='" & Left(cmbxTipo_Op.Text, 4) & "',"
    strSQL = strSQL & "emissao_not=" & F_GravaData(txtbxEmissao_Not.Text) & ","
    strSQL = strSQL & "entrada_not=" & F_GravaData(txtbxEntrada_Not.Text) & ","
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "', "
    strSQL = strSQL & "data_not=" & F_GravaDataTempo(Now)
    
    strSQL = strSQL & " where seq_not='" & Me.Tag & "' "
    banco.Execute strSQL
    
    lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.Text) + _
                                         CCur(txtbxFrete.Text))
    bActive = False
    Form_Activate
    RecalculaItens
    
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
            txtbxNome_ter.Tag & "' and cod_pro='" & p.cod_pro & "'") <> "" Then
        MsgBox "Este produto já se encontra lançado!", vbExclamation
        Exit Sub
    End If
    
    txtbxCodBarra.Text = p.nome_pro
    txtbxCodBarra.Tag = p.cod_pro
    P_PreencheEmbalagens p.cod_pro, cmbxEmb, True
    cmbxEmb.ListIndex = 0
    'lblCodBarra.caption = p.codbarra_pro
    txtbxCusto.Text = FormatVal(p.c2_pro)
    txtbxPreco.Text = FormatVal(p.p2_pro)
    txtbxLucro.Text = FormatVal(p.lucro_pro)
    txtbxER.Text = p.e2_pro
    cmbxEmb.SetFocus
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
    Dim sSeq As String
    Dim cCusto As Currency
    Dim cGastos As Currency
    sSeq = F_BuscaUmCampo("Max(Seq_ite)", sTableItens, "Num_Not='" & txtbxNum_Not.Text & "'" & _
        " and cod_for='" & txtbxNome_ter.Tag & "'")
    If sSeq = "" Then
        sSeq = "1"
    Else
        sSeq = Int(sSeq) + 1
    End If
    
    cGastos = CCur(lblGastos.caption) * cFator
    cCusto = unit + (cGastos / qtd) + (ipi / qtd)
    
    strSQL = "insert into " & sTableItens & "(num_not,cod_for,serie_not,cod_pro,"
    strSQL = strSQL & "unit_ite,qtd_ite,ipi_ite,custo_ite,gastos_ite,seq_ite)values("
    strSQL = strSQL & "'" & txtbxNum_Not.Text & "',"
    strSQL = strSQL & "'" & txtbxNome_ter.Tag & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit) & ","
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(ipi) & ","
    
    strSQL = strSQL & F_GravaDec(cCusto) & ","
    strSQL = strSQL & F_GravaDec(cGastos) & ","
    strSQL = strSQL & sSeq & ") "
    banco.Execute strSQL
    
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
    strSQL = strSQL & "and cod_for='" & txtbxNome_ter.Tag & "' "
    strSQL = strSQL & "and serie_not='" & txtbxSerie_Not.Text & "' "
    strSQL = strSQL & "and cod_pro= '" & cod_pro & "' "
    
    
    banco.Execute strSQL
    P_buscaitens
End Sub

Private Sub cmndbtnConfirmar_Click()
    
    If Me.Tag = "" Then
        If txtbxNome_ter.Tag = "" Then
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
    On Error GoTo trata_Erro
    sTipo = UCase(Left(lblTipo.caption, 1))
    If sTipo = "S" Then
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and tipo_not= 'S' and cod_Cli = '" & _
                strCOD_CLI) <> "" And _
                txtbxNum_Not.Text <> txtbxNum_Not.Tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and cod_ter ='" & txtbxNome_ter.Tag & "' " _
                & " and tipo_not= 'E' and cod_Cli = '" & strCOD_CLI & "'") <> "" _
                And txtbxNum_Not.Text <> txtbxNum_Not.Tag Then
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
    strSQL = strSQL & " and num_not ='" & txtbxNum_Not.Tag & "' "
    strSQL = strSQL & " and tipo_not='" & sTipo & "' "
    strSQL = strSQL & " and cod_ter='" & Me.Tag & "' "
    strSQL = strSQL & " and serie_not='" & txtbxSerie_Not.Tag & "'"
    banco.Execute strSQL
    MsgBox "Nota Alterada com sucesso!", vbInformation
    If frmEditaMov.Visible = True Then
        iPos = frmEditaMov.dtNot.Recordset.AbsolutePosition
        frmEditaMov.dtNot.Refresh
        frmEditaMov.dtNot.Recordset.Move iPos
    End If
    Unload Me
    
    Exit Sub
trata_Erro:
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
        If dbgrdProd.Row = -1 Then
            Exit Sub
        End If
        With dtProd.Recordset
            txtbxCusto.Text = FormatVal(!cf_pro)
            txtbxPreco.Text = FormatVal(!pV_pro)
            txtbxLucro.Text = FormatVal(!lucro_pro)
            txtbxCodBarra.Text = !nome_pro
            txtbxCodBarra.Tag = !cod_pro
            txtbxQtd.Text = !qtd_ite
            cmbxEmb.Text = !qtd_emb
            txtbxUnit.Text = !unit_ite
            txtbxTotal.Text = !total
            cmndbtnConfirma.Tag = !seq_ite
            txtbxIpi.Text = !ipi_ite
            txtbxSujerido.Text = !sujerido_ite
            
            txtbxIcms.Text = !icms_ite
            txtbxQtd_LostFocus
            txtbxPrecoFuturo_LostFocus
            txtbxIcms_LostFocus
            txtbxPrecoFuturo.Text = !precofuturo_ite
            
        End With
         
         
         
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Dim strSQL As String
    'On Error GoTo trata_erro
    If MsgBox("Confirma exlusão do item " & dtProd.Recordset!nome_pro & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from " & sTableItens & " where SEQ_NOT='" & Me.Tag & "' "

        strSQL = strSQL & " and cod_pro='" & dtProd.Recordset!cod_pro & "' "
        banco.Execute strSQL
        
        If banco.RecordsAffected > 0 Then
            MsgBox "Item excuído! ", vbInformation
            soma
        Else
            MsgBox "Nenhum Registro Foi Afetado!", vbInformation
        End If
        
        
    End If
    P_buscaitens
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnForn_Click()

End Sub

Private Sub cmndbtnConfirma_Click()
    
    If txtbxCodBarra.Tag = "" Then
        MsgBox "Informe o Nome do Produto e tecle <ENTER>!", vbCritical
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    If F_Qtd = 0 Then
        MsgBox "Informe o tipo de embalagem e quantidade", vbExclamation
        cmbxEmb.SetFocus
        Exit Sub
    End If
    If cmndbtnConfirma.Tag = "" Then
        insere
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
    dbgrdProd.Enabled = opcao
    frPr.Enabled = opcao
End Sub
Sub HabCab(opcao As Boolean)
    frCab.Enabled = opcao
    txtbxNome_ter.Enabled = opcao
    cmbxTipo_Op.Enabled = opcao
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
    frmSelectCli.txtbxNome.Tag = "F"
    frmSelectCli.txtbxNome.Text = txtbxNome_ter.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxNome_ter.Text = frmSelectCli.dtBusca.Recordset!nome
    txtbxNome_ter.Tag = frmSelectCli.dtBusca.Recordset!cod
    If Me.operacao = "F" Then
        txtbxNum_Not.Text = MaxKeySQL("tab_cmp", "num_not", "cod_for='" & txtbxNome_ter.Tag & "' ", "0000000#")
    End If
    Unload frmSelectCli
    txtbxNum_Not.SetFocus
End Sub

Private Sub cmndbtnNovoProd_Click()
    Dim fr As New frmCadPro
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
    If IsLoad(frmEdiEntr.Name) Then
        frmEdiEntr.dtNot.Refresh
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
    ElseIf dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não é possível concluir uma nota sem itens!", vbExclamation
        Exit Sub
    End If
    
    
    If lblIpi_Lan.caption <> txtbxIPI_Not.Text Then
        MsgBox "O valor do IPI não está completamente lançado!", vbExclamation
        Exit Sub
    ElseIf dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não é possível concluir uma nota sem itens!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma efetivação da Nota Fiscal? ", vbYesNo + vbQuestion) = vbYes Then
        With dtProd.Recordset
            .MoveFirst
            strSQL = "update " & sTableCab & " set status_not = '2' "
            strSQL = strSQL & " where Seq_not='" & Me.Tag & "' "
            
            banco.Execute strSQL
            
        End With
        MsgBox "Nota Fiscal concluída!", vbInformation
        If IsLoad(frmEdiEntr.Name) Then
            frmEdiEntr.dtNot.Refresh
        End If
        Unload Me
    End If
    
    
    
End Sub


Private Sub dbgrdProd_DblClick()
'    If cmndbtnEditProd.Enabled And cmndbtnEditProd.Visible = True Then
    
        editaItem
        On Error Resume Next
        txtbxQtd.SetFocus
 '   End If
End Sub

Private Sub dbgrdProd_GotFocus()
    dbgrdProd.MarqueeStyle = 3
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        editaItem
    End If
    If KeyCode = vbKeyDelete Then
        cmndbtnExc_Prod_Click
    End If
    
End Sub

Private Sub dbgrdProd_LostFocus()
    dbgrdProd.MarqueeStyle = 0
    
End Sub

Private Sub dtProd_Reposition()
    With dtProd.Recordset
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
    With dtProd.Recordset
    Err.Clear
    On Error GoTo trata_Erro
        
        If .RecordCount = 0 Then
            lblIpi_Lan.caption = "0,00"
            lblValoLan.caption = "0,00"
            lblItens.caption = "0"
            
            
            lblTotalProd.caption = "0,0"
            
            
            lblGastos.Tag = "0,00"
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
            
            .MoveNext
        Loop
            lblIpi_Lan.caption = FormatVal(cIPI)
            lblValoLan.caption = FormatVal(cValor)
            lblTotalProd.caption = FormatVal(cValor)
            'lblItens.caption = .RecordCount
            lblItens.caption = Format(iCount)
            lblGastos.Tag = cDesp
            lblRest.caption = F_Restante
            .MoveFirst
            .Move iPos
    End With
    If CCur(txtbxBC_Not.Text) > 0 Then
        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.Text) / CCur(txtbxBC_Not.Text) * 100, 2)) & " %"
    End If
trata_Erro:
End Sub

Sub insere()
    Dim strSQL As String
    Dim sSeq As String
    strSQL = F_BuscaUmCampo("seq_ite", sTableItens, "cod_pro='" & txtbxCodBarra.Tag & "'  and seq_not='" & Me.Tag & "' ")
    
    
    If strSQL <> "" Then
        MsgBox "O Item já se encontra lançado como Item: " & strSQL, vbCritical
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    sSeq = MaxKeySQL(sTableItens, "seq_ite", "seq_not='" & Me.Tag & "'", "")
    strSQL = "Insert into " & sTableItens & "(seq_not,seq_ite,cod_pro,unit_ite,qtd_ite,qtd_emb,ipi_ite,"
    strSQL = strSQL & "gastos_ite,custo_ite,sujerido_ite,precofuturo_ite,lucrosuj_ite,"
    strSQL = strSQL & "lucroliq_ite,custosai_ite,icms_ite)values("
    strSQL = strSQL & "'" & Me.Tag & "',"
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & txtbxCodBarra.Tag & "',"
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
    banco.Execute strSQL
    P_buscaitens
    P_LimpaItem
    
End Sub
Sub RecalculaItens()
    
    With dtProd.Recordset
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            editaItem
            altera 1 'não dá refresh
            .MoveNext
        Loop
        P_buscaitens
    End With
End Sub
Sub P_LimpaItem()
    On Error Resume Next
    txtbxCodBarra.Text = ""
    txtbxCodBarra.Tag = ""
    txtbxCodBarra.SetFocus
    cmndbtnConfirma.Tag = ""
End Sub
Sub altera(Optional opcao As Integer = 0)
    Dim strSQL As String
    Dim sSeq As String
    Dim i As Integer
    If opcao = 0 Then
        strSQL = F_BuscaUmCampo("seq_ite", sTableItens, "cod_pro='" & txtbxCodBarra.Tag & "' and seq_ite <> '" & cmndbtnConfirma.Tag & "' and seq_not='" & Me.Tag & "'")
    
    
        If MsgBox("Confirma alteração do Item?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    
        If strSQL <> "" Then
            MsgBox "O Item já se encontra lançado como Item: " & strSQL, vbCritical
            txtbxCodBarra.SetFocus
            Exit Sub
        End If
    End If
    
    strSQL = "update " & sTableItens & " set "
    strSQL = strSQL & "cod_pro='" & txtbxCodBarra.Tag & "',"
    strSQL = strSQL & "unit_ite=" & F_GravaDec(txtbxUnit.Text) & ","
    strSQL = strSQL & "qtd_ite=" & F_GravaDec(txtbxQtd.Text) & ","
    strSQL = strSQL & "qtd_emb=" & F_GravaDec(cmbxEmb.Text) & ","
    strSQL = strSQL & "ipi_ite=" & F_GravaDec(txtbxIpi.Text) & ","
    strSQL = strSQL & "gastos_ite=" & F_GravaDec(txtbxGastos.Text) & ","
    strSQL = strSQL & "custo_ite=" & F_GravaDec(txtbxNovoCusto.Text) & ","
    strSQL = strSQL & "sujerido_ite=" & F_GravaDec(txtbxSujerido.Text) & ","
    strSQL = strSQL & "precofuturo_ite=" & F_GravaDec(txtbxPrecoFuturo.Text) & ","
    strSQL = strSQL & "lucrosuj_ite=" & F_GravaDec(txtbxLucro.Text) & ","
    strSQL = strSQL & "lucroliq_ite=" & F_GravaDec(txtbxLucroLiq.Text) & ","
    strSQL = strSQL & "custosai_ite=" & F_GravaDec(txtbxCustoSaida.Text) & ","
    strSQL = strSQL & "icms_ite=" & F_GravaDec(txtbxIcms.Text) & " "
    strSQL = strSQL & " where seq_not='" & Me.Tag & "' "
    strSQL = strSQL & " and seq_ite='" & cmndbtnConfirma.Tag & "' "
    
    banco.Execute strSQL
    cmndbtnConfirma.Tag = ""
    If opcao = 0 Then
        i = dtProd.Recordset.AbsolutePosition
    
        P_buscaitens
        dtProd.Recordset.Move i
    
        P_LimpaItem
    End If
End Sub
Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    
    If operacao = "F" Then
        sTableCab = "tab_cmp"
        sTableItens = "tab_itcmp"
        pcOperacao.BackColor = &H80000018
    Else
        sTableCab = "tab_compras"
        sTableItens = "tab_itcompras"
        pcOperacao.BackColor = &HB1DDB0
    End If
    
    dtProd.Connect = banco.Connect
    
    
    
    If Me.Tag = "" Then
        iTam = Me.Height
        
        
        dbgrdProd.Visible = False
        frPr.Visible = False
        pcInfo.Visible = False
        frOrd.Visible = False
        pcResu.Visible = False
        dtProd.Connect = banco.Connect
        
    Else
        
        cmndbtnAbreVenda.Visible = False
        'cmndbtnCancelar.Visible = False
        dbgrdProd.Visible = True
        frPr.Visible = True
        pcInfo.Visible = True
        frOrd.Visible = True
        pcResu.Visible = True
      
        dtProd.Connect = banco.Connect
        cmndbtnAbreVenda.Visible = False
        'cmndbtnCancelar.Visible = False
        frCab.Enabled = False
        P_buscaitens
        lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.Text) + _
                             CCur(txtbxFrete.Text))
                             
                             
    End If
    
    
    
    bActive = True
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyI And Shift = 1 Then
        P_InfPro dtProd.Recordset!cod_pro
    
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
            frmEdiPro.Show 0
        
        frmEdiPro.SetFocus
    ElseIf KeyCode = vbKeyF5 Then
        Me.WindowState = vbMinimized
        Dim fr As New frmCadPro
        fr.Show 0
        
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    End If

End Sub

Sub P_buscaitens()
    Dim strSQL As String
    
        
    strSQL = "select distinct ti.*,(ti.unit_ite*ti.qtd_ite*ti.qtd_emb) as total,ti.gastos_ite,tp.*,"
    
    strSQL = strSQL & "tm.nome_mar,tc.nome_cat, tp.cr_pro, tp.pv_pro "
    strSQL = strSQL & " from " & sTableItens & " as ti, tab_produtos as tp, tab_codbarras as tb, "
    strSQL = strSQL & "tab_marcas as tm, tab_categorias as tc "
    strSQL = strSQL & " where ti.seq_not ='" & Me.Tag & "' "
    
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar and tp.cod_cat=tc.cod_cat "
    
    strSQL = strSQL & " and tp.cod_pro=tb.cod_pro "
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro "
    If optnCod.Value = True Then
        strSQL = strSQL & " order by ti.cod_pro"
    
    End If
    If optnDesc.Value = True Then
        strSQL = strSQL & " order by tp.nome_pro"
    
    End If
    If optnSeq.Value = True Then
        strSQL = strSQL & " order by ti.seq_ite"
    
    End If
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    soma
    If CCur(txtbxIpi.Text) = 0 Then
        lblIPI.Enabled = False
        txtbxIpi.Enabled = False
    End If
    If CCur(txtbxICMS_Not.Text) = 0 Then
        lblIcmsCap.Enabled = False
        txtbxIcms.Enabled = False
        txtbxValorICMS.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    P_CentraForm Me
    dtProd.Connect = banco.Connect
    P_PreencheOperacoes cmbxTipo_Op, "E"
    'cmbxBusca.ListIndex = 0
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub Option2_Click()
    
End Sub

Private Sub optnBar_Click()
    lblExpr.caption = "C&od. Barra:"
    lblExpr.Tag = "B"
End Sub

Private Sub optnC_Click()
        lblExpr.caption = "C&od. Prod.:"
        lblExpr.Tag = "C"
End Sub



Private Sub optnN_Click()
    lblExpr.caption = "N&ome Prod.:"
    lblExpr.Tag = "N"
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

Private Sub txtbxBC_Not_GotFocus()
    P_SelecTexto txtbxBC_Not
End Sub

Private Sub txtbxBC_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBC_Not_LostFocus()
    P_ValorLostFocus txtbxBC_Not
    If CCur(txtbxBC_Not.Text) > 0 Then
        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.Text) / CCur(txtbxBC_Not.Text), 2) * 100) & " %"
    End If
End Sub


Private Sub txtbxBCS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub txtbxCodBarra_GotFocus()
    P_SelecTexto txtbxCodBarra
End Sub
Private Sub txtbxCodBarra_LostFocus()
    P_AjustaTexto txtbxCodBarra
    If lblExpr.Tag = "C" Then
        txtbxCodBarra.Text = Format(txtbxCodBarra.Text, "0000000#")
    End If
    
End Sub

Private Sub txtbxCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnAddProd_Click
    End If
End Sub

Private Sub txtbxDepesas_Not_GotFocus()
    P_SelecTexto txtbxDepesas_Not
End Sub

Private Sub txtbxDepesas_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDepesas_Not_LostFocus()
    P_ValorLostFocus txtbxDepesas_Not
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
    P_ValorLostFocus txtbxDesconto
End Sub

Private Sub txtbxEmissao_Not_GotFocus()
    P_SelecTexto txtbxEmissao_Not
End Sub

Private Sub txtbxEmissao_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEmissao_Not_LostFocus()
    P_AjustaData txtbxEmissao_Not
End Sub




Private Sub txtbxEntrada_Not_GotFocus()
    P_SelecTexto txtbxEntrada_Not
End Sub

Private Sub txtbxEntrada_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEntrada_Not_LostFocus()
    On Error GoTo trata_Erro
    P_AjustaData txtbxEntrada_Not
    If txtbxEntrada_Not.Text = "" Then Exit Sub
    If CDate(txtbxEntrada_Not.Text) < CDate(txtbxEmissao_Not.Text) Then
        MsgBox "A data de Entrada não pode ser anterior a data de Emissão", vbExclamation
        txtbxEntrada_Not.SetFocus
    End If
trata_Erro:
    
End Sub

Private Sub txtbxFrete_GotFocus()
    P_SelecTexto txtbxFrete
End Sub

Private Sub txtbxFrete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFrete_LostFocus()
    P_ValorLostFocus txtbxFrete
End Sub

Private Sub txtbxIcms_GotFocus()
    P_SelecTexto txtbxIcms
End Sub

Private Sub txtbxIcms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIcms_LostFocus()
    P_ValorLostFocus txtbxIcms
    txtbxICMSSaida.Text = CCur(F_IcmsProd(txtbxCodBarra.Tag, 3)) / 100 'percentual
    P_CalcItem
End Sub

Private Sub txtbxICMS_Not_GotFocus()
    P_SelecTexto txtbxICMS_Not
End Sub

Private Sub txtbxICMS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxICMS_Not_LostFocus()
    P_ValorLostFocus txtbxICMS_Not
    If CCur(txtbxBC_Not.Text) > 0 Then
        lblPercIms.caption = FormatVal(Round(CCur(txtbxICMS_Not.Text) / CCur(txtbxBC_Not.Text) * 100, 2)) & " %"
    End If
End Sub

Sub P_CalcItem()
    Dim strSQL As String
    Dim cCusto As Currency
    Dim dInd As Double
    Dim cSuj As Currency
    Dim cLucro As Currency
    On Error GoTo trata_Erro
    txtbxTaxas.Text = F_TaxasEntradas
    txtbxValorICMS.Text = Round(CCur(txtbxIcms.Text) * (txtbxTotal.Text) / 100, 4)
    If F_Qtd > 0 Then
        dInd = Round(CDbl(txtbxTotal.Text) / CDbl(txtbxValor_Not.Text) / F_Qtd, 8)
    End If
    
    cLucro = txtbxLucro.Text
    cCusto = CCur(txtbxUnit.Text) + (dInd * F_Despesas) + (CCur(txtbxIpi.Text) / (CCur(txtbxQtd.Text) * CCur(cmbxEmb.Text)))
    txtbxNovoCusto.Text = FormatVal(cCusto)
    cSuj = cCusto + (cCusto * cLucro / 100)
    txtbxSujerido.Text = FormatVal(cSuj)
    txtbxPrecoFuturo.Text = FormatVal(cSuj)
    P_calcLucroLiq
    txtbxGastos.Text = FormatVal(dInd * F_Despesas, "###,###,##0.0000")
    
    
    'custo médio das ultimas duas
    cSuj = CCur(txtbxCM.Text) + (CCur(txtbxCM.Text) * cLucro / 100)
    txtbxSujM.Text = FormatVal(cSuj)
    
    Exit Sub
trata_Erro:
End Sub
Sub P_calcLucroLiq()
    Dim cLucro As Currency
    Dim cCusto As Currency
    Dim cIcms As Currency
    If CCur(txtbxSujerido.Text) > 0 Then
    On Error GoTo trata_Erro
        'cLucro = CCur(txtbxLucro.Text) * CCur(txtbxPrecoFuturo.Text) / CCur(txtbxSujerido.Text)
        cLucro = (CCur(txtbxPrecoFuturo.Text) / CCur(txtbxNovoCusto.Text)) - 1
        cLucro = cLucro * 100
    End If
    txtbxLucroLiq.Text = cLucro
    'cLucro = cLucro - (cLucro * CCur(txtbxTaxas.Text) / 100)
    cLucro = cLucro - CCur(txtbxTaxas.Text)
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
trata_Erro:
    MsgBox Err.Description, vbCritical
    
End Sub
Sub P_Unit()
    If F_Qtd <> 0 Then
        txtbxUnit.Text = FormatVal(CCur(txtbxTotal.Text) / F_Qtd, "###,###,##0.0000")
    End If
End Sub
Private Sub txtbxIpi_GotFocus()
    If CCur(txtbxIpi.Text) = 0 Then
        txtbxIpi.Text = F_RestIpi
    End If
    P_SelecTexto txtbxIpi
End Sub

Function F_Restante() As Currency
    F_Restante = F_ValorProdutos - CCur(lblValoLan.caption)

End Function
Function F_RestIpi() As Currency
    F_RestIpi = CCur(txtbxIPI_Not.Text) - CCur(lblIpi_Lan.caption)
End Function
Private Sub txtbxIpi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIpi_LostFocus()
    
    P_ValorLostFocus txtbxIpi
    If CCur(txtbxIpi.Text) > F_RestIpi Then
        MsgBox "O Ipi Está maior do que o restante", vbCritical
        txtbxIpi.Text = F_RestIpi
    End If
    P_CalcItem
End Sub

Private Sub txtbxIPI_Not_GotFocus()
    P_SelecTexto txtbxIPI_Not
End Sub

Private Sub txtbxIPI_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxIPI_Not_LostFocus()
    P_ValorLostFocus txtbxIPI_Not
End Sub

Private Sub txtbxNOme_Ter_GotFocus()
    P_SelecTexto txtbxNome_ter
End Sub

Private Sub txtbxNOme_Ter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxNum_Not_GotFocus()
    P_SelecTexto txtbxNum_Not
End Sub

Private Sub txtbxNum_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxNum_Not_LostFocus()
     txtbxNum_Not.Text = Format(F_SoNumero(txtbxNum_Not.Text), "0#######")
End Sub


Private Sub txtbxPrecoFuturo_GotFocus()
    P_SelecTexto txtbxPrecoFuturo
End Sub

Private Sub txtbxPrecoFuturo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPrecoFuturo_LostFocus()
    P_ValorLostFocus txtbxPrecoFuturo
    P_calcLucroLiq
End Sub

Function F_Qtd() As Currency
    If cmbxEmb.Text <> "" Then
        F_Qtd = CCur(cmbxEmb.Text) * CCur(txtbxQtd.Text)
    
        
    End If
End Function
Private Sub txtbxQtd_GotFocus()
        P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd, "###,##0.000"
    If txtbxIPI_Not.Text <> "0,00" Then
        txtbxIpi.Enabled = True
        lblIPI.Enabled = True
    Else
        txtbxIpi.Enabled = False
        lblIPI.Enabled = False
    End If
    
    P_Unit
    P_CalcItem
End Sub

Private Sub txtbxSerie_Not_GotFocus()
    P_SelecTexto txtbxSerie_Not
End Sub

Private Sub txtbxSerie_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSerie_Not_LostFocus()
    P_AjustaTexto txtbxSerie_Not
End Sub

Private Sub txtbxTotal_GotFocus()
    P_SelecTexto txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    P_ValorLostFocus txtbxTotal
    P_Unit
    P_CalcItem
End Sub

Private Sub txtbxValor_Not_GotFocus()
    P_SelecTexto txtbxValor_Not
End Sub

Private Sub txtbxValor_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Not_LostFocus()
    P_ValorLostFocus txtbxValor_Not
End Sub

Private Sub txtbxValorProd_GotFocus()
    P_SelecTexto txtbxValorProd
End Sub

Private Sub txtbxValorProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValorProd_LostFocus()
    P_ValorLostFocus txtbxValorProd
End Sub
