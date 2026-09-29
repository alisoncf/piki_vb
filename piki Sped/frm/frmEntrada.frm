VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntradaSped 
   Caption         =   "Entrada de Produtos"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14760
   Icon            =   "frmEntrada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   14760
   Begin VB.PictureBox pcProd 
      Align           =   2  'Align Bottom
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   14700
      TabIndex        =   159
      Top             =   5145
      Width           =   14760
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   66
         Top             =   240
         Width           =   735
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   5640
         ScaleHeight     =   825
         ScaleWidth      =   2625
         TabIndex        =   173
         Top             =   600
         Width           =   2655
         Begin VB.TextBox txtAliqIpiItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   93
            Text            =   "0,0"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtVlIpiItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   91
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtBcIpiItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   89
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label66 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Caption         =   "IPI"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   174
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label65 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2040
            TabIndex        =   92
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   1080
            TabIndex        =   90
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   825
         ScaleWidth      =   2625
         TabIndex        =   171
         Top             =   600
         Width           =   2655
         Begin VB.TextBox txtBcIcmsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   77
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtVlICMSItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   79
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtAliqIcmsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   81
            Text            =   "0,0"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   1080
            TabIndex        =   78
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
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2040
            TabIndex        =   80
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label75 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "ICMS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   172
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   11160
         ScaleHeight     =   825
         ScaleWidth      =   2625
         TabIndex        =   169
         Top             =   600
         Width           =   2655
         Begin VB.TextBox txtBCCofinsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   101
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtVlCofinsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   103
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtAliqCofinsItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   105
            Text            =   "0,0"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            Caption         =   "COFINS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   170
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label64 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2040
            TabIndex        =   104
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   1080
            TabIndex        =   102
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   8400
         ScaleHeight     =   825
         ScaleWidth      =   2625
         TabIndex        =   167
         Top             =   600
         Width           =   2655
         Begin VB.TextBox txtBcPisItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   95
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtVlPisItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   97
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtAliqPisItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   99
            Text            =   "0,0"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   1080
            TabIndex        =   96
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
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2040
            TabIndex        =   98
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "PIS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   168
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2880
         ScaleHeight     =   825
         ScaleWidth      =   2625
         TabIndex        =   165
         Top             =   600
         Width           =   2655
         Begin VB.TextBox txtAliqIcmsStItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   87
            Text            =   "0,0"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtVlIcmsStItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   85
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtBcIcmsStItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   83
            Text            =   "0,00"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "ICMS ST"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   166
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label60 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2160
            TabIndex        =   86
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor ($)"
            Height          =   195
            Left            =   1080
            TabIndex        =   84
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BC:"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox txtCst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   62
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNCM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   60
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCfop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   64
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbxEmb 
         Height          =   315
         Left            =   7920
         TabIndex        =   69
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox txtbxCodBarra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   24
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmndbtnConfirma 
         Caption         =   "F5 - &Inserir"
         Height          =   375
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   285
         Left            =   11160
         TabIndex        =   75
         Text            =   "0,00"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FEF4D8&
         Height          =   285
         Left            =   9240
         TabIndex        =   71
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AC67&
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   73
         Text            =   "0,0000"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Und Entr."
         Height          =   195
         Left            =   7080
         TabIndex        =   65
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST:"
         Height          =   195
         Left            =   5400
         TabIndex        =   61
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM:"
         Height          =   195
         Left            =   4320
         TabIndex        =   59
         Top             =   0
         Width           =   405
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP:"
         Height          =   195
         Left            =   6240
         TabIndex        =   63
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Emb.:"
         Height          =   195
         Left            =   7920
         TabIndex        =   68
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblExpr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   58
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
         TabIndex        =   72
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         Height          =   195
         Left            =   9240
         TabIndex        =   70
         Top             =   0
         Width           =   300
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   11160
         TabIndex        =   74
         Top             =   0
         Width           =   360
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
      ScaleWidth      =   14730
      TabIndex        =   154
      Top             =   8055
      Width           =   14760
      Begin VB.CommandButton frmExcluiItem 
         Caption         =   "F3 - Excluir Item"
         Height          =   375
         Left            =   6480
         TabIndex        =   176
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditItem 
         Caption         =   "F2 - Editar Item"
         Height          =   375
         Left            =   5040
         TabIndex        =   175
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "F12 - Sair"
         Height          =   495
         Left            =   12600
         TabIndex        =   158
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdConcluir 
         Caption         =   "F7 - Concluir NF"
         Height          =   495
         Left            =   10560
         TabIndex        =   157
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdNovoProd 
         Caption         =   "F4 - Novo Produto"
         Height          =   495
         Left            =   2160
         TabIndex        =   156
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdEditProd 
         Caption         =   "F4 - Edição de Produtos"
         Height          =   495
         Left            =   120
         TabIndex        =   155
         Top             =   120
         Width           =   1935
      End
   End
   Begin TabDlg.SSTab stb 
      Height          =   1935
      Left            =   0
      TabIndex        =   25
      Top             =   120
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   3413
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "F9 - Dados Gerais"
      TabPicture(0)   =   "frmEntrada.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "pcNota"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "F10 - Impostos"
      TabPicture(1)   =   "frmEntrada.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmndbtnEdit"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmndbtnAbreVenda"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "pcImpostos"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Outras informações"
      TabPicture(2)   =   "frmEntrada.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox pcImpostos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1320
         Left            =   120
         ScaleHeight     =   1290
         ScaleWidth      =   12045
         TabIndex        =   152
         Top             =   360
         Width           =   12075
         Begin VB.TextBox txtbxValorProd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   35
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   43
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxDepesas_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   45
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxIPI_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4440
            MaxLength       =   12
            TabIndex        =   47
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxValor_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   5520
            MaxLength       =   12
            TabIndex        =   37
            Tag             =   "!"
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxFrete 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   39
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxBC_Not 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   27
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
            TabIndex        =   29
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
            TabIndex        =   31
            Text            =   "0,00"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtVlIcmsStNot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   33
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
            TabIndex        =   41
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
            TabIndex        =   49
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
            TabIndex        =   51
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
            TabIndex        =   53
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtCofinsSTNot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8760
            MaxLength       =   12
            TabIndex        =   55
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
            Left            =   4440
            TabIndex        =   34
            Tag             =   "!"
            Top             =   120
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
            TabIndex        =   42
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
            TabIndex        =   44
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
            TabIndex        =   46
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor Contábil:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   36
            Tag             =   "!"
            Top             =   120
            Width           =   1020
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
            Left            =   11385
            TabIndex        =   153
            Top             =   120
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
            TabIndex        =   38
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
            TabIndex        =   26
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
            TabIndex        =   28
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
            TabIndex        =   30
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
            TabIndex        =   32
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
            TabIndex        =   40
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
            TabIndex        =   48
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
            TabIndex        =   50
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
            TabIndex        =   52
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
            TabIndex        =   54
            Top             =   720
            Width           =   885
         End
      End
      Begin VB.CommandButton cmndbtnAbreVenda 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   12360
         TabIndex        =   56
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   12360
         TabIndex        =   57
         Top             =   1080
         Width           =   975
      End
      Begin VB.PictureBox pcNota 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   -74880
         ScaleHeight     =   1425
         ScaleWidth      =   13305
         TabIndex        =   0
         Top             =   360
         Width           =   13335
         Begin VB.ComboBox cmbxTipo_Op 
            Height          =   315
            ItemData        =   "frmEntrada.frx":091E
            Left            =   9000
            List            =   "frmEntrada.frx":0920
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txtChaveNfe 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF80FF&
            Height          =   285
            Left            =   9000
            TabIndex        =   13
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox txtbxEmissao_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6840
            MaxLength       =   10
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxEntrada_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7920
            MaxLength       =   10
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbIndFre 
            Height          =   315
            ItemData        =   "frmEntrada.frx":0922
            Left            =   7680
            List            =   "frmEntrada.frx":0932
            TabIndex        =   21
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbIndPag 
            Height          =   315
            ItemData        =   "frmEntrada.frx":09B2
            Left            =   6360
            List            =   "frmEntrada.frx":09BF
            TabIndex        =   19
            Top             =   975
            Width           =   1215
         End
         Begin VB.ComboBox cmbxModelo 
            Height          =   315
            ItemData        =   "frmEntrada.frx":09E9
            Left            =   3840
            List            =   "frmEntrada.frx":09EB
            TabIndex        =   16
            Top             =   960
            Width           =   2415
         End
         Begin VB.ComboBox cmbxSituacao 
            Height          =   315
            ItemData        =   "frmEntrada.frx":09ED
            Left            =   120
            List            =   "frmEntrada.frx":09EF
            TabIndex        =   17
            Top             =   975
            Width           =   3615
         End
         Begin VB.TextBox txtbxSerie_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtbxNum_Not 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4680
            MaxLength       =   12
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtForn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   2
            Top             =   360
            Width           =   4215
         End
         Begin VB.CommandButton cmndbtnLocaliza 
            Height          =   285
            Left            =   4320
            Picture         =   "frmEntrada.frx":09F1
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "&Operação:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   9000
            TabIndex        =   22
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chave NFE:"
            Height          =   195
            Left            =   9000
            TabIndex        =   12
            Top             =   120
            Width           =   870
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "&Emissão: "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6840
            TabIndex        =   8
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "E&ntrada:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7920
            TabIndex        =   10
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador Frete"
            Height          =   195
            Left            =   7680
            TabIndex        =   20
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador Pgto"
            Height          =   195
            Left            =   6360
            TabIndex        =   18
            Top             =   735
            Width           =   1035
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo NF:"
            Height          =   195
            Left            =   3840
            TabIndex        =   15
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Situação NF:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   735
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Série:"
            Height          =   195
            Left            =   5760
            TabIndex        =   6
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Número:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4680
            TabIndex        =   4
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "&Fornecedor:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   14730
      TabIndex        =   107
      Top             =   6840
      Width           =   14760
      Begin VB.TextBox txtbxSujerido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   11880
         TabIndex        =   162
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxPrecoFuturo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   12840
         TabIndex        =   161
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox pcInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   0
         ScaleHeight     =   1110
         ScaleWidth      =   6465
         TabIndex        =   124
         Top             =   0
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
            TabIndex        =   138
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
            TabIndex        =   137
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
            TabIndex        =   136
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
            TabIndex        =   135
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
            TabIndex        =   134
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
            TabIndex        =   133
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
            TabIndex        =   132
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
            TabIndex        =   131
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
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   125
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
            TabIndex        =   151
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
            TabIndex        =   150
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
            TabIndex        =   149
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
            TabIndex        =   148
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
            TabIndex        =   147
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
            TabIndex        =   146
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
            TabIndex        =   145
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
            TabIndex        =   144
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
            TabIndex        =   143
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
            TabIndex        =   142
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   139
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox pcResu 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1DDB0&
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   6600
         ScaleHeight     =   1110
         ScaleWidth      =   5145
         TabIndex        =   108
         Top             =   0
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
            TabIndex        =   123
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
            TabIndex        =   122
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
            TabIndex        =   121
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
            TabIndex        =   120
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
            TabIndex        =   119
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
            TabIndex        =   118
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
            TabIndex        =   117
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
            TabIndex        =   116
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   109
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pç. Suj.:"
         Height          =   195
         Left            =   11880
         TabIndex        =   164
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pç. &Futuro:"
         Height          =   195
         Left            =   12840
         TabIndex        =   163
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.PictureBox pcOperacao 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00B1DDB0&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   14730
      TabIndex        =   67
      Top             =   0
      Width           =   14760
      Begin VB.PictureBox pcTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   14160
         ScaleHeight     =   135
         ScaleWidth      =   495
         TabIndex        =   177
         Top             =   0
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEntrada.frx":0D9C
      Height          =   2895
      Left            =   0
      TabIndex        =   106
      Top             =   2160
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   5106
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
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "seq_ite"
         Caption         =   "..."
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
         DataField       =   "cod_pro"
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
         DataField       =   "nome_pro"
         Caption         =   "Nome"
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
         DataField       =   "qtd_ite"
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
      BeginProperty Column04 
         DataField       =   "qtd_emb"
         Caption         =   "Emb"
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
         DataField       =   "unit_ite"
         Caption         =   "Preço"
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
      BeginProperty Column06 
         DataField       =   "total"
         Caption         =   "Total"
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
         DataField       =   "custo_ite"
         Caption         =   "Custo"
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
         DataField       =   "ipi_ite"
         Caption         =   "Ipi"
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
      BeginProperty Column09 
         DataField       =   "sujerido_ite"
         Caption         =   "Sujerido"
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
         DataField       =   "precofuturo_ite"
         Caption         =   "Pr. Futuro"
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
         DataField       =   "lucroliq_ite"
         Caption         =   "L.liq"
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
         DataField       =   "custosai_ite"
         Caption         =   "Custo Saida"
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
         DataField       =   "gastos_ite"
         Caption         =   "Gasto"
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
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3435,024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   645,165
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
End
Attribute VB_Name = "frmEntradaSped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean 'para saber se o frm já foi aberto
Dim iTam As Integer 'tam do form
Dim sTableCab As String
Dim sTableItens As String
Public operacao As String
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

Private Sub cmbIndFre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbIndPag_KeyDown(KeyCode As Integer, Shift As Integer)
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
    PerdeFoco cmbxEmb, "###,##0"
    confEmb
    txtbxQtd.SetFocus
End Sub
Sub confEmb()
    Dim strSQL As String
    Dim iCount As Integer
    Dim ok As Boolean
    On Error GoTo trata_erro
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
            conn.Execute strSQL
            P_PreencheEmbalagens txtbxCodBarra.Tag, cmbxEmb, True
            cmbxEmb.Text = cmbxEmb.Tag
            
        End If
        
    End If
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmbxModelo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSituacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Op_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
                & " and cod_ter ='" & txtForn.Tag & "' " _
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
    strSQL = strSQL & "'" & txtForn.Tag & "',"
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
    conn.Execute strSQL
    MsgBox "Nota incluída com sucesso!", vbInformation
    P_Clear
    txtForn.SetFocus
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub P_Clear()
    txtForn.Text = ""
    txtForn.Tag = ""
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

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnAbreVenda_Click()
    If Me.Tag = "" Then
        If abreNota = True Then
            HabCab False
        End If
    Else
        AlteraNota
        HabCab False
        HabItens True
    End If
End Sub
Function abreNota() As Boolean
    Dim strSQL As String
    Dim sSeq As String
    If txtForn.Tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtForn.SetFocus
        Exit Function
    End If
    If F_ValorContabilNota <> CCur(txtbxValor_Not.Text) Then
        MsgBox "OS valores não batem!", vbCritical
        Exit Function
    End If
    If txtbxValor_Not.Text = "0,00" Then
        On Error Resume Next
        MsgBox "Informe o Valor da Nota Fiscal!", vbExclamation
        txtbxValor_Not.SetFocus
        Exit Function
    End If
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtForn.Tag & "' " & _
        " and num_not='" & txtbxNum_Not.Text & "' and serie_not ='" & _
        txtbxSerie_Not.Text & "'") <> "" Then
        MsgBox "Já existe uma nota com este número e esta série para este fornecedor!", vbCritical
        Exit Function
    End If
    
    
    If MsgBox("Confirma Abertura da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Function
    End If
    
    'frBotoes.Visible = True
    sSeq = MaxKeySQL(sTableCab, "seq_not", "", "000000#")
    strSQL = "insert into " & sTableCab & "(seq_not,Cod_For,Num_not,Serie_not,desconto_not,valorprod_not,Valor_not,Bc_Not,"
    strSQL = strSQL & "frete_not,icms_not,ipi_not,despesas_not,operacao_not,emissao_not,"
    strSQL = strSQL & "entrada_not,cod_usu,status_not,data_not)Values("
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & "'" & txtForn.Tag & "', "
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
    'conn.Execute strSQL
    
    
    
    
    Me.Tag = sSeq
    lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.Text) + _
                                         CCur(txtbxFrete.Text))
    bActive = False
    Form_Activate
    'cmndbtnAddProd.SetFocus
    HabItens True
    txtbxCodBarra.SetFocus
    
    abreNota = True
    
End Function
Sub AlteraNota()
    Dim strSQL As String
    Dim sSeq As String
    If txtForn.Tag = "" Then
        MsgBox "Selecione o Terceiro primeiramente!", vbExclamation
        txtForn.SetFocus
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
    If F_BuscaUmCampo("num_not", sTableCab, "cod_for='" & txtForn.Tag & "' " & _
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
    strSQL = strSQL & "cod_for='" & txtForn.Tag & "', "
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
    conn.Execute strSQL
    
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
            txtForn.Tag & "' and cod_pro='" & p.cod_pro & "'") <> "" Then
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
    ElseIf Round((p.qtd_pro * p.preco) + CCur(lblValoLan.caption), 2) > CCur(txtbxValor_Not.Text) Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    ElseIf ((p.qtd_pro * p.preco)) = 0 Then
        MsgBox "O valor não está correto! tente novamente", vbExclamation
        Exit Sub
    End If
    
    
    cFator = (p.qtd_pro * p.preco) / (CCur(txtbxValor_Not.Text) - CCur(lblGastos.caption) - CCur(txtbxIPI_Not.Text))
    
    P_InsereProd p.cod_pro, p.preco, p.qtd_pro, p.ipi, cFator
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
        " and cod_for='" & txtForn.Tag & "'")
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
    strSQL = strSQL & "'" & txtForn.Tag & "',"
    strSQL = strSQL & "'" & txtbxSerie_Not.Text & "',"
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(unit) & ","
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(ipi) & ","
    
    strSQL = strSQL & F_GravaDec(cCusto) & ","
    strSQL = strSQL & F_GravaDec(cGastos) & ","
    strSQL = strSQL & sSeq & ") "
    conn.Execute strSQL
    
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
    strSQL = strSQL & "and cod_for='" & txtForn.Tag & "' "
    strSQL = strSQL & "and serie_not='" & txtbxSerie_Not.Text & "' "
    strSQL = strSQL & "and cod_pro= '" & cod_pro & "' "
    
    
    conn.Execute strSQL
    P_buscaitens
End Sub

Private Sub cmndbtnConfirmar_Click()
    
    If Me.Tag = "" Then
        If txtForn.Tag = "" Then
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
                txtbxNum_Not.Text <> txtbxNum_Not.Tag Then
            MsgBox "Esta Nota Fiscal já existe", vbExclamation
            txtbxNum_Not.SetFocus
            Exit Sub
        End If
    Else
        If F_BuscaUmCampo("Num_Not", "TAb_NotasFiscais", _
                "num_Not='" & txtbxNum_Not.Text & _
                "' and serie_not ='" & txtbxSerie_Not.Text & "' " _
                & " and cod_ter ='" & txtForn.Tag & "' " _
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
    conn.Execute strSQL
    MsgBox "Nota Alterada com sucesso!", vbInformation
    If frmEditaMov.Visible = True Then
        iPos = frmEditaMov.adc.Recordset.AbsolutePosition
        frmEditaMov.adc.Refresh
        frmEditaMov.adc.Recordset.Move iPos
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
        Dim preco As Currency
        Dim qtd_pro As Currency
        'On Error Resume Next
        If dtgrd.Row = -1 Then
            Exit Sub
        End If
        With adc.Recordset
            txtbxCusto.Text = FormatVal(!cf_pro)
            txtbxPreco.Text = FormatVal(!pv_pro)
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
    If MsgBox("Confirma exlusão do item " & adc.Recordset!nome_pro & "?", vbYesNo + vbCritical) = vbYes Then
        strSQL = "delete from " & sTableItens & " where SEQ_NOT='" & Me.Tag & "' "

        strSQL = strSQL & " and cod_pro='" & adc.Recordset!cod_pro & "' "
        conn.Execute strSQL
        
        If banco.RecordsAffected > 0 Then
            MsgBox "Item excuído! ", vbInformation
            soma
        Else
            MsgBox "Nenhum Registro Foi Afetado!", vbInformation
        End If
        
        
    End If
    P_buscaitens
    
    Exit Sub
trata_erro:
    MostraErro
End Sub

Private Sub cmndbtnForn_Click()

End Sub

Private Sub cmndbtnConfirma_Click()
    
    If Me.Tag = "" Then
        If MsgBox("A NF ainda não está aberta. Deseja abrir?", vbYesNo + vbQuestion) = vbYes Then
            If Not abreNota Then
                Exit Sub
            End If
        End If
    End If
    
    
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
    dtgrd.Enabled = opcao
    frPr.Enabled = opcao
End Sub
Sub HabCab(opcao As Boolean)
    pcTop.Enabled = opcao
    txtForn.Enabled = opcao
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
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtForn.Text)
    If f.cod_for <> "" Then
            txtForn.Text = f.nome_for
            txtForn.Tag = f.cod_for
            
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
            strSQL = strSQL & " where Seq_not='" & Me.Tag & "' "
            
            conn.Execute strSQL
            
        End With
        MsgBox "Nota Fiscal concluída!", vbInformation
'        If IsLoad(frmEdiEntr.Name) Then
'            frmEdiEntr.adc.Refresh
'        End If
        Unload Me
    End If
    
    
    
End Sub


Private Sub dtgrd_DblClick()
'    If cmndbtnEditProd.Enabled And cmndbtnEditProd.Visible = True Then
    
        editaItem
        On Error Resume Next
        txtbxQtd.SetFocus
 '   End If
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
    With adc.Recordset
    Err.Clear
    On Error GoTo trata_erro
        
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
trata_erro:
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
    conn.Execute strSQL
    P_buscaitens
    P_LimpaItem
    
End Sub
Sub RecalculaItens()
    
    With adc.Recordset
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
    
    conn.Execute strSQL
    cmndbtnConfirma.Tag = ""
    If opcao = 0 Then
        
        i = adc.Recordset.AbsolutePosition
    
        P_buscaitens
        adc.Recordset.Requery
        adc.Recordset.Move i
    
        P_LimpaItem
    End If
End Sub
Private Sub Form_Activate()
    If bActive = True Then Exit Sub
    
    If operacao = "E" Then
        sTableCab = "tab_cmp"
        sTableItens = "tab_itcmp"
        pcOperacao.BackColor = &H80000018
        pcTop.BackColor = vbYellow
    Else
        pcTop.BackColor = vbGreen
        sTableCab = "tab_compras"
        sTableItens = "tab_itcompras"
        pcOperacao.BackColor = &HB1DDB0
    End If
    
    
    
    
    
        If Me.Tag <> "" Then
            P_buscaitens
        lblGastos.caption = FormatVal(CCur(txtbxDepesas_Not.Text) + _
                             CCur(txtbxFrete.Text))
        End If
                             
    
    
    
    
    bActive = True
    'MudaAparencia Me
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
    adc.RecordSource = strSQL
    adc.Refresh
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
    resizeForm Me
    txtbxEntrada_Not.Text = Date
    
    Set ado = New ADODB.Recordset
    
    
    
    stb.Tab = 0
    P_PreencheOperacoes cmbxTipo_Op, "E"
    
    P_PreencheSituacaoDoc cmbxSituacao
    P_PreencheModeloNFTab411 cmbxModelo
    
    
    'cmbxBusca.ListIndex = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dtgrd.Width = Me.Width - 250
    dtgrd.Height = pcBot.Top - dtgrd.Top - 150
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




Private Sub txtBcIcmsItem_GotFocus()
    GanhaFoco txtBcIcmsItem
End Sub

Private Sub txtBcIcmsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtBcCofinsItem_LostFocus()
    PerdeFoco txtBCCofinsItem
End Sub

'
Private Sub txtVlIcmsItem_GotFocus()
    GanhaFoco txtVlICMSItem
End Sub

Private Sub txtVlIcmsItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtVlIcmsItem_LostFocus()
    PerdeFoco txtVlICMSItem
End Sub






Private Sub txtChaveNfe_GotFocus()
    GanhaFoco txtChaveNfe
End Sub

Private Sub txtChaveNfe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtChaveNfe_LostFocus()
    AjustaTexto txtChaveNfe, 0
End Sub

Private Sub txtPisNot_GotFocus()
    GanhaFoco txtPisNot
End Sub

Private Sub txtPisNot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtVlIcmsStItem_LostFocus()
    PerdeFoco txtVlICMSItem
End Sub


Private Sub txtVlIpiItem_GotFocus()
    GanhaFoco txtVlIpiItem
End Sub

Private Sub txtVlIpiItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBC_Not_LostFocus()
    PerdeFoco txtbxBC_Not
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
    GanhaFoco txtbxCodBarra
End Sub
Private Sub txtbxCodBarra_LostFocus()
    AjustaTexto txtbxCodBarra
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
    GanhaFoco txtbxDepesas_Not
End Sub

Private Sub txtbxDepesas_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFrete_LostFocus()
    PerdeFoco txtbxFrete
End Sub

Private Sub txtbxIcms_LostFocus()
    'PerdeFoco txtbxIcms
    txtbxICMSSaida.Text = CCur(F_IcmsProd(txtbxCodBarra.Tag, 3)) / 100 'percentual
    P_CalcItem
End Sub

Private Sub txtbxICMS_Not_GotFocus()
    GanhaFoco txtbxICMS_Not
End Sub

Private Sub txtbxICMS_Not_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxICMS_Not_LostFocus()
    PerdeFoco txtbxICMS_Not
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
    On Error GoTo trata_erro
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
trata_erro:
End Sub
Sub P_calcLucroLiq()
    Dim cLucro As Currency
    Dim cCusto As Currency
    Dim cIcms As Currency
    If CCur(txtbxSujerido.Text) > 0 Then
    On Error GoTo trata_erro
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
trata_erro:
    MsgBox Err.Description, vbCritical
    
End Sub
Sub P_Unit()
    If F_Qtd <> 0 Then
        txtbxUnit.Text = FormatVal(CCur(txtbxTotal.Text) / F_Qtd, "###,###,##0.0000")
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
        SendKeys Chr(9)
    End If
End Sub


Private Sub txtbxIPI_Not_GotFocus()
    GanhaFoco txtbxIPI_Not
End Sub

Private Sub txtbxIPI_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        txtbxNum_Not.Text = MaxKeySQL("tab_cmp", "num_not", "cod_for='" & txtForn.Tag & "'", "")
        
    End If
    GanhaFoco txtbxNum_Not
    
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
    GanhaFoco txtbxPrecoFuturo
End Sub

Private Sub txtbxPrecoFuturo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPrecoFuturo_LostFocus()
    PerdeFoco txtbxPrecoFuturo
    P_calcLucroLiq
End Sub

Function F_Qtd() As Currency
    If cmbxEmb.Text <> "" Then
        F_Qtd = CCur(cmbxEmb.Text) * CCur(txtbxQtd.Text)
    
        
    End If
End Function
Private Sub txtbxQtd_GotFocus()
        GanhaFoco txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    PerdeFoco txtbxQtd, "###,##0.000"
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
    GanhaFoco txtbxSerie_Not
End Sub

Private Sub txtbxSerie_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSerie_Not_LostFocus()
    AjustaTexto txtbxSerie_Not
End Sub

Private Sub txtbxTotal_GotFocus()
    GanhaFoco txtbxTotal
End Sub

Private Sub txtbxTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxTotal_LostFocus()
    PerdeFoco txtbxTotal
    P_Unit
    P_CalcItem
End Sub

Private Sub txtbxValor_Not_GotFocus()
    GanhaFoco txtbxValor_Not
End Sub

Private Sub txtbxValor_Not_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtNCM_LostFocus()
    AjustaTexto txtNCM
End Sub


Private Sub txtCST_GotFocus()
    GanhaFoco txtCst
End Sub

Private Sub txtCST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtCST_LostFocus()
    AjustaTexto txtCst
End Sub
Private Sub txtCFOP_GotFocus()
    GanhaFoco txtCfop
End Sub

Private Sub txtCFOP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtCFOP_LostFocus()
    AjustaTexto txtCfop
End Sub

