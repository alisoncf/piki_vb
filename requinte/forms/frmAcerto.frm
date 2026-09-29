VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAcerto 
   Caption         =   "Acerto de:"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17850
   Icon            =   "frmAcerto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   17850
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLembrete 
      Caption         =   "Lembrete"
      Height          =   495
      Left            =   16560
      Picture         =   "frmAcerto.frx":33E2
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   0
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Height          =   3015
      Left            =   0
      TabIndex        =   84
      Top             =   2760
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   7
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
         Caption         =   "Ref."
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
         DataField       =   "unit_ite"
         Caption         =   "Unit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "data_con"
         Caption         =   "Data"
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
            ColumnWidth     =   659,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2355,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4545,071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1124,787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2865,26
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcTotalItens 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   15075
      ScaleHeight     =   3270
      ScaleWidth      =   2745
      TabIndex        =   29
      Top             =   2670
      Width           =   2775
      Begin VB.TextBox txtbxDescontoAvista 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "25"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtbxDataBase 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtbxDifAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtbxValorNegociado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   20
         Text            =   "0,00"
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Negociado:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label lblTotalInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   1440
         TabIndex        =   24
         Tag             =   "!"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Inicial:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label44 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total dos Itens:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblDelExcluir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Del> Excluir"
         Height          =   195
         Left            =   1080
         MouseIcon       =   "frmAcerto.frx":385C
         MousePointer    =   4  'Icon
         TabIndex        =   30
         Top             =   2640
         Width           =   930
      End
      Begin VB.Label lblTotalAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1440
         TabIndex        =   26
         Tag             =   "!"
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.PictureBox pcCli 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   17820
      TabIndex        =   34
      Top             =   495
      Width           =   17850
      Begin TabDlg.SSTab stbTrab 
         DragMode        =   1  'Automatic
         Height          =   1350
         Left            =   9360
         TabIndex        =   88
         Top             =   600
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   2381
         _Version        =   393216
         Style           =   1
         MousePointer    =   2
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Órgão"
         TabPicture(0)   =   "frmAcerto.frx":3DE6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label18"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblGrupo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label27"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblEndGrupo"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label26"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblBairroGrupo"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label32"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblCidadeGrupo"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Info. Profissional"
         TabPicture(1)   =   "frmAcerto.frx":3E02
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture1"
         Tab(1).ControlCount=   1
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   -74880
            ScaleHeight     =   615
            ScaleWidth      =   6855
            TabIndex        =   89
            Top             =   360
            Width           =   6855
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Trabalho:"
               Height          =   195
               Left            =   0
               TabIndex        =   95
               Top             =   0
               Width           =   675
            End
            Begin VB.Label lblTrabalho 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
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
               Left            =   0
               TabIndex        =   94
               Top             =   240
               Width           =   45
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "End:"
               Height          =   195
               Left            =   1680
               TabIndex        =   93
               Top             =   0
               Width           =   330
            End
            Begin VB.Label lblEndTrab 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
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
               Left            =   1680
               TabIndex        =   92
               Top             =   240
               Width           =   45
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Setor"
               Height          =   195
               Left            =   4680
               TabIndex        =   91
               Top             =   0
               Width           =   375
            End
            Begin VB.Label lblSetorTrab 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
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
               Left            =   4680
               TabIndex        =   90
               Top             =   240
               Width           =   45
            End
         End
         Begin VB.Label lblCidadeGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000002&
            Height          =   555
            Left            =   7200
            TabIndex        =   103
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7200
            TabIndex        =   102
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblBairroGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000002&
            Height          =   555
            Left            =   4560
            TabIndex        =   101
            Top             =   600
            Width           =   2595
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4560
            TabIndex        =   100
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblEndGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            ForeColor       =   &H80000002&
            Height          =   555
            Left            =   2400
            TabIndex        =   99
            Top             =   600
            Width           =   2115
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2400
            TabIndex        =   98
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblGrupo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   555
            Left            =   120
            TabIndex        =   97
            Top             =   600
            Width           =   2205
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame frCli 
         Appearance      =   0  'Flat
         BackColor       =   &H0092E7A9&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         TabIndex        =   42
         Top             =   120
         Width           =   17655
         Begin VB.CommandButton cmdContasAReceber 
            Caption         =   "Contas"
            Height          =   255
            Left            =   2400
            TabIndex        =   131
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdZerar 
            Caption         =   "Zerar Estoque"
            Height          =   495
            Left            =   5160
            TabIndex        =   125
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmndbtnEditarCliente 
            Caption         =   "Abrir Cadastro"
            Height          =   255
            Left            =   3720
            TabIndex        =   83
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optnCli 
            Appearance      =   0  'Flat
            BackColor       =   &H0092E7A9&
            Caption         =   "&Cliente"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optnRep 
            Appearance      =   0  'Flat
            BackColor       =   &H0092E7A9&
            Caption         =   "&Representante"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   960
            TabIndex        =   3
            Top             =   0
            Width           =   1455
         End
         Begin VB.TextBox txtbxClienteAcerto 
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
            TabIndex        =   5
            Top             =   480
            Width           =   4575
         End
         Begin VB.CommandButton cmndbtnClienteAcerto 
            Height          =   285
            Left            =   4680
            Picture         =   "frmAcerto.frx":3E1E
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   15720
            TabIndex        =   87
            Top             =   0
            Width           =   675
         End
         Begin VB.Label lblTelefone 
            BackColor       =   &H007DBFDB&
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
            Height          =   525
            Left            =   15720
            TabIndex        =   86
            Top             =   240
            Width           =   1770
         End
         Begin VB.Label lblCidadeCli 
            BackColor       =   &H007DBFDB&
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
            Height          =   525
            Left            =   14160
            TabIndex        =   52
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label lblBairroCli 
            BackColor       =   &H007DBFDB&
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
            Height          =   525
            Left            =   11640
            TabIndex        =   51
            Top             =   240
            Width           =   2490
         End
         Begin VB.Label lblEnderecoCli 
            BackColor       =   &H007DBFDB&
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
            Height          =   525
            Left            =   8640
            TabIndex        =   50
            Top             =   240
            Width           =   2970
         End
         Begin VB.Label lblRgCli 
            BackColor       =   &H007DBFDB&
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
            Height          =   525
            Left            =   7200
            TabIndex        =   49
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label lblCpfCli 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5160
            TabIndex        =   48
            Top             =   240
            Width           =   2010
         End
         Begin VB.Label lblCli 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "C&liente (F4):"
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
            Left            =   120
            TabIndex        =   4
            Tag             =   "C"
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label40 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "CPF:"
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
            Left            =   5160
            TabIndex        =   47
            Top             =   0
            Width           =   420
         End
         Begin VB.Label Label39 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8640
            TabIndex        =   46
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   11640
            TabIndex        =   45
            Top             =   0
            Width           =   450
         End
         Begin VB.Label Label37 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   14160
            TabIndex        =   44
            Top             =   0
            Width           =   540
         End
         Begin VB.Label label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "RG:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7200
            TabIndex        =   43
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.PictureBox pcProdAcerto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   9495
         TabIndex        =   35
         Top             =   1080
         Width           =   9495
         Begin VB.TextBox txtbxReman 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   10
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtbxVlrFinalAcerto 
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
            Left            =   7080
            TabIndex        =   16
            Text            =   "0,00"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtbxDiferenca 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   8280
            TabIndex        =   18
            Text            =   "0,00"
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtbxUnitAcerto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   6120
            TabIndex        =   14
            Text            =   "0,00"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxQtdAcerto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   5520
            TabIndex        =   12
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtbxProdAcerto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label lblSeqConAcerto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Left            =   1080
            TabIndex        =   41
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seq. Con:"
            ForeColor       =   &H00CF614E&
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   705
         End
         Begin VB.Label lblDataConAcerto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Left            =   2520
            TabIndex        =   39
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Lanç:"
            ForeColor       =   &H00CF614E&
            Height          =   195
            Left            =   1560
            TabIndex        =   38
            Top             =   600
            Width           =   795
         End
         Begin VB.Label lblRefAcerto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Left            =   3960
            TabIndex        =   37
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            ForeColor       =   &H00CF614E&
            Height          =   195
            Left            =   3000
            TabIndex        =   36
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reman:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4920
            TabIndex        =   9
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor Final:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7080
            TabIndex        =   15
            Top             =   0
            Width           =   780
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença:"
            Height          =   195
            Left            =   8280
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Unitário:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6120
            TabIndex        =   13
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto (F2)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   870
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5640
            TabIndex        =   11
            Top             =   0
            Width           =   300
         End
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   17850
      TabIndex        =   31
      Top             =   0
      Width           =   17850
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   10455
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   495
         Left            =   10440
         Picture         =   "frmAcerto.frx":3F17
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11640
         TabIndex        =   33
         Top             =   45
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   11160
         TabIndex        =   32
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   17850
      TabIndex        =   53
      Top             =   5970
      Width           =   17850
      Begin VB.CommandButton cmndbtrnCancel 
         Caption         =   "F9 - Cancelar"
         Height          =   735
         Left            =   13320
         Picture         =   "frmAcerto.frx":4BE1
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnConfAcerto 
         Caption         =   "F7 - Confirmar"
         Height          =   735
         Left            =   10920
         Picture         =   "frmAcerto.frx":58AB
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "F12 - Sair"
         Height          =   735
         Left            =   15000
         Picture         =   "frmAcerto.frx":5A22
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   2040
         Width           =   1575
      End
      Begin VB.PictureBox pcTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H007DB4BD&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   10920
         ScaleHeight     =   705
         ScaleWidth      =   3465
         TabIndex        =   61
         Top             =   1200
         Width           =   3495
         Begin VB.Label lblVlrAcertoDiv 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   825
            Left            =   1200
            TabIndex        =   64
            Tag             =   "!"
            Top             =   -120
            Width           =   1260
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
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
            Left            =   840
            TabIndex        =   63
            Top             =   240
            Width           =   195
         End
         Begin VB.Label lblParcelas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   825
            Left            =   360
            TabIndex        =   62
            Top             =   -90
            Width           =   360
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   10920
         ScaleHeight     =   945
         ScaleWidth      =   3465
         TabIndex        =   54
         Top             =   120
         Width           =   3495
         Begin VB.Label lblValorFinal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            TabIndex        =   60
            Tag             =   "!"
            Top             =   292
            Width           =   540
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Final:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblCapComiss 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comissão:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   75
            Width           =   720
         End
         Begin VB.Label lblComissao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            TabIndex        =   57
            Tag             =   "!"
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblTotalDistribuido 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            TabIndex        =   56
            Tag             =   "!"
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Distribuído:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   675
            Width           =   1215
         End
      End
      Begin TabDlg.SSTab sstb 
         Height          =   2775
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "F8 - A Prazo"
         TabPicture(0)   =   "frmAcerto.frx":5B8D
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdRedistribuir"
         Tab(0).Control(1)=   "txtParcelas"
         Tab(0).Control(2)=   "txtbxTotalAprazo"
         Tab(0).Control(3)=   "txtbxEntrada"
         Tab(0).Control(4)=   "Picture3"
         Tab(0).Control(5)=   "txtbxVenIniAcerto"
         Tab(0).Control(6)=   "txtbxTipoCob"
         Tab(0).Control(7)=   "txtbxMsf"
         Tab(0).Control(8)=   "cmbxOpc"
         Tab(0).Control(9)=   "txtbxDescCob"
         Tab(0).Control(10)=   "txtbxFormaPag"
         Tab(0).Control(11)=   "msflxgrd"
         Tab(0).Control(12)=   "cmndbtnFormasPag"
         Tab(0).Control(13)=   "cmndbtnCob"
         Tab(0).Control(14)=   "Label21"
         Tab(0).Control(15)=   "Label8"
         Tab(0).Control(16)=   "Label43"
         Tab(0).Control(17)=   "lblTipoCob"
         Tab(0).Control(18)=   "Label11"
         Tab(0).Control(19)=   "Label12"
         Tab(0).Control(20)=   "lblTipoPag"
         Tab(0).Control(21)=   "Label4"
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "F9 - À Vista"
         TabPicture(1)   =   "frmAcerto.frx":5BA9
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "msflxgrdAv"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Picture5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmndbtnCh"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmndbtnDin"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmndbtnCartao"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "optDin"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Option1"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Option2"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "optnPix"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "cmdPix"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).ControlCount=   10
         Begin VB.CommandButton cmdPix 
            Caption         =   "Adicionar Pagamento PIX"
            Height          =   495
            Left            =   4440
            TabIndex        =   130
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton optnPix 
            Caption         =   "PIX"
            Height          =   195
            Left            =   9000
            TabIndex        =   129
            Top             =   1680
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton cmdRedistribuir 
            Enabled         =   0   'False
            Height          =   375
            Left            =   -66120
            Picture         =   "frmAcerto.frx":5BC5
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtParcelas 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -66720
            TabIndex        =   127
            Top             =   600
            Width           =   495
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Cartão Parcelado"
            Height          =   195
            Left            =   9000
            TabIndex        =   124
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cartão A Vista"
            Height          =   195
            Left            =   9000
            TabIndex        =   123
            Top             =   1200
            Width           =   1335
         End
         Begin VB.OptionButton optDin 
            Caption         =   "Dinheiro"
            Height          =   195
            Left            =   9000
            TabIndex        =   122
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxTotalAprazo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -67800
            TabIndex        =   80
            Text            =   "0,00"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtbxEntrada 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68760
            TabIndex        =   78
            Text            =   "0,00"
            Top             =   600
            Width           =   855
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   -65520
            ScaleHeight     =   465
            ScaleWidth      =   945
            TabIndex        =   115
            Top             =   480
            Width           =   975
            Begin VB.CommandButton cmndbtnExclAp 
               DragMode        =   1  'Automatic
               Enabled         =   0   'False
               Height          =   375
               Left            =   50
               Picture         =   "frmAcerto.frx":5F4F
               Style           =   1  'Graphical
               TabIndex        =   117
               Top             =   50
               Width           =   375
            End
            Begin VB.CommandButton cmndbtnEditAp 
               Height          =   375
               Left            =   480
               Picture         =   "frmAcerto.frx":62D9
               Style           =   1  'Graphical
               TabIndex        =   116
               Top             =   50
               Width           =   375
            End
         End
         Begin VB.CommandButton cmndbtnCartao 
            Caption         =   "Adcionar Pgto em Cartão"
            Height          =   495
            Left            =   6600
            TabIndex        =   114
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnDin 
            Caption         =   "Adcionar Pgto em Dinheiro"
            Height          =   495
            Left            =   200
            TabIndex        =   113
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtbxVlrFinalAvista 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   -65400
            TabIndex        =   112
            Text            =   "0"
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox txtbxVenIniAcerto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72480
            TabIndex        =   71
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtbxTipoCob 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -74880
            TabIndex        =   68
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtbxMsf 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   -72360
            TabIndex        =   111
            Top             =   1620
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cmbxOpc 
            Height          =   315
            ItemData        =   "frmAcerto.frx":6CDB
            Left            =   -74400
            List            =   "frmAcerto.frx":6CE5
            Style           =   2  'Dropdown List
            TabIndex        =   110
            Top             =   1860
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmndbtnCh 
            Caption         =   "Adcionar Pgto em Cheque"
            Height          =   495
            Left            =   2280
            TabIndex        =   109
            Top             =   360
            Width           =   2055
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   9480
            ScaleHeight     =   465
            ScaleWidth      =   945
            TabIndex        =   106
            Top             =   360
            Width           =   975
            Begin VB.CommandButton cmndbtnExclAv 
               Height          =   375
               Left            =   50
               Picture         =   "frmAcerto.frx":6D08
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   50
               Width           =   375
            End
            Begin VB.CommandButton cmndbnEditAv 
               Height          =   375
               Left            =   480
               Picture         =   "frmAcerto.frx":7092
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   50
               Width           =   375
            End
         End
         Begin VB.TextBox txtbxDescCob 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -71280
            TabIndex        =   73
            Text            =   "0"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtbxFormaPag 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70800
            TabIndex        =   75
            Top             =   600
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrd 
            Height          =   1575
            Left            =   -74880
            TabIndex        =   118
            Top             =   1080
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdAv 
            Height          =   1575
            Left            =   240
            TabIndex        =   119
            Top             =   960
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.CommandButton cmndbtnFormasPag 
            Height          =   285
            Left            =   -69120
            Picture         =   "frmAcerto.frx":7A94
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnCob 
            Height          =   285
            Left            =   -72840
            Picture         =   "frmAcerto.frx":7E1E
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Redistribuir parc."
            Enabled         =   0   'False
            Height          =   195
            Left            =   -66720
            TabIndex        =   126
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(%)"
            Height          =   195
            Left            =   -71280
            TabIndex        =   72
            Top             =   360
            Width           =   210
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Final:"
            Height          =   195
            Left            =   -66360
            TabIndex        =   121
            Top             =   2760
            Width           =   780
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ven. Ini:"
            Height          =   195
            Left            =   -72480
            TabIndex        =   70
            Top             =   360
            Width           =   585
         End
         Begin VB.Label lblTipoCob 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrança:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   67
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total A Prazo"
            Height          =   195
            Left            =   -67800
            TabIndex        =   79
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada:"
            Height          =   195
            Left            =   -68760
            TabIndex        =   77
            Top             =   360
            Width           =   600
         End
         Begin VB.Label lblTipoPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DIN"
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
            Left            =   -69120
            TabIndex        =   120
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Forma de Pagamento:"
            Height          =   195
            Left            =   -70800
            TabIndex        =   74
            Top             =   360
            Width           =   1560
         End
      End
      Begin VB.Label lblAcerto_Promocao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "Acerto Promocional"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   14640
         TabIndex        =   105
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAcerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sX As Single
Dim sY As Single
Public acerto_promocao  As Boolean

Dim bCpfInvalido As Boolean
Dim bClienteNegativado As Boolean
Dim WithEvents adoAcerto As ADODB.Recordset
Attribute adoAcerto.VB_VarHelpID = -1
Dim bEnterCell As Boolean
Dim prod As ProdutoSaida
  
Dim bDirecao As Integer
Dim bActive As Boolean
Dim preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer

Public Atacado As Boolean

Private Type prodPromo
    nome As String
    preco As String
End Type
Dim pPromo() As prodPromo
Dim QtdPromo As Integer
Dim iIndPromo As Integer
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim cFile As New Collection
Dim ai As New ADODB.stream
Dim Indice As Integer
Option Explicit









Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub














Private Sub cmdCartaoDividido_Click()
    Dim strSQL As String
    Dim p As pagamentocartao
    p = F_DadosCartao()
    If p.valor > "0" Then
        strSQL = Date & Chr(9)
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        If p.operacao = "C" Then
            strSQL = strSQL & "3 - Cartão de Crédito" & Chr(9) & Chr(9)
        Else
            strSQL = strSQL & "4 - Cartão de Débito" & Chr(9) & Chr(9)
        End If
        strSQL = strSQL & p.cod_fin & Chr(9)
        strSQL = strSQL & p.parcelas & Chr(9)
        
        
        msflxgrdAv.AddItem strSQL
        
    End If
    P_SomaItd
End Sub

Private Sub cmdContasAReceber_Click()
    If txtbxClienteAcerto.tag = "" Then
        MsgBox "Escolha um cliente", vbExclamation
        Exit Sub
    End If
    frmEditaMov.txtbxTerceiro.text = txtbxClienteAcerto.text
    If optnCli.Value = True Then
        frmEditaMov.cmbxTipoTer.ListIndex = 0
    Else
        frmEditaMov.cmbxTipoTer.ListIndex = 1
    End If
    
    frmEditaMov.tag = "x"
    frmEditaMov.Show 0
    
    frmEditaMov.SetFocus

End Sub

Private Sub cmdLembrete_Click()
    buscaLembreteAcerto txtbxRepresentante.tag, True
End Sub

Private Sub cmdPix_Click()
    Dim p As parcela
    Dim strSQL As String
    p = F_DadosPar(Date)
    If p.valor > 0 Then
        strSQL = Date & Chr(9)
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        strSQL = strSQL & "4 - PIX" & Chr(9) & Chr(9) & p.complemento & Chr(9) & Chr(9) & Chr(9)
        
     '   strSQL = strSQL & FormatVal(p.valor) & Chr(9)
       ' strSQL = strSQL & Date
        msflxgrdAv.AddItem strSQL
        
        
    End If
    P_SomaItd
End Sub

Private Sub cmdZerar_Click()
    
    
    If txtbxRepresentante.tag = "" Then
        MsgBox "Informe o representante!", vbExclamation
        txtbxRepresentante.SetFocus
        Exit Sub
    End If
    If txtbxClienteAcerto.tag = "" Then
        MsgBox "Informe o representante!", vbExclamation
        txtbxClienteAcerto.SetFocus
        Exit Sub
    End If
    If ExisteAcertoEmAberto Then
        Exit Sub
    End If
    
    
    If MsgBox("Confirma a busca de todo o estoque do consignado!", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    
    ZerarEstoque
End Sub

'Private Sub adc_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    On Error Resume Next
'    adc.caption = adoacerto.AbsolutePosition + 1 & " / " & adoacerto.RecordCount & " " & adoacerto!nome_pro
'
'End Sub

Private Sub cmndbnEditAv_Click()
    If msflxgrdAv.row <= 0 Then
        Exit Sub
    End If
    If Left(msflxgrdAv.text, 1) = "1" Then
        'editaParcela
    Else
        editaChequeAv
    End If
End Sub

Private Sub cmndbtnEditarCliente_Click()
    If txtbxClienteAcerto.tag <> "" Then
        If optnCli.Value Then
            Dim fr As New frmCadCli
            
            strHora = Time
    
            fr.tag = txtbxClienteAcerto.tag
    
            fr.Show 0
        Else
                Dim frm As New frmCadRep
            
                
                strHora = Time
                frm.tag = txtbxClienteAcerto.tag
    
                frm.Show 0
        End If
    Else
        Dim fr1 As New frmCadCli
        fr1.Show 0
        
    End If
End Sub

Private Sub cmndbtnExclAv_Click()
    If MsgBox("Deseja excluir o lançamento?", vbYesNo + vbCritical) = vbYes Then
        If msflxgrdAv.rows = 2 Then
            msflxgrdAv.rows = 1
        Else
            msflxgrdAv.RemoveItem msflxgrdAv.row
        End If
    End If
End Sub

Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    Dim strSQL As String
    P_CalcTotalAprazo
    f = F_SelectFormasPag(txtbxFormaPag.text, "N")
    
    If f.cod_For <> "" Then
        
        txtbxFormaPag.text = f.nome_for
        txtbxFormaPag.tag = f.cod_For
        
        
        lblParcelas.caption = f.parcelas_for
        txtParcelas.text = f.parcelas_for
        lblTipoPag.caption = f.tipo_cai
        lblTipoPag.tag = f.cod_tip
        
        
        distribuiParcela
        somaDesconto
        
        
    End If
    
    
    
End Sub


Private Sub cmbxFormaspag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        msflxgrd.SetFocus
    End If
    
End Sub

Private Sub cmbxOpc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxOpc_LostFocus()
    msflxgrd.text = cmbxOpc.text
    If Left(msflxgrd.text, 1) = "1" Then
        editaParcela
    Else
        editaCheque
    End If
    
    cmbxOpc.Visible = False
End Sub
Sub editaParcela()
    Dim p As parcela
    Dim sVen As String
    Dim sVal As String
    msflxgrd.col = 0
    sVen = msflxgrd.text
    msflxgrd.col = 1
    sVal = msflxgrd.text
    Exit Sub
    p = F_DadosPar(sVen, sVal)
    If p.valor > 0 Then
        If p.vencimento = "" Then
            MsgBox "O vencimento foi informado errado!", vbExclamation
            Exit Sub
        End If
        msflxgrd.col = 0
        msflxgrd.text = p.vencimento
        msflxgrd.col = 1
        msflxgrd.text = p.valor
                
        
    End If
End Sub
Sub editaCheque()
    Dim c As cheque
    With msflxgrd
        .col = 0
        c.Venc_che = .text
        .col = 1: c.Valor_che = .text
        .col = 3: c.num_che = .text
        .col = 4: c.cod_bco = .text
        .col = 5: c.age_che = .text
        .col = 6: c.cco_che = .text
        .col = 7: c.nome_che = .text
        .col = 8: c.cpf_che = .text
        
        If c.cpf_che = "" Then
            c.cpf_che = lblcpf.caption
            
        End If
        If c.nome_che = "" Then
            c.nome_che = txtbxClienteAcerto.text
        End If
        c = F_DadosCheque(True, c.Venc_che, c.nome_che, c.num_che, c.cco_che, c.age_che, c.cod_bco, CStr(c.Valor_che), c.cpf_che)
        If c.num_che <> "" Then
        
            .col = 1: .text = c.Valor_che
            .col = 0: .text = c.Venc_che
            .col = 3: .text = c.num_che
            .col = 4: .text = c.cod_bco
            .col = 5: .text = c.age_che
            .col = 6: .text = c.cco_che
            .col = 7: .text = c.nome_che
            .col = 8: .text = c.cpf_che
            
        End If
        
        
    End With
            
            
            
End Sub
Sub editaChequeAv()
    Dim c As cheque
    With msflxgrdAv
        .col = 0
        c.Venc_che = .text
        .col = 1: c.Valor_che = .text
        .col = 3: c.num_che = .text
        .col = 4: c.cod_bco = .text
        .col = 5: c.age_che = .text
        .col = 6: c.cco_che = .text
        .col = 7: c.nome_che = .text
        .col = 8: c.cpf_che = .text
        
        If c.cpf_che = "" Then
            c.cpf_che = lblcpf.caption
            
        End If
        If c.nome_che = "" Then
            c.nome_che = txtbxClienteAcerto.text
        End If
        c = F_DadosCheque(True, c.Venc_che, c.nome_che, c.num_che, c.cco_che, c.age_che, c.cod_bco, CStr(c.Valor_che), c.cpf_che)
        If c.num_che <> "" Then
        
            .col = 1: .text = c.Valor_che
            .col = 0: .text = c.Venc_che
            .col = 3: .text = c.num_che
            .col = 4: .text = c.cod_bco
            .col = 5: .text = c.age_che
            .col = 6: .text = c.cco_che
            .col = 7: .text = c.nome_che
            .col = 8: .text = c.cpf_che
            
        End If
        
        
    End With
            
            
            
End Sub

Private Sub cmndbtnBuscaRep_Click()
    Dim r As representante.representante
    
    If Me.tag <> "" Then
        MsgBox "Não é possível trocar o representante no meio de uma operação!", vbExclamation
        Exit Sub
    End If
    
    r = F_BuscaRep(txtbxRepresentante.text)
    
    If r.cod_Rep <> "" Then
        
        
        If txtbxRepresentante.tag <> r.cod_Rep Then
            msflxgrd.rows = 1
            'msflxgrdAcerto.rows = 1
            msflxgrd.rows = 1
            P_SomaItd
        End If
        txtbxRepresentante.tag = r.cod_Rep
        txtbxRepresentante.text = r.nome_rep
        lblcpf.caption = r.cpf_rep
        
        Me.caption = "Acerto de: " & r.nome_rep
        BuscaMensagem "R", r.cod_Rep
        buscaLembreteAcerto r.cod_Rep
        txtbxClienteAcerto.SetFocus
    End If
End Sub

Private Sub cmndbtnCartao_Click()
    Dim strSQL As String
    Dim p As pagamentocartao
    p = F_DadosCartao()
    If p.valor > "0" Then
        strSQL = Date & Chr(9)
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        If p.operacao = "C" Then
            strSQL = strSQL & "3 - Cartão de Crédito" & Chr(9) & Chr(9)
        Else
            strSQL = strSQL & "4 - Cartão de Débito" & Chr(9) & Chr(9)
        End If
        strSQL = strSQL & p.cod_fin & Chr(9)
        strSQL = strSQL & p.parcelas & Chr(9)
        
        
        msflxgrdAv.AddItem strSQL
        
    End If
    P_SomaItd
End Sub

Private Sub cmndbtnCh_Click()
    Dim c As cheque
    Dim strSQL As String
    c = F_DadosCheque(True, , txtbxClienteAcerto.text, "", "", "", "", "", lblcpf.caption)
    
    If c.Valor_che > 0 Then
        
        strSQL = c.Venc_che & Chr(9)
        strSQL = strSQL & FormatVal(c.Valor_che) & Chr(9)
        strSQL = strSQL & "2 - Cheque-pré" & Chr(9)
        strSQL = strSQL & c.num_che & Chr(9)
        strSQL = strSQL & c.cod_bco & Chr(9)
        strSQL = strSQL & c.age_che & Chr(9)
        strSQL = strSQL & c.cco_che & Chr(9)
        strSQL = strSQL & c.nome_che & Chr(9)
        strSQL = strSQL & c.cpf_che & Chr(9)
        strSQL = strSQL & "" & Chr(9)
        
        msflxgrdAv.AddItem strSQL
        
    End If
    P_SomaItd
End Sub

Private Sub cmndbtnClienteAcerto_Click()
    Dim c As Clientes.cliente
    If optnRep.Value = True Then
        MsgBox "Você optou por fazer acerto para o próprio representante, não é possível escolher qualquer cliente!", vbExclamation
        optnRep.SetFocus
        Exit Sub
    End If
    c = Clientes.F_BuscaCliente(txtbxClienteAcerto.text)
    If c.cod_cli <> "" Then
        txtbxClienteAcerto.text = c.nome_cli
        txtbxClienteAcerto.tag = c.cod_cli
        lblCpfCli.caption = c.cnpj_Cli
        lblRgCli.caption = c.ie_Cli
        If c.status_cli = "2" Then
            MsgBox "O Cliente encontra-se negativado! Emita essa mensagem ao Representante", vbExclamation
            bClienteNegativado = True
            
            Exit Sub
        Else
            bClienteNegativado = False
        End If
        lblEnderecoCli.caption = StrConv(c.endereco_cli, vbProperCase)
        lblBairroCli.caption = StrConv(c.bairro_cli, vbProperCase)
        lblCidadeCli.caption = StrConv(c.cidade_cli, vbProperCase)
        lblTrabalho.caption = c.localtrab_cli
        lblEndTrab.caption = c.endtrab_cli
        lblSetorTrab.caption = c.setortrab.nome & "( Cobrador " & F_BuscaContaCobPorCod(c.setortrab.cobrador_set).nome_cob & ")"
        lblTelefone.caption = c.telefone_cli
        
        lblGrupo.caption = c.org.nome_org
        lblBairroGrupo.caption = c.org.nome_set
        lblCidadeGrupo.caption = c.org.nome_cid
        lblEndGrupo.caption = c.org.endereco_org
        
        
        
        If F_CPF(c.cnpj_Cli) = False Then
            If F_CGC(c.cnpj_Cli) = False Then
                MsgBox "O CPF do cliente é inválido! Não é possível concluir o acerto sem corrigi-lo", vbExclamation
                bCpfInvalido = True
                lblCpfCli.BackColor = vbRed
                lblCpfCli.ForeColor = vbRed
                
            End If
        Else
            bCpfInvalido = False
            lblCpfCli.BackColor = vbWhite
            lblCpfCli.ForeColor = vbBlue
        End If
        txtbxProdAcerto.SetFocus
    End If
    atualiza_acerto
End Sub

Private Sub cmndbtnCob_Click()
    Dim t As tipocob
    t = F_BuscaTipoCobRep(txtbxRepresentante.tag)
    If t.cod_tip <> "" Then
        txtbxTipoCob.tag = t.cod_tip
        txtbxTipoCob.text = t.nome_tip
        txtbxDescCob.text = t.comiss_tip
        somaDesconto
    End If
    distribuiParcela
    txtbxVenIniAcerto.SetFocus
End Sub
Sub somaDesconto()
    Dim cCom As Currency
    
    If acerto_promocao Then
        txtbxDescontoAvista.text = "5"
        txtbxDescCob.text = "5"
        
        
    End If
        
    cCom = CCur(txtbxDescCob.text) / 100 * CCur(lblTotalInicial.caption)
    cCom = cCom + CCur(txtbxDifAcerto.text)
    lblComissao.caption = FormatVal(cCom)
    
    If optnRep.Value Then
        'lblCapComiss.caption = "Desconto"
        lblValorFinal.caption = FormatVal(CCur(lblTotalAcerto.caption) - CCur(lblComissao.caption))
    Else
        'lblCapComiss.caption = "Comissão"
        lblValorFinal.caption = lblTotalAcerto.caption
    End If
    
End Sub
Private Sub ConfAprazo()
    On Error GoTo trata_Erro
    P_Status ""
    If lblTotalDistribuido.caption <> lblValorFinal.caption Then
        MsgBox "O Valor distribuído é de " & lblTotalDistribuido.caption _
        & ". No entanto o Deve ser " & lblValorFinal.caption, vbExclamation
        Exit Sub
    End If
    If adoAcerto.RecordCount <= 0 Then
        MsgBox "Não há o que lançar!", vbExclamation
        Exit Sub
    End If
    If txtbxClienteAcerto.tag = "" Then
        MsgBox "Informe o cliente do acerto!", vbExclamation
        txtbxClienteAcerto.SetFocus
        Exit Sub
    End If
    
    If txtbxRepresentante.tag = "" Then
        MsgBox "Não há um representante escolhido!", vbExclamation
        txtbxRepresentante.SetFocus
        Exit Sub
    End If
    If txtbxFormaPag.tag = "" Then
        MsgBox "Escolha a forma de pagamento!", vbExclamation
        Exit Sub
    End If
    
    If txtbxVenIniAcerto.text = "" Then
        MsgBox "Informe a data para o vencimento inicial!", vbExclamation
        txtbxVenIniAcerto.SetFocus
        Exit Sub
    End If
    If txtbxTipoCob.tag = "" Or txtbxTipoCob.tag = "0" Then
        MsgBox "Escolha o tipo de cobrança!", vbExclamation
        cmndbtnCob_Click
    End If
    If bClienteNegativado = True Then
        MsgBox "Não é possível concluir o acerto por que o cliente encontra-se bloqueado!", vbExclamation
        Exit Sub
    End If
    P_Status ""
    If MsgBox("Confirma lançamento de acerto para o cliente " & txtbxClienteAcerto.text & "?", vbYesNo + vbQuestion) = vbNo Then
        
        Exit Sub
    End If
    
    P_Status "Lançando acerto..."
    
    lancaAcerto
    P_Status ""
Exit Sub
trata_Erro:
P_MostraErro

End Sub
Private Sub ConfAvista()
    
    If lblTotalDistribuido.caption <> lblValorFinal.caption Then
        MsgBox "O Valor distribuído é de " & lblTotalDistribuido.caption _
        & ". No entanto o Deve ser " & lblValorFinal.caption, vbExclamation
        Exit Sub
    End If
    If optnCli.Value = True Then
        MsgBox "Pagamento À Vista apenas para o Representante!", vbExclamation
        optnCli.SetFocus
        Exit Sub
        
    End If
    If adoAcerto.RecordCount <= 0 Then
        MsgBox "Não há o que lançar!", vbExclamation
        Exit Sub
    End If
    If txtbxClienteAcerto.tag = "" Then
        MsgBox "Informe o cliente do acerto!", vbExclamation
        txtbxClienteAcerto.SetFocus
        Exit Sub
    End If
    
    If txtbxRepresentante.tag = "" Then
        MsgBox "Não há um representante escolhido!", vbExclamation
        
        Exit Sub
    End If
    If msflxgrdAv.rows = 1 Then
        MsgBox "Não há um pagamento adicionado!", vbExclamation
        cmndbtnDin.SetFocus
        Exit Sub
        
    End If
    
    If txtbxVenIniAcerto.text = "" Then
        MsgBox "Informe a data para o vencimento inicial!", vbExclamation
        txtbxVenIniAcerto.SetFocus
        Exit Sub
    End If
'    If txtbxTipoCob.tag = "" Then
'        MsgBox "Escolha o tipo de cobrança!", vbExclamation
'        cmndbtnCob_Click
'    End If
    If MsgBox("Confirma lançamento de acerto para o cliente " & txtbxClienteAcerto.text & "?", vbYesNo + vbQuestion) = vbNo Then
        
        Exit Sub
    End If
    
        
    P_Status "Finalizando acerto..."
    
    lancaAcerto
    
    P_Status ""
    
End Sub
Sub calculaComissAcertoPromocional()
    Dim c As Currency
    c = CCur(lblValorFinal.caption) * 0.05
    lblAcerto_Promocao.caption = "Acerto Promocional" & Chr(13) & "Comissão: " & FormatVal(c) & "(5%)"
End Sub

Function confirma_representante() As Boolean
    Dim cod As String
    Dim rep As representante.representante
    cod = ""
    'If optnRep = True Then
        cod = InputBox("Confirme o representante")
    
        
        rep = F_BuscaRep(cod)
    
        If rep.cod_Rep <> txtbxRepresentante.tag Then
        
            MsgBox "Os representantes não coincidem!", vbExclamation
            txtbxClienteAcerto.SetFocus
            confirma_representante = False
            
        Else
            confirma_representante = True
        End If
    'Else
    '    confirma_representante = True
    'End If


End Function

Sub lancaAcerto()
    Dim strSQL As String
    Dim id_Cai As String
    Dim iCount As Integer
    Dim sSEQ_DOC As String
    Dim sSEQ_ITD As String
    Dim sSeq As String
    Dim cAux As Currency
    Dim seq_con As String
    Dim unit_ite As String
    Dim qtd_ite As String
    Dim cod_pro As String
    Dim dif_ite As String
    Dim cValorIte As Currency
    Dim cDif As Currency
    Dim sVen As String
    Dim sSeq_Com As String
    Dim sTipoCob As String
    Dim ComissaoAcertoPromo As Currency
    Dim sSeqCai As String
    
    If acerto_promocao Then
        ComissaoAcertoPromo = 5
    Else
        ComissaoAcertoPromo = 0
    End If
    
    
    
    
    
    
    
    
'    Dim sDesconto As String
    
    Dim sVal As String, sEsp As String, sNum_Che As String, sBco As String
    Dim sAge As String, sCCO As String, sNome_che As String
    Dim sCPF As String
    Dim sBol As String
    If Not F_confMsflx(msflxgrd) Then Exit Sub
    If Not F_confMsflx(msflxgrdAv) Then Exit Sub
    On Error GoTo trata_Erro
    
    
    
    
    
    If txtbxTipoCob.tag = "" Then
        If sstb.Tab = 0 Then
            MsgBox "Escolha a conta cobrança!", vbExclamation
            cmndbtnCob_Click
            Exit Sub
        Else
            txtbxTipoCob.tag = "0"
        End If
    End If
    
    'If Not F_IsCaixa Then
    '    MsgBox "Este computador não pode fazer acerto, pois não é um Computador Caixa!", vbExclamation
    '    Exit Sub
    'End If
    
    P_Status "Gravando Acerto..."
    sSeq = F_GravaAcerto(txtbxRepresentante.tag, CStr(txtbxDataBase.text), CStr(Time), _
    txtbxClienteAcerto.tag, txtbxFormaPag.tag, lblValorFinal.caption, _
    txtbxTipoCob.tag, lblCli.tag, lblComissao.caption, txtbxDifAcerto.text, txtbxEntrada.text, "2", IIf(acerto_promocao = True, "S", "N"))
            
    'grava o acerto e guarda o numero em sSEQ pra poder usar no resto
    P_Status ""
    If sSeq = "" Then
        MsgBox "Erro na gravação do duplicata!", vbExclamation
        Exit Sub
    Else
        With adoAcerto
            .MoveFirst
            Do While Not .EOF
               
                
                sSEQ_ITD = F_GravaItAcerto(sSeq, !seq_con, !cod_pro, !unit_ite, !qtd_ite, !dif_ite)
                
                
                
                If sSEQ_ITD = "" Then
                    MsgBox "Erro na gravação do item do acerto!", vbExclamation
                    strSQL = "delete from tab_acertos where seq_ace='" & sSeq & "'"
                    Conn.Execute strSQL
                    strSQL = "delete from tab_itacertos where seq_ace='" & sSeq & "'"
                    Conn.Execute strSQL
                    
                    Exit Sub
                Else
                    'If Produtos.F_AlteraEstRes(cod_pro, "-" & qtd_ite) Then
                     'já faz na gravaestcons
                    'End If
                    
                    If Not F_GravaEstCons(txtbxRepresentante.tag, !cod_pro, "-" & !qtd_ite) Then
                        MsgBox "Ocorreu um erro na volta do estoque, informe o responsável!", vbExclamation
                    End If
                    
                    strSQL = "update tab_produtos set r_pro=r_pro-" & F_GravaDec(!qtd_ite)
                    strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                    Conn.Execute strSQL
                    
                    strSQL = "update tab_consignacoes set reman_con=reman_con-" & F_GravaDec(!qtd_ite)
                    strSQL = strSQL & " where seq_con='" & !seq_con & "'"
                    strSQL = strSQL & " and cod_pro='" & !cod_pro & "'"
                    Conn.Execute strSQL
                    
                    'lembrar de fazer o cancelamento e voltar caso não de certo em algum ponto
                    
                    
                    
                End If
                .MoveNext
            Loop
        End With
    End If
    
    
    
    If sstb.Tab = 0 Then
        If optnCli.Value Then
            sSEQ_DOC = F_GravaDoc("ACERTO CONSIGN. " & txtbxClienteAcerto.text, sSeq, txtbxTotalAprazo.text, txtbxDataBase.text, _
                txtbxClienteAcerto.tag, "R", sSeq, lblParcelas.caption, "A", lblCli.tag, "002")
        Else
    
            sSEQ_DOC = F_GravaDoc("ACERTO CONSIGN. " & txtbxRepresentante.text, sSeq, txtbxTotalAprazo.text, txtbxDataBase.text, _
                txtbxRepresentante.tag, "R", sSeq, lblParcelas.caption, "A", lblCli.tag, "002")
        End If
        
        
        
        If sSEQ_DOC <> "" Then
            With msflxgrd
                For iCount = 1 To .rows - 1
                    .row = iCount
                    msflxgrd.col = 0
                    sVen = msflxgrd.text
                    msflxgrd.col = 1
                    sVal = msflxgrd.text
                    msflxgrd.col = 2
                    sEsp = Left(msflxgrd.text, 1)
                    msflxgrd.col = 3
                    sNum_Che = msflxgrd.text
                    msflxgrd.col = 4
                    sBco = msflxgrd.text
                    msflxgrd.col = 5
                    sAge = msflxgrd.text
                    msflxgrd.col = 6
                    sCCO = msflxgrd.text
                    msflxgrd.col = 7
                    sNome_che = msflxgrd.text
                    msflxgrd.col = 8
                    sCPF = msflxgrd.text
                    msflxgrd.col = 9
                    sBol = msflxgrd.text
                    If sEsp = "1" Then
                        sSEQ_ITD = doc.F_GravaItDoc(sSEQ_DOC, sVal, sVen, sVen, CStr(iCount), acerto_consignado, sBol, txtbxTipoCob.tag)
                        strSQL = F_GravaCai("C", txtbxDataBase.text, sVal, "ACERTO CONSIGN. " & txtbxRepresentante.text, sSeq, _
                        "DUP", "", 0, 0, strCod_Cai, strCod_Tur, "000", "N")
                    Else
                        
                        sSEQ_ITD = F_GravaCh(sSEQ_DOC, sNum_Che, sBco, sAge, sCCO, _
                            sNome_che, sVal, sVen, txtbxClienteAcerto.tag, lblCli.tag, _
                            "", "", CStr(iCount), "", sCPF)
                        
                        'sSEQ_ITD = doc.F_GravaItDoc(sSeq_Doc, sVal, sVen, sVen, CStr(iCount), acerto_consignado, sBol, txtbxTipoCob.tag)
                        
                        strSQL = F_GravaCai("C", txtbxDataBase.text, sVal, "ACERTO CONSIGN. " & txtbxRepresentante.text, sSeq, _
                        "CHP", "", 0, 0, strCod_Cai, strCod_Tur, "000", "N")
                        
                    End If
                    
                    If optnCli.Value Then
                        If Not F_GravaComissParc(sSEQ_ITD, txtbxRepresentante.tag, sSeq, _
                            CCur(lblComissao.caption) / (.rows - 1), sVen, sEsp) Then
                            MsgBox "Houve um problema na gravação da comissão, informe o responsável!", vbExclamation
                            
                        End If
                    Else
                        
                    End If
                    
                    
                    If sSEQ_ITD = "" Then
                        MsgBox "Houve um erro na gravação da parcela!", vbExclamation
                        strSQL = "delete from tab_acertos where seq_ace='" & sSeq & "'"
                        Conn.Execute strSQL
                        strSQL = "delete from tab_itacertos where seq_ace='" & sSeq & "'"
                        Conn.Execute strSQL
                        strSQL = "delete from tab_doc where id='" & sSEQ_DOC & "'"
                        Conn.Execute strSQL
                        strSQL = "delete from tab_itdoc where id_doc='" & sSEQ_DOC & "'"
                        Conn.Execute strSQL
                        Exit Sub
                    End If
                    
                    
                    
                Next
            End With
            Dim retorno
            P_Status "enviando documento pra internet... aguarde"
            retorno = APIInserirDoc(sSEQ_DOC, False)
            If retorno = "" Then
                insere_fila sSEQ_DOC, documento
            End If
            P_Status ""
        End If
        
        
        
        
        If optnCli.Value Then
            sSeq_Com = Acerto.F_GravaComissProv(txtbxRepresentante.tag, sSeq, lblComissao.caption, sSEQ_DOC)
             
            If sSeq_Com = "" Then
                MsgBox "Houve um erro na gravação da comissão temporária para o representante!", vbExclamation
                'Exit Sub
            End If
        End If
        
    Else
        'é pagamento a vista
        With msflxgrdAv
            For iCount = 1 To .rows - 1
                .row = iCount
                
                .col = 0
                sVen = .text
                .col = 1
                sVal = .text
                .col = 2
                sEsp = Left(.text, 1)
                .col = 3
                sNum_Che = .text
                .col = 4
                sBco = .text
                .col = 5
                sAge = .text
                .col = 6
                sCCO = .text
                .col = 7
                sNome_che = .text
                .col = 8
                sCPF = .text
                .col = 9
                sBol = .text
                
                
                
                
                If sEsp = "1" Then
                    .col = 1
                    sVal = .text
                    .col = 2
                    If Left(.text, 1) = "1" Then
                        sEsp = "DIN"
                    ElseIf Left(.text, 1) = "4" Then
                        sEsp = "PIX"
                    
                    ElseIf Left(.text, 1) = "2" Then
                        sEsp = "CHV"
                    Else
                        sEsp = "CRT"
                    End If
                    sSeqCai = F_GravaCai("C", txtbxDataBase.text, sVal, "ACERTO CONS. " & txtbxRepresentante.text, sSeq, _
                     sEsp, "", sVal, 0, strCod_Cai, strCod_Tur, "002", "N")
                     
                     
                    
                    If sSeqCai <> "" Then
                        strSQL = "insert into tab_acertofin(seq_fin,seq_ace,tipo_fin)values("
                        strSQL = strSQL & "'" & sSeqCai & "',"
                        strSQL = strSQL & "'" & sSeq & "', "
                        strSQL = strSQL & "'C')"
                        Conn.Execute strSQL
                    End If
                ElseIf sEsp = "4" Then
                    
                    
                    sSeqCai = F_GravaCai("C", txtbxDataBase.text, sVal, "PIX " & sCCO & " ACERTO CONS. " & txtbxRepresentante.text, sSeq, _
                      "PIX", "", sVal, sVal, strCod_Cai, strCod_Tur, "002", "N")
                    
                    If sSeqCai <> "" Then
                        strSQL = "insert into tab_acertofin(seq_fin,seq_ace,tipo_fin)values("
                        strSQL = strSQL & "'" & sSeqCai & "',"
                        strSQL = strSQL & "'" & sSeq & "', "
                        strSQL = strSQL & "'C')"
                        Conn.Execute strSQL
                    End If
                ElseIf sEsp = "2" Then
                    sSEQ_DOC = F_GravaDoc("ACERTO CONSIGN. " & txtbxClienteAcerto.text, sSeq, sVal, txtbxDataBase.text, _
                    txtbxRepresentante.tag, "R", sSeq, lblParcelas.caption, "A", lblCli.tag, "002")
                    strSQL = F_GravaCh(sSEQ_DOC, sNum_Che, sBco, sAge, sCCO, sNome_che, sVal, sVen, txtbxRepresentante.tag, _
                    "R", sSeq, "", "1", "", sCPF)
                    
                    
                    
                    'imprimeCarne sSeq_Doc
                    sSeqCai = F_GravaCai("C", txtbxDataBase.text, sVal, "ACERTO CONS. " & txtbxRepresentante.text, sSeq, "CHP", "", 0, 0, strCod_Cai, strCod_Tur, "002", "N")
                ElseIf sEsp = "3" Or sEsp = "4" Then
                    If sAge = "" Then sAge = "1"
                    If sEsp = "3" Then
                        sEsp = "CRT"
                    Else
                        sEsp = "CRD"
                    End If
                    strSQL = F_GravaVenFin(sSeq, sBco, CStr(sVal), txtbxDataBase.text, "", sAge, "A", CnpjTransacaoTEF)
                    id_Cai = F_GravaCai("C", txtbxDataBase.text, sVal, "ACERTO CONS. " & txtbxRepresentante.text, _
                        sSeq, sEsp, "", 0, 0, strCod_Cai, strCod_Tur, "002", "N")
                End If
            
            Next
            
        End With
        
    End If
    
    If CCur(txtbxEntrada.text) > 0 Then
            strSQL = F_GravaCai("C", txtbxDataBase.text, txtbxEntrada.text, _
                "ENTRADA ACERTO CONS " & txtbxRepresentante.text, "", "DIN", strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000", "N")
        End If
    On Error Resume Next
    
    strSQL = txtbxRepresentante.tag
    
    On Error GoTo trata_Erro
    cancela_acerto ' o acerto em aberto
    
    imprAcertoData strSQL, txtbxDataBase.text
    If sSEQ_DOC <> "" Then
        imprimeCarne sSEQ_DOC
    End If
    
    
    
    txtbxClienteAcerto.tag = ""
    txtbxClienteAcerto.text = ""
    txtbxClienteAcerto.text = ""
        Me.tag = ""
    txtbxRepresentante.Enabled = True
    cmndbtnBuscaRep.Enabled = True
'    msflxgrdAcerto.rows = 1
    msflxgrd.rows = 1
    msflxgrdAv.rows = 1
    
        
    txtbxValorNegociado.text = "0,00"
    txtbxDifAcerto.text = "0,00"
    lblValorFinal.caption = "0,00"
    lblTotalDistribuido.caption = "0,00"
    lblTotalAcerto.caption = "0,00"
    lblVlrAcertoDiv.caption = "0,00"
    txtbxValorNegociado.text = "0,00"
    txtbxDifAcerto.text = "0,00"
    lblTotalInicial.caption = "0,00"
    
    
    On Error Resume Next
    If frmAcerto.Visible = True And txtbxClienteAcerto.Enabled = True Then
        txtbxClienteAcerto.SetFocus
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnConfAcerto_Click()
'    P_Msg "aguarde..."
    Dim strSQL As String
    
    If bCpfInvalido = True Then
        MsgBox "Não é possível concluir este acerto antes de corrigir o CPF do cliente!", vbExclamation
        P_Status
        Exit Sub
    End If
    
    
    
    
    If Not confirma_representante Then
        Exit Sub
    End If
    
    
'    If Not F_SenhaRep(txtbxRepresentante.tag) Then
'
'
'        strSQL = F_senhaReturnUsu
'
'        If strSQL <> "042" Then 'laysa
'            MsgBox "Acesso não autorizado!", vbExclamation
'            Exit Sub
'        End If
'
'    End If
    
    
    If sstb.Tab = 0 Then
        
        ConfAprazo
        
    Else
        
        ConfAvista
    End If
    P_Msg
    
End Sub

Private Sub cmndbtnDin_Click()
    Dim p As parcela
    Dim strSQL As String
    p = F_DadosPar(Date)
    If p.valor > 0 Then
        strSQL = Date & Chr(9)
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        strSQL = strSQL & "1 - Dinheiro" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9)
        
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        strSQL = strSQL & Date
        msflxgrdAv.AddItem strSQL
        
        
    End If
    P_SomaItd
End Sub

Private Sub cmndbtnEditAp_Click()
    If msflxgrd.row <= 0 Then
        Exit Sub
    End If
    If Left(msflxgrd.text, 1) = "1" Then
        editaParcela
    Else
        editaCheque
    End If
End Sub

Sub deletaav()
    If msflxgrdAv.rows = 2 Then
            msflxgrdAv.rows = 1
            
        ElseIf msflxgrdAv.rows > 2 Then
            msflxgrdAv.RemoveItem msflxgrdAv.row
        End If
        
        P_SomaItd
End Sub
Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro "cmndbtnSair"
    Unload Me
End Sub











Private Sub cmndbtrnCancel_Click()
    If Me.tag = "" Then
        MsgBox "O Acerto não está gravado!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja Excluir todo o progresso do acerto, inclusive seus itens?", vbYesNo + vbCritical) = vbYes Then
        cancela_acerto
    End If
    
End Sub
Sub cancela_acerto()
    Dim strSQL As String
    On Error GoTo trata_Erro
    With adoAcerto
        .MoveFirst
        Do While Not .EOF
            deleta_item False
            .MoveNext
        Loop
    End With
    
    strSQL = "delete from tab_acertosabertos where seq_ace='" & Me.tag & "' "
    Conn.Execute strSQL
    Me.tag = ""
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub


Private Sub dtgrd_GotFocus()
    dtgrd.BorderStyle = 1
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        deleta_item
    End If
End Sub
Sub deleta_item(Optional question As Boolean = True)
    On Error GoTo trata_Erro
    Dim i As Integer
    With adoAcerto
        If question Then
            If MsgBox("Deseja realmente excluir o Item " & !nome_pro & "?", vbYesNo + vbCritical) = vbNo Then
            
            End If
        End If
        Acerto.P_DeletaItAcertoAberto !seq_ace, !seq_ite, !qtd_ite, !seq_con, !cod_pro
        i = adoAcerto.AbsolutePosition
        If question Then
            buscaItens
            adoAcerto.Move i
        End If
    End With
    Exit Sub
trata_Erro:
'    P_MostraErro
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.BorderStyle = 0
End Sub


Private Sub Form_Activate()
    On Error GoTo trata_Erro
    If strEfetuarAcerto_usu <> "S" Then
        MsgBox "Não tem permissão para efetuar acerto!", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    'lblnum_ven.caption = Me.tag
    
    
    
    txtbxDescontoAvista.text = F_BuscaUmCampo("txdescontoavista_par", "tab_parametros", "")
    
    
    If acerto_promocao Then
        txtbxDescontoAvista.text = "5"
    End If
    
    
    If bActive = True Then Exit Sub
    
    If Me.tag = "" Then
        If txtbxRepresentante.tag = "" Then
            cmndbtnBuscaRep_Click
        End If
    End If
    If Me.tag <> "" Then
        carrega
    End If
    bActive = True
    If strCod_Cai = "" Then
        MsgBox "Não existe caixa para este computador, o caixa ficará para o '01' ", vbExclamation
        strCod_Cai = "00"
    End If
    
    VerificaAcertoPromocao
    
    
    Exit Sub
trata_Erro:
    P_MostraErro "Form_Activate"
    
End Sub
Sub VerificaAcertoPromocao()
    If acerto_promocao Then
        pcProdAcerto.BackColor = &H80C0FF
        txtbxRepresentante.BackColor = &H80C0FF
        lblAcerto_Promocao.Visible = True
        sstb.TabVisible(0) = False
        
    Else
        sstb.TabVisible(0) = True
        lblAcerto_Promocao.Visible = False
        
    End If
End Sub
Sub atualiza_acerto()
    If Me.tag = "" Then
        Exit Sub
    End If
    Acerto.P_AlteraAcertoAbertos Me.tag, Date, Time, txtbxClienteAcerto.tag, _
        txtbxFormaPag.tag, lblValorFinal.caption, txtbxTipoCob.tag, IIf(optnCli.Value, "C", "R"), _
        lblComissao.caption, txtbxDifAcerto.text, txtbxEntrada.text, "1"
    
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    If Me.tag = "" Then
        Exit Sub
    End If
    
    P_Status "Carregando..."
    strSQL = "select ta.*,tr.nome_rep,tr.cpf_rep from tab_Representantes as tr, tab_acertosabertos as ta"
    strSQL = strSQL & " where ta.cod_rep=tr.cod_rep "
    strSQL = strSQL & " and ta.seq_ace='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        txtbxRepresentante.text = !nome_rep
        txtbxRepresentante.tag = !cod_Rep
        txtbxRepresentante.Enabled = False
        cmndbtnBuscaRep.Enabled = False
        If !tipo_ter = "C" Then
            optnCli.Value = True
            txtbxClienteAcerto.tag = !cod_cli
            txtbxClienteAcerto.text = F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_Cli='" & !cod_cli & "'")
        Else
            
            optnRep.Value = True
            txtbxClienteAcerto.tag = !cod_cli
            txtbxClienteAcerto.text = F_BuscaUmCampo("nome_rep", "tab_representates", "cod_rep='" & !cod_cli & "'")
        End If
        
        
        If !queima = "S" Then
            acerto_promocao = True
        Else
            acerto_promocao = False
        End If
        
        VerificaAcertoPromocao
        .Close
    End With
    buscaItens
    P_Status ""
    calculaComissAcertoPromocional
End Sub
Function F_VerificaItensJaBaixados() As Integer
    Dim strSQL As String
    Dim iCount As Integer
    With msflxgrd
        'for icount = 1 to
    End With
    
    
End Function


Sub abreDevolucao()
    
    frmDevolucaoConsign.txtbxRepresentante.text = txtbxRepresentante.text
    frmDevolucaoConsign.txtbxRepresentante.tag = txtbxRepresentante.tag
    frmDevolucaoConsign.Show 1
    
    
    
    
    
    
End Sub
Sub mudamodo()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set modo_cai='" & operacao & "' "
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
    Conn.Execute strSQL
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
'        If Me.ActiveControl.Name = dtgrd.Name Then
'            excluiItemAcerto
'        End If
    
    End If
    
    If KeyCode = vbKeyF2 Then
        txtbxProdAcerto.SetFocus
    ElseIf KeyCode = vbKeyF3 Then
    ElseIf KeyCode = vbKeyF4 Then
        txtbxClienteAcerto.SetFocus
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnConfAcerto_Click
    ElseIf KeyCode = vbKeyF8 Then
        sstb.Tab = 0
        txtbxTipoCob.SetFocus
    ElseIf KeyCode = vbKeyF9 Then
        sstb.Tab = 1
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    End If
    
End Sub

Private Sub Form_Load()
   
    bActive = False
    sstb.Tab = 0
    resizeForm Me
    ajustaGrid
    
    Set adoAcerto = New ADODB.Recordset
    
   txtbxDataBase = Date
    
    
    'P_PreencheFormasPag cmbxFormaspag, "02"
    
    
'    cmbxBuscar.ListIndex = 0
    
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
    
    
    
    ajustaGridAcerto
End Sub

Function F_QtdProdutoLancadoAcerto(cod_pro As String, seq_con As String) As Integer
    Dim i As Integer
    Dim cTotal As Integer
    Dim sSeqCon As String
    
    F_QtdProdutoLancadoAcerto = 0
    Exit Function
    With msflxgrd
        
        For i = 1 To .rows - 1
            .row = i
            .col = 9
            sSeqCon = .text
            .col = 0
            
            If .text = cod_pro And seq_con = sSeqCon Then
                .col = 3
                cTotal = cTotal + CCur(.text)
            End If
            
        Next
    End With
    F_QtdProdutoLancadoAcerto = cTotal
End Function
Sub ajustaGridAcerto()
'    With msflxgrdAcerto
'        .rows = 1
'        .Cols = 10
'        .Col = 0: .Text = "Código"
'        .Col = 1: .Text = "Referência"
'        .Col = 2: .Text = "Produto"
'        .Col = 3: .Text = "Qtd"
'        .Col = 4: .Text = "Unit. Inicial"
'        .Col = 5: .Text = "Unit Vendido"
'        .Col = 6: .Text = "Total"
'        .Col = 7: .Text = "Diferença"
'        .Col = 8: .Text = "Transferido Em:"
'        .Col = 9: .Text = "Seq.Con:"
'        .ColWidth(0) = 0
'        .ColWidth(1) = 1000
'        .ColWidth(2) = 3500
'        .ColWidth(3) = 1000
'        .ColWidth(4) = 1000
'        .ColWidth(5) = 1000
'        .ColWidth(6) = 1000
'        .ColWidth(7) = 1000
'        .ColWidth(8) = 2500
'    End With
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    
    
    
    
    
    dtgrd.height = pcBot.top - dtgrd.top - 100
    'adc.Top = pcBot.Top - 10
    dtgrd.width = pcTotalItens.Left - dtgrd.Left - 100
    frCli.width = pcCli.width
    
    
    
    
    
End Sub

Sub editaItem()
'    txtbxQtd.Enabled = True
'    txtbxQtd.Text = "1"
'    txtbxQtd.SetFocus
End Sub






Private Sub Form_Terminate()
    On Error Resume Next
    'Timer1.Enabled = False
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mdiFrm.Enabled = True
    adoAcerto.Close
    
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub frmCli_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub lblMsg2_Click()

End Sub




Private Sub msflxgrd_KeyPress(KeyAscii As Integer)
    editaCelula KeyAscii
End Sub

Private Sub msflxgrdAv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        deletaav
    End If
End Sub

Private Sub optDin_Click()
    msflxgrdAv.rows = 1
    txtbxDescontoAvista.text = F_BuscaUmCampo("txdescontoavista_par", "tab_parametros", "")
    txtbxDescCob.text = txtbxDescontoAvista.text
    
    somaDesconto
End Sub

Private Sub Option1_Click()
    msflxgrdAv.rows = 1
    txtbxDescontoAvista.text = "20"
    txtbxDescCob.text = "20"
    somaDesconto
End Sub

Private Sub Option2_Click()
    msflxgrdAv.rows = 1
    txtbxDescontoAvista.text = "15"
    txtbxDescCob.text = "15"
    somaDesconto
End Sub

Private Sub optnCli_Click()
    txtbxClienteAcerto.text = ""
    txtbxClienteAcerto.tag = ""
    lblCli.tag = "C"
    lblCli.caption = "C&liente (F4):"
    msflxgrd.rows = 1
    msflxgrdAv.rows = 1
    cmdZerar.Visible = False
End Sub

Private Sub lblDelExcluir_Click()
    excluiItemAcerto
End Sub
Sub excluiItemAcerto()
    
    deleta_item
    Exit Sub
    
    With msflxgrd
        If .col <= 0 Then
            MsgBox "Selecione um item para excluir", vbExclamation
            Exit Sub
        End If
        .col = 2
        If MsgBox("Deseja excluir o item " & .text & "?", vbYesNo + vbCritical) = vbYes Then
            If .rows = 2 Then
                .rows = 1
            Else
                .RemoveItem .row
            End If
        End If
    End With
    P_SomaAcerto
End Sub
Sub P_SomaAcerto()
    Dim cTotal As Currency
    Dim iCount As Integer
    Dim cDif As Currency
    Dim cIni As Currency
    Dim cQtd As Currency
    On Error GoTo trata_Erro:
    
    
'    With msflxgrdAcerto
'
'        For iCount = 1 To .rows - 1
'            .Row = iCount
'            .Col = 3
'            cQtd = CCur(.Text)
'            .Col = 6
'            cTotal = cTotal + CCur(.Text)
'            .Col = 7
'            cDif = cDif + CCur(.Text)
'            .Col = 4
'            cIni = cIni + (CCur(.Text) * cQtd)
'        Next
'    End With
    
    With adoAcerto
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            cQtd = cQtd + !qtd_ite
            cTotal = cTotal + !total
            cDif = cDif + !dif_ite
            cIni = cIni + (!total)
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        If .RecordCount > 1 Then
           ' dtgrd.FirstRow = adoAcerto.RecordCount - 1
        End If
    End With
    
    
    
    lblTotalAcerto.caption = FormatVal(cTotal)
    'lblTotalDifAcert0.caption = FormatVal(cDif)
    lblTotalInicial.caption = FormatVal(cIni)
    Exit Sub
    
trata_Erro:
P_MostraErro
End Sub
'Sub concluiSaidas()
'    Dim strSQL As String
'    Dim iCount As Integer
'    Dim sSeq As String
'    If dtSai.Recordset.RecordCount = 0 Then
'        MsgBox "Não existem transferências para serem concluídas!", vbExclamation
'        Exit Sub
'
'    End If
'
'    'If MsgBox("Confirma conclusão dos lançamentos para transferência?", vbYesNo + vbQuestion) = vbNo Then
'    '    Exit Sub
'    'End If
'
'    If F_senhaReturnUsu() = "" Then
'        Exit Sub
'    End If
'
'    sSeq = MaxKeySQL("tab_consignacoes", "seq_con", "", "000000#")  'buscando o sequencial q servirá para todos os produtos da transf
'    With dtSai.Recordset
'        Do While Not .EOF
'            If Consignado.F_GravaConsign(sSeq, !cod_pro, !qtd_con, !cod_Rep, !unit_con) = "" Then
'                MsgBox "Erro na transferência no item " & .AbsolutePosition, vbCritical
'
'            Else
'                If Not Consignado.F_DeletaConsignAberta(!seq_con) Then
'
'                End If
'            End If
'
'            .MoveNext
'        Loop
'
'    End With
'    buscaItens
'
'End Sub




Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'editaCelula
    End If
    If KeyCode = vbKeyDelete Then
        If msflxgrd.rows = 2 Then
            msflxgrd.rows = 1
        Else
            msflxgrd.RemoveItem msflxgrd.row
        End If
    End If
End Sub
Sub editaCelula(KeyAscii As Integer)
    Dim obj As Object
    
    
    
    
    
    
        If msflxgrd.col = 2 Then
        Set obj = cmbxOpc
    Else
        Set obj = txtbxMsf
    End If
    
    
    
    On Error Resume Next
    
    
    obj.Visible = True
    
    obj.text = msflxgrd.text
    obj.SetFocus
    obj.width = msflxgrd.CellWidth
    obj.height = msflxgrd.CellHeight
    obj.top = msflxgrd.CellTop + msflxgrd.top
    obj.Left = msflxgrd.CellLeft + msflxgrd.Left
    If KeyAscii <> 13 Then
        'If .obj.Text = "" Then
            obj.text = Chr(KeyAscii)
        'End If
    End If
    
    'obj.Alignment = msflxgrd.CellAlignment
    
End Sub

Private Sub optnPix_Click()
    msflxgrdAv.rows = 1
    txtbxDescontoAvista.text = F_BuscaUmCampo("txdescontoavista_par", "tab_parametros", "")
    txtbxDescCob.text = txtbxDescontoAvista.text
    
    somaDesconto
End Sub

Private Sub optnRep_Click()
    txtbxClienteAcerto.text = txtbxRepresentante.text
    txtbxClienteAcerto.tag = txtbxRepresentante.tag
    lblCli.tag = "R"
    lblCli.caption = "&Representante"
    lblCpfCli.caption = lblcpf.caption
    distribuiParcela
    
    bClienteNegativado = False
    msflxgrd.rows = 1
    msflxgrdAv.rows = 1
    
    cmdZerar.Visible = True
End Sub


Private Sub pcProdAcerto_GotFocus()
    Sendkeys Chr(9)
End Sub

Private Sub sstb_Click(PreviousTab As Integer)
    If sstb.Tab = 0 Then
        
        
        pcTotal.Visible = True
        msflxgrdAv.rows = 1
        txtbxDescCob.text = "0"
        txtbxTipoCob.text = ""
        
        P_SomaItd
        somaDesconto
        If modular.bUsaComissaoAcerto = True Then
            lblCapComiss.caption = "Comissão"
        Else
            lblCapComiss.caption = "Desconto"
        End If
        txtbxDataBase.Enabled = True
    ElseIf sstb.Tab = 1 Then
        txtbxDescontoAvista.text = F_BuscaUmCampo("txdescontoavista_par", "tab_parametros", "")
        
        
        
        pcTotal.Visible = False
        txtbxDataBase.text = Date
        txtbxDataBase.Enabled = False
        msflxgrd.rows = 1
        txtbxEntrada.text = "0,00"
        txtbxTipoCob.tag = "0"
        txtbxDescCob.text = txtbxDescontoAvista.text
        
        P_SomaItd
        somaDesconto
        lblCapComiss.caption = "Desconto"
    ElseIf sstb.Tab = 2 Then
    ElseIf sstb.Tab = 3 Then
        
    ElseIf sstb.Tab = 4 Then
    End If
End Sub

Private Sub txtbxBairro_GotFocus()
    'P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub stbTrab_DragDrop(Source As Control, x As Single, y As Single)
    stbTrab.Left = x + pcCli.Left
    
    stbTrab.top = y + pcCli.top
End Sub

Private Sub stbTrab_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    sY = y
    sX = x
End Sub

Private Sub txtbxClienteAcerto_Change()
    txtbxClienteAcerto.tag = ""
    
End Sub

Private Sub txtbxClienteAcerto_GotFocus()
    P_SelecTexto txtbxClienteAcerto
End Sub

Private Sub txtbxClienteAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnClienteAcerto_Click
    Else
        
    If optnRep.Value = True Then
        MsgBox "Você optou por fazer acerto para o próprio representante, não é possível escolher qualquer cliente!", vbExclamation
        optnRep_Click
        optnRep.SetFocus
        Exit Sub
    End If
        
    End If
End Sub








Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub


Private Sub txtbxDataBase_GotFocus()
    P_SelecTexto txtbxDataBase
End Sub

Private Sub txtbxDataBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataBase_LostFocus()
    On Error Resume Next
    P_AjustaData txtbxDataBase
    If txtbxDataBase.text <> "" Then
        If DateDiff("D", Date, txtbxDataBase.text) < 30 Then
            If MsgBox("Confirma lançamento da data retroativa para " & txtbxDataBase.text & "?", vbYesNo + vbCritical) = vbNo Then
                txtbxDataBase.SetFocus
                
            End If
        End If
    End If
End Sub

Private Sub txtbxDifAcerto_GotFocus()
    P_SelecTexto txtbxDifAcerto
End Sub

Private Sub txtbxDifAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDifAcerto_LostFocus()
    txtbxDifAcerto.text = Format(txtbxDifAcerto.text, "###,##0.00")
    
    If txtbxDifAcerto.text = "" Then
        txtbxDifAcerto.text = "0,00"
    End If
    
    
    calculaDif True
    
End Sub
Sub calculaDif(Optional bTotal As Boolean = True)
    On Error Resume Next
    
    If bTotal Then
        txtbxValorNegociado.text = FormatVal(CCur(txtbxDifAcerto.text) + CCur(lblTotalInicial.caption))
    Else
    
        txtbxDifAcerto.text = FormatVal(CCur(txtbxValorNegociado.text) - CCur(lblTotalInicial.caption))
    End If
    lblTotalAcerto.caption = FormatVal(CCur(lblTotalInicial.caption) + CCur(txtbxDifAcerto.text))
    
End Sub

Private Sub txtbxEntrada_GotFocus()
    P_SelecTexto txtbxEntrada
End Sub

Private Sub txtbxEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntrada_LostFocus()
    P_ValorLostFocus txtbxEntrada
    P_CalcTotalAprazo
    distribuiParcela
    somaDesconto
End Sub
Sub P_CalcTotalAprazo()
    txtbxTotalAprazo.text = FormatVal(CCur(lblValorFinal.caption) - CCur(txtbxEntrada.text))
End Sub

Private Sub txtbxFormaPag_GotFocus()
    P_SelecTexto txtbxFormaPag
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxMsf_Change()
    Dim i As Integer
    
End Sub

Private Sub txtbxMsf_GotFocus()
    'P_SelecTexto txtbxMsf
    If msflxgrd.col = 7 Then
        If txtbxMsf.text = "" Then
            txtbxMsf.text = txtbxClienteAcerto.text
        End If
    End If
    If msflxgrd.col = 8 Then
        If txtbxMsf.text = "" Then
            txtbxMsf.text = lblCpfCli.caption
        End If
    End If
    txtbxMsf.SelStart = Len(txtbxMsf.text)
End Sub

Private Sub txtbxMsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    Dim i As Integer
    Dim bCh As Boolean
    
    
    If KeyCode = vbKeyReturn Then
        
        i = msflxgrd.col
    msflxgrd.col = 2
    If Left(msflxgrd.text, 1) = "1" Then
        MsgBox "Esta linha está marcada como duplicata!", vbExclamation
    Else
        bCh = True
    End If
    msflxgrd.col = i
        
        
        If msflxgrd.col = 4 Then
            strSQL = F_SelectBanco(txtbxMsf.text).cod_bco
            msflxgrd.col = 4
            msflxgrd.text = strSQL
            msflxgrd.col = 5
            msflxgrd.SetFocus
            msflxgrd_KeyPress 13
            repete
        Else
            
            msflxgrd.text = txtbxMsf.text
            i = msflxgrd.row
            
            If msflxgrd.col > 3 Then repete
            
            msflxgrd.row = i
            msflxgrd.SetFocus
            Sendkeys Chr(9)
        End If
        
    
        
    ElseIf KeyCode = vbKeyRight And txtbxMsf.SelStart = Len(txtbxMsf) Then
        Sendkeys Chr(9)
    
    End If
    
    
    
    
    
End Sub
Sub repete()
Dim i As Integer
Dim Banco As String
    Dim Agencia As String
    Dim cpf As String
    Dim titular As String
    Dim conta As String
With msflxgrd
        .row = 1
        .col = 4
        Banco = .text
        .col = 5
        Agencia = .text
        .col = 6
        conta = .text
        .col = 7
        titular = .text
        .col = 8
        cpf = .text
        
        For i = 2 To .rows - 1
            .row = i
             .col = 4
             If .text = "" Then .text = Banco
             .col = 5
             If .text = "" Then .text = Agencia
             .col = 6
             If .text = "" Then .text = conta
             .col = 7
             If .text = "" Then .text = titular
             .col = 8
             If .text = "" Then .text = cpf
            
        Next
        End With
End Sub
Sub buscaBanco()
    Dim b As Bancos.TBanco
     
End Sub
Private Sub txtbxMsf_LostFocus()
    
    Dim iCol As Integer
    Dim iRow As Integer
    On Error Resume Next
    If msflxgrd.col = 1 Then
        P_ValorLostFocus txtbxMsf
    Else
        P_AjustaTexto txtbxMsf
    End If
    iCol = msflxgrd.col
    iRow = msflxgrd.row
    msflxgrd.text = txtbxMsf.text
    
    P_SomaItd
    
    
    msflxgrd.col = iCol + 1
    msflxgrd.row = iRow
    msflxgrd_KeyDown 13, 0
    
    txtbxMsf.Visible = False
    msflxgrd.SetFocus
End Sub

Private Sub txtbxProdAcerto_Change()
    txtbxProdAcerto.tag = ""
End Sub

Private Sub txtbxProdAcerto_DblClick()
    buscaProdAcerto
End Sub

Private Sub txtbxProdAcerto_GotFocus()
    P_SelecTexto txtbxProdAcerto
End Sub

Private Sub txtbxProdAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProdAcerto
        
    End If
End Sub
Sub ZerarEstoque()
    Dim strSQL As String
    Dim i As Integer
    Dim p As produtos.ProdutoConsig
    Dim rPesq As New ADODB.Recordset
    Dim prod As ProdutoConsig
   ' On Error GoTo trata_erro
    
    If txtbxRepresentante.tag = "" Then
        Exit Sub
    End If
    
    
    
    
        
        strSQL = "select  reman_con-reservado_con as restante,"
        strSQL = strSQL & " er_pro-r_pro as er, ts.nome_sec, tu.nome_sub,tu.comiss_sub,tu.descmax_sub, "
        
        strSQL = strSQL & "tp.*,tc.* "
        strSQL = strSQL & " from tab_produtos as tp, "
        strSQL = strSQL & " tab_subsec as tu,tab_secoes as ts,tab_consignacoes as tc  "
        strSQL = strSQL & " where tp.cod_sub=tu.cod_sub "
        strSQL = strSQL & " and tu.cod_sec=ts.cod_sec "
        strSQL = strSQL & " and tc.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tc.cod_rep='" & txtbxRepresentante.tag & "' "
        strSQL = strSQL & " and reman_con>0"
        
            
           
    
        
        'strSQL = strSQL & " group by tp.cod_pro "
        strSQL = strSQL & " order by tp.nome_pro,data_con"
    
        abreObjetoAdo rPesq, strSQL
        
        
        With rPesq
        
            If .RecordCount > 0 Then
                iniciaAcerto
            Else
                MsgBox "Não há itens para o representante!", vbExclamation
            End If
        
        
            Do While Not .EOF
                p.cod_Rep = !cod_Rep
                p.cod_usu = !cod_usu
                p.data_con = !data_con
                p.seq_con = !seq_con
                p.qtd_con = !qtd_con
                'F_BuscaProdConsign.reman_con = !reman_con
                p.reman_con = !Restante 'já tirando o reservado do que foi lançado em aberto
                
                If CCur(!Restante) = 0 And CCur(!reman_con) > 0 Then
                    MsgBox "Este produto está lançado  em acertos que estão em aberto restando apenas  " & !Restante, vbExclamation
                    
                End If
                
                p.unit_con = !unit_con
        
                p.pro.cod_pro = !cod_pro
                p.pro.fracionado_pro = !fracionado_pro
                p.pro.pa_pro = !pa_pro
                p.pro.lucro_pro = !lucro_pro
                p.pro.s_pro = !st_pro
                p.pro.cm_pro = !cm_pro
                p.pro.codbar_pro = TiraNull(!codbarra_pro)
                p.pro.cr_pro = !cr_pro
                p.pro.pq_pro = !pq_pro
                               
                
                
                
                p.pro.sub.cod_sub = !cod_sub
                p.pro.sub.nome_sub = !nome_sub
                p.pro.sub.descmax_sub = !descmax_sub
                p.pro.sub.comiss_sub = !comiss_sub
                p.pro.sub.sec.nome_sec = !nome_sec
                p.pro.pv_pro = !pv_pro
                p.pro.cr_pro = !cr_pro
                p.pro.ecf_icms = F_IcmsProd(!cod_pro, 1)
                p.pro.nome_pro = !nome_pro
                p.pro.ref_pro = !ref_pro
                p.pro.complemento_pro = !complemento_pro
                p.pro.pfab_pro = !pfab_pro
            
    
    
    
                If p.reman_con > 0 Then
                    strSQL = F_GravaItAcertoAberto(Me.tag, p.seq_con, p.pro.cod_pro, _
                     CStr(p.unit_con), CStr(p.reman_con), "0")
                     
                           
                     
                End If
                
                calculaComissAcertoPromocional
                
                .MoveNext
            Loop
            buscaItens
            
        End With
        Exit Sub
trata_Erro:
MsgBox Err.Description, vbExclamation
End Sub
Sub buscaProdAcerto()
    Dim p As produtos.ProdutoConsig
    
    If txtbxRepresentante.tag = "" Then
        cmndbtnBuscaRep_Click
    End If
    p = produtos.F_BuscaProdConsign(txtbxRepresentante.tag, txtbxProdAcerto.text)
    
    If p.reman_con > 0 Then
        txtbxProdAcerto.text = StrConv(p.pro.nome_pro, vbProperCase)
        txtbxProdAcerto.tag = p.pro.cod_pro
        lblRefAcerto.caption = p.pro.ref_pro
        
        txtbxReman.tag = p.reman_con - F_QtdProdutoLancadoAcerto(p.pro.cod_pro, p.seq_con)
        txtbxReman.text = txtbxReman.tag
        txtbxQtdAcerto.text = "1"
        If acerto_promocao = True Then
            
            If p.pro.pq_pro = 0 Then
                MsgBox "Este acerto é apenas para produtos em promoção. Escolha outra modalidade de acerto ou outro produto!", vbExclamation
                txtbxProdAcerto.tag = ""
                
                txtbxProdAcerto.SetFocus
                Exit Sub
            Else
                txtbxUnitAcerto.text = FormatVal(p.pro.pq_pro)
                txtbxUnitAcerto.tag = FormatVal(p.pro.pq_pro)
            End If
            
            
        Else
            txtbxUnitAcerto.text = FormatVal(p.unit_con)
            txtbxUnitAcerto.tag = FormatVal(p.unit_con)
        End If
        
        
        
        lblDataConAcerto.caption = p.data_con
        lblSeqConAcerto.caption = p.seq_con
        txtbxQtdAcerto.SetFocus
    Else
        MsgBox "Não há quantidade deste produto disponível para lançar. Verifique os acertos em aberto.", vbExclamation
        
                
    End If
    
    calculaComissAcertoPromocional
    
End Sub



Private Sub txtbxQtdAcerto_GotFocus()
    P_SelecTexto txtbxQtdAcerto
End Sub

Private Sub txtbxQtdAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxQtdAcerto_LostFocus
        insereItemAcerto
    End If
End Sub

Private Sub txtbxQtdAcerto_LostFocus()
    P_ValorLostFocus txtbxQtdAcerto
        P_TotalAcerto
End Sub


Private Sub txtbxRepresentante_GotFocus()
    P_SelecTexto txtbxRepresentante
End Sub

Private Sub txtbxRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaRep_Click
    Else
        txtbxRepresentante.tag = ""
    End If
End Sub

Private Sub txtbxRepresentante_LostFocus()
    P_AjustaTexto txtbxRepresentante
    If txtbxRepresentante.tag <> "" Then
        If optnRep.Value = True Then
            If txtbxRepresentante.tag <> txtbxClienteAcerto.tag Then
                txtbxClienteAcerto.text = txtbxRepresentante.text
                txtbxClienteAcerto.tag = txtbxRepresentante.tag
                
            End If
        End If
    End If
    
End Sub

Private Sub txtbxTipoCob_GotFocus()
    P_SelecTexto txtbxTipoCob
End Sub

Private Sub txtbxTipoCob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCob_Click
    End If
End Sub

Private Sub txtbxUnitAcerto_GotFocus()
    P_SelecTexto txtbxUnitAcerto
End Sub

Private Sub txtbxUnitAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxUnitAcerto_LostFocus()
    P_ValorLostFocus txtbxUnitAcerto
    P_TotalAcerto
End Sub
Sub P_TotalAcerto()
    
    'On Error Resume Next
    'If txtbxUnitAcerto.tag = "" Then
    '    txtbxUnitAcerto.tag = "0"
    'End If
    txtbxVlrFinalAcerto.text = FormatVal(CCur(txtbxUnitAcerto.text) * CCur(txtbxQtdAcerto.text))
    txtbxDiferenca.text = FormatVal(CCur(txtbxVlrFinalAcerto.text) - CCur(txtbxUnitAcerto.tag) * CCur(txtbxQtdAcerto.text))
    
End Sub

Private Sub txtbxValorNegociado_GotFocus()
    P_SelecTexto txtbxValorNegociado
End Sub

Private Sub txtbxValorNegociado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValorNegociado_LostFocus()
    P_ValorLostFocus txtbxValorNegociado
    calculaDif False
    
End Sub

Private Sub txtbxVenIniAcerto_GotFocus()
    
    P_SelecTexto txtbxVenIniAcerto
    txtbxVenIniAcerto.tag = txtbxVenIniAcerto.text
    
    
    
End Sub

Private Sub txtbxVenIniAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    If KeyCode = vbKeyRight Then
        If txtbxVenIniAcerto.text = "" Then
            txtbxVenIniAcerto.text = Date
        End If
        txtbxVenIniAcerto.text = DateAdd("d", 1, txtbxVenIniAcerto.text)
    End If
    If KeyCode = vbKeyLeft Then
    If txtbxVenIniAcerto.text = "" Then
            txtbxVenIniAcerto.text = Date
        End If
        txtbxVenIniAcerto.text = DateAdd("d", -1, txtbxVenIniAcerto.text)
    End If
End Sub

Private Sub txtbxVenIniAcerto_LostFocus()
    On Error GoTo trata_Erro
    P_AjustaData txtbxVenIniAcerto
    
    
    If DateDiff("M", txtbxVenIniAcerto.text, Date) > 6 Then
        MsgBox "Não é possível lançar carnê 6 meses anteriores!", vbExclamation
        txtbxVenIniAcerto.SetFocus
        Exit Sub
    End If
    
    
    If DateDiff("M", txtbxVenIniAcerto.text, Date) < -6 Then
        MsgBox "Não é possível lançar carnê 6 meses posteriores!", vbExclamation
        txtbxVenIniAcerto.SetFocus
        Exit Sub
    End If
    
    If txtbxVenIniAcerto.tag <> txtbxVenIniAcerto.text Or msflxgrd.rows = 1 Then
        distribuiParcela
    End If
    Exit Sub
trata_Erro:
    txtbxVenIniAcerto.text = Date
End Sub
Sub distribuiParcela()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iDias As Integer
    Dim bCheque As Boolean
    Dim iCount As Integer
    msflxgrd.rows = 1
    On Error GoTo trata_Erro
    lblVlrAcertoDiv.caption = FormatVal(CCur(txtbxTotalAprazo.text) / CCur(lblParcelas.caption))
    If txtbxVenIniAcerto.text = "" Then
        msflxgrd.rows = 1
        P_SomaItd
        Exit Sub
    End If
    If lblTipoPag.caption = "DUP" Then
        bCheque = False
    Else
        bCheque = True
    End If
    For iCount = 1 To Int(lblParcelas.caption)
            
            If strTipoPeriodoPar = "D" Then
                iDias = (30 * (iCount - IIf(1 = 1, 1, 0))) - 30
                strSQL = DateAdd("D", iDias, txtbxVenIniAcerto.text) & Chr(9)
            Else
                iDias = (1 * iCount) - 1
                strSQL = DateAdd("M", iDias, txtbxVenIniAcerto.text) & Chr(9)
            End If
            
            strSQL = strSQL & FormatVal(CCur(txtbxTotalAprazo.text) / CCur(lblParcelas.caption)) & Chr(9)
            
            If bCheque Then
                strSQL = strSQL & "2 - Cheque pré"
            Else
                strSQL = strSQL & "1 - Duplicata"
            End If
            msflxgrd.AddItem strSQL
            msflxgrd.row = msflxgrd.rows - 1
            msflxgrd.col = 1
            'msflxgrd.Text = txtbxBoleto.Text
            cTotal = cTotal + Round((CCur(txtbxTotalAprazo.text) / CCur(lblParcelas.caption)), 2)
            
        Next
            P_SomaItd
            If CCur(lblTotalDistribuido.caption) <> CCur(lblValorFinal.caption) Then
                msflxgrd.row = msflxgrd.rows - 1
                msflxgrd.col = 1
                msflxgrd.text = FormatVal(CCur(msflxgrd.text) + CCur(txtbxTotalAprazo.text) - CCur(lblTotalDistribuido.caption))
            End If
            P_SomaItd
            
                    
            Exit Sub
trata_Erro:
          P_MostraErro
            
End Sub
Sub ajustaGrid()
    AjustaGr msflxgrd
    AjustaGr msflxgrdAv
End Sub
Sub AjustaGr(obj As Object)
    Dim bCheque As Boolean
    bCheque = False
    txtbxVenIniAcerto.text = Date
    
    'P_CentraForm Me, 600
    obj.rows = 1
    obj.Cols = 11
    obj.col = 0
    obj.text = "Vencimento"
    obj.col = 1
    obj.text = "Valor"
    obj.col = 2
    obj.text = "Espécie"
    obj.col = 3
    obj.text = "Cheque Nº"
    obj.col = 4
    obj.text = "Ban/Con"
    obj.col = 5
    obj.text = "Agencia"
    obj.col = 6
    obj.text = "Conta C."
    obj.col = 7
    obj.text = "Titular"
    obj.col = 8
    obj.text = "CPF"
    obj.col = 9
    obj.text = "Boleto"
    obj.col = 10
    obj.text = "Comissão"
    
    obj.ColAlignment(2) = 1
    
    obj.ColWidth(0) = 1000
    
    obj.ColWidth(1) = 1000
    obj.ColWidth(2) = 1500
    obj.ColWidth(7) = 3000
    obj.ColWidth(8) = 1500
End Sub
Private Sub txtbxVlrFinalAcerto_GotFocus()
    P_SelecTexto txtbxVlrFinalAcerto
End Sub

Private Sub txtbxVlrFinalAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        insereItemAcerto
    End If
End Sub
Function ExisteAcertoEmAberto() As Boolean

    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_acertosabertos where cod_rep='" & txtbxRepresentante.tag & "' "
    abreObjetoAdo rPesq, strSQL
    strSQL = ""
    With rPesq
        Do While Not .EOF
            strSQL = strSQL & "IP: " & !ipusu_ace & "-> Usuário: " & !cod_usu & Chr(13)
                
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            MsgBox "Existem " & .RecordCount & " Acerto(s) em aberto para o representante!" & _
                Chr(13) & strSQL, vbExclamation
                                
            ExisteAcertoEmAberto = True
                
        End If
        
        .Close
    End With

End Function
Sub iniciaAcerto()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_acertosabertos where cod_rep='" & txtbxRepresentante.tag & "' "
    abreObjetoAdo rPesq, strSQL
    strSQL = ""
    With rPesq
        Do While Not .EOF
            strSQL = strSQL & "IP: " & !ipusu_ace & "-> Usuário: " & !cod_usu & Chr(13)
            
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            If MsgBox("Existem " & .RecordCount & " Acerto(s) em aberto para o representante!" & _
                Chr(13) & strSQL & Chr(13) & "Deseja Continuar mesmo assim?", vbYesNo + vbQuestion) = vbNo Then
                txtbxRepresentante.SetFocus
                Exit Sub
            End If
                
        End If
        
        .Close
    End With
    If MsgBox("Confirma abertura do acerto para o Representante " & txtbxRepresentante.tag & "-" & txtbxRepresentante.text & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    Me.tag = Acerto.F_GravaAcertoAbertos(txtbxRepresentante.tag, Date, Time, txtbxClienteAcerto.tag, _
        txtbxFormaPag.tag, lblValorFinal.caption, txtbxTipoCob.tag, IIf(optnCli.Value, "C", "R"), _
        lblComissao.caption, txtbxDifAcerto.text, txtbxEntrada.text, "1", IIf(acerto_promocao = True, "S", "N"))
        
    
    If Me.tag = "" Then
        MsgBox "Erro na abertura do acerto!", vbExclamation
    Else
        txtbxRepresentante.Enabled = False
    End If
    
    
    
    
End Sub



Sub insereItemAcerto()
    Dim strSQL As String
    txtbxVlrFinalAcerto_LostFocus
    If CCur(txtbxVlrFinalAcerto.text) = 0 Then
        Exit Sub
    End If
    
    If Me.tag = "" Then
        iniciaAcerto
    End If
    If CCur(txtbxQtdAcerto.text) > CCur(txtbxReman.text) Then
        MsgBox "A quantidade que está lançando é maior que a restante!", vbExclamation
        txtbxQtdAcerto.text = txtbxReman.text
        txtbxQtdAcerto.SetFocus
        Exit Sub
    End If

    If txtbxProdAcerto.tag = "" Then
        MsgBox "Escolha o produto e tecle enter", vbExclamation
        txtbxProdAcerto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxVlrFinalAcerto.text) = 0 Then
        MsgBox "O Valor final não pode ser zero!", vbExclamation
        Exit Sub
    End If
    strSQL = txtbxProdAcerto.tag & Chr(9)
    strSQL = strSQL & lblRefAcerto.caption & Chr(9)
    strSQL = strSQL & txtbxProdAcerto.text & Chr(9)
    strSQL = strSQL & txtbxQtdAcerto.text & Chr(9)
    strSQL = strSQL & txtbxUnitAcerto.tag & Chr(9)
    strSQL = strSQL & txtbxUnitAcerto.text & Chr(9)
    strSQL = strSQL & txtbxVlrFinalAcerto.text & Chr(9)
    strSQL = strSQL & txtbxDiferenca.text & Chr(9)
    strSQL = strSQL & lblDataConAcerto.caption & Chr(9)
    strSQL = strSQL & lblSeqConAcerto.caption
    'msflxgrdAcerto.AddItem strSQL
    
    
    strSQL = F_GravaItAcertoAberto(Me.tag, lblSeqConAcerto.caption, txtbxProdAcerto.tag, _
        txtbxUnitAcerto.text, txtbxQtdAcerto.text, txtbxDiferenca.text)
    
    txtbxUnitAcerto.text = "0"
    txtbxProdAcerto.tag = ""
    txtbxVlrFinalAcerto.text = "0"
    txtbxQtdAcerto.text = "1"
    txtbxDiferenca.text = "0"
    txtbxProdAcerto.SetFocus
    txtbxReman.tag = "0"
    txtbxReman.text = "0"
    
    txtbxUnitAcerto.tag = "0"
    buscaItens
    
'    msflxgrdAcerto.TopRow = msflxgrdAcerto.rows - 1
End Sub
Sub buscaItens()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    strSQL = "select tp.nome_pro,tp.ref_pro,tp.pv_pro,(ti.qtd_ite*ti.unit_ite) as total,ti.*,tc.data_con"
    strSQL = strSQL & " from tab_acertosabertos as ta,tab_produtos as tp,"
    strSQL = strSQL & " tab_itacertosabertos as ti,tab_consignacoes as tc "
    strSQL = strSQL & " where ta.seq_ace=ti.seq_ace "
    strSQL = strSQL & " and ti.seq_con=tc.seq_con"
    strSQL = strSQL & " and ti.cod_pro=tp.cod_pro"
    strSQL = strSQL & " and tc.cod_pro=tp.cod_pro"
    strSQL = strSQL & " and ta.seq_ace='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite desc "
    
    FechaObjeto adoAcerto
    abreObjetoAdo adoAcerto, strSQL, dtgrd
    
    
    
    soma
    
'    'rpesq.open strsql,conn
'    With rPesq
'        Do While Not .EOF
'            strSQL = !cod_pro & Chr(9)
'            strSQL = strSQL & !ref_pro & Chr(9)
'            strSQL = strSQL & !nome_pro & Chr(9)
'            strSQL = strSQL & !qtd_ite & Chr(9)
'            strSQL = strSQL & FormatVal(!pv_pro) & Chr(9)
'            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
'            strSQL = strSQL & FormatVal(!unit_ite * !qtd_ite) & Chr(9)
'            strSQL = strSQL & FormatVal(!dif_ite) & Chr(9)
'            strSQL = strSQL & !data_con & Chr(9)
'            strSQL = strSQL & !seq_con
'            'msflxgrdAcerto.AddItem strSQL
'            .MoveNext
'        Loop
'        .Close
'    End With
    
    
End Sub
Sub soma()
    P_SomaAcerto
    P_SomaItd
    somaDesconto
End Sub
Private Sub txtbxVlrFinalAcerto_LostFocus()
    On Error Resume Next
    P_ValorLostFocus txtbxVlrFinalAcerto
    txtbxUnitAcerto.text = FormatVal(CCur(txtbxVlrFinalAcerto.text) / CCur(txtbxQtdAcerto.text))
    P_TotalAcerto
End Sub


Sub P_SomaItd()
    Dim cVal As Currency
    Dim iCount As Integer
    Dim iRow As Integer
    Dim cIni As Currency
    Dim obj As Object
    
    If sstb.Tab = 0 Then
        Set obj = msflxgrd
    Else
        Set obj = msflxgrdAv
    End If
    
    iRow = obj.row
    With obj
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 1
            cVal = cVal + CCur(.text)
            
        Next
        
    End With
    
    obj.row = iRow
    lblTotalDistribuido.caption = FormatVal(cVal + CCur(txtbxEntrada.text))
End Sub

Function F_confMsflx(msf As MSFlexGrid) As Boolean
     Dim iCount As Integer
     Dim sTipo As String
     Dim retorno As Boolean
     Dim strSQL As String
     
     retorno = True
     strSQL = "os seguintes dados estão faltando: " & Chr(13) & Chr(13)
     With msf
        For iCount = 1 To .rows - 1
            .row = iCount
            If 1 = 0 Then
                .col = 2
                sTipo = Left(.text, 1)
                If sTipo = "2" Then
                    .col = 3
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        retorno = False
                    End If
                    .col = 4
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        retorno = False
                    End If
                    
                    
                    
                    
                    
                End If
            
            Else
                .col = 2
                sTipo = Left(.text, 1)
                If sTipo = "2" Then
                    .col = 3
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        retorno = False
                    End If
                    .col = 4
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        retorno = False
                    End If
                    
                    If F_BuscaUmCampo("Nome_bco", "tab_bancos", "cod_bco='" & .text & "'") = "" Then
                        strSQL = strSQL & "linha " & iCount & " não existe o banco de código " & .text & Chr(13)
                        retorno = False
                    End If
                    .col = 5
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & " agência" & Chr(13)
                        retorno = False
                    End If
                    .col = 6
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & " conta" & Chr(13)
                        retorno = False
                    End If
                    .col = 7
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & " nome do titular" & Chr(13)
                        retorno = False
                    End If
                    .col = 8
                    If .text = "" Then
                        strSQL = strSQL & "linha " & iCount & " CPF do titular" & Chr(13)
                        retorno = False
                    End If
                End If
            End If
        Next
     End With
     F_confMsflx = retorno
     If retorno = False Then
        MsgBox strSQL, vbExclamation
        
     End If
End Function

