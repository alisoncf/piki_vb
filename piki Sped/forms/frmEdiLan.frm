VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiLan 
   Caption         =   "Edição de Lançamentos"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11850
   Icon            =   "frmEdiLan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmndbtnBaixar 
      Caption         =   "&Baixar"
      Height          =   720
      Left            =   240
      Picture         =   "frmEdiLan.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7800
      Width           =   1455
   End
   Begin TabDlg.SSTab sstb 
      Height          =   3735
      Left            =   240
      TabIndex        =   43
      Top             =   3840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Prorrogações"
      TabPicture(0)   =   "frmEdiLan.frx":0454
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmndbtnProrr"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtbxVencAtual"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtbxAnot"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtbxProrrogar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtbxVisAnot"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "gridProrrog"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Memória das Cobranças"
      TabPicture(1)   =   "frmEdiLan.frx":0470
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridUsu"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "gridCob"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Outras Funções"
      TabPicture(2)   =   "frmEdiLan.frx":048C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSDataGridLib.DataGrid gridCob 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   60
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4683
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridProrrog 
         Height          =   1935
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtbxVisAnot 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000006&
         Height          =   1725
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   720
         Width           =   6495
      End
      Begin VB.TextBox txtbxProrrogar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtbxAnot 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   2640
         Width           =   8655
      End
      Begin VB.TextBox txtbxVencAtual 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnProrr 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   10080
         TabIndex        =   52
         Top             =   2640
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid gridUsu 
         Height          =   2655
         Left            =   -68520
         TabIndex        =   61
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4683
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Anotação Referente ao Dia"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   55
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Prorrogar Para:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Anotações do Contato:"
         Height          =   195
         Left            =   2880
         TabIndex        =   48
         Top             =   2400
         Width           =   1665
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2400
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EF9C7C&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11820
      TabIndex        =   14
      Top             =   0
      Width           =   11850
      Begin VB.Label lblTipoTer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   28
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   27
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   435
      End
      Begin VB.Label lblSeq_Doc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   25
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   24
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblNUM_Doc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   23
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5880
         TabIndex        =   22
         Top             =   0
         Width           =   585
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   21
         Top             =   0
         Width           =   945
      End
      Begin VB.Label lblSeq_Itd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seq:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblParc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   16
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5880
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11655
      Begin VB.OptionButton optnUsu 
         Caption         =   "Usuário (Func)"
         Height          =   195
         Left            =   3960
         TabIndex        =   58
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optnRepr 
         Caption         =   "Representante"
         Height          =   195
         Left            =   2520
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optnCli 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   1560
         TabIndex        =   56
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtbxClass 
         Appearance      =   0  'Flat
         BackColor       =   &H0005C9EB&
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Tag             =   "000"
         Text            =   "NÃO CLASSIFICADO"
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Frame frCobr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   5760
         TabIndex        =   29
         Top             =   960
         Width           =   5535
         Begin VB.TextBox txtbxFunc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3240
            TabIndex        =   49
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox chkTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Todas as Parcelas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txtbxTipoCob 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtbxContaCob 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   2775
         End
         Begin VB.CommandButton cmndbtnCob 
            Height          =   285
            Left            =   2880
            Picture         =   "frmEdiLan.frx":04A8
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnContaCob 
            Height          =   285
            Left            =   2880
            Picture         =   "frmEdiLan.frx":0832
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnFun 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5040
            Picture         =   "frmEdiLan.frx":0BBC
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Para o Funcionário:"
            Height          =   195
            Left            =   3240
            TabIndex        =   53
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label lblContaCob 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta Cobrança:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblTipoCob 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Cobrança:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Width           =   1125
         End
      End
      Begin VB.TextBox txtbxBoleto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtbxNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txtbxHist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtbxVenc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtbxValor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnPlanoContas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         Picture         =   "frmEdiLan.frx":0F46
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblTel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   42
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   41
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Classificação:"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Boleto:"
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblTerCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Terceiro:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Histórico:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Vencimento:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "V&alor:"
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   1440
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   8400
      TabIndex        =   11
      Top             =   7560
      Width           =   3255
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   735
         Left            =   1680
         Picture         =   "frmEdiLan.frx":12D0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Confirmar"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiLan.frx":143B
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito"
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
         Left            =   1320
         TabIndex        =   39
         Top             =   1080
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmEdiLan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Public WithEvents adoCob As ADODB.Recordset
Attribute adoCob.VB_VarHelpID = -1
Public WithEvents adoUsu As ADODB.Recordset
Attribute adoUsu.VB_VarHelpID = -1
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo trata_Erro
    txtbxVisAnot.Text = F_BuscaUmCampo("anotacao_pre", " " & TAB_ITDOC & "pre", "seq_pre=" & ado!seq_pre)
    Exit Sub
trata_Erro:
    txtbxVisAnot.Text = ""
End Sub

Private Sub cmndbtnBaixar_Click()
    On Error GoTo trata_Erro
    If strBaixadoc_usu <> "S" Then
        MsgBox "Não tem permissão pra baixar documento!", vbExclamation
        Exit Sub
    End If
    With adorMov
        frmBaixaDocReq.tag = txtbxNome.tag
        frmBaixaDocReq.txtbxCliente.Text = txtbxNome.Text
        
        frmBaixaDocReq.txtbxCliente.tag = txtbxNome.tag
        frmBaixaDocReq.lblCpf.caption = ""
        frmBaixaDocReq.txtbxSeq.Text = lblSeq_Doc.caption
        frmBaixaDocReq.Show 1
        
        Unload frmBaixaDocReq
    End With
    Exit Sub
trata_Erro:
P_MostraErro

End Sub

Private Sub cmndbtnCob_Click()
    Dim t As tipocob
    
    t = Cobranca.F_BuscaTipoCob("")
    If t.cod_tip <> "" Then
        txtbxTipoCob.Text = t.cod_tip & " " & t.nome_tip
        txtbxTipoCob.tag = t.cod_tip
        If t.cod_tip = "1" Then
            'tem q voltar pra geral
            txtbxContaCob.Text = "GERAL"
            txtbxContaCob.tag = "000"
            
        End If
    End If
End Sub

Private Sub cmndbtnContaCob_Click()
    Dim t As ContaCob
    
    t = Cobranca.F_BuscaContaCob("")
    If t.cod_cob <> "" Then
        txtbxContaCob.Text = t.nome_cob
        txtbxContaCob.tag = t.cod_cob
        If t.cod_cob = "099" Then
            txtbxFunc.Enabled = True
            cmndbtnFun.Enabled = True
        Else
            txtbxFunc.Enabled = False
            cmndbtnFun.Enabled = False
        End If
    End If
End Sub

Private Sub cmndbtnFun_Click()
    Dim f As funcionario.funcionario
    f = funcionario.F_SelectFuncionario(txtbxFunc.Text)
    If f.cod_fun <> "" Then
        txtbxFunc.Text = f.guerra_fun
        txtbxFunc.tag = f.cod_fun
        
        
    End If
    
End Sub

Private Sub cmndbtnPlanoContas_Click()
        Dim p As PlanoContas
    
    p = F_BuscaPlanoContas(Left(lblTipo.caption, 1), txtbxClass.Text)
    If p.cod_pla <> "" Then
        txtbxClass.tag = p.cod_pla
        txtbxClass.Text = p.nome_pla
        'txtbxParc.SetFocus
    End If
    
End Sub

Private Sub cmndbtnProrr_Click()
    If strProrrogarPar_Usu <> "S" Then
        MsgBox "Não tem permissão para prorrogar parcela!", vbExclamation
        Exit Sub
    End If
    
    
    If txtbxProrrogar.Text = "" Then
        MsgBox "Informe a data que deseja prorrogar!", vbExclamation
        txtbxProrrogar.SetFocus
        Exit Sub
    End If
    P_GravaItDocPre Me.tag, txtbxVenc.Text, txtbxProrrogar.Text, txtbxAnot.Text
    'If MsgBox("Deseja alterar as parcelas anteriores para este vencimento?", vbYesNo + vbQuestion) = vbYes Then
    '    P_AlteraDocsAnteriores lblSeq_Doc.caption, txtbxVenc.Text, txtbxProrrogar.Text
    'End If
    
    carrega
    
    
End Sub
Sub buscaPre()
    Dim strSQL As String
    strSQL = "select ti.*,tf.guerra_fun from " & TAB_ITDOC & "pre as ti,tab_funcionarios as tf "
    strSQL = strSQL & " where ti.cod_usu=tf.cod_fun "
    strSQL = strSQL & " and seq_itd='" & Me.tag & "' "
    abreObjetoAdo ado, strSQL, gridProrrog
    
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
  '  If Not F_senha Then
  '      Exit Sub
  '  End If
    
    If txtbxValor.Text = "0,00" Then
        MsgBox "Este valor é inválido!", vbCritical
        txtbxValor.SetFocus
        Exit Sub
    End If
    If txtbxHist.Text = "" Then
        MsgBox "Informe o Histórico!", vbCritical
        txtbxHist.SetFocus
        Exit Sub
        
    End If
    If txtbxNome.tag = "" Then
        MsgBox "Dê enter no campo 'NOME' para escolher um terceiro!", vbExclamation
        txtbxNome.SetFocus
        Exit Sub
        
    End If
    If MsgBox("Confirma alteração da Parcela?", vbYesNo + vbQuestion) = vbYes Then
        altera
    End If
    
End Sub
Sub altera()
    Dim strSQL As String
    Dim i As Integer
    On Error GoTo trata_Erro
    
    If lblTipoTer.caption = "" Then
        lblTipoTer.caption = "Cliente"
    End If
    
    
    strSQL = "update  " & TAB_DOC & " set historico_doc='" & txtbxHist.Text & "', "
    strSQL = strSQL & "cod_pla='" & txtbxClass.tag & "', "
    strSQL = strSQL & "cod_Ter='" & txtbxNome.tag & "', "
    
    strSQL = strSQL & "tipo_ter='" & Left(lblTipoTer.caption, 1) & "'    "
    strSQL = strSQL & " where seq_doc='" & lblSeq_Doc.caption & "' "
    conn.Execute strSQL
    
    strSQL = "update  " & TAB_ITDOC & "  set " 'venc_itd=" & F_GravaData(txtbxVenc.Text) & ","
    strSQL = strSQL & "reman_itd=" & F_GravaDec(txtbxValor.Text) & ","
    
    If txtbxTipoCob.tag <> "" Then
        strSQL = strSQL & "cod_tip='" & txtbxTipoCob.tag & "',"
    End If
    If txtbxContaCob.tag <> "" Then
        
        strSQL = strSQL & "cod_cob='" & txtbxContaCob.tag & "',"
        If lblContaCob.tag <> txtbxContaCob.tag Then
            doc.P_gravaDocHist lblSeq_Doc.caption, lblContaCob.tag, txtbxContaCob.tag, txtbxTipoCob.tag, lblTipoCob.tag
            'grava o histórico da mudança da cobrança
            
        End If
        If txtbxFunc.Enabled = True And txtbxFunc.tag <> "" Then
            doc.P_gravaDocHistUsu lblSeq_Doc.caption, txtbxFunc.tag
        End If
    End If
    
    strSQL = strSQL & "boleto_itd='" & txtbxBoleto.Text & "'"
    strSQL = strSQL & " where seq_itd='" & lblSeq_Itd & "'"
    conn.Execute strSQL, i
    
    If chkTodas.Value = 1 Then
        strSQL = "update  " & TAB_ITDOC & "  set cod_cob='" & txtbxContaCob.tag & "' "
        If txtbxTipoCob.tag <> "" Then
            strSQL = strSQL & ",cod_tip='" & txtbxTipoCob.tag & "' "
        End If
        
        strSQL = strSQL & " where seq_doc='" & lblSeq_Doc.caption & "' "
        
        strSQL = strSQL & " and reman_itd> 0"
        conn.Execute strSQL
    End If
    
    MsgBox "Documento Alterado!", vbInformation
    Unload Me
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub dt_Reposition()
   
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    
    carrega
    If Left(lblTipo.caption, 1) = "D" Then
        frCobr.Visible = False
    End If
    
    If strAlterarVlrParc_Usu <> "S" Then
        
        txtbxValor.Enabled = False
        
    End If
    If strAlterarTerceiroDoc_Usu <> "S" Then
        txtbxNome.Enabled = False
    End If
    bActive = True
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    Dim sTipoDoc As String
    sTipoDoc = "R"
    
    On Error GoTo trata_Erro
    
    lblTipoTer.caption = F_BuscaUmCampo("tipo_ter", TAB_DOC & " as td, " & TAB_ITDOC & "  as ti", "td.seq_doc=ti.seq_doc and seq_itd='" & Me.tag & "'")
    
    If lblTipoTer.caption = "C" Or lblTipoTer.caption = "" Then
            sTab = TAB_CLIENTES
            sCod = "cod_cli"
            sNome = "nome_cli"
            sTipoTer = "C"
        ElseIf lblTipoTer.caption = "U" Then
            sTab = TAB_FUNCIONARIOS
            sCod = "cod_fun"
            sNome = "nome_fun"
            sTipoTer = "U"
            optnUsu.Value = True
            optnCli.Enabled = False
            optnRepr.Enabled = False
        
        ElseIf lblTipoTer.caption = "F" Then
            sTab = TAB_FORNECEDORES
            sCod = "cod_for"
            sNome = "nome_for"
            sTipoTer = "F"
            sTipoDoc = "P"
        Else
        
            sTab = "tab_representantes"
            sNome = "nome_rep"
            sCod = "cod_rep"
            sTipoTer = "R"
        End If
    strSQL = "select tpc.nome_pla,td.tipo_ter,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
    strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
    strSQL = strSQL & " tcc.nome_cob,td.cod_ter,"
    strSQL = strSQL & "td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
    strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
    strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
    
    strSQL = strSQL & "from " & TAB_DOC & " as td inner join " & sTab & " as tc on tc." & sCod & "=td.cod_ter "
    strSQL = strSQL & " left join " & TAB_ITDOC & "  as ti on td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " left join tab_contacob as tcc on tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " left join tab_tipocob as tt on tt.cod_tip=ti.cod_tip  "
    strSQL = strSQL & " left join " & TAB_PLANOCONTAS & " as tpc on tpc.cod_pla=td.cod_pla , "
    strSQL = strSQL & " tab_parametros as tp  "
    strSQL = strSQL & "where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & "and td.tipo_doc='" & sTipoDoc & "' "
    strSQL = strSQL & " and td.tipo_ter='" & sTipoTer & "' "

    strSQL = strSQL & " and ti.seq_itd='" & Me.tag & "' "
        
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If !tipo_ter = "C" Then
            optnCli.Value = True
        ElseIf !tipo_ter = "U" Then
            optnUsu.Value = True
        
        Else
        
            optnRepr.Value = True
        End If
        lblData.caption = !data_doc
        lblNUM_Doc.caption = !num_doc
        lblParc.caption = !parcela_itd & "/" & !parcelas_doc
        lblSeq_Doc.caption = !seq_doc
        lblSeq_Itd.caption = !seq_itd
        lblValor.caption = FormatVal(!valor_itd)
        txtbxHist.Text = !historico_doc
        txtbxNome.Text = !nome
        txtbxNome.tag = !codigo
        txtbxVenc.Text = !venc_itd
        txtbxProrrogar.Text = !venc_itd
        txtbxValor.Text = FormatVal(!reman_itd)
        txtbxBoleto.Text = !boleto_itd
        txtbxClass.tag = TiraNull(!cod_pla)
        txtbxClass.Text = TiraNull(!nome_pla)
        txtbxContaCob.Text = TiraNull(!nome_cob)
        lblContaCob.tag = TiraNull(!cod_cob)
        'txtbxFunc.tag = !cod_fun
        'txtbxFunc.Text = !nome_fun
        txtbxContaCob.tag = !cod_cob
        txtbxTipoCob.Text = !cod_tip & " " & !nome_tip
        txtbxTipoCob.tag = !cod_tip
        lblTipoCob.tag = !cod_tip
        lblTipo.caption = "Crédito"
        txtbxVencAtual.Text = !venc_itd
        txtbxBoleto.Text = !boleto_itd
        
        'lblTel.caption = F_BuscaTel(!cod_ter, sTipoTer)
        .Close
    End With
    buscaPre
    buscaMemoria
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Sub buscaMemoria()
    Dim strSQL As String
    strSQL = doc.F_SelectDocHist(lblSeq_Doc.caption)
    
    abreObjetoAdo adoCob, strSQL, gridCob
    strSQL = doc.F_SelectDocHistUsu(lblSeq_Doc.caption)
    abreObjetoAdo adoUsu, strSQL, gridUsu
    
End Sub
Private Sub Form_Load()
        'adc.ConnectionString = Conn.ConnectionString
    Set ado = New ADODB.Recordset
    Set adoUsu = New ADODB.Recordset
    Set adoCob = New ADODB.Recordset
    
    sstb.Tab = 0
    bActive = False
End Sub

Private Sub txtbxAnot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAnot_LostFocus()
    P_AjustaTexto txtbxAnot, 3
End Sub

Private Sub txtbxBoleto_GotFocus()
    P_SelecTexto txtbxBoleto
End Sub

Private Sub txtbxBoleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBoleto_LostFocus()
    P_AjustaTexto txtbxBoleto
End Sub

Private Sub txtbxHist_GotFocus()
    P_SelecTexto txtbxHist
End Sub

Private Sub txtbxHist_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxHist_LostFocus()
    P_AjustaTexto txtbxHist
End Sub

Private Sub txtbxNome_GotFocus()
    P_SelecTexto txtbxNome
End Sub
Sub buscaTer()
    If optnCli.Value Then
        buscaCli
    ElseIf optnRepr.Value Then
        buscaRep
    ElseIf optnUsu.Value Then
        buscaFun
    End If
End Sub

Sub buscaCli()
    Dim c As Clientes.cliente
    c = Clientes.F_BuscaCliente(txtbxNome.Text)
    If c.cod_cli <> "" Then
        strSQL = "Mudar o Terceiro de uma parcela alterará"
        strSQL = strSQL & " o terceiro de todas as parcelas" & Chr(13)
        strSQL = strSQL & "referentes a este Documento!" & Chr(13)
        strSQL = strSQL & "Deseja Continunar?"
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
            txtbxNome.Text = c.nome_Cli
            txtbxNome.tag = c.cod_cli
            lblTipoTer.caption = "C"
            
        End If
    End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
    r = representante.F_BuscaRep(txtbxNome.Text)
    If r.cod_Rep <> "" Then
        strSQL = "Mudar o Terceiro de uma parcela alterará"
        strSQL = strSQL & " o terceiro de todas as parcelas" & Chr(13)
        strSQL = strSQL & "referentes a este Documento!" & Chr(13)
        strSQL = strSQL & "Deseja Continunar?"
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
            txtbxNome.Text = r.nome_rep
            txtbxNome.tag = r.cod_Rep
            lblTipoTer.caption = "R"
            
        End If
    End If
End Sub
Sub buscaFun()
    Dim r As funcionario.funcionario
    r = funcionario.F_SelectFuncionario(txtbxNome.Text)
    If r.cod_fun <> "" Then
        strSQL = "Mudar o Terceiro de uma parcela alterará"
        strSQL = strSQL & " o terceiro de todas as parcelas" & Chr(13)
        strSQL = strSQL & "referentes a este Documento!" & Chr(13)
        strSQL = strSQL & "Deseja Continunar?"
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
            txtbxNome.Text = r.nome_fun
            txtbxNome.tag = r.cod_fun
            lblTipoTer.caption = "U"
            
        End If
    End If
End Sub


Private Sub txtbxNome_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    
    If KeyCode = vbKeyReturn Then
        If strAlterarTerceiroDoc_Usu <> "S" Then
            MsgBox "Não tem permissão para alterar terceiro da parcela!", vbExclamation
            Exit Sub
        End If
                 
        buscaTer
            
            
    Exit Sub
            
            
        
        frmSelectCli.txtbxNome.tag = "C"
        frmSelectCli.txtbxNome.Text = txtbxNome.Text
    
        On Error Resume Next
    
        frmSelectCli.Show 1
        If frmSelectCli.adoR!cod <> "" Then
            strSQL = "Mudar o Terceiro de uma parcela alterará"
            strSQL = strSQL & " o terceiro de todas as parcelas" & Chr(13)
            strSQL = strSQL & "referentes a este Documento!" & Chr(13)
            strSQL = strSQL & "Deseja Continunar?"
            If MsgBox(strSQL, vbYesNo + vbCritical) = vbYes Then
                txtbxNome.Text = frmSelectCli.adoR!nome
                txtbxNome.tag = frmSelectCli.adoR!cod
                'lblTerCap.Tag =frmSelectCli.ador!tipo
                
            End If
            
            
        End If
        Unload frmSelectCli

    End If
    
End Sub

Private Sub txtbxNome_LostFocus()
    P_AjustaTexto txtbxNome
End Sub

Private Sub txtbxProrrogar_GotFocus()
    P_SelecTexto txtbxProrrogar
End Sub

Private Sub txtbxProrrogar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxAnot.SetFocus
    End If
End Sub

Private Sub txtbxProrrogar_LostFocus()
    P_AjustaData txtbxProrrogar
End Sub

Private Sub txtbxValor_GotFocus()
    P_SelecTexto txtbxValor
    txtbxValor.tag = txtbxValor.Text
End Sub

Private Sub txtbxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_LostFocus()
    P_ValorLostFocus txtbxValor
    If txtbxValor.Text <> "" Then
        If txtbxValor.Text <> txtbxValor.tag Then
            If strAlterarVlrParc_Usu <> "S" Then
                MsgBox "Você não tem permissão para alterar o valor da parcela!", vbExclamation
                txtbxValor.Text = txtbxValor.tag
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtbxVenc_GotFocus()
    P_SelecTexto txtbxVenc
End Sub

Private Sub txtbxVenc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxVenc_LostFocus()
    P_AjustaData txtbxVenc
End Sub
