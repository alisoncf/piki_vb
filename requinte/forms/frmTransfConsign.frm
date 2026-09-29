VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTransfConsign 
   Caption         =   "Transferência:"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16050
   Icon            =   "frmTransfConsign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   16050
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLembrete 
      Caption         =   "Lembrete"
      Height          =   615
      Left            =   10080
      Picture         =   "frmTransfConsign.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   0
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid dtg 
      Bindings        =   "frmTransfConsign.frx":0D44
      Height          =   3975
      Left            =   120
      TabIndex        =   29
      Top             =   1680
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   7011
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
      Caption         =   "Produtos a serem Transferidos"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "item"
         Caption         =   "Item"
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
         DataField       =   "ref_pro"
         Caption         =   "Ref"
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
         DataField       =   "qtd_con"
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
         DataField       =   "unit_con"
         Caption         =   "Unit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
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
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "cod_usu"
         Caption         =   "Usuário"
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
      BeginProperty Column08 
         DataField       =   "retificadora"
         Caption         =   "Retificadora"
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
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   6570,142
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1574,929
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   16020
      TabIndex        =   24
      Top             =   7680
      Width           =   16050
      Begin VB.CommandButton cmdReceber 
         Caption         =   "C. Rec."
         Height          =   570
         Left            =   2880
         Picture         =   "frmTransfConsign.frx":0D56
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Consig"
         Height          =   570
         Left            =   2040
         Picture         =   "frmTransfConsign.frx":0E0B
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdAcertos 
         Caption         =   "Acertos"
         Height          =   570
         Left            =   1200
         Picture         =   "frmTransfConsign.frx":0EC0
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3720
         ScaleHeight     =   1065
         ScaleWidth      =   7185
         TabIndex        =   35
         Top             =   120
         Width           =   7215
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite Atual:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblLimiteAtual 
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
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   1920
            TabIndex        =   50
            Tag             =   "!"
            Top             =   435
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite Restante:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label lblLimiteRestante 
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
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1920
            TabIndex        =   48
            Tag             =   "!"
            Top             =   720
            Width           =   525
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Débitos (+ 45 D)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3840
            TabIndex        =   47
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Débitos (1 a 45 D)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3840
            TabIndex        =   46
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Débitos A Vencer:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3840
            TabIndex        =   45
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3840
            TabIndex        =   44
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblDeb45 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5880
            TabIndex        =   43
            Tag             =   "!"
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblDeb 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5880
            TabIndex        =   42
            Tag             =   "!"
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblDebVencer 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5880
            TabIndex        =   41
            Tag             =   "!"
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblDebTotal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5880
            TabIndex        =   40
            Tag             =   "!"
            Top             =   720
            Width           =   390
         End
         Begin VB.Label lblLimiteIni 
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
            ForeColor       =   &H0000C000&
            Height          =   330
            Left            =   1920
            TabIndex        =   39
            Tag             =   "!"
            Top             =   -45
            Width           =   525
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite Inicial:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   0
            Width           =   1140
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblLimiteEmUso 
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
            ForeColor       =   &H0005C9EB&
            Height          =   330
            Left            =   1920
            TabIndex        =   36
            Tag             =   "!"
            Top             =   195
            Width           =   525
         End
      End
      Begin VB.CommandButton cmndbtnConf 
         Caption         =   "F7 -  Confirmar"
         Height          =   855
         Left            =   12720
         Picture         =   "frmTransfConsign.frx":0F75
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "F4 - Excluir"
         Height          =   855
         Left            =   11040
         Picture         =   "frmTransfConsign.frx":16DF
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   855
         Left            =   14400
         Picture         =   "frmTransfConsign.frx":1829
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnUsu 
         Caption         =   "..."
         Height          =   315
         Left            =   3360
         TabIndex        =   19
         Top             =   960
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
         Left            =   120
         TabIndex        =   18
         Text            =   "Nome do Funcionario"
         Top             =   975
         Width           =   3255
      End
      Begin VB.CommandButton cmndbtnBuscaFun 
         Height          =   315
         Left            =   120
         Picture         =   "frmTransfConsign.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   735
         Width           =   840
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.PictureBox pcProdAcerto 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   16020
      TabIndex        =   23
      Top             =   0
      Width           =   16050
      Begin VB.TextBox txtUltConf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Left            =   8280
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   12360
         ScaleHeight     =   945
         ScaleWidth      =   2505
         TabIndex        =   30
         Top             =   240
         Width           =   2535
         Begin VB.Label lblValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   465
            Left            =   840
            TabIndex        =   34
            Tag             =   "!"
            Top             =   0
            Width           =   750
         End
         Begin VB.Label lblVlrDiv3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   465
            Left            =   1320
            TabIndex        =   33
            Tag             =   "!"
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Left            =   480
            TabIndex        =   32
            Top             =   600
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   840
            TabIndex        =   31
            Top             =   645
            Width           =   105
         End
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7695
      End
      Begin VB.CheckBox chkRetif 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Retificadora"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10080
         TabIndex        =   16
         Top             =   960
         Width           =   1335
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
         Left            =   7800
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtbxDescReais 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6960
         TabIndex        =   12
         Text            =   "0,00"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtbxDescPercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6120
         TabIndex        =   10
         Text            =   "0,0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4680
         TabIndex        =   6
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4455
      End
      Begin VB.CommandButton cmndbtnInserir 
         Caption         =   "Inserir"
         Height          =   255
         Left            =   9120
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnBusca 
         Height          =   405
         Left            =   7800
         Picture         =   "frmTransfConsign.frx":1A8D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUltConf 
         Height          =   375
         Left            =   9480
         Picture         =   "frmTransfConsign.frx":1E17
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conferido em:"
         Height          =   195
         Left            =   8280
         TabIndex        =   56
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblCorrecao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correção de Estoque"
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
         Left            =   10080
         TabIndex        =   28
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(R$)"
         Height          =   195
         Left            =   6960
         TabIndex        =   11
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor Final:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7920
         TabIndex        =   13
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(%)"
         Height          =   195
         Left            =   6120
         TabIndex        =   9
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5280
         TabIndex        =   7
         Top             =   720
         Width           =   585
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
         TabIndex        =   3
         Top             =   720
         Width           =   555
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
         TabIndex        =   5
         Top             =   720
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmTransfConsign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim prod As Produto
  
Dim bDirecao As Integer
Dim bActive As Boolean
Dim preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer
Option Explicit
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
Dim ai As New ADODB.Stream
Dim Indice As Integer




Sub soma()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    Dim cDesconto As Currency
    Dim cFinal As Currency
    On Error GoTo trata_Erro:
    P_Status "Atualizando Valores..."
    
    dtg.Visible = False
    With ado
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            cTotal = cTotal + !total
            
            
            .MoveNext
        Loop
        '.Close
        If .RecordCount > 0 Then
            .MoveLast
            
        End If
    End With
    
    lblValor.caption = FormatVal(cTotal)
    lblVlrDiv3.caption = FormatVal(cTotal / 3)
    lblLimiteRestante.caption = FormatVal(CCur(lblLimiteAtual.caption) - CCur(lblValor.caption))
    

    
    
    dtg.Visible = True
    
    P_Status ""
    Exit Sub
trata_Erro:
P_Status ""
dtg.Visible = True
'    P_MostraErro
End Sub

Sub P_buscaitens()
    Dim strSQL As String
    'On Error GoTo trata_erro
    
    buscaConsignAbertas
    
    
    
    
    
    
    
    
    
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub






Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub









Private Sub cmndbtnAcerto_Click()
    frmAcerto.txtbxRepresentante.tag = txtbxNomeCli.tag
    frmAcerto.txtbxRepresentante.Text = txtbxNomeCli.Text
    frmAcerto.Show 0
    
    
End Sub

Private Sub chkRetif_Click()
    Dim strSQL As String
    
    If chkRetif.Value = 1 Then
        pcProdAcerto.BackColor = vbRed
        lblCorrecao.Visible = True
    Else
        pcProdAcerto.BackColor = vbWhite
        lblCorrecao.Visible = False
    End If
    
    
    If chkRetif.Value = 1 Then
        strSQL = "Update tab_consignabertas set retif_con='S' where cod_rep='" & Me.tag & "' "
        Conn.Execute strSQL
        
    Else
        strSQL = "Update tab_consignabertas set retif_con='N' where cod_rep='" & Me.tag & "' "
        Conn.Execute strSQL
    End If
    buscaConsignAbertas
End Sub

Private Sub cmdAcertos_Click()
    P_Acertos txtbxNomeCli.tag
End Sub

Private Sub cmdLembrete_Click()
    buscaLembrete txtbxNomeCli.tag, True
End Sub

Private Sub cmdReceber_Click()
    frmEditaMov.txtbxTerceiro.Text = txtbxNomeCli.Text
    frmEditaMov.cmbxTipoTer.ListIndex = 1
    frmEditaMov.tag = "x"
    frmEditaMov.Show 0
    
    frmEditaMov.SetFocus
End Sub

Private Sub cmdUltConf_Click()
    Dim strSQL As String
    If txtUltConf.Text <> "" And txtbxNomeCli.tag <> "" Then
            strSQL = "update tab_representantes set ultimaconferencia=" & F_GravaData(txtUltConf.Text)
            strSQL = strSQL & " where cod_rep='" & txtbxNomeCli.tag & "'"
            Conn.Execute strSQL
        End If
End Sub

Private Sub cmndbtnBusca_Click()
    cmndbtnCliente_Click
    
End Sub

Private Sub cmndbtnBuscaFun_Click()
    Dim f As funcionario.funcionario
    On Error GoTo trata_Erro
    
    
    f = funcionario.F_SelectVendedor("")
    
    
    
    If f.cod_fun <> "" Then
        If F_senhaReturnUsu(True, False, f.guerra_fun, False) Then
        
        
            txtbxVendedor.Text = f.nome_fun
            txtbxVendedor.tag = f.cod_fun
            lblDesc.caption = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & f.cod_fun & "'")
            
        End If
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Private Sub cmndbtnCliente_Click()
    Dim c As representante.representante
    Dim UltAcerto As String
    Dim dif As String
    On Error GoTo trata_Erro
    c = representante.F_BuscaRep(txtbxNomeCli.Text)
    If c.cod_Rep <> "" Then
        P_Status "."
        Me.tag = c.cod_Rep
        txtbxNomeCli.Text = c.nome_rep
        txtbxNomeCli.tag = c.cod_Rep
        P_Status ".."
        txtUltConf.Text = c.ultconf
        P_buscaitens
        buscaDebRep
        P_Status "..."
            
        
        txtbxNomeCli.tag = c.cod_Rep
        lblLimiteIni.caption = FormatVal(c.limite_rep)
        lblLimiteEmUso.caption = FormatVal(F_BuscaLimiteEmUsoRep(c.cod_Rep))
        lblLimiteAtual.caption = FormatVal(CCur(lblLimiteIni.caption) - CCur(lblLimiteEmUso.caption))
        lblLimiteRestante.caption = FormatVal(CCur(lblLimiteAtual.caption) - CCur(lblValor.caption))
        P_Status "...."
        BuscaMensagem "R", c.cod_Rep
        If c.status_rep = "2" Then
            MsgBox "O Representante encontra-se negativado, não é possível a transferência!", vbExclamation
            txtbxNomeCli.tag = ""
            txtbxNomeCli.Text = ""
            txtbxNomeCli.SetFocus
                
        End If
        
        
        
        
            
            
            P_Status "....."
        buscaLembrete c.cod_Rep
        UltAcerto = F_BuscaDataUltAcerto(c.cod_Rep)
        dif = DateDiff("D", UltAcerto, Date)
        
        
        If dif > 30 Then
            MsgBox "Este representante fez o último acerto há " & dif & " dias", vbInformation
        End If
         
         
         P_Status
        
        
        
        'pcLimite.Visible = True
        
        
        
        'carrega
        
    Else
        
    End If
     
            txtbxProduto.SetFocus
     Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmbxFormaPag_Change()

End Sub

Private Sub cmbxFormaPag_DropDown()

End Sub










Private Sub CmndbtnConf_Click()
    buscaDebRep
    concluiSaidas
End Sub

Private Sub cmndbtnLocaliza_Click()
    
    If F_Vazio(txtbxProduto.Text) Then
'        MsgBox "Informe o produto que deseja inserir!", vbExclamation, constNome_sis
'        txtbxProduto.SetFocus
'        Exit Sub
    End If
    If txtbxNomeCli.tag = "" Then
        MsgBox "Escolha o representante", vbExclamation
        Exit Sub
    End If
    
    buscaProduto
End Sub

Private Sub cmndbtnEdit_Click()
    On Error GoTo trata_Erro
    With ado
        cmndbtnInserir.tag = ado!seq_
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnExcl_Click()
    excluiItem
End Sub
Sub excluiItem()
    On Error GoTo trata_Erro
    If MsgBox("Confirma exclusão do item " & ado!nome_pro & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If Consignado.F_DeletaConsignAberta(ado!seq_con) Then
        buscaConsignAbertas
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub


Private Sub cmndbtnInserir_Click()
    
    insere
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    ado.Close
    
    
    
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro "cmndbtnSair"
    Unload Me
End Sub










Private Sub dtgProd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub



Private Sub cmndbtnUsu_Click()
    Dim strSQL As String
    strSQL = F_senhaReturnUsu(, , strNome_Usu)
    If strSQL <> "" Then
        
    End If
End Sub


Private Sub Command1_Click()
    frmConsign.txtbxNomeCli.Text = txtbxNomeCli.Text
    frmConsign.txtbxNomeCli.tag = txtbxNomeCli.tag
    frmConsign.tag = "t"
    frmConsign.Show
 End Sub

Private Sub Form_Activate()
    On Error GoTo trata_Erro
    
    
    
    If strEfetuarTransfCons_Usu <> "S" Then
        MsgBox "Não tem permissão para efetuar transferência!", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    If bActive = True Then Exit Sub
    
    
    
    
    If Me.tag <> "" Then
        txtbxNomeCli.tag = Me.tag
        txtbxNomeCli.Text = F_BuscaRepCodigo(Me.tag).nome_rep
    End If
    
    If txtbxNomeCli.tag = "" Then
        Exit Sub
    End If
    
    P_Status "Buscando Itens..."
    P_Status "..."
    P_buscaitens
    
    P_Status ""
    
    bActive = True
    
    If txtbxNomeCli.tag = "" Then
        txtbxNomeCli.SetFocus
    End If
    P_Status "....->"
    BuscaLimite
    P_Status ""
    
    Exit Sub
trata_Erro:
    P_MostraErro "Form_Activate"
    
End Sub
Sub BuscaLimite()
    'lblLimiteAtual.caption = FormatVal(F_BuscaLimiteEmUsoRep(txtbxNomeCli.tag))
    'lblLimiteRestante.caption = FormatVal(CCur(lblLimiteAtual.caption) - CCur(lblValor.caption))
    
    
End Sub
Sub buscaConsignAbertas()
    Dim strSQL As String
    
    
    strSQL = "select @rank := @rank + 1 as item, if(retif_con='S','SIM','NÃO') as retificadora,tp.nome_pro,"
    strSQL = strSQL & " ref_pro,unit_con*qtd_con as total,tc.* from tab_consignabertas as tc, tab_produtos as tp "
    strSQL = strSQL & " where tp.cod_pro=tc.cod_pro "
    strSQL = strSQL & " and tc.cod_rep='" & txtbxNomeCli.tag & "' "
    strSQL = strSQL & " order by data_con desc,seq_con desc"
    
    abreObjetoAdo ado, strSQL, dtg
    soma
    
    
End Sub
Sub playGif()
    'WebBrowser1.Navigate "c:\market\gifs\maebela2.gif"
    'WebBrowser1.Navigate "about:<html><body scroll='no'><img src='c:\maebela.gif'></img></body></html>"
    
End Sub

    


Sub abreDevolucao()
    
    frmDevolucaoConsign.txtbxRepresentante.Text = txtbxNomeCli.Text
    frmDevolucaoConsign.txtbxRepresentante.tag = txtbxNomeCli.tag
    frmDevolucaoConsign.Show 0
    
    
    
    
    
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    bDescancoTela = False
    iDescancoTela = 0
    
    
    
    If KeyCode = vbKeyF3 Then
        
    ElseIf KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    ElseIf KeyCode = vbKeyF5 Then
    
        
        dtg.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
         
    ElseIf KeyCode = vbKeyF6 Then
        abreDevolucao
    ElseIf KeyCode = vbKeyF7 Then
        
        CmndbtnConf_Click
    ElseIf KeyCode = vbKeyF2 Then
        lancaPosologia
    ElseIf KeyCode = vbKeyF8 Then
        
        
    
        operacao = "C"
        mudamodo
    ElseIf KeyCode = vbKeyF9 Then
        'If operacao = "C" And ado.RecordCount > 0 Then
        '    MsgBox "Operacação Impossível!", vbCritical
        '    Exit Sub
        'End If
        
        
        operacao = "B"
        mudamodo
    ElseIf KeyCode = vbKeyF10 Then
        
        
        
        
    
        
    ElseIf KeyCode = vbKeyF11 Then
        
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        frmECF.Show 1
    End If
    
    If KeyCode = vbKeyC And (Shift = vbAltMask Or Shift = vbCtrlMask) Then
        txtbxNomeCli.SetFocus
    ElseIf KeyCode = vbKeyP And Shift = vbCtrlMask Then
        
    ElseIf KeyCode = vbKeyE And Shift = vbAltMask Then
        
    End If
    
End Sub
Sub mudamodo()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set modo_cai='" & operacao & "' "
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
    Conn.Execute strSQL
End Sub
Private Sub Form_Load()
    resizeForm Me
    bActive = False
    Set ado = New ADODB.Recordset
    txtbxVendedor.Text = strNome_Usu
    txtbxVendedor.tag = strCod_Usu
    
    
    
    
    
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
    
    
    
    PosicionaGrade dtg, Me, pcStatusBar
    
    
    
 
    
    
    
    
End Sub

Sub editaItem()
    txtbxQtd.Enabled = True
    txtbxQtd.Text = "1"
    txtbxQtd.SetFocus
End Sub

Private Sub optnCod_Click()
    txtbxQtd.SetFocus
End Sub

Private Sub optnDesc_Click()
    txtbxQtd.SetFocus
End Sub

Private Sub optnEnd_Click()
    
End Sub

Private Sub optnSeq_Click()
    P_buscaitens
End Sub



Private Sub Form_Terminate()
    On Error Resume Next
    
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFrm.Enabled = True
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



Private Sub optnCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtbxNomeCli.SetFocus
    End If

End Sub

Private Sub lblDelExcluir_Click()
    
End Sub
Sub concluiSaidas()
    Dim strSQL As String
    Dim iCount As Integer
    Dim sSeq As String
    Dim sCab As String
    If ado.RecordCount > 0 Then
        ado.MoveFirst
    End If
    If ado.RecordCount = 0 Then
        MsgBox "Não existem transferências para serem concluídas!", vbExclamation
        Exit Sub
        
    End If
    If chkRetif.Value <> 1 Then
        If CCur(lblLimiteRestante.caption) < 0 Then
            MsgBox "Observe que o limite é insuficiente para a transação!", vbExclamation
            Exit Sub
        End If
    End If
    If CCur(lblDeb45.caption) > 0 And chkRetif.Value <> 1 Then
        MsgBox "A Transferência não pode ser concluída pois o cliente possui débito superior a 45 dias!", vbExclamation
        Exit Sub
    End If
    If chkRetif.Value = 0 Then
        If Not F_SenhaRep(txtbxNomeCli.tag) Then
            Exit Sub
        End If
    Else
        If strCorrigeEstCons_Usu <> "S" Then
            MsgBox "Você não tem permissão para corrigir estoque!", vbExclamation
            Exit Sub
        End If
    End If
    'If MsgBox("Confirma conclusão dos lançamentos para transferência?", vbYesNo + vbQuestion) = vbNo Then
    '    Exit Sub
    'End If
    
    If F_senhaReturnUsu() = "" Then
        Exit Sub
    End If
    
    If strEfetuarTransfCons_Usu <> "S" Then
        MsgBox "Não tem permissão para efetuar transferência!", vbExclamation
        Exit Sub
    End If
    dtg.Visible = False
    sSeq = MaxKeySQL("tab_consignacoes", "seq_con", "", "000000#") 'buscando o sequencial q servirá para todos os produtos da transf
    With ado
        Do While Not .EOF
            If Consignado.F_GravaConsign(sSeq, !cod_pro, !qtd_con, !cod_Rep, !unit_con, True, !retif_con = "S") = "" Then
                MsgBox "Erro na transferência no item " & .AbsolutePosition, vbCritical
                
            Else
                If Not Consignado.F_DeletaConsignAberta(!seq_con) Then
                'aqui
                End If
            End If
            
            .MoveNext
        Loop
        
    End With
    dtg.Visible = True
    P_buscaitens
        
    imprimeConsign sSeq
        
End Sub
Sub lancaPosologia()
    'I'f ado.RecordCount = 0 Then
    '    MsgBox "Não existe itens na sua venda!", vbExclamation
    '    Exit Sub
    'e 'nd If
    'i ''f dtgProd.Row = -1 Then
    '    MsgBox "Selecione o Item para lançar a posologia!", vbExclamation
    '    dtgProd.SetFocus
    '    Exit Sub
    'End If
    'Venda.P_AddPosologia Me.tag, ado!seq_ite
    
End Sub


Private Sub pcMsg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub pcPromo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub




Private Sub pcProd_GotFocus()
    On Error Resume Next
    txtbxProduto.SetFocus
End Sub



Private Sub pcProdAcerto_GotFocus()
    Sendkeys Chr(9)
End Sub


Private Sub txtbxBairro_GotFocus()
    'P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub txtbxDescPercent_GotFocus()
    P_SelecTexto txtbxDescPercent
End Sub

Private Sub txtbxDescPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescPercent_LostFocus()
    On Error GoTo trata_Erro
    txtbxDescReais.Text = FormatVal(CCur(txtbxDescPercent.Text) * CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) / 100)
    Exit Sub
trata_Erro:
    txtbxDescReais.Text = "0,00"
End Sub

Private Sub txtbxDescReais_GotFocus()
    P_SelecTexto txtbxDescReais
End Sub

Private Sub txtbxDescReais_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescReais_LostFocus()
    On Error Resume Next
    txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    p_totalItem
End Sub

Private Sub txtbxEndereco_GotFocus()
 '   P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub txtbxFinal_GotFocus()
    P_SelecTexto txtbxFinal
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
    On Error Resume Next
    txtbxDescReais.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxFinal.Text))
    txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    Exit Sub
trata_Erro:
    
End Sub



Private Sub txtbxNomeCli_GotFocus()
    P_SelecTexto txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
    
End Sub


Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub

Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli, 1
    
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    Dim p As Produto
    Dim rPesq As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click

    ElseIf KeyCode = vbKeyTab Then
        txtbxQtd.Enabled = True
    ElseIf KeyCode = vbKeyDown Then
        
    ElseIf KeyCode = vbKeyF6 Then
        'cmndbtnLocaliza_Click
    ElseIf KeyCode = vbKeyF11 Then
        
        
    End If
    If KeyCode = vbKeyUp Then
        
    End If
End Sub

Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim e As EstadoImp
    Dim sQtd As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoQtd As String
    Dim sUsu As String
    
    If CCur(txtbxFinal.Text) = 0 Then
        MsgBox "O Valor final está zerado!", vbExclamation
        Exit Sub
    End If
    
    
    If chkRetif.Value = 1 Then
        
        
    
        strSQL = "A Transferência Retificadora deve ser feita quando existe um produto que foi transferido errado anteriormente!"
        strSQL = strSQL & Chr(13) & "Deseja continuar?"
        If MsgBox(strSQL, vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        
        If strCorrigeEstCons_Usu <> "S" Then
            MsgBox "Você não tem permissão para corrigir estoque!", vbExclamation
            Exit Sub
            
        End If
        
    End If
    If CCur(lblDeb45.caption) > 0 Then
        If MsgBox("A Transferência não pode ser concluída pois o cliente possui débito superior a 45 dias!" _
            & Chr(13) & "Deseja Continuar", vbExclamation + vbYesNo) = vbNo Then
        
            Exit Sub
        End If
    End If
    
    If txtbxNomeCli.tag = "" Then
        cmndbtnBusca_Click
    End If
    If CCur(txtbxDescPercent.Text) > CCur(lblDesc.caption) Then
        MsgBox "O desconto máximo permitido para o usuário é de " & lblDesc.caption, vbExclamation
        Exit Sub
    End If
    
    If txtbxProduto.Text = "" Then
        MsgBox "Informe pelo menos parte do nome do produto para que seja localizado!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    On Error GoTo trata_Erro
    strSQL = F_BuscaUmCampo("seq_con", "tab_consignabertas", "cod_rep='" & Me.tag & "' and cod_pro='" & prod.cod_pro & "'")
    If sQtd = "" Then
        
    End If
    
    
    If Int(txtbxQtd.Text) <> CCur(txtbxQtd.Text) Then
        
        If prod.fracionado_pro <> "S" Then
            MsgBox "Este produto não pode ser transferido em quantidade fracionada!", vbExclamation
            Exit Sub
        End If
        
        
    End If
    
        
        
        If strSQL <> "" Then
            If MsgBox("Já existe esteeste produto. Deseja acrescentar na quantidade total?", vbYesNo + vbQuestion) = vbYes Then
                If Consignado.F_AumentaQtdConsignAberta(strSQL, txtbxQtd.Text) Then
                    P_buscaitens
                End If
                Exit Sub
            End If
        Else
                                
                                
                                
                iItem = iItem + 1
                If Consignado.F_GravaConsignAberta(prod.cod_pro, txtbxQtd.Text, txtbxNomeCli.tag, txtbxUnit.Text, chkRetif.Value = 1) = "" Then
                    MsgBox "Erro na inserção do item!", vbExclamation
                End If
                    
                P_buscaitens
                Me.Refresh
                
                txtbxQtd.Text = "1"
                'ado.Move ado.RecordCount - 1
                txtbxUnit.Text = "0"
                txtbxDescPercent.Text = "0,0"
                txtbxDescReais.Text = "0,00"
                'chkRetif.Value = 0
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
            
        End If
        BuscaLimite
        
        'On Error Resume Next

        Exit Sub
trata_Erro:
    P_MostraErro "Insere()"
End Sub
Sub buscaProduto()
    
    
    prod = Produtos.F_BuscaProdSaida(txtbxProduto.Text, "B")
            
            
            
    
    If prod.cod_pro <> "" Then
            txtbxDescPercent.Text = "0"
            txtbxDescReais.Text = "0,00"
            txtbxProduto.tag = prod.cod_pro
            txtbxProduto.Text = prod.nome_pro
        
            'lblNomePro.caption = p.nome_pro
            
            'lblCodPro.caption = p.cod_pro
'            lblSeqPro.caption = p.est.seq_pro
            
            
            'lblSec.caption = p.sub.sec.nome_sec
            'lblSubSec.caption = p.sub.nome_sub
            
            'lblDescSub.caption = p.sub.descmax_sub
            'lblPreco.caption = p.pv_pro
            If bUsaPrecoAtacado > 0 Then
                txtbxUnit.Text = FormatVal(prod.pa_pro)
            Else
                txtbxUnit.Text = FormatVal(prod.pv_pro)
            End If
            'txtbxDescPercent.Text = CCur(txtbxDescForma.Text) + CCur(p.sub.descmax_sub)
            
            Sendkeys Chr(9)
    End If
End Sub
Sub buscaDesc()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_descontos where cod_pro='" & txtbxProduto.tag & "' "
    strSQL = strSQL & " and cod_car='"
    
    
    
    
    
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo trata_Erro
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
      
    End If
Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub p_totalItem()
    txtbxFinal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxDescReais.Text))
    
End Sub
Private Sub txtbxQtd_LostFocus()
    On Error Resume Next
    If CCur(txtbxQtd.Text) > 999 Then
        If MsgBox("A quantidade digitada é muito grande." & _
            "Deseja Continuar?", vbYesNo + vbCritical + vbDefaultButton2, "Atenção!!!") = vbNo Then
            txtbxQtd.Text = "1"
            Exit Sub
        End If
    End If
    If txtbxQtd.Text = "0" Then
        MsgBox "A quantidade está zerada!", vbCritical
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If Int(txtbxQtd.Text) <> CCur(txtbxQtd.Text) Then
        P_ValorLostFocus txtbxQtd, "###,##0.000"
    Else
        P_ValorLostFocus txtbxQtd, "###,##0"
    End If
    p_totalItem
    'txtbxQtd.Enabled = False
    'txtbxProduto.Text = ""
    'txtbxProduto.Enabled = True
    'txtbxProduto.SetFocus
    
    
    
End Sub



Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit, "###,##0.000"
    p_totalItem
End Sub

















Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         
    End If
End Sub
Sub buscaDebRep()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cDeb As Currency, cDeb45 As Currency, cDebTotal As Currency, cDebVencer As Currency
    strSQL = "select ti.venc_itd,ti.reman_itd "
    strSQL = strSQL & " from tab_itdoc as ti,tab_doc as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and ti.reman_itd>0 "
    strSQL = strSQL & " and td.cod_ter='" & txtbxNomeCli.tag & "' "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and td.tipo_ter='R' "
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If CDate(!venc_itd) < CDate(DateAdd("D", -15, Date)) Then
                cDeb45 = cDeb45 + !reman_itd
            ElseIf CDate(!venc_itd < CDate(Date)) Then
                cDeb = cDeb + !reman_itd
            Else
                cDebVencer = cDebVencer + !reman_itd
            End If
            .MoveNext
        Loop
        .Close
    End With
    lblDeb.caption = FormatVal(cDeb)
    lblDeb45.caption = FormatVal(cDeb45)
    lblDebVencer.caption = FormatVal(cDebVencer)
    lblDebTotal.caption = FormatVal(cDeb + cDeb45 + cDebVencer)
    
    
End Sub

Private Sub txtUltConf_GotFocus()
    P_SelecTexto txtUltConf
End Sub

Private Sub txtUltConf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtUltConf_LostFocus()
    P_AjustaData txtUltConf
End Sub
