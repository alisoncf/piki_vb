VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMovInt 
   Caption         =   "Movimento Interno"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15660
   Icon            =   "frmMovInt2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   15660
   Begin VB.CommandButton cmndbtnLocaliza 
      Caption         =   "..."
      Height          =   255
      Left            =   14760
      TabIndex        =   130
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Data dtProd 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox pcAtalhos 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   15630
      TabIndex        =   105
      Top             =   0
      Width           =   15660
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
         TabIndex        =   127
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
            Picture         =   "frmMovInt2.frx":030A
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   128
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
            TabIndex        =   129
            Top             =   120
            Width           =   735
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
         TabIndex        =   124
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmMovInt2.frx":0475
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   125
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
            TabIndex        =   126
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
         Index           =   5
         Left            =   9000
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   121
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
            Picture         =   "frmMovInt2.frx":07FF
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   122
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
            TabIndex        =   123
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
         Index           =   4
         Left            =   7200
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   118
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
            Picture         =   "frmMovInt2.frx":2BC6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   119
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
            TabIndex        =   120
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
         Index           =   3
         Left            =   5400
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   115
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
            Picture         =   "frmMovInt2.frx":35C8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   116
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F4 -Exclui Item"
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
            TabIndex        =   117
            Top             =   120
            Width           =   1170
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
         TabIndex        =   112
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
            Picture         =   "frmMovInt2.frx":3712
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   113
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F3 -Muda Qtd"
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
            TabIndex        =   114
            Top             =   120
            Width           =   1065
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
         TabIndex        =   109
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
            Picture         =   "frmMovInt2.frx":3986
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   110
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
            TabIndex        =   111
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
         Index           =   1
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   106
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   120
            Picture         =   "frmMovInt2.frx":3D10
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   107
            Top             =   120
            Width           =   240
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
            TabIndex        =   108
            Top             =   120
            Width           =   1050
         End
      End
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   15630
      TabIndex        =   90
      Top             =   9555
      Width           =   15660
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
         TabIndex        =   92
         Text            =   "Nome do Funcionario"
         Top             =   15
         Width           =   3255
      End
      Begin VB.CommandButton cmndbtnBuscaFun 
         Height          =   315
         Left            =   3840
         Picture         =   "frmMovInt2.frx":4712
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atendente:"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   104
         Top             =   15
         Width           =   840
      End
      Begin VB.Label lblnum_ven 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
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
         Left            =   6600
         TabIndex        =   103
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5760
         TabIndex        =   102
         Top             =   60
         Width           =   630
      End
      Begin VB.Label lblNomeEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
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
         Left            =   12480
         TabIndex        =   101
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8160
         TabIndex        =   100
         Top             =   60
         Width           =   435
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
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
         Left            =   8760
         TabIndex        =   99
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   98
         Top             =   60
         Width           =   465
      End
      Begin VB.Label lblBloqueado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
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
         Left            =   10560
         TabIndex        =   97
         Top             =   0
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblTur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00"
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
         Left            =   9960
         TabIndex        =   96
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   95
         Top             =   60
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
         Left            =   4680
         TabIndex        =   94
         Top             =   0
         Width           =   135
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
         TabIndex        =   93
         Top             =   52
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   8985
      ScaleWidth      =   5385
      TabIndex        =   56
      Top             =   540
      Width           =   5415
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   0
         ScaleHeight     =   4665
         ScaleWidth      =   5385
         TabIndex        =   70
         Top             =   5400
         Width           =   5415
         Begin VB.Label lblCapPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto:"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lblNomePro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Width           =   5085
         End
         Begin VB.Label lblCapPreco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço:"
            Height          =   195
            Left            =   2640
            TabIndex        =   84
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label lblPreco 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   83
            Top             =   1560
            Width           =   1605
         End
         Begin VB.Label lblQtd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4320
            TabIndex        =   82
            Top             =   1560
            Width           =   885
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   4320
            TabIndex        =   81
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seção:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   720
            Width           =   510
         End
         Begin VB.Label lblSec 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   79
            Top             =   960
            Width           =   5085
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Seção:"
            Height          =   195
            Left            =   2640
            TabIndex        =   78
            Top             =   3360
            Width           =   840
         End
         Begin VB.Label lblSubSec 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   675
            Left            =   2640
            TabIndex        =   77
            Top             =   3600
            Width           =   2565
         End
         Begin VB.Label lblCodPro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   76
            Top             =   2160
            Width           =   2565
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   2640
            TabIndex        =   75
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label lblRef 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   74
            Top             =   2880
            Width           =   1605
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            Height          =   195
            Left            =   2640
            TabIndex        =   73
            Top             =   2640
            Width           =   825
         End
         Begin VB.Label lblDescSub 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4320
            TabIndex        =   72
            Top             =   2880
            Width           =   885
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc %"
            Height          =   195
            Left            =   4440
            TabIndex        =   71
            Top             =   2640
            Width           =   540
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   5385
         TabIndex        =   58
         Top             =   3240
         Width           =   5415
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   64
            Top             =   0
            Width           =   5415
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
               TabIndex        =   68
               Top             =   0
               Width           =   900
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
               TabIndex        =   67
               Tag             =   "!"
               Top             =   0
               Width           =   105
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
               TabIndex        =   66
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
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
               TabIndex        =   65
               Top             =   0
               Width           =   510
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   59
            Top             =   1680
            Width           =   5415
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
               TabIndex        =   63
               Top             =   0
               Width           =   930
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
               TabIndex        =   62
               Tag             =   "!"
               Top             =   0
               Width           =   105
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
               TabIndex        =   61
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor:"
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
               TabIndex        =   60
               Top             =   0
               Width           =   570
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
            ForeColor       =   &H000000FF&
            Height          =   1440
            Left            =   1560
            TabIndex        =   69
            Tag             =   "!"
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         Picture         =   "frmMovInt2.frx":480B
         ScaleHeight     =   2865
         ScaleWidth      =   2385
         TabIndex        =   57
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label lblPExtenso 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Left            =   2520
         TabIndex        =   89
         Top             =   5160
         Width           =   150
      End
      Begin VB.Label lblTotalQueima 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   4200
         TabIndex        =   88
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Queima:"
         Height          =   195
         Left            =   2520
         TabIndex        =   87
         Top             =   3000
         Width           =   1215
      End
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   5385
      TabIndex        =   17
      Top             =   600
      Width           =   5415
      Begin VB.CommandButton cmnbtnRecalc 
         Caption         =   "Recalcula Desconto"
         Height          =   435
         Left            =   3720
         Picture         =   "frmMovInt2.frx":EBBA
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   315
         Left            =   2760
         Picture         =   "frmMovInt2.frx":EF44
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         TabIndex        =   49
         Tag             =   "00000"
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   330
         Left            =   5040
         Picture         =   "frmMovInt2.frx":F03D
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   840
         Width           =   255
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
         TabIndex        =   47
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtbxDescForma 
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
         Left            =   3120
         TabIndex        =   46
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox pcEnd 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   -360
         ScaleHeight     =   1455
         ScaleWidth      =   9135
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   9135
         Begin VB.TextBox txtbxEndereco 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   5175
         End
         Begin VB.TextBox txtbxBairro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   120
            TabIndex        =   41
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txtbxFone 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   3840
            TabIndex        =   40
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkEntrega 
            Caption         =   "En&trega (ctrl + E)"
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblBairro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   450
         End
         Begin VB.Label lblFone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fone:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3840
            TabIndex        =   43
            Top             =   840
            Width           =   405
         End
      End
      Begin VB.PictureBox pcLimite 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   5175
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   615
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
            TabIndex        =   36
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1080
            TabIndex        =   35
            Top             =   0
            Width           =   405
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
            Left            =   1080
            TabIndex        =   34
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   1920
            TabIndex        =   33
            Top             =   0
            Width           =   600
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
            Left            =   1920
            TabIndex        =   32
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   2880
            TabIndex        =   31
            Top             =   0
            Width           =   690
         End
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
            Left            =   2880
            TabIndex        =   30
            Top             =   240
            Width           =   390
         End
      End
      Begin VB.PictureBox pcDesconto 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1065
         ScaleWidth      =   5145
         TabIndex        =   18
         Top             =   1680
         Width           =   5175
         Begin VB.CommandButton cmndbtnDesconto 
            Caption         =   "Confirmar Gerencial"
            Height          =   525
            Left            =   1920
            Picture         =   "frmMovInt2.frx":F136
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtbxDescPerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   20
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtbxDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   19
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
            TabIndex        =   28
            Top             =   480
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final Previsto:"
            Height          =   195
            Left            =   4080
            TabIndex        =   27
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   1200
            TabIndex        =   26
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Autorizado Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
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
            TabIndex        =   23
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
            Left            =   90
            TabIndex        =   22
            Top             =   -30
            Width           =   1860
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Forma de Pagamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   0
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3120
         TabIndex        =   54
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   525
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
         Left            =   1800
         TabIndex        =   52
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5520
      ScaleHeight     =   945
      ScaleWidth      =   10545
      TabIndex        =   0
      Top             =   600
      Width           =   10575
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4680
         TabIndex        =   7
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxDescPercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6840
         TabIndex        =   5
         Text            =   "0,0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxDescReais 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   7680
         TabIndex        =   4
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
         Left            =   8520
         TabIndex        =   3
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
         Left            =   6000
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmndbtnInserir 
         Caption         =   "Inserir"
         Height          =   285
         Left            =   8520
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         Left            =   5160
         TabIndex        =   14
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(%)"
         Height          =   195
         Left            =   6840
         TabIndex        =   13
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
         Left            =   8640
         TabIndex        =   12
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(R$)"
         Height          =   195
         Left            =   7680
         TabIndex        =   11
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblQueima 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "Queima!!!"
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
         TabIndex        =   10
         Top             =   480
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
         Left            =   6000
         TabIndex        =   9
         Top             =   0
         Width           =   405
      End
   End
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmMovInt2.frx":F8A0
      Height          =   8175
      Left            =   5520
      OleObjectBlob   =   "frmMovInt2.frx":F8B5
      TabIndex        =   131
      Top             =   1440
      Width           =   9375
   End
End
Attribute VB_Name = "frmMovInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p As ProdutoSaida
Dim bBlock As Boolean
Dim bDirecao As Integer
Dim bActive As Boolean
Dim Preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer

Public Atacado As Boolean

Private Type prodPromo
    Nome As String
    Preco As String
End Type

Dim pPromo() As prodPromo
Dim QtdPromo As Integer
Dim iIndPromo As Integer

Sub carrega()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    If bActive = False Then
        strSQL = F_VerificaVendaAberta(Me.tag)
        If strSQL <> "" Then
            MsgBox "A venda se encontra aberta no computador " & strSQL, vbExclamation
            Unload Me
            Exit Sub
        End If
    End If
    strSQL = "select tv.*,tf.*,tpp.nome_tip from tab_vendas as tv, tab_formaspag as tf, tab_tipopag as tpp "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for "
    strSQL = strSQL & " and num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tpp.cod_tip=tf.cod_tip "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        txtbxNomeCli.Text = !nomecli_ven
        txtbxNomeCli.tag = !cod_cli
        
        If UCase(txtbxNomeCli.Text) = "VENDA A CONSUMIDOR" Then
            txtbxNomeCli.Text = ""
        End If
        txtbxFormaPag.Text = !nome_for
        txtbxDescForma.Text = !desconto_for
        txtbxFormaPag.tag = !cod_for
        txtbxEndereco.Text = Format(!endereco_ven)
        txtbxBairro.Text = Format(!bairro_ven)
        txtbxFone.Text = Format(!fone_ven)
        lblTipoPag.tag = !cod_tip
        lblTipoPag.caption = !nome_tip
        lblDescontoMax.caption = !desconto_for + CCur(lblDesc.caption)
        If !entrega_ven = "S" Then
            chkEntrega.Value = 1
            
        End If
        P_GravaVendaAberta Me.tag
        .Close
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
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
    
    iPos = dtProd.Recordset.AbsolutePosition
    dtProd.Refresh
    With dtProd.Recordset
        
        
        Do While Not .EOF
            cTotal = cTotal + !Final
            cPq = cPq + !pq_itd * !qtd_ite
            cItens = cItens + !qtd_ite
            cDesc = cDesc + !desconto_ite
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cTotal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = dtProd.Recordset.RecordCount
        lblTotalQueima.caption = FormatVal(cPq)
        lblDesconto.caption = FormatVal(cDesc)
        lblTotal.caption = FormatVal(cTotal + cDesc)
    End With
        
    
    
    
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    
    lblProd.caption = dtProd.Recordset.RecordCount
    lblItens.caption = cItens

    strSQL = "update tab_vendas set valor_ven =" & F_GravaDec(cTotal) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(cDesc) & ","
    strSQL = strSQL & "nomecli_ven= '" & txtbxNomeCli.Text & "', "
    strSQL = strSQL & " cod_cli= '" & txtbxNomeCli.tag & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "', "
    strSQL = strSQL & " abrevenda_usu='" & txtbxVendedor.tag & "' "
'    strSQL = strSQL & " obs_VEN='" & txtbxObs.Text & "' "
    strSQL = strSQL & " where num_Ven='" & Me.tag & "' "
    Banco.Execute strSQL
    
    
    P_Status ""
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = Vendas.F_BuscaItens(Me.tag)
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    
            
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub cmndbtnCancel_Click()
    Dim strSQL  As String
    Dim i As Integer
    On Error Resume Next
    If dtProd.Recordset.RecordCount > 0 Then
    
    If strExclVenda_usu <> "S" Then
        strSQL = F_senhaReturnUsu
        If F_BuscaUmCampo("exclvenda_usu", "tab_defusu", "cod_usu='" & strSQL & "'") <> "S" Then
            MsgBox "Usuário Não Autorizado!", vbCritical
            Exit Sub
        End If
    End If
    End If
    If F_BuscaUmCampo("num_ven", "tab_venfec", "num_ven='" & Me.tag & "'") <> "" Then
        MsgBox "Cancelamento impossível", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja cancelar esta venda?", vbYesNo + vbCritical) = vbYes Then
        
        With dtProd.Recordset
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                excluir False
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
        Banco.Execute strSQL
        
        
        If F_BuscaUmCampo("num_ven", "tab_venfec", "num_ven='" & Me.tag & "'") <> "" Then
            MsgBox "Cancelamento impossível", vbExclamation
            Exit Sub
        End If
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
            RefreshDataPosition frmEdiVendas.dtVendas
        End If
        
        If MsgBox("Deseja Gerar uma Nova Venda?" & Chr(13) & "Clique em não caso queira abandonar o modo de vendas!", vbYesNo + vbQuestion) = vbYes Then
            strSQL = F_GeraNovaVenda("00000")
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
    excluir True
End Sub
Sub excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    With dtProd.Recordset
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
        strSQL = "delete from tab_itvendas where seq_ite='" & !seq_ite & "' "
        
        strSQL = strSQL & " and num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        P_GravaItemCancelado
        
        
            strSQL = "update tab_produtos set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))
            
           
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Banco.Execute strSQL
        
        dtProd.Refresh
        If iPos = .RecordCount - 1 Then
            If dtProd.Recordset.RecordCount > 0 Then
                dtProd.Recordset.MoveLast
            End If
        Else
            dtProd.Recordset.Move iPos
        End If
        If bQuestion Then
            P_AtualizaVenda
        End If
        If dtProd.Recordset.RecordCount = 0 Then
            txtbxQtd.Enabled = False
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro "Excluir()"
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
     
        frmSelectCli.txtbxNome.tag = "C"
        frmSelectCli.txtbxNome.Text = txtbxNomeCli.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    
    If F_BuscaUmCampo("status_cli", "tab_clientes", "cod_cli='" & _
        frmSelectCli.dtBusca.Recordset!cod & "'") <> "1" Then
        MsgBox "Este cliente está negativado! " & Chr(13) & _
            "A venda só poderá ser fechada a vista ou a cartão. ", vbCritical
    End If
    
    txtbxNomeCli.Text = frmSelectCli.dtBusca.Recordset!Nome
    txtbxNomeCli.tag = frmSelectCli.dtBusca.Recordset!cod
    
    
    
    
    Unload frmSelectCli
    
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
    
    sDes = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & sUsu & "'")
    If sDes = "" Then
        sDes = "0"
    End If
    If CCur(sDes) < CCur(txtbxDescPerc.Text) Then
        MsgBox "O Desconto possível para este usuário é de " & sDes & "%", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Confirma desconto de " & txtbxDesconto.Text & "?", vbYesNo + vbCritical) = vbYes Then
        Vendas.P_RecalculaVendaDesconto lblFinalPrev.caption, txtbxDescPerc.Text, Me.tag, sUsu
        P_AtualizaVenda
    End If
    
End Sub

Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    Dim strSQL As String
    f = F_SelectFormasPag(txtbxFormaPag.Text)
    
    If f.cod_for <> "" Then
        If CCur(txtbxDescForma.Text) > f.desconto_for Then
            If dtProd.Recordset.RecordCount > 0 Then
                strSQL = "Você optou por mudar a forma de pagamento de " & txtbxFormaPag.Text
                strSQL = strSQL & " para " & f.nome_for & Chr(13)
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
        txtbxFormaPag.Text = f.nome_for
        txtbxFormaPag.tag = f.cod_for
        txtbxDescForma.Text = f.desconto_for
        txtbxDescForma.tag = f.avista_for
        lblDescontoMax.caption = f.desconto_for + CCur(lblDesc.caption)
        txtbxDescPercent.Text = txtbxDescForma.Text
        lblTipoPag.caption = f.nome_tip
        lblTipoPag.tag = f.cod_tip
        
        'P_recalculaVenda
        P_AtualizaFormaPag Me.tag, txtbxFormaPag.tag
        P_AtualizaVenda
        txtbxNomeCli.SetFocus
    End If
End Sub

Private Sub cmndbtnInserir_Click()
    insere
End Sub

Private Sub cmndbtnLocaliza_Click()
    
    buscaProd
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    If dtProd.Recordset.RecordCount > 0 Then
        
            If MsgBox("Confirma Saída da Venda?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            If txtbxNomeCli.Text = "" Or txtbxNomeCli.Text = "VENDA A CONSUMIDOR" Then
                MsgBox "Informe o nome do cliente!", vbExclamation
                txtbxNomeCli.SetFocus
                Exit Sub
            End If
            
            
    End If
    If IsLoad(frmEdiVendas.Name) Then
        frmEdiVendas.dtVendas.Refresh
    End If
    If dtProd.Recordset.RecordCount = 0 Then
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        
    End If
    P_DeletaVendaAberta Me.tag
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro "cmndbtnSair"
    
End Sub


Private Sub cmndbtnSalvar_Click()
    Dim sTroco As String
    Dim strSQL As String
    Dim sVen As String
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não há itens na sua venda!", vbExclamation
        Exit Sub
    End If
    
    
    strSQL = F_senhaReturnUsu(False, 3)
    If strSQL = "" Then
        MsgBox "não autenticado!", vbExclamation
        Exit Sub
    End If
    If strSQL <> txtbxVendedor.tag Then
        txtbxVendedor.tag = strSQL
        txtbxVendedor.Text = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & strSQL & "'")
        Vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text
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
                Vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text
            End If
            
            
            
        End If
        If lblTipoPag.tag = "02" Then
            If F_GravaCrediarioVen(Me.tag) Then
                MsgBox "Esta venda venda deve ser liberada pelo setor de Crediário para ser concretizada!", vbExclamation
            End If
        End If
        
        If Vendas.F_AlteraStatusVenda(Me.tag, "2") Then
            P_DeletaVendaAberta Me.tag
            Me.tag = F_GeraNovaVenda
            If Me.tag = "" Then
                MsgBox "Não foi possível gerar uma nova venda!", vbExclamation
                Exit Sub
            End If
            bActive = False
            Form_Activate
        End If
    
    
    
End Sub


Private Sub dbgrdProd_DblClick()
    editaItem
End Sub

Private Sub dbgrdProd_GotFocus()
    
    dbgrdProd.MarqueeStyle = 3
    
    
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo trata_Erro

    '    On Error GoTo trata_erro

    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        
        
                
                    If dbgrdProd.Row = -1 Then
                        MsgBox "Selecione o item que deseja excluir!", vbCritical
                        dbgrdProd.SetFocus
                        Exit Sub
                    End If
                    
            excluir False
        
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxQtd.SetFocus
    End If
    P_AtualizaVenda
    Exit Sub
trata_Erro:
P_MostraErro

End Sub

Private Sub dbgrdProd_LostFocus()
    dbgrdProd.MarqueeStyle = 0
End Sub


Private Sub dbgrdProd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub dtProd_Reposition()
    With dtProd.Recordset
    On Error Resume Next
    If Not .EOF Then
        lblNomePro.caption = !nome_pro
            
            
                lblCapPreco.caption = "Preço:"
                lblPreco.caption = "R$ " & FormatVal(!pv_pro)
            
            lblSec.caption = !nome_sec
            lblCodPro.caption = !cod_pro
            lblRef.caption = !ref_pro
            lblSubSec.caption = !nome_sub
            lblQtd.caption = !qtd_ite
            
        End If
    End With
        
End Sub


Sub buscaPromo()
    Dim strSQL As String
    Dim rPesq As Recordset
        
    strSQL = "select nome_pro,pv_pro from tab_produtos as tp,"
    strSQL = strSQL & " tab_prodpromo as tpp where tpp.cod_Pro=tp.cod_pro"
    strSQL = strSQL & " and datafin_pro >=" & F_GravaData(Date)
    strSQL = strSQL & " order by nome_pro"
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rPesq
        QtdPromo = .RecordCount
        If .RecordCount > 0 Then
            ReDim pPromo(.RecordCount) As prodPromo
        End If
        strSQL = ""
        Do While Not .EOF
            pPromo(.AbsolutePosition).Nome = !nome_pro
            pPromo(.AbsolutePosition).Preco = "R$ " & FormatVal(!pv_pro)
            
            .MoveNext
        Loop
        .Close
        
        
        
    
    End With
    
End Sub
Private Sub Form_Activate()
    On Error GoTo trata_Erro
    mdiFrm.Enabled = False
    lblDesc.caption = crcDesconto_usu
    lblNomeEmp.caption = strFantasia_Emp
    lblnum_ven.caption = Me.tag
    txtbxVendedor.tag = strCod_Usu
    txtbxVendedor.Text = strNome_Usu
    lblNomeEmp.caption = strFantasia_Emp
    lblCaixa = strCod_Cai
    lblTur.caption = strCod_Tur
    If bActive = True Then Exit Sub
    carrega
    
    
    P_Status "Buscando Itens..."
    txtbxProduto.SetFocus
    P_buscaitens
    P_AtualizaVenda
    P_Status ""
    
    bActive = True
    
    
    Exit Sub
trata_Erro:
    P_MostraErro "Form_Activate"
    
End Sub
Sub playGif()
    'WebBrowser1.Navigate "c:\market\gifs\maebela2.gif"
    'WebBrowser1.Navigate "about:<html><body scroll='no'><img src='c:\maebela.gif'></img></body></html>"
    
End Sub

    


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    bDescancoTela = False
    iDescancoTela = 0
    
    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyF5 Then
        
        dbgrdProd.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
         dbgrdProd_KeyDown vbKeyDelete, 0
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF2 Then
        ImprVenda Me.tag
    ElseIf KeyCode = vbKeyF8 Then
        
    ElseIf KeyCode = vbKeyF9 Then
        
    ElseIf KeyCode = vbKeyF10 Then
        
    ElseIf KeyCode = vbKeyF11 Then
        
    ElseIf KeyCode = vbKeyF12 Or KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    
    End If
End Sub
Sub mudamodo()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set modo_cai='" & operacao & "' "
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
    Banco.Execute strSQL
End Sub
Private Sub Form_Load()
'    buscaPromo
    bActive = False

    dtProd.Connect = Banco.Connect

    lblNomeEmp.caption = strNome_Emp
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", "tab_parametros", "")
    txtbxNomeCli.tag = "00000"
    txtbxNomeCli.Text = "VENDA A CONSUMIDOR"
    
    
    
    
'    cmbxBuscar.ListIndex = 0
    'pcECF.BackColor = vbGreen
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
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
    dbgrdProd.Width = Me.Width - dbgrdProd.Left - 200
    
    
    dbgrdProd.Height = pcStatusBar.Top - dbgrdProd.Top - 100
    lblNomeEmp.Left = Me.Width - lblNomeEmp.Width - 200
    'pcCli.Width = pcProd.Width
    
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).Width = pcAtalhos.Width / (pcBotao.count - 1)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).Width + 50
        End If
        
    Next
 
    pcProd.Width = dbgrdProd.Width
    'pcCli.Width = pcProd.Width
    

End Sub

Sub editaItem()
    'On Error GoTo trata_Erro
    Exit Sub 'para terminar depois
    With dtProd.Recordset
        cmndbtnInserir.tag = !seq_ite
        txtbxProduto.Text = !desc_ite
        txtbxProduto.tag = !cod_pro
        txtbxUnit.Text = !unit_ite
        txtbxQtd.Text = !qtd_ite
        txtbxPQ_pro.Text = !pq_itd
        txtbxDescReais.Text = !desconto_ite
        txtbxDescReais_LostFocus
        txtbxQtd.SetFocus
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
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
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BorderStyle = 0
        
    Next
    pcBotao(Indice).BorderStyle = 1
'    pcbotao(Indice).SetFocus
End Sub

Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_GotFocus(Index As Integer)
    SendKeys Chr(9)
End Sub



Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If Index = 0 Then
            pcBotao(pcBotao.count - 1).SetFocus
        Else
            pcBotao(Index - 1).SetFocus
        End If
        
    ElseIf KeyCode = vbKeyDown Then
        If Index = pcBotao.count - 1 Then
            pcBotao(0).SetFocus
        Else
            pcBotao(Index + 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcBotao_Click Index
    End If
    
     
    
End Sub
Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
End Sub

Private Sub txtbxDescForma_LostFocus()
    If CCur(txtbxDescForma.Text) > CCur(lblDescontoMax.caption) Then
        MsgBox "O Desconto máximo permitido para o usuário com esta forma de pagamento é de " & lblDescontoMax.caption, vbInformation
        Exit Sub
    End If
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
    On Error Resume Next
    P_ValorLostFocus txtbxDesconto
   txtbxDescPerc.Text = FormatVal(CCur(txtbxDesconto.Text) * 100 / CCur(lblValor.caption), "###,##0.0000")

    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.Text))

End Sub

Private Sub txtbxDescPerc_GotFocus()
    P_SelecTexto txtbxDescPerc
End Sub

Private Sub txtbxDescPerc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub




Private Sub txtbxDescPerc_LostFocus()
    P_ValorLostFocus txtbxDescPerc, "###,##0.0000"
    txtbxDesconto.Text = FormatVal(CCur(txtbxDescPerc.Text) / 100 * (CCur(lblValor.caption) - CCur(lblTotalQueima.caption)))
    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.Text))
    
End Sub


Private Sub txtbxDescPercent_GotFocus()
    P_SelecTexto txtbxDescPercent
End Sub
Private Sub pcBotao_Click(Index As Integer)
    If Index = 7 Then
        cmndbtnSair_Click
    ElseIf Index = 6 Then
        cmndbtnSalvar_Click
    ElseIf Index = 5 Then
        cmndbtnCancel_Click
    ElseIf Index = 4 Then
        dbgrdProd.SetFocus
    ElseIf Index = 3 Then
        cmndbtnExc_Prod_Click
    ElseIf Index = 2 Then
        editaItem
    ElseIf Index = 1 Then
        
        cmndbtnFormasPag_Click
    ElseIf Index = 0 Then

    ElseIf Index = 7 Then
        
    End If
End Sub

Private Sub txtbxDescPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDescPercent_LostFocus()
    P_ValorLostFocus txtbxDescPercent
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDescReais_LostFocus()
    'On Error Resume Next
    P_ValorLostFocus txtbxDescReais
    
    'txtbxDescPercent.Text = FormatVal((CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text))
    txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxQtd.Text) / CCur(txtbxUnit.Text) * 100)
    
    
    p_totalItem
    
    
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
'    On Error Resume Next
    'txtbxDescReais.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxFinal.Text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    
    
    
    txtbxDescReais.Text = FormatVal((CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text)) - CCur(txtbxFinal.Text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    txtbxDescReais_LostFocus
trata_Erro:
    
End Sub


Private Sub txtbxFormaPag_GotFocus()
    P_SelecTexto txtbxFormaPag
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxNomeCli_GotFocus()
    P_SelecTexto txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub

Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli, -1
    If txtbxNomeCli.tag <> "" Then
        P_AtualizaVenda
    End If
End Sub

Private Sub txtbxObs_GotFocus()
    
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub

Private Sub txtbxProduto_Change()
    cmndbtnInserir.tag = ""
    txtbxProduto.tag = ""
    
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub
Sub buscaProd()
    
    p = Produtos.F_BuscaProdSaida(txtbxProduto.Text)
    If p.cod_pro <> "" Then
    
        txtbxProduto.Text = p.nome_pro
        txtbxProduto.tag = p.cod_pro
        txtbxUnit.Text = FormatVal(p.pv_pro)
        txtbxUnit.tag = FormatVal(p.pv_pro)
        txtbxUnit.Enabled = False
        
        
        
        
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
            txtbxDescPercent.Text = CCur(txtbxDescForma.Text) + CCur(p.sub.descmax_sub)
        End If
        
        
        
        txtbxQtd.Enabled = True
        txtbxQtd.SetFocus
        
    End If
    
    
End Sub
Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    Dim p As Produto
    Dim rPesq As Recordset
    If KeyCode = vbKeyReturn Then
        buscaProd

    ElseIf KeyCode = vbKeyTab Then
        txtbxQtd.Enabled = True
    ElseIf KeyCode = vbKeyDown Then
        dbgrdProd.SetFocus
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
    Dim rPesq As Recordset
    Dim sTipoQtd As String
    If lblQueima.Visible = False Then
        If CCur(Round(txtbxDescPercent.Text)) > CCur(lblDesc.caption) + CCur(txtbxDescForma.Text) Then
            MsgBox "O desconto máximo permitido para o usuário é de " & CCur(lblDesc.caption) + CCur(txtbxDescForma.Text) & "%", vbExclamation
            Exit Sub
        End If
    End If
    If txtbxProduto.Text = "" Then
        MsgBox "Informe pelo menos parte do nome do produto para que seja localizado!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    On Error GoTo trata_Erro
                
        If cmndbtnInserir.tag = "" Then
            iItem = iItem + 1
        Else
            
        End If
                
                
                If Vendas.F_AdicionaItem(Me.tag, p, txtbxQtd.Text, txtbxDescReais.Text, txtbxFinal.Text, txtbxPQ_pro.Text) = "" Then
                    MsgBox "Erro na inserção do item!", vbExclamation
                Else
                    'If Not F_alteraitem(p.cod_pro, "-" & txtbxQtd.Text) Then
                    '    MsgBox "Erro na baixa do estoque.", vbExclamation
                    'End If
                    
                End If
                    
                P_buscaitens
                Me.Refresh
                
                txtbxQtd.Text = "1"
                dtProd.Recordset.Move dtProd.Recordset.RecordCount - 1
                txtbxUnit.Text = "0"
                txtbxDescPercent.Text = "0,0"
                txtbxDescReais.Text = "0,00"
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
            
        
        P_AtualizaVenda
        On Error Resume Next

        Exit Sub
trata_Erro:
    P_MostraErro "Insere()"
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub txtbxQtd_LostFocus()
    On Error Resume Next
    
    P_ValorLostFocus txtbxQtd, "###"
    If txtbxPQ_pro.Visible = True Then
        txtbxDescReais.Text = FormatVal(CCur(txtbxPQ_pro.tag) * CCur(txtbxQtd.Text))
        txtbxDescReais_LostFocus
    End If
    p_totalItem
    
        
    
End Sub


Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit, "###,##0.00"
    
    p_totalItem
End Sub
Sub P_GravaItemCancelado()
    On Error Resume Next
    Dim strSQL As String
    strSQL = "insert into tab_itvndcncl(seq_itv,num_ven,seq_ite,qtd_ite,unit_ite,cod_cai,cod_tur)values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_itvndcncl", "seq_itv", "", "000000#") & "',"
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "" & dtProd.Recordset!seq_ite & ","
    strSQL = strSQL & F_GravaDec(dtProd.Recordset!qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(dtProd.Recordset!unit_ite) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "') "
    Banco.Execute strSQL
    
    
    
End Sub






