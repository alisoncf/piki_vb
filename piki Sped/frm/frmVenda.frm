VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVenda 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "-Venda de Produtos:"
   ClientHeight    =   10365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14085
   ControlBox      =   0   'False
   Icon            =   "frmVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmVenda.frx":08CA
      Height          =   6615
      Left            =   5880
      OleObjectBlob   =   "frmVenda.frx":08DF
      TabIndex        =   7
      Top             =   840
      Width           =   6495
   End
   Begin VB.PictureBox pcAtalhos 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   14055
      TabIndex        =   62
      Top             =   0
      Width           =   14085
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
         TabIndex        =   84
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
            Picture         =   "frmVenda.frx":1994
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   85
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
            TabIndex        =   86
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
         TabIndex        =   81
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmVenda.frx":1AFF
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   82
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
            TabIndex        =   83
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
         TabIndex        =   78
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
            Picture         =   "frmVenda.frx":1E89
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   79
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
            TabIndex        =   80
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
         TabIndex        =   75
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
            Picture         =   "frmVenda.frx":4250
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   76
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
            TabIndex        =   77
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
         TabIndex        =   72
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
            Picture         =   "frmVenda.frx":4C52
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   73
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
            TabIndex        =   74
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
         TabIndex        =   69
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
            Picture         =   "frmVenda.frx":4D9C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   70
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F3 - Quantidade"
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
            TabIndex        =   71
            Top             =   120
            Width           =   1260
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
         TabIndex        =   66
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
            Picture         =   "frmVenda.frx":5010
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   67
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F1 - TEF"
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
            TabIndex        =   68
            Top             =   120
            Width           =   615
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
         TabIndex        =   63
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
            Picture         =   "frmVenda.frx":539A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   64
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F2 - Multi"
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
            TabIndex        =   65
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   1320
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14025
      TabIndex        =   15
      Top             =   9990
      Width           =   14085
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
         Left            =   6360
         TabIndex        =   25
         Top             =   0
         Width           =   270
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
         Left            =   7320
         TabIndex        =   24
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
         Left            =   5880
         TabIndex        =   23
         Top             =   60
         Width           =   465
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
         Left            =   5400
         TabIndex        =   22
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   21
         Top             =   60
         Width           =   435
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
         Left            =   10350
         TabIndex        =   20
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   60
         Width           =   630
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
         Left            =   3840
         TabIndex        =   18
         Top             =   0
         Width           =   945
      End
      Begin VB.Label lblNome_Fun 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Atendente"
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
         Left            =   840
         TabIndex        =   17
         Top             =   0
         Width           =   2250
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   60
         Width           =   780
      End
   End
   Begin VB.PictureBox pcMsg 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14055
      TabIndex        =   11
      Top             =   9615
      Width           =   14085
      Begin VB.Label lblHora 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         Height          =   195
         Left            =   10680
         TabIndex        =   14
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblMensagem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5745
         TabIndex        =   13
         Top             =   0
         Width           =   135
      End
      Begin VB.Label lblAtacado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda a Atacado"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C6AE86&
         Height          =   390
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   2625
      End
   End
   Begin VB.Data dtProd 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmndbtnLocaliza 
      Caption         =   "..."
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pcECF 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   14085
      TabIndex        =   8
      Top             =   9570
      Width           =   14085
   End
   Begin VB.PictureBox pcBaixa 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   14085
      TabIndex        =   9
      Top             =   9525
      Visible         =   0   'False
      Width           =   14085
   End
   Begin VB.PictureBox pcValor 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8985
      Left            =   0
      ScaleHeight     =   8985
      ScaleWidth      =   5295
      TabIndex        =   26
      Top             =   540
      Width           =   5295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   240
         Picture         =   "frmVenda.frx":5D9C
         ScaleHeight     =   1005
         ScaleWidth      =   2010
         TabIndex        =   87
         Top             =   8520
         Width           =   2010
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   5295
         TabIndex        =   56
         Top             =   1560
         Width           =   5295
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   60
            Top             =   30
            Width           =   855
         End
         Begin VB.Label lblProd 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1200
            TabIndex        =   59
            Tag             =   "!"
            Top             =   0
            Width           =   135
         End
         Begin VB.Label lblItens 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   4440
            TabIndex        =   58
            Tag             =   "!"
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Itens:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3720
            TabIndex        =   57
            Top             =   37
            Width           =   465
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00000000&
            FillColor       =   &H00FF8080&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   5175
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   5295
         TabIndex        =   39
         Top             =   2040
         Width           =   5295
         Begin VB.Label lblCapPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto:"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   53
            Top             =   960
            Width           =   2325
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categoria:"
            Height          =   195
            Left            =   2640
            TabIndex        =   52
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lblCat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   51
            Top             =   960
            Width           =   2325
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   330
         End
         Begin VB.Label lblCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Top             =   1560
            Width           =   2325
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. Barra:"
            Height          =   195
            Left            =   2640
            TabIndex        =   48
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label lblCodBarr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2640
            TabIndex        =   47
            Top             =   1560
            Width           =   2325
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   2640
            TabIndex        =   46
            Top             =   1920
            Width           =   870
         End
         Begin VB.Label lblCapPreco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1920
            Width           =   465
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balança:"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   2580
            Width           =   630
         End
         Begin VB.Label lblBalanca 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   840
            TabIndex        =   43
            Top             =   2520
            Width           =   285
         End
         Begin VB.Label lblRef 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2640
            TabIndex        =   42
            Top             =   2520
            Width           =   2325
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            Height          =   195
            Left            =   1680
            TabIndex        =   41
            Top             =   2580
            Width           =   825
         End
         Begin VB.Label lblNome_pro 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   60
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00000000&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   3015
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   5175
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   5175
         TabIndex        =   36
         Top             =   120
         Width           =   5175
         Begin VB.Label lblExtenso 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2520
            TabIndex        =   61
            Tag             =   "!"
            Top             =   1080
            Width           =   195
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   38
            Top             =   0
            Width           =   600
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
               Size            =   50.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1140
            Left            =   1680
            TabIndex        =   37
            Tag             =   "!"
            Top             =   0
            Width           =   1785
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00000000&
            FillColor       =   &H000000C0&
            FillStyle       =   0  'Solid
            Height          =   1335
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   5175
         End
      End
      Begin VB.TextBox txtbxNomeCli 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtbxCod_Cli 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1080
         TabIndex        =   32
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox pcPromo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3105
         Left            =   120
         ScaleHeight     =   3105
         ScaleWidth      =   5385
         TabIndex        =   27
         Top             =   5280
         Width           =   5385
         Begin VB.Label lblPromo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos em Promoção"
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
            Left            =   1080
            TabIndex        =   31
            Top             =   720
            Width           =   2880
         End
         Begin VB.Label lblPromoPreco 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "R$ 0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   600
            Left            =   1560
            TabIndex        =   30
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos em Promoção"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   345
            Left            =   960
            TabIndex        =   29
            Top             =   240
            Width           =   3120
         End
         Begin VB.Label lblEmp 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-----------"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   28
            Top             =   2400
            Width           =   855
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00000000&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   3015
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   5175
         End
      End
      Begin VB.Label lblPreco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   3600
         Width           =   2325
      End
      Begin VB.Label lblQtd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2640
         TabIndex        =   34
         Top             =   3600
         Width           =   2325
      End
   End
   Begin VB.Frame frProd 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   5280
      TabIndex        =   0
      Top             =   7440
      Width           =   6615
      Begin VB.TextBox txtbxProduto 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
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
         Left            =   5640
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   600
         Width           =   735
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   855
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   5520
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   975
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unit:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   330
      End
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7095
      Left            =   5400
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7815
   End
End
Attribute VB_Name = "frmVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  
  
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
Sub P_Msg(msg As String)
    
    lblMensagem.caption = msg
    Me.Refresh
End Sub

Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    On Error GoTo trata_erro:
    P_Status "Atualizando Valores..."
    
    iPos = dtProd.Recordset.AbsolutePosition
    dtProd.Refresh
    With dtProd.Recordset
        
        
        Do While Not .EOF
            cTotal = cTotal + !total
            cItens = cItens + !qtd_ite
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cTotal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = dtProd.Recordset.RecordCount
        lblExtenso.caption = Extenso(cTotal)
    End With
        
    
    
    
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    
    lblProd.caption = dtProd.Recordset.RecordCount
    lblItens.caption = cItens

    strSQL = "update tab_vendas set valor_ven =" & F_GravaDec(cTotal)
    strSQL = strSQL & ", nomecli_ven= '" & txtbxNomeCli.Text & "', "
    strSQL = strSQL & " cod_cli= '" & txtbxNomeCli.Tag & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "' "
'    strSQL = strSQL & " obs_VEN='" & txtbxObs.Text & "' "
    strSQL = strSQL & " where num_Ven='" & Me.Tag & "' "
    Banco.Execute strSQL
    
    
    P_Status ""
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = "select tv.*,tp.*,tg.nome_cat,tg.nome_cat,"
    strSQL = strSQL & "te.nome_mar,te.cod_mar,ti.*,(ti.unit_ite*ti.qtd_ite) as total "
    strSQL = strSQL & " from tab_produtos as tp,tab_vendas as tv, tab_categorias as tg,"
    strSQL = strSQL & " tab_marcas as te, tab_Itvendas as ti   "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and tp.cod_cat=tg.cod_cat "
    
    strSQL = strSQL & " and tp.cod_mar=te.cod_mar "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven and tv.num_ven='" & Me.Tag & "' "
    strSQL = strSQL & " order by seq_ite "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub buscaCupomAberto()
    Dim i As Integer
    Dim strSQL As String
    i = F_RetornoVariaveis()
    
    If i = 1 Or i = 33 Then
        If MsgBox("Deseja continuar no cupom fiscal aberto?", vbYesNo + vbQuestion) = vbYes Then
            Me.operacao = "C"
            pcECF.BackColor = vbGreen
            strSQL = MaxKeySQL("tab_itvendas", "seq_ite", "num_ven='" & Me.Tag & "'", "")
            iItem = Int(strSQL)
        Else
            Me.operacao = "B"
            pcECF.BackColor = vbYellow
             ecf_CancelaCupom
        End If
    End If
End Sub


Sub P_InsereProd(sCod_Pro As String, cUnit As Currency, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    On Error GoTo trata_erro
    strSQL = F_BuscaUmCampo("max(seq_ite)", "tab_itvendas", "num_ven='" & Me.Tag & "'")
'    If strSQL = "" Then
'        iIndex = 1
'    Else
'        iIndex = val(strSQL) + 1
'    End If
    strSQL = "insert into tab_itvendas(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite,final_ite)values("
    
    strSQL = strSQL & "'" & Me.Tag & "', "
    strSQL = strSQL & "'" & sCod_Pro & "',"
    strSQL = strSQL & F_GravaDec(CStr(cUnit)) & ","
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & ","
    strSQL = strSQL & iItem & ","
    
    strSQL = strSQL & F_GravaDec(cUnit * cQtd) & ") "
    Banco.Execute strSQL
    'RESERVANDO O ESTOQUE
    If pcBaixa.Visible = False Then  'baixa se não estiver visivel
        strSQL = "update tab_produtos set er_pro=er_pro - " & F_GravaDec(CStr(cQtd))
        strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
        Banco.Execute strSQL
    Else
        
    End If
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub P_AumentaQTD(sCod_Pro As String, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    
    strSQL = "Update tab_itvendas set "
    
    strSQL = strSQL & "qtd_ite = qtd_ite+"
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & " "
    strSQL = strSQL & " where num_ven='" & Me.Tag & "'  "
    strSQL = strSQL & " and cod_pro ='" & sCod_Pro & "' "
    Banco.Execute strSQL
    
    
    
    'RESERVANDO O ESTOQUE
    strSQL = "update tab_produtos set er_pro=er_pro - " & F_GravaDec(CStr(cQtd))
    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
    Banco.Execute strSQL
    
End Sub



Private Sub cmbxBuscar_Click()
    If cmbxBuscar.ListIndex = 0 Then
       cmndbtnLocaliza.Tag = "B"
    Else
        cmndbtnLocaliza.Tag = Left(cmbxBuscar.Text, 1)
    End If
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
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
    If MsgBox("Tem certeza que deseja cancelar esta venda?", vbYesNo + vbCritical) = vbYes Then
        If operacao = "C" Then
            i = F_RetornoVariaveis
            If i Mod 2 > 0 Then
                ecf_CancelaCupom
                iItem = 0
            End If
        End If
        With dtProd.Recordset
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                excluir False
                .MoveNext
            Loop
        End With
        strSQL = "insert into tab_cancelamentos(num_ven,data_can,valor_ven,tipo_ven,"
        strSQL = strSQL & "cod_cai,cod_tur)values("
        strSQL = strSQL & "'" & Me.Tag & "',"
        strSQL = strSQL & F_GravaDataTempo(Now) & ","
        strSQL = strSQL & F_GravaDec(lblValor.caption) & ","
        strSQL = strSQL & "'" & operacao & "',"
        strSQL = strSQL & "'" & strCod_Cai & "',"
        strSQL = strSQL & "'" & strCod_Tur & "') "
        Banco.Execute strSQL
        
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.Tag & "' "
        Banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.Tag & "' "
        Banco.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
            RefreshDataPosition frmEdiVendas.dtVendas
        End If
        
        If MsgBox("Deseja Gerar uma Nova Venda?" & Chr(13) & "Clique em não caso queira abandonar o modo de vendas!", vbYesNo + vbQuestion) = vbYes Then
            strSQL = F_GeraNovaVenda("00000", strCod_Usu)
            iItem = 0
            Me.Tag = strSQL
            Form_Activate
        Else
            iItem = 0
            Unload Me
        End If
    End If
End Sub



Private Sub cmndbtnEditProd_Click()
    On Error GoTo trata_erro
    Dim iPos As Integer
    Dim strSQL As String
    Dim cUlt_Qtd As Currency
    
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        iPos = .AbsolutePosition
        frmValoresProd.lblCod_Pro.caption = !cod_pro
        frmValoresProd.lblNome_pro.caption = !nome_pro
        frmValoresProd.txtbxCusto.Text = FormatVal(!custo_pro)
        frmValoresProd.txtbxPreco.Text = FormatVal(!unit_ite)
        frmValoresProd.txtbxQtd.Text = FormatVal(!qtd_ite)
        frmValoresProd.Tag = Me.Tag
        cUlt_Qtd = !qtd_ite
        On Error Resume Next
        frmValoresProd.Show 1
        If CCur(frmValoresProd.txtbxQtd.Text) <= 0 Then
            MsgBox "Informe um valor decente!", vbExclamation
            Exit Sub
        End If
        strSQL = "Update tab_itvendas set qtd_ite=" & _
        F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & ", unit_ite=" & _
        F_GravaDec(frmValoresProd.txtbxPreco.Text)
        
        strSQL = strSQL & " where num_ven='" & Me.Tag & "' "
        strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
        Banco.Execute strSQL
        'ajustando estoque reservado
        strSQL = "Update tab_produtos set reservado_pro=reservado_pro-" & F_GravaDec(CStr(cUlt_Qtd))
        strSQL = strSQL & "+" & F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        Banco.Execute strSQL
        Unload frmValoresProd
    End With
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    P_AtualizaVenda
trata_erro:
MsgBox "Err: cmndbtnEditProd_Click() "
End Sub

Private Sub cmndbtnExc_Prod_Click()
    excluir True
End Sub
Sub excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & !nome_pro & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        iPos = .AbsolutePosition
        strSQL = "delete from tab_itvendas where seq_ite='" & !seq_ite & "' "
        
        strSQL = strSQL & " and num_ven='" & Me.Tag & "' "
        Banco.Execute strSQL
        P_GravaItemCancelado
        'ajustando estoque reservado
        If pcBaixa.Visible = False Then
            strSQL = "update tab_produtos set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))
            
            If operacao = "C" Then
                strSQL = strSQL & ", ef_pro=ef_pro + " & F_GravaDec(CStr(!qtd_ite))
            End If
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Banco.Execute strSQL
        End If
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
trata_erro:
    P_MostraErro "Excluir()"
End Sub


Private Sub cmndbtnCliente_Click()
     
        frmSelectCli.txtbxNOme.Tag = "C"
        frmSelectCli.txtbxNOme.Text = txtbxNomeCli.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    
    If F_BuscaUmCampo("status_cli", "tab_clientes", "cod_cli='" & _
        frmSelectCli.dtBusca.Recordset!COD & "'") <> "1" Then
        MsgBox "Este cliente está negativado! " & Chr(13) & _
            "A venda só poderá ser fechada a vista ou a cartão. ", vbCritical
    End If
    
    txtbxNomeCli.Text = frmSelectCli.dtBusca.Recordset!Nome
    txtbxNomeCli.Tag = frmSelectCli.dtBusca.Recordset!COD
    txtbxCod_Cli.Text = frmSelectCli.dtBusca.Recordset!COD
    
    
    
    Unload frmSelectCli
    
End Sub

Private Sub cmndbtnLocaliza_Click()
    
    If F_Vazio(txtbxProduto.Text) Then
        MsgBox "Informe o produto que deseja inserir!", vbExclamation, constNome_sis
        txtbxProduto.SetFocus
        Exit Sub
    End If
    
    insere
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    
    If dtProd.Recordset.RecordCount > 0 Then
        If lblAtacado.Tag <> "" Then
            If MsgBox("Confirma Saída da Venda?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "Impossível Abandonar a Venda!", vbCritical
            Exit Sub
        End If
    End If
    If IsLoad("frmEdiVendas") Then
        frmEdiVendas.dtVendas.Refresh
    End If
    If dtProd.Recordset.RecordCount = 0 Then
        strSQL = "delete from tab_vendas where num_ven='" & Me.Tag & "' "
        Banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.Tag & "' "
        Banco.Execute strSQL
        
    End If
    i = F_RetornoVariaveis
        If i Mod 2 > 0 Then
            ecf_CancelaCupom
        End If
    
    Unload Me
    Exit Sub
trata_erro:
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
    
    
    strSQL = "update tab_itvendas set final_ite=(qtd_ite*unit_ite) "
    strSQL = strSQL & " where num_ven='" & Me.Tag & "' "
    Banco.Execute strSQL
    
    sTroco = F_Troco(CCur(lblValor.caption), Me.Tag, sVen, txtbxNomeCli.Tag, txtbxNomeCli.Text, operacao)
    If sTroco <> "" Then
        If Atacado = True Then
            
                ImprVenda Me.Tag
            
        End If
        VerificaAtualizacao
        strSQL = F_GeraNovaVenda("00000", strCod_Usu)
        iItem = 0
        Me.Tag = strSQL
        lblnum_ven.caption = strSQL
        P_buscaitens
        P_AtualizaVenda
        On Error Resume Next
        frmVenda.SetFocus
        txtbxProduto.Text = ""
        pcBaixa.Visible = False
        On Error Resume Next
        txtbxProduto.SetFocus
        dtProd.Refresh
    End If
End Sub


Private Sub dbgrdProd_DblClick()
    cmndbtnEditProd_Click
End Sub

Private Sub dbgrdProd_GotFocus()
    
    dbgrdProd.MarqueeStyle = 3
    
    
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo trata_erro
    Dim stSQL As String
    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        strSQL = F_senhaReturnUsu()
        
        If strSQL = "" Then
            MsgBox "Erro de autenticação!"
            Exit Sub
        Else
            strSQL = F_BuscaUmCampo("exclvenda_usu", "tab_defusu", "cod_usu='" & strSQL & "'")
            If strSQL <> "S" Then
                MsgBox "Você não tem permissão para cancelar o item. Chame o responsável!", vbExclamation
                Exit Sub
                
            End If
        End If
        
        If operacao = "C" Then
            If dtProd.Recordset.RecordCount = 1 Then
                i = ecf_CancelaCupom
                If i <> 1 Then
                        MsgBox "Ocorreu um erro na tentativa de exclusão. Tente novamente!", vbCritical
                        P_AtualizaVenda
                        Exit Sub
                End If
                iItem = 0
            Else
                
                    If dbgrdProd.Row = -1 Then
                        MsgBox "Selecione o item que deseja excluir!", vbCritical
                        dbgrdProd.SetFocus
                        Exit Sub
                    End If
                    i = ecf_CancelaItemGenerico(dtProd.Recordset!seq_ite)
                    If i <> 1 Then
                        MsgBox "Ocorreu um erro na tentativa de exclusão. Tente novamente!", vbCritical
                        P_AtualizaVenda
                        Exit Sub
                    
                End If
            End If
            excluir False
            
        Else
            excluir False
        End If
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxProduto.SetFocus
    End If
    P_AtualizaVenda
    Exit Sub
trata_erro:
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
        lblProduto.caption = !nome_pro
            
            
                lblCapPreco.caption = "Preço:"
                lblPreco.caption = "R$ " & FormatVal(!pV_pro)
            
            lblCat.caption = !nome_cat
            lblCod.caption = FormatCodigo(!cod_pro)
            lblCodBarr.caption = F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'")
            lblMarca.caption = !nome_mar
            lblQtd.caption = !qtd_ite
            lblRef.caption = !codfab_pro
            lblBalanca.caption = !balanca_pro
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
            pPromo(.AbsolutePosition).Preco = "R$ " & FormatVal(!pV_pro)
            
            .MoveNext
        Loop
        .Close
        
        
        
    
    End With
    
End Sub
Private Sub Form_Activate()
    On Error GoTo trata_erro
    mdiFrm.Enabled = False
    lblEmp.caption = strFantasia_Emp
    lblnum_ven.caption = Me.Tag
    lblNome_Fun.caption = strCod_Usu & "-" & strNome_Usu
    lblNomeEmp.caption = strFantasia_Emp
    lblCaixa = strCod_Cai
    lblTur.caption = strCod_Tur
    If bActive = True Then Exit Sub
    
    If Me.Atacado = True Then
        pcPromo.Visible = False
        Timer1.Enabled = False
        
    Else
        pcPromo.Visible = True
        Timer1.Enabled = True
        
    End If
    P_Status "Buscando Itens..."
    txtbxProduto.SetFocus
    P_buscaitens
    If dtProd.Recordset.RecordCount > 0 Then
        buscaCupomAberto
    Else
        If Atacado Then
            lblAtacado.Visible = True
            Form_KeyDown vbKeyF9, 0
        Else
            lblAtacado.Visible = False
        End If
    End If
    P_AtualizaVenda
    P_Status ""
    
    bActive = True
    If strCod_Cai = "" Then
        MsgBox "Não existe caixa para este computador, o caixa ficará para o '01' ", vbExclamation
        strCod_Cai = "00"
    End If
    playGif
    Exit Sub
trata_erro:
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
    ElseIf KeyCode = vbKeyF1 Then
        P_AdmTEF
    
    ElseIf KeyCode = vbKeyF5 Then
        
        dbgrdProd.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
         dbgrdProd_KeyDown vbKeyDelete, 0
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF2 Then
        frmRecebConta.Show 1
    ElseIf KeyCode = vbKeyF8 Then
        If operacao = "B" And dtProd.Recordset.RecordCount > 0 Then
            MsgBox "Operacação Impossível!", vbCritical
            
            Exit Sub
        End If
        
        pcECF.BackColor = vbGreen
        operacao = "C"
        mudamodo
    ElseIf KeyCode = vbKeyF9 Then
        If operacao = "C" And dtProd.Recordset.RecordCount > 0 Then
            MsgBox "Operacação Impossível!", vbCritical
            Exit Sub
        End If
        If lblBloqueado.caption = "S" Then
            Exit Sub
        End If
        pcECF.BackColor = vbYellow
        operacao = "B"
        mudamodo
    ElseIf KeyCode = vbKeyF10 Then
        If lblBloqueado.caption = "S" Then
            Exit Sub
        End If
        
        If dtProd.Recordset.RecordCount > 0 Then
            MsgBox "Operação impossível. A Venda já possui itens!", vbCritical
            Exit Sub
        End If
        If pcBaixa.Visible = True Then
            pcBaixa.Visible = False
        Else
            pcBaixa.Visible = True
        End If
    
        mudamodo
    ElseIf KeyCode = vbKeyF11 Then
        frmReemissao.Show 1
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        frmECF.Show 1
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
    buscaPromo
    bActive = False

    dtProd.Connect = Banco.Connect

    lblNomeEmp.caption = strNome_Emp
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", "tab_parametros", "")
    txtbxNomeCli.Tag = "00000"
    txtbxNomeCli.Text = "VENDA A CONSUMIDOR"
    
    
    
    
'    cmbxBuscar.ListIndex = 0
    pcECF.BackColor = vbGreen
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
'    frControle.Width = Me.Width - 10
    
    pcStatusBar.Top = Me.Height - pcStatusBar.Height - 50
'    pcStatusBar.Width = Me.Width - 100
 
    pcMsg.Top = pcStatusBar.Top - pcMsg.Height - 10
'    pcMsg.Width = pcStatusBar.Width
    
    
'    frControle.Height = pcStatusBar.Top - 10
'    pcAtalhos.Width = Me.Width - 10
    lblNomeEmp.Left = pcStatusBar.Width - lblNomeEmp.Width - 50
    lblHora.Left = pcStatusBar.Width - lblHora.Width - 50
    frProd.Top = pcMsg.Top - frProd.Height - 10
'    frProd.Left = frCli.Width + 200
    frProd.Width = Me.Width - frProd.Left - 100
    dbgrdProd.Width = frProd.Width
    dbgrdProd.Height = frProd.Top - 1000
'    frCli.Height = pcMsg.Top - frCli.Top - 150
'    pcPromo.Top = frCli.Height - pcPromo.Height - 10
    frProd.Top = dbgrdProd.Top + 20 + dbgrdProd.Height
    frProd.Left = dbgrdProd.Left
    pcECF.Top = dbgrdProd.Height + dbgrdProd.Top + 50
    pcBaixa.Top = pcECF.Top + pcECF.Height + 50
'    pcECF.Width = Me.Width - 300
'    pcBaixa.Width = pcECF.Width
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).Width = pcAtalhos.Width / (pcBotao.count)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).Width + 50
        End If
        
    Next
    
End Sub

Sub editaItem()
    txtbxQtd.Enabled = True
    txtbxQtd.Text = "1"
    txtbxQtd.SetFocus
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



Private Sub Form_Terminate()
    On Error Resume Next
    Timer1.Enabled = False
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mdiFrm.Enabled = True
    dtProd.Recordset.Close
    
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

Private Sub lblTef_Click()
    P_AdmTEF
End Sub

Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
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
        
        
    ElseIf Index = 0 Then

    ElseIf Index = 7 Then
        
    End If
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

Private Sub pcMsg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub pcPromo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    
iGif = iGif + 1
If iGif >= 15 Then
    If iIndPromo >= QtdPromo - 1 Then
        buscaPromo
        iIndPromo = 0
    End If
    lblPromo.caption = QuebraLinha(StrConv(pPromo(iIndPromo).Nome, vbProperCase), 30)
    lblPromoPreco.caption = pPromo(iIndPromo).Preco
    iIndPromo = iIndPromo + 1
    iGif = 0
    
End If
    lblBloqueado.caption = F_BuscaUmCampo("bloqueado_cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    
    If lblBloqueado.caption = "S" Then
        If dtProd.Recordset.RecordCount = 0 Then
            
        End If
    End If
    
    
    iDescancoTela = iDescancoTela + 1
    
    If iDescancoTela = 100 Then
        On Error Resume Next
'        frmPropa.Show 0
'        frmPropa.SetFocus
        
        iDescancoTela = 0
    End If
    
    lblHora.caption = Now
    lblHora.Refresh
   End Sub

Private Sub txtbxCod_Cli_GotFocus()
    P_SelecTexto txtbxCod_Cli
End Sub

Private Sub txtbxCod_Cli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
        
    End If
End Sub
Sub buscaCodCli()
    Dim strSQL As String
    
    strSQL = F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & _
        txtbxCod_Cli.Text & "'")
    
    If strSQL = "" Then
        MsgBox "Cliente não encontrado!", vbExclamation
        txtbxCod_Cli.Text = ""
        txtbxCod_Cli.SetFocus
    Else
        txtbxNomeCli.Text = strSQL
        txtbxNomeCli.Tag = txtbxCod_Cli.Text

    End If
    
    
End Sub
Private Sub txtbxCod_Cli_LostFocus()
    If txtbxCod_Cli.Text = "" Then Exit Sub
    P_AjustaTexto txtbxCod_Cli
    txtbxCod_Cli.Text = Format(txtbxCod_Cli.Text, "0000#")
    
    buscaCodCli
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
    P_AjustaTexto txtbxNomeCli
    If txtbxNomeCli.Tag <> "" Then
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

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    Dim p As Produto
    Dim rPesq As Recordset
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click

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
    Dim p As Produto
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTipoQtd As String
    
    On Error GoTo trata_erro
            p = F_BuscaProdSaida(txtbxProduto.Text, "B")
            If p.cod_pro <> "" Then
                txtbxProduto.Text = p.nome_pro
                lblNome_pro.caption = p.nome_pro
                lblBalanca.caption = p.balanca_pro
                If p.balanca_pro <> "S" Then
                    If CCur(txtbxQtd.Text) - Int(txtbxQtd.Text) <> 0 Then
                        strSQL = "A quantidade informada para o produto: " & p.nome_pro & Chr(13)
                        strSQL = strSQL & "não pôde ser lançada, pois o mesmo não se trata de "
                        strSQL = strSQL & " um produto de balança!" & Chr(13) & Chr(13)
                        strSQL = strSQL & "Qualquer dúvida, chame o Responsável!"
                        MsgBox strSQL, vbExclamation, "Incompatibilidade de Tipos"
                        txtbxQtd.Text = Int(txtbxQtd.Text)
                        txtbxQtd.Enabled = True
                        txtbxQtd.SetFocus
                        Exit Sub
                    End If
                Else
                    If p.totalbalanca_pro <> 0 Then
                        If strBALANCA = "V" Then 'valor
                            txtbxQtd.Text = p.totalbalanca_pro / p.p2_pro
                        Else
                            txtbxQtd.Text = Round(p.totalbalanca_pro / p.p2_pro, 3)
                        End If
                        txtbxQtd_LostFocus
                        txtbxUnit.Text = p.p2_pro
                    End If
                    
                End If
                                
                If p.p2_pro = 0 Then
                    MsgBox "Este produto encontra-se com o preço de venda zerado!", vbCritical
                    Exit Sub
                End If
                lblNome_pro.caption = p.nome_pro
                             
                iItem = iItem + 1
                P_InsereProd p.cod_pro, p.p2_pro, CCur(txtbxQtd.Text)
                    
                P_buscaitens
                Me.Refresh
                If operacao = "C" Then
                    i = F_RetornoVariaveis()
                    If i = 4 Or i = 36 Or i = 0 Or i = 32 Then
                        strNum_Cpm = Space(6)
                        
                        i = ecf_NumeroCupom(strNum_Cpm)
                         P_Msg "Abrindo Cupom Fiscal..."
                        
                        i = ecf_AbreCupom("")
                        P_Msg ""
                    Else
                        'MsgBox "Impressora fiscal não localizada!", vbExclamation
                    End If
                    P_Msg "Registrando Item... " & p.nome_pro
                    
                    
                    
                    
                    If p.s_pro = "R" Then
                        p.ecf_icms = "FF"
                    ElseIf p.s_pro = "F" Then
                        p.ecf_icms = "FF"
                    ElseIf p.s_pro = "I" Then
                        p.ecf_icms = "II"
                    ElseIf p.s_pro = "" Then
                        p.ecf_icms = "01"
                    End If
                    If Int(txtbxQtd.Text) <> CCur(txtbxQtd.Text) Then
                        sTipoQtd = "F"
                    Else
                        sTipoQtd = "I"
                    End If
                    i = ecf_VendeItem(p.cod_pro, Left(p.nome_pro, 29), p.ecf_icms, sTipoQtd, txtbxQtd.Text, 3, Format(p.p2_pro, "###,##0.000"), "%", "0")
                    If i <> 1 Then
                        MsgBox "Erro na passagem do item para a impressora fiscal!", vbCritical
                    End If
                    ecf_verificaImpressora
                    P_Msg ""
                End If
                txtbxQtd.Text = "1"
                dtProd.Recordset.Move dtProd.Recordset.RecordCount - 1
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
            End If
        
        P_AtualizaVenda
        On Error Resume Next
        rPesq.Close
        Exit Sub
trata_erro:
    P_MostraErro "Insere()"
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        If CCur(txtbxQtd.Text) = 0 Then
            MsgBox "O Valor Informado não está Correto!", vbCritical
            txtbxQtd.SetFocus
            Exit Sub
        End If
        txtbxQtd.Enabled = False
        txtbxProduto.SetFocus
    End If
Exit Sub
trata_erro:
    P_MostraErro
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
    txtbxQtd.Enabled = False
    txtbxProduto.Text = ""
    txtbxProduto.Enabled = True
    txtbxProduto.SetFocus
    
    
    
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    Timer1_Timer
End Sub

Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit, "###,##0.000"
End Sub
Sub P_GravaItemCancelado()
    On Error Resume Next
    strSQL = "insert into tab_itvndcncl(seq_itv,num_ven,seq_ite,qtd_ite,unit_ite,cod_cai,cod_tur,data_ven)values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_itvndcncl", "seq_itv", "", "000000#") & "',"
    strSQL = strSQL & "'" & Me.Tag & "',"
    strSQL = strSQL & "" & dtProd.Recordset!seq_ite & ","
    strSQL = strSQL & F_GravaDec(dtProd.Recordset!qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(dtProd.Recordset!unit_ite) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "', "
    strSQL = strSQL & F_GravaDataTempo(Now) & ")"
    Banco.Execute strSQL
    
    
    
End Sub
