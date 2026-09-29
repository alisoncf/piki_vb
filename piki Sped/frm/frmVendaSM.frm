VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVendaSM 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "-Venda de Produtos:"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   ControlBox      =   0   'False
   Icon            =   "frmVendaSM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12000
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "frmVendaSM.frx":08CA
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   8700
      Left            =   5280
      OLEDropMode     =   1  'Manual
      Picture         =   "frmVendaSM.frx":0C54
      ScaleHeight     =   8700
      ScaleWidth      =   9000
      TabIndex        =   53
      Top             =   600
      Width           =   9000
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   7935
         Left            =   360
         TabIndex        =   57
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   13996
         _Version        =   393216
         BackColorBkg    =   15132392
         AllowBigSelection=   0   'False
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   5400
      Picture         =   "frmVendaSM.frx":A46B
      ScaleHeight     =   1500
      ScaleWidth      =   9000
      TabIndex        =   46
      Top             =   9240
      Width           =   9000
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   480
         TabIndex        =   48
         Text            =   "1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtbxProduto 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
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
         Left            =   1800
         TabIndex        =   47
         Text            =   "Informe o nome do produto"
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label lblProdutoAtacado 
         Caption         =   "Produto em atacado"
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
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Tag             =   "N"
         Top             =   1440
         Visible         =   0   'False
         Width           =   8775
      End
      Begin VB.Label lblUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7440
         TabIndex        =   52
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unit:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7440
         TabIndex        =   51
         Top             =   120
         Width           =   330
      End
      Begin VB.Label lblUltItem 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         TabIndex        =   50
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd(F3):"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   49
         Top             =   120
         Width           =   570
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
      ScaleWidth      =   14370
      TabIndex        =   21
      Top             =   0
      Width           =   14400
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
         TabIndex        =   43
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
            Picture         =   "frmVendaSM.frx":12C56
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   44
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
            TabIndex        =   45
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
         TabIndex        =   40
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmVendaSM.frx":12DC1
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   41
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
            TabIndex        =   42
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
         TabIndex        =   37
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
            Picture         =   "frmVendaSM.frx":1314B
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
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
            TabIndex        =   39
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
         TabIndex        =   34
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
            Picture         =   "frmVendaSM.frx":15512
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   35
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
            TabIndex        =   36
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
         TabIndex        =   31
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
            Picture         =   "frmVendaSM.frx":15F14
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   32
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
            TabIndex        =   33
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
         TabIndex        =   28
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
            Picture         =   "frmVendaSM.frx":1605E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   29
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
            TabIndex        =   30
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
         TabIndex        =   25
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
            Picture         =   "frmVendaSM.frx":162D2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   26
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
            TabIndex        =   27
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
         TabIndex        =   22
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   120
            Picture         =   "frmVendaSM.frx":1665C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   23
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F2 - CPF"
            Enabled         =   0   'False
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
            TabIndex        =   24
            Top             =   120
            Width           =   645
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   2040
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14370
      TabIndex        =   4
      Top             =   11625
      Width           =   14400
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
         Left            =   6720
         TabIndex        =   14
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
         Left            =   7560
         TabIndex        =   13
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
         Left            =   6120
         TabIndex        =   12
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
         Left            =   5640
         TabIndex        =   11
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         TabIndex        =   8
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
         Left            =   4080
         TabIndex        =   7
         Top             =   0
         Width           =   945
      End
      Begin VB.Label lblNome_Fun 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   60
         Width           =   780
      End
   End
   Begin VB.PictureBox pcMsg 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ScaleWidth      =   14370
      TabIndex        =   2
      Top             =   11250
      Width           =   14400
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
         TabIndex        =   3
         Top             =   0
         Width           =   135
      End
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
      ScaleWidth      =   14400
      TabIndex        =   0
      Top             =   11205
      Width           =   14400
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
      ScaleWidth      =   14400
      TabIndex        =   1
      Top             =   11160
      Visible         =   0   'False
      Width           =   14400
   End
   Begin VB.PictureBox pcValor 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10620
      Left            =   0
      ScaleHeight     =   10620
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   540
      Width           =   5535
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4500
         Left            =   360
         Picture         =   "frmVendaSM.frx":1705E
         ScaleHeight     =   4500
         ScaleWidth      =   4500
         TabIndex        =   58
         Top             =   3240
         Width           =   4500
         Begin VB.Label lblAtacado 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Item em Acadado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1440
            TabIndex        =   69
            Tag             =   "N"
            Top             =   3840
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label lblProduto 
            BackStyle       =   0  'Transparent
            Caption         =   "-------------->"
            Height          =   675
            Left            =   1440
            TabIndex        =   68
            Top             =   2400
            Width           =   2205
         End
         Begin VB.Label lblProd 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2640
            TabIndex        =   67
            Top             =   1320
            Width           =   135
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos"
            Height          =   195
            Left            =   1440
            TabIndex        =   66
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label lblItens 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2640
            TabIndex        =   65
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Itens"
            Height          =   195
            Left            =   1440
            TabIndex        =   64
            Top             =   1680
            Width           =   345
         End
         Begin VB.Label lblHora 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Não Informada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   720
            TabIndex        =   63
            Top             =   4200
            Width           =   1680
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF:"
            Height          =   255
            Left            =   1440
            TabIndex        =   62
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label lblCPF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   61
            Top             =   3120
            Width           =   45
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selecionado"
            Height          =   195
            Left            =   1440
            TabIndex        =   60
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label lblItemSelec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2640
            TabIndex        =   59
            Top             =   2040
            Width           =   135
         End
      End
      Begin VB.PictureBox pcLogo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1680
         Left            =   -120
         Picture         =   "frmVendaSM.frx":1BD17
         ScaleHeight     =   1680
         ScaleWidth      =   5520
         TabIndex        =   54
         Top             =   8520
         Width           =   5520
      End
      Begin VB.PictureBox pcTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         Picture         =   "frmVendaSM.frx":24735
         ScaleHeight     =   1935
         ScaleWidth      =   5175
         TabIndex        =   18
         Top             =   120
         Width           =   5175
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
            Left            =   0
            TabIndex        =   20
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
            TabIndex        =   19
            Tag             =   "!"
            Top             =   360
            Width           =   1785
         End
         Begin VB.Label lblValor2 
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
            ForeColor       =   &H00000000&
            Height          =   1140
            Left            =   1710
            TabIndex        =   55
            Tag             =   "!"
            Top             =   380
            Width           =   1785
         End
      End
      Begin VB.TextBox txtbxNomeCli 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   3360
         TabIndex        =   17
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtbxCod_Cli 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   0
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblAvisoAtacado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atenção. Vendas em atacado apenas pagamento em dinheiro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   360
         TabIndex        =   70
         Top             =   2160
         Visible         =   0   'False
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmVendaSM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado  As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1

Dim xAux As Single
Dim yAux As Single
  
Dim seq_ite As String
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
Dim ProdutoSelecionado As Produtos.Produto



Sub P_Msg(msg As String)
    
    lblMensagem.caption = msg
    Me.Refresh
End Sub

Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    On Error GoTo trata_Erro:
    lblItemSelec.caption = ""
    P_Status "Atualizando Valores..."
    
    iPos = ado.AbsolutePosition
    ado.Requery
    With ado
        
        
        Do While Not .EOF
            cTotal = cTotal + !total
            cItens = cItens + !qtd_ite
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cTotal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        If .AbsolutePosition > -1 Then
            lblItens.caption = ado.RecordCount
        End If
        lblValor2.caption = lblValor.caption
        'lblExtenso.caption = Extenso(cTotal)
    End With
        
    
    
    
    ado.Requery
    If iPos > -1 Then
        ado.Move iPos
        lblProd.caption = ado.RecordCount
    End If
    
    lblItens.caption = cItens

    strSQL = "update " & TAB_VENDAS & " set valor_ven =" & F_GravaDec(cTotal)
    strSQL = strSQL & ", nomecli_ven= '" & txtbxNomeCli.Text & "', "
    strSQL = strSQL & " cod_cli= '" & txtbxNomeCli.tag & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "' "
'    strSQL = strSQL & " obs_VEN='" & txtbxObs.Text & "' "
    strSQL = strSQL & " where num_Ven='" & Me.tag & "' "
    conn.Execute strSQL
    
    
    P_Status ""
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select tv.*,tp.*,tg.nome_cat,tg.nome_cat,ti.*,"
    strSQL = strSQL & "te.nome_mar,te.cod_mar,ti.*,(ti.unit_ite*ti.qtd_ite) as total "
    strSQL = strSQL & " from tab_produtos as tp," & TAB_VENDAS & " as tv, tab_categorias as tg,"
    strSQL = strSQL & " tab_marcas as te, tab_Itvendas as ti   "
    strSQL = strSQL & " where tp.cod_pro=ti.cod_pro and tp.cod_cat=tg.cod_cat "
    
    strSQL = strSQL & " and tp.cod_mar=te.cod_mar "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven and tv.num_ven='" & Me.tag & "' "
    strSQL = strSQL & " order by seq_ite asc "
    
    abreObjetoAdo ado, strSQL
    ajustaGrid
    lblAvisoAtacado.Visible = False
    With ado
        Do While Not .EOF
            
            lblCPF.caption = Format(!CPF)
            
            strSQL = .AbsolutePosition & Chr(9) & !cod_pro & Chr(9) & !qtd_ite & Chr(9) & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite * !qtd_ite) & Chr(9)
            strSQL = strSQL & !descricao_ite & Chr(9)
            strSQL = strSQL & !Atacado & Chr(9)
            strSQL = strSQL & !seq_ite
            grid.AddItem strSQL
            grid.Row = grid.rows - 1
            If grid.Row Mod 2 = 0 Then
                
                ZebraLinha grid, grid.rows - 1, &HB8D6D8
                
                         
            End If
            If !Atacado = "S" Then
                ZebraLinha grid, grid.rows - 1, &HA8D2D8, vbBlue, True
                lblAvisoAtacado.Alignment = 2
                lblAvisoAtacado.Visible = True
            End If
            
            .MoveNext
        Loop
    End With
If grid.rows > 1 Then
    grid.FixedRows = 1
    grid.TopRow = grid.rows - 1
End If

'    If ado.RecordCount > 0 Then
'        ado.MoveLast
'    End If
'    dtgrd.FirstRow = ado.RecordCount
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub ajustaGrid()
    With grid
        
        .Cols = 8
        .rows = 1
        '.FixedRows = 1
        '.FixedCols = 1
        .col = 0: .Text = "...": .ColWidth(0) = 500
        .col = 1: .Text = "Código": .ColWidth(1) = 0
        .col = 2: .Text = "Qtd": .ColWidth(2) = 1000
        .col = 3: .Text = "Unitário": .ColWidth(3) = 1000
        .col = 4: .Text = "Total": .ColWidth(4) = 1200
        .col = 5: .Text = "Descrição": .ColWidth(5) = 4500
        .col = 6: .Text = "Atcd": .ColWidth(6) = 0
        .col = 7: .Text = "Seq": .ColWidth(7) = 0
        .ColAlignment(2) = 3
        .ColAlignment(3) = 3
        .ColAlignment(4) = 3
    End With
End Sub
Sub buscaCupomAberto()
    Dim i As Integer
    Dim strSQL As String
    i = F_RetornoVariaveis()
    
    If i = 1 Or i = 33 Then
        If MsgBox("Deseja continuar no cupom fiscal aberto?", vbYesNo + vbQuestion) = vbYes Then
            Me.operacao = "C"
            pcECF.BackColor = vbGreen
            strSQL = MaxKeySQL("" & TAB_ITVENDAS & "", "seq_ite", "num_ven='" & Me.tag & "'", "")
            iItem = Int(strSQL)
        Else
            Me.operacao = "B"
            pcECF.BackColor = vbYellow
             ecf_CancelaCupom
        End If
    End If
End Sub


Sub P_InsereProd(sCod_Pro As String, sDescricao As String, cUnit As Currency, cQtd As Currency, Atacado As String)
    Dim strSQL As String
    Dim iIndex As Integer
    On Error GoTo trata_Erro
    strSQL = F_BuscaUmCampo("max(seq_ite)", "" & TAB_ITVENDAS & "", "num_ven='" & Me.tag & "'")
'    If strSQL = "" Then
'        iIndex = 1
'    Else
'        iIndex = val(strSQL) + 1
'    End If
    strSQL = "insert into " & TAB_ITVENDAS & "(num_ven,cod_pro,descricao_ite,unit_ite,qtd_ite,seq_ite,final_ite,atacado)values("
    
    strSQL = strSQL & "'" & Me.tag & "', "
    strSQL = strSQL & "'" & sCod_Pro & "',"
    strSQL = strSQL & "'" & sDescricao & "',"
    strSQL = strSQL & F_GravaDec(CStr(cUnit)) & ","
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & ","
    strSQL = strSQL & iItem & ","
    
    strSQL = strSQL & F_GravaDec(cUnit * cQtd) & ", "
    strSQL = strSQL & "'" & Atacado & "')"
    conn.Execute strSQL
    'RESERVANDO O ESTOQUE
    If pcBaixa.Visible = False Then  'baixa se não estiver visivel
        strSQL = "update tab_produtos set er_pro=er_pro - " & F_GravaDec(CStr(cQtd))
        strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
        conn.Execute strSQL
    Else
        
    End If
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub P_AumentaQTD(sCod_Pro As String, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    
    strSQL = "Update " & TAB_ITVENDAS & " set "
    
    strSQL = strSQL & "qtd_ite = qtd_ite+"
    strSQL = strSQL & F_GravaDec(CStr(cQtd)) & " "
    strSQL = strSQL & " where num_ven='" & Me.tag & "'  "
    strSQL = strSQL & " and cod_pro ='" & sCod_Pro & "' "
    conn.Execute strSQL
    
    
    
    'RESERVANDO O ESTOQUE
    strSQL = "update tab_produtos set er_pro=er_pro - " & F_GravaDec(CStr(cQtd))
    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
    conn.Execute strSQL
    
End Sub



Private Sub cmbxBuscar_Click()
    If cmbxBuscar.ListIndex = 0 Then
       cmndbtnLocaliza.tag = "B"
    Else
        cmndbtnLocaliza.tag = Left(cmbxBuscar.Text, 1)
    End If
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnCancel_Click()
    Dim strSQL  As String
    Dim i As Integer
    On Error Resume Next
    If ado.RecordCount > 0 Then
    
    If strExclVenda_usu <> "S" Then
        strSQL = F_senhaReturnUsu
        If F_BuscaUmCampo("exclvenda_usu", "" & TAB_DEFUSU & "", "cod_usu='" & strSQL & "'") <> "S" Then
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
        With ado
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                excluir False, seq_ite
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
        
        
        strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        strSQL = "delete from " & TAB_ITVENDAS & " where num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
            RefreshDataPosition frmEdiVendas.adc
        End If
        
        If MsgBox("Deseja Gerar uma Nova Venda?" & Chr(13) & "Clique em não caso queira abandonar o modo de vendas!", vbYesNo + vbQuestion) = vbYes Then
            strSQL = F_GeraNovaVenda("00000", strCod_Usu)
            iItem = 0
            Me.tag = strSQL
            bActive = False
            Form_Activate
        Else
            iItem = 0
            Unload Me
        End If
    End If
End Sub



Private Sub cmndbtnEditProd_Click()
    On Error GoTo trata_Erro
    Dim iPos As Integer
    Dim strSQL As String
    Dim cUlt_Qtd As Currency
    
    With ado
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        iPos = .AbsolutePosition
        frmValoresProd.lblCod_Pro.caption = !cod_pro
        frmValoresProd.lblNome_pro.caption = !nome_pro
        frmValoresProd.txtbxCusto.Text = FormatVal(!custo_pro)
        frmValoresProd.txtbxPreco.Text = FormatVal(!unit_ite)
        frmValoresProd.txtbxQtd.Text = FormatVal(!qtd_ite)
        frmValoresProd.tag = Me.tag
        cUlt_Qtd = !qtd_ite
        On Error Resume Next
        frmValoresProd.Show 1
        If CCur(frmValoresProd.txtbxQtd.Text) <= 0 Then
            MsgBox "Informe um valor decente!", vbExclamation
            Exit Sub
        End If
        strSQL = "Update " & TAB_ITVENDAS & " set qtd_ite=" & _
        F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & ", unit_ite=" & _
        F_GravaDec(frmValoresProd.txtbxPreco.Text)
        
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
        strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
        conn.Execute strSQL
        'ajustando estoque reservado
        strSQL = "Update tab_produtos set reservado_pro=reservado_pro-" & F_GravaDec(CStr(cUlt_Qtd))
        strSQL = strSQL & "+" & F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        conn.Execute strSQL
        Unload frmValoresProd
    End With
    ado.Requery
    ado.Move iPos
    P_AtualizaVenda
trata_Erro:
MsgBox "Err: cmndbtnEditProd_Click() "
End Sub

Private Sub cmndbtnExc_Prod_Click()
    dtgrd_KeyDown vbKeyDelete, 0
End Sub
Sub excluir(bQuestion As Boolean, seq_ite As String)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    Dim cod_pro As String
    Dim qtd_ite As String
    Dim seq As String
    Dim nome As String
    
    With grid
    
        
        .col = 1: cod_pro = .Text
        .col = 2: qtd_ite = .Text
        
        .col = 5: nome = .Text
        
        .col = 7: seq = .Text
        
    
        If .rows = 1 Then
            Exit Sub
        End If
        
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & ProdutoSelecionado.nome_pro & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        
        
        If pcBaixa.Visible = False Then
            strSQL = "update tab_produtos set er_pro=er_pro + " & F_GravaDec(ProdutoSelecionado.qtd)
            
            If operacao = "C" Then
                strSQL = strSQL & ", ef_pro=ef_pro + " & F_GravaDec(ProdutoSelecionado.qtd)
            End If
            
            strSQL = strSQL & " where cod_pro='" & ProdutoSelecionado.cod_pro & "' "
            conn.Execute strSQL
        End If
        
        
        strSQL = "delete from " & TAB_ITVENDAS & " where seq_ite='" & ProdutoSelecionado.s_pro & "' "
        
        strSQL = strSQL & " and num_ven='" & Me.tag & "' "
        conn.Execute strSQL
        
        
        
        
        
        P_GravaItemCancelado
        'ajustando estoque reservado
        
        
        
        'If bQuestion Then
            P_AtualizaVenda
        'End If
        
        If ado.RecordCount = 0 Then
            txtbxQtd.Enabled = False
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
    Exit Sub
trata_Erro:
    MostraErro "Excluir()"
End Sub

Private Sub cmndbtnLocaliza_Click()
    
    If F_Vazio(txtbxProduto.Text) Then
        MsgBox "Informe o produto que deseja inserir!", vbExclamation, strNomeSis
        txtbxProduto.SetFocus
        Exit Sub
    End If
    
    insere
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    If ado.RecordCount > 0 Then
        If lblAtacado.tag <> "" Then
            If MsgBox("Confirma Saída da Venda?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "Impossível Abandonar a Venda!", vbCritical
            Exit Sub
        End If
    End If
    If IsLoad("frmEdiVendas") Then
        frmEdiVendas.adc.Refresh
    End If
    If ado.RecordCount = 0 Then
        If F_BuscaUmCampo("cod_pro", "" & TAB_ITVENDAS & "", "num_ven='" & Me.tag & "'") = "" Then
            strSQL = "delete from " & TAB_VENDAS & " where num_ven='" & Me.tag & "' "
            conn.Execute strSQL
            strSQL = "delete from " & TAB_ITVENDAS & " where num_ven='" & Me.tag & "' "
            conn.Execute strSQL
        End If
        
    End If
    i = F_RetornoVariaveis
        If i Mod 2 > 0 Then
            ecf_CancelaCupom
        End If
    
    Unload Me
    Exit Sub
trata_Erro:
    MostraErro "cmndbtnSair"
    Unload Me
End Sub


Private Sub cmndbtnSalvar_Click()
    Dim sTroco As String
    Dim strSQL As String
    Dim sVen As String
    If ado.RecordCount = 0 Then
        MsgBox "Não há itens na sua venda!", vbExclamation
        Exit Sub
    End If
    
    
    strSQL = "update " & TAB_ITVENDAS & " set final_ite=(qtd_ite*unit_ite) "
    strSQL = strSQL & " where num_ven='" & Me.tag & "' "
    conn.Execute strSQL
    
    sTroco = F_Troco(CCur(lblValor.caption), Me.tag, sVen, txtbxNomeCli.tag, txtbxNomeCli.Text, operacao, lblCPF.caption)
    If sTroco <> "" Then
        If Atacado = True Then
            
                ImprimeComprovanteNaoFiscal Me.tag
            
        End If
        VerificaAtualizacao
        strSQL = F_GeraNovaVenda("00000", strCod_Usu)
        iItem = 0
        Me.tag = strSQL
        lblnum_ven.caption = strSQL
        P_buscaitens
        
        P_AtualizaVenda
        On Error Resume Next
        frmVenda.SetFocus
        txtbxProduto.Text = ""
        pcBaixa.Visible = False
        On Error Resume Next
        txtbxProduto.SetFocus
        ado.Requery
    End If
End Sub



Private Sub dtgrd_DblClick()
    cmndbtnEditProd_Click
End Sub

Private Sub dtgrd_GotFocus()
    
    dtgrd.MarqueeStyle = 3
    
    
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo trata_Erro
    Dim stSQL As String
    Dim Item As String
    
    grid.col = 7
    
    Item = grid.Text
    seq_ite = Item
    
    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        strSQL = F_senhaReturnUsu()
            
        If strSQL = "" Then
            MsgBox "Erro de autenticação!"
            Exit Sub
        Else
            strSQL = F_BuscaUmCampo("exclvenda_usu", "" & TAB_DEFUSU & "", "cod_usu='" & strSQL & "'")
            If strSQL <> "S" Then
                MsgBox "Você não tem permissão para cancelar o item. Chame o responsável!", vbExclamation
                Exit Sub
                
            End If
        End If
        
        If operacao = "C" Then
            If ado.RecordCount = 1 Then
                i = ecf_CancelaCupom
                If i <> 1 Then
                        MsgBox "Ocorreu um erro na tentativa de exclusão. Tente novamente!", vbCritical
                        P_AtualizaVenda
                        Exit Sub
                End If
                iItem = 0
            Else
                
                    If grid.Row = -1 Then
                        MsgBox "Selecione o item que deseja excluir!", vbCritical
                        dtgrd.SetFocus
                        Exit Sub
                    End If
                    i = ecf_CancelaItemGenerico(Item)
                    If i <> 1 Then
                        MsgBox "Ocorreu um erro na tentativa de exclusão. Tente novamente!", vbCritical
                        P_AtualizaVenda
                        Exit Sub
                    
                End If
            End If
            excluir False, seq_ite
            
        Else
            excluir False, seq_ite
        End If
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxProduto.SetFocus
    End If
    P_buscaitens
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

'Private Sub adc_Reposition()
'    With ado
'    On Error Resume Next
'     If Not .EOF Then
'        lblProduto.caption = !nome_pro
'
'
'                lblCapPreco.caption = "Preço:"
'                lblPreco.caption = "R$ " & FormatVal(!pv_pro)
'
'            lblCat.caption = !nome_cat
'            lblCod.caption = FormatCodigo(!cod_pro)
'            lblCodBarr.caption = F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'")
'            lblMarca.caption = !nome_mar
'            lblQtd.caption = !qtd_ite
'
'            lblBalanca.caption = !balanca_pro
'        End If
'    End With
'
'End Sub


Sub buscaPromo()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
        
    strSQL = "select nome_pro,pv_pro from tab_produtos as tp,"
    strSQL = strSQL & " tab_prodpromo as tpp where tpp.cod_Pro=tp.cod_pro"
    strSQL = strSQL & " and datafin_pro >=" & F_GravaData(Date)
    strSQL = strSQL & " order by nome_pro"
    
    rPesq.Open strSQL, conn_client
    
    With rPesq
        QtdPromo = .RecordCount
        If .RecordCount > 0 Then
            ReDim pPromo(.RecordCount) As prodPromo
            'lblPromoPreco.Visible = True
            'lblPromo.Visible = True
        Else
            'lblPromoPreco.Visible = False
            'lblPromo.Visible = False
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

Sub BuscaPic()
    On Error Resume Next
    pcLogo.Picture = LoadPicture(App.Path & "\pic\logo.jpg")
End Sub
Private Sub Form_Activate()
    'On Error GoTo trata_Erro
    mdiFrm.Enabled = False
    BuscaPic
    
    lblnum_ven.caption = Me.tag
    lblNome_Fun.caption = strCod_Usu & "-" & strNome_Usu
    lblNomeEmp.caption = strFantasia_Emp
    lblCaixa = strCod_Cai
    lblTur.caption = strCod_Tur
    P_CentraObejct lblMensagem, Me, False
    
    
    If bActive = True Then Exit Sub
    
    ajustaGrid
    
    P_Status "Buscando Itens..."
    txtbxProduto.SetFocus
    P_buscaitens
    If ado.RecordCount > 0 Then
        P_Status "Buscando dados do cupom aberto..."
        
        buscaCupomAberto
        P_Status
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
trata_Erro:
    MostraErro "Form_Activate"
    
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
        
        grid.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
         dtgrd_KeyDown vbKeyDelete, 0
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF2 Then
        'frmRecebConta.Show 1
    ElseIf KeyCode = vbKeyF8 Then
        If operacao = "B" And ado.RecordCount > 0 Then
            MsgBox "Operacação Impossível!", vbCritical
            
            Exit Sub
        End If
        
        pcECF.BackColor = vbGreen
        operacao = "C"
        mudamodo
    ElseIf KeyCode = vbKeyF9 Then
        If operacao = "C" And ado.RecordCount > 0 Then
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
        
        If ado.RecordCount > 0 Then
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
    conn.Execute strSQL
End Sub
Private Sub Form_Load()
    Set ado = New ADODB.Recordset
    buscaPromo
    bActive = False

    

    lblNomeEmp.caption = strNome_Emp
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", "tab_parametros", "")
    txtbxNomeCli.tag = "00000"
    txtbxNomeCli.Text = "VENDA A CONSUMIDOR"
    
    
    
    
'    cmbxBuscar.ListIndex = 0
    pcECF.BackColor = vbGreen
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDescancoTela = False
    iDescancoTela = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Resize()
    'ajustaForm
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).Width = pcAtalhos.Width / (pcBotao.count)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).Width + 50
        End If
        
    Next
End Sub
Sub ajustaForm()
'    frControle.Width = Me.Width - 10
    
'    pcStatusBar.Top = Me.Height - pcStatusBar.Height - 50
'    pcStatusBar.Width = Me.Width - 100
 
    pcMsg.Top = pcStatusBar.Top - pcMsg.Height - 10
'    pcMsg.Width = pcStatusBar.Width
    
    
'    frControle.Height = pcStatusBar.Top - 10
'    pcAtalhos.Width = Me.Width - 10
    lblNomeEmp.Left = pcStatusBar.Width - lblNomeEmp.Width - 50
    lblHora.Left = pcStatusBar.Width - lblHora.Width - 50
    pcProd.Top = pcMsg.Top - pcProd.Height - 10
'    pcprod.Left = frCli.Width + 200
    pcProd.Width = Me.Width - pcProd.Left - 100
    dtgrd.Width = pcProd.Width
    dtgrd.Height = pcProd.Top - 1000
'    frCli.Height = pcMsg.Top - frCli.Top - 150
'    pcPromo.Top = frCli.Height - pcPromo.Height - 10
    pcProd.Top = dtgrd.Top + 20 + dtgrd.Height
    pcProd.Left = dtgrd.Left
    pcECF.Top = dtgrd.Height + dtgrd.Top + 50
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
    ado.Close
    
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub frmCli_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub lblTef_Click()
    P_AdmTEF
End Sub

Private Sub grid_GotFocus()
    Dim i As Integer
    grid.BorderStyle = flexBorderSingle
     
    
End Sub
Sub SelecionarProduto()
    
    With grid
    
        .col = 1: ProdutoSelecionado.cod_pro = .Text
        .col = 2: ProdutoSelecionado.qtd = .Text
        .col = 3: ProdutoSelecionado.preco = CCur(.Text)
        .col = 5: ProdutoSelecionado.nome_pro = .Text
        .col = 7: ProdutoSelecionado.s_pro = .Text
    End With
End Sub

Private Sub grid_LostFocus()
    grid.BorderStyle = flexBorderNone
End Sub

Private Sub grid_RowColChange()
    Dim i As Integer
    With grid
        If .Row > 0 Then
            
            SelecionarProduto
            
            .col = 7
            seq_ite = .Text
            lblItemSelec.caption = seq_ite
            .col = 5
            lblProduto.caption = grid.Text
            .col = 6
            If .Text = "S" Then
                lblAtacado.caption = "Item em atacado"
                lblAtacado.Visible = True
            Else
                lblAtacado.caption = ""
                lblAtacado.Visible = False
                
            End If
            
            i = grid.Row
            P_buscaitens
            ZebraLinha grid, i, grid.BackColorSel, grid.ForeColorSel, True
            grid.Row = i
            
            
            
        End If
    End With
End Sub

Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
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
        dtgrd.SetFocus
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


Private Sub pcMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub


Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    Picture1.Left = X
    Picture1.Top = Y
End Sub



Private Sub Timer1_Timer()

    lblBloqueado.caption = F_BuscaUmCampo("bloqueado_cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    
    If lblBloqueado.caption = "S" Then
        If ado.RecordCount = 0 Then
            
        End If
    End If
    
    
            
    
    
    
    lblHora.caption = Now
    lblHora.Refresh
   End Sub

Private Sub txtbxCod_Cli_GotFocus()
    ganhaFoco txtbxCod_Cli
End Sub

Private Sub txtbxCod_Cli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
        
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
        txtbxNomeCli.tag = txtbxCod_Cli.Text

    End If
    
    
End Sub
Private Sub txtbxCod_Cli_LostFocus()
    If txtbxCod_Cli.Text = "" Then Exit Sub
    AjustaTexto txtbxCod_Cli
    txtbxCod_Cli.Text = Format(txtbxCod_Cli.Text, "0000#")
    
    buscaCodCli
End Sub

Private Sub txtbxNomeCli_GotFocus()
    ganhaFoco txtbxNomeCli
End Sub



Private Sub txtbxNomeCli_LostFocus()
    AjustaTexto txtbxNomeCli
    If txtbxNomeCli.tag <> "" Then
        P_AtualizaVenda
    End If
End Sub

Private Sub txtbxObs_GotFocus()
    
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub

Private Sub txtbxProduto_GotFocus()
    ganhaFoco txtbxProduto
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
    Dim p As Produto
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoQtd As String
    Dim pa As produtoAtacado
    
    
    
    If iItem = 0 And ado.RecordCount > 0 Then
        'buscar ultimo iitem
        
        strSQL = F_BuscaUmCampo("max(seq_ite)", "" & TAB_ITVENDAS & "", "num_ven='" & Me.tag & "'")
        If strSQL <> "" Then
            iItem = Int(strSQL)
        End If
        
    End If
    
'    On Error GoTo trata_erro
            p = F_BuscaProdSaida(txtbxProduto.Text, txtbxQtd.Text)
            If p.cod_pro <> "" Then
                If p.qtd = "" Then p.qtd = "1"
                If p.qtd > 1 Then
                    pa = buscaProduto_atacado(p.qtd, p.cod_pro)
                End If
                txtbxProduto.Text = p.nome_pro
                'lblNome_pro.caption = p.nome_pro
                'lblBalanca.caption = p.balanca_pro
                'lblMarca.caption = p.nome_mar
                'lblCodBarr.caption = p.codbarra_pro
                'lblCat.caption = p.nome_cat
                'lblCod.caption = p.cod_pro
                txtbxQtd.Text = p.qtd
                'lblQtd.caption = txtbxQtd.Text
                    
                
                If pa.codigo = "" Then
                
                
                    'lblPreco.caption = "R$ " & FormatVal(p.pv_pro)
                    lblUnit.caption = FormatVal(p.pv_pro)
                    lblAtacado.tag = "N"
                    lblAtacado.Visible = False
                
                Else
                    'lblPreco.caption = "R$ " & FormatVal(pa.preco)
                    'lblUnit.caption = FormatVal(pa.preco)
                    lblAtacado.tag = "S"
                    lblAtacado.Visible = True
                    lblAtacado.caption = "Unitário: " & FormatVal(p.pv_pro) & " Atacado: " & pa.qtd & " X " & FormatVal(pa.preco)
                    p.pv_pro = pa.preco
                End If
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
                            txtbxQtd.Text = p.totalbalanca_pro / p.pv_pro
                        Else
                            txtbxQtd.Text = Round(p.totalbalanca_pro / p.pv_pro, 3)
                        End If
                        txtbxQtd_LostFocus
                        lblUnit.caption = p.pv_pro
                    End If
                    
                End If
                                
                If p.pv_pro = 0 Then
                    MsgBox "Este produto encontra-se com o preço de venda zerado!", vbCritical
                    Exit Sub
                End If
                'lblNome_pro.caption = p.nome_pro
                P_msgProduto txtbxQtd.Text & " X " & p.nome_pro
                iItem = iItem + 1
                
                    
                
                If operacao = "C" Then
                    
                    
                    '    i = ecf_VerificaImpressoraLigada
                    
                    
                    i = F_RetornoVariaveis()
                    
                    If i = 4 Or i = 36 Or i = 0 Or i = 32 Then
                        strNum_Cpm = Space(6)
                        
                        i = ecf_NumeroCupom(strNum_Cpm)
                        
                         P_Msg "Abrindo Cupom Fiscal..."
                        
                        i = ecf_AbreCupom(lblCPF.caption)
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
                    i = ecf_VendeItem(p.cod_pro, Left(p.nome_pro, 29), p.ecf_icms, sTipoQtd, txtbxQtd.Text, 3, Format(p.pv_pro, "###,##0.000"), "%", "0")
                    If i = 1 Then
                        P_InsereProd p.cod_pro, p.nome_pro, p.pv_pro, CCur(txtbxQtd.Text), lblAtacado.tag
                    Else
                        If i = -3 Then
                            MsgBox "Alíquota não programada!", vbExclamation
                        ElseIf i = -2 Then
                            MsgBox "Parâmetro inválido na função de VendeItem!", vbExclamation
                        ElseIf i = -4 Then
                            MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows", vbExclamation
                        ElseIf i = 0 Then
                            MsgBox "Erro de comunicação com a impressora. ", vbExclamation
                        ElseIf i <> 1 Then
                            MsgBox "Erro na passagem do item para a impressora fiscal!", vbCritical
                        End If
                    
                    End If
                    ecf_verificaImpressora False
                    P_Msg ""
                    P_msgProduto ""
                Else
                    If txtbxQtd.Text = "" Then
                        MsgBox "Informe uma quantidade válida!", vbExclamation
                        'txtbxQtd.SetFocus
                        Exit Sub
                        'txtbxQtd.Text = "1"
                    End If
                    P_InsereProd p.cod_pro, p.nome_pro, p.pv_pro, CCur(txtbxQtd.Text), lblAtacado.tag
                End If
                            
                            
                
                            
                
            End If
        
        P_buscaitens
        txtbxQtd.Text = "1"
                
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
        Me.Refresh
        P_AtualizaVenda
        
        
        FechaObjeto rPesq
        
        Exit Sub
trata_Erro:
    MostraErro "Insere()"
End Sub
Private Sub txtbxQtd_GotFocus()
    
    ganhaFoco txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo trata_Erro
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
trata_Erro:
    MostraErro
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
        PerdeFoco txtbxQtd, "###,##0.000"
    Else
        PerdeFoco txtbxQtd, "###,##0"
    End If
    txtbxQtd.Enabled = False
    txtbxProduto.Text = ""
    txtbxProduto.Enabled = True
    txtbxProduto.SetFocus
    
    
    
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    Timer1_Timer
End Sub


Sub P_GravaItemCancelado()
    On Error Resume Next
    Dim i As Integer
    strSQL = "insert into tab_itvndcncl(seq_itv,cod_pro,num_ven,seq_ite,qtd_ite,unit_ite,cod_cai,cod_tur,data_ven)values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_itvndcncl", "seq_itv", "", "000000#") & "',"
    strSQL = strSQL & "'" & ado!cod_pro & "', "
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "" & seq_ite & ","
    strSQL = strSQL & F_GravaDec(ado!qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(ado!unit_ite) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "', "
    strSQL = strSQL & F_GravaDataTempo(Now) & ")"
    conn.Execute strSQL, i
    
    
    
End Sub
