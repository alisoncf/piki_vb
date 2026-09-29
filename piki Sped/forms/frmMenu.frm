VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Barra de Ferramentas"
   ClientHeight    =   9960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14565
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   14565
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   3315
      TabIndex        =   92
      Top             =   7920
      Visible         =   0   'False
      Width           =   3375
      Begin BARCODEXLib.BarcodeX bcxVenda 
         Height          =   615
         Left            =   120
         Top             =   120
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   1085
         _StockProps     =   13
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "æ0123456"
         Title           =   ""
         hasText         =   0
         Scaling         =   -1
         BarcodeType     =   12
      End
   End
   Begin VB.PictureBox pcCaixa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3375
      ScaleWidth      =   1815
      TabIndex        =   74
      Top             =   3000
      Width           =   1815
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   11
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   79
         Top             =   120
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   11
            Left            =   600
            Picture         =   "frmMenu.frx":0000
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lançamentos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   315
            TabIndex        =   80
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   25
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   77
         Top             =   1080
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fechamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   345
            TabIndex        =   78
            Top             =   600
            Width           =   915
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   25
            Left            =   600
            Picture         =   "frmMenu.frx":0CCA
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   19
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   75
         Top             =   2040
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NFEs"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   600
            TabIndex        =   76
            Top             =   600
            Width           =   405
         End
         Begin VB.Image imgIco 
            Height          =   510
            Index           =   19
            Left            =   480
            Picture         =   "frmMenu.frx":121F
            Top             =   0
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox pcShapeLongo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   840
      Picture         =   "frmMenu.frx":1680
      ScaleHeight     =   360
      ScaleWidth      =   8055
      TabIndex        =   71
      Top             =   7560
      Visible         =   0   'False
      Width           =   8085
   End
   Begin VB.PictureBox pcShape 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   7800
      Picture         =   "frmMenu.frx":ADE2
      ScaleHeight     =   1050
      ScaleWidth      =   5010
      TabIndex        =   70
      Top             =   8880
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.PictureBox pcBarcode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   9135
      TabIndex        =   69
      Top             =   8880
      Visible         =   0   'False
      Width           =   9165
      Begin BARCODEXLib.BarcodeX bcx 
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   9135
         _Version        =   65536
         _ExtentX        =   16113
         _ExtentY        =   1508
         _StockProps     =   13
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "12345678901234567890123456789012345678901234"
         Title           =   ""
         hasText         =   0
         Scaling         =   -1
         BarcodeType     =   11
      End
   End
   Begin VB.PictureBox pcCalend 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   9960
      Left            =   12105
      ScaleHeight     =   9960
      ScaleWidth      =   2460
      TabIndex        =   34
      Top             =   0
      Width           =   2460
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   7
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2175
         TabIndex        =   65
         Top             =   8280
         Width           =   2175
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sair"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   1020
            TabIndex        =   66
            Top             =   600
            Width           =   285
         End
         Begin VB.Image imgIco 
            Height          =   570
            Index           =   7
            Left            =   840
            Picture         =   "frmMenu.frx":1C0AC
            Top             =   0
            Width           =   570
         End
      End
      Begin VB.PictureBox pcVendasAb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   2385
         TabIndex        =   51
         Top             =   120
         Width           =   2415
         Begin MSDataGridLib.DataGrid grid 
            Height          =   1575
            Left            =   0
            TabIndex        =   53
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2778
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            ColumnHeaders   =   -1  'True
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "cliente"
               Caption         =   "Cliente"
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
               DataField       =   "data"
               Caption         =   "Data"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/mm HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "numero"
               Caption         =   "Número"
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
               RecordSelectors =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas em aberto"
            Height          =   195
            Left            =   480
            TabIndex        =   52
            Top             =   0
            Width           =   1290
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3825
         ScaleWidth      =   2385
         TabIndex        =   43
         Top             =   2160
         Width           =   2415
         Begin VB.PictureBox pcPrinterCopiadeVenda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            Picture         =   "frmMenu.frx":1D226
            ScaleHeight     =   330
            ScaleWidth      =   330
            TabIndex        =   91
            Top             =   3120
            Width           =   330
         End
         Begin VB.PictureBox pcPrinterNfce 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            Picture         =   "frmMenu.frx":1D840
            ScaleHeight     =   330
            ScaleWidth      =   330
            TabIndex        =   90
            Top             =   2400
            Width           =   330
         End
         Begin VB.PictureBox pcPrinterPadrao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            Picture         =   "frmMenu.frx":1DE5A
            ScaleHeight     =   330
            ScaleWidth      =   330
            TabIndex        =   89
            Top             =   1800
            Width           =   330
         End
         Begin VB.PictureBox pcEditCaixa 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2040
            Picture         =   "frmMenu.frx":1E474
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   54
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label lblPrintCopiaVenda 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00916024&
            Height          =   240
            Left            =   120
            TabIndex        =   88
            Top             =   3480
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cópia de Venda"
            Height          =   195
            Left            =   600
            TabIndex        =   87
            Top             =   3120
            Width           =   1140
         End
         Begin VB.Label lblTurno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   960
            TabIndex        =   73
            Top             =   1560
            Width           =   105
         End
         Begin VB.Label lblSituacaoCaixa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1320
            TabIndex        =   72
            Top             =   1560
            Width           =   180
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NFCe:"
            Height          =   195
            Left            =   600
            TabIndex        =   64
            Top             =   2520
            Width           =   450
         End
         Begin VB.Label lblImpressoraNFCe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00916024&
            Height          =   240
            Left            =   120
            TabIndex        =   63
            Top             =   2760
            Width           =   315
         End
         Begin VB.Label lblImpressora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   62
            Top             =   2160
            Width           =   180
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Padrao"
            Height          =   195
            Left            =   600
            TabIndex        =   61
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label lblIp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "000"
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
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblUs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   195
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caixa:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   435
         End
         Begin VB.Label lblCaixa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
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
            Left            =   120
            TabIndex        =   46
            Top             =   1560
            Width           =   225
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computador/Host:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1305
         End
         Begin VB.Label lblhost 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblHost"
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
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   810
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   0
         ScaleHeight     =   2025
         ScaleWidth      =   2385
         TabIndex        =   35
         Top             =   6120
         Width           =   2415
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visão"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   765
            Width           =   390
         End
         Begin VB.Label lblVisao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblVisao"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   1005
            Width           =   675
         End
         Begin VB.Label lblCodUsu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "000"
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
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   405
         End
         Begin VB.Label lblUsu 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome do Usuário"
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
            Left            =   600
            TabIndex        =   39
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   285
            Width           =   585
         End
         Begin VB.Image imgUsu 
            Height          =   180
            Left            =   2040
            MouseIcon       =   "frmMenu.frx":1E5CD
            Picture         =   "frmMenu.frx":1FF4F
            Top             =   240
            Width           =   210
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblVisao"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conexão"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox pcProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   2055
      TabIndex        =   27
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   32
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "frmMenu.frx":20004
            Picture         =   "frmMenu.frx":2164E
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   33
            Top             =   600
            Width           =   645
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   21
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   30
         Top             =   1200
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seções"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   645
            TabIndex        =   31
            Top             =   600
            Width           =   555
         End
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   360
            Index           =   21
            Left            =   720
            MouseIcon       =   "frmMenu.frx":22318
            Picture         =   "frmMenu.frx":23962
            Top             =   120
            Width           =   360
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   22
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   28
         Top             =   2160
         Width           =   1815
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   22
            Left            =   720
            MouseIcon       =   "frmMenu.frx":23D40
            Picture         =   "frmMenu.frx":2538A
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subseções"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   510
            TabIndex        =   29
            Top             =   600
            Width           =   825
         End
      End
   End
   Begin VB.PictureBox pcClientes 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   2040
      ScaleHeight     =   8655
      ScaleWidth      =   2055
      TabIndex        =   24
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   8
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   81
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   8
            Left            =   720
            Picture         =   "frmMenu.frx":258E1
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Representantes"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   345
            TabIndex        =   82
            Top             =   600
            Width           =   1155
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   23
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   67
         Top             =   2160
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   23
            Left            =   720
            Picture         =   "frmMenu.frx":25DA4
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionários"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   465
            TabIndex        =   68
            Top             =   600
            Width           =   915
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   9
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   57
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   9
            Left            =   720
            Picture         =   "frmMenu.frx":2624E
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedores"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   420
            TabIndex        =   58
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   25
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   1
            Left            =   720
            MouseIcon       =   "frmMenu.frx":26702
            Picture         =   "frmMenu.frx":28084
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clientes"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   26
            Top             =   600
            Width           =   585
         End
      End
   End
   Begin VB.PictureBox pcVendas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   4080
      ScaleHeight     =   8655
      ScaleWidth      =   2055
      TabIndex        =   19
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   13
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   85
         Top             =   4080
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Orçamentos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   480
            TabIndex        =   86
            Top             =   600
            Width           =   885
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   13
            Left            =   720
            Picture         =   "frmMenu.frx":284E2
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   12
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   83
         Top             =   3120
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo Orçamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   300
            TabIndex        =   84
            Top             =   600
            Width           =   1245
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   12
            Left            =   720
            Picture         =   "frmMenu.frx":28917
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   10
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   59
         Top             =   2160
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NFEs"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   720
            TabIndex        =   60
            Top             =   600
            Width           =   405
         End
         Begin VB.Image imgIco 
            Height          =   510
            Index           =   10
            Left            =   600
            Picture         =   "frmMenu.frx":28D38
            Top             =   0
            Width           =   750
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   5
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   55
         Top             =   5040
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entregas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   600
            TabIndex        =   56
            Top             =   600
            Width           =   645
         End
         Begin VB.Image imgIco 
            Height          =   360
            Index           =   5
            Left            =   720
            Picture         =   "frmMenu.frx":29199
            Top             =   120
            Width           =   360
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   22
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   4
            Left            =   720
            Picture         =   "frmMenu.frx":2962E
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   645
            TabIndex        =   23
            Top             =   600
            Width           =   555
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   3
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   20
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   3
            Left            =   720
            Picture         =   "frmMenu.frx":29A63
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nova Venda"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   465
            TabIndex        =   21
            Top             =   600
            Width           =   915
         End
      End
   End
   Begin VB.PictureBox pcCompras 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   8160
      ScaleHeight     =   8655
      ScaleWidth      =   2055
      TabIndex        =   14
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   2
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   17
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   2
            Left            =   600
            Picture         =   "frmMenu.frx":29E84
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compras"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   18
            Top             =   600
            Width           =   645
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   24
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nova Compra"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   300
            TabIndex        =   16
            Top             =   600
            Width           =   1005
         End
         Begin VB.Image imgIco 
            Height          =   510
            Index           =   24
            Left            =   600
            Picture         =   "frmMenu.frx":2AB4E
            Top             =   120
            Width           =   510
         End
      End
   End
   Begin VB.PictureBox pcUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   10200
      ScaleHeight     =   8655
      ScaleWidth      =   2055
      TabIndex        =   11
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   6
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   12
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   600
            Index           =   6
            Left            =   600
            Picture         =   "frmMenu.frx":2B570
            Top             =   0
            Width           =   600
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   645
            TabIndex        =   13
            Top             =   600
            Width           =   555
         End
      End
   End
   Begin VB.PictureBox pcFinanceiro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   6120
      ScaleHeight     =   8655
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   18
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   9
         Top             =   3120
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   18
            Left            =   615
            Picture         =   "frmMenu.frx":2C872
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Receber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   135
            TabIndex        =   10
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   17
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   7
         Top             =   2160
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimento Financeiro"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   75
            TabIndex        =   8
            Top             =   600
            Width           =   1575
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   17
            Left            =   600
            Picture         =   "frmMenu.frx":2D53C
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   16
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Pagar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   300
            TabIndex        =   6
            Top             =   600
            Width           =   1125
         End
         Begin VB.Image imgIco 
            Height          =   630
            Index           =   16
            Left            =   480
            Picture         =   "frmMenu.frx":2D95B
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   15
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   495
            Index           =   15
            Left            =   600
            Picture         =   "frmMenu.frx":2DDDA
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   210
            TabIndex        =   4
            Top             =   600
            Width           =   1305
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   27
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   1
         Top             =   4080
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Pagar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   1245
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   27
            Left            =   615
            Picture         =   "frmMenu.frx":2E215
            Top             =   90
            Width           =   480
         End
      End
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   2040
      TabIndex        =   50
      Top             =   8280
      Width           =   10095
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bBlock As Boolean
Dim bActive As Boolean

Sub P_foco(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    
    

    
    On Error Resume Next
    For iCount = 0 To 50
        lblBotao(iCount).FontBold = False
        lblBotao(iCount).FontUnderline = False
        
    Next
    lblBotao(Indice).FontBold = True
    lblBotao(Indice).FontUnderline = True

End Sub

Sub buscVendasAbertas()
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    strSQL = "select data_ven as data,nomecli_ven as cliente,num_ven as numero from " & TAB_VENDAS & " where abrevenda_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " and abrevenda_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " and situacao_ven=1 order by data_ven"
    
    
    abreObjetoAdo ado, strSQL, grid
    
    
    
'    AvisaVendasAbertasPorUsuario 1
    
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    Dim strSQL As String
    habilita_controles
     
    
    
    
    lblDia.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    lblNumDia.caption = Day(Date)
    lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    lblAno.caption = Year(Date)
    strSQL = LCase(strNomeBanco)
    If Left(strSQL, 8) <> "requinte" Then
    
        pcLogo.Picture = Nothing
        
        'pcLogo.Picture = LoadPicture(app.Path & "\img
        
    End If
    lblIp.caption = strIP_Usu
    mudaVariaveis
    
    resizeForm Me

    mudaCor strCorFundoUsu
    If Not bUsaConfeccao Then
        buscVendasAbertas
    End If
    If Not bUsaEntrega Then
        pcBotao(5).Visible = False
    End If
End Sub
Sub habilita_controles()
    If strVisaoEstoque_usu <> "S" And strVisaoGeral_usu <> "S" Then
        pcProduto.Visible = False
        mdiFrm.mnProdutos.Enabled = False
        mdiFrm.mnServ.Enabled = False
        mdiFrm.mnTipoServ.Enabled = False
        mdiFrm.mnEntradas.Enabled = False
        pcCompras.Visible = False
    End If
    If strCadFun_Usu <> "S" Then
        pcBotao(8).Visible = False
        pcBotao(23).Visible = False
        mdiFrm.mnFun.Enabled = False
        mdiFrm.mnCadCargos.Enabled = False
        mdiFrm.mnEdiUsu.Enabled = False
        
    End If
    If strCadcli_usu <> "S" Then
        mdiFrm.mnEditCli.Enabled = False
        pcBotao(1).Visible = False
        
    End If
    If strCadfor_Usu <> "S" Then
        mdiFrm.mnForn.Enabled = False
        pcBotao(9).Visible = False
    End If
    If strVisaoFinanceiro_usu <> "S" And strVisaoGeral_usu <> "S" Then
        mdiFrm.mnFinan.Enabled = False
        pcFinanceiro.Visible = False
    End If
    
    If strVisaoCaixa_usu <> "S" And strVisaoGeral_usu <> "S" Then
        mdiFrm.mnCaixa.Enabled = False
    End If
    
    If strVisaoGeral_usu <> "S" Then
        mdiFrm.mnParam.Enabled = False
        mdiFrm.mnRel.Enabled = False
    Else
        mdiFrm.mnSped.Enabled = True
    End If
    
    mdiFrm.mnSped.Enabled = True
    
    If bUsaConfeccao Then
        pcVendasAb.Visible = False
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.height = mdiFrm.height - 200
    Me.width = mdiFrm.ScaleWidth - 50
    Me.Top = 500
    Me.Left = 0
    'Me.Visible = True
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.PopupMenu mdiFrm.mnOptCor
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    P_mensagem
End Sub

Private Sub grid_DblClick()
    On Error Resume Next
    
    editaVenda grid.Columns(2).Value
    
    
End Sub

Private Sub grid_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    
    pcVendas.width = 800
    pcVendas.width = 400
    pcVendas.Left = 200
    pcVendas.container = mdiFrm
    
    
End Sub

Private Sub imgIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub imgIco_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pcBotao_MouseMove Index, Button, Shift, X, Y
End Sub


Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub lblImpressora_Click()
    MudaImpressoraPadrao
    lblImpressora.caption = Printer.deviceName
End Sub

Private Sub lblImpressoraNFCe_Click()
    MudaImpressoraPadraoNFCe
     
     
     mudaVariaveis

End Sub
Public Sub mudaVariaveis()
    Dim strSQL As String
    lblIp.caption = strIP_Usu
    lblhost.caption = strHost_Usu
    lblCaixa.caption = strCod_Cai
    
    strSQL = F_BuscaUmCampo("status_cai", TAB_CAIXAS, "cod_cai='" & strCod_Cai & "'")
    If strSQL = "F" Then
        lblSituacaoCaixa.caption = "Fechado"
        lblSituacaoCaixa.ForeColor = vbRed
    
    ElseIf strSQL = "A" Then
        lblSituacaoCaixa.caption = "Aberto"
        lblSituacaoCaixa.ForeColor = vbGreen
    Else
        lblSituacaoCaixa.caption = "Desconhecido"
        lblSituacaoCaixa.ForeColor = vbMagenta
    End If
    lblTurno.caption = F_BuscaUmCampo("cod_tur", TAB_CAIXAS, "cod_cai='" & strCod_Cai & "'")
    lblUsu.caption = strNome_Usu
    lblCodUsu.caption = strCod_Usu
    lblVisao.caption = strVisao
    lblBanco.caption = conn.DefaultDatabase
    'lblECF.caption = strImprFiscal
    lblImpressora.caption = Printer.deviceName
    lblImpressoraNFCe.caption = ImpressoraPadraoCupomNFCe.deviceName
    lblPrintCopiaVenda.caption = ImpressoraCopiaVenda.deviceName
    
    
End Sub

Private Sub lblPrintCopiaVenda_Click()
    MudaImpressoraCopiaVenda
    mudaVariaveis
    
     
    

End Sub

Private Sub lblVisao_Click()
    frmEscolherVisao.Show 1
End Sub
Public Sub MudaCorFundo(objeto As Object, cor)
    On Error Resume Next
    objeto.BackColor = cor
    
End Sub
Sub mudaCor(Optional cor = &HEEDFD5)
  On Error Resume Next
    MudaCorFundo Me, cor
    MudaCorFundo pcCaixa, cor
    MudaCorFundo pcClientes, cor
    MudaCorFundo pcCompras, cor
'    MudaCorFundo pcConsign, cor
'    MudaCorFundo pcPacote, cor
    MudaCorFundo pcProduto, cor
    MudaCorFundo pcVendas, cor
    MudaCorFundo pcFinanceiro, cor
    MudaCorFundo pcUsuario, cor
    
    
    Dim i As Integer
    For i = 0 To 99
        MudaCorFundo pcBotao(i), cor
        lblBotao(i).FontName = "Arial"
        lblBotao(i).fontsize = 10
    Next
End Sub
 Private Sub pcBotao_Click(Index As Integer)
    'On Error Resume Next
    If Index = 0 Then
        frmEdiPro.Show 0, mdiFrm
    ElseIf Index = 1 Then
        frmEdiClientes.Show 0, mdiFrm
    ElseIf Index = 2 Then
        frmEntradasSped.Show 0, mdiFrm
    ElseIf Index = 3 Then
        NovaVenda
    ElseIf Index = 4 Then
        frmEdiVendas.Show 0, mdiFrm
    ElseIf Index = 5 Then
        frmEntregas.Show 0, mdiFrm
    ElseIf Index = 6 Then
        frmOptUsu.Show 0, mdiFrm
    ElseIf Index = 7 Then
        If MsgBox("Deseja sair do sistema?", vbYesNo + vbQuestion) = vbYes Then
            P_Fim
        End If
    ElseIf Index = 8 Then
        frmEdiRep.Show 0
        frmEdiRep.SetFocus
    ElseIf Index = 9 Then
        frmEdiFor.Show 0, mdiFrm
        
    ElseIf Index = 10 Then
        'If strCreditario_Usu = "S" Then
        '     frmCrediario.Show 0
        'Else
        '    MsgBox "Você não tem permissão para este acesso!", vbExclamation
            
        'End If
        frmEdiNFe.Show 0, mdiFrm
    ElseIf Index = 11 Then
    
        frmEdiCai.Show 0, mdiFrm
    ElseIf Index = 12 Then
    
        frmOrcamentoUniformes.Show 0, mdiFrm
    ElseIf Index = 13 Then
        frmEdiOrc.Show 0, mdiFrm
        
    ElseIf Index = 14 Then
    
        P_AdmTEF
    ElseIf Index = 15 Then
        frmEditaMov.Show 0
    ElseIf Index = 16 Then
        frmContasPagar.Show 0
    ElseIf Index = 17 Then
        frmEdiFin.Show 0
    
    ElseIf Index = 18 Then
        frmControleChequesReceber.Show 0
    ElseIf Index = 19 Then
        frmEdiNFe.Show 0, mdiFrm
    ElseIf Index = 20 Then
        frmCobranca.Show 0
    ElseIf Index = 21 Then
        'secao.P_EditaSecao
    ElseIf Index = 22 Then
            frmEdiSubSec.Show 0
    frmEdiSubSec.SetFocus

    ElseIf Index = 23 Then
        frmEdiFun.Show 0
        frmEdiFun.SetFocus
    ElseIf Index = 24 Then
        frmEntradaSped.Show 0, mdiFrm
    ElseIf Index = 25 Then
        If strFechacai_usu = "S" Then
            frmFechaCaixa.Show 1
        Else
            MsgBox "Acesso inválido!", vbExclamation
        End If
        
    ElseIf Index = 26 Then
        frmECF.Show 0
    ElseIf Index = 27 Then
        frmControleChequesPagar.Show 0
    ElseIf Index = 28 Then
        frmRelComiss.Show 0
        
    ElseIf Index = 29 Then
        frmPacote.Show 0
    ElseIf Index = 30 Then
        frmAcerto.acerto_promocao = True
        frmAcerto.Show 0
        
    End If
End Sub

Private Sub pcBotao_GotFocus(Index As Integer)
    P_focoMouse Index
End Sub

Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Index < pcBotao.Count - 1 Then
            pcBotao(Index + 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyLeft Then
        
    ElseIf KeyCode = vbKeyRight Then
    
    ElseIf KeyCode = vbKeyUp Then
        If Index > 0 Then
            pcBotao(Index - 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcBotao_Click Index
    End If
End Sub

Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
    
    
    
   ' P_mensagem Index
    
    
    
End Sub

Sub P_mensagem(Optional Index As Integer = -1)
    On Error Resume Next
    If Index = -1 Then
        lblMsg.caption = ""
        lblMsg.Visible = False
        Exit Sub
    ElseIf Index = 0 Then
        lblMsg.caption = "Abre lista de produtos. Consulta, Alteração, Inclusão de um Novo Produto"
        
    ElseIf Index = 1 Then
        lblMsg.caption = "Caso queira consultar, incluir, alterar um cliente. Clique para entrar."
    ElseIf Index = 2 Then
        lblMsg.caption = "Notas Fiscais de compra. Buscar um nota, concluir um lançamento, lançar um nota Fiscal"
    ElseIf Index = 3 Then
        lblMsg.caption = "Clique para efetuar uma nova Venda."
    ElseIf Index = 4 Then
        lblMsg.caption = "Para concluir uma venda ou pesquisar nas vendas já efetuadas, clique aqui."
    ElseIf Index = 5 Then
        lblMsg.caption = "Lançamentos diários do Caixa."
    ElseIf Index = 6 Then
        lblMsg.caption = "Alterar senha, relogar no sistema, e outras opções pessoais."
    ElseIf Index = 7 Then
        lblMsg.caption = "Sair para o Windows."
    End If
    
  '  pcMsg.Left = Me.Width - pcMsg.Width
  '  pcMsg.Top = Me.Height - pcMsg.Width
    lblMsg.Visible = True
    
    
End Sub

Private Sub pcEditCaixa_Click()
    If strExclcai_usu = "S" Then
        frmCadCaixa.Show 1
    End If
    
    
        
    
    
    
    
    
End Sub

Private Sub pcPrinterCopiadeVenda_Click()
    lblPrintCopiaVenda_Click
End Sub

Private Sub pcPrinterNfce_Click()
    lblImpressoraNFCe_Click
End Sub

Private Sub pcPrinterPadrao_Click()
    lblImpressora_Click
End Sub

Private Sub pcProduto_GotFocus()
    pcProduto.BorderStyle = 1
End Sub

Private Sub pcProduto_LostFocus()
    pcProduto.BorderStyle = 0
End Sub

