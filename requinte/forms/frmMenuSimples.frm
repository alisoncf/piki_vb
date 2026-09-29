VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Área de Trabalho"
   ClientHeight    =   9930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Menu de Opções"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      ScaleHeight     =   945
      ScaleWidth      =   7905
      TabIndex        =   103
      Top             =   8040
      Visible         =   0   'False
      Width           =   7935
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
   Begin VB.PictureBox pcShapeLongo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1920
      Picture         =   "frmMenuSimples.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   8055
      TabIndex        =   102
      Top             =   240
      Visible         =   0   'False
      Width           =   8085
   End
   Begin VB.PictureBox pcShape 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   6360
      Picture         =   "frmMenuSimples.frx":9762
      ScaleHeight     =   1050
      ScaleWidth      =   5010
      TabIndex        =   101
      Top             =   8640
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.PictureBox pcPac 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   4320
      ScaleHeight     =   1215
      ScaleWidth      =   2055
      TabIndex        =   78
      Top             =   6480
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   29
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   79
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   29
            Left            =   600
            Picture         =   "frmMenuSimples.frx":1AA2C
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pacote"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   540
            TabIndex        =   80
            Top             =   600
            Width           =   525
         End
      End
      Begin VB.Label pcPacote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pacote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   81
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox pcConsign 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   2160
      ScaleHeight     =   5295
      ScaleWidth      =   2055
      TabIndex        =   51
      Top             =   2400
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   30
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   86
         Top             =   4200
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acerto Promocional"
            Height          =   195
            Index           =   30
            Left            =   150
            TabIndex        =   87
            Top             =   600
            Width           =   1395
         End
         Begin VB.Image imgIco 
            Height          =   285
            Index           =   30
            Left            =   720
            Picture         =   "frmMenuSimples.frx":1B6F6
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
         Index           =   13
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   58
         Top             =   3240
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   13
            Left            =   600
            Picture         =   "frmMenuSimples.frx":1B853
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acerto"
            Height          =   195
            Index           =   13
            Left            =   600
            TabIndex        =   59
            Top             =   600
            Width           =   495
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
         TabIndex        =   56
         Top             =   2160
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   12
            Left            =   720
            Picture         =   "frmMenuSimples.frx":1C11D
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Devolução"
            Height          =   195
            Index           =   12
            Left            =   465
            TabIndex        =   57
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   11
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   54
         Top             =   1080
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transferência"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   55
            Top             =   600
            Width           =   1005
         End
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   11
            Left            =   720
            Picture         =   "frmMenuSimples.frx":1C6A7
            Top             =   120
            Width           =   240
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
         TabIndex        =   52
         Top             =   120
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   9
            Left            =   600
            Picture         =   "frmMenuSimples.frx":1CC31
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consignações"
            Height          =   195
            Index           =   9
            Left            =   405
            TabIndex        =   53
            Top             =   600
            Width           =   1035
         End
      End
   End
   Begin VB.PictureBox pcFinanceiro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   6480
      ScaleHeight     =   7695
      ScaleWidth      =   2055
      TabIndex        =   22
      Top             =   0
      Width           =   2055
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
         TabIndex        =   82
         Top             =   6120
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pendências"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   495
            TabIndex        =   83
            Top             =   600
            Width           =   855
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   10
            Left            =   720
            Picture         =   "frmMenuSimples.frx":1D4FB
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
         Index           =   27
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   49
         Top             =   4080
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   27
            Left            =   615
            Picture         =   "frmMenuSimples.frx":1DDC5
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Pagar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   1245
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
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   210
            TabIndex        =   47
            Top             =   600
            Width           =   1305
         End
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   15
            Left            =   720
            Picture         =   "frmMenuSimples.frx":1EA8F
            Top             =   120
            Width           =   240
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
         TabIndex        =   25
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   16
            Left            =   720
            Picture         =   "frmMenuSimples.frx":1ED13
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Pagar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   300
            TabIndex        =   46
            Top             =   600
            Width           =   1125
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
         TabIndex        =   26
         Top             =   2160
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   17
            Left            =   600
            Picture         =   "frmMenuSimples.frx":1F12A
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimento Financeiro"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   75
            TabIndex        =   45
            Top             =   600
            Width           =   1575
         End
      End
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
         TabIndex        =   27
         Top             =   3120
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Receber"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   135
            TabIndex        =   44
            Top             =   600
            Width           =   1455
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   18
            Left            =   615
            Picture         =   "frmMenuSimples.frx":1F9F4
            Top             =   90
            Width           =   480
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   20
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   28
         Top             =   5040
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrança"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   510
            TabIndex        =   43
            Top             =   600
            Width           =   705
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   20
            Left            =   600
            Picture         =   "frmMenuSimples.frx":206BE
            Top             =   90
            Width           =   480
         End
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financeiro"
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
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox pcCaixa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   2055
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   19
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   95
         Top             =   2280
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   510
            Index           =   19
            Left            =   480
            Picture         =   "frmMenuSimples.frx":21388
            Top             =   0
            Width           =   750
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NFEs"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   600
            TabIndex        =   96
            Top             =   600
            Width           =   405
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   14
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   66
         Top             =   3120
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   14
            Left            =   720
            Picture         =   "frmMenuSimples.frx":217E9
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TEF"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   645
            TabIndex        =   67
            Top             =   600
            Width           =   315
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   26
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ECF"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   645
            TabIndex        =   41
            Top             =   600
            Width           =   315
         End
         Begin VB.Image imgIco 
            Height          =   420
            Index           =   26
            Left            =   600
            Picture         =   "frmMenuSimples.frx":21B73
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
         Index           =   25
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   25
            Left            =   600
            Picture         =   "frmMenuSimples.frx":224F8
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fechamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   345
            TabIndex        =   40
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
         Index           =   5
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lançamentos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   315
            TabIndex        =   39
            Top             =   600
            Width           =   975
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   5
            Left            =   600
            Picture         =   "frmMenuSimples.frx":22A4D
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox pcCompras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   4320
      ScaleHeight     =   2295
      ScaleWidth      =   2055
      TabIndex        =   14
      Top             =   4080
      Width           =   2055
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   510
            Index           =   24
            Left            =   600
            Picture         =   "frmMenuSimples.frx":23717
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nova Compra"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   300
            TabIndex        =   38
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
         Index           =   2
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compras"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   37
            Top             =   600
            Width           =   645
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   2
            Left            =   600
            Picture         =   "frmMenuSimples.frx":24139
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox pcVendas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   4320
      ScaleHeight     =   3975
      ScaleWidth      =   2055
      TabIndex        =   18
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   33
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   88
         Top             =   3000
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   420
            Index           =   33
            Left            =   600
            Picture         =   "frmMenuSimples.frx":24E03
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NFe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   645
            TabIndex        =   89
            Top             =   600
            Width           =   315
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   28
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   84
         Top             =   2040
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comissão Vendedor"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   150
            TabIndex        =   85
            Top             =   600
            Width           =   1425
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   28
            Left            =   600
            Picture         =   "frmMenuSimples.frx":25788
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
         Index           =   3
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   20
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nova Venda"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   345
            TabIndex        =   36
            Top             =   600
            Width           =   915
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   3
            Left            =   600
            Picture         =   "frmMenuSimples.frx":26052
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
         Index           =   4
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   21
         Top             =   1200
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   525
            TabIndex        =   35
            Top             =   600
            Width           =   555
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   4
            Left            =   600
            Picture         =   "frmMenuSimples.frx":2691C
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas"
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
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.PictureBox pcClientes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   2160
      ScaleHeight     =   2295
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
         Index           =   8
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   8
            Left            =   720
            Picture         =   "frmMenuSimples.frx":275E6
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
            TabIndex        =   42
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
         Index           =   1
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   12
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clientes"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   34
            Top             =   600
            Width           =   585
         End
         Begin VB.Image imgIco 
            Height          =   450
            Index           =   1
            Left            =   720
            MouseIcon       =   "frmMenuSimples.frx":282B0
            Picture         =   "frmMenuSimples.frx":29C32
            Top             =   120
            Width           =   435
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   690
      End
   End
   Begin VB.PictureBox pcProduto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
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
         Index           =   22
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   4
         Top             =   2160
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Etiquetas"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   555
            TabIndex        =   33
            Top             =   600
            Width           =   675
         End
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   22
            Left            =   720
            MouseIcon       =   "frmMenuSimples.frx":2A19C
            Picture         =   "frmMenuSimples.frx":2B7E6
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
         Index           =   21
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   360
            Index           =   21
            Left            =   720
            MouseIcon       =   "frmMenuSimples.frx":2BD3D
            Picture         =   "frmMenuSimples.frx":2D387
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seções"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   645
            TabIndex        =   32
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
         Index           =   0
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   2
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produtos"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   31
            Top             =   600
            Width           =   645
         End
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "frmMenuSimples.frx":2D765
            Picture         =   "frmMenuSimples.frx":2EDAF
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.PictureBox pcCalend 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9930
      Left            =   12795
      ScaleHeight     =   9930
      ScaleWidth      =   2580
      TabIndex        =   48
      Top             =   0
      Width           =   2580
      Begin VB.PictureBox pcBotoesUsu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1935
         ScaleWidth      =   2415
         TabIndex        =   90
         Top             =   7200
         Width           =   2415
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEDFD5&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   7
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   1935
            TabIndex        =   93
            Top             =   960
            Width           =   1935
            Begin VB.Image imgIco 
               Height          =   405
               Index           =   7
               Left            =   720
               Picture         =   "frmMenuSimples.frx":2FA79
               Top             =   120
               Width           =   420
            End
            Begin VB.Label lblBotao 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sair - F12"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   570
               TabIndex        =   94
               Top             =   600
               Width           =   705
            End
         End
         Begin VB.PictureBox pcBotao 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEDFD5&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   6
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   1935
            TabIndex        =   91
            Top             =   0
            Width           =   1935
            Begin VB.Image imgIco 
               Height          =   435
               Index           =   6
               Left            =   720
               Picture         =   "frmMenuSimples.frx":2FFFF
               Top             =   120
               Width           =   390
            End
            Begin VB.Label lblBotao 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Usuário"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   6
               Left            =   645
               TabIndex        =   92
               Top             =   600
               Width           =   555
            End
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   2415
         TabIndex        =   68
         Top             =   5760
         Width           =   2415
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   1440
            TabIndex        =   105
            Top             =   480
            Width           =   360
         End
         Begin VB.Label lblNivelUsu 
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
            Left            =   1440
            TabIndex        =   104
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conexão"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   960
            Width           =   630
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
            Left            =   840
            TabIndex        =   74
            Top             =   960
            Width           =   675
         End
         Begin VB.Image imgUsu 
            Height          =   180
            Left            =   2040
            MouseIcon       =   "frmMenuSimples.frx":3057B
            Picture         =   "frmMenuSimples.frx":31EFD
            Top             =   0
            Width           =   210
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   45
            Width           =   585
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
            TabIndex        =   72
            Top             =   240
            Width           =   1995
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
            TabIndex        =   71
            Top             =   240
            Width           =   405
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
            TabIndex        =   70
            Top             =   765
            Width           =   675
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Visão"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   525
            Width           =   390
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   0
         Picture         =   "frmMenuSimples.frx":31FB2
         ScaleHeight     =   3015
         ScaleWidth      =   2655
         TabIndex        =   60
         Top             =   2760
         Width           =   2655
         Begin VB.PictureBox pcEditCaixa 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2040
            Picture         =   "frmMenuSimples.frx":4C6E4
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   65
            Top             =   240
            Width           =   285
         End
         Begin VB.Label lblImpressoraNFCe 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   240
            TabIndex        =   100
            Top             =   2520
            Width           =   225
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impressora de NFCe:"
            Height          =   195
            Left            =   240
            TabIndex        =   99
            Top             =   2280
            Width           =   1485
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impressora de Doc:"
            Height          =   195
            Left            =   240
            TabIndex        =   98
            Top             =   1800
            Width           =   1380
         End
         Begin VB.Label lblImpressora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   240
            TabIndex        =   97
            Top             =   2040
            Width           =   225
         End
         Begin VB.Label lblhost 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblHost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   240
            TabIndex        =   77
            Top             =   840
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Computador/Host:"
            Height          =   195
            Left            =   240
            TabIndex        =   76
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label lblCaixa 
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
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   64
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caixa:"
            Height          =   195
            Left            =   240
            TabIndex        =   63
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label lblUs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP:"
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   120
            Width           =   195
         End
         Begin VB.Label lblIp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.PictureBox pcLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   120
         Picture         =   "frmMenuSimples.frx":4C83D
         ScaleHeight     =   2655
         ScaleWidth      =   2415
         TabIndex        =   29
         Top             =   0
         Width           =   2415
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
      TabIndex        =   30
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
Public Sub mudaVariaveis()
    lblIp.caption = strIP_Usu
    lblhost.caption = strHost_Usu
    lblCaixa.caption = strCod_Cai
    lblUsu.caption = strNome_Usu
    lblCodUsu.caption = strCod_Usu
    lblVisao.caption = strVisao
    lblBanco.caption = Conn.DefaultDatabase
    'lblECF.caption = strImprFiscal
    lblImpressora.caption = Printer.deviceName
    lblImpressoraNFCe.caption = ImpressoraPadraoCupomNFCe.deviceName
    If strNivel_Usu = "1" Then
        lblNivelUsu.caption = "Operador"
        
    ElseIf strNivel_Usu = "2" Then
        lblNivelUsu.caption = "Supervisor"
    ElseIf strNivel_Usu = "3" Then
        lblNivelUsu.caption = "Gerente"
    ElseIf strNivel_Usu = "4" Then
        lblNivelUsu.caption = "Gestor"
    Else
        lblNivelUsu.caption = ""
    End If
    
    
End Sub

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
    For iCount = 0 To pcBotao.Count - 1
        lblBotao(iCount).FontBold = False
        lblBotao(iCount).FontUnderline = False
        
    Next
    lblBotao(Indice).FontBold = True
    lblBotao(Indice).FontUnderline = True
trata_ERRO:
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Dim strSQL As String
    
    strSQL = LCase(strNomeBanco)
    If Left(strSQL, 8) <> "requinte" Then
        
       ' pcLogo.Picture = Nothing
        
        pcLogo.Picture = LoadPicture(App.path & "\pic\logo.jpg")
        
    End If
    mudaVariaveis
    resizeForm Me
   mudaCor strCorFundoUsu
End Sub
Sub mudaCor(Optional cor = &HEEDFD5)
  On Error Resume Next
     MudaCorFundo Me, cor
    MudaCorFundo pcCaixa, cor
    MudaCorFundo pcClientes, cor
    MudaCorFundo pcCompras, cor
    MudaCorFundo pcConsign, cor
    MudaCorFundo pcPacote, cor
    MudaCorFundo pcProduto, cor
    MudaCorFundo pcVendas, cor
    MudaCorFundo pcFinanceiro, cor
    MudaCorFundo pcPac, cor
    
    
    Dim i As Integer
    For i = 0 To 99
        MudaCorFundo pcBotao(i), cor
    Next
End Sub
Public Sub MudaCorFundo(objeto As Object, cor)
    On Error Resume Next
    objeto.BackColor = cor
    
End Sub
Private Sub Form_Load()
    
    Me.height = mdiFrm.ScaleHeight - mdiFrm.pcStatus.height
    Me.width = mdiFrm.ScaleWidth - 100
    Me.top = 100
    
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mdiFrm.mnOptCor
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    P_mensagem
End Sub

Private Sub Form_Resize()
    pcBotoesUsu.top = Me.height - pcBotoesUsu.height - 100
End Sub

Private Sub imgIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub imgIco_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    pcBotao_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub lblBanco_Click()
    Dim sBanco As String
    
    sBanco = InputBox("Conectar a", "Conexão", sBanco)
    
    If F_AbreBanco(sBanco) Then
        Exit Sub
    End If
    
    
    
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

Private Sub lblVisao_Click()
    frmEscolherVisao.Show 1
End Sub



Private Sub pcBotao_Click(Index As Integer)
    If Index = 0 Then
        frmEdiPro.Show 0
        frmEdiPro.SetFocus
    ElseIf Index = 1 Then
        frmEdiClientes.Show 0
        frmEdiClientes.SetFocus
    ElseIf Index = 2 Then
        frmEntradasSped.Show 0
        frmEntradasSped.SetFocus
    ElseIf Index = 3 Then
        novaVenda
    ElseIf Index = 4 Then
        frmEdiVendas.Show 0, mdiFrm
    ElseIf Index = 5 Then
        frmEdiCai.Show 0
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
        frmConsign.Show 0
        frmConsign.SetFocus
    ElseIf Index = 10 Then
        If strCreditario_Usu = "S" Then
            frmCrediario.Show 0
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
            
        End If
    ElseIf Index = 11 Then
    
        frmTransfConsign.Show 0
    ElseIf Index = 12 Then
    
        frmDevolucaoConsign.Show 0
    ElseIf Index = 13 Then
        frmAcerto.acerto_promocao = False
        frmAcerto.Show 0
        
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
        frmEdiNFe.Show 1
    ElseIf Index = 20 Then
        frmCobranca.Show 0
    ElseIf Index = 21 Then
        secao.P_EditaSecao
    ElseIf Index = 22 Then
        frmGeraEtiquetas.Show 0
        frmGeraEtiquetas.SetFocus

    ElseIf Index = 23 Then
        frmEdiFun.Show 0
        frmEdiFun.SetFocus
    ElseIf Index = 24 Then
        frmEntradas.Show 0
        frmEntradas.SetFocus
    ElseIf Index = 25 Then
        frmMovDiarioCx.Show 1
    ElseIf Index = 26 Then
        frmECF.Show 0
    ElseIf Index = 27 Then
        frmControleChequesPagar.Show 0
    ElseIf Index = 28 Then
        frmRelComissGeral.Show 1
        
    ElseIf Index = 29 Then
        frmPacote.Show 0
    ElseIf Index = 30 Then
        frmAcerto.acerto_promocao = True
        frmAcerto.Show 0
    ElseIf Index = 33 Then
        If strAbrenota_usu = "S" Then
            frmEdiNFe.Show 0
            frmEdiNFe.SetFocus
        Else
            MsgBox "Você não tem permissão para este acesso!", vbExclamation
        End If
        
    End If
End Sub
Sub novaVenda()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = Vendas.F_GeraNovaVenda("00000")
    
    
    
    frmVenda.caption = strNome_Emp & " - Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    
    frmVenda.Show 0, mdiFrm
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

Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
    
    
    
    P_mensagem Index
    
    
    
End Sub

Sub P_mensagem(Optional Index As Integer = -1)
    
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
    frmCadCaixa.Show 1
    CarregarParametros
    carrega_caixa
    mudaVariaveis
    Me.Refresh
End Sub

Private Sub pcProduto_GotFocus()
 '   pcProduto.BorderStyle = 1
End Sub

Private Sub pcProduto_LostFocus()
  '  pcProduto.BorderStyle = 0
End Sub

