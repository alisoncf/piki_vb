VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00755C33&
   BorderStyle     =   0  'None
   Caption         =   "Área de Trabalho"
   ClientHeight    =   9930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Menu de Opções"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   4320
      ScaleHeight     =   1185
      ScaleWidth      =   2025
      TabIndex        =   25
      Top             =   5400
      Width           =   2055
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
         TabIndex        =   27
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pacote"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   540
            TabIndex        =   76
            Top             =   600
            Width           =   525
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   11
            Left            =   600
            Picture         =   "frmMenu2.frx":0000
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2160
      ScaleHeight     =   1185
      ScaleWidth      =   2025
      TabIndex        =   18
      Top             =   5400
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
         TabIndex        =   20
         Top             =   240
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   10
            Left            =   720
            Picture         =   "frmMenu2.frx":0CCA
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Crediário"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   600
            TabIndex        =   75
            Top             =   600
            Width           =   645
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crediário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   8640
      ScaleHeight     =   6585
      ScaleWidth      =   2025
      TabIndex        =   34
      Top             =   0
      Width           =   2055
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
         TabIndex        =   82
         Top             =   4080
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   27
            Left            =   615
            Picture         =   "frmMenu2.frx":1594
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
            TabIndex        =   83
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
         TabIndex        =   36
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
            TabIndex        =   74
            Top             =   600
            Width           =   1305
         End
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   15
            Left            =   720
            Picture         =   "frmMenu2.frx":225E
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
         TabIndex        =   37
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   16
            Left            =   720
            Picture         =   "frmMenu2.frx":24E2
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
            TabIndex        =   73
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
         TabIndex        =   38
         Top             =   2160
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   17
            Left            =   600
            Picture         =   "frmMenu2.frx":28F9
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
            TabIndex        =   72
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
         TabIndex        =   39
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
            TabIndex        =   71
            Top             =   600
            Width           =   1455
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   18
            Left            =   615
            Picture         =   "frmMenu2.frx":31C3
            Top             =   90
            Width           =   480
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   20
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   40
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
            TabIndex        =   70
            Top             =   600
            Width           =   705
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   20
            Left            =   600
            Picture         =   "frmMenu2.frx":3E8D
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
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   10800
      ScaleHeight     =   6585
      ScaleWidth      =   2025
      TabIndex        =   41
      Top             =   0
      Width           =   2055
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
         TabIndex        =   80
         Top             =   2400
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monitoramento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   330
            TabIndex        =   81
            Top             =   600
            Width           =   1065
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   19
            Left            =   600
            Picture         =   "frmMenu2.frx":4B57
            Top             =   90
            Width           =   480
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
         TabIndex        =   78
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   23
            Left            =   600
            Picture         =   "frmMenu2.frx":5821
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
            Left            =   345
            TabIndex        =   79
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
         Index           =   6
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   43
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuário"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   525
            TabIndex        =   69
            Top             =   600
            Width           =   555
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   6
            Left            =   600
            Picture         =   "frmMenu2.frx":64EB
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
         Index           =   7
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   44
         Top             =   5520
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sair"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   660
            TabIndex        =   68
            Top             =   600
            Width           =   285
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   7
            Left            =   600
            Picture         =   "frmMenu2.frx":71B5
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   6480
      ScaleHeight     =   6585
      ScaleWidth      =   2025
      TabIndex        =   28
      Top             =   0
      Width           =   2055
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
         TabIndex        =   33
         Top             =   3120
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acerto"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   615
            TabIndex        =   67
            Top             =   600
            Width           =   495
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   14
            Left            =   600
            Picture         =   "frmMenu2.frx":7E7F
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
         Index           =   13
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   32
         Top             =   2160
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Devolução"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   465
            TabIndex        =   66
            Top             =   600
            Width           =   795
         End
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   13
            Left            =   720
            Picture         =   "frmMenu2.frx":8749
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
         Index           =   12
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   31
         Top             =   1200
         Width           =   1815
         Begin VB.Image imgIco 
            Height          =   240
            Index           =   12
            Left            =   720
            Picture         =   "frmMenu2.frx":8CD3
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transferência"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   65
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
         Index           =   9
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   1815
         TabIndex        =   30
         Top             =   240
         Width           =   1815
         Begin VB.Label lblBotao 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consignações"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   405
            TabIndex        =   64
            Top             =   600
            Width           =   1035
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   9
            Left            =   720
            Picture         =   "frmMenu2.frx":925D
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3105
      ScaleWidth      =   2025
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
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
            TabIndex        =   62
            Top             =   600
            Width           =   315
         End
         Begin VB.Image imgIco 
            Height          =   420
            Index           =   26
            Left            =   600
            Picture         =   "frmMenu2.frx":9B27
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
            Picture         =   "frmMenu2.frx":A4AC
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
            TabIndex        =   61
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
            TabIndex        =   60
            Top             =   600
            Width           =   975
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   5
            Left            =   600
            Picture         =   "frmMenu2.frx":AA01
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
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   2160
      ScaleHeight     =   2865
      ScaleWidth      =   2025
      TabIndex        =   14
      Top             =   2400
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
            Picture         =   "frmMenu2.frx":B6CB
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
            TabIndex        =   59
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
            TabIndex        =   58
            Top             =   600
            Width           =   645
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   2
            Left            =   600
            Picture         =   "frmMenu2.frx":C0ED
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
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   4320
      ScaleHeight     =   5265
      ScaleWidth      =   2025
      TabIndex        =   21
      Top             =   0
      Width           =   2055
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
         TabIndex        =   23
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
            TabIndex        =   57
            Top             =   600
            Width           =   915
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   3
            Left            =   600
            Picture         =   "frmMenu2.frx":CDB7
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
         TabIndex        =   24
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
            TabIndex        =   56
            Top             =   600
            Width           =   555
         End
         Begin VB.Image imgIco 
            Height          =   480
            Index           =   4
            Left            =   600
            Picture         =   "frmMenu2.frx":D681
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
         ForeColor       =   &H00EF9C7C&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.PictureBox pcClientes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   2160
      ScaleHeight     =   2265
      ScaleWidth      =   2025
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
            Picture         =   "frmMenu2.frx":E34B
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
            TabIndex        =   63
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
            TabIndex        =   55
            Top             =   600
            Width           =   585
         End
         Begin VB.Image imgIco 
            Height          =   450
            Index           =   1
            Left            =   720
            MouseIcon       =   "frmMenu2.frx":F015
            Picture         =   "frmMenu2.frx":10997
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
         ForeColor       =   &H00EF9C7C&
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
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   2025
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
            Caption         =   "Subseções"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   510
            TabIndex        =   54
            Top             =   600
            Width           =   825
         End
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   22
            Left            =   720
            MouseIcon       =   "frmMenu2.frx":10F01
            Picture         =   "frmMenu2.frx":1254B
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
            MouseIcon       =   "frmMenu2.frx":12AA2
            Picture         =   "frmMenu2.frx":140EC
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
            TabIndex        =   53
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
            TabIndex        =   52
            Top             =   600
            Width           =   645
         End
         Begin VB.Image imgIco 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "frmMenu2.frx":144CA
            Picture         =   "frmMenu2.frx":15B14
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
         ForeColor       =   &H00EF9C7C&
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
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9930
      Left            =   12690
      ScaleHeight     =   9930
      ScaleWidth      =   2460
      TabIndex        =   77
      Top             =   0
      Width           =   2460
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   0
         Picture         =   "frmMenu2.frx":167DE
         ScaleHeight     =   2865
         ScaleWidth      =   2385
         TabIndex        =   45
         Top             =   0
         Width           =   2415
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1080
            Picture         =   "frmMenu2.frx":20B8D
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   46
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblDia"
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
            Left            =   885
            TabIndex        =   50
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblNumDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   825
            Left            =   1020
            TabIndex        =   49
            Top             =   795
            Width           =   375
         End
         Begin VB.Label lblMes 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês"
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
            Left            =   960
            TabIndex        =   48
            Top             =   1560
            Width           =   465
         End
         Begin VB.Label lblAno 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2007"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   840
            TabIndex        =   47
            Top             =   1875
            Width           =   735
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
      TabIndex        =   51
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
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    
    

    
    On Error Resume Next
    For iCount = 0 To pcBotao.count - 1
        lblBotao(iCount).FontBold = False
        lblBotao(iCount).FontUnderline = False
        
    Next
    lblBotao(Indice).FontBold = True
    lblBotao(Indice).FontUnderline = True
trata_Erro:
End Sub

Private Sub Form_Activate()
    lblDia.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    lblNumDia.caption = Day(Date)
    lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    lblAno.caption = Year(Date)
End Sub

Private Sub Form_Load()
    Me.Height = mdiFrm.ScaleHeight - mdiFrm.pcStatus.Height
    Me.Width = mdiFrm.ScaleWidth - 100
    Me.Top = 100
    
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    P_mensagem
End Sub

Private Sub imgIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub imgIco_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    pcBotao_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcBotao_Click(Index As Integer)
    If Index = 0 Then
        frmEdiPro.Show 0
    ElseIf Index = 1 Then
        frmEdiClientes.Show 0
    ElseIf Index = 2 Then
        frmEdiEntr.Show 0
    ElseIf Index = 3 Then
        novaVenda
    ElseIf Index = 4 Then
        frmEdiVendas.Show 0
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
        frmCrediario.Show 0
    ElseIf Index = 11 Then
    
        frmPacote.Show 1
    ElseIf Index = 12 Then
    
        frmTransfConsign.Show 0
    ElseIf Index = 13 Then
    
        frmDevolucaoConsign.Show 0
        
    ElseIf Index = 14 Then
    
        frmAcerto.Show 0
    ElseIf Index = 15 Then
        frmEditaMov.Show 0
    ElseIf Index = 16 Then
        frmContasPagar.Show 0
    ElseIf Index = 17 Then
        frmEdiMovFin.Show 0
    
    ElseIf Index = 18 Then
        frmControleChequesReceber.Show 0
    ElseIf Index = 19 Then
        frmMonitoramentoCaixas.Show 0
    ElseIf Index = 20 Then
        frmCobranca.Show 0
    ElseIf Index = 21 Then
        secao.P_EditaSecao
    ElseIf Index = 22 Then
            frmEdiSubSec.Show 0
    frmEdiSubSec.SetFocus

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
        If Index < pcBotao.count - 1 Then
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

Private Sub pcProduto_GotFocus()
    pcProduto.BorderStyle = 1
End Sub

Private Sub pcProduto_LostFocus()
    pcProduto.BorderStyle = 0
End Sub
