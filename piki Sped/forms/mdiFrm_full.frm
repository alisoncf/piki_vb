VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiFrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00C0C0C0&
   Caption         =   "Piki - 2020"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   -135
   ClientWidth     =   15240
   Icon            =   "mdiFrm_full.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   15210
      TabIndex        =   50
      Top             =   9300
      Width           =   15240
      Begin VB.PictureBox pcHora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   11160
         ScaleHeight     =   615
         ScaleWidth      =   4695
         TabIndex        =   53
         Top             =   0
         Width           =   4695
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   360
            Picture         =   "mdiFrm_full.frx":3282
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   58
            Top             =   45
            Width           =   240
         End
         Begin VB.Label lblConexao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Conexão"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   0
            TabIndex        =   71
            Top             =   360
            Width           =   570
         End
         Begin VB.Label lblAno 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2007"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   3840
            TabIndex        =   57
            Top             =   0
            Width           =   675
         End
         Begin VB.Label lblMes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dezembro"
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
            Left            =   2760
            TabIndex        =   56
            Top             =   30
            Width           =   975
         End
         Begin VB.Label lblNumDia 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   2280
            TabIndex        =   55
            Top             =   -15
            Width           =   330
         End
         Begin VB.Label lblDia 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Segunda-Feira"
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
            Left            =   720
            TabIndex        =   54
            Top             =   30
            Width           =   1410
         End
      End
      Begin VB.PictureBox pcUsu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "mdiFrm_full.frx":35BA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   51
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblEmpresa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   480
         TabIndex        =   70
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Left            =   9165
         TabIndex        =   69
         Top             =   0
         Width           =   195
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         Height          =   195
         Left            =   480
         TabIndex        =   52
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.PictureBox pcMax 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9300
      Left            =   2295
      Picture         =   "mdiFrm_full.frx":3E84
      ScaleHeight     =   9270
      ScaleWidth      =   150
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pcMenu 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9300
      Left            =   0
      ScaleHeight     =   9270
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.PictureBox pcVendas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   59
         Top             =   4920
         Width           =   1935
         Begin VB.PictureBox pcPesqVenda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_full.frx":4CE5
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   64
            Top             =   720
            Width           =   240
         End
         Begin VB.PictureBox pcNovaVenda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_full.frx":56E7
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   63
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox Picture16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   61
            Top             =   0
            Width           =   1935
            Begin VB.Label lblVendas 
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
               Height          =   195
               Left            =   240
               TabIndex        =   62
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.PictureBox Picture3 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1680
            Picture         =   "mdiFrm_full.frx":584B
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   60
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblPesqVenda 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   66
            Top             =   765
            Width           =   690
         End
         Begin VB.Label lblNovaVenda 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   65
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.PictureBox pcSaidas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   42
         Top             =   3720
         Width           =   1935
         Begin VB.PictureBox Picture24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   45
            Top             =   0
            Width           =   1935
            Begin VB.Label lblSaidas 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saidas"
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
               Left            =   240
               TabIndex        =   46
               Top             =   0
               Width           =   585
            End
         End
         Begin VB.PictureBox pcNovaSaida 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_full.frx":5BD5
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   44
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox pcPesqSaida 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_full.frx":5D39
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   43
            Top             =   720
            Width           =   240
         End
         Begin VB.Label lblNovaSaida 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   48
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblPesqSaida 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   47
            Top             =   765
            Width           =   690
         End
      End
      Begin VB.PictureBox pcClientes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000000&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   27
         Top             =   120
         Width           =   1935
         Begin VB.PictureBox pcNovoCli 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_full.frx":673B
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   29
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox pcPesqCli 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   0
            Picture         =   "mdiFrm_full.frx":689F
            ScaleHeight     =   450
            ScaleWidth      =   435
            TabIndex        =   28
            Top             =   720
            Width           =   435
         End
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   30
            Top             =   0
            Width           =   1935
            Begin VB.Label lblClientes 
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
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Label lblNovoCli 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   33
            Top             =   405
            Width           =   390
         End
         Begin VB.Label lblPesqCli 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   32
            Top             =   765
            Width           =   690
         End
      End
      Begin VB.PictureBox pcMin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   2100
         Picture         =   "mdiFrm_full.frx":6E09
         ScaleHeight     =   2265
         ScaleWidth      =   135
         TabIndex        =   26
         Top             =   0
         Width           =   165
      End
      Begin VB.PictureBox pcEntradas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   23
         Top             =   2520
         Width           =   1935
         Begin VB.PictureBox pcPesqEntrada 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   82
            Picture         =   "mdiFrm_full.frx":752C
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   36
            Top             =   720
            Width           =   360
         End
         Begin VB.PictureBox pcNovaEntrada 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_full.frx":790A
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   34
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   24
            Top             =   0
            Width           =   1935
            Begin VB.Label lblEntradas 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Entradas"
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
               Left            =   240
               TabIndex        =   25
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.Label lblPesqEntrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   37
            Top             =   765
            Width           =   690
         End
         Begin VB.Label lblNovaEntrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   405
            Width           =   390
         End
      End
      Begin VB.PictureBox pcProdutos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
         Begin VB.PictureBox pcPesqPro 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            Picture         =   "mdiFrm_full.frx":7A6E
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   40
            Top             =   720
            Width           =   345
         End
         Begin VB.PictureBox pcNovoPro 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_full.frx":7EF8
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   38
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   21
            Top             =   0
            Width           =   1935
            Begin VB.Label lblProdutos 
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
               Height          =   195
               Left            =   240
               TabIndex        =   22
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.Label lblPesqPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   41
            Top             =   765
            Width           =   690
         End
         Begin VB.Label lblNovoPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   39
            Top             =   405
            Width           =   390
         End
      End
      Begin VB.PictureBox pcFinanceiro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2145
         ScaleWidth      =   1905
         TabIndex        =   6
         Top             =   6120
         Width           =   1935
         Begin VB.PictureBox pcMov 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_full.frx":805C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   11
            Top             =   360
            Width           =   240
         End
         Begin VB.PictureBox pcCpag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_full.frx":8473
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   10
            Top             =   720
            Width           =   240
         End
         Begin VB.PictureBox pcChPag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   113
            Picture         =   "mdiFrm_full.frx":86F7
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   8
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox pcChRec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   113
            Picture         =   "mdiFrm_full.frx":8A05
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   7
            Top             =   1800
            Width           =   255
         End
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   12
            Top             =   0
            Width           =   1935
            Begin VB.Label lblFinanceiro 
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
               Height          =   195
               Left            =   240
               TabIndex        =   13
               Top             =   0
               Width           =   900
            End
         End
         Begin VB.PictureBox pcCRec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   98
            Picture         =   "mdiFrm_full.frx":8D13
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   9
            Top             =   1080
            Width           =   285
         End
         Begin VB.Label lblMov 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentos"
            Height          =   195
            Left            =   480
            TabIndex        =   19
            Top             =   383
            Width           =   855
         End
         Begin VB.Label lblContasPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Pagar"
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblContasRec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber"
            Height          =   195
            Left            =   480
            TabIndex        =   17
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   0
            Width           =   1290
         End
         Begin VB.Label lblChPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Pagar"
            Height          =   195
            Left            =   480
            TabIndex        =   15
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label lblChRec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Receb."
            Height          =   195
            Left            =   480
            TabIndex        =   14
            Top             =   1800
            Width           =   1335
         End
      End
      Begin VB.PictureBox pcSair 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   705
         ScaleWidth      =   1905
         TabIndex        =   1
         Top             =   8280
         Width           =   1935
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   3
            Top             =   0
            Width           =   1935
            Begin VB.Label lblSair 
               AutoSize        =   -1  'True
               BackColor       =   &H000000C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sair - F12"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   195
               Left            =   240
               TabIndex        =   4
               Top             =   0
               Width           =   840
            End
         End
         Begin VB.PictureBox picSair 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            Picture         =   "mdiFrm_full.frx":8E6F
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   2
            Top             =   398
            Width           =   195
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sair"
            Height          =   195
            Left            =   480
            TabIndex        =   5
            Top             =   398
            Width           =   270
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         Picture         =   "mdiFrm_full.frx":8FDA
         ScaleHeight     =   1095
         ScaleWidth      =   2295
         TabIndex        =   67
         Top             =   9000
         Width           =   2295
         Begin VB.Label lblBD 
            Caption         =   "..."
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   840
            Width           =   1935
         End
      End
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   4320
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnPrinc 
      Caption         =   "&Principal"
      Begin VB.Menu mnEditCli 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu mnEDiCid 
         Caption         =   "&Cidades"
      End
      Begin VB.Menu div6 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnForn 
         Caption         =   "&Fornecedores"
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImo 
         Caption         =   "Imobilizado"
         Visible         =   0   'False
      End
      Begin VB.Menu divArq2 
         Caption         =   "-"
      End
      Begin VB.Menu mnProdutos 
         Caption         =   "&Produtos"
         Begin VB.Menu mnProd 
            Caption         =   "&Produtos"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnHistProd 
            Caption         =   "&Histórico de Produtos"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnSecao 
            Caption         =   "Seções"
         End
         Begin VB.Menu mnSubSec 
            Caption         =   "Subseções"
         End
         Begin VB.Menu mnCadCat 
            Caption         =   "C&ategorias"
         End
         Begin VB.Menu mnSubcat 
            Caption         =   "Subcategorias"
         End
         Begin VB.Menu mnCadMar 
            Caption         =   "&Marcas"
         End
         Begin VB.Menu mnSimilares 
            Caption         =   "Similares"
         End
         Begin VB.Menu mnTransfComp 
            Caption         =   "&Transferência de Marcas/Categorias"
         End
         Begin VB.Menu mnUnidades 
            Caption         =   "Unidades"
         End
      End
      Begin VB.Menu mnServ 
         Caption         =   "Servicos"
      End
      Begin VB.Menu mnTipoServ 
         Caption         =   "Tipos de Serviço"
      End
      Begin VB.Menu div4 
         Caption         =   "-"
      End
      Begin VB.Menu mnVeiculos 
         Caption         =   "Veiculos"
      End
      Begin VB.Menu mnVeicFab 
         Caption         =   "Fabricantes de Veículos"
      End
      Begin VB.Menu divvei 
         Caption         =   "-"
      End
      Begin VB.Menu mnCadCargos 
         Caption         =   "Ca&rgos"
      End
      Begin VB.Menu mnFun 
         Caption         =   "&Funcionários"
      End
      Begin VB.Menu mnEdiUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnDiv6 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair do Piki Box"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnEntradas 
      Caption         =   "&Entrada de Produtos"
      Begin VB.Menu mnIncMer 
         Caption         =   "&Nova Compra"
      End
      Begin VB.Menu mnEdiComp 
         Caption         =   "&Edição de Compras"
      End
      Begin VB.Menu mnEntrada 
         Caption         =   "E&ntrada de Produtos"
         Visible         =   0   'False
      End
      Begin VB.Menu DivPed 
         Caption         =   "-"
      End
      Begin VB.Menu mnSujComp 
         Caption         =   "&Novo Processo de Compras"
      End
      Begin VB.Menu mnEdiProComp 
         Caption         =   "Editar processos de compra"
      End
      Begin VB.Menu div10 
         Caption         =   "-"
      End
      Begin VB.Menu mnEdiDev 
         Caption         =   "&Devoluções"
      End
      Begin VB.Menu mnDEv 
         Caption         =   "Nova Dev&olução"
      End
      Begin VB.Menu mnDivVen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnReposicao 
         Caption         =   "Reposição de Estoque"
      End
   End
   Begin VB.Menu mnOrc 
      Caption         =   "&Orçamentos/vendas"
      Begin VB.Menu mnVendaAtac 
         Caption         =   "&Nova Venda"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnEdiVenAtac 
         Caption         =   "&Edição de Vendas"
      End
      Begin VB.Menu mnBuscaOrc 
         Caption         =   "&Busca de Orçamentos"
      End
      Begin VB.Menu mnNovoOrc 
         Caption         =   "&Novo Orçamento"
      End
      Begin VB.Menu mnEntregas 
         Caption         =   "&Entregas"
         Shortcut        =   ^E
      End
      Begin VB.Menu divOrc1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnSaidas 
      Caption         =   "&Saídas"
      Begin VB.Menu mnEdiNotaSai 
         Caption         =   "Edição de Notas Fiscais"
      End
      Begin VB.Menu mnCFOP 
         Caption         =   "&Lista de CFOP"
      End
      Begin VB.Menu divSai1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEdiMovInt 
         Caption         =   "Edição de &Movimento Interno"
      End
      Begin VB.Menu mnNovoMovInt 
         Caption         =   "Novo Movimento &Interno"
      End
      Begin VB.Menu mnNFDev 
         Caption         =   "&Nova Fiscal Avulsa/Devolução"
      End
   End
   Begin VB.Menu mnFinan 
      Caption         =   "&Financeiro"
      Begin VB.Menu mnEfetVen 
         Caption         =   "&Efetivação de Vendas"
      End
      Begin VB.Menu mnDupRec 
         Caption         =   "Duplicatas a Receber"
      End
      Begin VB.Menu mnCheq 
         Caption         =   "Cheques a Receber"
      End
      Begin VB.Menu mnContasChequesPagar 
         Caption         =   "&Duplicatas a Pagar"
      End
      Begin VB.Menu mnChPag 
         Caption         =   "Cheques a Pagar"
      End
      Begin VB.Menu mnMovFin 
         Caption         =   "&Movimento Financeiro"
      End
      Begin VB.Menu mnContasPagar 
         Caption         =   "&Edição de Documentos"
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu mnCalGuias 
         Caption         =   "&Lançamento de Documentos"
      End
      Begin VB.Menu mnEdiMovFin 
         Caption         =   "Movimento de &Financeiras"
      End
      Begin VB.Menu mnCons 
         Caption         =   "&Consolidação Diária Financeira"
      End
      Begin VB.Menu mnVinc 
         Caption         =   "&Vinculação de Documentos"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnRelBaixas 
         Caption         =   "Baixas"
      End
      Begin VB.Menu div11 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnBan 
         Caption         =   "&Bancos"
      End
      Begin VB.Menu mnCon 
         Caption         =   "&Contas"
      End
      Begin VB.Menu mnFinanceiras 
         Caption         =   "&Financeiras"
      End
      Begin VB.Menu mnCadPla 
         Caption         =   "&Plano de Receitas e Despesas"
      End
      Begin VB.Menu mnBol 
         Caption         =   "&Boletos"
      End
      Begin VB.Menu mnMonit 
         Caption         =   "&Monitoramento de Caixas"
      End
      Begin VB.Menu mnMovUsu 
         Caption         =   "&Movimentação por Usuário"
      End
      Begin VB.Menu mnEdiCaixas 
         Caption         =   "Edição de &Caixas"
      End
   End
   Begin VB.Menu mnCaixa 
      Caption         =   "Cai&xas"
      Begin VB.Menu mnLan 
         Caption         =   "&Lançamento de Depesa/Receita"
      End
      Begin VB.Menu mnAbreCx 
         Caption         =   "&Abertura de Caixa"
      End
      Begin VB.Menu mnLanCaixa 
         Caption         =   "&Edição de Lançamentos"
      End
      Begin VB.Menu mnFechaCaixa 
         Caption         =   "Fechamento de Caixa"
      End
      Begin VB.Menu mnEcf 
         Caption         =   "EC&F"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnRedZ 
         Caption         =   "Reduções Z"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnSped 
      Caption         =   "&NFe/SPED"
      Begin VB.Menu mnNfe 
         Caption         =   "Notas Fiscais Eletrônicas"
      End
      Begin VB.Menu mnInfAdd 
         Caption         =   "Informações Adicionais "
      End
      Begin VB.Menu mnEmpresaSped 
         Caption         =   "&Empresas"
      End
      Begin VB.Menu mnParticipantes 
         Caption         =   "Participantes"
      End
      Begin VB.Menu mnLanNotEnergiaEletrica 
         Caption         =   "Lançamento de Notas de Energia Elétrica e Água"
      End
      Begin VB.Menu mnLanRedZ 
         Caption         =   "Lançamento de Redução Z"
      End
      Begin VB.Menu mnListaIcms 
         Caption         =   "Lista de CST ICMS"
      End
      Begin VB.Menu mnListaIPI 
         Caption         =   "Lista CST IPI"
      End
      Begin VB.Menu mnListaCofins 
         Caption         =   "Lista CST Cofins"
      End
      Begin VB.Menu mnListaPis 
         Caption         =   "Lista CST PIS"
      End
   End
   Begin VB.Menu mnRel 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnRelVendas 
         Caption         =   "&Vendas"
         Begin VB.Menu mnRelVend 
            Caption         =   "Relatório de &Vendas"
         End
         Begin VB.Menu mnRelFreqVen 
            Caption         =   "Relatório de Frequência de Vendas"
         End
         Begin VB.Menu mnRelDev 
            Caption         =   "&Relatório de Devoluções"
         End
         Begin VB.Menu mnConsPre 
            Caption         =   "Consulta de Pr&eços"
         End
         Begin VB.Menu mnRelVendaBandeira 
            Caption         =   "Relatório de Vendas por Bandeiras (Cartões)"
         End
         Begin VB.Menu mnRelComissVen 
            Caption         =   "Relatório de Vendas por Vendedor"
         End
         Begin VB.Menu mnRelProdsNaoVendidos 
            Caption         =   "Relatório de Produtos não vendidos"
         End
      End
      Begin VB.Menu mnRelComp 
         Caption         =   "&Compras"
         Begin VB.Menu mnRelProdForn 
            Caption         =   "&Relatório de Produtos por Fornecedor"
         End
         Begin VB.Menu mnRelCmpFor 
            Caption         =   "Relatório de &Compras por Fornecedor"
         End
      End
      Begin VB.Menu div12 
         Caption         =   "-"
      End
      Begin VB.Menu mnRelFatur 
         Caption         =   "&Faturamento"
         Begin VB.Menu mnRelComparaCli 
            Caption         =   "Comparativo de Clientes"
         End
         Begin VB.Menu mnRelAcompMensal 
            Caption         =   "Acompanhamento Mensal"
         End
         Begin VB.Menu mnRelAcompMensalCliProd 
            Caption         =   "Acompanhamento mensal de clientes por produtos"
         End
         Begin VB.Menu mnRelFatVenda 
            Caption         =   "Relatório de &Faturamento por Vendas"
         End
         Begin VB.Menu mnRelAnaFat 
            Caption         =   "&Análise de Faturamento"
         End
         Begin VB.Menu mnRelAnaIndFluxo 
            Caption         =   "Relatório de Análise Direcionada de Fl&uxo"
         End
      End
      Begin VB.Menu mnRelEst 
         Caption         =   "&Análise de Estoque"
         Begin VB.Menu mnRelAltPre 
            Caption         =   "&Relatório de Alterações de Preços"
         End
         Begin VB.Menu mnRelProdLucr 
            Caption         =   "Realtório de Produtos por &Lucratividade"
         End
         Begin VB.Menu mnRelComport 
            Caption         =   "Relatório de Comportamento de Produtos"
         End
         Begin VB.Menu mnRelGiro 
            Caption         =   "Relatório de Giro Periódico"
         End
         Begin VB.Menu mnCompProCli 
            Caption         =   "Relatório de Comp. de Prod. p/ C&liente"
         End
         Begin VB.Menu mnCOmport 
            Caption         =   "Painel Mensal de Vendas"
         End
         Begin VB.Menu mnRelCompEst 
            Caption         =   "&Listagem de Comparativo de Estoque"
         End
         Begin VB.Menu mnRelLucrComp 
            Caption         =   "Lucratividade Por &Componente"
         End
         Begin VB.Menu mnDemoInv 
            Caption         =   "&Demonstrativo de Inventário"
         End
         Begin VB.Menu mnRelAltPreco 
            Caption         =   "Relatório de Alteração de Preços"
         End
      End
      Begin VB.Menu div13 
         Caption         =   "-"
      End
      Begin VB.Menu mnRelFinan 
         Caption         =   "&Financeiro"
         Begin VB.Menu mnResumo 
            Caption         =   "&Resumo de Documentos a Pagar"
         End
         Begin VB.Menu mnHistCheCPf 
            Caption         =   "&Histórico de Cheques por CPF"
         End
         Begin VB.Menu mnRelPagReceb 
            Caption         =   "Relatório de Contas a Pagar/Receber"
         End
         Begin VB.Menu mnRelChq 
            Caption         =   "Relatório de Cheques a Pagar/Receber"
         End
         Begin VB.Menu mnRelAtvCli 
            Caption         =   "Relatório de &Atividade de Clientes por Vendas"
         End
         Begin VB.Menu mNRelAtvPro 
            Caption         =   "Relatório de &Atividade de Clientes por Produtos"
         End
         Begin VB.Menu mnRecCli 
            Caption         =   "Relação de Recebimentos"
         End
         Begin VB.Menu mnRelReceitasDesp 
            Caption         =   "&Relatório de Receitas e Despesas"
         End
      End
      Begin VB.Menu mnInvent 
         Caption         =   "&Inventário de Conferência"
      End
      Begin VB.Menu mnImpInvent 
         Caption         =   "&Impressão de Inventário"
      End
   End
   Begin VB.Menu mnParam 
      Caption         =   "Parâ&metros"
      Begin VB.Menu mnAcBr 
         Caption         =   "AcBrMonitor"
      End
      Begin VB.Menu mnSendMail 
         Caption         =   "Configurações de Email"
      End
      Begin VB.Menu mnFecharMes 
         Caption         =   "Fechar Mês"
      End
      Begin VB.Menu mnTurnos 
         Caption         =   "Edição de &Turnos"
      End
      Begin VB.Menu mnAdm 
         Caption         =   "&Aministração do Sistema"
         Begin VB.Menu mnSQLPlu 
            Caption         =   "&SQL Plus - Console SQL"
         End
         Begin VB.Menu mnBAckUp 
            Caption         =   "&BackUp"
         End
         Begin VB.Menu mnTransf 
            Caption         =   "&Utilitário de Transferência de Dados"
         End
         Begin VB.Menu mnMigra 
            Caption         =   "&Utilitário de Migração"
         End
      End
      Begin VB.Menu mnConfigNF 
         Caption         =   "&Configuar Impressão NF"
      End
      Begin VB.Menu mnUtilECF 
         Caption         =   "&Utilitário de ECF"
      End
      Begin VB.Menu mnImp 
         Caption         =   "&Impostos"
      End
      Begin VB.Menu mnCont 
         Caption         =   "C&ontagens"
      End
      Begin VB.Menu mnCaixas 
         Caption         =   "&Caixas"
      End
      Begin VB.Menu mnDePara 
         Caption         =   "&De Para"
      End
      Begin VB.Menu mnCalMed 
         Caption         =   "&Cálculo de Média Diária"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnConfig 
         Caption         =   "&Configurações do Sistema"
      End
      Begin VB.Menu mnUsu 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnProm 
         Caption         =   "&Promoções"
      End
      Begin VB.Menu mnMescComp 
         Caption         =   "&Mesclar Componentes"
      End
      Begin VB.Menu mnMudSen 
         Caption         =   "&Mudar Senha"
      End
      Begin VB.Menu mnGeraEti 
         Caption         =   "&Gerar Etiquetas"
      End
      Begin VB.Menu div5 
         Caption         =   "-"
      End
      Begin VB.Menu mnFormasPag 
         Caption         =   "&Formas de Pagamento"
      End
      Begin VB.Menu mnApar 
         Caption         =   "&Aparência"
      End
      Begin VB.Menu mnCalc 
         Caption         =   "Ca&lculadora"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnCriarReplica 
         Caption         =   "Replicar"
      End
      Begin VB.Menu mnCidadesIBGE 
         Caption         =   "Cidades &IBGE"
      End
      Begin VB.Menu mnShell 
         Caption         =   "Linha de Comando"
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnJan 
      Caption         =   "&Janelas"
      WindowList      =   -1  'True
      Begin VB.Menu mnBarAtalho 
         Caption         =   "Barra de Atalhos"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnCasc 
         Caption         =   "&Cascata"
      End
      Begin VB.Menu mnLado 
         Caption         =   "&Lado a Lado"
      End
      Begin VB.Menu mnTodasJan 
         Caption         =   "&Todas"
      End
      Begin VB.Menu mnOptCor 
         Caption         =   "Opções de Cor"
         Begin VB.Menu mnMudarCordeFundo 
            Caption         =   "Mudar Cor"
         End
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "A&juda"
      Begin VB.Menu mnAjuda 
         Caption         =   "&Indices da Ajuda"
      End
      Begin VB.Menu mnMostraCont 
         Caption         =   "Mostra Controles"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnCred 
         Caption         =   "&Créditos"
      End
   End
End
Attribute VB_Name = "mdiFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Dim bActive As Boolean
Private Sub mnCli_Click()
    frmEdiClientes.Show 0
End Sub

Public Sub muda_visao(cod As tipo_visao)
    
    
    pcClientes.Enabled = False
    lblClientes.Enabled = False
    lblNovoCli.Enabled = False
    pcNovoCli.Enabled = False
    lblPesqCli.Enabled = False
    pcPesqCli.Enabled = False
    
    pcProdutos.Enabled = False
    lblProdutos.Enabled = False
    lblNovoPro.Enabled = False
    lblPesqPro.Enabled = False
    pcNovoPro.Enabled = False
    pcPesqPro.Enabled = False
    
    
    pcFinanceiro.Enabled = False
    lblMov.Enabled = False
    pcMov.Enabled = False
    lblChPag.Enabled = False
    pcChPag.Enabled = False
    
    lblContasPag.Enabled = False
    pcCpag.Enabled = False
    lblContasRec.Enabled = False
    pcCRec.Enabled = False
    
    lblChRec.Enabled = False
    pcChRec.Enabled = False
    
    
    
    pcEntradas.Enabled = False
    pcVendas.Enabled = False
    pcSaidas.Enabled = False
    pcClientes.Enabled = False
    lblPesqCli.Enabled = False
    pcPesqCli.Enabled = False
    
    
    
    
    If cod = tCAIXA Then
        pcMenu.Visible = False
        Unload frmEscolherVisao
        frmMenu.Show 0
    ElseIf cod = tEstoque Then
        strVisao = "Estoque"
        
    ElseIf cod = tAtacado Then
        strVisao = "Atacado"
        
        
    
    ElseIf cod = tFinanceiro Then
        strVisao = "Financeiro"
    Else
        strVisao = "Geral"
        
            pcClientes.Enabled = True
        lblClientes.Enabled = True
        lblNovoCli.Enabled = True
        pcNovoCli.Enabled = True
        lblPesqCli.Enabled = False
        pcPesqCli.Enabled = True
        
        pcProdutos.Enabled = True
        lblProdutos.Enabled = True
        lblNovoPro.Enabled = True
        lblPesqPro.Enabled = True
        pcNovoPro.Enabled = True
        pcPesqPro.Enabled = True
        
        
        pcFinanceiro.Enabled = True
        lblMov.Enabled = True
        pcMov.Enabled = True
        lblChPag.Enabled = True
        pcChPag.Enabled = True
        
        lblContasPag.Enabled = True
        pcCpag.Enabled = True
        lblContasRec.Enabled = True
        pcCRec.Enabled = True
        
        lblChRec.Enabled = True
        pcChRec.Enabled = True
        
        
        
        pcEntradas.Enabled = True
        pcVendas.Enabled = True
        pcSaidas.Enabled = True
            
        pcClientes.Enabled = True
        lblPesqCli.Enabled = True
        pcPesqCli.Enabled = True
                
    End If
    If cod = tGeral Then
        pcClientes.Enabled = True
    lblClientes.Enabled = True
    lblNovoCli.Enabled = True
    pcNovoCli.Enabled = True
    lblPesqCli.Enabled = True
    pcPesqCli.Enabled = True
    
    pcProdutos.Enabled = True
    lblProdutos.Enabled = True
    lblNovoPro.Enabled = True
    lblPesqPro.Enabled = True
    pcNovoPro.Enabled = True
    pcPesqPro.Enabled = True
    
    
    pcFinanceiro.Enabled = True
    lblMov.Enabled = True
    pcMov.Enabled = True
    lblChPag.Enabled = True
    pcChPag.Enabled = True
    
    lblContasPag.Enabled = True
    pcCpag.Enabled = True
    lblContasRec.Enabled = True
    pcCRec.Enabled = True
    
    lblChRec.Enabled = True
    pcChRec.Enabled = True
    
    
    
    pcEntradas.Enabled = True
    pcVendas.Enabled = True
    pcSaidas.Enabled = True
    pcClientes.Enabled = True
    lblPesqCli.Enabled = True
    pcPesqCli.Enabled = True
    
    
    End If

End Sub



Private Sub cmndbtnCompras_Click()
    frmEdiEntr.Show 0
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmndbtnFechar_Click()

End Sub

Private Sub cmndbtnDocs_Click()
    frmEditaMov.Show 0
End Sub

Private Sub cmndbtnJan_Click()
    frmConsPreco.Show 0
    frmConsPreco.SetFocus
'    Me.PopupMenu mnJan
End Sub


Private Sub cmndbtnSair_Click()
    mnSair_Click
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0, mdiFrm
End Sub

Private Sub mnAjEstRes_Click()
    'P_AjustaEstoqueReservado
End Sub





Private Sub Label8_Click()
    picSair_Click
End Sub

Private Sub lblChPag_Click()
    pcChPag_Click
End Sub

Private Sub lblChRec_Click()
    pcChRec_Click
End Sub

Private Sub lblContasPag_Click()
    pcCpag_Click
End Sub

Private Sub lblContasRec_Click()
    pcCRec_Click
End Sub

Private Sub lblEmpresa_Click()
    buscaEmpresa
    MDIForm_Activate
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    
    e = sped_empresa.SelecionarSped_Empresa("")
    If e.codigo <> "" Then
        InicializaEmpresa e.codigo
    End If
    
End Sub
Private Sub lblNovaVenda_Click()
    pcNovaVenda_Click
End Sub

Private Sub lblNovoCli_Click()
    pcNovoCli_Click
End Sub

Private Sub lblNovoPro_Click()
    pcNovoPro_Click
End Sub

Private Sub lblPesqCli_Click()
    pcPesqCli_Click
End Sub

Private Sub lblPesqEntrada_Click()
    pcPesqEntrada_Click
End Sub

Private Sub lblPesqPro_Click()
    pcPesqPro_Click
End Sub

Private Sub lblPesqSaida_Click()
    pcPesqSaida_Click
End Sub

Private Sub lblPesqVenda_Click()
    pcPesqVenda_Click
End Sub

Private Sub MDIForm_Resize()
    pcHora.Left = pcBot.width - pcHora.width
End Sub

Private Sub mnAbreCx_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
        
    End If
    
    frmAberturaCaixa.Show 1
    
End Sub

Private Sub mnAcBr_Click()
    frmAcbrMonitor.Show 1
End Sub

Private Sub mnBAckUp_Click()
    frmBackup.Show 1
End Sub

Private Sub mnBarAtalho_Click()
    frmMenu.Show 0
    frmMenu.SetFocus
    
End Sub

Private Sub mnBol_Click()
    frmBoletos.Show 0
End Sub

Private Sub mnBuscaOrc_Click()
    frmEdiOrc.Show 0
    frmEdiOrc.SetFocus
End Sub

Private Sub mnCaixas_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbCritical
        Exit Sub
    
    End If
    frmCaixas.Show 0
End Sub

Private Sub mnCalc_Click()
    On Error Resume Next
    Shell "C:\windows\calc.exe", vbNormalFocus
    Shell "C:\windows\system32\calc.exe", vbNormalFocus
End Sub

Private Sub mnCalMed_Click()
    If strViscusto_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para esta operação!", vbExclamation
        Exit Sub
        
    End If
    frmMediaConsumo_Mer.Show 1
End Sub

Private Sub mnCFOP_Click()
    frmEdiCfop.Show 1
End Sub

Private Sub mnChPag_Click()
    frmControleChequesPagar.Show 0
End Sub

Private Sub mnCidadesIBGE_Click()
    frmImportCid.Show 1
End Sub

Private Sub mnCOmport_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelPainelVendasMensais.Show
End Sub

Private Sub mnCompProCli_Click()
    frmComportProCli.Show 0
    frmComportProCli.SetFocus
End Sub

Private Sub mnCons_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmFechaFinan.Show 1
    End If
End Sub

Private Sub mnConsPre_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmConsPreco.Show 0
    frmConsPreco.SetFocus
End Sub

Private Sub mnCont_Click()
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    frmContagem.Show 0
    frmContagem.SetFocus
End Sub


Private Sub cmndbtnCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub cmndbtnDev_Click()
    'frmDevolucoes.Show 0
End Sub

Private Sub cmndbtnNewVEnda_Click()
    On Error Resume Next
    frmVenda.Show 1
End Sub

Private Sub cmndbtnPEd_Click()
    frmEdiPed.Show 0
End Sub

Private Sub cmndbtnProd_Click()
    frmEdiPro.Show 0, mdiFrm
    frmEdiPro.SetFocus
End Sub

Private Sub cmndbtnVendas_Click()
    frmMonitoramentoCaixas.Show 0, mdiFrm
    frmMonitoramentoCaixas.SetFocus
End Sub
Sub showItens(opcao As Integer)
    mnRelFatur.Visible = opcao
    mnRelEst.Visible = opcao
    mnRelFinan.Visible = opcao
    mnMonit.Visible = opcao
    pcPesqCli.Enabled = opcao
    mnInvent.Visible = opcao
    
    
    
End Sub

Private Sub mnDupRec_Click()
    frmEditaMov.Show 0, mdiFrm
End Sub

Private Sub mnEdiDev_Click()
    frmEdiDEv.Show 0, mdiFrm
End Sub

Private Sub mnEmpresaSped_Click()
    If strAbrenota_usu = "S" Then
    
    
        frmEdiSpedEmpresa.Show 1
    Else
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
    End If
End Sub

Private Sub mnEntregas_Click()
    frmEntregas.Show 0, mdiFrm
End Sub

Private Sub mnFechaCaixa_Click()
    If strFechacai_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    frmFechaCaixa.Show 1
End Sub

Private Sub mnHistProd_Click()
    frmHistProd.Show 1
End Sub

Private Sub mnInfAdd_Click()
    frmNfeInfAddPadrao.Show 0, mdiFrm
End Sub

Private Sub mnLan_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmLanCaixa
    fr.Show 1
End Sub

Private Sub mnLanNotEnergiaEletrica_Click()
    frmSpedEnergiaEletrica.Show 0, mdiFrm
End Sub

Private Sub mnLanRedZ_Click()
    frmEdiRedZ.Show 0, mdiFrm
End Sub

Private Sub mnListaIcms_Click()
    Sped_Lista_Cst_Icms ("")
End Sub

Private Sub mnMudarCordeFundo_Click()
       
    Dim cor
    mdiFrm.cmndlg.ShowColor
    cor = mdiFrm.cmndlg.Color
    
    
    frmMenu.mudaCor cor
    If MsgBox("Deseja utilizar essa cor na próxima vez que acessar o sistema?", vbYesNo + vbQuestion) = vbYes Then
        If Not F_senha Then Exit Sub
        funcionario.AlteraCorFundoUsu strCod_Usu, cor
    End If
    
    
End Sub

Private Sub mnNfe_Click()
    frmEdiNFe.Show 0
End Sub

Private Sub mnParticipantes_Click()
    frmEdiParticipante.Show 0, mdiFrm
End Sub

Private Sub mnRedZ_Click()
    frmEdiRedZ.Show 0, mdiFrm
    frmEdiRedZ.SetFocus
End Sub

Private Sub mnRelAcompMensalCliProd_Click()
    frmRelAcompanhamentoCliMensalProd.Show 1
End Sub

Private Sub mnRelAltPreco_Click()
    frmRelAltPreco.Show 1
End Sub

Private Sub mnRelBaixas_Click()
    frmRelCarnesRecebidos.Show 0, mdiFrm
End Sub

Private Sub mnRelComissVen_Click()
    If strVisrelfin_usu = "S" Then
        frmRelVendasVendedor.Show 1
    End If
End Sub

Private Sub mnRelComparaCli_Click()
    frmRelAcompanhamentoProduto.Show 0, mdiFrm
End Sub

Private Sub mnRelComport_Click()
    frmRelComportamentoProd.Show 0
End Sub

Private Sub mnRelFreqVen_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    frmRelVendasFrequencia.Show 0, mdiFrm
End Sub

Private Sub mnRelProdsNaoVendidos_Click()
    frmRelProdutosNaoVendidosPorMarca.Show 0, mdiFrm
End Sub

Private Sub mnReposicao_Click()
    frmEdiReposicao.Show 0, mdiFrm
End Sub

Private Sub mnSecao_Click()
    frmEdiSecao.Show 0, mdiFrm
    frmEdiSecao.SetFocus
End Sub

Private Sub mnSendMail_Click()
    If strNivel_Usu < 3 Then
        MsgBox "Seu usuário não tem nível de permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    frmEdiMails.Show 1
End Sub

Private Sub mnServ_Click()
    frmEdiServico.Show 1
End Sub

Private Sub mnShell_Click()
    frmShell.Show 1
End Sub

Private Sub mnSimilares_Click()
    frmCadSim.Show 0, mdiFrm
End Sub

Private Sub mnSubcat_Click()
    frmEdiSubCat.Show 0
End Sub

Private Sub mnSubSec_Click()
    frmEdiSubSec.Show 0, mdiFrm
    frmEdiSubSec.SetFocus
End Sub

Private Sub mnTipoServ_Click()
    frmEdiTipoServ.Show 0, mdiFrm
End Sub

Private Sub mnUnidades_Click()
    EditaUnidade
End Sub

Private Sub mnVeicFab_Click()
    frmEdiVeiculoFabricante.Show 1
End Sub

Private Sub mnVeiculos_Click()
    frmEdiVeiculo.Show 1
End Sub

Private Sub mNVendasAntigas_Click()
    frmEdiVendasAntigas.Show 0, mdiFrm
End Sub

Private Sub pcMax_Click()
    pcMax.Visible = False
    pcMenu.Visible = True
End Sub

Private Sub pcMin_Click()
    pcMenu.Visible = False
    pcMax.Visible = True
End Sub

Public Sub MudaCorFundo(cor)
    Me.BackColor = cor
End Sub
Private Sub MDIForm_Activate()
     On Error Resume Next
    Dim versao As String
    
     MudaCorFundo strCorFundoUsu
    
    versao = "Piki"
    lblBD.caption = sServer
    
    VerificaAtualizacao
    MDIForm_Resize
    Me.caption = versao & " : " & strNome_Emp & " - " & strCidade_Emp & "/" & strUF_Emp & " .:. CNPJ.: " & strCNPJ_Emp
    'TaskBar False
   ' If F_BuscaUmCampo("cod_loj", tab_parametros, "") = "" Then
    '    MsgBox "Os Dados desta Loja não estão definidos corretamente!" & Chr(13) & _
     '       "Informe esta mensagem ao Responsável!", vbExclamation
      '  P_Fim
        'End If
   lblConexao.caption = Winsock1.LocalIP & "-" & Winsock1.LocalHostName & " BD: " & conn.DefaultDatabase & " em " & sServer
   strIP_Usu = Winsock1.LocalIP
   strHost_Usu = Winsock1.LocalHostName
   strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixasip", "ip_cai='" & strIP_Usu & "'")
   
   
   If strCod_Cai = "" Then
        strCod_Cai = "00"
   End If
    strSerie_Nfe = F_BuscaUmCampo("serie_nfe", "tab_caixas", "cod_Cai='" & strCod_Cai & "'")
   If strSerie_Nfe = "" Then
    strSerie_Nfe = "0"
   End If
   
   
    strCod_Tur = F_BuscaUmCampo("Cod_tur", TAB_CAIXAS, "cod_Cai='" & strCod_Cai & "'")
    If strCod_Tur = "" Then
        strCod_Tur = "1"
    End If
    If bActive <> True Then
'        frmMenu.Show 0, Me
'        frmMenu.Left = 0
'        frmMenu.Top = 600
'        frmMenu.Height = StatusBar1.Top - 50
        bActive = True
        
    End If
    
    
    
    lblDia.caption = StrConv(WeekdayName(Weekday(Date)), vbProperCase)
    lblNumDia.caption = Day(Date)
    lblMes.caption = StrConv(MonthName(Month(Date)), vbProperCase)
    lblAno.caption = Year(Date)
    lblEmpresa.caption = strCod_Emp & "-" & strNome_Emp
    frmMenu.mudaVariaveis
    
    
    
    If Not bUsaVeiculo Then
        mnVeicFab.Visible = False
        mnVeiculos.Visible = False
    End If
    
    If Not bUsaServico Then
        mnServ.Visible = False
        mnTipoServ.Visible = False
    End If
    
    If Not bUsaEntrega Then
        mnEntregas.Visible = False
        
    End If
End Sub

Private Sub MDIForm_DblClick()
    'frmVenda.Show 0, Me
End Sub

Private Sub MDIForm_Load()
    'showItens False 'esconder controles
    lblUsuario.caption = strNome_Usu
    'frmLogo.Show 0
    frmMenu.Show 0
    frmMenu.SetFocus
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnPrinc
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'TaskBar True
    P_Fim
End Sub

Private Sub mnAjuda_Click()
    MsgBox "em breve"
End Sub

Private Sub mnApar_Click()
    frmConfig.Show 0
End Sub

Private Sub mnBan_Click()
    frmEdiBan.Show 0
End Sub

Private Sub mnCadDiv_Click()
    frmEdiDiv.Show 0
End Sub

Private Sub mnCadEsp_Click()
    frmEdiMar.Show 0
End Sub

Private Sub mnCadGen_Click()
    frmEdiCat.Show 0
End Sub

Private Sub mnCadCargos_Click()
    frmEdiCargos.Show 0
End Sub

Private Sub mnCadCat_Click()
    frmEdiCat.Show 1
End Sub

Private Sub mnCadMar_Click()
    frmEdiMar.Show 1
End Sub

Private Sub mnCadPla_Click()
    frmEdiPlanoContas.Show 0, mdiFrm
    frmEdiPlanoContas.SetFocus
End Sub

Private Sub mnCalGuias_Click()
    frmLanDoc.Show 1
End Sub

Private Sub mnCasc_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnCheq_Click()
    frmControleChequesReceber.Show 1
End Sub

Private Sub mnComDes_Click()
    frmEdiDes.Show 1
End Sub

Private Sub mnCOn_Click()
    frmEdiCon.Show 0
End Sub

Private Sub mnConfig_Click()
    frmParametros.Show 0
End Sub

Private Sub mnContasChequesPagar_Click()
    If strVisrelfin_usu = "S" Then
    
        frmContasPagar.Show 0
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub mnContasPagar_Click()
    frmEditaMov.Show 0, mdiFrm
End Sub

Private Sub mnCred_Click()
    On Error Resume Next
    frmCred.Show 1
End Sub

Private Sub mnCriarReplica_Click()
    frmCriarReplica.Show 1
End Sub

Private Sub mnDemoInv_Click()
    frmDemonstrativoInvent.Show 0
    
End Sub

Private Sub mnDePara_Click()
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    frmDePara.Show 1
End Sub

Private Sub mnDEv_Click()
    Dim strSQL As String
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    
    
    
    
        
    
  
frmDevolucao.Show 1
    
End Sub

Private Sub mnEcf_Click()
    frmECF.Show 0, mdiFrm
End Sub

Private Sub mnEDiCid_Click()
    frmEdiCid.Show 0
End Sub

Private Sub mnEdiComp_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEntradasSped.Show 0, mdiFrm
    
End Sub


Private Sub mnEdiMovFin_Click()
    
    frmEdiMovFin.Show 0, mdiFrm
End Sub

Private Sub mnEdiMovInt_Click()
    frmEdiMovInt.Show 0
    frmEdiMovInt.SetFocus
End Sub

Private Sub mnEdiNotaSai_Click()
    frmEdiNf.Show 0
    frmEdiNf.SetFocus
End Sub

Private Sub mnEdiPedCom_Click()
    frmEdiPedidoCompra.Show 0
End Sub

Private Sub mnEdiPer_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiPerdas.Show 0
    frmEdiPerdas.SetFocus
End Sub

Private Sub mnEdiProcomp_Click()
    frmEdiProcomp.Show 0
    frmEdiProcomp.SetFocus
End Sub

Private Sub mnEditCli_Click()
    frmEdiClientes.Show 0, mdiFrm
End Sub

Private Sub mnEditLan_Click()
     
    frmEditaMov.Show 0, mdiFrm
End Sub

Private Sub mnEdiTransf_Click()
    frmEdiTransf.Show 0
    frmEdiTransf.SetFocus
End Sub

Private Sub mnEdiTro_Click()
    frmEdiTrocas.Show 0
    frmEdiTrocas.SetFocus
End Sub

Private Sub mnEdiUsu_Click()
    frmEdiUsu.Show 0
End Sub

Private Sub mnEdiVen_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiVendas.Show 0, mdiFrm
End Sub

Private Sub mnEfetivVen_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    If IsLoad("frmEfetivaVen") Then
        frmEfetivaVen.SetFocus
    Else
        frmEfetivaVen.Show 0
    End If
End Sub

Private Sub mnEdiVenAtac_Click()
    frmEdiVendas.Show 0, mdiFrm
End Sub

Private Sub mnEfetVen_Click()
    frmEfetivaVen.Show 0
End Sub

Private Sub mnEntrada_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmNotaEntrada.Show 1
End Sub

Private Sub mnEntrPed_Click()
    frmEfetivaPed.Show 0
End Sub

Private Sub mnEntrVen_Click()
    frmEntrega.Show 0
End Sub

Private Sub mnEstIni_Click()
    On Error Resume Next
    frmAjEstIni.Show 1
End Sub

Private Sub mnEti_Click()
    P_ImprxEstiquetas frmVisualiza.Pic
    frmVisualiza.Show 1
End Sub

Private Sub mnEspelho_Click()
    frmEspelhoNota.Show 0
    frmEspelhoNota.SetFocus
End Sub

Private Sub mnFecharMes_Click()
    If strFechames_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    frmFecharMes.Show 1
End Sub

Private Sub mnFinanceiras_Click()
    frmEdiFinan.Show 0
End Sub

Private Sub mnFormasPag_Click()
    frmEdiFormasPag.Show 1
End Sub

Private Sub mnForn_Click()
    frmEdiFor.Show 0
End Sub

Private Sub mnIncCli_Click()
    frmIncClientes.Show 0
End Sub

Private Sub mnIncTer_Click()
    frmIncTer.Show 0
End Sub



Private Sub mnLivIcms_Click()
    frmLivIcms.Show 0
End Sub

Private Sub mnLivIpi_Click()
    frmLivIpi.Show 0
End Sub

Private Sub mnMudaCli_Click()
    On Error Resume Next
    frmSeleccliente.Show 1
End Sub

Private Sub mnPer_Click()
    frmPerfis.Show 0
End Sub

Private Sub mnFreteiros_Click()
    frmEdiFreteiros.Show 0
    frmEdiFreteiros.SetFocus
End Sub

Private Sub mnFun_Click()
    frmEdiFun.Show 0
End Sub

Private Sub mnGeraEti_Click()
    frmGeraEtiquetas.Show 0
    frmGeraEtiquetas.SetFocus
End Sub

Private Sub mnGeraNota_Click()
    frmGeraNota.Show 1
End Sub

Private Sub mnHistCheCPf_Click()
    frmHistoricoChe.Show 0, mdiFrm
End Sub

Private Sub mnImo_Click()
    frmEdiImo.Show 0
End Sub

Private Sub mnImp_Click()
    frmAlteraImpostos.Show 1
End Sub

Private Sub mnImpInvent_Click()
    frmImpInvent.Show 0, mdiFrm
End Sub

Private Sub mnIncMer_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmEntradaSped
    fr.Show 0, mdiFrm
End Sub

Private Sub mnInvent_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmInventario.Show 0
    frmInventario.SetFocus
End Sub

Private Sub mnLado_Click()
    Me.Arrange vbVertical
End Sub

Private Sub mnLanCaixa_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiCai.Show 0
End Sub

Private Sub mnLojas_Click()
    frmEdiLoj.Show 0
End Sub

Private Sub mnLanCont_Click()
    frmLanCont.Show 1
End Sub

Private Sub mnLoj_Click()
    frmEdiLoj.Show 0
    frmEdiLoj.SetFocus
End Sub

Private Sub mnMescComp_Click()
    frmMesclaComp.Show 0, mdiFrm
End Sub

Private Sub mnMonit_Click()
    frmMonitoramentoCaixas.Show 0, mdiFrm
    frmMonitoramentoCaixas.SetFocus
End Sub

Private Sub mnMostraCont_Click()
    showItens True
End Sub

Private Sub mnMovFin_Click()
    frmEdiFin.Show 0
    frmEdiFin.SetFocus
End Sub

Private Sub mnMovUsu_Click()
    frmFechaCaixaUsu.Show 0, mdiFrm
End Sub

Private Sub mnMudSen_Click()
    On Error Resume Next
    frmMudaSenha.Show 1
End Sub

Private Sub mnNovaVenda_Click()
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim strSQL As String
    
    strSQL = F_GeraNovaVenda("00000", strFormaPagamentoPadrao)
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.tag = strSQL
    On Error Resume Next
    frmVenda.Atacado = False
    frmVenda.Show 0, mdiFrm
    
End Sub

Private Sub mnPed_Click()
    frmPedidoCompra.Show 0
End Sub

Private Sub mnPisos_Click()
    frmEstoqueIni.Show 1
End Sub

Private Sub mnNovoMovInt_Click()
    Dim fr As New frmMovInt
    fr.Show 0
End Sub

Private Sub mnNovoOrc_Click()
    frmOrcamentoUniformes.Show 0
    frmOrcamentoUniformes.SetFocus
End Sub

Private Sub mnProd_Click()
    frmEdiPro.Show 0, mdiFrm
End Sub

Private Sub mnRelGuias_Click()
    frmRelGuias.Show 0
End Sub

Private Sub mnRecCli_Click()
    frmRecebimentos.Show 0
    frmRecebimentos.SetFocus
End Sub

Private Sub mnRelAcompMensal_Click()
    frmRelAcompanhamentoCliMensal.Show 0
End Sub

Private Sub mnRelAltPre_Click()
    frmRelAltPreco.Show 0, mdiFrm
End Sub

Private Sub mnRelComp_Click()
If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
End Sub

Private Sub mnProdProm_Click()
    
        frmProdProm.Show 0
        frmProdProm.SetFocus
    
    
End Sub

Private Sub mnProm_Click()
    
        
        frmProdProm.Show 0
    
End Sub

Private Sub mnRelAnaFat_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFatMensal.Show 0
        frmRelFatMensal.SetFocus
    End If
End Sub

Private Sub mnRelAnaIndFluxo_Click()
    If strVisrelest_usu = "S" Then
        frmRelAnaliseIndFluxo.Show 0
        frmRelAnaliseIndFluxo.SetFocus
    Else
        MsgBox "Acesso Inválido!", vbCritical
    End If
End Sub

Private Sub mnRelAtvCli_Click()
    If strVisrelfin_usu = "S" Then
        frmRelAtvCli.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mNRelAtvPro_Click()
    If strVisrelfin_usu = "S" Then
        frmRelAtvCliProd.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnRelChq_Click()
    frmRelCheques.Show 0
    frmRelCheques.SetFocus
End Sub

Private Sub mnRelCmpFor_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelProdFor.Show 0
    frmRelProdFor.SetFocus
End Sub

Private Sub mnRelCOmiss_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Você não tem permissão para este aesso!", vbExclamation
    Else
        frmRelComiss.Show 0
        frmRelComiss.SetFocus
    End If
End Sub

Private Sub mnRelCompEst_Click()
    If strVisrelest_usu = "S" Then
        frmRelEstFis.Show
    End If
End Sub

Private Sub mnRelDev_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelDevProd.Show 0
End Sub

Private Sub mnRelFatVenda_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFat.Show 0
        frmRelFat.SetFocus
    End If
End Sub

Private Sub mnRelFluxo_Click()
    If strVisrelest_usu = "S" Then
        frmRelAnaliseFluxo.Show 0
        frmRelAnaliseFluxo.SetFocus
    Else
        MsgBox "Acesso Inválido!", vbCritical
    End If
End Sub

Private Sub mnRelGiro_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelGiro.Show 0, mdiFrm
    frmRelGiro.SetFocus
End Sub

Private Sub mnRelLucrComp_Click()
    If strVisrelest_usu = "S" Then
        frmRelLucrComp.Show 0
        frmRelLucrComp.SetFocus
    End If
End Sub

Private Sub mnRelMMD_Click()
    frmRelMMD.Show 0
    frmRelMMD.SetFocus
End Sub

Private Sub mnRelPagReceb_Click()
    frmRelMov.Show 0
    frmRelMov.SetFocus
End Sub

Private Sub mnRelProdForn_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelFornProd.Show 0
    frmRelFornProd.SetFocus
End Sub

Private Sub mnRelProdLucr_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbCritical
        Exit Sub
    End If
    frmRelProdLucr.Show 0
End Sub

Private Sub mnRelReceitasDesp_Click()
    frmRelReceitasDespesas.Show 0
End Sub

Private Sub mnRelVend_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelSaiCaixa.Show 0, mdiFrm
    frmRelSaiCaixa.SetFocus
End Sub

Private Sub mnTer_Click()
    If strCOD_CLI = "" Then
        If MsgBox("Não há cliente selecionado." & Chr(13) & "Deseja selecionar um agora?" _
            , vbYesNo + vbQuestion) = vbYes Then
            On Error Resume Next
            frmSeleccliente.Show 1
            If strCOD_CLI = "" Then
                MsgBox "Não foi selecionado nehum cliente", vbExclamation
            Else
                frmSelectTer.Show
            End If
            
        Else
            MsgBox "Não posso manipular terceiros sem que você escolha um cliente", vbCritical
            Exit Sub
        End If
    Else
        frmSelectTer.Show 0
    End If
    
End Sub

Private Sub mnSaiDia_Click()
    On Error Resume Next
    frmSaidaDiaria.Show 1
End Sub

Private Sub mnRelVendaBandeira_Click()
    frmRelVendasBandeiras.Show 0
End Sub

Private Sub mnResumo_Click()
    frmResumoContasPagar.Show 0
End Sub

Private Sub mnSair_Click()
    If MsgBox("Deseja Realmente Sair?", vbYesNo + vbQuestion) = vbYes Then
        P_Fim
    Else
        Exit Sub
    End If
End Sub

Private Sub mnServicos_Click()
    frmEdiServ.Show 0
End Sub

Private Sub mnSujComp_Click()
    If strVisrelest_usu = "S" Then
        frmProcessoDeCompra.Show 0
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
    
End Sub

Private Sub mnTodasJan_Click()
    'On Error Resume Next
   
   If Forms.Count > 1 Then
        frmForms.Show , mdiFrm
    End If
   
End Sub

Private Sub mnTransf_Click()
    On Error Resume Next
    frmTransf.Show 1
End Sub

Private Sub mnTroca_Click()
    If strAbresaida_usu = "S" Then
        On Error Resume Next
        frmTroca.Show 1
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnTransfComp_Click()
    frmTransfComp.Show 0, mdiFrm
    
End Sub

Private Sub mnUsu_Click()
    Dim fr As New frmIncUsu
    fr.Show 0
End Sub

Private Sub mnUtilECF_Click()
    P_PortaComunicacaoECF
    frmUtilitarioECF.Show 1
End Sub

Public Sub mnVendaAtac_Click()
    NovaVenda
End Sub

Private Sub mnVinc_Click()
    frmVinculaDoc.Show 1
End Sub

Private Sub pcChPag_Click()
    frmControleChequesPagar.Show 0
End Sub

Private Sub pcChRec_Click()
    frmControleChequesReceber.Show 0
End Sub

Private Sub pcCpag_Click()
    frmContasPagar.Show 0
End Sub

Private Sub pcCRec_Click()
    frmEditaMov.Show 0
End Sub

Private Sub pcMov_Click()
    frmEdiFin.Show 0
End Sub

Private Sub pcNovaEntrada_Click()
    frmNotaEntrada.Show 1
End Sub

Private Sub pcNovaSaida_Click()
    frmGeraNota.Show 1
End Sub

Private Sub pcNovaVenda_Click()
    mnVendaAtac_Click
End Sub

Private Sub pcNovoCli_Click()
    frmIncClientes.Show 0
End Sub

Private Sub pcNovoPro_Click()
    frmCadPro.Show 1
End Sub

Private Sub pcPesqCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub pcPesqEntrada_Click()
    frmEdiEntr.setOperacao "F"
    frmEdiEntr.Show 0, mdiFrm
End Sub

Private Sub pcPesqPro_Click()
    frmEdiPro.Show 0, mdiFrm
End Sub







Private Sub pcPesqSaida_Click()
    frmEdiNf.Show 0
End Sub

Private Sub pcPesqVenda_Click()
    mnEdiVenAtac_Click
End Sub

Private Sub pcSair_Click()
    picSair_Click
End Sub

Private Sub pcUsu_Click()
    frmOptUsu.Show 0
End Sub

Private Sub picSair_Click()
    mnSair_Click
End Sub

Private Sub Picture6_Click()
    If UCase(strNome_Usu) = "ALISON" Then
        frmMigrarEstoque.Show 1
    End If
End Sub
