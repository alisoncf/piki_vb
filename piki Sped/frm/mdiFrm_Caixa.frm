VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiFrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Piki 2013"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   Icon            =   "mdiFrm_Caixa.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcMenu 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7905
      ScaleWidth      =   2265
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.PictureBox pcSair 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   705
         ScaleWidth      =   1905
         TabIndex        =   62
         Top             =   8880
         Width           =   1935
         Begin VB.PictureBox picSair 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":3132
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   65
            Top             =   398
            Width           =   195
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   63
            Top             =   0
            Width           =   1935
            Begin VB.Label Label7 
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
               Height          =   195
               Left            =   240
               TabIndex        =   64
               Top             =   0
               Width           =   840
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sair"
            Height          =   195
            Left            =   480
            TabIndex        =   66
            Top             =   398
            Width           =   270
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
         TabIndex        =   48
         Top             =   6600
         Width           =   1935
         Begin VB.PictureBox pcCRec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   98
            Picture         =   "mdiFrm_Caixa.frx":329D
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   55
            Top             =   1080
            Width           =   285
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
            TabIndex        =   53
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
               TabIndex        =   54
               Top             =   0
               Width           =   900
            End
         End
         Begin VB.PictureBox pcChRec 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   113
            Picture         =   "mdiFrm_Caixa.frx":33F9
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   52
            Top             =   1800
            Width           =   255
         End
         Begin VB.PictureBox pcChPag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   113
            Picture         =   "mdiFrm_Caixa.frx":3707
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   51
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox pcCpag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":3A15
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   50
            Top             =   720
            Width           =   240
         End
         Begin VB.PictureBox pcMov 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":3C99
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   49
            Top             =   360
            Width           =   240
         End
         Begin VB.Label lblChRec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Receb."
            Height          =   195
            Left            =   480
            TabIndex        =   61
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblChPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques a Pagar"
            Height          =   195
            Left            =   480
            TabIndex        =   60
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber"
            Height          =   195
            Left            =   360
            TabIndex        =   59
            Top             =   0
            Width           =   1290
         End
         Begin VB.Label lblContasRec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Receber"
            Height          =   195
            Left            =   480
            TabIndex        =   58
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label lblContasPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contas a Pagar"
            Height          =   195
            Left            =   480
            TabIndex        =   57
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblMov 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimentos"
            Height          =   195
            Left            =   480
            TabIndex        =   56
            Top             =   383
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   41
         Top             =   1320
         Width           =   1935
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   44
            Top             =   0
            Width           =   1935
            Begin VB.Label Label5 
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
               TabIndex        =   45
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.PictureBox pcNovoPro 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":40B0
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   43
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox pcPesqPro 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":4214
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   42
            Top             =   720
            Width           =   345
         End
         Begin VB.Label lblNovoPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   47
            Top             =   405
            Width           =   390
         End
         Begin VB.Label lblPesqPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   46
            Top             =   765
            Width           =   690
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   34
         Top             =   2640
         Width           =   1935
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   37
            Top             =   0
            Width           =   1935
            Begin VB.Label Label4 
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
               TabIndex        =   38
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.PictureBox pcNovaEntrada 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":469E
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   36
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox pcPesqEntrada 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   82
            Picture         =   "mdiFrm_Caixa.frx":4802
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   35
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblNovaEntrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   40
            Top             =   405
            Width           =   390
         End
         Begin VB.Label lblPesqEntrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   39
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
         Picture         =   "mdiFrm_Caixa.frx":4BE0
         ScaleHeight     =   2265
         ScaleWidth      =   135
         TabIndex        =   33
         Top             =   0
         Width           =   165
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000000&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   26
         Top             =   0
         Width           =   1935
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   29
            Top             =   0
            Width           =   1935
            Begin VB.Label Label1 
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
               TabIndex        =   30
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.PictureBox pcPesqCli 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   0
            Picture         =   "mdiFrm_Caixa.frx":5303
            ScaleHeight     =   450
            ScaleWidth      =   435
            TabIndex        =   28
            Top             =   720
            Width           =   435
         End
         Begin VB.PictureBox pcNovoCli 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":586D
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   27
            Top             =   360
            Width           =   285
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
         Begin VB.Label lblNovoCli 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   31
            Top             =   405
            Width           =   390
         End
      End
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   19
         Top             =   3960
         Width           =   1935
         Begin VB.PictureBox pcPesqSaida 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":59D1
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   23
            Top             =   720
            Width           =   240
         End
         Begin VB.PictureBox pcNovaSaida 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":63D3
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   22
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox Picture24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   20
            Top             =   0
            Width           =   1935
            Begin VB.Label Label17 
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
               TabIndex        =   21
               Top             =   0
               Width           =   585
            End
         End
         Begin VB.Label lblPesqSaida 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   25
            Top             =   765
            Width           =   690
         End
         Begin VB.Label lblNovaSaida 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   24
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   1905
         TabIndex        =   11
         Top             =   5280
         Width           =   1935
         Begin VB.PictureBox Picture3 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1680
            Picture         =   "mdiFrm_Caixa.frx":6537
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   16
            Top             =   0
            Width           =   240
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
            TabIndex        =   14
            Top             =   0
            Width           =   1935
            Begin VB.Label Label10 
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
               TabIndex        =   15
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.PictureBox pcNovaVenda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":68C1
            ScaleHeight     =   270
            ScaleWidth      =   285
            TabIndex        =   13
            Top             =   360
            Width           =   285
         End
         Begin VB.PictureBox pcPesqVenda 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "mdiFrm_Caixa.frx":6A25
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   12
            Top             =   720
            Width           =   240
         End
         Begin VB.Label lblNovaVenda 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblPesqVenda 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar"
            Height          =   195
            Left            =   480
            TabIndex        =   17
            Top             =   765
            Width           =   690
         End
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11610
      TabIndex        =   0
      Top             =   7935
      Width           =   11610
      Begin VB.PictureBox pcUsu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "mdiFrm_Caixa.frx":7427
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox pcHora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11160
         ScaleHeight     =   375
         ScaleWidth      =   4695
         TabIndex        =   1
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
            Picture         =   "mdiFrm_Caixa.frx":7CF1
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   2
            Top             =   45
            Width           =   240
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
            TabIndex        =   6
            Top             =   30
            Width           =   1410
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
            TabIndex        =   5
            Top             =   -15
            Width           =   330
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
            TabIndex        =   4
            Top             =   30
            Width           =   975
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
            TabIndex        =   3
            Top             =   0
            Width           =   675
         End
      End
      Begin VB.Label lblMsg 
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
         Left            =   6720
         TabIndex        =   9
         Top             =   120
         Width           =   135
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   120
         Width           =   540
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8029
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8087
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":80E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8143
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":81A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":81FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":825D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":82BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8319
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8377
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":83D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8433
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8491
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":84EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":854D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":85AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8609
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8667
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":86C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8723
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":8781
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListPrint 
      Left            =   7320
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   419
      ImageHeight     =   192
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiFrm_Caixa.frx":87DF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnPrinc 
      Caption         =   "&Principal"
      Visible         =   0   'False
      Begin VB.Menu mnEditCli 
         Caption         =   "&Clientes"
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
      Begin VB.Menu mnProd 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mnCadCat 
         Caption         =   "C&ategorias"
      End
      Begin VB.Menu mnCadMar 
         Caption         =   "&Marcas"
      End
      Begin VB.Menu div4 
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
   Begin VB.Menu mnCadTer 
      Caption         =   "&Compras"
      Visible         =   0   'False
      Begin VB.Menu mnIncMer 
         Caption         =   "&Nova Compra"
      End
      Begin VB.Menu mnEdiComp 
         Caption         =   "&Edição de Compras"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnEntrada 
         Caption         =   "E&ntrada de Produtos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnEdiPer 
         Caption         =   "&Perdas"
      End
      Begin VB.Menu mnSujComp 
         Caption         =   "&Sujestão de Compras"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnMov 
      Caption         =   "&Vendas"
      Visible         =   0   'False
      Begin VB.Menu mnNovaVenda 
         Caption         =   "&Nova Venda"
      End
      Begin VB.Menu mnEdiVen 
         Caption         =   "Edição de Vendas"
         Shortcut        =   {F5}
      End
      Begin VB.Menu div10 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnEdiDev 
         Caption         =   "&Devoluções"
         Visible         =   0   'False
      End
      Begin VB.Menu mnDEv 
         Caption         =   "Nova Dev&olução"
      End
      Begin VB.Menu mnDivVen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnProdProm 
         Caption         =   "&Produtos em Promoção"
      End
      Begin VB.Menu mnDivVen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnTroca 
         Caption         =   "&Troca de Produtos"
      End
      Begin VB.Menu mnEdiTro 
         Caption         =   "&Edição de Trocas"
      End
   End
   Begin VB.Menu mnCaixa 
      Caption         =   "Cai&xa"
      Visible         =   0   'False
      Begin VB.Menu mnAberturaDia 
         Caption         =   "&Abertura do Dia"
      End
      Begin VB.Menu mnLanCaixa 
         Caption         =   "&Edição de Lançamentos"
      End
      Begin VB.Menu mnFechaCaixa 
         Caption         =   "&Lançamentos"
      End
      Begin VB.Menu mnCOntEcf 
         Caption         =   "Controle ECF"
      End
      Begin VB.Menu mnFechamento 
         Caption         =   "&Fechamento do Dia"
      End
   End
   Begin VB.Menu mnRel 
      Caption         =   "&Relatórios"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnRelVend 
         Caption         =   "Relatório de &Vendas"
      End
      Begin VB.Menu mnRelDev 
         Caption         =   "&Relatório de Devoluções"
      End
      Begin VB.Menu mnRelProdForn 
         Caption         =   "&Relatório de Produtos por Fornecedor"
      End
      Begin VB.Menu mnRelCmpFor 
         Caption         =   "Relatório de &Compras por Fornecedor"
      End
      Begin VB.Menu mnInvent 
         Caption         =   "&Inventário de Conf."
      End
      Begin VB.Menu div12 
         Caption         =   "-"
      End
      Begin VB.Menu mnRelCOmiss 
         Caption         =   "Relatório de &Comissões"
      End
      Begin VB.Menu mnRelFatVenda 
         Caption         =   "Relatório de &Faturamento por Vendas"
      End
      Begin VB.Menu mnRelAnaFat 
         Caption         =   "&Análise de Faturamento"
      End
      Begin VB.Menu mnRelGiro 
         Caption         =   "Relatório de Giro Periódico"
      End
      Begin VB.Menu mnCompProCli 
         Caption         =   "Relatório de Comp. de Prod. p/ C&liente"
      End
      Begin VB.Menu mnCOmport 
         Caption         =   "Relatório de &Comportamento de Produtos"
      End
      Begin VB.Menu div13 
         Caption         =   "-"
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
   End
   Begin VB.Menu mnParam 
      Caption         =   "Parâ&metros"
      Visible         =   0   'False
      Begin VB.Menu mnCont 
         Caption         =   "C&ontagens"
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
         Enabled         =   0   'False
      End
      Begin VB.Menu mnMudSen 
         Caption         =   "&Mudar Senha"
      End
      Begin VB.Menu div5 
         Caption         =   "-"
      End
      Begin VB.Menu mnFormasPag 
         Caption         =   "&Formas de Pagamento"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnApar 
         Caption         =   "&Aparência"
      End
   End
   Begin VB.Menu mnJan 
      Caption         =   "&Janelas"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnCasc 
         Caption         =   "&Cascata"
      End
      Begin VB.Menu mnLado 
         Caption         =   "&Lado a Lado"
      End
      Begin VB.Menu mnTodasJan 
         Caption         =   "&Todas"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "A&juda"
      Visible         =   0   'False
      Begin VB.Menu mnAjuda 
         Caption         =   "&Indices da Ajuda"
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
Dim bActive As Boolean
Private Sub mnCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub cmndbtnAbertura_Click()
    mnAberturaDia_Click
End Sub

Private Sub cmndbtnCaixa_Click()
    frmEdiCai.Show 0
    frmEdiCai.SetFocus
End Sub

Private Sub cmndbtnECF_Click()
    On Error Resume Next
    frmImpressoraECF.Show 1
End Sub
Private Sub cmndbtnJan_Click()
    Me.PopupMenu mnJan
End Sub

Private Sub cmndbtnFechamento_Click()
    mnFechaCaixa_Click
End Sub

Private Sub cmndbtnNovaVenda_Click()
    mnNovaVenda_Click
End Sub

Private Sub cmndbtnSair_Click()
    mnSair_Click
End Sub

Private Sub cmndbtnUsu_Click()
    frmOptUsu.Show 0, mdiFrm
End Sub

Private Sub mnAjEstRes_Click()
    P_AjustaEstoqueReservado
End Sub

Private Sub mnArr_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiArr.Show 0
    frmEdiArr.SetFocus
End Sub

Private Sub mnBol_Click()
    frmBoletos.Show 0
End Sub

Private Sub mnAberturaDia_Click()
    frmAberturaCaixa.Show 1
End Sub

Private Sub mnCalMed_Click()
    If strViscusto_usu <> "S" Then
        MsgBox "VocÊ não tem permissão para esta operação!", vbExclamation
        Exit Sub
        
    End If
    frmMediaConsumo_Mer.Show 1
End Sub

Private Sub mnCOmport_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmComport.Show 0
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

Private Sub mnCont_Click()
    frmContagens.Show 0
    frmContagens.SetFocus
End Sub

Sub mnCOntEcf_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmImpressoraECF.Show 1
End Sub

Private Sub cmndbtnCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub cmndbtnDev_Click()
    'frmDevolucoes.Show 0
End Sub

Private Sub cmndbtnNewVEnda_Click()
    On Error Resume Next
    frmVenda.Show 0, mdiFrm
End Sub

Private Sub cmndbtnPEd_Click()
    frmEdiPed.Show 0
End Sub

Private Sub cmndbtnProd_Click()
    frmEdiPro.Show 0
    frmEdiPro.SetFocus
End Sub

Private Sub cmndbtnVendas_Click()
    frmEdiVendas.Show 0
    frmEdiVendas.SetFocus
End Sub


Private Sub MDIForm_Activate()
    On Error Resume Next
    Dim strSQL As String
    Me.caption = constNome_sis & " " & constVersao_Sis & " - " & strNome_Emp
    VerificaAtualizacao
    StatusBar1.Panels(6).Text = Winsock1.LocalIP & "-" & Winsock1.LocalHostName
    strIP_Usu = Winsock1.LocalIP
    strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixas", "ip_cai='" & strIP_Usu & "'")
    strSQL = F_UsuLogadoCaixa(strCod_Cai)
    If strSQL = "" Then
'        MsgBox "Não existe usuário logado neste caixa!", vbExclamation
        
    ElseIf strSQL <> strNome_Usu Then
        strSQL = "update tab_turncai set cod_usu = '" & strCod_Usu & "' "
        strSQL = strSQL & " where cod_cai='" & sCod_Cai & "' "
        strSQL = strSQL & " and cod_Tur='" & strCod_Tur & "' "
        strSQL = strSQL & " and data_tur=" & F_GravaData(Date)
        banco.Execute strSQL
    End If
   
    If strCod_Cai = "" Then
        strCod_Cai = "01"
    End If
    If bActive <> True Then

        bActive = True
        
    End If
    frmMenu.Show 0
End Sub

Private Sub MDIForm_DblClick()
    'frmVenda.Show 0, Me
End Sub

Private Sub MDIForm_Load()
    
    
    'frmLogo.Show 0
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    frmEdiCat.Show 0
End Sub

Private Sub mnCadMar_Click()
    frmEdiMar.Show 0
End Sub

Private Sub mnCadPla_Click()
    frmEdiPlanoContas.Show 0
    frmEdiPlanoContas.SetFocus
End Sub

Private Sub mnCalGuias_Click()
    frmLanDoc.Show 0
End Sub

Private Sub mnCasc_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnCheq_Click()
    frmControleCheques.Show 0
End Sub

Private Sub mnComDes_Click()
    frmEdiDes.Show 0
End Sub

Private Sub mnCOn_Click()
    frmEdiCon.Show 0
End Sub

Private Sub mnConfig_Click()
    frmParametros.Show 0
End Sub

Private Sub mnContasPagar_Click()
    frmEditaMov.Show 0
End Sub

Private Sub mnCred_Click()
    On Error Resume Next
    frmCred.Show 1
End Sub

Private Sub mnDEv_Click()
    Dim strSQL As String
    If strAbrevenda_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    
    
    
    
    'frmDevolucoes.caption = "STYLO PRÓPRIO -Devolução a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
  
    On Error Resume Next
    frmDevolucoes.Show 1
    
End Sub

Private Sub mnEdiComp_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiEntr.Show 0
End Sub


Private Sub mnEdiMovFin_Click()
    
    frmEdiMovFin.Show 0
End Sub

Private Sub mnEdiPer_Click()
    If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmEdiPerdas.Show 0
    frmEdiPerdas.SetFocus
End Sub

Private Sub mnEditCli_Click()
    frmEdiClientes.Show 0
End Sub

Private Sub mnEditLan_Click()
     
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
    frmEdiVendas.Show 0
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

Private Sub mnFechaCaixa_Click()
    If strLancaixa_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    Dim fr As New frmLanCaixa
    fr.Show 0
End Sub

Private Sub mnFinanceiras_Click()
    frmEdiFinan.Show 0
End Sub

Private Sub mnFechamento_Click()
    If strFechacai_usu = "S" Then
        frmFechaCaixa.Show 1
    Else
        MsgBox "Você não tem permissão para fechar o caixa!", vbExclamation
    End If
End Sub

Private Sub mnFormasPag_Click()
    frmEdiFormasPag.Show 0
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
    frmSelectCliente.Show 1
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

Private Sub mnIncMer_Click()
     If strAbrenota_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    'Dim fr As New frmEntradas
    'fr.Show 0
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

Private Sub mnMovFin_Click()
    frmEdiFin.Show 0
    frmEdiFin.SetFocus
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
    
    strSQL = F_GeraNovaVenda("00000", strCod_Usu)
    
    
    
    frmVenda.caption = ConstEmpresa & "-Venda a Consumidor" & " [CAIXA " & strNome_Usu & "]"
    
    
    
    frmVenda.Tag = strSQL
    On Error Resume Next
    frmVenda.Show 0, mdiFrm
    
End Sub

Private Sub mnPed_Click()
    frmEdiPed.Show 0
End Sub

Private Sub mnPisos_Click()
    frmEstoqueIni.Show 1
End Sub

Private Sub mnProd_Click()
    frmEdiPro.Show 0
End Sub

Private Sub mnRelGuias_Click()
    frmRelGuias.Show 0
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
    If strAltpreco_usu = "S" Then
        If Not F_senha Then
            Exit Sub
        End If
        frmProm.Show 0
    Else
        MsgBox "Acesso Negado!", vbCritical
    End If
End Sub

Private Sub mnRelAnaFat_Click()
    If strVisrelfin_usu <> "S" Then
        MsgBox "Acesso Negado!", vbCritical
    Else
        frmRelFatMensal.Show 0
        frmRelFatMensal.SetFocus
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

Private Sub mnRelDev_Click()
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

Private Sub mnRelGiro_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelGiro.Show 0
    frmRelGiro.SetFocus
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

Private Sub mnRelVend_Click()
    If strVisrelest_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation, "Acesso Restrito!"
        Exit Sub
    End If
    frmRelVendas.Show 0
    
End Sub

Private Sub mnTer_Click()
    If strCOD_CLI = "" Then
        If MsgBox("Não há cliente selecionado." & Chr(13) & "Deseja selecionar um agora?" _
            , vbYesNo + vbQuestion) = vbYes Then
            On Error Resume Next
            frmSelectCliente.Show 1
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
        frmSujComp.Show 1
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
    End If
    
End Sub

Private Sub mnTodasJan_Click()
    'On Error Resume Next
   
   If Forms.count > 1 Then
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

Private Sub mnUsu_Click()
    'Dim fr As New frmIncUsu
   ' fr.Show 0
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
     If Panel.Index = 1 Then
        mnMudaCli_Click
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim iCount As Integer
    If Button.Index = 1 Then
        mnEditCli_Click
    ElseIf Button.Index = 2 Then
        mnForn_Click
    ElseIf Button.Index = 3 Then
        mnProd_Click
    ElseIf Button.Index = 4 Then
        mnEdiVen_Click
    ElseIf Button.Index = 5 Then
        mnNovaVenda_Click
    ElseIf Button.Index = 6 Then
        mnLanCaixa_Click
    ElseIf Button.Index = 7 Then
        mnEdiComp_Click
    ElseIf Button.Index = 8 Then
        mnServicos_Click
    ElseIf Button.Index = 9 Then
        If Forms.count = 1 Then
            mnSair_Click
        Else
            'For iCount = 1 To Forms.count - 1
                
               ' MsgBox Forms(iCount).Name
                 Unload Screen.ActiveForm
            'Next
        End If
    End If
    
End Sub



