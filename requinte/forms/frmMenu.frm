VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00755C33&
   BorderStyle     =   0  'None
   Caption         =   "Área de Trabalho"
   ClientHeight    =   9930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Menu de Opções"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   20
      Left            =   6240
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   48
      Top             =   5520
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   20
         Left            =   600
         Picture         =   "frmMenu.frx":0000
         Top             =   90
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrança"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   20
         Left            =   510
         TabIndex        =   49
         Top             =   600
         Width           =   705
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   19
      Left            =   6240
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   46
      Top             =   4440
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monitoramento"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   19
         Left            =   330
         TabIndex        =   47
         Top             =   600
         Width           =   1065
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   19
         Left            =   600
         Picture         =   "frmMenu.frx":0CCA
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   18
      Left            =   6240
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   44
      Top             =   3360
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   18
         Left            =   615
         Picture         =   "frmMenu.frx":1994
         Top             =   90
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle de Cheques"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   18
         Left            =   105
         TabIndex        =   45
         Top             =   600
         Width           =   1515
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   17
      Left            =   6240
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   42
      Top             =   2280
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Movimento Financeiro"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   17
         Left            =   75
         TabIndex        =   43
         Top             =   600
         Width           =   1575
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   17
         Left            =   600
         Picture         =   "frmMenu.frx":265E
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   16
      Left            =   6240
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   40
      Top             =   1200
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contas a Pagar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   16
         Left            =   300
         TabIndex        =   41
         Top             =   600
         Width           =   1125
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   16
         Left            =   720
         Picture         =   "frmMenu.frx":2F28
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   15
      Left            =   6240
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   38
      Top             =   120
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   15
         Left            =   720
         Picture         =   "frmMenu.frx":333F
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contas a Receber"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   210
         TabIndex        =   39
         Top             =   600
         Width           =   1305
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   14
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   36
      Top             =   2280
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   14
         Left            =   600
         Picture         =   "frmMenu.frx":35C3
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acerto"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   615
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   13
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   34
      Top             =   1200
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   13
         Left            =   720
         Picture         =   "frmMenu.frx":3E8D
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Devolução"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   465
         TabIndex        =   35
         Top             =   600
         Width           =   795
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   12
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   32
      Top             =   120
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transferência"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   33
         Top             =   600
         Width           =   1005
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   12
         Left            =   720
         Picture         =   "frmMenu.frx":4417
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   11
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   30
      Top             =   3360
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   11
         Left            =   720
         Picture         =   "frmMenu.frx":49A1
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pacote"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   660
         TabIndex        =   31
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   10
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   28
      Top             =   2280
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crediário"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   29
         Top             =   600
         Width           =   645
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   10
         Left            =   720
         Picture         =   "frmMenu.frx":566B
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   9
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   26
      Top             =   1200
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   9
         Left            =   720
         Picture         =   "frmMenu.frx":5F35
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignações"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   405
         TabIndex        =   27
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   8
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   24
      Top             =   120
      Width           =   1815
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Representantes"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   345
         TabIndex        =   25
         Top             =   600
         Width           =   1155
      End
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   8
         Left            =   720
         Picture         =   "frmMenu.frx":67FF
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox pcCalend 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9930
      Left            =   8565
      ScaleHeight     =   9930
      ScaleWidth      =   2460
      TabIndex        =   16
      Top             =   0
      Width           =   2460
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   0
         Picture         =   "frmMenu.frx":74C9
         ScaleHeight     =   2865
         ScaleWidth      =   2385
         TabIndex        =   17
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
            Picture         =   "frmMenu.frx":11878
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   18
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   1875
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   7
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   14
      Top             =   7680
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   7
         Left            =   600
         Picture         =   "frmMenu.frx":13572
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   660
         TabIndex        =   15
         Top             =   600
         Width           =   285
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   6
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   12
      Top             =   6600
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   6
         Left            =   600
         Picture         =   "frmMenu.frx":1423C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   525
         TabIndex        =   13
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   5
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   10
      Top             =   5520
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   5
         Left            =   600
         Picture         =   "frmMenu.frx":14F06
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   405
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   4
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   4
         Left            =   600
         Picture         =   "frmMenu.frx":15BD0
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   525
         TabIndex        =   9
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   3
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   3
         Left            =   600
         Picture         =   "frmMenu.frx":1649A
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nova Venda"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   345
         TabIndex        =   7
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   2
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   2
         Left            =   600
         Picture         =   "frmMenu.frx":16D64
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compras"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   645
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
      Begin VB.Image imgIco 
         Height          =   480
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmMenu.frx":17A2E
         Picture         =   "frmMenu.frx":193B0
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   510
         TabIndex        =   3
         Top             =   600
         Width           =   585
      End
   End
   Begin VB.PictureBox pcBotao 
      Appearance      =   0  'Flat
      BackColor       =   &H00755C33&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   0
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.Image imgIco 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmMenu.frx":1AD32
         Picture         =   "frmMenu.frx":1C37C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBotao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produtos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   645
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
      TabIndex        =   23
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
    For iCount = 0 To pcBotao.count - 1
        lblBotao(iCount).FontBold = False
        lblBotao(iCount).FontUnderline = False
        
    Next
    lblBotao(Indice).FontBold = True
    lblBotao(Indice).FontUnderline = True
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    P_mensagem
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
    
      '  frmAcerto.Show 0
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
    On Error Resume Next
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

Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
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
