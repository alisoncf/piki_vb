VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVendaFunc 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "-Venda de Produtos:"
   ClientHeight    =   11310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16170
   ControlBox      =   0   'False
   Icon            =   "frmvendafunc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11310
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5520
      ScaleHeight     =   945
      ScaleWidth      =   10545
      TabIndex        =   9
      Top             =   600
      Width           =   10575
      Begin VB.CommandButton cmndbtnInserir 
         Caption         =   "Inserir"
         Height          =   285
         Left            =   8520
         TabIndex        =   114
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtbxPQ_pro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   112
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
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDescReais 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   7680
         TabIndex        =   19
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
         TabIndex        =   17
         Text            =   "0,0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4680
         TabIndex        =   13
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
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
         TabIndex        =   113
         Top             =   0
         Width           =   405
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
         TabIndex        =   111
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(R$)"
         Height          =   195
         Left            =   7680
         TabIndex        =   18
         Top             =   0
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
         Left            =   8640
         TabIndex        =   20
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(%)"
         Height          =   195
         Left            =   6840
         TabIndex        =   16
         Top             =   0
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
         Left            =   5160
         TabIndex        =   14
         Top             =   0
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
         TabIndex        =   10
         Top             =   0
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
         TabIndex        =   12
         Top             =   0
         Width           =   300
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
      TabIndex        =   85
      Top             =   600
      Width           =   5415
      Begin VB.PictureBox pcDesconto 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1065
         ScaleWidth      =   5145
         TabIndex        =   119
         Top             =   1680
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtbxDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   122
            Text            =   "0,00"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtbxDescPerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   121
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmndbtnDesconto 
            Caption         =   "Confirmar Gerencial"
            Height          =   525
            Left            =   1920
            Picture         =   "frmvendafunc.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   240
            Width           =   1575
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
            TabIndex        =   130
            Top             =   -30
            Width           =   1860
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
            TabIndex        =   128
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Autorizado Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   127
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto:"
            Height          =   195
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   1200
            TabIndex        =   125
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final Previsto:"
            Height          =   195
            Left            =   4080
            TabIndex        =   124
            Top             =   240
            Width           =   990
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
            TabIndex        =   123
            Top             =   480
            Width           =   630
         End
      End
      Begin VB.PictureBox pcLimite 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   5175
         TabIndex        =   101
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
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
            TabIndex        =   109
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   2880
            TabIndex        =   108
            Top             =   0
            Width           =   690
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
            TabIndex        =   107
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   1920
            TabIndex        =   106
            Top             =   0
            Width           =   600
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
            TabIndex        =   105
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1080
            TabIndex        =   104
            Top             =   0
            Width           =   405
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
            TabIndex        =   103
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox pcEnd 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   -360
         ScaleHeight     =   1455
         ScaleWidth      =   9135
         TabIndex        =   100
         Top             =   2640
         Visible         =   0   'False
         Width           =   9135
         Begin VB.CheckBox chkEntrega 
            Caption         =   "En&trega (ctrl + E)"
            Height          =   255
            Left            =   1200
            TabIndex        =   118
            Top             =   120
            Width           =   1575
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
            TabIndex        =   27
            Top             =   1080
            Width           =   1455
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
            TabIndex        =   25
            Top             =   1080
            Width           =   3615
         End
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
            TabIndex        =   23
            Top             =   480
            Width           =   5175
         End
         Begin VB.Label lblFone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fone:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3840
            TabIndex        =   26
            Top             =   840
            Width           =   405
         End
         Begin VB.Label lblBairro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   450
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtbxDescForma 
         Appearance      =   0  'Flat
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
         Left            =   3120
         TabIndex        =   4
         Text            =   "20"
         Top             =   240
         Width           =   375
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
         TabIndex        =   1
         Text            =   "A Vista"
         Top             =   240
         Width           =   2655
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
         TabIndex        =   7
         Tag             =   "00000"
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   315
         Left            =   2760
         Picture         =   "frmvendafunc.frx":1034
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmnbtnRecalc 
         Caption         =   "Recalcula Desconto"
         Enabled         =   0   'False
         Height          =   435
         Left            =   3720
         Picture         =   "frmvendafunc.frx":112D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   330
         Left            =   5040
         Picture         =   "frmvendafunc.frx":14B7
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   255
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
         TabIndex        =   110
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Funcionário:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Forma de Pagamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   1560
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10275
      Left            =   0
      ScaleHeight     =   10245
      ScaleWidth      =   5385
      TabIndex        =   61
      Top             =   540
      Width           =   5415
      Begin VB.PictureBox pcLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         Picture         =   "frmvendafunc.frx":15B0
         ScaleHeight     =   2865
         ScaleWidth      =   2385
         TabIndex        =   117
         Top             =   6840
         Width           =   2415
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   5385
         TabIndex        =   77
         Top             =   3240
         Width           =   5415
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   86
            Top             =   1680
            Width           =   5415
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
               TabIndex        =   90
               Top             =   0
               Width           =   570
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
               TabIndex        =   89
               Tag             =   "!"
               Top             =   0
               Width           =   105
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
               TabIndex        =   88
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
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
               TabIndex        =   87
               Top             =   0
               Width           =   930
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   79
            Top             =   0
            Width           =   5415
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
               TabIndex        =   83
               Top             =   0
               Width           =   510
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
               TabIndex        =   82
               Tag             =   "!"
               Top             =   0
               Width           =   105
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
               TabIndex        =   81
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
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
               TabIndex        =   80
               Top             =   0
               Width           =   900
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
            TabIndex        =   78
            Tag             =   "!"
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   0
         ScaleHeight     =   4665
         ScaleWidth      =   5385
         TabIndex        =   62
         Top             =   5400
         Width           =   5415
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc %"
            Height          =   195
            Left            =   4440
            TabIndex        =   99
            Top             =   2640
            Width           =   540
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
            TabIndex        =   98
            Top             =   2880
            Width           =   885
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            Height          =   195
            Left            =   2640
            TabIndex        =   76
            Top             =   2640
            Width           =   825
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
            TabIndex        =   75
            Top             =   2880
            Width           =   1605
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   2640
            TabIndex        =   74
            Top             =   1920
            Width           =   540
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
            TabIndex        =   73
            Top             =   2160
            Width           =   2565
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
            TabIndex        =   72
            Top             =   3600
            Width           =   2565
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Seção:"
            Height          =   195
            Left            =   2640
            TabIndex        =   71
            Top             =   3360
            Width           =   840
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
            TabIndex        =   70
            Top             =   960
            Width           =   5085
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seção:"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   4320
            TabIndex        =   68
            Top             =   1320
            Width           =   870
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
            TabIndex        =   67
            Top             =   1560
            Width           =   885
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
            TabIndex        =   66
            Top             =   1560
            Width           =   1605
         End
         Begin VB.Label lblCapPreco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço:"
            Height          =   195
            Left            =   2640
            TabIndex        =   65
            Top             =   1320
            Width           =   465
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
            TabIndex        =   64
            Top             =   360
            Width           =   5085
         End
         Begin VB.Label lblCapPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Queima:"
         Height          =   195
         Left            =   2520
         TabIndex        =   116
         Top             =   3000
         Width           =   1215
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
         TabIndex        =   115
         Top             =   2880
         Width           =   1155
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
         TabIndex        =   84
         Top             =   5160
         Width           =   150
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
      ScaleWidth      =   16140
      TabIndex        =   51
      Top             =   10815
      Width           =   16170
      Begin VB.CommandButton cmndbtnBuscaFun 
         Height          =   315
         Left            =   3840
         Picture         =   "frmvendafunc.frx":B95F
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   0
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
         Left            =   840
         TabIndex        =   94
         Text            =   "Nome do Funcionario"
         Top             =   15
         Width           =   3255
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
         TabIndex        =   129
         Top             =   52
         Visible         =   0   'False
         Width           =   90
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
         TabIndex        =   97
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   96
         Top             =   60
         Width           =   420
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
         TabIndex        =   60
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
         Left            =   10560
         TabIndex        =   59
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
         Left            =   9360
         TabIndex        =   58
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
         Left            =   8760
         TabIndex        =   57
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8160
         TabIndex        =   56
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
         Left            =   12480
         TabIndex        =   55
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5760
         TabIndex        =   54
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
         Left            =   6600
         TabIndex        =   53
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atendente:"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   52
         Top             =   15
         Width           =   840
      End
   End
   Begin VB.PictureBox pcAtalhos 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   16140
      TabIndex        =   50
      Top             =   0
      Width           =   16170
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   91
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
            Picture         =   "frmvendafunc.frx":BA58
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   92
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
            TabIndex        =   93
            Top             =   120
            Width           =   1050
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   28
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
            Picture         =   "frmvendafunc.frx":C45A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   29
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
            TabIndex        =   30
            Top             =   120
            Width           =   1200
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   3600
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
            Index           =   2
            Left            =   120
            Picture         =   "frmvendafunc.frx":C7E4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   32
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
            TabIndex        =   33
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   5400
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
            Index           =   3
            Left            =   120
            Picture         =   "frmvendafunc.frx":CA58
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   35
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
            TabIndex        =   36
            Top             =   120
            Width           =   1170
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   7200
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   37
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
            Picture         =   "frmvendafunc.frx":CBA2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
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
            TabIndex        =   39
            Top             =   120
            Width           =   990
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   9000
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   40
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
            Picture         =   "frmvendafunc.frx":D5A4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   41
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
            TabIndex        =   42
            Top             =   120
            Width           =   1050
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   10800
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   43
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmvendafunc.frx":F96B
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   44
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
            TabIndex        =   45
            Top             =   120
            Width           =   1020
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   12600
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   46
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
            Picture         =   "frmvendafunc.frx":FCF5
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   47
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
            TabIndex        =   48
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmndbtnLocaliza 
      Caption         =   "..."
      Height          =   255
      Left            =   14760
      TabIndex        =   49
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmvendafunc.frx":FE60
      Height          =   2775
      Left            =   5520
      TabIndex        =   131
      Top             =   1680
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4895
      _Version        =   393216
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "seq_ite"
         Caption         =   "..."
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
         Caption         =   "Referência"
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
         DataField       =   "desc_ite"
         Caption         =   "Descrição do Produto"
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
         DataField       =   "qtd_ite"
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
         DataField       =   "unit_ite"
         Caption         =   "Preço"
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
      BeginProperty Column05 
         DataField       =   "total"
         Caption         =   "Total"
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
      BeginProperty Column06 
         DataField       =   "desconto"
         Caption         =   "Desconto"
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
         DataField       =   "final"
         Caption         =   "Final"
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
         DataField       =   "pq"
         Caption         =   "Queima"
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
      BeginProperty Column09 
         DataField       =   "nome_sub"
         Caption         =   "Subseção"
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
         BeginProperty Column00 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2849,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   689,953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adc 
      Height          =   495
      Left            =   8160
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmVendaFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p As Produto
Dim bBlock As Boolean
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

Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    If bActive = False Then
        strSQL = F_VerificaVendaAberta(Me.tag)
        If strSQL <> "" Then
            MsgBox "A venda se encontra aberta no computador " & strSQL, vbExclamation
            Unload Me
            Exit Sub
        End If
    End If
    txtbxDescForma.Text = F_BuscaUmCampo("descontofun_par", "tab_parametros", "")
    lblDescontoMax.caption = txtbxDescForma.Text
    
    strSQL = "select tv.*,tf.*,tpp.nome_tip from tab_vendas as tv, tab_formaspag as tf, tab_tipopag as tpp "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for "
    strSQL = strSQL & " and num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tpp.cod_tip=tf.cod_tip "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        txtbxNomeCli.Text = !nomecli_ven
        txtbxNomeCli.tag = !cod_cli
        
        If UCase(txtbxNomeCli.Text) = "VENDA A CONSUMIDOR" Then
            txtbxNomeCli.Text = ""
        End If
        txtbxFormaPag.Text = !nome_for
        'txtbxDescForma.Text = !desconto_for
        txtbxFormaPag.tag = !cod_For
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
trata_erro:
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
    On Error GoTo trata_erro:
    Dim cPq As Currency
    P_Status "Atualizando Valores..."
    
    iPos = adc.Recordset.AbsolutePosition
    adc.Refresh
    With adc.Recordset
        
        
        Do While Not .EOF
            cTotal = cTotal + !final
            cPq = cPq + !pq_itd * !qtd_ite
            cItens = cItens + !qtd_ite
            cDesc = cDesc + !desconto_ite
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cTotal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = adc.Recordset.RecordCount
        lblTotalQueima.caption = FormatVal(cPq)
        lblDesconto.caption = FormatVal(cDesc)
        lblTotal.caption = FormatVal(cTotal + cDesc)
    End With
        
    
    
    
    adc.Refresh
    adc.Recordset.Move iPos
    
    lblProd.caption = adc.Recordset.RecordCount
    lblItens.caption = cItens

    strSQL = "update tab_vendas set valor_ven =" & F_GravaDec(cTotal) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(cDesc) & ","
    strSQL = strSQL & "nomecli_ven= '" & txtbxNomeCli.Text & "', "
    strSQL = strSQL & " cod_cli= '" & txtbxNomeCli.tag & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "', "
    strSQL = strSQL & " abrevenda_usu='" & txtbxVendedor.tag & "' "
'    strSQL = strSQL & " obs_VEN='" & txtbxObs.Text & "' "
    strSQL = strSQL & " where num_Ven='" & Me.tag & "' "
    Conn.Execute strSQL
    
    
    P_Status ""
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = Vendas.F_BuscaItens(Me.tag)
    adc.RecordSource = strSQL
    adc.Refresh
    
            
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub cmndbtnCancel_Click()
    Dim strSQL  As String
    Dim i As Integer
    On Error Resume Next
    If adc.Recordset.RecordCount > 0 Then
    
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
        
        With adc.Recordset
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                Excluir False
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
        Conn.Execute strSQL
        
        
        If F_BuscaUmCampo("num_ven", "tab_venfec", "num_ven='" & Me.tag & "'") <> "" Then
            MsgBox "Cancelamento impossível", vbExclamation
            Exit Sub
        End If
        
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
            'RefreshDataPosition frmEdiVendas.adc
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
    


trata_erro:
MsgBox "Err: cmndbtnEditProd_Click() "
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Excluir True
End Sub
Sub Excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_erro
    With adc.Recordset
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
        Conn.Execute strSQL
        P_GravaItemCancelado
        
        
            strSQL = "update tab_produtos set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))
            
           
            
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
        
        adc.Refresh
        If iPos = .RecordCount - 1 Then
            If adc.Recordset.RecordCount > 0 Then
                adc.Recordset.MoveLast
            End If
        Else
            adc.Recordset.Move iPos
        End If
        If bQuestion Then
            P_AtualizaVenda
        End If
        If adc.Recordset.RecordCount = 0 Then
            txtbxQtd.Enabled = False
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
    Exit Sub
trata_erro:
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
     Dim f As funcionario.funcionario
     f = F_SelectFuncionario(txtbxNomeCli.Text)
     
     If f.cod_fun <> "" Then
        txtbxNomeCli.Text = f.nome_fun
        txtbxNomeCli.tag = f.cod_fun
     End If
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
    f = F_SelectFormasPag("")
    
    If f.cod_For <> "" Then
        'If CCur(txtbxDescForma.Text) > f.desconto_for Then
        '    If adc.Recordset.RecordCount > 0 Then
        '        strSQL = "Você optou por mudar a forma de pagamento de " & txtbxFormaPag.Text
        '        strSQL = strSQL & " para " & f.nome_for & Chr(13)
        '        strSQL = strSQL & "Continuando com esta operação, todo o desconto da venda será cancelado!"
        '        strSQL = strSQL & "Deseja continuar?"
        '        If MsgBox(strSQL, vbOKCancel) = vbCancel Then
        '            Exit Sub
        '        Else
        '            P_CancelaDescontoVenda Me.tag
        '           P_buscaitens
        '        End If
        '    End If
        'End If
        txtbxFormaPag.Text = f.nome_for
        txtbxFormaPag.tag = f.cod_For
        'txtbxDescForma.Text = f.desconto_for
        txtbxDescForma.tag = f.avista_For
        'lblDescontoMax.caption = f.desconto_for + CCur(lblDesc.caption)
        'txtbxDescPercent.Text = txtbxDescForma.Text
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
    On Error GoTo trata_erro
    
    If adc.Recordset.RecordCount > 0 Then
        
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
        'frmEdiVendas.adc.Refresh
    End If
    If adc.Recordset.RecordCount = 0 Then
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
    End If
    P_DeletaVendaAberta Me.tag
    Unload Me
    Exit Sub
trata_erro:
    P_MostraErro "cmndbtnSair"
    
End Sub


Private Sub cmndbtnSalvar_Click()
    Dim sTroco As String
    Dim strSQL As String
    Dim sVen As String
    Dim s As String
    If adc.Recordset.RecordCount = 0 Then
        MsgBox "Não há itens na sua venda!", vbExclamation
        Exit Sub
    End If
    
    
    strSQL = F_senhaReturnUsu(False, True)
    If strSQL = "" Then
        MsgBox "não autenticado!", vbExclamation
        Exit Sub
    End If
    If txtbxNomeCli.tag = "" Or txtbxNomeCli.tag = "00000" Then
        MsgBox "Informe o nome de um funcionário válido!", vbExclamation
        txtbxNomeCli.SetFocus
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
        
        If bUsaCrediario Then
            If lblTipoPag.tag = "02" Then
                If F_GravaCrediarioVen(Me.tag) Then
                    MsgBox "Esta venda venda deve ser liberada pelo setor de Crediário para ser concretizada!", vbExclamation
                Else
                    MsgBox "Houve erro na gravação do crediário", vbExclamation
                End If
            End If
        End If
        
        
        If Vendas.F_AlteraStatusVenda(Me.tag, "2") Then
            
            If Not bUsaCaixaSeparado Then
                s = F_Troco(Me.tag)
                If s = "" Then
                    Exit Sub
                End If
            End If
            
            P_DeletaVendaAberta Me.tag
            Me.tag = F_GeraNovaVenda
            If Me.tag = "" Then
                MsgBox "Não foi possível gerar uma nova venda!", vbExclamation
                Unload Me
                Exit Sub
            End If
            bActive = False
            Form_Activate
        End If
    
        
    
End Sub


Private Sub dtgrd_DblClick()
    editaItem
End Sub

Private Sub dtgrd_GotFocus()
    
    dtgrd.MarqueeStyle = 3
    
    
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo trata_erro

    '    On Error GoTo trata_erro

    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        
        
                
                    If dtgrd.row = -1 Then
                        MsgBox "Selecione o item que deseja excluir!", vbCritical
                        dtgrd.SetFocus
                        Exit Sub
                    End If
                    
            Excluir False
        
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxQtd.SetFocus
    End If
    P_AtualizaVenda
    Exit Sub
trata_erro:
P_MostraErro

End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub


Private Sub dtgrd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub adc_Reposition()
    With adc.Recordset
    On Error Resume Next
    If Not .EOF Then
        lblNomePro.caption = !nome_pro
            
            
                lblCapPreco.caption = "Preço:"
                lblPreco.caption = "R$ " & FormatVal(!pv_pro)
            
            lblSec.caption = !Nome_Sec
            lblCodPro.caption = !cod_pro
            lblRef.caption = !ref_pro
            lblSubSec.caption = !nome_sub
            lblQtd.caption = !qtd_ite
            
        End If
    End With
        
End Sub


Sub buscaPromo()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
        
    strSQL = "select nome_pro,pv_pro from tab_produtos as tp,"
    strSQL = strSQL & " tab_prodpromo as tpp where tpp.cod_Pro=tp.cod_pro"
    strSQL = strSQL & " and datafin_pro >=" & F_GravaData(Date)
    strSQL = strSQL & " order by nome_pro"
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        QtdPromo = .RecordCount
        If .RecordCount > 0 Then
            ReDim pPromo(.RecordCount) As prodPromo
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
Private Sub Form_Activate()
    On Error GoTo trata_erro
    mdiFrm.Enabled = False
    'If LCase(Left(strNomeBanco, 8)) <> "requinte" Then
    '    pcLogo.Visible = False
    'End If
    lblDesc.caption = crcDesconto_usu
    lblNomeEmp.caption = strFantasia_Emp
    lblNum_Ven.caption = Me.tag
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
    ElseIf KeyCode = vbKeyF5 Then
        
        dtgrd.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
         dtgrd_KeyDown vbKeyDelete, 0
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
    Conn.Execute strSQL
End Sub
Private Sub Form_Load()
'    buscaPromo
    bActive = False

    adc.ConnectionString = Conn.ConnectionString
    

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
    dtgrd.width = Me.width - dtgrd.Left - 200
    
    
    dtgrd.height = pcStatusBar.top - dtgrd.top - 100
    lblNomeEmp.Left = Me.width - lblNomeEmp.width - 200
    'pcCli.Width = pcProd.Width
    
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).width = pcAtalhos.width / (pcBotao.Count - 1)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).width + 50
        End If
        
    Next
 
    pcProd.width = dtgrd.width
    'pcCli.Width = pcProd.Width
    

End Sub

Sub editaItem()
    'On Error GoTo trata_Erro
    Exit Sub 'para terminar depois
    With adc.Recordset
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
trata_erro:
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
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BorderStyle = 0
        
    Next
    pcBotao(Indice).BorderStyle = 1
'    pcbotao(Indice).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    adc.Recordset.Close
End Sub

Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_GotFocus(Index As Integer)
    Sendkeys Chr(9)
End Sub



Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If Index = 0 Then
            pcBotao(pcBotao.Count - 1).SetFocus
        Else
            pcBotao(Index - 1).SetFocus
        End If
        
    ElseIf KeyCode = vbKeyDown Then
        If Index = pcBotao.Count - 1 Then
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
        txtbxDescForma.Text = lblDescontoMax.caption
        Exit Sub
    End If
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
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
        dtgrd.SetFocus
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescPercent_LostFocus()
    P_ValorLostFocus txtbxDescPercent
    On Error GoTo trata_erro
    txtbxDescReais.Text = FormatVal(CCur(txtbxDescPercent.Text) * CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) / 100)
    
    
    
    
    Exit Sub
trata_erro:
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
    On Error GoTo trata_erro
'    On Error Resume Next
    'txtbxDescReais.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxFinal.Text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    
    
    
    txtbxDescReais.Text = FormatVal((CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text)) - CCur(txtbxFinal.Text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    txtbxDescReais_LostFocus
trata_erro:
    
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
        Sendkeys Chr(9)
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
        
        
        
        
        If p.pq_pro > 0 Then  'And lblTipoPag.tag = "01"
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
    Dim rPesq As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        buscaProd

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
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoQtd As String
    If lblQueima.Visible = False Then
        If CCur(Round(txtbxDescPercent.Text)) > CCur(lblDesc.caption) + CCur(txtbxDescForma.Text) Then
            MsgBox "O desconto máximo permitido para o usuário é de " & CCur(lblDesc.caption) + CCur(txtbxDescForma.Text) & "%", vbExclamation
            Exit Sub
        End If
    Else
        
    End If
    If txtbxProduto.Text = "" Then
        MsgBox "Informe pelo menos parte do nome do produto para que seja localizado!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    On Error GoTo trata_erro
                
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
                adc.Recordset.Move adc.Recordset.RecordCount - 1
                txtbxUnit.Text = "0"
                txtbxDescPercent.Text = "0,0"
                txtbxDescReais.Text = "0,00"
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
            
        
        P_AtualizaVenda
        On Error Resume Next

        Exit Sub
trata_erro:
    P_MostraErro "Insere()"
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
Exit Sub
trata_erro:
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
    strSQL = strSQL & "" & adc.Recordset!seq_ite & ","
    strSQL = strSQL & F_GravaDec(adc.Recordset!qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(adc.Recordset!unit_ite) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "') "
    Conn.Execute strSQL
    
    
    
End Sub




