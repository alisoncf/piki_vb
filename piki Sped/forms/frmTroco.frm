VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTroco 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Troco"
   ClientHeight    =   9105
   ClientLeft      =   225
   ClientTop       =   -435
   ClientWidth     =   16875
   ControlBox      =   0   'False
   Icon            =   "frmTroco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   16875
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcDesconto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   5865
      TabIndex        =   126
      Top             =   7440
      Width           =   5895
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   129
         Text            =   "0,00"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxDescPerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   128
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmndbtnDesconto 
         Caption         =   "Confirmar"
         Height          =   525
         Left            =   4800
         Picture         =   "frmTroco.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   120
         TabIndex        =   133
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   1200
         TabIndex        =   132
         Top             =   120
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final Previsto:"
         Height          =   195
         Left            =   1800
         TabIndex        =   131
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblFinalPrev 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Left            =   1800
         TabIndex        =   130
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   5775
      TabIndex        =   122
      Top             =   4800
      Width           =   5775
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2640
         TabIndex        =   123
         Top             =   0
         Width           =   150
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   5775
      TabIndex        =   117
      Top             =   5160
      Width           =   5775
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
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
         Left            =   120
         TabIndex        =   121
         Top             =   0
         Width           =   1110
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
         ForeColor       =   &H00916024&
         Height          =   240
         Left            =   1320
         TabIndex        =   120
         Tag             =   "!"
         Top             =   0
         Width           =   105
      End
      Begin VB.Label lblDesconto 
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
         ForeColor       =   &H00916024&
         Height          =   240
         Left            =   4800
         TabIndex        =   119
         Tag             =   "!"
         Top             =   0
         Width           =   105
      End
      Begin VB.Label Label2 
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
         Left            =   3720
         TabIndex        =   118
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1815
      ScaleWidth      =   5775
      TabIndex        =   110
      Top             =   5280
      Width           =   5775
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   120
         Picture         =   "frmTroco.frx":1434
         ScaleHeight     =   1140
         ScaleWidth      =   2220
         TabIndex        =   115
         Top             =   600
         Width           =   2220
         Begin VB.Label lblTotalQueima 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   720
            TabIndex        =   116
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   3075
         Picture         =   "frmTroco.frx":9846
         ScaleHeight     =   1140
         ScaleWidth      =   2220
         TabIndex        =   113
         Top             =   600
         Width           =   2220
         Begin VB.Label lblDemaisProd 
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   840
            TabIndex        =   114
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demais Produtos:"
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
         Left            =   3075
         TabIndex        =   112
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Promoção:"
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
         Left            =   120
         TabIndex        =   111
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar Venda (F6)"
      Height          =   855
      Left            =   11400
      Picture         =   "frmTroco.frx":11C58
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox txtbxTipoCob 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      TabIndex        =   60
      Tag             =   "1"
      Text            =   "Pagamento na Loja"
      Top             =   8640
      Width           =   1695
   End
   Begin VB.PictureBox pcAprazo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   10200
      ScaleHeight     =   4785
      ScaleWidth      =   6225
      TabIndex        =   77
      Top             =   3000
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtbxMsflx 
         Height          =   285
         Left            =   1200
         TabIndex        =   106
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnEdi 
         Height          =   375
         Left            =   4920
         Picture         =   "frmTroco.frx":11DA2
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Editar o cheque ou  parcela que você selecionou na grade abaixo..."
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkCh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cheque:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   0
         Width           =   1095
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   6015
         TabIndex        =   78
         Top             =   240
         Width           =   6015
         Begin VB.TextBox txtbxCPFChe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   40
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtbxAgencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   49
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtbxNome_Che 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   42
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtbxNum_Che 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            MaxLength       =   12
            TabIndex        =   47
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxContaCorrente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            MaxLength       =   12
            TabIndex        =   51
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtbxBanco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   44
            Top             =   960
            Width           =   2535
         End
         Begin VB.CommandButton cmndbtnBuscaBanco 
            Height          =   285
            Left            =   2640
            Picture         =   "frmTroco.frx":127A4
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   960
            Width           =   255
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&CPF:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   345
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Agência:"
            Height          =   195
            Left            =   3960
            TabIndex        =   48
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Cheque Nº:"
            Height          =   195
            Left            =   3000
            TabIndex        =   46
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Titular:"
            Height          =   195
            Left            =   2040
            TabIndex        =   41
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&onta C.:"
            Height          =   195
            Left            =   4800
            TabIndex        =   50
            Top             =   720
            Width           =   660
         End
      End
      Begin VB.TextBox txtbxVencimento 
         Appearance      =   0  'Flat
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   53
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtbxValor_Che 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   14
         TabIndex        =   55
         Text            =   "0,00"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnAdd 
         Height          =   375
         Left            =   4440
         Picture         =   "frmTroco.frx":1289D
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Adiciona o Cheque ou parcela que você digitou.."
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnExcl 
         Height          =   375
         Left            =   5400
         Picture         =   "frmTroco.frx":12E27
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Exclui o cheque selecionado na grade abaixo"
         Top             =   1680
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2535
         Left            =   120
         TabIndex        =   64
         Top             =   2160
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   14737632
         AllowUserResizing=   3
         Appearance      =   0
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Height          =   195
         Left            =   1800
         TabIndex        =   54
         Top             =   1560
         Width           =   405
      End
   End
   Begin VB.PictureBox pcDescAutorizadoPor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   3015
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmndbtnCancelarDescGer 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2040
         TabIndex        =   105
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblAutorizaDesconto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   104
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto Autorizado Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   103
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "Efetivar (F7/F8)"
      Height          =   855
      Left            =   13080
      Picture         =   "frmTroco.frx":151EE
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   8280
      Width           =   1575
   End
   Begin VB.PictureBox Picture10 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   16875
      TabIndex        =   82
      Top             =   0
      Width           =   16875
      Begin VB.PictureBox pcLimite 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AC67&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6120
         ScaleHeight     =   615
         ScaleWidth      =   10575
         TabIndex        =   85
         Top             =   0
         Width           =   10575
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblStatus_CLi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liberado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1440
            TabIndex        =   91
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblLimite 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1440
            TabIndex        =   90
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   3240
            TabIndex        =   89
            Top             =   0
            Width           =   600
         End
         Begin VB.Label lblEmUso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Left            =   3240
            TabIndex        =   88
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   5040
            TabIndex        =   87
            Top             =   0
            Width           =   690
         End
         Begin VB.Label lblLimiteRest 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   5040
            TabIndex        =   86
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   1560
         TabIndex        =   101
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente:"
         Height          =   195
         Left            =   240
         TabIndex        =   100
         Top             =   0
         Width           =   780
      End
      Begin VB.Label lblCaixa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   240
         TabIndex        =   84
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblNome_ven 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   83
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.PictureBox pcFormasPag 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6C1A6&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   6120
      ScaleHeight     =   6105
      ScaleWidth      =   10545
      TabIndex        =   79
      Top             =   2040
      Width           =   10575
      Begin VB.PictureBox pcEntrada 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   120
         ScaleHeight     =   4785
         ScaleWidth      =   3825
         TabIndex        =   94
         Top             =   960
         Width           =   3855
         Begin VB.TextBox txtDeposito 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   29
            Text            =   "0,00"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtDescricaoPIX 
            Appearance      =   0  'Flat
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   35
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox txtPIX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   26
            Text            =   "0,00"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxEntradaChv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   32
            Text            =   "0,00"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtbxEntradaDin 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   21
            Text            =   "0,00"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxEntradaCart 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmndbtnDadosCard 
            Height          =   285
            Left            =   2640
            Picture         =   "frmTroco.frx":16EE8
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnDadosChv 
            Height          =   285
            Left            =   2640
            Picture         =   "frmTroco.frx":17272
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton cmdPix 
            Height          =   285
            Left            =   2640
            Picture         =   "frmTroco.frx":175FC
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton cmdDep 
            Height          =   285
            Left            =   2640
            Picture         =   "frmTroco.frx":17986
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblCCoDep 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta"
            Height          =   195
            Left            =   3000
            TabIndex        =   136
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label lblCcoPix 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta"
            Height          =   195
            Left            =   3000
            TabIndex        =   135
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito:"
            Height          =   195
            Left            =   765
            TabIndex        =   28
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label lblDescricaoPIX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento:"
            Height          =   195
            Left            =   435
            TabIndex        =   34
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PIX:"
            Height          =   195
            Left            =   1140
            TabIndex        =   25
            Top             =   960
            Width           =   300
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&heque:"
            Height          =   195
            Left            =   840
            TabIndex        =   31
            Top             =   1680
            Width           =   600
         End
         Begin VB.Label lblEntrada 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dinheiro:"
            Height          =   195
            Left            =   840
            TabIndex        =   20
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cartão:"
            Height          =   195
            Left            =   930
            TabIndex        =   22
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lblTotalEntrada 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   1680
            TabIndex        =   96
            Top             =   2880
            Width           =   630
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total A Vista:"
            Height          =   195
            Left            =   240
            TabIndex        =   95
            Top             =   3000
            Width           =   945
         End
      End
      Begin VB.TextBox txtbxValorDoc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8160
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDataBase 
         Appearance      =   0  'Flat
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
         Left            =   9240
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxInterv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtbxParcelas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   13
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtbxFormaPag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtbxDescForma 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   285
         Left            =   4680
         Picture         =   "frmTroco.frx":17D10
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo:"
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
         Left            =   4440
         TabIndex        =   137
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo:"
         Height          =   195
         Left            =   8160
         TabIndex        =   16
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblCapEntrada 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada/A Vista:"
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
         Left            =   120
         TabIndex        =   97
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label lblDataBase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base:"
         Height          =   195
         Left            =   9240
         TabIndex        =   18
         Top             =   120
         Width           =   795
      End
      Begin VB.Label lblInterMes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mês:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7200
         TabIndex        =   81
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblInternv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo."
         Height          =   195
         Left            =   6720
         TabIndex        =   14
         Top             =   120
         Width           =   660
      End
      Begin VB.Label lblParcelas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas:"
         Height          =   195
         Left            =   6000
         TabIndex        =   12
         Top             =   120
         Width           =   660
      End
      Begin VB.Label lblTipoCai 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIN"
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
         Left            =   4320
         TabIndex        =   80
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pagamento (F3):"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1875
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   5040
         TabIndex        =   11
         Top             =   120
         Width           =   120
      End
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   6120
      ScaleHeight     =   1185
      ScaleWidth      =   10545
      TabIndex        =   76
      Top             =   720
      Width           =   10575
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   125
         Top             =   840
         Width           =   6615
      End
      Begin VB.TextBox txtbxCNPJ 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   4320
         Picture         =   "frmTroco.frx":17E09
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   124
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/CNPJ:"
         Height          =   195
         Left            =   4800
         TabIndex        =   5
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblCliCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente (F2):"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   -120
      ScaleHeight     =   135
      ScaleWidth      =   17295
      TabIndex        =   73
      Top             =   9240
      Width           =   17295
   End
   Begin VB.CommandButton cmndbtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar (F12)"
      Height          =   855
      Left            =   15120
      Picture         =   "frmTroco.frx":17F02
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CheckBox chkEmite 
      Caption         =   "Imprimir Recibo"
      Height          =   735
      Left            =   120
      Picture         =   "frmTroco.frx":18450
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   8400
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   120
      Picture         =   "frmTroco.frx":185B4
      ScaleHeight     =   1170
      ScaleWidth      =   5835
      TabIndex        =   69
      Top             =   1200
      Width           =   5835
      Begin VB.Label lblValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F8AC67&
         Height          =   480
         Left            =   3390
         TabIndex        =   71
         Tag             =   "!"
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Final:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcDin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   120
      Picture         =   "frmTroco.frx":2E9D6
      ScaleHeight     =   1170
      ScaleWidth      =   5835
      TabIndex        =   68
      Top             =   2400
      Width           =   5835
      Begin VB.TextBox txtbxDinheiro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   660
         Left            =   2400
         TabIndex        =   1
         Text            =   "0,00"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Dinheiro:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.PictureBox pcTroco 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   120
      Picture         =   "frmTroco.frx":44DF8
      ScaleHeight     =   1140
      ScaleWidth      =   5805
      TabIndex        =   65
      Top             =   3600
      Width           =   5805
      Begin VB.Label lblTroco 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   600
         Left            =   3315
         TabIndex        =   67
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Troco:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   240
         TabIndex        =   66
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.CommandButton cmndbtnCob 
      Height          =   285
      Left            =   8280
      Picture         =   "frmTroco.frx":5A7CA
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   8640
      Width           =   255
   End
   Begin VB.OptionButton optnAgora 
      Caption         =   "Lançar Dados Agora"
      Height          =   195
      Left            =   6240
      TabIndex        =   36
      Top             =   4320
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optnDepois 
      Caption         =   "Lançar Dados Mais Tarde"
      Height          =   195
      Left            =   8880
      TabIndex        =   37
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto Gerencial (F8):"
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
      Left            =   120
      TabIndex        =   134
      Top             =   7200
      Width           =   2145
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restante:"
      Height          =   195
      Left            =   2520
      TabIndex        =   108
      Top             =   8400
      Width           =   690
   End
   Begin VB.Label lblRestante 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   2520
      TabIndex        =   107
      Top             =   8640
      Width           =   630
   End
   Begin VB.Label lblTipoCob 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cobr:"
      Height          =   195
      Left            =   6480
      TabIndex        =   59
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblTotalDistribuido 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   4080
      TabIndex        =   99
      Top             =   8640
      Width           =   630
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Distribuído:"
      Height          =   195
      Left            =   4080
      TabIndex        =   98
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atendente:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   75
      Top             =   0
      Width           =   780
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2640
      TabIndex        =   74
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "frmTroco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Public operacao As String
Dim sBloqueado_cai As String
'B - buffer, C - concomitante
Dim che As cheque
Dim pagcartao As pagamentocartao
Option Explicit
Dim sTab As String
Dim sNome As String
Dim sCod As String
Dim sCPF As String
Dim sStatus As String
Dim sLimite As String
Dim sTipo As String
Dim cTotalParcelas As Currency
Private Sub cmndbtnAv_Click()

End Sub


Private Sub cmdDep_Click()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente("")
    If c.cod_con <> "" Then
        lblCCoDep.caption = c.nome_con
        txtDeposito.tag = c.cod_con
        txtDescricaoPIX.text = c.nome_con
    End If
End Sub
Function restante() As Currency
    restante = CCur(lblRestante.caption)
End Function
Private Sub cmdFechar_Click()
    
    Dim f As FormasPag.FormasPag
     f = FormasPag.F_BuscaFormasPag(txtbxFormaPag.tag)
    
     If f.avista_For = "N" And lblStatus_CLi.tag <> "1" Then
            
                MsgBox "Essa venda não pode ser concluída para esse cliente nessa forma de pagamento!", vbExclamation
                Exit Sub
            
        End If
    
    If lblStatus_CLi.tag = "2" Then
            MsgBox "Não é possível a realização da venda a crédito pois o cliente se encontra Bloqueado!", vbExclamation
            Exit Sub
    End If
    If CCur(lblLimiteRest.caption) < CCur(txtbxValorDoc.text) Then
            MsgBox "Limite insuficiente para a operação!", vbExclamation
            P_Status ""
            Exit Sub
        End If
    If MsgBox("Confirma fechamento da venda?", vbYesNo + vbQuestion) = vbYes Then
        vendas.F_AlteraStatusVenda Me.tag, 2
        P_DeletaVendaAberta Me.tag
        ImprVenda Me.tag
        Me.Hide
        
    End If
    
    
    
End Sub

Private Sub cmdPix_Click()
    Dim c As ContaCorrente
    c = F_BuscaContaCorrente("")
    If c.cod_con <> "" Then
        lblCcoPix.caption = c.nome_con
        txtPIX.tag = c.cod_con
        txtDescricaoPIX.text = c.nome_con
    End If
End Sub

Private Sub cmndbtnAdd_Click()
    
    If chkCh.Value = 1 Then
        If txtbxNum_Che.text = "" Then
            MsgBox "Informe o numero do cheque!", vbExclamation
            Exit Sub
        End If
    End If
    
    If cmndbtnAdd.tag = "" Then
    Else
        msflxgrd.row = cmndbtnAdd.tag
        msflxgrd.col = 1:
        If chkCh.Value = 1 Then
            msflxgrd.text = "Cheque"
        Else
            msflxgrd.text = "Duplicata"
        End If
        msflxgrd.col = 2:
        msflxgrd.text = txtbxVencimento.text
        msflxgrd.col = 3:
        msflxgrd.text = txtbxValor_Che.text
        msflxgrd.col = 4:
        msflxgrd.text = txtbxNum_Che.text
        msflxgrd.col = 5:
        msflxgrd.text = txtbxAgencia.text
        msflxgrd.col = 6:
        msflxgrd.text = txtbxContaCorrente.text
        msflxgrd.col = 7:
        msflxgrd.text = txtbxBanco.tag
        msflxgrd.col = 8:
        msflxgrd.text = txtbxCPFChe.text
        msflxgrd.col = 9:
        msflxgrd.text = txtbxNome_Che.text
        cmndbtnAdd.tag = ""
        somaParcelas
    End If
End Sub

Private Sub cmndbtnBuscaBanco_Click()
    buscaBanco
End Sub

Private Sub cmndbtnCancelarDescGer_Click()
    If MsgBox("Deseja cancelar o desconto da venda?", vbYesNo + vbQuestion) = vbYes Then
        
        P_CancelaDescontoVenda Me.tag
        bActive = False
        Form_Activate
    End If
End Sub

Private Sub cmndbtnCob_Click()
        Dim t As tipocob
    
    t = Cobranca.F_BuscaTipoCob("")
    If t.cod_tip <> "" Then
        txtbxTipoCob.text = t.nome_tip
        txtbxTipoCob.tag = t.cod_tip
        
    End If

End Sub

Private Sub cmndbtnDadosCard_Click()
    If CCur(txtbxEntradaCart.text) > 0 Then
        buscaDadosCartao
    End If
End Sub

Private Sub cmndbtnDadosChv_Click()
    If CCur(txtbxEntradaChv.text) > 0 Then
        buscaDadosChv
    End If
End Sub
Sub buscaDadosChv()
    If txtbxNomeCli.tag = "" Then
        MsgBox "Para Compras com Cheques, é necessário que o cliente seja cadastrado!", vbExclamation
        Exit Sub
        
    End If
    
    If che.Venc_che = "" Then
        che.Venc_che = Date
    End If
    If che.cpf_che = "" Then
        che.cpf_che = txtbxCNPJ.text
    End If
    If che.nome_che = "" Then
        che.nome_che = txtbxNomeCli.text
    End If
    If che.Valor_che = 0 Then
        che.Valor_che = CCur(txtbxEntradaChv.text)
    End If
    
    che = F_DadosCheque(True, che.Venc_che, che.nome_che, che.num_che, che.cco_che, che.age_che, che.cod_bco, CStr(che.Valor_che), che.cpf_che)
    
    somaParcelas
    
End Sub
Private Sub cmndbtnEdi_Click()
    With msflxgrd
        If .row < 1 Then
            MsgBox "Escolha uma parcela para editar!", vbExclamation
            Exit Sub
            Exit Sub
            
        End If
                
        cmndbtnAdd.tag = msflxgrd.row
        .col = 1:
        If Left(.text, 1) = "C" Then
            chkCh.Value = 1
        Else
            chkCh.Value = 0
        End If
        
        .col = 2:
        If .text <> "" Then
            txtbxVencimento.text = .text
        End If
        .col = 3:
        If .text <> "" Then
            txtbxValor_Che.text = .text
        End If
        .col = 4:
        If .text <> "" Then
            txtbxNum_Che.text = .text
        End If
        .col = 5:
        If .text <> "" Then
            txtbxAgencia.text = .text
        End If
        .col = 6:
        If .text <> "" Then
            txtbxContaCorrente.text = .text
        End If
        .col = 7:
        If .text <> "" Then
            txtbxBanco.tag = .text
            txtbxBanco.text = .text
        End If
        
        .col = 8:
        If .text <> "" Then
            txtbxCPFChe.text = .text
        End If
        .col = 9:
        If .text <> "" Then
            txtbxNome_Che.text = .text
        End If
        If chkCh.Value Then
            If txtbxCPFChe.text = "" Then
                txtbxCPFChe.text = txtbxCNPJ.text
                txtbxNome_Che.text = txtbxNomeCli.text
            End If
            txtbxNum_Che.SetFocus
        Else
            txtbxVencimento.SetFocus
        End If
    End With
End Sub



Private Sub msflxgrd_DblClick()
    cmndbtnEdi_Click
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdi_Click
    End If
End Sub

Private Sub msflxgrd_KeyPress1(KeyAscii As Integer)
    If msflxgrd.col > 1 Then
        
        
        txtbxMsflx.width = msflxgrd.CellWidth
        txtbxMsflx.height = msflxgrd.CellHeight
        txtbxMsflx.Top = msflxgrd.Top + msflxgrd.CellTop
        txtbxMsflx.Left = msflxgrd.Left + msflxgrd.CellLeft
        If KeyAscii <> 13 Then
            txtbxMsflx.text = Chr(KeyAscii)
        End If
        txtbxMsflx.Visible = True
        
        
        
        txtbxMsflx.SetFocus
        txtbxMsflx.SelStart = Len(txtbxMsflx.text)
        
        
    End If
End Sub


Private Sub txtbxAgencia_GotFocus()
    ganhaFoco txtbxAgencia
End Sub

Private Sub txtbxAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAgencia_LostFocus()
    AjustaTexto txtbxAgencia
End Sub

Private Sub txtbxBanco_GotFocus()
    ganhaFoco txtbxBanco
End Sub

Private Sub txtbxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaBanco_Click
    End If
End Sub

Private Sub txtbxBanco_LostFocus()
    AjustaTexto txtbxBanco
End Sub

Private Sub txtbxContaCorrente_GotFocus()
    ganhaFoco txtbxContaCorrente
End Sub

Private Sub txtbxContaCorrente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxContaCorrente_LostFocus()
    AjustaTexto txtbxContaCorrente
End Sub

Private Sub txtbxDataBase_GotFocus()
    ganhaFoco txtbxDataBase
End Sub

Private Sub txtbxDataBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataBase_LostFocus()
    P_AjustaData txtbxDataBase
End Sub

Private Sub cmndbtnCancel_Click()
    
    P_DeletaVendaAberta Me.tag
    Me.tag = ""
    Me.Hide
End Sub

Private Sub cmndbtnDesconto_Click()
    Dim sUsu As String
    Dim sDes As String
    If txtbxDesconto.text = 0 Then
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
    
    sDes = F_BuscaUmCampo("desconto_usu", "" & TAB_DEFUSU & "", "cod_usu='" & sUsu & "'")
    If sDes = "" Then
        sDes = "0"
    End If
    If CCur(sDes) < CCur(txtbxDescPerc.text) Then
        MsgBox "O Desconto possível para este usuário é de " & sDes & "%", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Confirma desconto de " & txtbxDesconto.text & "?", vbYesNo + vbCritical) = vbYes Then
        vendas.P_RecalculaVendaDesconto lblFinalPrev.caption, txtbxDescPerc.text, Me.tag, sUsu
        Carrega
    End If
    
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    Dim sVisivel As String
    Dim sSEQ_DOC As String
    On Error GoTo trata_Erro
    Dim rest As Currency
    Dim iParc As Integer
    Dim iCount As Integer
    Dim sValor As String
    Dim sFCard As String
    Dim sEspCard As String
    Dim sChe As String
    Dim sTip As String
    Dim sAge As String
    Dim sConta As String
    Dim sVen As String
    Dim sVal As String
    Dim sBanco As String
    Dim sNome As String
    Dim sCPF As String
    Dim i As Integer
    Dim cTotalDup As Currency
    Dim cTotalChp As Currency
    Dim cLancado As Currency
    Dim f As FormasPag.FormasPag
    Dim totalLancado As Currency
    
        
    
    
    If CCur(lblTotalDistribuido.caption) <> CCur(lblValor.caption) Then
        MsgBox "Os valores não estão completamente distribuídos!", vbExclamation
        Exit Sub
    End If
    
    If CCur(txtDeposito.text) > 0 And txtDeposito.tag = "" Then
        MsgBox "Informe a conta de deposito", vbExclamation
        Exit Sub
    End If
    If CCur(txtPIX.text) > 0 And txtPIX.tag = "" Then
        MsgBox "Informe a conta de PIX", vbExclamation
        Exit Sub
    End If
    f = FormasPag.F_BuscaFormasPag(txtbxFormaPag.tag)
    
     If f.avista_For = "N" And lblStatus_CLi.tag <> "1" Then
            
                MsgBox "Essa venda não pode ser concluída para esse cliente nessa forma de pagamento!", vbExclamation
                Exit Sub
            
        End If
    
    
    If strEfetivarVenda_Usu <> "S" Then
        MsgBox "O usuário não tem permissão para efetivar venda!", vbExclamation
        Exit Sub
    End If
    
    If Not bUsaCaixaSeparado Then
        If strCod_Cai = "" Then
            strCod_Cai = "01"
        End If
    End If
    
    
    If Not F_IsCaixa And bUsaCaixaSeparado Then
         
        MsgBox "Este computador não é um caixa, portanto,não pode efetivar venda!", vbExclamation
            Exit Sub
        End If
            
        strSQL = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
        If strSQL = "F" Then
            MsgBox "Não é possível efetuar o recebimento pois este caixa está fechado!", vbExclamation
            Exit Sub
        End If
    
    txtbxDinheiro_LostFocus
    
    If lblTipoCai.caption = "DUP" Or lblTipoCai.caption = "CAR" Or lblTipoCai.caption = "CHP" Then
        somaParcelas
        If cTotalParcelas <= 0 Then
            MsgBox "Os valores distribuídos não estão corretos para o tipo DUP, CAR ou CHP!", vbExclamation
            Exit Sub
        End If
        
    Else
    
    End If
    If CCur(lblRestante.caption) > 0 Then
        MsgBox "Não há valor distribuído!", vbExclamation
        
        cmndbtnFormasPag_Click
        Exit Sub
    End If
    If CCur(lblTroco.caption) < 0 Then
        MsgBox "Informe o Valor Pago!", vbExclamation
        'txtbxDinheiro.Text = lblValor.caption
        'txtbxDinheiro.SetFocus
        Exit Sub
       
    End If
    cLancado = CCur(txtbxDinheiro.text) + CCur(txtbxEntradaDin.text) + CCur(txtbxEntradaCart.text) + CCur(txtbxEntradaChv.text)
    If cLancado < CCur(lblValor.caption) And lblTipoCai.caption = "DIN" Then
        MsgBox "O valor não está lançado corretamente!", vbExclamation
        Exit Sub
    End If
    If optnAgora.Value Then
        If Not F_confMsflxCH(msflxgrd, 2, 4, 5, 6, 7, 8, 9, 1) Then
            Exit Sub
        End If
    
        
    End If
    
    If lblRestante.caption <> 0 Then
        MsgBox "Está restando " & lblRestante.caption & " para lançar!", vbExclamation
        Exit Sub
    End If
    If pc.BackColor = vbGreen Then
        sVisivel = "S"
    Else
        sVisivel = "S"
    End If
        
    If lblTipoCai.caption = "DUP" Or lblTipoCai.caption = "CHP" Or lblTipoCai.caption = "CHV" Then
        
        If CCur(lblLimiteRest.caption) <= CCur(txtbxValorDoc.text) Then
            MsgBox "Limite insuficiente para a operação!", vbExclamation
            P_Status ""
            Exit Sub
        End If
        If CCur(lblTotalDistribuido.caption) <> CCur(txtbxValorDoc) + CCur(lblTotalEntrada) Then
            MsgBox "O Valor a prazo está diferente do distribuído!", vbExclamation
            Exit Sub
        End If
        If txtbxNomeCli.tag = "" Or txtbxNomeCli.tag = "00000" Or txtbxNomeCli.tag = "0" Then
            MsgBox "Esta opção de pagamento requer que o cliente tenha sido cadastrado!", vbExclamation
            txtbxNomeCli.SetFocus
            Exit Sub
        End If
        If lblStatus_CLi.tag = "2" Then
            MsgBox "Não é possível a realização da venda a crédito pois o cliente se encontra Bloqueado!", vbExclamation
            Exit Sub
        End If
        If lblStatus_CLi.tag = "3" Then
            MsgBox "Não é possível a realização da venda a crédito pois o cliente só é liberado para compra com Cheques", vbExclamation
            Exit Sub
        End If
    End If
    If lblTipoCai.caption = "DUP" Then
        If txtbxTipoCob.tag = "" Then
            MsgBox "Escolha o tipo de cobrança!", vbExclamation
            txtbxTipoCob.SetFocus
            Exit Sub
        End If
    End If
    If F_Vazio(txtbxNomeCli.text) Then
        MsgBox "Informe o nome do cliente!", vbExclamation
        txtbxNomeCli.SetFocus
        Exit Sub
    End If
    If txtbxNomeCli.tag = "" Then
        MsgBox "Escolha o Cliente", vbExclamation
        txtbxNomeCli.SetFocus
        Exit Sub
    End If
    
    
    If lblTipoCai.caption <> "DIN" Then
        'txtbxDinheiro.Text = lblValor.caption
    End If
    
    If CCur(txtbxEntradaCart.text) > 0 Then
        If pagcartao.cod_fin = "" Then
            MsgBox "Informe Os dados do cartão", vbExclamation
            buscaDadosCartao
            Exit Sub
        End If
        
    End If
    If CCur(txtbxEntradaChv.text) > 0 Then
        If che.num_che = "" Or che.Valor_che <> CCur(txtbxEntradaChv.text) Then
            MsgBox "Informe Os dados do cheque", vbExclamation
            buscaDadosChv
            Exit Sub
        End If
        
    End If
    P_Status "Finalizando Venda..."
    
    If CCur(txtbxEntradaCart.text) > 0 Then
        If pagcartao.tipo_fin = "3" Then
            
            If bUsaTEF = True Then
                If Not TefNaoFiscalSemECF(txtbxEntradaCart.text, "", "", "", "", pagcartao.parcelas) Then
                    P_Status ""
                    MsgBox "Erro no TEF", vbExclamation
                    Exit Sub
                End If
            End If
            
              
        End If
        
        If pagcartao.operacao = "C" Then
            sFCard = "09"
        Else
            sFCard = "10"
        End If
        If pagcartao.operacao = "C" Then
            sEspCard = "CRT"
        Else
            sEspCard = "CRD"
        End If
        If Not F_AdicionaVenFec(Me.tag, txtbxEntradaCart.text, sFCard, _
            strCod_Cai, strCod_Tur, sEspCard) Then
            MsgBox "Erro na gravação de VENFEC o processo irá continuar", vbExclamation
            
        End If
        
        For iParc = 1 To pagcartao.parcelas
            If NumeroCupomTef = "" Or bUsaTEF = False Then
                NumeroCupomTef = pagcartao.autorizacao
            End If
            If Not F_GravaVenFin(Me.tag, pagcartao.cod_fin, CCur(txtbxEntradaCart.text) / pagcartao.parcelas, Date, NumeroCupomTef, pagcartao.parcelas, "V", CStr(iParc), DateAdd("D", IIf(pagcartao.operacao = "C", 30 * iParc, 1), Date)) Then
                MsgBox "Erro na gravação da parcela da financeira"
            End If
        Next
        
        strSQL = F_GravaCai("C", Date, txtbxEntradaCart.text, "PAG. CARTÃO " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, sEspCard, "", txtbxEntradaCart.text, txtbxEntradaCart.text, strCod_Cai, strCod_Tur, "000", sVisivel)
            
        If strSQL = "" Then
            MsgBox "Houve um erro no lançamento do caixa. Avise o responsável!", vbExclamation
            
        End If
            
    End If
            
    If CCur(txtbxEntradaDin.text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtbxEntradaDin.text, "PAG. A VISTA VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, "DIN", "", txtbxEntradaDin.text, 0, strCod_Cai, strCod_Tur, "000", sVisivel)
        
        If strSQL = "" Then
            MsgBox "Houve um erro no lançamento do caixa. Avise o responsável!", vbExclamation
            
        End If
        
        If Not F_AdicionaVenFec(Me.tag, txtbxEntradaDin.text, "01", _
            strCod_Cai, strCod_Tur, "DIN") Then
            MsgBox "Erro na gravação de VENFEC o processo irá continuar", vbExclamation
            
        End If
        
        
    End If
    If CCur(txtPIX.text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtPIX.text, "PAG. PIX VENDA: " & Me.tag & lblCcoPix.caption & "-" & txtbxNomeCli.text, _
            Me.tag, "PIX", "", txtPIX.text, txtPIX.text, strCod_Cai, strCod_Tur, "000", sVisivel)
        
        If strSQL = "" Then
            MsgBox "Houve um erro no lançamento do caixa. Avise o responsável!", vbExclamation
            
        End If
        
        strSQL = F_GravaFinanceiro("C", Date, txtPIX.text, "PIX VENDA: " & Me.tag & lblCcoPix.caption & "-" & txtbxNomeCli.text, _
             txtPIX.tag, "N", strPlaRecebDupl)
             
        
        
        If Not F_AdicionaVenFec(Me.tag, txtPIX.text, "01", _
            strCod_Cai, strCod_Tur, "PIX") Then
            MsgBox "Erro na gravação de VENFEC o processo irá continuar", vbExclamation
            
        End If
        
        
    End If
    If CCur(txtDeposito.text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtDeposito.text, "PAG. DEP VENDA: " & Me.tag & lblCCoDep.caption & "-" & txtbxNomeCli.text, _
            Me.tag, "DEP", "", txtDeposito.text, txtDeposito.text, strCod_Cai, strCod_Tur, "000", sVisivel)
        
        If strSQL = "" Then
            MsgBox "Houve um erro no lançamento do caixa. Avise o responsável!", vbExclamation
            
        End If
        
        strSQL = F_GravaFinanceiro("C", Date, txtPIX.text, "DEP VENDA: " & Me.tag & lblCcoPix.caption & "-" & txtbxNomeCli.text, _
             txtDeposito.tag, "N", strPlaRecebDupl)
        
        If Not F_AdicionaVenFec(Me.tag, txtDeposito, "01", _
            strCod_Cai, strCod_Tur, "DEP") Then
            MsgBox "Erro na gravação de VENFEC o processo irá continuar", vbExclamation
            
        End If
        
        
    End If
    
            
    If CCur(txtbxEntradaChv.text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtbxEntradaChv.text, "PAG. CHV VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, "CHV", "", txtbxEntradaChv.text, txtbxEntradaChv.text, strCod_Cai, strCod_Tur, "000", sVisivel)
        
        If strSQL = "" Then
            MsgBox "Houve um erro no lançamento do caixa. Avise o responsável!", vbExclamation
            
        End If
        
        If Not F_AdicionaVenFec(Me.tag, txtbxEntradaChv.text, "02", _
            strCod_Cai, strCod_Tur, "CHV") Then
            MsgBox "Erro na gravação de VENFEC o processo irá continuar", vbExclamation
            
        End If
        sSEQ_DOC = F_GravaDoc("VENDA CHV " & txtbxNomeCli.text, Me.tag, txtbxEntradaChv.text, _
            Date, txtbxNomeCli.tag, "R", Me.tag, 1, "CHV", sTipo, "000")
        strSQL = F_GravaCh(sSEQ_DOC, che.num_che, che.cod_bco, che.age_che, che.cco_che, che.nome_che, CStr(che.Valor_che), _
            che.Venc_che, txtbxNomeCli.tag, "R", Me.tag, "", "1", "", che.cpf_che)
            
        
    End If
            
            
            
    If CCur(txtbxValorDoc.text) > 0 Then  'A PRAZO
                
        
        If optnAgora.Value Then
        
        
        
            sSEQ_DOC = F_GravaDoc("VENDA " & Me.tag & " " & txtbxNomeCli.text, Me.tag, txtbxValorDoc.text, _
             Date, txtbxNomeCli.tag, "R", Me.tag, txtbxParcelas.text, "V", sTipo, "000")
                
                
                
            If sSEQ_DOC <> "" Then
                With msflxgrd
                    
                    For iCount = 1 To .rows - 1
                        .row = iCount
                        .col = 1: sTip = .text
                        .col = 2: sVen = .text
                        .col = 3: sVal = .text
                        If Left(sTip, 1) = "D" Then
                            strSQL = F_GravaItDoc(sSEQ_DOC, sVal, sVen, sVen, CStr(iCount), a_receber_avulso, "", txtbxTipoCob.tag)
                            If strSQL = "" Then
                                MsgBox " Erro na gravação da parcela " & iCount
                            End If
                            cTotalDup = cTotalDup + CCur(sVal)
                        Else
                            .col = 4: sChe = .text
                            .col = 5: sAge = .text
                            .col = 6: sConta = .text
                            .col = 7: sBanco = .text
                            .col = 8: sCPF = .text
                            .col = 9: sNome = .text
                            strSQL = F_GravaCh(sSEQ_DOC, sChe, sBanco, sAge, sConta, sNome, sVal, sVen, txtbxNomeCli.tag, _
                                    "R", Me.tag, "", CStr(iCount), "", sCPF)
                            cTotalChp = cTotalChp + CCur(sVal)
                                    
                            If strSQL = "" Then
                                MsgBox " Erro na gravação da parcela " & iCount
                            End If
                        End If
                            
                    Next
                End With
                
                
                If cTotalChp > 0 Then
                    strSQL = F_GravaCai("C", Date, CStr(cTotalChp), "PAG. A CHP VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
                    Me.tag, "CHP", "", CStr(cTotalChp), CStr(cTotalChp), strCod_Cai, strCod_Tur, "000", sVisivel)
                    If strSQL = "" Then
                        MsgBox "Houve um erro no lançamento do caixa. Avise o responsável!", vbExclamation
            
                    End If
                End If
                If cTotalDup > 0 Then
                    strSQL = F_GravaCai("C", Date, CStr(cTotalDup), "PAG. A DUP VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
                    Me.tag, "DUP", "", CStr(cTotalDup), CStr(cTotalDup), strCod_Cai, strCod_Tur, "000", sVisivel)
                    If strSQL = "" Then
                        MsgBox "Houve um erro no lançamento do caixa. Avise o responsável!", vbExclamation
            
                    End If
                End If
            Else
                MsgBox "Erro na gravação do documento!", vbExclamation
                Exit Sub
            End If
                    
        Else
            
        End If
        
        If Not F_AdicionaVenFec(Me.tag, txtbxValorDoc.text, txtbxFormaPag.tag, _
                strCod_Cai, strCod_Tur, lblTipoCai.caption) Then
                MsgBox "Erro na gravação de VENFEC o processo irá continuar", vbExclamation
        End If
                
                
    End If
                
                
            
            
            
    strSQL = "update " & TAB_VENDAS & " set nomecli_ven = '" & txtbxNomeCli.text & "' "
    strSQL = strSQL & ",valor_ven=" & F_GravaDec(CCur(lblValor.caption))
    strSQL = strSQL & ",desconto_ven=" & F_GravaDec(CCur(lblDesconto.caption))
    
    If optnAgora.Value Then
        strSQL = strSQL & ",situacao_ven='4', " 'efetivada
            
    Else
        strSQL = strSQL & ",situacao_ven='3', " 'a efetivar
    End If
    strSQL = strSQL & " DATaefet_ven=" & F_GravaData(Date) & ", "
    
    
    
    strSQL = strSQL & " fechavenda_usu='" & strCod_Usu & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "' "
        'strSQL = strSQL & "entrega_ven='" & IIf(chkEntrega.Value = 1, "S", "N") & "',"
        'strSQL = strSQL & "endereco_ven='" & txtbxEndereco.Text & "',"
        'strSQL = strSQL & "bairro_ven='" & txtbxBairro.Text & "',"
        'strSQL = strSQL & "fone_ven='" & txtbxFone.Text & "'"
    strSQL = strSQL & " where num_ven='" & Me.tag & "' "
    conn.Execute strSQL
    
    P_BaixaVenda Me.tag
        
      
    
    
    
    P_DeletaVendaAberta Me.tag
    
        
    strSQL = pacote.F_GravaPacote(Me.tag)
    If sSEQ_DOC <> "" And sTipo <> "U" And bUsaContrato = True Then
        imprimeContrato sSEQ_DOC
     End If
    If sSEQ_DOC <> "" And bUsaCarne Then
        imprimeCarne sSEQ_DOC
    End If
    If chkEmite.Value = 1 Then
        If bUsaVendaSupermercado Then
            ImprimeComprovanteNaoFiscal Printer, Me.tag, False
        Else
            ImprVenda Me.tag
        End If
    End If
    
    If txtbxNomeCli.tag = "0" Or bUsaVendaSupermercado = True Then
        If pc.BackColor = vbGreen Then '
            strSQL = GravaNFCeVenda(Me.tag, strTipo_IcmsNFe)
            
            strSQL = TransmitirNFcE(strSQL, "", True)
    
             If InStr(1, "CStat=100", strSQL) > 0 Then
                 AlteraStatusNfe Me.tag, "2"
                 vendas.AlteraStatusVen Me.tag, "2"
                 
             'MsgBox strSQL
             Else
            ' MsgBox strSQL
             End If
        Else
            'ImprimeComprovanteNaoFiscal Printer, Me.tag, True
        End If
    End If
    
    
    
    pagcartao.tipo_fin = ""
    
    P_Status ""
            
    If IsLoad("frmDadosCartao") Then
        Unload frmDadosCartao
        
    End If
    pagcartao.autorizacao = ""
    pagcartao.cod_fin = ""
    pagcartao.nome_fin = ""
    
    pagcartao.parcelas = "1"
    Me.tag = "X"
    
    Me.Hide
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
        
    If bActive = True Then Exit Sub
    che.Valor_che = 0
    che.num_che = ""
    che.nome_che = ""
    che.Venc_che = ""
    che.num_che = ""
    Carrega
    bActive = True
    
    If bUsaCaixaSeparado = True Then
        cmdFechar.Visible = False
    End If
    
End Sub


Sub somaParcelas()
    Dim iCount As Integer
    Dim cTotal As Currency
    Dim iCol As Integer
    Dim iRow As Integer
    cTotal = 0
    With msflxgrd
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 3
            cTotal = cTotal + CCur(.text)
        Next
        lblTotalDistribuido.caption = FormatVal(cTotal + CCur(lblTotalEntrada.caption))
        .row = iRow
        .col = iCol
    End With
    cTotalParcelas = cTotal
End Sub
Sub P_EscondeOpcoesPagamento()
    If lblTipoCai.tag = "02" Then
        pcAprazo.Visible = True
        lblInternv.Visible = True
        txtbxParcelas.Visible = True
        lblParcelas.Visible = True
        lblInterMes.Visible = True
        txtbxInterv.Visible = True
        txtbxDataBase.Visible = True
        lblDataBase.Visible = True
        pcEntrada.Visible = True
        pcDin.Visible = False
        pcTroco.Visible = False
        
            optnAgora.Enabled = True
            optnDepois.Enabled = True
            optnAgora.Value = True
        
        
        
    Else
        pcAprazo.Visible = False
        
        If lblTipoCai.caption = "DIN" Then
            txtbxParcelas.Visible = False
            lblParcelas.Visible = False
            
            
        Else
            txtbxParcelas.Visible = True
            lblParcelas.Visible = True
        End If
        optnAgora.Enabled = False
        optnDepois.Enabled = False
        optnAgora.Value = True
            
        pcDin.Visible = True
        pcTroco.Visible = True
        lblInterMes.Visible = False
        txtbxInterv.Visible = False
        lblInternv.Visible = False
        txtbxDataBase.Visible = False
        lblDataBase.Visible = False
        
        
        
    End If
    If lblTipoCai.caption = "CHP" Or lblTipoCai.caption = "CHV" Then
        
    End If
    pcEntrada.Visible = True
    P_TotalEntrada
End Sub
Sub Carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro
    
    strSQL = F_VerificaVendaAberta(Me.tag)
    If strSQL <> "" And strSQL <> strIP_Usu Then
        
        MsgBox "A Venda se encontra aberta no computador " & strSQL, vbExclamation
        Unload Me
    End If
    
    
    P_GravaVendaAberta Me.tag
    
    
    sTipo = F_BuscaUmCampo("tipocli_ven", TAB_VENDAS, "num_ven='" & Me.tag & "'")
    
    If sTipo = "U" Then
        sTab = TAB_FUNCIONARIOS
        sCod = "cod_fun"
        sNome = "nome_fun"
        sCPF = "cpf_fun"
        sLimite = "limite_fun"
        sStatus = "status_fun"
    Else
        sTab = TAB_CLIENTES
        sCod = "cod_cli"
        sNome = "nome_cli"
        sCPF = "cnpj_cli"
        sLimite = "limite_cli"
        sStatus = "status_cli"
        sTipo = "C" 'cliente
    End If
    
    strSQL = "select tc." & sLimite & " as limite_cli ,tc." & sCPF & " as cpf_cli,"
    strSQL = strSQL & " tt.nome_tip,tc." & sStatus & " as status_cli,tv.*,tf.*,tfu.guerra_fun from " & TAB_VENDAS & " as tv, "
    strSQL = strSQL & sTab & " as tc, tab_formaspag as tf,tab_funcionarios as tfu,tab_tipopag as tt "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for  and tc." & sCod & "=tv.cod_cli "
    strSQL = strSQL & " and tt.cod_tip=tf.cod_tip "
    strSQL = strSQL & " and num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tv.abrevenda_usu=tfu.cod_fun "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        txtbxNomeCli.text = !nomecli_ven
        
        txtbxNomeCli.tag = !cod_cli
        If !cod_cli = "0" Or !cod_cli = "" Or !cod_cli = "000000" Or !cod_cli = "00000" Then
            txtbxCNPJ.text = TiraNull(!cpfcli_ven)
        Else
            txtbxCNPJ.text = TiraNull(!cpfcli_ven)
        End If
        
        
        txtbxFormaPag.text = !nome_For
        lblTipoCai.caption = !tipo_cai
        
        txtbxDescForma.text = !desconto_for 'FormatVal(!desconto_ven / !valor_ven) * 100 '!desconto_for
        txtbxFormaPag.tag = !cod_For
        lblValor.caption = FormatVal(!valor_ven)
        lblValor.tag = FormatVal(!valor_ven - !desconto_ven)
        lblDesconto.caption = FormatVal(!desconto_ven)
        lblTotal.caption = FormatVal(!valor_ven)
        lblPExtenso.caption = Extenso(!valor_ven)
        lblCaixa.caption = strNome_Usu
        lblNome_ven.caption = !guerra_fun
        'txtbxDinheiro.Text = lblValor.caption
        lblTipoCai.caption = !tipo_cai
        lblTipoCai.tag = !cod_tip
        If !autorizadesconto_ven <> "" Then
            pcDescAutorizadoPor.Visible = True
            lblAutorizaDesconto.caption = F_BuscaUmCampo("nome_fun", TAB_FUNCIONARIOS, "cod_fun='" & !autorizadesconto_ven & "'")
            lblAutorizaDesconto.tag = !autorizadesconto_ven
        Else
            lblAutorizaDesconto.caption = ""
            lblAutorizaDesconto.tag = ""
            pcDescAutorizadoPor.Visible = False
        End If
        lblTotalQueima.caption = FormatVal(F_BuscaUmCampo("sum(pq_itd)", "tab_itvendas", "num_ven='" & Me.tag & "'"))
        lblDemaisProd.caption = FormatVal(CCur(lblValor.caption) - CCur(lblTotalQueima.caption))
        'txtbxEndereco.Text = !endereco_ven
        'txtbxBairro.Text = !bairro_ven
        'txtbxFone.Text = !fone_ven
        If !cod_cli <> 0 Then
            lblLimite.caption = FormatVal(!limite_cli)
            lblEmUso.caption = FormatVal(F_BuscaLimiteEmUsoCli(!cod_cli))
            lblLimiteRest.caption = FormatVal(!limite_cli - CCur(lblEmUso.caption))
            lblStatus_CLi.tag = Format(!status_cli)
            If !status_cli = "2" Then
                lblStatus_CLi.caption = "Bloqueado"
                lblStatus_CLi.ForeColor = vbRed
            ElseIf !status_cli = "3" Then
                lblStatus_CLi.caption = "Cheques"
                lblStatus_CLi.ForeColor = vbYellow
            Else
                lblStatus_CLi.caption = "Liberado"
                lblStatus_CLi.ForeColor = vbGreen
            End If
            
        End If
        If txtbxNomeCli.text = "" Then
                txtbxNomeCli.text = "Venda a Consumidor"
            End If
        mudaCorCliente
        cmndbtnFormasPag_Click
        
        If sTipo = "U" Then
            txtbxNomeCli.Locked = True
            cmndbtnCliente.Enabled = False
            optnDepois.Visible = False
        End If
        
        If lblTipoCai.caption = "CRT" Or lblTipoCai.caption = "CRD" Then
            txtbxEntradaCart.text = lblValor.caption
            txtbxEntradaCart.SetFocus
        ElseIf lblTipoCai.caption = "CHV" Then
            txtbxEntradaChv.text = lblValor.caption
            txtbxEntradaChv.SetFocus
        End If
        
        .Close
    End With
    Exit Sub
trata_Erro:
    
End Sub
Sub ajustaGrid()
    With msflxgrd
        .rows = 1
        .Cols = 10
        .col = 0: .text = "..."
        .col = 1: .text = "Tipo"
        .col = 2: .text = "Vencimento"
        .col = 3: .text = "Valor"
        .col = 4: .text = "Cheque"
        .col = 5: .text = "Agência"
        .col = 6: .text = "Conta"
        .col = 7: .text = "banco"
        .col = 8: .text = "CPF"
        .col = 9: .text = "Titular"
        .ColWidth(0) = 300
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(7) = 2000
        .ColWidth(8) = 2000
        .ColWidth(9) = 3000
        
    End With
End Sub
Sub P_TotalEntrada()
    'On Error Resume Next
    Dim cTotalLancado As Currency
    lblTotalEntrada.caption = FormatVal(CCur(txtbxEntradaCart.text) + CCur(txtbxEntradaDin.text) + CCur(txtbxEntradaChv.text)) + CCur(txtPIX.text) + CCur(txtDeposito.text)
    If txtbxValorDoc.text > 0 Then
        txtbxValorDoc.text = CCur(lblValor.caption) - CCur(lblTotalEntrada.caption)
    End If
    cTotalLancado = CCur(txtbxValorDoc.text) + CCur(txtbxEntradaDin.text) + CCur(txtbxEntradaCart.text) + CCur(txtbxEntradaChv.text) + CCur(txtDeposito.text) + CCur(txtPIX.text)
       
       lblRestante.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxValorDoc.text) - CCur(lblTotalEntrada.caption))
        lblTotalDistribuido.caption = FormatVal(cTotalLancado)
        
        
        
        
End Sub
Sub distribuiparcelas()
    msflxgrd.rows = 1
    Dim strSQL As String
    Dim iCount As Integer
    Dim cTotal As Currency
    Dim sTipo As String
    Dim cValor As Currency
    Dim sVlrParcela As String
    
    If CCur(txtbxValorDoc.text) = 0 Then
        msflxgrd.rows = 1
        Exit Sub
    End If
    
    
    cValor = CCur(lblValor.caption) - CCur(lblTotalEntrada.caption)
    
    If txtbxDataBase.text = "" Then
        txtbxDataBase.text = Format(Date, "dd/mm/yyyy")
    End If
    If chkCh.Value = 1 Or (lblTipoCai.caption = "CHP" Or lblTipoCai.caption = "CHV") Then
        sTipo = "Cheque"
    Else
        sTipo = "Duplicata"
    End If
    For iCount = 1 To Int(txtbxParcelas.text)
        strSQL = iCount & Chr(9) & sTipo & Chr(9)
        strSQL = strSQL & DateAdd("M", iCount, txtbxDataBase.text) & Chr(9)
        
        
        If iCount = Int(txtbxParcelas) Then
            sVlrParcela = FormatVal(CCur(cValor) - cTotal)
        Else
            sVlrParcela = FormatVal(CCur(cValor) / Int(txtbxParcelas.text))
            cTotal = cTotal + CCur(sVlrParcela)
        End If
        strSQL = strSQL & sVlrParcela & Chr(9)
        msflxgrd.AddItem strSQL
            
    Next
    
    somaParcelas
    
End Sub
Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    Dim strSQL As String
    If txtbxFormaPag.tag <> "" Then
        f = F_SelectFormasPag("", "", txtbxFormaPag.tag)
    Else
        f = F_SelectFormasPag("")
    End If
    
    If f.cod_For <> "" Then
        txtbxValorDoc.text = "0"
        If CCur(txtbxDescForma.text) > f.desconto_for Then
            
                strSQL = "Você optou por mudar a forma de pagamento de " & txtbxFormaPag.text
                strSQL = strSQL & " para " & f.nome_For & Chr(13)
                strSQL = strSQL & "Continuando com esta operação, todo o desconto da venda será cancelado!"
                strSQL = strSQL & "Deseja continuar?"
                If MsgBox(strSQL, vbOKCancel, "CANCELAMENTO DE DESCONTO") = vbCancel Then
                    Exit Sub
                Else
                    P_CancelaDescontoVenda Me.tag, CStr(f.desconto_for)
                    Carrega
                End If
            
        End If
        
        
        'txtbxEntradaCart.Text = "0,00"
        'txtbxEntradaChv.Text = "0,00"
        'txtbxEntradaDin.Text = "0,00"
        
        txtbxFormaPag.text = f.nome_For
        txtbxFormaPag.tag = f.cod_For
        txtbxDescForma.text = f.desconto_for
        lblTipoCai.caption = f.tipo_cai
        lblTipoCai.tag = f.cod_tip
        txtbxParcelas.text = f.parcelas_for
        If f.avista_For = "N" Then
            txtbxValorDoc.text = lblValor.caption
        End If
        distribuiparcelas
        vendas.P_AtualizaFormaPag Me.tag, txtbxFormaPag.tag
        
    
        
         
        If lblTipoCai.caption <> "DIN" Then
            'txtbxDinheiro.Text = lblValor.caption
            'txtbxDinheiro_LostFocus
        End If
        
        P_EscondeOpcoesPagamento
        If f.cod_For = "01" Then
            'avista
            pcTroco.Visible = True
        Else
            pcTroco.Visible = False
        End If
                
        If txtbxDinheiro.Enabled = True And txtbxDinheiro.Visible = True And f.tipo_cai = "DIN" Then
            txtbxDinheiro.SetFocus
        End If
        If f.tipo_cai = "CRT" Or f.tipo_cai = "CRD" Then
            txtbxEntradaCart.text = lblValor.caption
            txtbxEntradaCart.SetFocus
            Exit Sub
        End If
        If f.tipo_cai = "PIX" Then
            txtPIX.text = lblValor.caption
            txtPIX.SetFocus
        End If
        If f.tipo_cai = "DEP" Then
            txtDeposito.text = lblValor.caption
            txtDeposito.SetFocus
        End If
    End If
    
    
    
    
End Sub
Private Sub cmndbtnCliente_Click()
    Dim c As Clientes.cliente
    
    c = Clientes.F_BuscaCliente(txtbxNomeCli.text)
    If c.cod_cli <> "" Then
        txtbxNomeCli.text = c.nome_Cli
        txtbxNomeCli.tag = c.cod_cli
        If c.tipo_cli = "J" Then
            txtbxCNPJ.text = c.cnpj_Cli
            
        Else
            txtbxCNPJ.text = c.cpf_cli
        End If
        
        lblLimite.caption = FormatVal(c.limite_cli)
        lblEmUso.caption = FormatVal(F_BuscaLimiteEmUsoCli(c.cod_cli))
        lblLimiteRest.caption = FormatVal(c.limite_cli - CCur(lblEmUso.caption))
        lblStatus_CLi.tag = c.status_cli
        If c.status_cli = "1" Then
            lblStatus_CLi.caption = "Liberado"
            lblStatus_CLi.ForeColor = vbGreen
        ElseIf c.status_cli = "3" Then
            lblStatus_CLi.caption = "Cheques"
            lblStatus_CLi.ForeColor = vbYellow
        Else
            lblStatus_CLi.caption = "Bloqueado"
            lblStatus_CLi.ForeColor = vbRed
        End If
        If c.status_cli = "2" Then
            MsgBox "O Cliente encontra-se inativo e baixado, portanto, não é possível realizar a venda!", vbExclamation
            Exit Sub
            
        End If
        If c.status_cli = "3" Then
            MsgBox "O Cliente encontra-se com crédito bloqueado, portanto, só é possível venda a vista!", vbExclamation
            
            
        End If
        
        vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text
        mudaCorCliente
        txtEndereco.SetFocus
    End If
    
    
End Sub
Sub mudaCorCliente()
    If txtbxNomeCli.tag = "" Or txtbxNomeCli.tag = "00000" Then
        txtbxNomeCli.BackColor = vbRed
        lblCliCaption.caption = "&Cliente: (Não Cadastrado)(F2)"
    Else
        txtbxNomeCli.BackColor = vbGreen
        lblCliCaption.caption = "&Cliente: (F2)"
    End If
    
    
End Sub
Private Sub txtbxCPFChe_GotFocus()
    ganhaFoco txtbxCPFChe
End Sub

Private Sub txtbxCPFChe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxCPFChe_LostFocus
        
    End If
End Sub
Sub buscaCH()
    Dim c As cheque
    On Error Resume Next
    If txtbxCPFChe.text = "" Then Exit Sub
    c = F_HistCheCPF(txtbxCPFChe.text)
    
    If c.cod_bco <> "" Then
        txtbxNome_Che.text = c.nome_che
        txtbxAgencia.text = c.age_che
        txtbxContaCorrente.text = c.cco_che
        txtbxBanco.text = F_BuscaUmCampo("nome_bco", "tab_Bancos", "cod_bco='" & c.cod_bco & "'")
        txtbxBanco.tag = c.cod_bco
        
        txtbxNum_Che.SetFocus
    Else
        txtbxNome_Che.SetFocus
    End If
    
    
    
End Sub
Private Sub txtbxCPFChe_LostFocus()
    
    
    
    txtbxCPFChe.text = F_SoNumero(txtbxCPFChe.text)
    If Len(txtbxCPFChe.text) = 0 Then
        
    ElseIf Len(txtbxCPFChe.text) >= 14 Then
        'lblCpf.caption = "&CNPJ:"
        P_FormataCGC txtbxCPFChe
        If Not F_CGC(txtbxCPFChe.text) Then
            MsgBox "CNPJ Inválido!", vbExclamation
            txtbxCPFChe.SetFocus
            Exit Sub
        End If
    Else
        'lblCpf.caption = "&CPF:"
        P_FormataCPF txtbxCPFChe
        If Not F_CPF(txtbxCPFChe.text) Then
            MsgBox "CPF Inválido!", vbExclamation
            txtbxCPFChe.SetFocus
            Exit Sub
        End If
    End If
    buscaCH
End Sub

Private Sub txtbxDesconto_GotFocus()
    ganhaFoco txtbxDesconto
End Sub

Private Sub txtbxDescPerc_GotFocus()
    ganhaFoco txtbxDescPerc
End Sub

Private Sub txtbxDescPerc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescPerc_LostFocus()
    PerdeFoco txtbxDescPerc, "###,##0.0000"
    txtbxDesconto.text = FormatVal(CCur(txtbxDescPerc.text) / 100 * CCur(lblDemaisProd.caption))
    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.text))
    
End Sub

Private Sub txtbxDinheiro_Change()
    txtbxEntradaDin.text = txtbxDinheiro.text
End Sub

Private Sub txtbxEntradaCart_GotFocus()
    txtbxEntradaCart.text = "0"
    txtbxEntradaCart.text = restante
    ganhaFoco txtbxEntradaCart
End Sub

Private Sub txtbxEntradaCart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntradaCart_LostFocus()
    PerdeFoco txtbxEntradaCart
    If CCur(txtbxEntradaCart.text) < 0 Then
        txtbxEntradaCart.text = "0"
    End If
    P_TotalEntrada
    If txtbxEntradaCart.text > 0 Then
        buscaDadosCartao
    End If
End Sub
Sub buscaDadosCartao()
    If CCur(txtbxEntradaCart.text) > 0 Then
        pagcartao = F_DadosCartao(pagcartao.cod_fin, pagcartao.nome_fin, pagcartao.parcelas, _
        pagcartao.autorizacao, txtbxEntradaCart.text)
        'pcDadosCartao.Visible = True
        If pagcartao.cod_fin = "" Then
                pagcartao.valor = "0"
                txtbxEntradaCart.text = "0"
            MsgBox "Não foi escolhida uma financeira!", vbExclamation
            
        Else
            txtbxEntradaCart.text = pagcartao.valor
            P_TotalEntrada
        End If
    Else
        pagcartao.cod_fin = ""
        pagcartao.nome_fin = ""
        pagcartao.valor = "0"
        'pcDadosCartao.Visible = False
    End If
    
    
End Sub

Private Sub txtbxEntradaChv_GotFocus()
    ganhaFoco txtbxEntradaChv
End Sub

Private Sub txtbxEntradaChv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntradaChv_LostFocus()
    PerdeFoco txtbxEntradaChv
    P_TotalEntrada
    cmndbtnDadosChv_Click
    If msflxgrd.rows > 1 Then
    If MsgBox("Deseja recalcular as parcelas já lançadas?", vbYesNo + vbQuestion) = vbYes Then
        cmndbtnFormasPag_Click
    End If
    End If
End Sub

Private Sub txtbxEntradaDin_GotFocus()
    
    txtbxEntradaDin.text = "0"
    txtbxEntradaDin.text = restante
    ganhaFoco txtbxEntradaDin
End Sub

Private Sub txtbxEntradaDin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntradaDin_LostFocus()
    If CCur(txtbxDinheiro.text) > CCur(lblValor.caption) Then
        txtbxEntradaDin.text = lblValor.caption
    
    End If
    'txtbxEntradaDin.Text = restante
    PerdeFoco txtbxEntradaDin
    If CCur(txtbxEntradaDin.text) < 0 Then
        txtbxEntradaDin.text = "0"
    End If
    
    P_TotalEntrada
    
    'cmndbtnFormasPag_Click
End Sub

Private Sub txtbxFormaPag_Change()
    txtbxFormaPag.tag = ""
End Sub

Private Sub txtbxFormaPag_GotFocus()
    P_SelecTexto txtbxFormaPag
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxMsflx_GotFocus()
'    ganhafoco txtbxMsflx
End Sub

Private Sub txtbxMsflx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMsflx_LostFocus()
    If msflxgrd.col = 2 Then
        PerdeFoco txtbxMsflx
    Else
    End If
    msflxgrd.text = txtbxMsflx.text
    If msflxgrd.col <> msflxgrd.Cols - 1 Then
    
        msflxgrd.col = msflxgrd.col + 1
        msflxgrd.SetFocus
    End If
End Sub

Private Sub txtbxNome_Che_GotFocus()
    ganhaFoco txtbxNome_Che
End Sub

Private Sub txtbxNome_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Che_LostFocus()
    AjustaTexto txtbxNome_Che
End Sub

Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, "", txtbxCNPJ.text
End Sub

Private Sub txtbxNum_Che_GotFocus()
    ganhaFoco txtbxNum_Che
End Sub

Private Sub txtbxNum_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_Che_LostFocus()
    AjustaTexto txtbxNum_Che
End Sub

Private Sub txtbxValor_Che_GotFocus()
    ganhaFoco txtbxValor_Che
End Sub

Private Sub txtbxValor_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Che_LostFocus()
    PerdeFoco txtbxValor_Che
End Sub

Private Sub txtbxVencimento_GotFocus()
    ganhaFoco txtbxVencimento
End Sub

Private Sub txtbxVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxVencimento_LostFocus()
    P_AjustaData txtbxVencimento
    If txtbxVencimento.text <> "" Then
        If DateDiff("D", Date, txtbxVencimento.text) < 30 Then
            If MsgBox("Confirma lançamento da data retroativa para " & txtbxVencimento.text & "?", vbYesNo + vbCritical) = vbNo Then
                txtbxDataBase.SetFocus
                
            End If
        End If
    End If
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If operacao <> "C" Then
    If KeyCode = vbKeyF2 Then
        txtbxNomeCli.SetFocus
        
    ElseIf KeyCode = vbKeyF3 Then
        
        txtbxFormaPag.text = ""
        cmndbtnFormasPag_Click
    ElseIf KeyCode = vbKeyF7 Then
        
        'pc.BackColor = vbGreen
        'txtbxDinheiro.Text = lblValor.caption
        'cmndbtnOk.SetFocus
        
        chkEmite.Value = 1
        cmndbtnOk_Click
    ElseIf KeyCode = vbKeyF9 Then
        If sBloqueado_cai = "S" Then
            Exit Sub
        End If
        pc.BackColor = vbRed
        'txtbxDinheiro.Text = lblValor.caption
        cmndbtnOk.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
        pc.BackColor = vbBlue
        cmndbtnOk.SetFocus
    ElseIf KeyCode = vbKeyF10 Then
        
    End If
    
        
    End If
    
    
    If KeyCode = vbKeyF6 Then
        If bUsaCaixaSeparado = False And bUsaVendaSupermercado = False Then
            cmdFechar_Click
        End If
        
        
        
    End If
    If KeyCode = vbKeyF8 Then
        chkEmite.Value = 0
        pc.BackColor = vbGreen
        cmndbtnOk_Click
    End If
    If KeyCode = vbKeyF12 Then
        cmndbtnCancel_Click
    ElseIf KeyCode = vbKeyE And Shift = vbAltMask Then
        
    
    End If
    
End Sub

Private Sub Form_Load()
    resizeForm Me
    txtbxDataBase.text = Date
    bActive = False
    If bUsaCaixaSeparado = True Then
        'pc.BackColor = vbGreen
        'pc.ForeColor = vbGreen
        pc.BackColor = vbRed
        pc.ForeColor = vbRed
    Else
        pc.BackColor = vbRed
        pc.ForeColor = vbRed
    End If
    ajustaGrid
End Sub

Private Sub optnAv_Click()
'    frNome_cli.Visible = False
    
End Sub

Private Sub optnAv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub optnCH_Click()
'    frNome_cli.Visible = True
    'txtbxDinheiro.Text = lblValor.caption
    txtbxDinheiro_LostFocus
End Sub

Private Sub optnCH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub pcCli_GotFocus()
    'txtbxNome_Cli.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbxCNPJ_GotFocus()
    ganhaFoco txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    txtbxCNPJ.text = F_SoNumero(txtbxCNPJ.text)
    If Len(txtbxCNPJ.text) > 11 Then
        
        txtbxCNPJ.text = Format(txtbxCNPJ.text, "0#\.###\.###\/####\-##")
        If Not F_CGC(txtbxCNPJ.text) Then
            MsgBox "CNPJ inválido!", vbExclamation
            txtbxCNPJ.SetFocus
        End If
    ElseIf Len(txtbxCNPJ.text) > 1 Then
        txtbxCNPJ.text = Format(txtbxCNPJ.text, "00#\.###\.###\-##")
        
        If Not F_CPF(txtbxCNPJ.text) Then
            MsgBox "CPF inválido!", vbExclamation
            txtbxCNPJ.SetFocus
        End If
    Else
        txtbxCNPJ.text = ""
    End If
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, "", txtbxCNPJ.text
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
   PerdeFoco txtbxDesconto
   txtbxDescPerc.text = FormatVal(CCur(txtbxDesconto.text) * 100 / CCur(lblDemaisProd.caption), "###,##0.0000")

    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.text))
   
   
End Sub

Private Sub txtbxDinheiro_GotFocus()
    ganhaFoco txtbxDinheiro
End Sub

Private Sub txtbxDinheiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
        
    End If
End Sub

Private Sub txtbxDinheiro_LostFocus()
    
    PerdeFoco txtbxDinheiro
    If CCur(txtbxDinheiro.text) >= CCur(lblValor.caption) Then
        txtbxEntradaDin.text = lblValor.caption
        P_TotalEntrada
'        txtbxEntradaDin_LostFocus
    End If
    
    lblTroco.caption = FormatVal(CCur(txtbxDinheiro.text) - CCur(lblValor.caption))
    If CCur(lblTroco.caption) < 0 Then
        lblTroco.caption = "0,00"
    End If
    
End Sub




Private Sub txtbxNomeCli_GotFocus()
    ganhaFoco txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub

Private Sub txtbxParcelas_GotFocus()
    ganhaFoco txtbxParcelas
End Sub

Private Sub txtbxParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxParcelas_LostFocus()
    AjustaTexto txtbxParcelas
    distribuiparcelas
End Sub
Sub buscaBanco()
    Dim b As TBanco
    b = F_BuscaBanco(txtbxBanco.text)
    If b.cod_bco <> "" Then
        txtbxBanco.text = b.nome_bco
        txtbxBanco.tag = b.cod_bco
        txtbxNum_Che.SetFocus
    End If
    
End Sub

Sub buscaTotalProd()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
End Sub
Sub imprimeCarne(seq_doc As String)
    P_ImprimeCarne seq_doc, frmVisualiza.Pic
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        P_ImprimeCarne seq_doc, Printer
        Printer.EndDoc
        
    End If
    Unload frmVisualiza
    
    
End Sub

Sub ImprVenda(num_ven As String)
    
    imprimeVenda frmVisualiza.Pic, num_ven
    frmVisualiza.Show 1
    If frmVisualiza.tag <> "" Then
        imprimeVenda Printer, num_ven
        Printer.EndDoc
        
        If MsgBox("Deseja imprimir outra via?", vbYesNo + vbQuestion) = vbYes Then
            Set Printer = Printers(ImpressoraCopiaVenda.Indice)
            imprimeVenda Printer, num_ven
            Printer.EndDoc
            
            
            
        End If

    End If
            Unload frmVisualiza
End Sub

Function P_ImprimeCarne(seq_doc As String, obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoTer As String
    Dim yIni As Single
    Dim xIni As Single
    Dim iImp As Integer
    Dim rEnd As New ADODB.Recordset
    Dim sNomeTer As String
    Dim sVia As String
    Dim iVia As Integer
    
'    On Error GoTo trata_erro
On Error Resume Next
    
    sVia = F_BuscaUmCampo("count(*)", " " & TAB_DOC & " impress", "seq_doc='" & seq_doc & "'")
    If sVia = "" Then
       sVia = "0"
    End If
    iVia = Int(sVia)
    sVia = Int(sVia) + 1 & "ª Via"
    If iVia > -1 Then
        
        obj.fontsize = 76
        obj.FontName = "arial"
        obj.ForeColor = &H8000000F
        'obj.Print sVia;
        'obj.CurrentY = 100
        P_CentraImpressao obj, sVia, True
        obj.ForeColor = vbBlack
    End If
    sTipoTer = F_BuscaUmCampo("tipo_ter", " " & TAB_DOC & " ", "seq_doc='" & seq_doc & "'")
    If sTipoTer = "" Then
        Exit Function
    End If
    strSQL = F_SelectDocsReceber(sTipoTer, , , , , , , , , seq_doc)
    
    rPesq.Open strSQL, conn
    
    With rPesq
        
        
        
        
        If .RecordCount = 0 Then
            'MsgBox "Documento " & seq_doc & " não encontrado!", vbExclamation
            .Close
            Exit Function
        End If
        
        If sTipoTer = "C" Then
            strSQL = "select tci.nome_cid,tipo_set,cep_cli,logradouro,numero,complemento,endereco_cli,uf_cid,nome_set "
            strSQL = strSQL & " from " & TAB_CLIENTES & " as tc," & TAB_CIDADES & " as tci, tab_setores as ts "
            strSQL = strSQL & " where tci.cod_cid=ts.cod_cid "
            strSQL = strSQL & " and tc.cod_set=ts.cod_set "
            strSQL = strSQL & " and cod_cli='" & !cod_ter & "' "
            rEnd.Open strSQL, conn_client
            
        Else
            strSQL = "select tci.nome_cid,tipo_set,cep_rep as cep_cli,endereco_rep as endereco_cli,uf_cid,nome_set "
            strSQL = strSQL & " from tab_representantes as tc," & TAB_CIDADES & " as tci, tab_setores as ts "
            strSQL = strSQL & " where tci.cod_cid=ts.cod_cid "
            strSQL = strSQL & " and tc.cod_set=ts.cod_set "
            strSQL = strSQL & " and cod_rep='" & !cod_ter & "' "
            rEnd.Open strSQL, conn_client
        End If
        
        Do While Not .EOF
            sNomeTer = !nome
            If iImp = 4 Then
                P_RodapeVis obj
                P_AddPicVis obj
                iImp = 0
            End If
            iImp = iImp + 1
            frmVisualiza.brcdx.caption = !seq_itd
            
            obj.fontsize = 12
            'obj.Print
            P_ImprX obj, "", 0, False
            obj.fontsize = 8
            P_ImprX obj, sNomeTer, 5, False
            P_ImprX obj, "CPF:" & !cpf, 5, False
            
            
            yIni = obj.CurrentY
            obj.DrawWidth = 2
            
            obj.fontsize = 12
            P_Grades obj, obj.CurrentY + 0.4, 4, 195, 1, DrawStyleConstants.vbSolid
            obj.DrawWidth = 1
            obj.FontBold = False
            P_ImprX obj, "Documento:", 5, True
            obj.fontsize = 8
            
            P_ImprX obj, "Local de Pagamento:", 75, True
            
            obj.fontsize = 10
            P_ImprX obj, strFantasia_Emp, 110, True
            obj.FontBold = True
            obj.fontsize = 12
            P_ImprX obj, seq_doc, 35, False
            obj.FontBold = False
            
            
            
            
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Parcela:", 5, True
            obj.fontsize = 10
            P_ImprX obj, "Parcela:", 75, True, esquerda, -1
            P_ImprX obj, "Nº Seq.:", 90, True, esquerda, 0
            P_ImprX obj, "Doc./Venda:", 110, True, esquerda, 0
            P_ImprX obj, "Dt. Proc.:", 140, True, esquerda, 0
            
            obj.FontBold = True
            obj.fontsize = 12
            P_ImprX obj, !parcela_itd & "/" & !parcelas_doc, 35, False
            
            obj.FontBold = False
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Valor:", 5, True
            obj.FontBold = True
            P_ImprX obj, FormatVal(!reman_itd), 35, True
            
            P_ImprX obj, !parcela_itd & "/" & !parcelas_doc, 75, True, esquerda, -1
            P_ImprX obj, !seq_doc, 90, True, esquerda, 0
            P_ImprX obj, !num_doc, 110, True, esquerda, 0
            P_ImprX obj, !data_doc, 140, False, esquerda, 0
            obj.FontBold = False
            
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_Grades obj, obj.CurrentY, 4, 60, 1, vbSolid
            P_ImprX obj, "Vencimento:", 5, True
            obj.FontBold = True
            P_ImprX obj, !venc_itd, 35, True
            obj.FontBold = False
            obj.fontsize = 10
            P_ImprX obj, "Valor Doc:", 75, True, esquerda, -1
            P_ImprX obj, "Valor Parc.", 100, True, esquerda, 0
            P_ImprX obj, "Seq.Parc", 120, True, esquerda, 0
            P_ImprX obj, "Vencimento", 150, True, esquerda, 0
            obj.fontsize = 12
            P_ImprX obj, "", 0, False
            
            
            obj.FontBold = False
            
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Desconto:", 5, True
            obj.FontBold = True
            P_ImprX obj, FormatVal(!valor_doc), 75, True, esquerda, -1
            P_ImprX obj, FormatVal(!reman_itd), 100, True, esquerda, 0
            P_ImprX obj, !seq_itd, 120, True, esquerda, 0
            P_ImprX obj, !venc_itd, 150, False, esquerda, 0
            
            obj.FontBold = False
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            P_ImprX obj, "Juros:", 5, True
            obj.fontsize = 8
            obj.FontBold = False
            P_ImprX obj, "Cliente:", 75, True
            obj.fontsize = 10
            P_ImprX obj, !nome, 85, True
            obj.fontsize = 12
            P_ImprX obj, "", 100, False
            P_Grades obj, obj.CurrentY, 4, 60, 1, vbSolid
            P_ImprX obj, "Outros:", 5, True
            obj.fontsize = 8
            obj.FontBold = False
            P_ImprX obj, "CPF:", 75, True
            obj.fontsize = 10
            P_ImprX obj, !cpf, 90, True
            obj.fontsize = 8
            P_ImprX obj, "Cidade:", 120, True
            obj.fontsize = 10
            P_ImprX obj, rEnd!nome_cid & "/" & rEnd!uf_cid, 131, True
            obj.fontsize = 12
            P_ImprX obj, "", 100, False
            
            P_Grades obj, obj.CurrentY, 4, 60, 1, vbSolid
            P_ImprX obj, "Total:", 5, True
            obj.fontsize = 8
            P_ImprX obj, "End: ", 75, True
            
            P_ImprX obj, StrConv(rEnd!Logradouro & " N°" & rEnd!numero & " " & rEnd!complemento, vbProperCase) & " " & StrConv(rEnd!tipo_set, vbProperCase) & " " & StrConv(rEnd!nome_set, vbProperCase) & " CEP: " & rEnd!cep_cli, 82, True
            obj.fontsize = 12
            P_ImprX obj, "", 100, False
            
            
            obj.FontBold = False
            'P_Line obj, 30, yIni, 30, obj.CurrentY
            P_Grades obj, 4, yIni, obj.CurrentY, ccOrientationVertical, vbSolid
            P_Grades obj, 30, yIni, obj.CurrentY, ccOrientationVertical, vbSolid
            P_Grades obj, 60, yIni, obj.CurrentY, ccOrientationVertical, vbSolid
            P_Grades obj, 72, yIni, obj.CurrentY, ccOrientationVertical, vbSolid
            P_Grades obj, 195, yIni, obj.CurrentY, ccOrientationVertical, vbSolid
            'obj.ScaleMode = vbMillimeters
            P_Grades obj, obj.CurrentY, 4, 195, 1, vbSolid
            
            
            
            
            obj.PaintPicture frmVisualiza.brcdx.Picture, 5, obj.CurrentY
            obj.PaintPicture frmVisualiza.brcdx.Picture, 75, obj.CurrentY
            obj.fontsize = 8
            P_ImprX obj, "Impresso em:" & Now, 195, False, direita
            P_ImprX obj, "", 0, False
            P_Grades obj, obj.CurrentY, 0, 200, 1, vbDot
            If sTipoTer = "R" Then
                
                P_ImprX obj, "Eu, " & sNomeTer & ", reconheço a dívida acima, e estou de acordo com seus vencimentos e valores.", 5, False
                
                P_ImprX obj, strCidade_Emp & ", " & Date & " ______________________________________", 5, False
            End If
            P_ImprX obj, sVia, 0, False
            .MoveNext
        Loop
        .Close
        rEnd.Close
    End With
    
    P_AddPicVis obj
    
    If TypeName(obj) = "Printer" Then
        Printer.EndDoc
        strSQL = "A Impressão Saiu Correta?" & Chr(13)
        strSQL = strSQL & "caso a impressão não tenha sido proveitosa escolha NÃO"
        If MsgBox(strSQL, vbYesNo + vbQuestion) = vbYes Then
            'P_GravaDocImperss seq_doc
        End If
    End If
    
    Exit Function
trata_Erro:
    P_MostraErro
End Function


Private Sub txtDeposito_GotFocus()
    txtDeposito.text = "0"
    txtDeposito.text = restante
    P_SelecTexto txtDeposito
End Sub

Private Sub txtDeposito_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDeposito_LostFocus()
    If CCur(txtDeposito.text) < 0 Then
        txtDeposito.text = "0"
    End If
    
    PerdeFoco txtDeposito
    
    If CCur(txtDeposito.text) > 0 Then
        cmdDep_Click
    End If
    
    
    P_TotalEntrada
End Sub

Private Sub txtendereco_GotFocus()
    P_SelecTexto txtEndereco
End Sub

Private Sub txtendereco_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtendereco_LostFocus()
    P_AjustaTexto txtEndereco
    P_AtualizaVendaEndereco Me.tag, "", txtEndereco.text, ""
End Sub

Private Sub txtPIX_GotFocus()
    txtPIX.text = "0"
    txtPIX.text = restante
    P_SelecTexto txtPIX
    
End Sub

Private Sub txtPIX_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtPIX_LostFocus()
    P_TotalEntrada
    If CCur(txtPIX.text) > 0 Then
        cmdPix_Click
    End If
    If CCur(txtPIX.text) < 0 Then
        txtPIX.text = "0"
    End If
End Sub
