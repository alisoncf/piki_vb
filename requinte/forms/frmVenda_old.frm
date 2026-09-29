VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVenda 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "-Venda de Produtos:"
   ClientHeight    =   10665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15060
   ControlBox      =   0   'False
   Icon            =   "frmVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5520
      ScaleHeight     =   705
      ScaleWidth      =   9345
      TabIndex        =   11
      Top             =   600
      Width           =   9375
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
         Left            =   7920
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDescReais 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   7080
         TabIndex        =   21
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxDescPercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   6240
         TabIndex        =   19
         Text            =   "0,0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(R$)"
         Height          =   195
         Left            =   7080
         TabIndex        =   20
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
         Left            =   8040
         TabIndex        =   22
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(%)"
         Height          =   195
         Left            =   6240
         TabIndex        =   18
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
         Left            =   5400
         TabIndex        =   16
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
         TabIndex        =   12
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
         Left            =   4800
         TabIndex        =   14
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
      TabIndex        =   94
      Top             =   600
      Width           =   5415
      Begin VB.PictureBox pcLimite 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1680
         ScaleHeight     =   495
         ScaleWidth      =   4095
         TabIndex        =   111
         Top             =   1200
         Visible         =   0   'False
         Width           =   4095
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
            TabIndex        =   119
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   2880
            TabIndex        =   118
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
            TabIndex        =   117
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   1920
            TabIndex        =   116
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
            TabIndex        =   115
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1080
            TabIndex        =   114
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
            TabIndex        =   113
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   0
            TabIndex        =   112
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.CheckBox chkEntrega 
         Caption         =   "En&trega (ctrl + E)"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
      Begin VB.PictureBox pcEnd 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         ScaleHeight     =   1455
         ScaleWidth      =   9135
         TabIndex        =   110
         Top             =   1320
         Width           =   9135
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
            TabIndex        =   30
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
            TabIndex        =   28
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
            TabIndex        =   26
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
            TabIndex        =   29
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
            TabIndex        =   27
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
            TabIndex        =   25
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   5040
         Picture         =   "frmVenda.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   240
         Width           =   975
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
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   315
         Left            =   2760
         Picture         =   "frmVenda.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
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
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   330
         Left            =   5040
         Picture         =   "frmVenda.frx":0B0D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   255
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
         TabIndex        =   9
         Tag             =   "00000"
         Top             =   840
         Width           =   4935
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
         TabIndex        =   120
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   4080
         TabIndex        =   5
         Top             =   0
         Width           =   735
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
   Begin VB.PictureBox pcECF 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   15060
      TabIndex        =   53
      Top             =   9750
      Width           =   15060
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9210
      Left            =   0
      ScaleHeight     =   9180
      ScaleWidth      =   5385
      TabIndex        =   68
      Top             =   540
      Width           =   5415
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   5385
         TabIndex        =   86
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
            TabIndex        =   95
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
               TabIndex        =   99
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
               TabIndex        =   98
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
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Left            =   4440
               TabIndex        =   97
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
               TabIndex        =   96
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
            TabIndex        =   88
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
               TabIndex        =   92
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   89
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
            TabIndex        =   87
            Tag             =   "!"
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3825
         ScaleWidth      =   5385
         TabIndex        =   69
         Top             =   5400
         Width           =   5415
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc %"
            Height          =   195
            Left            =   4560
            TabIndex        =   109
            Top             =   2760
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
            Left            =   4560
            TabIndex        =   108
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            Height          =   195
            Left            =   1800
            TabIndex        =   85
            Top             =   720
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
            Left            =   1800
            TabIndex        =   84
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   720
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
            Left            =   120
            TabIndex        =   82
            Top             =   960
            Width           =   1605
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
            Height          =   315
            Left            =   120
            TabIndex        =   81
            Top             =   3000
            Width           =   4365
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Seção:"
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   2760
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
            TabIndex        =   79
            Top             =   2280
            Width           =   5085
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seção:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   2040
            Width           =   510
         End
         Begin VB.Label lblFabricante 
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
            TabIndex        =   77
            Top             =   1680
            Width           =   5085
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   750
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   4320
            TabIndex        =   75
            Top             =   720
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
            TabIndex        =   74
            Top             =   960
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
            Left            =   2760
            TabIndex        =   73
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label lblCapPreco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço:"
            Height          =   195
            Left            =   2760
            TabIndex        =   72
            Top             =   720
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
            TabIndex        =   71
            Top             =   360
            Width           =   5085
         End
         Begin VB.Label lblCapPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto:"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   600
         End
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
         TabIndex        =   93
         Top             =   5160
         Width           =   150
      End
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   15000
      TabIndex        =   58
      Top             =   10170
      Width           =   15060
      Begin VB.CommandButton cmndbtnBuscaFun 
         Height          =   315
         Left            =   3840
         Picture         =   "frmVenda.frx":0C06
         Style           =   1  'Graphical
         TabIndex        =   104
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
         TabIndex        =   103
         Text            =   "Nome do Funcionario"
         Top             =   15
         Width           =   3255
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         Left            =   9000
         TabIndex        =   67
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
         Left            =   9960
         TabIndex        =   66
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
         Left            =   8400
         TabIndex        =   65
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
         Left            =   5880
         TabIndex        =   64
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7560
         TabIndex        =   63
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
         Left            =   11520
         TabIndex        =   62
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5160
         TabIndex        =   61
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
         Left            =   6480
         TabIndex        =   60
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Atendente:"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   59
         Top             =   15
         Width           =   840
      End
   End
   Begin VB.PictureBox pcMsg 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E2FDFE&
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
      Picture         =   "frmVenda.frx":0CFF
      ScaleHeight     =   345
      ScaleWidth      =   15030
      TabIndex        =   56
      Top             =   9795
      Width           =   15060
      Begin VB.Label lblAtacado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "."
         Height          =   195
         Left            =   960
         TabIndex        =   105
         Top             =   120
         Width           =   45
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
         TabIndex        =   57
         Top             =   0
         Width           =   135
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
      ScaleWidth      =   15030
      TabIndex        =   55
      Top             =   0
      Width           =   15060
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
         TabIndex        =   100
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
            Picture         =   "frmVenda.frx":9365
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   101
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
            TabIndex        =   102
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
         Index           =   0
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   31
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
            Picture         =   "frmVenda.frx":9D67
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   32
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
            TabIndex        =   33
            Top             =   120
            Width           =   1200
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
            Index           =   2
            Left            =   120
            Picture         =   "frmVenda.frx":A0F1
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   35
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
            TabIndex        =   36
            Top             =   120
            Width           =   1065
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
            Index           =   3
            Left            =   120
            Picture         =   "frmVenda.frx":A365
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
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
            TabIndex        =   39
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
         Index           =   4
         Left            =   7200
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   40
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
            Picture         =   "frmVenda.frx":A4AF
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   41
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
            TabIndex        =   42
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
         Index           =   5
         Left            =   9000
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
            Height          =   240
            Index           =   5
            Left            =   120
            Picture         =   "frmVenda.frx":AEB1
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   44
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
            TabIndex        =   45
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
         Index           =   6
         Left            =   10800
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   46
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmVenda.frx":D278
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   47
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
            TabIndex        =   48
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
         Index           =   7
         Left            =   12600
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   49
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
            Picture         =   "frmVenda.frx":D602
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   50
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
            TabIndex        =   51
            Top             =   120
            Width           =   735
         End
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
      Left            =   480
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
      Left            =   14760
      TabIndex        =   54
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmVenda.frx":D76D
      Height          =   8415
      Left            =   5520
      OleObjectBlob   =   "frmVenda.frx":D782
      TabIndex        =   52
      Top             =   1320
      Width           =   9375
   End
End
Attribute VB_Name = "frmVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As ProdutoSaida
  
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
Sub P_Msg(msg As String)
    
    lblMensagem.caption = msg
    Me.Refresh
End Sub

Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    Dim cDesconto As Currency
    Dim cFinal As Currency
    On Error GoTo trata_erro:
    P_Status "Atualizando Valores..."
    
    iPos = dtProd.Recordset.AbsolutePosition
    dtProd.Refresh
    With dtProd.Recordset
        
        
        Do While Not .EOF
            cTotal = cTotal + !total
            cFinal = cFinal + !Final
            cDesconto = cDesconto + !Desconto
            cItens = cItens + !qtd_ite
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cFinal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = dtProd.Recordset.RecordCount
        lblTotal.caption = FormatVal(cTotal)
        lblDesconto.caption = FormatVal(cDesconto)
        lblPExtenso.caption = Extenso(CDbl(lblValor.caption))
    End With
        
    
    
    
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    
    lblProd.caption = dtProd.Recordset.RecordCount
    lblItens.caption = cItens

    strSQL = "update tab_vendas set valor_ven =" & F_GravaDec(cTotal) & ","
    strSQL = strSQL & " nomecli_ven= '" & txtbxNomeCli.Text & "', "
    strSQL = strSQL & " cod_cli= '" & txtbxNomeCli.tag & "', "
    strSQL = strSQL & " abrevenda_usu='" & txtbxVendedor.tag & "',"
    
    strSQL = strSQL & "cod_For='" & txtbxFormaPag.tag & "', "
    strSQL = strSQL & "entrega_ven='" & IIf(chkEntrega.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "endereco_ven='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "bairro_ven='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "fone_ven='" & txtbxFone.Text & "',"
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(lblDesconto.caption)
    
    strSQL = strSQL & " where num_Ven='" & Me.tag & "' "
    Banco.Execute strSQL
    
    
    P_Status ""
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    strSQL = "select tv.*,tf.* from tab_vendas as tv, tab_formaspag as tf "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for "
    strSQL = strSQL & " and num_ven='" & Me.tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        txtbxNomeCli.Text = !nomecli_ven
        txtbxNomeCli.tag = !cod_cli
        If !seq_con <> "" Then
            txtbxConveniado.tag = !seq_con
            txtbxConveniado.Text = f_BuscaConveniadoNome(!seq_con, !cod_cli)
            txtbxConveniado.Enabled = ture
            lblConveniado.Enabled = True
            cmndbtnConveniado.Enabled = True
        End If
        txtbxFormaPag.Text = !nome_for
        txtbxDescForma.Text = !desconto_for
        txtbxFormaPag.tag = !cod_for
        txtbxEndereco.Text = Format(!endereco_ven)
        txtbxBairro.Text = Format(!bairro_ven)
        txtbxFone.Text = Format(!fone_ven)
        If !entrega_ven = "S" Then
            chkEntrega.Value = 1
            chkEntrega_Click
        End If
        .Close
    End With
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_erro
    strSQL = Venda.F_BuscaItens(Me.tag)
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    
            
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub


Sub P_InsereProd(sCod_Pro As String, cUnit As Currency, cQtd As Currency)
    Dim strSQL As String
    Dim iIndex As Integer
    On Error GoTo trata_erro
    strSQL = F_BuscaUmCampo("max(seq_ite)", "tab_itvendas", "num_ven='" & Me.tag & "'")
'    If strSQL = "" Then
'        iIndex = 1
'    Else
'        iIndex = val(strSQL) + 1
'    End If
    strSQL = "insert into tab_itvendas(num_ven,cod_pro,unit_ite,qtd_ite,seq_ite,final_ite)values("
    
    strSQL = strSQL & "'" & Me.tag & "', "
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
    strSQL = strSQL & " where num_ven='" & Me.tag & "'  "
    strSQL = strSQL & " and cod_pro ='" & sCod_Pro & "' "
    Banco.Execute strSQL
    
    
    
    'RESERVANDO O ESTOQUE
    strSQL = "update tab_produtos set er_pro=er_pro - " & F_GravaDec(CStr(cQtd))
    strSQL = strSQL & " where cod_pro='" & sCod_Pro & "' "
    Banco.Execute strSQL
    
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
        'If operacao = "C" Then
        '    i = F_RetornoVariaveis
        '    If i Mod 2 > 0 Then
        '        ecf_CancelaCupom
        '        iItem = 0
        '    End If
        'End If
        With dtProd.Recordset
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                excluir False
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
        Banco.Execute strSQL
        
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
            RefreshDataPosition frmEdiVendas.dtVendas
        End If
        
        If MsgBox("Deseja Gerar uma Nova Venda?" & Chr(13) & "Clique em não caso queira abandonar o modo de vendas!", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
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
        frmValoresProd.tag = Me.tag
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
        
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
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
            If MsgBox("Confirma Exclusão do Produto " & !desc_ite & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        iPos = .AbsolutePosition
        strSQL = "delete from tab_itvendas where seq_ite='" & !seq_ite & "' "
        
        strSQL = strSQL & " and num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        'P_GravaItemCancelado
        'voltando o estoque
'        If pcBaixa.Visible = False Then
           strSQL = "update tab_produtos set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))

            
            

            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            Banco.Execute strSQL
        
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
            txtbxQtd.Enabled = True
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
    Exit Sub
trata_erro:
    P_MostraErro "Excluir()"
End Sub


Private Sub chkEntrega_Click()
    If chkEntrega.Value = 1 Then
        lblEnd.Enabled = True
        lblBairro.Enabled = True
        lblFone.Enabled = True
        txtbxEndereco.Enabled = True
        txtbxBairro.Enabled = True
        txtbxFone.Enabled = True
        pcEnd.Visible = True
        txtbxEndereco.SetFocus
    Else
        pcEnd.Visible = False
        lblEnd.Enabled = False
        lblBairro.Enabled = False
        lblFone.Enabled = False
        txtbxEndereco.Enabled = False
        txtbxBairro.Enabled = False
        txtbxFone.Enabled = False
    End If
    
End Sub

Private Sub cmndbtnBuscaFun_Click()
    Dim f As funcionario.funcionario
    On Error GoTo trata_erro
    
    
    f = funcionario.F_SelectVendedor("")
    
    
    
    If f.cod_fun <> "" Then
        If F_senhaReturnUsu(True, False, f.guerra_fun, False) Then
        
        
            txtbxVendedor.Text = f.nome_fun
            txtbxVendedor.tag = f.cod_fun
            lblDesc.caption = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & f.cod_fun & "'")
            P_AtualizaVenda
        End If
    End If
    
    Exit Sub
trata_erro:
    P_MostraErro
    
End Sub
Private Sub cmndbtnCliente_Click()
    Dim c As Clientes.Cliente
    
    c = Clientes.F_BuscaCliente(txtbxNomeCli.Text)
    If c.cod_cli <> "" Then
        txtbxNomeCli.Text = c.nome_Cli
        txtbxNomeCli.tag = c.cod_cli
        lblLimite.caption = FormatVal(c.limite_cli)
        lblEmUso.caption = FormatVal(F_BuscaLimiteEmUsoCli(c.cod_cli))
        lblLimiteRest.caption = FormatVal(c.limite_cli - CCur(lblEmUso.caption))
        lblStatus_CLi.tag = c.status_cli
        If c.status_cli = "1" Then
            lblStatus_CLi.caption = "Liberado"
            lblStatus_CLi.ForeColor = vbGreen
        ElseIf c.status_cli = "2" Then
            lblStatus_CLi.caption = "Cheques"
            lblStatus_CLi.ForeColor = vbYellow
        Else
            lblStatus_CLi.caption = "Bloqueado"
            lblStatus_CLi.ForeColor = vbRed
        End If
        If c.status_cli = "2" Then
            MsgBox "O Cliente encontra-se negativado, portanto, só é possível venda a vista!", vbExclamation
            
            
        End If
        
        
        If chkEntrega.Value = 1 Then
            txtbxEndereco.Text = c.endereco_cli
            txtbxBairro.Text = c.bairro_cli
            txtbxFone.Text = c.telefone_cli
        End If
        
            
            
            
            
        
        
        pcLimite.Visible = True
        
        
        
        P_AtualizaVenda
        
    Else
        
    End If
     
     If txtbxNomeCli.tag = "00000" Or txtbxNomeCli.tag = "" Then
        txtbxNomeCli.BackColor = &H100FFF
        lblNomeCli.caption = "&Cliente(Não Cadastrado)"
    Else
        txtbxNomeCli.BackColor = &H80FFFF
        lblNomeCli.caption = "&Cliente(Cadastrado)"
    End If
     
     
     
            txtbxProduto.SetFocus
     
    
End Sub

Private Sub cmbxFormaPag_Change()

End Sub

Private Sub cmbxFormaPag_DropDown()

End Sub



Private Sub cmndbtnConveniado_Click()
    Dim c As conveniado.conveniado
    c = conveniado.F_SelectConveniado(txtbxConveniado.Text, txtbxNomeCli.tag)
    
    
    If c.cod_cli <> "" Then
        txtbxConveniado.Text = c.nome_con
        txtbxConveniado.tag = c.seq_con
        
        P_AtualizaVenda
        txtbxProduto.SetFocus
    End If
End Sub

Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    f = F_SelectFormasPag(txtbxFormaPag.Text)
    
    If f.cod_for <> "" Then
        txtbxFormaPag.Text = f.nome_for
        txtbxFormaPag.tag = f.cod_for
        txtbxDescForma.Text = f.desconto_for
        txtbxDescForma.tag = f.avista_for
        txtbxDescPercent.Text = txtbxDescForma.Text
        lblTipoPag.caption = f.nome_tip
        lblTipoPag.tag = f.cod_tip
        
        'P_recalculaVenda
        
        P_AtualizaVenda
        txtbxNomeCli.SetFocus
    End If
    
    
    
    
End Sub

Private Sub cmndbtnImprPos_Click()
    Venda.P_ImprPosologia Me.tag
End Sub

Private Sub cmndbtnLocaliza_Click()
    
    If F_Vazio(txtbxProduto.Text) Then
'        MsgBox "Informe o produto que deseja inserir!", vbExclamation, constNome_sis
'        txtbxProduto.SetFocus
'        Exit Sub
    End If
    
    buscaProduto
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    
    If dtProd.Recordset.RecordCount > 0 Then
        
            If MsgBox("Confirma Saída da Venda?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
    End If
    If IsLoad(frmEdiVendas.Name) Then
        frmEdiVendas.dtVendas.Refresh
    End If
    If dtProd.Recordset.RecordCount = 0 Then
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        
    End If
    'i = F_RetornoVariaveis
    '    If i Mod 2 > 0 Then
    '        ecf_CancelaCupom
    '    End If
    
    Unload Me
    Exit Sub
trata_erro:
    P_MostraErro "cmndbtnSair"
    Unload Me
End Sub


Private Sub cmndbtnSalvar_Click()
    Dim sTroco As String
    Dim strSQL As String
    Dim sVen As String
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não há itens na sua venda!", vbExclamation
        Exit Sub
    End If
    
    
    strSQL = F_senhaReturnUsu(False, 3)
    If strSQL = "" Then
        MsgBox "não autenticado!", vbExclamation
        Exit Sub
    End If
    If strSQL <> txtbxVendedor.tag Then
        txtbxVendedor.tag = strSQL
        txtbxVendedor.Text = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & strSQL & "'")
        Venda.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text, txtbxConveniado.tag
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
                Venda.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text, ""
            End If
            
            
            
        End If
        If lblTipoPag.tag = "02" Then
            If F_GravaCrediarioVen(Me.tag) Then
                MsgBox "Esta venda venda deve ser liberada pelo setor de Crediário para ser concretizada!", vbExclamation
            End If
        End If
        
        If Venda.F_AlteraStatusVenda(Me.tag, "2") Then
            Me.tag = F_GeraNovaVenda
            bActive = False
            Form_Activate
        End If
    
    
    
End Sub


Private Sub dbgrdProd_DblClick()
    cmndbtnEditProd_Click
End Sub

Private Sub dbgrdProd_GotFocus()
    
    dbgrdProd.MarqueeStyle = 3
    
    
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error GoTo trata_erro

    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        
        
                
                    If dbgrdProd.Row = -1 Then
                        MsgBox "Selecione o item que deseja excluir!", vbCritical
                        dbgrdProd.SetFocus
                        Exit Sub
                    End If
                    
            excluir False
        
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxQtd.SetFocus
    End If
    P_AtualizaVenda
    Exit Sub
trata_erro:
P_MostraErro
End Sub

Private Sub dbgrdProd_LostFocus()
    dbgrdProd.MarqueeStyle = 0
End Sub


Private Sub dbgrdProd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub dtProd_Reposition()
    With dtProd.Recordset
    On Error Resume Next
    If Not .EOF Then
        lblProduto.caption = !nome_pro
            
            
                lblCapPreco.caption = "Preço:"
                lblPreco.caption = "R$ " & FormatVal(!pv_pro)
            
            lblCat.caption = !nome_cat
            lblCod.caption = FormatCodigo(!cod_pro)
            lblCodBarr.caption = F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'")
            lblMarca.caption = !nome_mar
            lblQtd.caption = !qtd_ite
            lblRef.caption = !codfab_pro
            lblBalanca.caption = !balanca_pro
            lblPosologia.caption = F_BuscaUmCampo("descricao_pos", "tab_posologia", "num_ven='" & Me.tag & _
                "' and seq_ite=" & dtProd.Recordset!seq_ite)
            
            
        End If
    End With
        
End Sub


Private Sub Form_Activate()
    'On Error GoTo trata_erro
    mdiFrm.Enabled = False
    
    lblnum_ven.caption = Me.tag
    'lblNome_Fun.caption = strCod_Usu & "-" & strNome_Usu
    lblNomeEmp.caption = strFantasia_Emp
    lblCaixa = strCod_Cai
    lblTur.caption = strCod_Tur
    If bActive = True Then Exit Sub
    
    
    
    carrega
    If txtbxNomeCli.tag = "00000" Or txtbxNomeCli.tag = "" Then
        txtbxNomeCli.BackColor = &H100FFF
        lblNomeCli.caption = "&Cliente(Não Cadastrado)"
    Else
        txtbxNomeCli.BackColor = &H80FFFF
        lblNomeCli.caption = "&Cliente(Cadastrado)"
    End If
    
    P_Status "Buscando Itens..."
    txtbxProduto.SetFocus
    P_buscaitens
    P_AtualizaVenda
    P_Status ""
    If Atacado Then
        lblAtacado.Visible = True
        Form_KeyDown vbKeyF9, 0
    Else
        lblAtacado.Visible = False
    End If
    bActive = True
    If strCod_Cai = "" Then
        MsgBox "Não existe caixa para este computador, o caixa ficará para o '01' ", vbExclamation
        strCod_Cai = "00"
    End If
    
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
    ElseIf KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    ElseIf KeyCode = vbKeyF5 Then
    
        
        dbgrdProd.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
         dbgrdProd_KeyDown vbKeyDelete, 0
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF2 Then
        
    ElseIf KeyCode = vbKeyF1 Then
        pcBotao_Click 1
    ElseIf KeyCode = vbKeyF8 Then
        If operacao = "B" And dtProd.Recordset.RecordCount > 0 Then
            MsgBox "Operacação Impossível!", vbCritical
            
            Exit Sub
        End If
        
        pcECF.BackColor = vbGreen
        operacao = "C"
        mudamodo
    ElseIf KeyCode = vbKeyF9 Then
        'If operacao = "C" And dtProd.Recordset.RecordCount > 0 Then
        '    MsgBox "Operacação Impossível!", vbCritical
        '    Exit Sub
        'End If
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
    
    If KeyCode = vbKeyC And (Shift = vbAltMask Or Shift = vbCtrlMask) Then
        txtbxNomeCli.SetFocus
    ElseIf KeyCode = vbKeyP And Shift = vbCtrlMask Then
        
    ElseIf KeyCode = vbKeyE And Shift = vbAltMask Then
        chkEntrega_Click
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
    
    bActive = False

    dtProd.Connect = Banco.Connect

    lblNomeEmp.caption = strNome_Emp
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", "tab_parametros", "")
    
    
    
    
    
'    cmbxBuscar.ListIndex = 0
    pcECF.BackColor = vbGreen
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
    txtbxVendedor.tag = strCod_Usu
    txtbxVendedor.Text = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & strCod_Usu & "'")
    lblDesc.caption = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & strCod_Usu & "'")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    dbgrdProd.Width = Me.Width - dbgrdProd.Left - 200
    
    
    dbgrdProd.Height = pcECF.Top - 2500
    lblNomeEmp.Left = Me.Width - lblNomeEmp.Width - 200
    'pcCli.Width = pcProd.Width
    
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).Width = pcAtalhos.Width / (pcBotao.count - 1)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).Width + 50
        End If
        
    Next
 
    pcProd.Width = dbgrdProd.Width
    'pcCli.Width = pcProd.Width
    
    
End Sub

Sub editaItem()
    txtbxQtd.Enabled = True
    txtbxQtd.Text = "1"
    txtbxQtd.SetFocus
End Sub

Private Sub optnCod_Click()
    txtbxQtd.SetFocus
End Sub

Private Sub optnDesc_Click()
    txtbxQtd.SetFocus
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
    mdiFrm.Enabled = True
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

Private Sub lblMsg2_Click()

End Sub



Private Sub optnCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtbxNomeCli.SetFocus
    End If

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
        
        cmndbtnFormasPag_Click
    ElseIf Index = 0 Then
'        MsgBox "Não há TEF"
    ElseIf Index = 7 Then
        
    End If
End Sub
Sub lancaPosologia()
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não existe itens na sua venda!", vbExclamation
        Exit Sub
    End If
    If dbgrdProd.Row = -1 Then
        MsgBox "Selecione o Item para lançar a posologia!", vbExclamation
        dbgrdProd.SetFocus
        Exit Sub
    End If
    Venda.P_AddPosologia Me.tag, dtProd.Recordset!seq_ite
    
End Sub
Private Sub pcBotao_GotFocus(Index As Integer)
    P_foco Index
End Sub

Private Sub pcbotao_LostFocus(Index As Integer)
    pcBotao(Index).BorderStyle = 0
End Sub

Private Sub pcMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub pcPromo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    
iGif = iGif + 1
If iGif >= 15 Then
    If iIndPromo >= QtdPromo - 1 Then
        
        iIndPromo = 0
    End If
    lblPromo.caption = QuebraLinha(StrConv(pPromo(iIndPromo).nome, vbProperCase), 30)
    lblPromoPreco.caption = pPromo(iIndPromo).preco
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


Private Sub pcProd_GotFocus()
    SendKeys Chr(9)
End Sub

Private Sub txtbxBairro_GotFocus()
    P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    P_AjustaTexto txtbxBairro
    P_AtualizaVenda
End Sub


Private Sub txtbxDescPercent_GotFocus()
    P_SelecTexto txtbxDescPercent
End Sub

Private Sub txtbxDescPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDescPercent_LostFocus()
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDescReais_LostFocus()
    On Error Resume Next
    txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    p_totalItem
End Sub

Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
    P_AtualizaVenda
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
    On Error Resume Next
    txtbxDescReais.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxFinal.Text))
    txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    Exit Sub
trata_erro:
    
End Sub

Private Sub txtbxFone_GotFocus()
    P_SelecTexto txtbxFone
End Sub

Private Sub txtbxFone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxFone_LostFocus()
    P_AjustaTexto txtbxFone
    P_AtualizaVenda
End Sub

Private Sub txtbxFormaPag_GotFocus()
    P_SelecTexto txtbxFormaPag
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxFormaPag_LostFocus()
    P_AtualizaVenda
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
    If KeyCode = vbKeyDown Then
        If txtbxConveniado.Enabled Then
            txtbxConveniado.SetFocus
        Else
            txtbxProduto.SetFocus
        End If
    End If
End Sub


Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub

Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli, 1
    If txtbxNomeCli.tag <> "00000" Then
        txtbxNomeCli.Text = F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & txtbxNomeCli.tag & "'")
        Venda.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text, ""
    Else
        Venda.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.Text, ""
    End If
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
    If KeyCode = vbKeyUp Then
        dbgrdProd.SetFocus
    End If
End Sub

Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim e As EstadoImp
    
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTipoQtd As String
    If CCur(txtbxDescPercent.Text) > CCur(lblDesc.caption) Then
        MsgBox "O desconto máximo permitido para o usuário é de " & lblDesc.caption, vbExclamation
        Exit Sub
    End If
    
    If txtbxProduto.Text = "" Then
        MsgBox "Informe pelo menos parte do nome do produto para que seja localizado!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    On Error GoTo trata_erro
                
                
                                
                iItem = iItem + 1
                If Venda.F_AdicionaItem(Me.tag, p, txtbxQtd.Text, txtbxDescReais.Text, txtbxFinal.Text) = "" Then
                    MsgBox "Erro na inserção do item!", vbExclamation
                Else
                    If Not F_MudaQtdEstoque(p.cod_pro, "-" & txtbxQtd.Text) Then
                        MsgBox "Erro na baixa do estoque.", vbExclamation
                    End If
                    
                End If
                    
                P_buscaitens
                Me.Refresh
                
                txtbxQtd.Text = "1"
                dtProd.Recordset.Move dtProd.Recordset.RecordCount - 1
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
Sub buscaProduto()
    
    
    p = Produtos.F_BuscaProdSaida(txtbxProduto.Text, "B")
            
            
            
        If p.cod_pro <> "" Then
                               
            txtbxProduto.Text = p.nome_pro
                
            lblNomePro.caption = p.nome_pro
            
            lblCodPro.caption = p.cod_pro
'            lblSeqPro.caption = p.est.seq_pro
            
            
            lblSec.caption = p.sub.sec.nome_sec
            lblSubSec.caption = p.sub.nome_sub
            
            lblDescSub.caption = p.sub.descmax_sub
            lblPreco.caption = p.pv_pro
            lblRef.caption = p.ref_pro
            If p.pb_pro > 0 Then
                txtbxUnit.Text = FormatVal(p.pb_pro)
            Else
                txtbxUnit.Text = FormatVal(p.pv_pro)
            End If
            txtbxDescPercent.Text = CCur(txtbxDescForma.Text) + CCur(p.sub.descmax_sub)
            
            SendKeys Chr(9)
        End If
End Sub
Sub buscaDesc()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_descontos where cod_pro='" & txtbxProduto.tag & "' "
    strSQL = strSQL & " and cod_car='"
    
    
    
    
    
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
      
    End If
Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub p_totalItem()
    txtbxFinal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxDescReais.Text))
    
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
    p_totalItem
    'txtbxQtd.Enabled = False
    'txtbxProduto.Text = ""
    'txtbxProduto.Enabled = True
    'txtbxProduto.SetFocus
    
    
    
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    Timer1_Timer
End Sub

Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit, "###,##0.000"
    p_totalItem
End Sub
Sub P_GravaItemCancelado()
    On Error Resume Next
    strSQL = "insert into tab_itvndcncl(seq_itv,num_ven,seq_ite,qtd_ite,unit_ite,cod_cai,cod_tur)values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_itvndcncl", "seq_itv", "", "000000#") & "',"
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "" & dtProd.Recordset!seq_ite & ","
    strSQL = strSQL & F_GravaDec(dtProd.Recordset!qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(dtProd.Recordset!unit_ite) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "') "
    Banco.Execute strSQL
    
    
    
End Sub
Sub P_recalculaVenda()
    Venda.P_recalculaVenda txtbxDescForma.Text, Me.tag
    
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
Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
End Sub










Private Sub txtbxVendedor_GotFocus()
    P_SelecTexto txtbxVendedor
End Sub

Private Sub txtbxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaFun_Click
    End If
End Sub
