VERSION 5.00
Begin VB.Form frmPendenciaCrediario 
   Caption         =   "Pendência de Crediário"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   11145
   Icon            =   "frmPendenciaCrediario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   11145
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8205
      Left            =   9090
      ScaleHeight     =   8175
      ScaleWidth      =   2025
      TabIndex        =   26
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdCheques 
         Caption         =   "Cheques"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4320
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "F5 - Carnês"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":0344
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "F5 - Histórico"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnEditar 
         Caption         =   "F6 - Editar Cliente"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":1756
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnEstoque 
         Caption         =   "F9 - Clientes"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":2020
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "F8 - Novo Cliente"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":2CEA
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnAcerto 
         Caption         =   "F7 - Finalizar"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":2E4E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnSair 
         Caption         =   "F12 - Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmPendenciaCrediario.frx":31D8
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   7320
         Width           =   1815
      End
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4305
      ScaleWidth      =   9105
      TabIndex        =   0
      Top             =   660
      Width           =   9135
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   16
         Tag             =   "00000"
         Top             =   960
         Width           =   7935
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
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   315
         Left            =   2760
         Picture         =   "frmPendenciaCrediario.frx":3343
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
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
         TabIndex        =   12
         Text            =   "0"
         Top             =   240
         Width           =   375
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
         Left            =   3600
         TabIndex        =   11
         Text            =   "0,00"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   4560
         Picture         =   "frmPendenciaCrediario.frx":343C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pcLimite 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   10215
         TabIndex        =   1
         Top             =   1560
         Width           =   10215
         Begin VB.Label lblCEP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Left            =   3480
            TabIndex        =   47
            Top             =   2400
            Width           =   195
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEP"
            Height          =   195
            Left            =   3480
            TabIndex        =   46
            Top             =   2160
            Width           =   315
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   5760
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label lblCidade 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            TabIndex        =   38
            Top             =   2400
            Width           =   195
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   0
            TabIndex        =   37
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label lblSetor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            TabIndex        =   36
            Top             =   1680
            Width           =   195
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Setor"
            Height          =   195
            Left            =   0
            TabIndex        =   35
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Left            =   1800
            TabIndex        =   34
            Top             =   960
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End:"
            Height          =   195
            Left            =   1800
            TabIndex        =   33
            Top             =   720
            Width           =   330
         End
         Begin VB.Label LBLcPF 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            TabIndex        =   32
            Top             =   960
            Width           =   195
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/CNPJ:"
            Height          =   195
            Left            =   0
            TabIndex        =   31
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   615
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
            TabIndex        =   8
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1080
            TabIndex        =   7
            Top             =   0
            Width           =   405
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
            TabIndex        =   6
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   1920
            TabIndex        =   5
            Top             =   0
            Width           =   600
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
            TabIndex        =   4
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   2880
            TabIndex        =   3
            Top             =   0
            Width           =   690
         End
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
            TabIndex        =   2
            Top             =   240
            Width           =   390
         End
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   450
         Left            =   8040
         Picture         =   "frmPendenciaCrediario.frx":3586
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Forma de Pagamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3120
         TabIndex        =   20
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   3600
         TabIndex        =   19
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblTipoPag 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DUP"
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
         TabIndex        =   17
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8205
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   10425
      TabIndex        =   22
      Top             =   0
      Width           =   10455
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         ScaleHeight     =   255
         ScaleWidth      =   3615
         TabIndex        =   57
         Top             =   5040
         Width           =   3615
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
            TabIndex        =   61
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
            Left            =   1080
            TabIndex        =   60
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
            Left            =   2880
            TabIndex        =   59
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
            Left            =   1800
            TabIndex        =   58
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4815
         TabIndex        =   52
         Top             =   5040
         Width           =   4815
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
            Left            =   1800
            TabIndex        =   56
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
            Left            =   2880
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   585
         ScaleWidth      =   10305
         TabIndex        =   39
         Top             =   0
         Width           =   10335
         Begin VB.Label lblVendedor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            Left            =   3120
            TabIndex        =   45
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblDataVen 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            TabIndex        =   43
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
            Height          =   195
            Left            =   1080
            TabIndex        =   42
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblNumVen 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
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
            TabIndex        =   41
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Venda Nº:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   960
         Picture         =   "frmPendenciaCrediario.frx":3D6D
         ScaleHeight     =   2295
         ScaleWidth      =   6855
         TabIndex        =   23
         Top             =   5760
         Width           =   6855
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
            ForeColor       =   &H00000000&
            Height          =   1440
            Left            =   2400
            TabIndex        =   24
            Tag             =   "!"
            Top             =   240
            Width           =   2250
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
         Left            =   4920
         TabIndex        =   25
         Top             =   5400
         Width           =   150
      End
   End
End
Attribute VB_Name = "frmPendenciaCrediario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub cmdCheques_Click()
        
    frmControleChequesReceber.txtbxExpr.text = txtbxNomeCli.text
    frmControleChequesReceber.txtbxExpr.tag = txtbxNomeCli.tag
    frmControleChequesReceber.txtbxDataFin.text = ""
    frmControleChequesReceber.txtbxDataIni.text = ""
    
    frmControleChequesReceber.Show 0
    frmControleChequesReceber.SetFocus
    

End Sub

Private Sub cmndbtnAcerto_Click()
    Dim s As String
    Dim sStatus As String
    'Permissao_usu.cred_resolver_pendencias_crediario = True
    
   ' If Permissao_usu.cred_resolver_pendencias_crediario = False Then
   '     MsgBox "Você não tem permissão para esta operação!", vbExclamation
   '     Exit Sub
   ' End If
    If txtbxNomeCli.tag = "" Or txtbxNomeCli.tag = "0" Then
        MsgBox "Escolha o cliente!", vbExclamation
        txtbxNomeCli.SetFocus
        Exit Sub
        
    End If
    
    
    If Not F_CPF(lblCPF.caption) And Not F_CGC(lblCPF.caption) Then
        MsgBox "O cliente está com CPF inválido!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirma Fechamento da Pendência?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    If Vendas.F_AlteraCrediarioVen(Me.tag, "2", "Avaliação Positiva!") Then
        Vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text
        
        MsgBox "A Venda deve ser concluída No Caixa!", vbInformation
        
        
        sStatus = F_BuscaUmCampo("status_Cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
        If sStatus = "F" Then
            MsgBox "Não é possível efetuar o recebimento pois este caixa está fechado!", vbExclamation
            Exit Sub
        End If
        P_DeletaVendaAberta Me.tag
        s = F_Troco(Me.tag)
                
        Unload Me
        
    End If
    P_DeletaVendaAberta Me.tag
End Sub

Private Sub cmndbtnCliente_Click()
    Dim c As cliente
    
    If txtbxNomeCli.tag <> "" And txtbxNomeCli.tag <> "0" Then
        c = Clientes.F_BuscaCliente(txtbxNomeCli.tag)
    Else
        If txtbxNomeCli.text = "" Then
            'MsgBox "Informe o nome ou parte do nome do cliente. Seu código ou CPF", vbExclamation
            txtbxNomeCli.SetFocus
            Exit Sub
        End If
        c = Clientes.F_BuscaCliente(txtbxNomeCli.text)
    End If
    
    If c.cod_cli <> "" Then
        txtbxNomeCli.BackColor = vbYellow
        txtbxNomeCli.text = c.nome_Cli
        txtbxNomeCli.tag = c.cod_cli
        lblStatus_CLi.caption = c.status_cli
        lblLimite.caption = FormatVal(c.limite_cli)
        lblEmUso.caption = FormatVal(F_BuscaLimiteEmUsoCli(c.cod_cli))
        lblLimiteRest.caption = FormatVal(c.limite_cli - CCur(lblEmUso.caption))
        lblStatus_CLi.tag = c.status_cli
        lblCPF.caption = c.cnpj_Cli
        lblCEP.caption = c.cep_cli
        strSQL = "update tab_vendas set cod_cli='" & txtbxNomeCli.tag & "' "
        strSQL = strSQL & " where num_ven='" & Me.tag & "'"
        Conn.Execute strSQL
        
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
        lblEnd.caption = c.endereco_cli
        lblSetor.caption = c.set.nome
        lblCidade.caption = c.set.Cidade
        lblCEP.caption = c.cep_cli
        
        
        
    End If
    If txtbxNomeCli.tag <> "" And txtbxNomeCli.tag <> "00000" Then
        pcLimite.Visible = True
    End If
    P_Status ""
End Sub

Sub busca()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    
    strSQL = F_VerificaVendaAberta(Me.tag)
    If strSQL <> "" And strSQL <> strIP_Usu Then
        MsgBox "A Venda se encontra aberta no computador " & strSQL, vbExclamation
        Unload Me
    End If
    
    
    
    P_GravaVendaAberta Me.tag
    
    strSQL = "select tfp.*,tc.cod_cli,tc.nome_cli,tv.valor_ven,data_ven,tv.nomecli_ven,data_cre,tp.*,tf.guerra_fun from tab_vendas as tv,"
    strSQL = strSQL & "tab_clientes as tc,tab_Crediarioven as tp,tab_formaspag as tfp,"
    strSQL = strSQL & "tab_funcionarios as tf where tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " and tfp.cod_for=tv.cod_for "
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    strSQL = strSQL & " and tp.num_ven=tv.num_ven "
    strSQL = strSQL & " and tv.num_ven='" & Me.tag & "'"
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        
        lblValor.caption = FormatVal(!valor_ven)
        txtbxNomeCli.text = !nomecli_ven
        txtbxNomeCli.tag = !cod_cli
        lblNumVen.caption = !num_ven
        lblVendedor.caption = !guerra_fun
        lblDataVen.caption = Format(!data_cre)
        txtbxFormaPag.text = !nome_for
        txtbxFormaPag.tag = !cod_For
        lblTipoPag.tag = !cod_tip
        .Close
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Sub buscaCli()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
End Sub

Private Sub cmndbtnDeb_Click()
    frmEditaMov.cmbxTipoTer.ListIndex = 0
    
    frmEditaMov.txtbxTerceiro.text = txtbxNomeCli.text
    frmEditaMov.txtbxTerceiro.tag = txtbxNomeCli.tag
    frmEditaMov.txtbxDataFin.text = ""
    frmEditaMov.txtbxDataIni.text = ""
    frmEditaMov.tag = "crediário"
    frmEditaMov.Show 0
End Sub

Private Sub cmndbtnEditar_Click()
    Dim fr As New frmCadCli
    If txtbxNomeCli.tag = "" Or txtbxNomeCli.tag = 0 Then
        MsgBox "Informe primeiro cliente!", vbExclamation
        txtbxNomeCli.SetFocus
        Exit Sub
    End If
        
    strHora = Time
    
    fr.tag = txtbxNomeCli.tag
    
    fr.Show 0
End Sub

Private Sub cmndbtnEstoque_Click()
    frmEdiClientes.txtbxExpr.text = txtbxNomeCli.text
    frmEdiClientes.Show 0, mdiFrm
End Sub

Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    f = F_SelectFormasPag(txtbxFormaPag.text)
    
    If f.cod_For <> "" Then
        txtbxFormaPag.text = f.nome_for
        txtbxFormaPag.tag = f.cod_For
        txtbxDescForma.text = f.desconto_for
        txtbxDescForma.tag = f.avista_For
        'txtbxDescPercent.Text = txtbxDescForma.Text
        
        lblTipoPag.caption = f.nome_tip
        lblTipoPag.tag = f.cod_tip
        
        'P_recalculaVenda
        
        'P_AtualizaVenda
        txtbxNomeCli.SetFocus
    End If
    

End Sub

Private Sub cmndbtnHist_Click()
    If txtbxNomeCli.tag = "" Then
        Exit Sub
    End If
    P_histCli txtbxNomeCli.tag
End Sub


Private Sub cmndbtnSair_Click()
    P_DeletaVendaAberta Me.tag
    Unload Me
    
End Sub

Private Sub cmndbtnTransf_Click()
    Dim fr As New frmCadCli
    
    fr.txtbxRazao.text = txtbxNomeCli.text
    fr.Show 0
    
    Do While fr.Visible = True
        DoEvents
    Loop
    cmndbtnCliente_Click
    
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    cmndbtnCliente_Click
    busca
    bActive = True
End Sub

Private Sub Form_Load()
    bActive = False
    resizeForm Me
End Sub

Private Sub txtbxNomeCli_Change()
    txtbxNomeCli.tag = ""
    
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
End Sub
