VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMovDiarioCx 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimentação Diária do Caixa"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      Height          =   255
      Left            =   120
      TabIndex        =   81
      Top             =   6120
      Width           =   1335
   End
   Begin TabDlg.SSTab sstb 
      Height          =   1095
      Left            =   1680
      TabIndex        =   72
      Top             =   6120
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Por Caixa"
      TabPicture(0)   =   "frmMovDiarioCx.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Por Usuário"
      TabPicture(1)   =   "frmMovDiarioCx.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Usuário"
         Height          =   615
         Left            =   240
         TabIndex        =   79
         Top             =   360
         Width           =   2775
         Begin VB.ComboBox cmbxUsu 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "&Turno:"
         Height          =   615
         Left            =   -71880
         TabIndex        =   77
         Top             =   360
         Width           =   1935
         Begin VB.ComboBox cmbxTurno 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "F&undo de Caixa:"
         Height          =   615
         Left            =   -69840
         TabIndex        =   75
         Top             =   360
         Width           =   1455
         Begin VB.TextBox txtbxFundoCaixa 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   76
            Text            =   "0,00"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cai&xa"
         Height          =   615
         Left            =   -74760
         TabIndex        =   73
         Top             =   360
         Width           =   2775
         Begin VB.ComboBox cmbxCaixa 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   240
            Width           =   2175
         End
      End
   End
   Begin VB.CommandButton cmndbtnFecharCaixa 
      Caption         =   "&Fechar Caixa"
      Height          =   735
      Left            =   120
      Picture         =   "frmMovDiarioCx.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   735
      Left            =   1920
      Picture         =   "frmMovDiarioCx.frx":0182
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnLan 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   3720
      Picture         =   "frmMovDiarioCx.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   8640
      Picture         =   "frmMovDiarioCx.frx":0768
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   7320
      Width           =   1695
   End
   Begin VB.PictureBox pcECF 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   10095
      TabIndex        =   60
      Top             =   6000
      Width           =   10095
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Data:"
      Height          =   735
      Left            =   120
      TabIndex        =   58
      Top             =   6480
      Width           =   1335
      Begin VB.TextBox txtbxData 
         Height          =   285
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox pcControle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5865
      ScaleWidth      =   10065
      TabIndex        =   14
      Top             =   0
      Width           =   10095
      Begin VB.Frame frControle 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   0
         TabIndex        =   15
         Top             =   -120
         Width           =   9975
         Begin VB.CommandButton cmndbtnDetalhesBaixa 
            Caption         =   "+"
            Height          =   255
            Left            =   6000
            TabIndex        =   71
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnDet 
            Caption         =   "Detalhar"
            Height          =   255
            Left            =   8040
            TabIndex        =   16
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label lblSaidasFun 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2040
            TabIndex        =   70
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   4560
            Width           =   480
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saídas Para Func:"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label lblTotalAvulso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   8640
            TabIndex        =   68
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   960
            Width           =   480
         End
         Begin VB.Label lblAvulsoDeb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   67
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   960
            Width           =   480
         End
         Begin VB.Label lblAvulsoCred 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   66
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Movimento Avulso"
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
            Left            =   120
            TabIndex        =   65
            Top             =   960
            Width           =   1920
         End
         Begin VB.Label lblTotalRec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   57
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   3600
            Width           =   480
         End
         Begin VB.Label lblTotalRecAP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   56
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total À Prazo"
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
            Left            =   120
            TabIndex        =   55
            Top             =   3240
            Width           =   1560
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   9600
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   9600
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label lblTotalRecAv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   54
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total À Vista"
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
            Left            =   120
            TabIndex        =   53
            Top             =   2160
            Width           =   1560
         End
         Begin VB.Label lblRecCHP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   52
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   2760
            Width           =   480
         End
         Begin VB.Label lblTotalCRD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   8640
            TabIndex        =   51
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label lblSaiCRD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   50
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label lblCRD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   49
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cartão de Débito"
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
            Left            =   120
            TabIndex        =   48
            Top             =   1800
            Width           =   1920
         End
         Begin VB.Label lblRCD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   47
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1800
            Width           =   480
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   9600
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   9600
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label lblRecCar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   46
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label lblRCC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   45
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label lblRecCHV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   44
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Left            =   8640
            TabIndex        =   43
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saídas"
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
            Left            =   6360
            TabIndex        =   42
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recebimentos"
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
            Left            =   4560
            TabIndex        =   0
            Top             =   480
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendas/Entradas"
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
            Left            =   2400
            TabIndex        =   1
            Top             =   480
            Width           =   1800
         End
         Begin VB.Label lblRec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5040
            TabIndex        =   2
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Venda em Dinheiro"
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
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   2040
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carteira:"
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
            Left            =   120
            TabIndex        =   4
            Top             =   2520
            Width           =   1080
         End
         Begin VB.Label lblCar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   5
            Tag             =   "!"
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label lblDin 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   6
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblTotalEntradas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   240
            Left            =   3120
            TabIndex        =   7
            Tag             =   "!"
            Top             =   3600
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total.........:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   8
            Top             =   3600
            Width           =   1800
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque A Prazo:"
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
            Left            =   120
            TabIndex        =   9
            Top             =   2760
            Width           =   1800
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque à Vista:"
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
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1800
         End
         Begin VB.Label lblCHV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   11
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cartão de Crédito"
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
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   2040
         End
         Begin VB.Label lblFin 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   41
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label lblTotalAv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8640
            TabIndex        =   40
            Tag             =   "!"
            ToolTipText     =   "Total do movimento à vista"
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label lblCHP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   39
            Tag             =   "!"
            Top             =   2760
            Width           =   480
         End
         Begin VB.Label lblTotalAp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8640
            TabIndex        =   38
            Tag             =   "!"
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label lblTotalSai 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   6480
            TabIndex        =   37
            Tag             =   "!"
            ToolTipText     =   "Total de saídas"
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label lblTotalDin 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   8640
            TabIndex        =   36
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(C)"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   240
            Left            =   3360
            TabIndex        =   35
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(D)"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   240
            Left            =   6600
            TabIndex        =   34
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblTotalSai2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   6480
            TabIndex        =   33
            Tag             =   "!"
            Top             =   3600
            Width           =   480
         End
         Begin VB.Label lblTotalAB 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   8640
            TabIndex        =   32
            Tag             =   "!"
            Top             =   3600
            Width           =   480
         End
         Begin VB.Line Line8 
            X1              =   120
            X2              =   9600
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line9 
            X1              =   120
            X2              =   9600
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Espécie"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000C&
            Height          =   240
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelamentos:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   4200
            Width           =   1140
         End
         Begin VB.Label lblCancela 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   240
            Left            =   2040
            TabIndex        =   29
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   4200
            Width           =   480
         End
         Begin VB.Label lblDinDeb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   28
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblCHVDeb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   27
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label lblFinanDeb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   26
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label lblCarDeb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   25
            Tag             =   "!"
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label lblCHPDeb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   24
            Tag             =   "!"
            Top             =   2760
            Width           =   480
         End
         Begin VB.Label lblTotalCHV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   8640
            TabIndex        =   23
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label lblTotalRCC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   8640
            TabIndex        =   22
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label lblTotalEntr 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   21
            Tag             =   "!"
            ToolTipText     =   "total de entradas à v ista"
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label lblTotalCar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   8640
            TabIndex        =   20
            Tag             =   "!"
            ToolTipText     =   "Vendas a Vista"
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label lblTotalCHP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   8640
            TabIndex        =   19
            Tag             =   "!"
            Top             =   2760
            Width           =   480
         End
         Begin VB.Label lblTotalApDeb 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6480
            TabIndex        =   18
            Tag             =   "!"
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label lblTotalApCred 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3120
            TabIndex        =   17
            Tag             =   "!"
            Top             =   3240
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pcRestr 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   11400
      ScaleHeight     =   1515
      ScaleWidth      =   4515
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "frmMovDiarioCx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim operacao As String
Private Sub cmbxCaixa_Click()
        'cmbxUsu.ListIndex = -1
    'busca
End Sub
Function VendasFechadas() As Boolean
    If F_BuscaUmCampo("num_ven", "tab_vendas", "situacao_ven=2 and cod_cai='" & strCod_Cai & "'") <> "" Then
        VendasFechadas = True
    Else
        VendasFechadas = False
    End If
End Function

Private Sub cmbxTurno_Click()
    'busca
End Sub

Private Sub cmbxTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxUsu_Click()
    'cmbxCaixa.ListIndex = -1
    cmbxUsu.tag = F_BuscaUmCampo("cod_fun", "tab_funcionarios", "guerra_fun='" & cmbxUsu.text & "'")
    'busca
End Sub

Private Sub cmndbtnDet_Click()
    'frmEdiCai.Show 1
End Sub

Private Sub cmndbtnDetalhesBaixa_Click()
    If sstb.Tab = 0 Then
        P_ImprimeBaixaCarnesPorCaixa Left(cmbxCaixa.text, 2), txtbxData.text, txtbxData.text, ""
    Else
        P_ImprimeBaixaCarnesPorCaixa "", txtbxData.text, txtbxData.text, cmbxUsu.tag
    End If
    
End Sub

Private Sub cmndbtnFecharCaixa_Click()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim sCod_Con  As String
    Dim id As String
    If pcECF.BackColor = vbGreen Then
        MsgBox "Impossível fechar esse caixa. Observe os valores e as cores!", vbExclamation
        Exit Sub
    End If
    If sstb.Tab = 0 Then
'        If VerificaVendaAbertaNoCaixa(Left(cmbxCaixa.Text, 2)) Then
'            MsgBox "Não é possível fechar o caixa pois existem vendas em aberto! Procure o responsável!", vbCritical
'            Exit Sub
'
'        End If
    End If
    'On Error GoTo trata_erro
    strSQL = F_BuscaUmCampo("count(*)", "tab_vendas", "situacao_ven=2 and cod_cai='" & strCod_Cai & "'")
    If Int(strSQL) > 0 Then
        MsgBox "Não é possível fechar o caixa pois existem vendas a efetivar!", vbExclamation
        Exit Sub
    End If
'    If F_BuscaUmCampo("data_cai", "tab_fechacaixa", "data_cai=" & _
'        F_PesqData(txtbxData.Text) & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "'") <> "" Then
'        MsgBox "Este caixa já foi fechado neste dia!", vbExclamation
'        Exit Sub
'    End If

    If CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.text) < 0 Then
        MsgBox "O fundo de caixa está maior que o valor disponível!", vbExclamation
        Exit Sub
    End If

    If VendasFechadas = True Then
        MsgBox "Existem vendas condicionais a serem efetivadas!", vbExclamation
        On Error Resume Next
        frmEfetivaVen.Show 1
        Exit Sub
    End If
    If MsgBox("Deseja realmente fechar o Turno do dia " & txtbxData.text _
        & " Com " & txtbxFundoCaixa.text & " de Fundo de caixa?", vbYesNo + vbCritical) = vbYes Then
        If CCur(txtbxFundoCaixa.text) = 0 Then
            
            If MsgBox("Este fechamento não contém fundo de caixa!" & Chr(13) & _
                "Deseja realmente continuar?", vbYesNo + vbQuestion) = vbYesNo Then
                Exit Sub
            End If
        End If
        If F_BuscaUmCampo("status_tur", "tab_turncai", "cod_Cai='" & Left(cmbxCaixa.text, 2) & "' and cod_tur='" & Left(cmbxTurno.text, 1) & "' and data_tur=" & F_GravaData(txtbxData.text)) = "F" Then
            MsgBox "Este turno já se encontra Fechado!", vbCritical
            Exit Sub
         End If
'        strSQL = "insert into tab_fechacaixa(data_Cai,valor_FEC,fundo_cai,cod_usu,COD_CAI)values("
'        strSQL = strSQL & F_GravaData(txtbxData.Text) & ","
'        strSQL = strSQL & F_GravaDec(lblTotalAB.caption) & ","
'        strSQL = strSQL & F_GravaDec(txtbxFundoCaixa.Text) & ","
'        strSQL = strSQL & "'" & strCod_Usu & "','" & Left(cmbxCaixa.Text, 2) & "') "
'        conn.execute strSQL
        
        
        strSQL = "update tab_turncai set logout_usu='" & strCod_Usu & "',"
        strSQL = strSQL & "Fechamento_tur=" & F_GravaHora(Now) & ","
        strSQL = strSQL & "valorfechamento_tur=" & F_GravaDec(lblTotalAB.caption) & ","
        strSQL = strSQL & "status_tur='F' "
        strSQL = strSQL & " where cod_Cai='" & Left(cmbxCaixa.text, 2) & "' "
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.text, 1) & "' "
        Conn.Execute strSQL
        
        'strSQL = F_GravaCai("D", txtbxData.Text, CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.Text), _
            "TRANSFERÊNCIA DE CAIXA", "", "DIN", "01", 0, 0, Left(cmbxCaixa.Text, 2), strCod_Tur)
        
        id = F_GravaFinanceiro("C", txtbxData.text, CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.text), _
             "TRANSFERÊNCIA DE CAIXA", constContaCaixa, "S", "000")
         
        
            
            
        strSQL = "Update tab_Caixas set status_Cai='F' where cod_Cai='" & Left(cmbxCaixa.text, 2) & "'"
        Conn.Execute strSQL
        MsgBox "Caixa " & cmbxCaixa.text & "  Fechado com sucesso!", vbInformation
        
        
    End If
        
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnImprimir_Click()

        imprime frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            imprime Printer
            Printer.EndDoc
        End If
        Unload frmVisualiza
    
End Sub
Sub imprime(obj As Object)
    obj.FontName = "arial"
    Dim yIni As Single
    Dim yFin As Single
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim sHora As String
    
    
    
    
    
    P_CabRelTimbr "Fechamento de Caixa " & txtbxData.text & " " & cmbxCaixa.text & " Turno: " & cmbxTurno.text & " - (Usuário: " & strNome_Usu & ")", obj
    strSQL = "select now();"
    abreObjetoAdo rPesq, strSQL
    sHora = rPesq.Fields(0).Value
    
    
    rPesq.Close
    
    P_CentraImpressao obj, "Hora no Servidor " & sHora
    obj.Print ""
    
    
    
    
    obj.Print ""
    
    
    
    
    obj.Print ""
    
    yIni = obj.CurrentY
    P_Line obj, 8, obj.CurrentY, 182, obj.CurrentY
    P_ImprX obj, "Espécie: ", 10
    P_ImprX obj, "(C): ", 60, True, direita
    P_ImprX obj, "(D): ", 140, True, direita
    P_ImprX obj, "(Total): ", 180, False, direita
    P_Line obj, 8, obj.CurrentY, 182, obj.CurrentY
    
    P_ImprX obj, "Entrada ", 60, True, direita
    P_ImprX obj, "Recebimento ", 100, True, direita
    P_ImprX obj, "Saída: ", 140, True, direita
    P_ImprX obj, "(Saldo): ", 180, False, direita
    P_Line obj, 8, obj.CurrentY, 182, obj.CurrentY
    
    
    
    P_ImprX obj, "Dinheiro: ", 10
    P_ImprX obj, lblDin.caption, 60, True, direita
    P_ImprX obj, lblRec.caption, 100, True, direita
    P_ImprX obj, lblDinDeb.caption, 140, True, direita
    P_ImprX obj, lblTotalDin.caption, 180, False, direita
     
    
    P_ImprX obj, "Cheque À Vista: ", 10
    P_ImprX obj, lblChv.caption, 60, True, direita
    P_ImprX obj, lblRecCHV.caption, 100, True, direita
    P_ImprX obj, lblCHVDeb.caption, 140, True, direita
    P_ImprX obj, lblTotalCHV.caption, 180, False, direita
     
    obj.Print
    
    P_ImprX obj, "Cartão de Crédito: ", 10
    P_ImprX obj, lblFin.caption, 60, True, direita
    P_ImprX obj, lblRCC.caption, 100, True, direita
    P_ImprX obj, lblFinanDeb.caption, 140, True, direita
    P_ImprX obj, lblTotalRCC.caption, 180, False, direita
     
    
    P_ImprX obj, "Cartão de Débito: ", 10
    P_ImprX obj, lblCrd.caption, 60, True, direita
    P_ImprX obj, lblRCD.caption, 100, True, direita
    P_ImprX obj, lblSaiCRD.caption, 140, True, direita
    P_ImprX obj, lblTotalCRD.caption, 180, False, direita
     
    
    
    P_Line obj, 8, obj.CurrentY, 182, obj.CurrentY
    
    P_ImprX obj, "Total À Vista: ", 10
    P_ImprX obj, lblTotalEntr.caption, 60, True, direita
    P_ImprX obj, lblTotalRecAv.caption, 100, True, direita
    P_ImprX obj, lblTotalSai.caption, 140, True, direita
    P_ImprX obj, lblTotalAv.caption, 180, False, direita
     
    obj.Print
    
    
    P_ImprX obj, "Car/Duplicata: ", 10
    P_ImprX obj, lblCar.caption, 60, True, direita
    P_ImprX obj, lblRecCar.caption, 100, True, direita
    P_ImprX obj, lblCarDeb.caption, 140, True, direita
    P_ImprX obj, lblTotalCar.caption, 180, False, direita
     
    
    P_ImprX obj, "Cheque Pré: ", 10
    P_ImprX obj, lblCHP.caption, 60, True, direita
    P_ImprX obj, lblRecCHP.caption, 100, True, direita
    P_ImprX obj, lblCHPDeb.caption, 140, True, direita
    P_ImprX obj, lblTotalCHP.caption, 180, False, direita
     
    P_Line obj, 8, obj.CurrentY, 182, obj.CurrentY
     
    P_ImprX obj, "Total A Prazo: ", 10
    P_ImprX obj, lblTotalApCred.caption, 60, True, direita
    P_ImprX obj, lblTotalRecAP.caption, 100, True, direita
    P_ImprX obj, lblTotalApDeb.caption, 140, True, direita
    P_ImprX obj, lblTotalAp.caption, 180, False, direita
    obj.Print
    P_Line obj, 8, obj.CurrentY, 182, obj.CurrentY
    
    P_ImprX obj, "Total: ", 10
    P_ImprX obj, lblTotalEntradas.caption, 60, True, direita
    P_ImprX obj, lblTotalRec.caption, 100, True, direita
    P_ImprX obj, lblTotalSai2.caption, 140, True, direita
    P_ImprX obj, lblTotalAB.caption, 180, False, direita
      
     
      
      
     
    obj.FontBold = False
    yFin = obj.CurrentY
    P_Line obj, 8, obj.CurrentY, 182, obj.CurrentY
    P_Line obj, 8, yIni, 8, yFin
    P_Line obj, 182, yIni, 182, yFin
    obj.Print ""
    imprimeCheques obj
    imprimeCHP obj
    imprimeCAR obj
    imprimeFIN obj
    
    
    
    
End Sub

Sub imprimeCheques(obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select tc.num_ven,tx.* from tab_chv as tx,tab_caixa as tc "
    strSQL = strSQL & " where tc.seq_cai=tx.seq_cai and data_cai=" & F_GravaData(txtbxData.text)
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.text, 1)
    strSQL = strSQL & " and excluido_cai=false "
    strSQL = strSQL & " order by tc.seq_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Cheques à Vista"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "CPF", 20
        P_ImprX obj, "Titular", 50
        P_ImprX obj, "Nº Cheque", 110
        P_ImprX obj, "Agência", 135
        P_ImprX obj, "C/C", 155
        P_ImprX obj, "Banco", 173
        P_ImprX obj, "Valor", 198, False, direita
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !cpf_che, 20
            P_ImprX obj, Left(StrConv(!nome_che, vbProperCase), 40), 50
            P_ImprX obj, !num_che, 110
            P_ImprX obj, !age_che, 135
            P_ImprX obj, !cco_che, 155
            P_ImprX obj, !cod_bco, 173
            P_ImprX obj, FormatVal(!Valor_che), 198, False, direita
            
            cTotal = cTotal + !Valor_che
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 198, False, direita
        obj.Print ""
        
    End With


End Sub
Sub imprimeCHP(obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select tv.num_ven,tx.* from tab_cheques as tx,tab_venfec as tf,tab_doc as td,tab_vendas as tv "
    strSQL = strSQL & " where tv.num_ven=td.num_ven and dataefet_ven=" & F_GravaData(txtbxData.text)
    strSQL = strSQL & " and tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.text, 1)
    strSQL = strSQL & " and tx.seq_doc=td.seq_doc "
    strSQL = strSQL & " order by tv.num_ven "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Cheques a Prazo"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "CPF", 22
        P_ImprX obj, "Titular", 50
        P_ImprX obj, "Nº Cheque", 112
        P_ImprX obj, "Agência", 135
        P_ImprX obj, "C/C", 155
        P_ImprX obj, "Banco", 173
        P_ImprX obj, "Valor", 198, False, direita
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !cpf_che, 22
            P_ImprX obj, Left(StrConv(!nome_che, vbProperCase), 40), 50
            P_ImprX obj, !num_che, 112
            P_ImprX obj, !age_che, 135
            P_ImprX obj, !cco_che, 155
            P_ImprX obj, !cod_bco, 173
            P_ImprX obj, FormatVal(!Valor_che), 198, False, direita
            
            cTotal = cTotal + !Valor_che
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 198, False, direita
        
        
    End With


End Sub
Sub imprimeCAR(obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select tv.nomecli_ven,tv.num_ven,ti.* from tab_itdoc as ti,tab_venfec as tf,tab_doc as td,tab_vendas as tv "
    strSQL = strSQL & " where tv.num_ven=td.num_ven and dataefet_ven=" & F_GravaData(txtbxData.text)
    strSQL = strSQL & " and tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.text, 1)
    strSQL = strSQL & " and ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " order by tv.num_ven "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Vendas em Carteira"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "Cliente", 25
        P_ImprX obj, "Vencimento", 100
        P_ImprX obj, "Valor", 128, False, direita
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !nomecli_ven, 25
            P_ImprX obj, !venc_itd, 100
            P_ImprX obj, FormatVal(!valor_itd), 128, False, direita
            
            cTotal = cTotal + !valor_itd
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 128, False, direita
        
        
    End With


End Sub
Sub imprimeFIN(obj As Object)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select tv.nomecli_ven,tfi.nome_fin,tv.num_ven,ti.* from tab_venfin as ti,tab_venfec as tf,tab_vendas as tv, "
    strSQL = strSQL & " tab_financeiras as tfi where tv.num_ven=ti.num_ven and dataefet_ven=" & F_GravaData(txtbxData.text)
    strSQL = strSQL & " and ti.cod_fin=tfi.cod_fin "
    strSQL = strSQL & " and tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.text, 1)
    
    strSQL = strSQL & " order by tv.num_ven "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Vendas por Financeiras"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "Cliente", 25
        P_ImprX obj, "Financeira", 85
        P_ImprX obj, "Valor", 128, False, direita
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !nomecli_ven, 25
            P_ImprX obj, !nome_fin, 85
            P_ImprX obj, FormatVal(!valor_fin), 128, False, direita
            cTotal = cTotal + !valor_fin
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 128, False, direita
        
        
    End With


End Sub

Private Sub cmndbtnLan_Click()
    busca
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Command3_Click()
    
End Sub

Private Sub Form_Activate()
    If strVisrelfin_usu = "S" Then
        frControle.Visible = True
        pcRestr.Visible = False
    Else
        frControle.Visible = False
        pcRestr.Visible = True
    End If
    If strFechames_usu <> "S" Then
        cmbxCaixa.Enabled = False
        cmndbtnLan.Enabled = False
    End If
    busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        operacao = "F"
        pcECF.BackColor = vbGreen
        busca
    ElseIf KeyCode = vbKeyF9 Then
        operacao = "R"
        pcECF.BackColor = vbYellow
        busca
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtbxData.text = Date
    
    P_PreencheFunc cmbxUsu, 1
    P_PreencheCaixas cmbxCaixa
    P_PreencheTurnos cmbxTurno
    P_LocalizaOperacao strCod_Cai, cmbxCaixa
    P_LocalizaOperacao strCod_Tur, cmbxTurno
    pcECF.BackColor = vbGreen
    operacao = "F"
    pcECF.BackColor = vbYellow
    operacao = "R"
End Sub

Private Sub txtbxData_GotFocus()
    P_SelecTexto txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
    busca
End Sub
Sub buscaR()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    lblCar.caption = "0,00"
    lblCHP.caption = "0,00"
    lblChv.caption = "0,00"
    lblDin.caption = "0,00"
    lblFin.caption = "0,00"
    lblDinDeb.caption = "0,00"
    lblCHVDeb.caption = "0,00"
    lblFinanDeb.caption = "0,00"
    lblCarDeb.caption = "0,00"
    lblCHPDeb.caption = "0,00"
    lblTotalSai.caption = "0,00"
    lblRec.caption = "0,00"
    
    
    strSQL = "select tipo_cai,esp_cai,sum(tc.valor_cai) as total "
    strSQL = strSQL & " from tab_caixa as tc "
    strSQL = strSQL & " where tc.data_cai=" & F_PesqData(txtbxData.text)
    'strSQL = strSQL & " and tipo_cai='C' "
    strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    strSQL = strSQL & " and tc.excluido_cai=false "
    strSQL = strSQL & " group by esp_cai,tipo_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If .RecordCount > 0 Then
                If !esp_cai = "DIN" Then
                    If !tipo_cai = "C" Then
                        lblDin.caption = FormatVal(!total)
                    Else
                        lblDinDeb.caption = FormatVal(!total)
                    End If
                    
                End If
                
                If !esp_cai = "MOV" Then
                    If !tipo_cai = "C" Then
                        lblAvulsoCred.caption = FormatVal(!total)
                    Else
                        lblAvulsoDeb.caption = FormatVal(!total)
                    End If
                    
                End If
                
                If !esp_cai = "RDI" Then
                    lblRec.caption = FormatVal(!total)
                End If
                If !esp_cai = "FIN" Or !esp_cai = "CRT" Then
                    If !tipo_cai = "C" Then
                        lblFin.caption = FormatVal(!total)
                    Else
                        lblFinanDeb.caption = FormatVal(!total)
                    End If
                End If
                If !esp_cai = "CRD" Then
                    If !tipo_cai = "C" Then
                        lblCrd.caption = FormatVal(!total)
                    Else
                        lblSaiCRD.caption = FormatVal(!total)
                    End If
                End If
                
                
                If !esp_cai = "CHP" Then
                    If !tipo_cai = "C" Then
                        lblCHP.caption = FormatVal(!total)
                    Else
                        lblCHPDeb.caption = FormatVal(!total)
                    End If
                End If
                If !esp_cai = "CHV" Then
                    If !tipo_cai = "C" Then
                        lblChv.caption = FormatVal(!total)
                    Else
                        lblCHVDeb.caption = FormatVal(!total)
                    End If
                End If
                If !esp_cai = "CAR" Or !esp_cai = "DUP" Then
                    If !tipo_cai = "C" Then
                        lblCar.caption = FormatVal(!total)
                    Else
                        lblCarDeb.caption = FormatVal(!total)
                    End If
                End If
                If !esp_cai = "RCH" Then
                    If !tipo_cai = "C" Then
                        lblRecCHP.caption = FormatVal(!total)
                    Else
                        'lblMultideb.caption = FormatVal(!total)
                    End If
                End If
                If !esp_cai = "RDI" Then
                    If !tipo_cai = "C" Then
                        lblRec.caption = FormatVal(!total)
                    Else
                        'lblMultideb.caption = FormatVal(!total)
                    End If
                End If
                If !esp_cai = "RCD" Then
                    If !tipo_cai = "C" Then
                        lblRCD.caption = FormatVal(!total)
                    Else
                        'lblMultideb.caption = FormatVal(!total)
                    End If
                End If
                If !esp_cai = "RCC" Then
                    If !tipo_cai = "C" Then
                        lblRCC.caption = FormatVal(!total)
                    Else
                        'lblMultideb.caption = FormatVal(!total)
                    End If
                End If
                
            End If
            
            .MoveNext
        Loop
        
        lblTotalAv.caption = FormatVal(CCur(lblDin.caption) + CCur(lblChv.caption) + CCur(lblFin.caption) + CCur(lblRec.caption))
        lblTotalAp.caption = FormatVal(CCur(lblCHP.caption) + CCur(lblCar.caption))
        lblTotalEntradas.caption = FormatVal(CCur(lblTotalAv.caption) + CCur(lblTotalAp.caption))
        
        
        
        .Close
        
    End With
    
    'strSQL = "select tf.* from tab_financeiro as tf, tab_caixas as tc where tc.cod_con=tf.cod_con "
    'strSQL = strSQL & " and tc.cod_cai='" & strCod_Cai & "' "
    'strsql = strsql & " and data_
    
    
    
    strSQL = "select sum(tc.valor_cai) as total from tab_Caixa as tc "
    strSQL = strSQL & " where tipo_cai='D' and data_cai=" & F_PesqData(txtbxData.text)
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    strSQL = strSQL & " and tc.excluido_cai=false "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            lblTotalSai.caption = FormatVal(!total)
        End If
        .Close
    End With
    lblTotalAB.caption = FormatVal(CCur(lblTotalEntradas.caption) - CCur(lblTotalSai.caption))
    lblTotalDin.caption = FormatVal(CCur(lblDin.caption) - CCur(lblTotalSai.caption))
    
    
    
    lblTotalSai2.caption = lblTotalSai.caption
    'parei aki
    strSQL = "select sum(tc.valor_ven) as total from tab_cancelamentos  as tc,"
    strSQL = strSQL & " tab_venfec as tf "
    strSQL = strSQL & " where tc.num_ven=tf.num_ven "
    strSQL = strSQL & " and left(data_can,10) = " & F_PesqData(txtbxData.text)
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
    strSQL = strSQL & " and tf.cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            lblCancela.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    strSQL = "select sum(tc.valor_cai) as total from tab_caixa as tc, tab_valescai as tv "
    strSQL = strSQL & "  where data_cai = " & F_PesqData(txtbxData.text)
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    strSQL = strSQL & " and tc.seq_cai=tv.seq_cai "
    strSQL = strSQL & " and tc.excluido_cai=false "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            lblSaidasFun.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    
    soma
    
End Sub
Sub soma()
    lblTotalRecAv.caption = FormatVal(CCur(lblRec.caption) + CCur(lblRecCHV.caption) + CCur(lblRCD.caption) + CCur(lblRCC.caption))
    lblTotalDin.caption = FormatVal(CCur(lblDin.caption) + CCur(lblRec.caption) - CCur(lblDinDeb.caption))
    lblTotalCHV.caption = FormatVal(CCur(lblChv.caption) - CCur(lblCHVDeb.caption))
    lblTotalRCC.caption = FormatVal(CCur(lblFin.caption) + CCur(lblRCC.caption) - CCur(lblFinanDeb.caption))
    lblTotalEntr.caption = FormatVal(CCur(lblChv.caption) + CCur(lblFin.caption) + CCur(lblDin.caption) + CCur(lblCrd.caption))
    lblTotalSai.caption = FormatVal(CCur(lblCHVDeb.caption) + CCur(lblFinanDeb.caption) + CCur(lblDinDeb))
    lblTotalAv.caption = FormatVal(CCur(lblTotalEntr.caption) + CCur(lblTotalRecAv.caption) - CCur(lblTotalSai.caption))
    lblTotalCar.caption = FormatVal(CCur(lblCar.caption) - CCur(lblCarDeb.caption))
    lblTotalCHP.caption = FormatVal(CCur(lblCHP.caption) - CCur(lblCHPDeb.caption))
    lblTotalApCred.caption = FormatVal(CCur(lblCHP.caption) + CCur(lblCar.caption))
    lblTotalApDeb.caption = FormatVal(CCur(lblCHPDeb.caption) + CCur(lblCarDeb.caption))
    lblTotalAp.caption = FormatVal(CCur(lblTotalApCred.caption) - CCur(lblTotalApDeb.caption))
    lblTotalCRD.caption = FormatVal(CCur(lblCrd.caption) + CCur(lblRCD.caption) - CCur(lblSaiCRD.caption))
        
    lblTotalEntradas.caption = FormatVal(CCur(lblTotalEntr.caption) + CCur(lblTotalApCred.caption))
        
        
    lblTotalRec.caption = FormatVal(CCur(lblTotalRecAv.caption) + CCur(lblTotalRecAP.caption))
    lblTotalAB.caption = FormatVal(CCur(lblTotalEntradas.caption) + CCur(lblTotalRec.caption) - CCur(lblTotalSai2.caption))
        
    
End Sub
    
Sub busca()
    If operacao = "F" Then
        buscaF
    Else
        buscaR
    End If
End Sub
Sub buscaF()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    lblCar.caption = "0,00"
    lblCHP.caption = "0,00"
    lblChv.caption = "0,00"
    lblDin.caption = "0,00"
    lblFin.caption = "0,00"
    lblTotalSai.caption = "0,00"
    strSQL = "select esp_cai,sum(tc.valor_cai) as total "
    strSQL = strSQL & " from tab_caixa as tc,"
    strSQL = strSQL & " tab_saidas as ts "
    strSQL = strSQL & " where ts.num_ven=tc.num_ven"
    strSQL = strSQL & " and tipo_cai='C' "
    strSQL = strSQL & " and serie_sai='CUPM' "
    strSQL = strSQL & " and tc.excluido_cai=false "
    If sstb.Tab = 0 Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    Else
        strSQL = strSQL & " and tc.cod_usu='" & cmbxUsu.tag & "' "
    End If
    strSQL = strSQL & " and data_cai=" & F_GravaData(txtbxData.text)
    strSQL = strSQL & " group by esp_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If .RecordCount > 0 Then
                If !esp_cai = "DIN" Then
                    lblDin.caption = FormatVal(!total)
                End If
                If !esp_cai = "FIN" Then
                    lblFin.caption = FormatVal(!total)
                End If
                If !esp_cai = "CHP" Then
                    lblCHP.caption = FormatVal(!total)
                End If
                If !esp_cai = "CHV" Then
                    lblChv.caption = FormatVal(!total)
                End If
                If !esp_cai = "CAR" Then
                    lblCar.caption = FormatVal(!total)
                End If
            End If
            
            .MoveNext
        Loop
        lblTotalAv.caption = FormatVal(CCur(lblDin.caption) + CCur(lblChv.caption) + CCur(lblFin.caption))
        lblTotalAp.caption = FormatVal(CCur(lblCHP.caption) + CCur(lblCar.caption))
        lblTotalEntradas.caption = FormatVal(CCur(lblTotalAv.caption) + CCur(lblTotalAp.caption))
        .Close
        
    End With
    strSQL = "select sum(tc.valor_cai) as total from tab_Caixa as tc "
    strSQL = strSQL & " where tipo_cai='D' and data_cai=" & F_PesqData(txtbxData.text)
    If sstb.Tab = 0 Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    Else
        strSQL = strSQL & " and tc.cod_usu='" & cmbxUsu.tag & "' "
    End If
    strSQL = strSQL & " and excluido_cai=false "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            lblTotalSai.caption = FormatVal(!total)
        End If
        .Close
    End With
    lblTotalAB.caption = FormatVal(CCur(lblTotalEntradas.caption) - CCur(lblTotalSai.caption))
    lblTotalDin.caption = FormatVal(CCur(lblDin.caption) - CCur(lblTotalSai.caption))
    lblTotalSai2.caption = lblTotalSai.caption
    'parei aki
    strSQL = "select sum(tc.valor_ven) as total from tab_cancelamentos  as tc "
    
    strSQL = strSQL & " where "
    strSQL = strSQL & " left(data_can,10) = " & F_PesqData(txtbxData.text)
    If sstb.Tab = 0 Then
        strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.text, 2) & "' "
        strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.text, 1) & "' "
    Else
        strSQL = strSQL & " and tc.cod_usu='" & cmbxUsu.tag & "' "
    End If
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!total) Then
            lblCancela.caption = FormatVal(!total)
        End If
        .Close
    End With
End Sub

Private Sub txtbxFundoCaixa_GotFocus()
    P_SelecTexto txtbxFundoCaixa
End Sub

Private Sub txtbxFundoCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxFundoCaixa_LostFocus()
    P_ValorLostFocus txtbxFundoCaixa
End Sub
