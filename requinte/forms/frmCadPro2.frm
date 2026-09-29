VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadPro2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Produtos:"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10845
   Icon            =   "frmCadPro2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFracionado 
      BackColor       =   &H007DBFDB&
      Caption         =   "&Fracionado"
      Height          =   255
      Left            =   2880
      TabIndex        =   80
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtbxICMS 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Tag             =   "01"
      Text            =   "17%"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10845
      TabIndex        =   68
      Top             =   7335
      Width           =   10845
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   2280
         TabIndex        =   79
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblNumNot 
         AutoSize        =   -1  'True
         Caption         =   "00000000"
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
         Left            =   3000
         TabIndex        =   78
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblEmb 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   3240
         TabIndex        =   77
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embalagem:"
         Height          =   195
         Left            =   2280
         TabIndex        =   76
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   5040
         TabIndex        =   75
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Unitário:"
         Height          =   195
         Left            =   3840
         TabIndex        =   74
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblQtd 
         AutoSize        =   -1  'True
         Caption         =   "0,00"
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
         Left            =   1320
         TabIndex        =   73
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd Comprada:"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblNomeFor 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Fornecedor"
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
         Left            =   3840
         TabIndex        =   71
         Top             =   0
         Width           =   1740
      End
      Begin VB.Label lblUltComp 
         AutoSize        =   -1  'True
         Caption         =   "00/00/0000"
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
         Left            =   1320
         TabIndex        =   70
         Top             =   0
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Última Compra:"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10845
      TabIndex        =   49
      Top             =   0
      Width           =   10845
      Begin VB.Label lblNomePro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Novo Produto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   240
         TabIndex        =   50
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   7440
      ScaleHeight     =   2385
      ScaleWidth      =   3345
      TabIndex        =   58
      Top             =   2040
      Width           =   3375
      Begin MSFlexGridLib.MSFlexGrid msflxgrdPrecos 
         Height          =   2415
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4260
         _Version        =   393216
         BackColor       =   -2147483637
         BackColorFixed  =   8454143
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         GridLinesFixed  =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Códigos de &Barra:"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   7440
      TabIndex        =   51
      Top             =   360
      Width           =   3375
      Begin VB.ListBox lstCodBarra 
         Appearance      =   0  'Flat
         Height          =   810
         ItemData        =   "frmCadPro2.frx":0ECA
         Left            =   120
         List            =   "frmCadPro2.frx":0ECC
         TabIndex        =   55
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtbxCodBarra 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton cmndbtnExclBarcode 
         Height          =   315
         Left            =   3000
         Picture         =   "frmCadPro2.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1320
         Width           =   315
      End
      Begin VB.CommandButton cmndbtnAddBarcode 
         Height          =   315
         Left            =   2640
         Picture         =   "frmCadPro2.frx":1458
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adicionar Barcode:"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   1350
      End
   End
   Begin VB.TextBox txtbxQtdEmb 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   14
      Text            =   "1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtbxPFab 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   18
      Text            =   "0,00"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtbxNome_Pro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   5895
   End
   Begin VB.TextBox txtbxRef 
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
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtbxLocal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtbxMargem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   16
      Text            =   "30"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   7065
      TabIndex        =   24
      Top             =   3720
      Width           =   7095
      Begin VB.TextBox txtbxFabricante 
         Appearance      =   0  'Flat
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
         Left            =   1200
         TabIndex        =   26
         Top             =   120
         Width           =   5295
      End
      Begin VB.CommandButton cmndbtnBuscaFab 
         Height          =   315
         Left            =   6480
         Picture         =   "frmCadPro2.frx":19E2
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fabricante"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   25
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   10545
      TabIndex        =   27
      Top             =   4560
      Width           =   10575
      Begin VB.TextBox txtbxCB_Pro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7920
         TabIndex        =   35
         Text            =   "0,00"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxQueima 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H007DBFDB&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6360
         TabIndex        =   33
         Text            =   "0,00"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkQueima 
         BackColor       =   &H007DBFDB&
         Caption         =   "&Queima"
         Height          =   255
         Left            =   6360
         TabIndex        =   32
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtbxPa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   1680
         TabIndex        =   57
         Text            =   "0,00"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxPv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3240
         TabIndex        =   29
         Text            =   "0,00"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxCr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4800
         TabIndex        =   31
         Text            =   "0,00"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo Bruto:"
         Height          =   195
         Left            =   7920
         TabIndex        =   34
         Top             =   0
         Width           =   870
      End
      Begin VB.Label lblEr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Esto&que:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Width           =   630
      End
      Begin VB.Label lblPrecoCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Preço Varejo:"
         Height          =   195
         Left            =   3240
         TabIndex        =   28
         Top             =   0
         Width           =   960
      End
      Begin VB.Label lblestCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço &Atacado:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1680
         TabIndex        =   44
         Top             =   0
         Width           =   1110
      End
      Begin VB.Label lblCusCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&usto:"
         Height          =   195
         Left            =   4800
         TabIndex        =   30
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   10845
      TabIndex        =   38
      Top             =   7830
      Width           =   10845
      Begin VB.CommandButton cmndbtnNovo 
         Caption         =   "&Novo"
         Height          =   855
         Left            =   5520
         Picture         =   "frmCadPro2.frx":1ADB
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "&Salvar F7"
         Height          =   855
         Left            =   7200
         Picture         =   "frmCadPro2.frx":1BCF
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar (ESC)"
         Height          =   855
         Left            =   8880
         Picture         =   "frmCadPro2.frx":1D2B
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   1575
      End
      Begin VB.PictureBox pcInfo 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   2415
         TabIndex        =   41
         Top             =   120
         Width           =   2415
         Begin VB.PictureBox pcEnt 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "frmCadPro2.frx":2279
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   43
            Top             =   67
            Width           =   240
         End
         Begin VB.PictureBox pcSai 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            Picture         =   "frmCadPro2.frx":23C3
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   42
            Top             =   420
            Width           =   240
         End
         Begin VB.Label lblInfoEnt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações de Entrada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   39
            Top             =   82
            Width           =   1755
         End
         Begin VB.Label lblInfoSai 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Informações de Saída"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   40
            Top             =   435
            Width           =   1605
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   7095
      Begin VB.ComboBox cmbxSecao 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   5625
      End
      Begin VB.ComboBox cmbxSubSec 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   720
         Width           =   5625
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seção:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   80
         TabIndex        =   20
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Subseção:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   80
         TabIndex        =   22
         Top             =   720
         Width           =   1125
      End
   End
   Begin VB.TextBox txtbxUnidade 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      Text            =   "UN"
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbxSt 
      Height          =   315
      ItemData        =   "frmCadPro2.frx":250D
      Left            =   1320
      List            =   "frmCadPro2.frx":251A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Similares"
      Height          =   1815
      Left            =   120
      TabIndex        =   61
      Top             =   5400
      Width           =   10575
      Begin VB.Data dt 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmCadPro2.frx":2539
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmCadPro2.frx":254A
         TabIndex        =   67
         Top             =   600
         Width           =   10215
      End
      Begin VB.CommandButton cmndbtnAdicionar 
         Caption         =   "Adicionar"
         Height          =   315
         Left            =   6360
         TabIndex        =   66
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxSimilar 
         Appearance      =   0  'Flat
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
         TabIndex        =   62
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton cmndbtnBuscaSim 
         Height          =   315
         Left            =   6000
         Picture         =   "frmCadPro2.frx":3287
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblCodProSim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
         Height          =   195
         Left            =   8880
         TabIndex        =   65
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmndbtnICMS 
      Height          =   315
      Left            =   2400
      Picture         =   "frmCadPro2.frx":3380
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblProvavel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(Provável)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3960
      TabIndex        =   60
      Top             =   840
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qtd. Emb:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5145
      TabIndex        =   13
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pç. Fab:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5145
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&ICMS:"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   185
      TabIndex        =   9
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sit. Trib.:"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   185
      TabIndex        =   6
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Nome:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   185
      TabIndex        =   0
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Referência.:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   185
      TabIndex        =   2
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Localiz.:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   185
      TabIndex        =   4
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M.&Luc(%):"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5145
      TabIndex        =   15
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Unidade:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5145
      TabIndex        =   11
      Top             =   840
      Width           =   840
   End
End
Attribute VB_Name = "frmCadPro2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Dim cPreco As String 'para carregar o preço atual no caso de mudança de preço


Private Sub chkProd_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub



Private Sub chkBonif_Click()
    If chkBonif.Value = 1 Then
        txtbxBonif.Enabled = True
        txtbxBonif.SetFocus
    Else
        txtbxBonif.Enabled = False
    End If
    
End Sub

Private Sub chkEstoque_Click()
    If Me.tag <> "" Then
        
    End If
End Sub

Private Sub cmbxControle_Change()

End Sub

Private Sub cmbxControle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub chkQueima_Click()
    If chkQueima.Value = 1 Then
        txtbxQueima.Enabled = True
        txtbxQueima.SetFocus
    Else
        txtbxQueima.Enabled = False
        txtbxQueima.Text = "0,00"
    End If
End Sub

Private Sub cmbxICMS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxMarca_Click()
    lblMar.caption = F_BuscaUmCampo("Cod_MAR", "tab_marcas", "nome_mar='" & cmbxMarca.Text & "'")
End Sub

Private Sub cmbxEsp_Click()
    lblEsp.caption = F_BuscaUmCampo("Cod_Esp", "tab_Especies", "nome_Esp='" & cmbxEsp.Text & "'")
End Sub

Private Sub cmbxCat_Click()
    lblCat.caption = F_BuscaUmCampo("Cod_Cat", "tab_Categorias", "nome_Cat='" & cmbxCat.Text & "'")
    
End Sub


Private Sub cmbxList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxPrincipio_Click()
    cmbxPrincipio.tag = F_BuscaUmCampo("cod_pri", "tab_principios", "nome_pri='" & cmbxPrincipio.Text & "'")
    
End Sub

Private Sub cmbxPrincipio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
    End If
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSecao_Click()
    cmbxSecao.tag = F_BuscaUmCampo("cod_sec", "tab_secoes", "nome_sec='" & cmbxSecao.Text & "'")
    
    P_PreencheSubSec cmbxSubSec, cmbxSecao.tag
End Sub

Private Sub cmbxSecao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        P_PreencheSecao cmbxSecao
        
    End If
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSt_Click()
    On Error Resume Next
    If Left(cmbxSt.Text, 1) = "R" Then
        txtbxIcms.Text = "Retido"

        txtbxIcms.tag = "FF"
        txtbxIcms.Enabled = False
    ElseIf Left(cmbxSt.Text, 1) = "I" Then
        txtbxIcms.Text = "Isento"
        txtbxIcms.tag = "II"
        txtbxIcms.Enabled = False
        
    Else
        
        txtbxIcms.Enabled = True
        txtbxIcms.Text = "17%"
        txtbxIcms.tag = "01"
        txtbxIcms.SetFocus
    End If
End Sub

Private Sub cmbxSt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSt_LostFocus()
    On Error Resume Next
    
End Sub


Private Sub cmndbtnCat_Click()
    buscaCat
End Sub


Private Sub cmndbtnAddBarcode_Click()
    If txtbxCodBarra.Text = "" Then
        MsgBox "Informe os dígitos do código de barras!", vbExclamation
        txtbxCodBarra.SetFocus
        Exit Sub
    End If
    
    If Me.tag = "" Then
        lstCodBarra.AddItem txtbxCodBarra.Text
    Else
        If Produtos.F_AddCodBar(Me.tag, txtbxCodBarra.Text) Then
            P_PreencheBarCode Me.tag, lstCodBarra
        End If
    End If
End Sub

Private Sub cmndbtnAdicionar_Click()
    If Me.tag = "" Then
        MsgBox "O produto não está gravado!", vbExclamation
        Exit Sub
    End If
    If Me.tag = txtbxSimilar.tag Then
        MsgBox "Os produtos não podem ser iguais!", vbExclamation
        Exit Sub
    End If
    If txtbxSimilar.tag = "" Then
        MsgBox "Informe o produto que deseja adicionar!", vbExclamation
        Exit Sub
    End If
    P_GravaSimilar Me.tag, txtbxSimilar.tag
    buscaSim
End Sub

Private Sub cmndbtnBuscaSim_Click()
    Dim p As Produto
    p = F_BuscaProdCabecalho(txtbxSimilar.Text)
    If p.cod_pro <> "" Then
        lblCodProSim.caption = p.cod_pro
        txtbxSimilar.Text = p.nome_pro
        txtbxSimilar.tag = p.cod_pro
    End If
End Sub

Private Sub cmndbtnExclBarcode_Click()
    Dim strSQL As String
    If lstCodBarra.ListIndex = -1 Then
        MsgBox "É necessário que se escolha um código de barras!", vbCritical
        lstCodBarra.SetFocus
        Exit Sub
    End If
    If lstCodBarra.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstCodBarra.ListCount = 1 Then
        MsgBox "Este código de barras não pode ser excluído por ser o único!", vbCritical
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o código de barras?", vbYesNo + vbCritical) = vbYes Then
        If Me.tag = "" Then
            lstCodBarra.RemoveItem lstCodBarra.ListIndex
        Else
            strSQL = "delete from tab_codbarras where cod_pro='" & Me.tag & "' "
            strSQL = strSQL & " and cod_bar='" & lstCodBarra.Text & "' "
            Banco.Execute strSQL
            buscaCodBar
        End If
    End If
End Sub

Private Sub cmndbtnMarcas_Click()
    buscaMarca
End Sub

Private Sub cmndbtnRefresh_Click()
    RefresheComp
End Sub

Private Sub cmndbtnExclEmb_Click()
    If lstEmb.ListCount = 0 Then
        MsgBox "Não há o que excluir!", vbExclamation
        Exit Sub
    End If
    If lstEmb.ListIndex = -1 Then
        MsgBox "É necessário que se escolha m item!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir a embalagem?", vbYesNo + vbCritical) = vbYes Then
        lstEmb.RemoveItem lstEmb.ListIndex
    End If
End Sub


Private Sub cmndbtnFoto_click()
    frmFoto.caption = txtbxNome_Pro.Text
    frmFoto.tag = Me.tag
    frmFoto.Show 1
End Sub

Private Sub cmbxSubSec_Click()
    cmbxSubSec.tag = F_BuscaUmCampo("cod_sub", "tab_subsec", "nome_sub='" & cmbxSubSec.Text & "' and cod_sec='" & cmbxSecao.tag & "'")
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxSubSec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxFabricante.SetFocus
    End If
End Sub



Private Sub cmndbtnBuscaFab_Click()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_SelectForn(txtbxFabricante.Text)
    If f.cod_for <> "" Then
        txtbxFabricante.Text = f.fantasia_for
        txtbxFabricante.tag = f.cod_for
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscarControl_Click()
    Dim c As controle
    c = Produtos.F_BuscaProdControle(txtbxControle.Text)
    If c.cod_con <> "" Then
        txtbxControle.Text = c.nome_con
        txtbxControle.tag = c.cod_con
    End If
    
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    If strExclVenda_usu <> "S" Then
        MsgBox "Você não tem permissão para isto!", vbExclamation
        Exit Sub
    End If
    With dtPro.Recordset
        If MsgBox("Confirma exclusão do lote? ", vbYesNo + vbCritical) = vbYes Then
            strSQL = F_BuscaUmCampo("num_ven", "tab_itvendas as ti, tab_vendas as tv", _
                "tv.num_ven=ti.nuM_ven and ti.seq_pro='" & !seq_pro & "' and situacao_ven<>'3' and cod_pro='" & !cod_pro & "'")
            If strSQL = "" Then
                strSQL = "delete from tab_estoque where cod_pro='" & !cod_pro & "' and seq_pro=" & !seq_pro
                Banco.Execute strSQL
                
            Else
                MsgBox "O produto não pode ser excluído!", vbExclamation
                Exit Sub
            End If
                
        End If
    End With
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnMatriz_Click()
    Dim p As Produtos.Produto
    If txtbxMatriz.Text = "" Then
        MsgBox "Informe o nome do produto!", vbExclamation
        Exit Sub
    End If
    p = F_BuscaProdCabecalho(txtbxMatriz.Text)
    If p.cod_pro <> "" Then
        txtbxMatriz.tag = p.cod_pro
        txtbxMatriz.Text = p.nome_pro
        lblPVMatriz.caption = FormatVal(p.pv_pro)
        lblCustoMatriz.caption = FormatVal(p.cr_pro)
        'cmbxPrincipio.Text = p.prin.nome_pri
        'cmbxSecao.Text = p.sub.sec.nome_sec
        'cmbxSubSec.Text = p.sub.nome_sub
    End If
    
End Sub

Private Sub cmndbtnICMS_Click()
    buscaImp
End Sub
Sub buscaImp()
    Dim i As imposto.imposto
    
    i = imposto.F_SelectImp(txtbxIcms.Text)
    If i.cod_imp <> "" Then
        txtbxIcms.Text = i.desc_imp
        txtbxIcms.tag = i.cod_imp
        If i.cod_imp = "II" Then
            cmbxSt.Text = "Isenta"
        ElseIf i.cod_imp = "FF" Then
            cmbxSt.Text = "Retida"
        End If
        cmbxSecao.SetFocus
    End If
    
End Sub
Sub buscaICMS()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_impostos where cod_imp='" & txtbxIcms.tag & "'"
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            txtbxIcms.Text = !desc_imp
            txtbxIcms.tag = !cod_imp
        End If
        .Close
        
    End With
End Sub
Private Sub cmndbtnNovo_Click()
    Dim fr As New frmCadPro
    fr.tag = ""
    fr.Show 0
    Unload Me
End Sub

Private Sub cmndbtnPro_Click()
    frmProdProm.Show 0
End Sub

Private Sub cmndbtnQueima_Click()
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
On Error GoTo trata_Erro

    If txtbxIcms.tag = "" Then
        MsgBox "Informe o ICMS!", vbExclamation
        Exit Sub
    End If
    If cmbxSubSec.tag = "" Then
        MsgBox "Informe a subseção!", vbExclamation
        cmbxSubSec.SetFocus
        Exit Sub
    End If
    
    
    If Me.tag = "" Then
        strSQL = F_BuscaUmCampo("nome_pro", "tab_produtos", "NOME_PRO='" & _
            txtbxNome_Pro.Text & "'")
        If strSQL <> "" Then
            MsgBox "Já existe um produto com este nome!", vbCritical
            txtbxNome_Pro.SetFocus
            Exit Sub
        End If
        If txtbxRef.Text <> "" Then
        
            strSQL = F_BuscaUmCampo("nome_pro", "tab_produtos", "codfab_pro='" & _
            txtbxRef.Text & "'")
            
        End If
            If strSQL <> "" Then
                If MsgBox("Já existe um produto com esta referência!" & _
                     Chr(13) & strSQL & Chr(13) & "Deseja continuar?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
                End If
            End If
            
        If cmbxSt.Text = "" Then
            MsgBox "Informe a Situacação Tributária!", vbCritical
            cmbxSt.SetFocus
            Exit Sub
        End If
        If MsgBox("confirma inclusão do Produto?", vbYesNo + vbQuestion) = vbYes Then
            
            Me.tag = Produtos.F_GravaProduto(txtbxNome_Pro.Text, cmbxSubSec.tag, _
                  txtbxRef.Text, txtbxCr.Text, txtbxCr.Text, txtbxPv.Text, txtbxPa.Text, txtbxUnidade.Text, _
                    txtbxPFab.Text, txtbxFabricante.tag, txtbxLocal.Text, Left(cmbxSt.Text, 1), txtbxIcms.tag, txtbxMargem.Text, txtbxCB_Pro.Text, txtbxQueima.Text)
            
            If Me.tag <> "" Then
                MsgBox "Produto Gravado com o código " & Me.tag, vbInformation
                
                For iCount = 0 To lstCodBarra.ListCount - 1
                    lstCodBarra.ListIndex = iCount
                    If Produtos.F_AddCodBar(Me.tag, lstCodBarra.Text) Then
                    End If
                Next
            End If
            
            
        
        End If
    Else
        If MsgBox("Confirma alteração do Produto?", vbYesNo + vbQuestion) = vbYes Then
            If Produtos.F_EditaProduto(Me.tag, txtbxNome_Pro.Text, cmbxSubSec.tag, txtbxRef.Text, _
                txtbxCr.Text, txtbxPv.Text, txtbxPa.Text, txtbxUnidade.Text, _
                txtbxPFab.Text, txtbxFabricante.tag, txtbxLocal.Text, Left(cmbxSt.Text, 1), txtbxIcms.tag, txtbxCB_Pro.Text, _
                    txtbxQueima.Text, IIf(chkFracionado.Value = 1, "S", "N")) Then
                MsgBox "Produto alterado com sucesso!", vbInformation
                P_AltPreco Me.tag, CCur(txtbxPv.Text), CCur(txtbxPv.tag)
            End If
        End If
    End If
        
        Exit Sub
trata_Erro:
        P_MostraErro "Cadastro de Produto"
End Sub
Sub P_Altera()
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "update tab_produtos set "
    strSQL = strSQL & "nome_pro='" & txtbxNome_Pro.Text & "',"
    strSQL = strSQL & "unidade_pro='" & txtbxUnidade.Text & "',"
    strSQL = strSQL & "codfab_pro='" & txtbxRef.Text & "', "
    strSQL = strSQL & "balanca_pro='" & IIf(chkProd.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "cod_mar='" & lblMar.caption & "', "
    strSQL = strSQL & "cod_cat='" & lblCat.caption & "', "
    
    strSQL = strSQL & "endereco_pro='" & txtbxLocal.Text & "',"
    strSQL = strSQL & "lucro_pro=" & F_GravaDec(txtbxMargem.Text) & ",  "
    strSQL = strSQL & "pv_pro=" & F_GravaDec(txtbxPreco.Text) & ","
    strSQL = strSQL & "cr_pro=" & F_GravaDec(txtbxCusto.Text) & ","
   ' If strAltpreco_usu = "S" Then
  '      strSQL = strSQL & "er_pro=" & F_GravaDec(txtbxEstoque.Text) & ","
   ' End If
    strSQL = strSQL & "estmin_pro=" & F_GravaDec(txtbxEstMin.Text) & ","
    strSQL = strSQL & "st_pro='" & Left(cmbxSt.Text, 1) & "', "  ' situacao trib"
    strSQL = strSQL & "cod_imp='" & Left(cmbxICMS.Text, 3) & "'  "
    strSQL = strSQL & " where cod_pro='" & Me.tag & "' "
    Banco.Execute strSQL
    P_AltPreco Me.tag, CCur(txtbxPreco.Text), CCur(cPreco)
    strSQL = "delete from tab_codbarras where cod_pro='" & Me.tag & "' "
    Banco.Execute strSQL
    With lstCodBarra
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values("
            strSQL = strSQL & "'" & Me.tag & "',"
            strSQL = strSQL & "'" & .Text & "')"
            Banco.Execute strSQL
        Next
    End With
    gravaEmb
    
    
   ' strSQL = F_GeraLog(strSQL, "tab_produtos")
    MsgBox "Produto Alterado com Sucesso!", vbInformation
    Unload Me
    On Error Resume Next
    If IsLoad("frmEdiPro") Then
        RefreshDataPosition frmEdiPro.dtPro
    End If
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub
Sub gravaEmb()
    Dim strSQL As String
    strSQL = "delete from tab_embalagens where cod_pro='" & Me.tag & "' "
    Banco.Execute strSQL
    With lstEmb
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_embalagens(cod_pro,qtd_emb)values("
            strSQL = strSQL & "'" & Me.tag & "',"
            strSQL = strSQL & .Text & ")"
            Banco.Execute strSQL
        Next
    End With
End Sub
Sub P_Grava()
On Error GoTo trata_Erro

    Dim strSQL As String
    Dim sCod As String
    Dim sMar As String
    Dim iCount As Integer
    
    sCod = F_BuscaUmCampo(" max(cod_pro)", "tab_produtos", "cod_pro like '" & lblCat.caption & "%'")
    If sCod = "" Then
        sCod = "0001"
    Else
        
        sCod = Format(Int(Mid(sCod, 4, 4)) + 1, "000#")
    End If
    sCod = F_Ean8(lblCat.caption & sCod)
    If lstCodBarra.ListCount = 0 Then
'        If MsgBox("Este produto não possui um código de barras!" & _
            Chr(13) & "Deseja que o Sistema gere um automaticamente?", vbYesNo + _
                vbCritical) = vbYes Then
            lstCodBarra.AddItem sCod
 '       End If
    End If
    strSQL = "insert into tab_produtos(cod_pro,nome_pro,unidade_pro,"
    strSQL = strSQL & "cod_cat,cod_mar,codfab_pro,"
    strSQL = strSQL & "lucro_pro,"
    strSQL = strSQL & "endereco_pro,er_pro,ef_pro,st_pro,cr_pro,cf_pro,cm_pro,pv_pro,balanca_pro,cod_imp,estmin_pro)Values("
    strSQL = strSQL & "'" & sCod & "' ,"
    
    strSQL = strSQL & "'" & txtbxNome_Pro.Text & "',"
    strSQL = strSQL & "'" & txtbxUnidade.Text & "',"
    strSQL = strSQL & "'" & lblCat.caption & "',"
    strSQL = strSQL & "'" & lblMar.caption & "', "
    strSQL = strSQL & "'" & txtbxRef.Text & "',"
    strSQL = strSQL & F_GravaDec(txtbxMargem.Text) & ","
    strSQL = strSQL & "'" & txtbxLocal.Text & "', "
    strSQL = strSQL & F_GravaDec(txtbxEstoque.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxEstoque.Text) & ","
    strSQL = strSQL & "'" & Left(cmbxSt.Text, 1) & "',"
    strSQL = strSQL & F_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxCusto.Text) & ","
    strSQL = strSQL & F_GravaDec(txtbxPreco.Text) & ","
    strSQL = strSQL & "'" & IIf(chkProd.Value = 1, "S", "N") & "',"
    strSQL = strSQL & "'" & Left(cmbxICMS.Text, 3) & "', "
    strSQL = strSQL & F_GravaDec(txtbxEstMin.Text) & ") "

    
    Banco.Execute strSQL
    
    
    
    
    strSQL = "delete from tab_codbarras where cod_pro='" & sCod & "' "
    Banco.Execute strSQL
    With lstCodBarra
        For iCount = 0 To .ListCount - 1
            .ListIndex = iCount
            strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values("
            strSQL = strSQL & "'" & sCod & "',"
            strSQL = strSQL & "'" & .Text & "')"
            Banco.Execute strSQL
            
        Next
    End With
    gravaEmb
    MsgBox "Produto Gravado com Sucesso", vbInformation
    txtbxRef.Text = ""
    txtbxCusto.Text = "0,00"
    txtbxEstoque.Text = "0,00"
    cmbxSt.ListIndex = -1
    cmbxCat.ListIndex = -1
    cmbxMarca.ListIndex = -1
    txtbxUnidade.Text = ""
    txtbxLocal.Text = ""
    txtbxMargem.Text = "0,00"
    txtbxPreco.Text = "0,00"
    txtbxNome_Pro.Text = ""
    txtbxRef.Text = ""
    txtbxCodBarra.Text = ""
    lstCodBarra.Clear
    If IsLoad("frmEdiPro") Then
        RefreshDataPosition frmEdiPro.dtPro
    End If
    cmbxCat.SetFocus
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmnxdbtnBuscaFor_Click()
    frmSelectCli.txtbxNome.tag = "F"
    frmSelectCli.txtbxNome.Text = txtbxFornecedor.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxFornecedor.Text = frmSelectCli.dtBusca.Recordset!nome
    txtbxFornecedor.tag = frmSelectCli.dtBusca.Recordset!cod
    Unload frmSelectCli
End Sub


Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    strSQL = "select num_not,entrada_not,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " tab_fornecedores as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & Me.tag & "' "
    strSQL = strSQL & " order by entrada_not desc limit 1 "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount = 0 Then
            lblUltComp.caption = ""
            lblNomeFor.caption = ""
            lblNumNot.caption = ""
            lblQtd.caption = ""
            lblUnit.caption = ""
            lblEmb.caption = ""
        Else
            lblUltComp.caption = !entrada_not
            lblNomeFor.caption = !nome_for
            lblNumNot.caption = !Num_Not
            lblQtd.caption = !qtd_ite
            lblUnit.caption = FormatVal(!unit_ite)
            lblEmb.caption = !qtd_emb
        End If
        .Close
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnSusp_Click()
    If Me.tag <> "" Then
        
    End If
End Sub



Private Sub Form_Activate()

    MudaAparencia Me
    If bActive Then Exit Sub
    bActive = True
   ' cPreco = txtbxPreco.Text
    buscaCodBar
    If strAltpreco_usu = "S" Then
        lblPrecoCap.Enabled = True
        lblCusCap.Enabled = True
      '  txtbxPreco.Enabled = True
'        txtbxCusto.Enabled = True
        
    Else
        lblPrecoCap.Enabled = False
        lblCusCap.Enabled = False
        'txtbxPreco.Enabled = False
        'txtbxCusto.Enabled = False
        'txtbxCustoFiscal.Enabled = False
    End If
    If Me.tag <> "" Then
        carrega
        buscaPrecos
        buscaSim
        BuscaUltEntrada
        
    End If
    
End Sub

Sub buscaSim()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "select tp.* from tab_produtos as tp,tab_similares as ts "
    strSQL = strSQL & " where tp.cod_pro<>'" & Me.tag & "' "
    strSQL = strSQL & " and tp.cod_pro=ts.cod_pro "
    dt.RecordSource = strSQL
    dt.Refresh
End Sub


Sub carrega()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error Resume Next
    strSQL = Produtos.F_SelectProduto("cod_pro='" & Me.tag & "'")
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
    
        Me.caption = !cod_pro & " - " & !nome_pro
        lblNomePro.caption = !nome_pro
        
        
        txtbxFabricante.tag = !cod_fab
        txtbxFabricante.Text = F_BuscaUmCampo("nome_for", "tab_fornecedores", "cod_for='" & !cod_fab & "'")
        cmbxSecao.Text = !Nome_Sec
        cmbxSubSec.Text = !nome_sub
        If !fracionado_pro = "S" Then
            chkFracionado.Value = 1
        Else
            chkFracionado.Value = 0
        End If
        txtbxPv.Text = FormatVal(!pv_pro)
        txtbxPa.Text = FormatVal(!pa_pro)
        txtbxPa.tag = FormatVal(!pa_pro)
        txtbxPv.tag = FormatVal(!pv_pro)
        txtbxRef.Text = !ref_pro
        txtbxCr.Text = FormatVal(!cr_pro)
        lblEr.caption = !er_pro
        txtbxQueima.Text = FormatVal(!pq_pro)
        txtbxIcms.tag = !cod_imp
        buscaICMS
        If !pq_pro > 0 Then
            chkQueima.Value = 1
            
        Else
        
        End If
        txtbxCB_Pro.Text = FormatVal(!cb_pro)
        txtbxNome_Pro.Text = !nome_pro
        txtbxUnidade.Text = !unidade_pro
        txtbxMargem.Text = FormatVal(!lucro_pro)
        
        txtbxLocal.Text = !endereco_pro
        
        
        P_LocalizaOperacao !st_pro, cmbxSt
        
        If !st_pro = "T" Then
            cmbxICMS.Enabled = True
        End If
        .Close
    End With
End Sub
Sub buscaPrecos()
    Dim strSQL As String
    Dim rPesq As Recordset
    msflxgrdPrecos.rows = 1
    strSQL = "select * from tab_tipoven "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    msflxgrdPrecos.ColAlignment(0) = 3
    msflxgrdPrecos.ColAlignment(1) = 3
    With rPesq
        Do While Not .EOF
        
            strSQL = !descricao_tip & Chr(9) & FormatVal(CCur(txtbxPv.Text) * !fator_tip / !parcelas_tip)
            msflxgrdPrecos.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    ZebraGrid msflxgrdPrecos, 1, &HB4DFE0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        buscaCat
    End If
    If KeyCode = vbKeyF3 Then
       buscaMarca
    End If
    If KeyCode = vbKeyF4 Then
        
    End If
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    End If
    
End Sub
Sub buscaCat()
    frmEdiCat.Show 0
    frmEdiCat.SetFocus
End Sub
Sub buscaMarca()
     frmEdiMar.Show 0
        frmEdiMar.SetFocus
End Sub
Sub RefresheComp()
    
End Sub
Private Sub Form_Load()
'    Me.Height = 5280
'    Me.Width = 4935
    bActive = False
    resizeForm Me
    dt.Connect = Banco.Connect
    With msflxgrdPrecos
        .FixedCols = 0
        .rows = 1
        .Cols = 2
        .Col = 0: .Text = "Plano"
        .Col = 1: .Text = "Preço"
        
        .ColWidth(0) = 1800
    End With
    'P_PreencheMarca cmbxMarca
    'P_PreencheCateg cmbxCat
    
    buscaICMS
    
    secao.P_PreencheSecao cmbxSecao
    cmbxSt.ListIndex = 0
    cmbxSt_LostFocus
    
End Sub
Sub buscaCodBar()
    P_PreencheBarCode Me.tag, lstCodBarra
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
End Sub

Private Sub Form_Resize()
'     P_CentraObejct frControle, Me
End Sub

Private Sub frControle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub lblInfoEnt_Click()
    pcEnt_Click
End Sub

Private Sub lblInfoEnt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfoEnt.FontBold = True
    lblInfoEnt.FontUnderline = True
    pcEnt.BorderStyle = 1
End Sub

Private Sub lblInfoSai_Click()
    pcSai_Click
End Sub

Private Sub lblInfoSai_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfoSai.FontBold = True
    lblInfoSai.FontUnderline = True
    pcSai.BorderStyle = 1
End Sub

Private Sub lstCodBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub lstEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub pcEnt_Click()
    If Me.tag <> "" Then
        P_InfoEntradasPro Me.tag
    Else
        MsgBox "Não há produto gravado!", vbCritical
    End If
End Sub

Private Sub pcEnt_GotFocus()
    lblInfoEnt_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcEnt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcEnt_Click
    End If
End Sub

Private Sub pcEnt_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcInfo_Click()
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub pcInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfoEnt.FontBold = False
    lblInfoEnt.FontUnderline = False
    pcEnt.BorderStyle = 0
    lblInfoSai.FontBold = False
    lblInfoSai.FontUnderline = False
    pcSai.BorderStyle = 0
End Sub

Private Sub pcSai_Click()
    If Me.tag = "" Then Exit Sub
    P_InfoSaidasPro Me.tag
End Sub

Private Sub pcSai_GotFocus()
    lblInfoSai_MouseMove 0, 0, 0, 0
End Sub

Private Sub pcSai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        pcSai_Click
    End If
End Sub

Private Sub pcSai_LostFocus()
    frControle_MouseMove 0, 0, 0, 0
End Sub


Private Sub txtbxPrecoMin_GotFocus()
    
End Sub


Private Sub txtbxCB_Pro_GotFocus()
    P_SelecTexto txtbxCB_Pro
End Sub

Private Sub txtbxCB_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCB_Pro_LostFocus()
    P_AjustaTexto txtbxCB_Pro
End Sub

Private Sub txtbxCr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCr_LostFocus()
    P_ValorLostFocus txtbxCr
End Sub

Private Sub txtbxFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaFab_Click
    End If
End Sub


Private Sub txtbxICMS_Change()
    txtbxIcms.tag = ""
End Sub

Private Sub txtbxIcms_GotFocus()
    P_SelecTexto txtbxIcms
End Sub

Private Sub txtbxIcms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnICMS_Click
    End If
End Sub

Private Sub txtbxPv_GotFocus()
    P_SelecTexto txtbxPv
End Sub

Private Sub txtbxPv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPv_LostFocus()
    P_ValorLostFocus txtbxPv
End Sub


Private Sub txtbxQtdEmb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub


Private Sub txtbxQueima_GotFocus()
    P_SelecTexto txtbxQueima
End Sub

Private Sub txtbxQueima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQueima_LostFocus()
    P_ValorLostFocus txtbxQueima
End Sub

Private Sub txtbxRef_GotFocus()
    If txtbxRef.Text = "" And Me.tag = "" Then
        txtbxRef.Text = MaxKeySQL("tab_produtos", "cod_pro", "", "0000000#")
        lblProvavel.Visible = True
    End If
    P_SelecTexto txtbxRef
End Sub

Private Sub txtbxRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxRef_LostFocus()
    P_AjustaTexto txtbxRef
End Sub

Private Sub txtbxDescricao_GotFocus()
    'P_SelecTexto txtbxDescricao
End Sub

Private Sub txtbxDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDescricao_LostFocus()
    'P_AjustaTexto txtbxDescricao
End Sub







Private Sub txtbxFornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmnxdbtnBuscaFor_Click
    End If
End Sub

Private Sub txtbxLocal_GotFocus()
    P_SelecTexto txtbxLocal
End Sub

Private Sub txtbxLocal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxLocal_LostFocus()
    P_AjustaTexto txtbxLocal
End Sub


Private Sub txtbxMargem_GotFocus()
    P_SelecTexto txtbxMargem
End Sub

Private Sub txtbxMargem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMargem_LostFocus()
    P_ValorLostFocus txtbxMargem
End Sub

Private Sub txtbxNome_Pro_GotFocus()
    P_SelecTexto txtbxNome_Pro
End Sub

Private Sub txtbxNome_Pro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Pro_LostFocus()
    P_AjustaTexto txtbxNome_Pro
End Sub

Private Sub txtbxPei_GotFocus()
    
End Sub



Private Sub txtbxPrecoMin_LostFocus()
    
End Sub


Private Sub txtbxUltComp_Change()

End Sub

Private Sub txtbxUnidade_GotFocus()
    P_SelecTexto txtbxUnidade
End Sub

Private Sub txtbxUnidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnidade_LostFocus()
    P_AjustaTexto txtbxUnidade
End Sub

