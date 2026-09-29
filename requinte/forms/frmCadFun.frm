VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadFun 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Funcionários"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   Icon            =   "frmCadFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11250
   Begin VB.PictureBox Picture10 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   8715
      Left            =   7515
      ScaleHeight     =   8685
      ScaleWidth      =   3705
      TabIndex        =   54
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmndbtnPerm 
         Caption         =   "Permissões"
         Height          =   735
         Left            =   1920
         Picture         =   "frmCadFun.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   6000
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Histórico"
         Height          =   735
         Left            =   120
         Picture         =   "frmCadFun.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   6840
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "ESC - &Fecha&r"
         Height          =   735
         Left            =   1920
         Picture         =   "frmCadFun.frx":0F3E
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   7680
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "F7 - Sal&var"
         Height          =   735
         Left            =   120
         Picture         =   "frmCadFun.frx":10A9
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   7680
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   1920
         Picture         =   "frmCadFun.frx":1205
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   6840
         Width           =   1695
      End
      Begin VB.PictureBox pc3X4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3015
         Index           =   0
         Left            =   360
         ScaleHeight     =   2985
         ScaleWidth      =   3105
         TabIndex        =   59
         Top             =   240
         Width           =   3135
         Begin VB.PictureBox p 
            BorderStyle     =   0  'None
            Height          =   1695
            Index           =   0
            Left            =   480
            ScaleHeight     =   1695
            ScaleWidth      =   1575
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   3105
         TabIndex        =   56
         Tag             =   "!"
         Top             =   0
         Width           =   3135
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            Picture         =   "frmCadFun.frx":1612
            ScaleHeight     =   180
            ScaleWidth      =   210
            TabIndex        =   57
            Top             =   0
            Width           =   210
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fotografia:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   360
            TabIndex        =   58
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.CommandButton cmndbtnFoto 
         Caption         =   "Fotos"
         Height          =   735
         Left            =   120
         Picture         =   "frmCadFun.frx":16C7
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   6000
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   6105
      TabIndex        =   48
      Tag             =   "!"
      Top             =   5760
      Width           =   6135
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         Picture         =   "frmCadFun.frx":177C
         ScaleHeight     =   270
         ScaleWidth      =   285
         TabIndex        =   53
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   360
         TabIndex        =   49
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   6105
      TabIndex        =   44
      Tag             =   "!"
      Top             =   0
      Width           =   6135
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Left            =   0
         Picture         =   "frmCadFun.frx":1B90
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   52
         Top             =   0
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dados do Funcionário"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   360
         TabIndex        =   45
         Top             =   0
         Width           =   1770
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CheckBox chkVale 
         Caption         =   "Vale liberado"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   5160
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ComboBox cmbxCar 
         Height          =   315
         ItemData        =   "frmCadFun.frx":1D0D
         Left            =   1200
         List            =   "frmCadFun.frx":1D1A
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   4800
         Width           =   4815
      End
      Begin VB.TextBox txtbxGuerra 
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
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         Top             =   480
         Width           =   4815
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmCadFun.frx":1D55
         Left            =   1200
         List            =   "frmCadFun.frx":1D5F
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   4440
         Width           =   2775
      End
      Begin VB.TextBox txtbxNatural 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   30
         Top             =   4080
         Width           =   4815
      End
      Begin VB.ComboBox cmbxSexo 
         Height          =   315
         ItemData        =   "frmCadFun.frx":1D7B
         Left            =   4800
         List            =   "frmCadFun.frx":1D85
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtbxConjuge 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   28
         Top             =   3720
         Width           =   4815
      End
      Begin VB.TextBox txtbxEndPais 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   26
         Top             =   3360
         Width           =   4815
      End
      Begin VB.TextBox txtbxMae 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2640
         Width           =   4815
      End
      Begin VB.TextBox txtbxRg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   18
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtbxDataNasc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   20
         Top             =   2280
         Width           =   990
      End
      Begin VB.TextBox txtbxCEP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   14
         TabIndex        =   18
         Text            =   "75.000-000"
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1905
         Width           =   615
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   14
         Text            =   "Anápolis"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtbxBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtbxEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtbxCNPJ 
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
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtbxRazao 
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
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   120
         Width           =   4815
      End
      Begin VB.TextBox txtbxPai 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3000
         Width           =   4815
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Car&go:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Chama-se"
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S&ituação:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Naturalidade:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Sexo:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4080
         TabIndex        =   31
         Top             =   4440
         Width           =   720
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Côn&juge:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "End. Pais:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lblIdade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5400
         TabIndex        =   47
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Idade:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4800
         TabIndex        =   46
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lblMae 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome da &Mãe:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&RG:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Data Nasc:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   19
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CE&P:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&UF:"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4800
         TabIndex        =   15
         Top             =   1905
         Width           =   600
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ci&dade:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Bairro:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Endereço:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblCpf 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&CPF:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Nome:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblPai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome do &Pai:"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1095
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      TabIndex        =   50
      Top             =   6000
      Width           =   6135
      Begin VB.TextBox txtbxTel 
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cmbxTipoTel 
         Height          =   315
         ItemData        =   "frmCadFun.frx":1D9E
         Left            =   1680
         List            =   "frmCadFun.frx":1DB1
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CommandButton cmndbtnAddTel 
         Height          =   315
         Left            =   4560
         Picture         =   "frmCadFun.frx":1DF7
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1560
         Width           =   315
      End
      Begin VB.CommandButton cmndbtnEditTel 
         Height          =   315
         Left            =   4920
         Picture         =   "frmCadFun.frx":2381
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1560
         Width           =   315
      End
      Begin VB.CommandButton cmndbtnExclTel 
         Height          =   315
         Left            =   5280
         Picture         =   "frmCadFun.frx":2D83
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1560
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrdTel 
         Height          =   1335
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   1320
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmCadFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Dim bActive As Boolean


Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim cFile As New Collection
Dim ai As New ADODB.Stream
Dim Indice As Integer
Sub buscaFoto()
    Dim sFile As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    Set cFile = Nothing
    ai.type = adTypeBinary
    strSQL = "select seq_fot,foto_fot from tab_fotofun where cod_fun='" & Me.tag & "' order by seq_fot desc"
    c.Open strNomeBanco
    Set r = c.Execute(strSQL)
    If Dir(sFoto_Temp, vbDirectory) = "" Then
        FileSystem.MkDir sFoto_Temp
    End If
    
    With r
        Do While Not .EOF
            sFile = sFoto_Temp & "\F" & !seq_fot & ".tmp"
            ai.Open
            ai.Write r!foto_fot
            ai.SaveToFile (sFile)
            ai.Close
            cFile.Add sFile
            .MoveNext
        Loop
        
        .Close
    
    End With
    c.Close
    pcDataFoto_click
    Exit Sub
trata_Erro:
    
End Sub

Private Sub pcDataFoto_click()
    Dim sFile As String
    Dim wid As Long
    Dim hei As Long
    Dim iCount As Integer
    
    For iCount = 0 To cFile.Count - 1
    
        'If pc3X4(iCount).Tag = "" Then
        '    If cFile.count > 0 Then
        '        pc3X4(iCount).Tag = 1
        '    Else
        '        Exit Sub
        '    End If
        'End If
            
        If iCount >= pc3X4.Count Then
            Load pc3X4(iCount)
            Load p(iCount)
            pc3X4(iCount).Visible = True
            pc3X4(iCount).top = (pc3X4(iCount - 1).height) + pc3X4(iCount - 1).top + 100
            pc3X4(iCount).Left = (pc3X4(iCount - 1).Left)
        End If
            
        'sFile = cFile(Int(pc3X4(iCount).Tag))
        sFile = cFile(iCount + 1)
        
        'lblFoto.caption = pc3X4(iCount).Tag & " de " & cFile.count
        'lblFoto.Tag = sFile
        On Error GoTo trata_Erro
                pc3X4(iCount).ScaleMode = 1
                p(iCount).ScaleMode = 1
                p(iCount).AutoRedraw = True
                p(iCount).AutoSize = True
                p(iCount).Picture = LoadPicture(sFile)
                wid = p(iCount).image.width
                hei = p(iCount).image.height
                w = pc3X4(iCount).width
                h = pc3X4(iCount).height
                Do Until hei < h And wid < w
                    hei = hei * 0.995
                    wid = wid * 0.995
                Loop
                pc3X4(iCount).Picture = Nothing
                pc3X4(iCount).PaintPicture LoadPicture(sFile), 0, 0, ScaleX(wid, pc3X4(iCount).ScaleMode, vbTwips), ScaleY(hei, pc3X4(iCount).ScaleMode, vbTwips)
                pc3X4(iCount).AutoSize = True
                pc3X4(iCount).Refresh
                'lblFoto.caption = .AbsolutePosition & " de " & .RecordCount
               ' Kill sFile
        Next
        
        
        
        
        
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub


Sub P_Salva()
    P_GravaTelefones
End Sub

Sub P_GravaTelefones()
    Dim strSQL As String
    Dim iCount As Integer
    Dim sTipo As String
    Dim sFone As String
    strSQL = "delete from tab_telefones where cod_cli='" & Me.tag & "' "
    strSQL = strSQL & " and tipo_ter='U' "
    Conn.Execute strSQL
    With msflxgrdTel
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0: sFone = .text
            .col = 1: sTipo = Left(.text, 1)
            If F_GravaTelefone(Me.tag, sFone, sTipo, "U") Then
            End If
            
        Next
    End With
End Sub



Private Sub cmbxEstadoCiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub cmbxCar_Click()
    cmbxCar.tag = F_BuscaUmCampo("cod_Car", "tab_cargos", "nome_car='" & cmbxCar.text & "'")
    
    
End Sub

Private Sub cmbxCar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxStatus_click()
    If cmbxStatus.ListIndex <> 0 Then
        
    Else
        
    End If
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub










Private Sub cmndbtnHist_Click()
    If Me.tag = "" Then
    Else
        P_histCli Me.tag
    End If
End Sub

Private Sub cmndbtnFoto_click()
    frmGravaFoto.sTab = "tab_fotofun"
    frmGravaFoto.sCampo = "cod_fun"
    frmGravaFoto.sTipo = "F"
    
    frmGravaFoto.caption = Me.tag & "-" & Me.caption
    frmGravaFoto.tag = Me.tag
    frmGravaFoto.Show 1
    deletaArquivos
    
    buscaFoto
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    Dim iCount As Integer
    If strCadFun_Usu <> "S" Then
        MsgBox "Você não tem permissão para cadastrar funcionários", vbExclamation
        Exit Sub
    End If
    
    If txtbxRazao.text = "" Then
        MsgBox "Informe um nome para o funcionario!", vbExclamation
        Exit Sub
        
    ElseIf F_Vazio(txtbxCNPJ.text) And F_Vazio(txtbxRg.text) Then
        
            MsgBox "Informe o CPF  ou CI do funcionario !", vbExclamation
            txtbxCNPJ.SetFocus
            Exit Sub
    ElseIf cmbxCar.tag = "" Then
        MsgBox "Informe o cargo do Funcionário", vbExclamation
        cmbxCar.SetFocus
        Exit Sub
    ElseIf txtbxEndereco.text = "" Then
        MsgBox "Informe o endereço do funcionario!", vbExclamation
        Exit Sub
    ElseIf txtbxBairro.text = "" Then
        MsgBox "informe o Bairro do funcionario!,vbExclamation  "
        Exit Sub
    ElseIf txtbxCidade.text = "" Then
        MsgBox "Informe a Cidade do funcionario!", vbExclamation
        Exit Sub
    ElseIf cmbxUf.text = "" Then
        MsgBox "Informe o Estado de Origem do funcionario!", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("cod_fun", "tab_funcionarios", "cpf_fun='" & txtbxCNPJ.text & "'") <> "" Then
            MsgBox "Já existe um funcionario com este CPF/CNPJ!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_fun", "tab_funcionarios", "nome_fun='" & txtbxRazao.text & "'") <> "" Then
            MsgBox "Já existe um funcionario com este Nome!", vbCritical
            Exit Sub
        
        End If
    
        strSQL = funcionario.F_GravaFun(txtbxRazao.text, txtbxGuerra, txtbxEndereco, _
            txtbxBairro.text, txtbxCidade.text, cmbxUf.text, txtbxCEP.text, txtbxDataNasc.text, txtbxNatural.text, _
            txtbxRg.text, txtbxCNPJ.text, Left(cmbxSexo.text, 1), txtbxConjuge.text, txtbxPai.text, _
            txtbxMae.text, txtbxEndPais.text, "", cmbxCar.tag)
        
        
        P_GravaTelefones
        
        MsgBox "funcionario Cadastrado com sucesso!", vbInformation
        P_Limpa
        
        txtbxRazao.text = ""
        Me.tag = ""
    Else
        If MsgBox("Confirma alteração do funcionario " & Me.tag & " ?", vbYesNo + vbQuestion) = vbYes Then
            If funcionario.F_AlteraFun(Me.tag, txtbxRazao.text, txtbxGuerra.text, txtbxEndereco.text, txtbxBairro.text, _
                 txtbxCidade.text, cmbxUf.text, txtbxCEP.text, txtbxDataNasc.text, txtbxNatural.text, _
                 txtbxRg.text, txtbxCNPJ.text, Left(cmbxSexo.text, 1), txtbxConjuge.text, _
                    txtbxPai.text, txtbxMae.text, txtbxEndPais.text, "", Left(cmbxStatus.text, 1), cmbxCar.tag, IIf(chkVale.Value = 1, "S", "N")) Then
                MsgBox "funcionario alterado com sucesso!", vbInformation
            End If
                
            Unload Me
        End If
    End If
    
End Sub
Sub P_Limpa()
    Dim i As Integer
    For i = 0 To Me.Controls.Count - 1
        If LCase(Left(Me.Controls(i).Name, 3)) = "txt" Then
            Me.Controls(i).text = ""
        End If
    Next
   
    cmbxUf.ListIndex = -1
End Sub







Private Sub Form_Activate()
    
    If bActive Then Exit Sub
    
    
   ' MudaAparencia Me
    
    
    
    
    If Me.tag <> "" Then
        carrega
        buscaFoto
    End If
    
    
    
    
    
bActive = True
    
End Sub
Sub carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select * from tab_funcionarios where cod_fun = '" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do Until .EOF
            If TiraNull(!vale_fun) <> "S" Then
                chkVale.Value = 0
            Else
                chkVale.Value = 1
            End If
            txtbxBairro.text = !bairro_fun
            txtbxCEP.text = !cep_fun
            txtbxCidade.text = !cidade_fun
            txtbxCNPJ.text = !cpf_fun
            txtbxRg.text = !rg_fun
            txtbxEndereco.text = !endereco_fun
            txtbxMae.text = Format(!mae_fun)
            txtbxPai.text = Format(!pai_fun)
            txtbxRazao.text = !nome_fun
            txtbxDataNasc.text = Format(!nasc_fun)
            txtbxEndPais.text = !endpais_fun
            txtbxNatural.text = !natural_fun
            P_LocalizaOperacao !sexo_fun, cmbxSexo
            txtbxConjuge.text = !conjuge_fun
            calcIdade
            
            txtbxRg.text = !rg_fun
            
            txtbxGuerra.text = !guerra_fun
            cmbxCar.text = F_BuscaUmCampo("nome_car", "tab_cargos", "cod_car='" & !cod_Car & "'")
            If !uf_fun <> "" Then cmbxUf.text = !uf_fun
            
            
            
            
            
            


            
            
            P_LocalizaOperacao !status_fun, cmbxStatus
            P_CarregaTelefone
            
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    bActive = False
    'Me.Height = 6540
    'Me.Width = 6255
   ' P_CentraForm Me
    
    resizeForm Me
    P_PreencheCargos cmbxCar
    
    If Me.tag = "" Then
        bCNPJ = True
    End If
    'stb.Tab = 0
    
    
    
    
    If strContrMens_usu <> "S" Then
'        chkMens.Enabled = False
'        txtbxLimite.Enabled = False
    Else
'        chkMens.Enabled = True
'        txtbxLimite.Enabled = True
    End If
    With msflxgrdTel
        .FixedCols = 0
        .rows = 1
        .Cols = 2
        .col = 0: .text = "Fone:"
        .col = 1: .text = "Tipo:"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1200
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
    On Error Resume Next
    P_PreencheUF cmbxUf
    cmbxUf.text = "GO"
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsLoad(frmEdiFun.Name) Then
        'RefreshDataPosition frmEdiFun.adc
    End If
End Sub







Private Sub txtbxBairro_GotFocus()
    P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    P_AjustaTexto txtbxBairro
End Sub



Private Sub txtbxCEP_GotFocus()
    P_SelecTexto txtbxCEP
End Sub

Private Sub txtbxCEP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCEP_LostFocus()
    P_CEPLostFocus txtbxCEP
End Sub


Private Sub txtbxCidade_GotFocus()
    P_SelecTexto txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    P_AjustaTexto txtbxCidade
End Sub

Private Sub txtbxCNPJ_GotFocus()
    P_SelecTexto txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    
    
    
        If Not F_CPF(txtbxCNPJ.text) Then
            MsgBox "CPF Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCPF txtbxCNPJ
    
End Sub



Private Sub txtbxConjuge_GotFocus()
    P_SelecTexto txtbxConjuge
End Sub

Private Sub txtbxConjuge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxConjuge_LostFocus()
    P_AjustaTexto txtbxConjuge
End Sub

Private Sub txtbxDataNasc_GotFocus()
    P_SelecTexto txtbxDataNasc
End Sub

Private Sub txtbxDataNasc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataNasc_LostFocus()
    P_AjustaData txtbxDataNasc
    calcIdade
End Sub
Sub calcIdade()
    On Error Resume Next
    If txtbxDataNasc.text = "" Then
        Exit Sub
    End If
    lblIdade.caption = DateDiff("YYYY", txtbxDataNasc.text, Date)
End Sub
Private Sub txtbxEndereco_GotFocus()
    P_SelecTexto txtbxEndereco
End Sub

Private Sub txtbxEndereco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEndereco_LostFocus()
    P_AjustaTexto txtbxEndereco
End Sub


Private Sub txtbxEndPais_GotFocus()
    P_SelecTexto txtbxEndPais
End Sub

Private Sub txtbxEndPais_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEndPais_LostFocus()
    P_AjustaTexto txtbxEndPais
End Sub

Private Sub txtbxGuerra_GotFocus()
    P_SelecTexto txtbxGuerra
End Sub

Private Sub txtbxGuerra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxGuerra_LostFocus()
    P_AjustaTexto txtbxGuerra
End Sub

Private Sub txtbxMae_GotFocus()
    P_SelecTexto txtbxMae
End Sub

Private Sub txtbxMae_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMae_LostFocus()
    P_AjustaTexto txtbxMae
End Sub

Private Sub txtbxNatural_GotFocus()
    P_SelecTexto txtbxNatural
End Sub

Private Sub txtbxNatural_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNatural_LostFocus()
    P_AjustaTexto txtbxNatural
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSalvar.SetFocus
    End If
End Sub

Private Sub txtbxPai_GotFocus()
    P_SelecTexto txtbxPai
End Sub

Private Sub txtbxPai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxPai_LostFocus()
    P_AjustaTexto txtbxPai
End Sub

Private Sub txtbxRazao_GotFocus()
    P_SelecTexto txtbxRazao
End Sub

Private Sub txtbxRazao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxRazao_LostFocus()
    P_AjustaTexto txtbxRazao
End Sub


Private Sub txtbxRg_GotFocus()
    P_SelecTexto txtbxRg
End Sub

Private Sub txtbxRg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxRg_LostFocus()
    P_AjustaTexto txtbxRg
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.text = F_AjustaTelefone(txtbxTel.text)
End Sub

Sub P_CarregaTelefone()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipo As String
    msflxgrdTel.rows = 1
    strSQL = "select * from tab_telefones where cod_CLI='" & Me.tag & "' and tipo_ter='U' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If !tipo_tel = "1" Then
                sTipo = "1 - Residencial"
            ElseIf !tipo_tel = "2" Then
                sTipo = "2 - Comercial"
            ElseIf !tipo_tel = "3" Then
                sTipo = "3 - Celular"
            ElseIf !tipo_tel = "4" Then
                sTipo = "4 - Recado"
            ElseIf !tipo_tel = "5" Then
                sTipo = "5 - Fax"
            End If
            strSQL = !num_tel & Chr(9) & sTipo
            msflxgrdTel.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub
Private Sub txtbxTel_GotFocus()
    P_SelecTexto txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub


Private Sub cmbxTipoTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAddTel_Click()
    If F_Vazio(txtbxTel.text) Then
        MsgBox "Infome o nº do Telefone!", vbExclamation
        txtbxTel.SetFocus
        Exit Sub
    End If
    If cmbxTipoTel.text = "" Then
        MsgBox "Informe o tipo do Telefone!", vbExclamation
        cmbxTipoTel.SetFocus
        Exit Sub
    End If
    If F_InFlxgrd(msflxgrdTel, txtbxTel.text, 0) > -1 Then
        MsgBox "Este telefone já se encontra Adicionado!", vbExclamation
        txtbxTel.SetFocus
        Exit Sub
    End If
    If cmndbtnAddTel.tag = "" Then
        msflxgrdTel.AddItem txtbxTel.text & Chr(9) & cmbxTipoTel.text
    Else
        msflxgrdTel.col = 0: msflxgrdTel.text = txtbxTel.text
        msflxgrdTel.col = 1: msflxgrdTel.text = cmbxTipoTel.text
    End If
    If Me.tag <> "" Then
        If MsgBox("Deseja salvar alterações?", vbYesNo + vbQuestion) = vbYes Then
            P_GravaTelefones
        End If
    End If
    cmndbtnAddTel.tag = ""
    txtbxTel.SetFocus
    
End Sub

Private Sub cmndbtnEditTel_Click()
    With msflxgrdTel
        If .row <= 0 Then
            Exit Sub
        End If
        .col = 0
        txtbxTel.text = .text
        .col = 1
        P_LocalizaOperacao .text, cmbxTipoTel
        cmndbtnAddTel.tag = .row
    End With
End Sub

Private Sub cmndbtnExclTel_Click()
    With msflxgrdTel
        If .row <= 0 Then
            Exit Sub
        End If
        If MsgBox("Deseja excluir o telefone?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        If .rows = 2 Then
            .rows = 1
        Else
            .RemoveItem .row
        End If
    End With
End Sub

Sub deletaArquivos()
    On Error Resume Next
    Dim iCount As Integer
    For iCount = 1 To cFile.Count
        Kill cFile(iCount)
    Next
    Set cFile = Nothing
End Sub

