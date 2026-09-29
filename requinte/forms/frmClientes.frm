VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIncClientes 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13155
   Icon            =   "frmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   13155
   Begin VB.PictureBox Picture9 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   6315
      TabIndex        =   80
      Top             =   4800
      Width           =   6375
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   10920
      Picture         =   "frmClientes.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   255
      TabIndex        =   79
      Top             =   160
      Width           =   255
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Left            =   6720
      Picture         =   "frmClientes.frx":098F
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   78
      Top             =   145
      Width           =   285
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   210
      Left            =   8790
      Picture         =   "frmClientes.frx":0DA3
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   77
      Top             =   145
      Width           =   210
   End
   Begin TabDlg.SSTab sstb 
      Height          =   4695
      Left            =   6600
      TabIndex        =   25
      Top             =   120
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Telefones"
      TabPicture(0)   =   "frmClientes.frx":0F20
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Autorizados"
      TabPicture(1)   =   "frmClientes.frx":0F3C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Referências"
      TabPicture(2)   =   "frmClientes.frx":0F58
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00B36A09&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         ScaleHeight     =   225
         ScaleWidth      =   6105
         TabIndex        =   45
         Tag             =   "!"
         Top             =   480
         Width           =   6135
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência"
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
            Left            =   120
            TabIndex        =   76
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "&Referências:"
         Height          =   3855
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   6135
         Begin VB.TextBox txtbxTelRef 
            Height          =   285
            Left            =   4320
            TabIndex        =   53
            Top             =   2280
            Width           =   1575
         End
         Begin VB.ComboBox cmbxParentRef 
            Height          =   315
            ItemData        =   "frmClientes.frx":0F74
            Left            =   3000
            List            =   "frmClientes.frx":0F8A
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmndbtnModifRef 
            Caption         =   "Modificar"
            Height          =   375
            Left            =   1560
            TabIndex        =   55
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CommandButton cmndbtnExclRef 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   3000
            TabIndex        =   56
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox txtbxNomeRef 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   49
            Top             =   2280
            Width           =   2775
         End
         Begin VB.CommandButton cmndbtnAddRef 
            Caption         =   "&Adicionar"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   2760
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdRef 
            Height          =   1695
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2990
            _Version        =   393216
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Nº:"
            Height          =   195
            Left            =   4320
            TabIndex        =   52
            Top             =   2040
            Width           =   225
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco:"
            Height          =   195
            Left            =   3000
            TabIndex        =   50
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N&ome:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   2040
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "&Pessoas Autorizadas:"
         Height          =   3615
         Left            =   -74880
         TabIndex        =   36
         Top             =   840
         Width           =   6135
         Begin VB.CommandButton cmndbtnAddAut 
            Caption         =   "&Adicionar"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtbxAutorizado 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   39
            Top             =   1920
            Width           =   3975
         End
         Begin VB.CommandButton cmndbrnExclAut 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   3000
            TabIndex        =   44
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmndbtnEditAut 
            Caption         =   "Modificar"
            Height          =   375
            Left            =   1560
            TabIndex        =   43
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox cmbxParentesco 
            Height          =   315
            ItemData        =   "frmClientes.frx":0FD8
            Left            =   4200
            List            =   "frmClientes.frx":0FEE
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1920
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdAut 
            Height          =   1335
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2355
            _Version        =   393216
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N&ome:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   465
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco:"
            Height          =   195
            Left            =   4200
            TabIndex        =   40
            Top             =   1680
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00B36A09&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         ScaleHeight     =   225
         ScaleWidth      =   6105
         TabIndex        =   35
         Tag             =   "!"
         Top             =   480
         Width           =   6135
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pessoas Autorizadas"
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
            Left            =   120
            TabIndex        =   75
            Top             =   0
            Width           =   1755
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3735
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   6135
         Begin VB.TextBox txtbxTel 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   1455
         End
         Begin VB.ComboBox cmbxTipoTel 
            Height          =   315
            ItemData        =   "frmClientes.frx":103C
            Left            =   1680
            List            =   "frmClientes.frx":104F
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2160
            Width           =   2775
         End
         Begin VB.CommandButton cmndbtnAddTel 
            Caption         =   "Adicionar"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmndbtnEditTel 
            Caption         =   "Modificar"
            Height          =   375
            Left            =   1680
            TabIndex        =   33
            Top             =   2640
            Width           =   1335
         End
         Begin VB.CommandButton cmndbtnExclTel 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   3120
            TabIndex        =   34
            Top             =   2640
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdTel 
            Height          =   1575
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2778
            _Version        =   393216
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Nº:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1920
            Width           =   225
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "&Tipo:"
            Height          =   255
            Left            =   1680
            TabIndex        =   30
            Top             =   1920
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00B36A09&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   6105
         TabIndex        =   73
         Tag             =   "!"
         Top             =   480
         Width           =   6135
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
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   74
            Top             =   0
            Width           =   840
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00B36A09&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      ScaleHeight     =   225
      ScaleWidth      =   6345
      TabIndex        =   69
      Tag             =   "!"
      Top             =   4920
      Width           =   6375
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stautus do Cliente"
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
         Left            =   120
         TabIndex        =   70
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B36A09&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   6345
      TabIndex        =   67
      Tag             =   "!"
      Top             =   120
      Width           =   6375
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dados do Cliente"
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
         Left            =   120
         TabIndex        =   68
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   6600
      TabIndex        =   57
      Top             =   5160
      Width           =   6375
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmClientes.frx":1095
         Left            =   4080
         List            =   "frmClientes.frx":109F
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkMens 
         Caption         =   "&Cliente Mensalista"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtbxLimite 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   60
         Text            =   "0,00"
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação do Cliente:"
         Height          =   195
         Left            =   4080
         TabIndex        =   61
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Limite de Débito:"
         Height          =   195
         Left            =   2160
         TabIndex        =   59
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   6600
      TabIndex        =   63
      Top             =   6720
      Width           =   6375
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   735
         Left            =   120
         Picture         =   "frmClientes.frx":10BF
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fecha&r"
         Height          =   735
         Left            =   3840
         Picture         =   "frmClientes.frx":1649
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "Sal&var"
         Height          =   735
         Left            =   1920
         Picture         =   "frmClientes.frx":1793
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Anotações:"
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   6375
      Begin VB.TextBox txtbxObs 
         Height          =   495
         Left            =   240
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin VB.TextBox txtbxMae 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtbxRg 
         Height          =   285
         Left            =   4320
         MaxLength       =   18
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtbxDataNasc 
         Height          =   285
         Left            =   3840
         MaxLength       =   40
         TabIndex        =   22
         Top             =   2760
         Width           =   990
      End
      Begin VB.TextBox txtbxCEP 
         Height          =   285
         Left            =   1200
         MaxLength       =   14
         TabIndex        =   20
         Text            =   "75.000-000"
         Top             =   2760
         Width           =   975
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtbxCidade 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   16
         Text            =   "Anápolis"
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtbxBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox txtbxEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   12
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox txtbxCNPJ 
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
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtbxRazao 
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
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtbxPai 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label lblIdade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5520
         TabIndex        =   72
         Top             =   2835
         Width           =   135
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Idade:"
         Height          =   195
         Left            =   4965
         TabIndex        =   71
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da &Mãe:"
         Height          =   195
         Left            =   30
         TabIndex        =   3
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&RG:"
         Height          =   195
         Left            =   3720
         TabIndex        =   9
         Top             =   1320
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Nasc:"
         Height          =   195
         Left            =   2760
         TabIndex        =   21
         Top             =   2805
         Width           =   810
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CE&P:"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   2805
         Width           =   360
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&UF:"
         Height          =   195
         Left            =   5160
         TabIndex        =   17
         Top             =   2445
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   2445
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bairro:"
         Height          =   195
         Left            =   630
         TabIndex        =   13
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         Height          =   255
         Left            =   -15
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblCpf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   660
         TabIndex        =   7
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   525
         TabIndex        =   1
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do &Pai:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmIncClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Dim bActive As Boolean
Sub P_Salva()
    Dim strSQL As String
    Dim sCod As String
    sCod = MaxKeySQL("tab_clientes", "cod_cli", "", "0000#")
    
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,datanasc_cli,cpf_cli,rg_cli,"
    strSQL = strSQL & "endereco_Cli,bairro_cli,cidade_Cli,"
    strSQL = strSQL & "uf_cli,cep_cli,"
    strSQL = strSQL & "data_cli,obs_cli,mae_cli,pai_cli,mens_cli,limite_cli,status_cli) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & txtbxRazao.Text & "',"
    strSQL = strSQL & F_GravaData(txtbxDataNasc.Text) & ","
    strSQL = strSQL & "'" & txtbxCnpj.Text & "',"
    strSQL = strSQL & "'" & txtbxRg.Text & "',"
    strSQL = strSQL & "'" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "'" & txtbxBairro.Text & "',"
    strSQL = strSQL & "'" & txtbxCidade.Text & "',"
    strSQL = strSQL & "'" & cmbxUF.Text & "',"
    strSQL = strSQL & "'" & txtbxCep.Text & "',"
    
    
    
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & txtbxObs.Text & "',"
    strSQL = strSQL & "'" & txtbxMae.Text & "',"
    strSQL = strSQL & "'" & txtbxPai.Text & "',"
    If chkMens.Value = 1 Then
        strSQL = strSQL & "'S',"
    Else
        strSQL = strSQL & "'N',"
    End If
    strSQL = strSQL & F_GravaDec(txtbxLimite.Text) & ","
    strSQL = strSQL & "'" & Left(cmbxStatus.Text, 1) & "')"
    Banco.Execute strSQL
    Me.Tag = sCod
    P_GravaAutorizados
    P_GravaReferencia
End Sub

Sub P_GravaAutorizados()
    Dim strSQL As String
    Dim iCount As Integer
    strSQL = "delete from tab_autorizados where cod_cli='" & Me.Tag & "' "
    Banco.Execute strSQL
    With msflxgrdAut
        For iCount = 1 To .rows - 1
            
            .Row = iCount
            strSQL = "insert into tab_autorizados(cod_cli,seq_aut,nome_aut,status_aut,cod_tip)values("
            strSQL = strSQL & "'" & Me.Tag & "',"
            strSQL = strSQL & iCount & ","
            .Col = 0
            strSQL = strSQL & "'" & .Text & "',"
            strSQL = strSQL & "'1', "
            .Col = 2
            strSQL = strSQL & "'" & .Text & "')"
            Banco.Execute strSQL
        Next
    End With
End Sub
Sub P_GravaReferencia()
    Dim strSQL As String
    Dim iCount As Integer
    strSQL = "delete from tab_referencias where cod_cli='" & Me.Tag & "' "
    Banco.Execute strSQL
    With msflxgrdRef
        For iCount = 1 To .rows - 1
            
            .Row = iCount
            strSQL = "insert into tab_referencias(cod_cli,seq_ref,nome_ref,cod_tip,tel_ref)values("
            strSQL = strSQL & "'" & Me.Tag & "',"
            strSQL = strSQL & iCount & ","
            .Col = 0
            strSQL = strSQL & "'" & .Text & "',"
            .Col = 2
            strSQL = strSQL & "'" & Left(.Text, 1) & "',"
            .Col = 3
            strSQL = strSQL & "'" & .Text & "')"
            Banco.Execute strSQL
        Next
    End With
End Sub

Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_clientes set "
    strSQL = strSQL & "nome_cli='" & txtbxRazao.Text & "',"
    strSQL = strSQL & "cpf_Cli='" & txtbxCnpj.Text & "',"
    strSQL = strSQL & "rg_cli='" & txtbxRg.Text & "',"

    strSQL = strSQL & "endereco_Cli='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "bairro_Cli='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "cidade_Cli='" & txtbxCidade.Text & "',"
    strSQL = strSQL & "uf_Cli='" & cmbxUF.Text & "',"
    strSQL = strSQL & "cep_Cli='" & txtbxCep.Text & "',"
    strSQL = strSQL & "fone_Cli='" & txtbxTel.Text & "',  "
    'strSQL = strSQL & "fonecom_Cli='" & txtbxTelCom.Text & "',  "
    'strSQL = strSQL & "celular_Cli='" & txtbxCelular.Text & "',  "
    strSQL = strSQL & "dataNASC_CLI=" & F_GravaData(txtbxDataNasc.Text) & ","
    strSQL = strSQL & "obs_cli='" & txtbxObs.Text & "', "
    strSQL = strSQL & "mae_cli='" & txtbxMae.Text & "',"
    strSQL = strSQL & "pai_cli='" & txtbxPai.Text & "',"
    strSQL = strSQL & "limite_cli=" & F_GravaDec(txtbxLimite.Text) & ","
    strSQL = strSQL & "mens_cli='" & IIf(chkMens.Value = 1, "S", "N") & "', "
    strSQL = strSQL & "status_cli='" & Left(cmbxStatus.Text, 1) & "' "
    
         
         
    strSQL = strSQL & " where cod_Cli = '" & Me.Tag & "'"
    P_GravaAutorizados
    Banco.Execute strSQL
End Sub

Private Sub cmbxEstadoCiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub chkMens_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxParentesco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxParentRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxUF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbrnExclAut_Click()
    If msflxgrdAut.rows > 1 Then
    
    
        If MsgBox("Confirma Exclusão?", vbYesNo + vbCritical) = vbYes Then
            If msflxgrdAut.rows = 2 Then
                msflxgrdAut.rows = 1
            Else
                msflxgrdAut.RemoveItem msflxgrdAut.Row
            End If
        End If
    Else
        MsgBox "Não há o que excluir!", vbCritical
    End If
End Sub

Private Sub cmndbtnAddAut_Click()
    If cmbxParentesco.Text = "" Then
        MsgBox "Informe o grau de parentesco!", vbExclamation
        cmbxParentesco.SetFocus
        Exit Sub
    End If
    If F_Vazio(txtbxAutorizado.Text) Then
        MsgBox "Informe um nome!", vbExclamation
        txtbxAutorizado.SetFocus
        Exit Sub
    End If
    If cmndbtnAddAut.Tag = "" Then
        If F_InFlxgrd(msflxgrdAut, txtbxAutorizado.Text, 0) > -1 Then
            MsgBox "O Nome já se encontra incluído!", vbCritical
        
            txtbxAutorizado.SetFocus
            Exit Sub
        End If
        msflxgrdAut.AddItem txtbxAutorizado.Text & Chr(9) & Mid(cmbxParentesco.Text, 6, Len(cmbxParentesco.Text)) & Chr(9) & Left(cmbxParentesco.Text, 2)
        cmndbtnAddAut.Tag = ""
        txtbxAutorizado.SetFocus
    Else
        msflxgrdAut.Row = msflxgrdAut.Tag
        msflxgrdAut.Col = 0: msflxgrdAut.Text = txtbxAutorizado.Text
        msflxgrdAut.Col = 1: msflxgrdAut.Text = Mid(cmbxParentesco.Text, 6, Len(cmbxParentesco.Text))
        msflxgrdAut.Col = 2: msflxgrdAut.Text = Left(cmbxParentesco.Text, 2)
    End If
End Sub

Private Sub cmndbtnAddRef_Click()
    If cmbxParentRef.Text = "" Then
        MsgBox "Informe o grau de parentesco!", vbExclamation
        cmbxParentesco.SetFocus
        Exit Sub
    End If
    If F_Vazio(txtbxTelRef.Text) Then
        MsgBox "Informe um nome!", vbExclamation
        txtbxAutorizado.SetFocus
        Exit Sub
    End If
    If cmndbtnAddRef.Tag = "" Then
        If F_InFlxgrd(msflxgrdRef, txtbxNomeRef.Text, 0) > -1 Then
            MsgBox "O Nome já se encontra incluído!", vbCritical
        
            txtbxNomeRef.SetFocus
            Exit Sub
        End If
        msflxgrdRef.AddItem txtbxNomeRef.Text & Chr(9) & Mid(cmbxParentRef.Text, 6, Len(cmbxParentRef.Text)) & Chr(9) & Left(cmbxParentRef.Text, 2) & Chr(9) & txtbxTelRef.Text
        cmndbtnAddRef.Tag = ""
        txtbxNomeRef.SetFocus
    Else
        msflxgrdRef.Row = msflxgrdRef.Tag
        msflxgrdRef.Col = 0: msflxgrdRef.Text = txtbxNomeRef.Text
        msflxgrdRef.Col = 1: msflxgrdRef.Text = Mid(cmbxParentRef.Text, 6, Len(cmbxParentRef.Text))
        msflxgrdRef.Col = 2: msflxgrdRef.Text = Left(cmbxParentRef.Text, 2)
        msflxgrdRef.Col = 3: msflxgrdRef.Text = txtbxTelRef.Text
    End If

End Sub

Private Sub cmndbtnEditAut_Click()
    With msflxgrdAut
        If .Row <= 0 Then
            Exit Sub
        End If
        .Col = 0
        txtbxAutorizado.Text = .Text
        .Col = 2
        P_LocalizaOperacao .Text, cmbxParentesco
        cmndbtnAddAut.Tag = .Row
        msflxgrdAut.Tag = .Row
    End With
End Sub

Private Sub cmndbtnExclRef_Click()
        If msflxgrdRef.rows > 1 Then
    
    
        If MsgBox("Confirma Exclusão?", vbYesNo + vbCritical) = vbYes Then
            If msflxgrdRef.rows = 2 Then
                msflxgrdRef.rows = 1
            Else
                msflxgrdRef.RemoveItem msflxgrdRef.Row
            End If
        End If
    Else
        MsgBox "Não há o que excluir!", vbCritical
    End If

End Sub

Private Sub cmndbtnModifRef_Click()
        With msflxgrdRef
        If .Row <= 0 Then
            Exit Sub
        End If
        .Col = 0
        txtbxNomeRef.Text = .Text
        .Col = 2
        P_LocalizaOperacao .Text, cmbxParentRef
        .Col = 3
        txtbxTelRef.Text = .Text
        cmndbtnAddRef.Tag = .Row
        msflxgrdRef.Tag = .Row
    End With

End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()

    If txtbxRazao.Text = "" Then
        MsgBox "Informe um nome para o Cliente!", vbExclamation
        Exit Sub
        
    ElseIf F_Vazio(txtbxCnpj.Text) And F_Vazio(txtbxRg.Text) Then
        
            MsgBox "Informe o CPF  ou CI do Cliente!", vbExclamation
            txtbxCnpj.SetFocus
            Exit Sub
        
    
    
    ElseIf txtbxEndereco.Text = "" Then
        MsgBox "Informe o endereço do Cliente!", vbExclamation
        Exit Sub
    ElseIf txtbxBairro.Text = "" Then
        MsgBox "informe o Bairro do Cliente!,vbExclamation  "
        Exit Sub
    ElseIf txtbxCidade.Text = "" Then
        MsgBox "Informe a Cidade do Cliente!", vbExclamation
        Exit Sub
    ElseIf cmbxUF.Text = "" Then
        MsgBox "Informe o Estado de Origem do Cliente!", vbExclamation
        Exit Sub
    End If
    
    If Me.Tag = "" Then
        If F_BuscaUmCampo("cod_cli", "tab_clientes", "cpf_cli='" & txtbxCnpj.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este CPF/CNPJ!", vbCritical
            Exit Sub
        ElseIf F_BuscaUmCampo("cod_cli", "tab_clientes", "nome_cli='" & txtbxRazao.Text & "'") <> "" Then
            MsgBox "Já existe um cliente com este Nome/Razão Social!", vbCritical
            Exit Sub
        
        End If
    
        P_Salva
        MsgBox "Cliente Cadastrado com sucesso!", vbInformation
        P_Limpa
        
        txtbxRazao.Text = ""
        Me.Tag = ""
    Else
        If MsgBox("Confirma alteração do cliente " & Me.Tag & " ?", vbYesNo + vbQuestion) = vbYes Then
            P_Altera
            MsgBox "Cliente Alterado com sucesso!", vbInformation
            Unload Me
        End If
    End If
    
End Sub
Sub P_Limpa()
    Dim i As Integer
    For i = 0 To Me.Controls.count - 1
        If LCase(Left(Me.Controls(i).Name, 3)) = "txt" Then
            Me.Controls(i).Text = ""
        End If
    Next
   
    cmbxUF.ListIndex = -1
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub




Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    Dim rPesq As Recordset
    Dim strSQL As String
    If bActive Then Exit Sub
    
    
    MudaAparencia Me
    
    
    
    If Me.Tag <> "" Then
        strSQL = "select * from tab_clientes where cod_cli = '" & Me.Tag & "' "
        Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
        With rPesq
            Do Until .EOF
                txtbxBairro.Text = !bairro_cli
                txtbxCep.Text = !cep_Cli
                txtbxCidade.Text = !cidade_Cli
                txtbxCnpj.Text = !cpf_cli
                txtbxRg.Text = !rg_cli
                txtbxEndereco.Text = !endereco_cli
                txtbxMae.Text = Format(!mae_cli)
                txtbxPai.Text = Format(!pai_cli)
                txtbxRazao.Text = !nome_cli
                txtbxDataNasc.Text = Format(!datanasc_cli)
'
                calcIdade
                txtbxObs.Text = Format(!obs_cli)
                txtbxRg.Text = !rg_cli
                If !mens_cli = "S" Then
                    chkMens.Value = 1
                Else
                    chkMens.Value = 0
                End If
                txtbxLimite.Text = FormatVal(!limite_cli)
                If !UF_Cli <> "" Then cmbxUF.Text = !UF_Cli
                
                
                
                
                
                


                
                
                P_LocalizaOperacao !status_cli, cmbxStatus
                P_CarregaTelefone
                P_CarregaRef
                
                .MoveNext
            Loop
        End With
    End If
    strSQL = "select ta.*,tt.nome_tip from tab_autorizados as ta, tab_tipopar as tt "
    strSQL = strSQL & " where ta.cod_tip=tt.cod_tip and cod_cli='" & Me.Tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            msflxgrdAut.AddItem !nome_aut & Chr(9) & !nome_tip & Chr(9) & !cod_tip
            .MoveNext
        Loop
        .Close
    End With
    sstb.Tab = 0
bActive = True
    
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
    P_CentraForm Me
    
    resizeForm Me
    
    
    If Me.Tag = "" Then
        bCNPJ = True
    End If
    'stb.Tab = 0
    With msflxgrdAut
        .FixedCols = 0
        .Cols = 3
        .rows = 1
        .Row = 0
        
        .Col = 0: .Text = "Nome"
        .Col = 1: .Text = "Tipo"
        .Col = 2: .Text = "Cód."
        
        .ColWidth(0) = 3000
        .ColWidth(1) = 1000
        .ColWidth(2) = 500
        
    End With
    
    
    With msflxgrdRef
        .FixedCols = 0
        .Cols = 4
        .rows = 1
        .Row = 0
        
        .Col = 0: .Text = "Nome"
        .Col = 1: .Text = "Tipo"
        .Col = 2: .Text = "Cód."
        .Col = 3: .Text = "Fone."
        
        .ColWidth(0) = 3000
        .ColWidth(1) = 1000
        .ColWidth(2) = 500
        .ColWidth(3) = 1500
    End With
    
    
    P_PreencheTipoPar cmbxParentesco
    P_PreencheTipoPar cmbxParentRef
    If strContrMens_usu <> "S" Then
        chkMens.Enabled = False
        txtbxLimite.Enabled = False
    Else
        chkMens.Enabled = True
        txtbxLimite.Enabled = True
    End If
    With msflxgrdTel
        .FixedCols = 0
        .rows = 1
        .Cols = 2
        .Col = 0: .Text = "Fone:"
        .Col = 1: .Text = "Tipo:"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1200
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
    On Error Resume Next
    P_PreencheUF cmbxUF
    cmbxUF.Text = "GO"
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    If IsLoad(frmEdiClientes.Name) Then
        RefreshDataPosition frmEdiClientes.dtClientes
    End If
End Sub

Private Sub optnPessoa_Click(Index As Integer)
    If Index = 0 Then
        lblCpf.caption = "&CPF:"
        
        txtbxCnpj.MaxLength = 14
    Else
        lblCpf.caption = "&CNPJ:"
        
        txtbxCnpj.MaxLength = 18
    End If
    'txtbxCNPJ.Text = ""
End Sub




Private Sub msflxgrdAut_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete Then
        cmndbrnExclAut_Click
    End If
    
End Sub

Private Sub txtbxAutorizado_GotFocus()
    P_SelecTexto txtbxAutorizado
End Sub

Private Sub txtbxAutorizado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxAutorizado_LostFocus()
    P_AjustaTexto txtbxAutorizado
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
End Sub


Private Sub txtbxCelular_GotFocus()
    
End Sub

Private Sub txtbxCelular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCelular_LostFocus()
    txtbxCelular.Text = F_AjustaTelefone(txtbxCelular.Text)
    stb.Tab = 1
    txtbxAutorizado.SetFocus
End Sub

Private Sub txtbxCEP_GotFocus()
    P_SelecTexto txtbxCep
End Sub

Private Sub txtbxCEP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCEP_LostFocus()
    P_CEPLostFocus txtbxCep
End Sub


Private Sub txtbxCidade_GotFocus()
    P_SelecTexto txtbxCidade
End Sub

Private Sub txtbxCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCidade_LostFocus()
    P_AjustaTexto txtbxCidade
End Sub

Private Sub txtbxCNPJ_GotFocus()
    P_SelecTexto txtbxCnpj
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxCNPJ_LostFocus()
    
    
    
        If Not F_CPF(txtbxCnpj.Text) Then
            MsgBox "CPF Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCPF txtbxCnpj
    
End Sub



Private Sub txtbxDataNasc_GotFocus()
    P_SelecTexto txtbxDataNasc
End Sub

Private Sub txtbxDataNasc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataNasc_LostFocus()
    P_AjustaData txtbxDataNasc
    calcIdade
End Sub
Sub calcIdade()
    On Error Resume Next
    If txtbxDataNasc.Text = "" Then
        Exit Sub
    End If
    lblIdade.caption = DateDiff("YYYY", txtbxDataNasc.Text, Date)
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
End Sub


Private Sub txtbxLimite_GotFocus()
    P_SelecTexto txtbxLimite
End Sub

Private Sub txtbxLimite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxLimite_LostFocus()
    P_ValorLostFocus txtbxLimite
End Sub

Private Sub txtbxMae_GotFocus()
    P_SelecTexto txtbxMae
End Sub

Private Sub txtbxMae_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxMae_LostFocus()
    P_AjustaTexto txtbxMae
End Sub

Private Sub txtbxNomeRef_GotFocus()
    P_SelecTexto txtbxNomeRef
End Sub

Private Sub txtbxNomeRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNomeRef_LostFocus()
    P_AjustaTexto txtbxNomeRef
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
        SendKeys Chr(9)
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
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxRazao_LostFocus()
    P_AjustaTexto txtbxRazao
End Sub

Private Sub txtbxRef_GotFocus()
    
End Sub

Private Sub txtbxRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxRef_LostFocus()
    
End Sub

Private Sub txtbxRg_GotFocus()
    P_SelecTexto txtbxRg
End Sub

Private Sub txtbxRg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxRg_LostFocus()
    P_AjustaTexto txtbxRg
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
End Sub

Private Sub txtbxTelCom_GotFocus()
    
End Sub

Private Sub txtbxTelCom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxTelCom_LostFocus()
    
End Sub
Sub P_CarregaTelefone()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTipo As String
    msflxgrdTel.rows = 1
    strSQL = "select * from tab_telefones where cod_cli='" & Me.Tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            If !tipo_Tel = "1" Then
                sTipo = "1 - Residencial"
            ElseIf !tipo_Tel = "2" Then
                sTipo = "2 - Comercial"
            ElseIf !tipo_Tel = "3" Then
                sTipo = "3 - Celular"
            ElseIf !tipo_Tel = "4" Then
                sTipo = "4 - Recado"
            ElseIf !tipo_Tel = "5" Then
                sTipo = "5 - Fax"
            End If
            strSQL = !num_tel & Chr(9) & sTipo
            msflxgrdTel.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub
Sub P_CarregaRef()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTipo As String
    On Error GoTo trata_erro
    msflxgrdRef.rows = 1
    strSQL = "select * from tab_referencias where cod_cli='" & Me.Tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        Do While Not .EOF
            strSQL = !nome_ref & Chr(9)
            strSQL = strSQL & !cod_tip & Chr(9)
            strSQL = strSQL & !seq_ref & Chr(9)
            strSQL = strSQL & !tel_ref
            msflxgrdRef.AddItem strSQL
            .MoveNext
        Loop
        .Close
        
    End With
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub
Private Sub txtbxTel_GotFocus()
    P_SelecTexto txtbxTel
End Sub

Private Sub txtbxTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    
End Sub


Private Sub cmbxTipoTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnAddTel_Click()
    If F_Vazio(txtbxTel.Text) Then
        MsgBox "Infome o nº do Telefone!", vbExclamation
        txtbxTel.SetFocus
        Exit Sub
    End If
    If cmbxTipoTel.Text = "" Then
        MsgBox "Informe o tipo do Telefone!", vbExclamation
        cmbxTipoTel.SetFocus
        Exit Sub
    End If
    If F_InFlxgrd(msflxgrdTel, txtbxTel.Text, 0) > -1 Then
        MsgBox "Este telefone já se encontra Adicionado!", vbExclamation
        txtbxTel.SetFocus
        Exit Sub
    End If
    If cmndbtnAddTel.Tag = "" Then
        msflxgrdTel.AddItem txtbxTel.Text & Chr(9) & cmbxTipoTel.Text
    Else
        msflxgrdTel.Col = 0: msflxgrdTel.Text = txtbxTel.Text
        msflxgrdTel.Col = 1: msflxgrdTel.Text = cmbxTipoTel.Text
    End If
    cmndbtnAddTel.Tag = ""
    txtbxTel.SetFocus
End Sub

Private Sub cmndbtnEditTel_Click()
    With msflxgrdTel
        If .Row <= 0 Then
            Exit Sub
        End If
        .Col = 0
        txtbxTel.Text = .Text
        .Col = 1
        P_LocalizaOperacao .Text, cmbxTipoTel
        cmndbtnAddTel.Tag = .Row
    End With
End Sub

Private Sub cmndbtnExclTel_Click()
    With msflxgrdTel
        If .Row <= 0 Then
            Exit Sub
        End If
        If MsgBox("Deseja excluir o telefone?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
        If .rows = 2 Then
            .rows = 1
        Else
            .RemoveItem .Row
        End If
    End With
End Sub

Private Sub txtbxTelRef_GotFocus()
    P_SelecTexto txtbxTelRef
End Sub

Private Sub txtbxTelRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxTelRef_LostFocus()
    P_AjustaTelefone txtbxTelRef
End Sub
