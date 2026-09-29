VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCadCli 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14910
   Icon            =   "frmCadCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   14910
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEDFF3&
      ForeColor       =   &H80000008&
      Height          =   295
      Left            =   6720
      ScaleHeight     =   270
      ScaleWidth      =   8145
      TabIndex        =   110
      Tag             =   "!"
      Top             =   5040
      Width           =   8175
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações Gerais"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   111
         Top             =   0
         Width           =   2160
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BEDFF3&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   295
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   6465
      TabIndex        =   106
      Tag             =   "!"
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox chkClienteEspecial 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cliente Especial"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   112
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dados do Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   120
         TabIndex        =   107
         Top             =   0
         Width           =   1830
      End
   End
   Begin VB.TextBox txtbxObs 
      Appearance      =   0  'Flat
      Height          =   2535
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   95
      Top             =   5400
      Width           =   8175
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6495
      Begin VB.CommandButton cmdCEP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   480
         Picture         =   "frmCadCli.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   3060
         Width           =   255
      End
      Begin VB.ComboBox cmbxSexo 
         Height          =   315
         ItemData        =   "frmCadCli.frx":0914
         Left            =   4200
         List            =   "frmCadCli.frx":091E
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Tecle enter para selecionar um bairro"
         Top             =   2760
         Width           =   6015
      End
      Begin VB.TextBox txtComplemento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtLogradouro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   94
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox txtbxBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Tecle enter para selecionar um bairro"
         Top             =   2160
         Width           =   3735
      End
      Begin VB.CommandButton cmndbtnCobrador 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         Picture         =   "frmCadCli.frx":0937
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox txtbxMae 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   25
         Top             =   3960
         Width           =   3135
      End
      Begin VB.TextBox txtbxRg 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2520
         MaxLength       =   18
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxDataNasc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   23
         Top             =   3360
         Width           =   990
      End
      Begin VB.TextBox txtbxCEP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   14
         TabIndex        =   21
         Text            =   "75.000-000"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Anápolis"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtbxCNPJ 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaxLength       =   18
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtbxRazao 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   6255
      End
      Begin VB.TextBox txtbxPai 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   27
         Top             =   3960
         Width           =   3015
      End
      Begin VB.CommandButton cmdCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         Picture         =   "frmCadCli.frx":0CC1
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gênero"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4200
         TabIndex        =   97
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lblCodCidIBGE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
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
         Height          =   195
         Left            =   6120
         TabIndex        =   96
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Município IBGE(NFe)"
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
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4215
         TabIndex        =   12
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3375
         TabIndex        =   10
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2775
         TabIndex        =   93
         Top             =   3120
         Width           =   435
      End
      Begin VB.Label lblIdade 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2400
         TabIndex        =   55
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Idade:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2235
         TabIndex        =   54
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label lblMae 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome da &Mãe:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   3720
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "&RG:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Nasc:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   22
         Top             =   3120
         Width           =   810
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "CE&P:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Ci&dade:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4215
         TabIndex        =   16
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label lblBairro 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Set&or:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label lblCpf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblPai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do &Pai:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3375
         TabIndex        =   26
         Top             =   3720
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   14880
      TabIndex        =   69
      Top             =   7995
      Width           =   14910
      Begin VB.CommandButton cmdInternet 
         Caption         =   "Internet"
         Height          =   855
         Left            =   4200
         Picture         =   "frmCadCli.frx":104B
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "&Débitos"
         Height          =   855
         Left            =   2160
         Picture         =   "frmCadCli.frx":1E39
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "F7 - Sal&var"
         Height          =   855
         Left            =   10200
         Picture         =   "frmCadCli.frx":1F83
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Esc - &Fecha&r"
         Height          =   855
         Left            =   12360
         Picture         =   "frmCadCli.frx":230D
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnHist 
         Caption         =   "&Histórico"
         Height          =   855
         Left            =   120
         Picture         =   "frmCadCli.frx":289B
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   6465
      TabIndex        =   57
      Top             =   5040
      Width           =   6495
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00BEDFF3&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   295
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   6495
         TabIndex        =   108
         Tag             =   "!"
         Top             =   0
         Width           =   6495
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contatos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   109
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.TextBox txtbxTel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cmbxTipoTel 
         Height          =   315
         ItemData        =   "frmCadCli.frx":2E25
         Left            =   1680
         List            =   "frmCadCli.frx":2E38
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmndbtnAddTel 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   3600
         TabIndex        =   33
         Top             =   1980
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnEditTel 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   5520
         TabIndex        =   35
         Top             =   1980
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnExclTel 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   4560
         TabIndex        =   34
         Top             =   1980
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrdTel 
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2566
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblDataAlt 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Left            =   1440
         TabIndex        =   92
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label lblDataCad 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         TabIndex        =   91
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Cadastro:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   90
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ultima Alteração:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   89
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   1680
         TabIndex        =   31
         Top             =   1800
         Width           =   375
      End
   End
   Begin TabDlg.SSTab sstb 
      CausesValidation=   0   'False
      Height          =   4695
      Left            =   6720
      TabIndex        =   36
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Preferências:"
      TabPicture(0)   =   "frmCadCli.frx":2E7E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLimite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtbxLimite"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmndbtnSpc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frCob"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbxStatus"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkCreditoSuspenso"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Informações Profissionais"
      TabPicture(1)   =   "frmCadCli.frx":2E9A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtbxTelefoneTrab"
      Tab(1).Control(1)=   "txtbxCepTrab"
      Tab(1).Control(2)=   "txtbxTrabalho"
      Tab(1).Control(3)=   "txtbxSetTrab"
      Tab(1).Control(4)=   "txtbxEndTrab"
      Tab(1).Control(5)=   "cmndbtnSetTrab"
      Tab(1).Control(6)=   "lblCidadeTrab"
      Tab(1).Control(7)=   "Label11"
      Tab(1).Control(8)=   "Label17"
      Tab(1).Control(9)=   "Label16"
      Tab(1).Control(10)=   "Label7"
      Tab(1).Control(11)=   "lblBairroTrab"
      Tab(1).Control(12)=   "Label15"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Referências"
      TabPicture(2)   =   "frmCadCli.frx":2EB6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Anotações"
      TabPicture(3)   =   "frmCadCli.frx":2ED2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdAddAnot"
      Tab(3).Control(1)=   "txtAnot"
      Tab(3).Control(2)=   "gridAnot"
      Tab(3).Control(3)=   "Picture9"
      Tab(3).Control(4)=   "Label38"
      Tab(3).ControlCount=   5
      Begin VB.CommandButton cmdAddAnot 
         Caption         =   "Adicionar"
         Height          =   375
         Left            =   -69240
         TabIndex        =   118
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtAnot 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   -74880
         TabIndex        =   117
         Top             =   720
         Width           =   5415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridAnot 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   115
         Top             =   1440
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4048
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CheckBox chkCreditoSuspenso 
         Caption         =   "Suspender Crédito"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   3360
         Width           =   1695
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   -73560
         ScaleHeight     =   1545
         ScaleWidth      =   6585
         TabIndex        =   99
         Top             =   3720
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtbxEndereco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   113
            Top             =   1200
            Width           =   5175
         End
         Begin VB.TextBox txtbxBairroAntigo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   101
            Tag             =   "Tecle enter para selecionar um bairro"
            Top             =   240
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox txtbxCidadeAntiga 
            Appearance      =   0  'Flat
            BackColor       =   &H00B4DFE0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   100
            Text            =   "Anápolis"
            Top             =   600
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "&Endereço:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   114
            Top             =   960
            Width           =   765
         End
         Begin VB.Label lblUfAntigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00B4DFE0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "&UF:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4200
            TabIndex        =   104
            Top             =   600
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cid. Antiga:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   103
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "&Bairro Antigo:"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   102
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbxStatus 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCadCli.frx":2EEE
         Left            =   120
         List            =   "frmCadCli.frx":2EF8
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   720
         Width           =   3255
      End
      Begin VB.Frame frCob 
         Caption         =   "Informações de Cobrança"
         Height          =   1935
         Left            =   120
         TabIndex        =   75
         Top             =   1200
         Width           =   7935
         Begin VB.TextBox txtbxOrgao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            MaxLength       =   40
            TabIndex        =   76
            Tag             =   "00000"
            Text            =   "Geral"
            Top             =   480
            Width           =   4575
         End
         Begin VB.CommandButton cmndbtnOrgao 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6600
            Picture         =   "frmCadCli.frx":2F19
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   480
            Width           =   255
         End
         Begin VB.ComboBox cmbxCobranca 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmCadCli.frx":32A3
            Left            =   120
            List            =   "frmCadCli.frx":32B0
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrança:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   900
            Width           =   480
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Órgão"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2040
            TabIndex        =   86
            Top             =   240
            Width           =   435
         End
         Begin VB.Label lblGrupo 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
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
            Left            =   720
            TabIndex        =   85
            Top             =   900
            Width           =   525
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3600
            TabIndex        =   84
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label lblEndGrupo 
            AutoSize        =   -1  'True
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
            Left            =   3600
            TabIndex        =   83
            Top             =   1620
            Width           =   195
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label lblBairroGrupo 
            AutoSize        =   -1  'True
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
            Left            =   240
            TabIndex        =   81
            Top             =   1680
            Width           =   195
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label lblCidadeGrupo 
            AutoSize        =   -1  'True
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
            Left            =   960
            TabIndex        =   79
            Top             =   1200
            Width           =   195
         End
      End
      Begin VB.CommandButton cmndbtnSpc 
         Caption         =   "SPC"
         Height          =   375
         Left            =   5400
         TabIndex        =   74
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "&Referências:"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   67
         Top             =   360
         Width           =   7095
         Begin VB.CommandButton cmndbtnAddRef 
            Caption         =   "&Adicionar"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox txtbxNomeRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   59
            Top             =   2280
            Width           =   2655
         End
         Begin VB.CommandButton cmndbtnExclRef 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   2520
            TabIndex        =   66
            Top             =   2640
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnModifRef 
            Caption         =   "Modificar"
            Height          =   375
            Left            =   1320
            TabIndex        =   65
            Top             =   2640
            Width           =   1095
         End
         Begin VB.ComboBox cmbxParentRef 
            Height          =   315
            ItemData        =   "frmCadCli.frx":32EE
            Left            =   2880
            List            =   "frmCadCli.frx":3304
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox txtbxTelRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5280
            TabIndex        =   63
            Top             =   2280
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdRef 
            Height          =   1815
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3201
            _Version        =   393216
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "N&ome:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Parentesco:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2880
            TabIndex        =   60
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "&Nº:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5280
            TabIndex        =   62
            Top             =   2040
            Width           =   225
         End
      End
      Begin VB.TextBox txtbxTelefoneTrab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         MaxLength       =   40
         TabIndex        =   51
         Top             =   3000
         Width           =   4455
      End
      Begin VB.TextBox txtbxCepTrab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70080
         MaxLength       =   40
         TabIndex        =   53
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtbxTrabalho 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         MaxLength       =   40
         TabIndex        =   42
         ToolTipText     =   "Informe o nome da empresa ou local de trabalho"
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox txtbxSetTrab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         MaxLength       =   20
         TabIndex        =   46
         ToolTipText     =   "Tecle enter para selecionar um bairro"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox txtbxEndTrab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         MaxLength       =   40
         TabIndex        =   44
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox txtbxLimite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   40
         Text            =   "0,00"
         Top             =   720
         Width           =   1110
      End
      Begin VB.CommandButton cmndbtnSetTrab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70560
         Picture         =   "frmCadCli.frx":3352
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anotação:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   116
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "S&tatus:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblCidadeTrab 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cidade"
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
         Left            =   -74760
         TabIndex        =   49
         Top             =   2400
         Width           =   5775
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74760
         TabIndex        =   48
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Telefone:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74760
         TabIndex        =   50
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -70125
         TabIndex        =   52
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74760
         TabIndex        =   41
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblBairroTrab 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Setor:"
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
         Height          =   195
         Left            =   -74760
         TabIndex        =   45
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74760
         TabIndex        =   43
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblLimite 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "&Limite  (R$)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   39
         Top             =   480
         Width           =   825
      End
   End
   Begin VB.Menu mnOpt 
      Caption         =   "Opções"
      Visible         =   0   'False
      Begin VB.Menu mnAnot 
         Caption         =   "Anotações"
         Begin VB.Menu mnExcluirAnot 
            Caption         =   "Excluir"
         End
      End
   End
End
Attribute VB_Name = "frmCadCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Dim bActive As Boolean
Dim bAlterado As Boolean
Sub P_Salva()
    P_GravaAutorizados
    P_GravaReferencia
End Sub

Sub P_GravaAutorizados()
    Dim strSQL As String
    Dim iCount As Integer
    strSQL = "delete from tab_autorizados where cod_cli='" & Me.tag & "' "
    Conn.Execute strSQL
    With msflxgrdAut
        For iCount = 1 To .rows - 1
            
            .row = iCount
            strSQL = "insert into tab_autorizados(cod_cli,seq_aut,nome_aut,status_aut,cod_tip)values("
            strSQL = strSQL & "'" & Me.tag & "',"
            strSQL = strSQL & iCount & ","
            .col = 0
            strSQL = strSQL & "'" & .Text & "',"
            strSQL = strSQL & "'1', "
            .col = 2
            strSQL = strSQL & "'" & .Text & "')"
            Conn.Execute strSQL
        Next
    End With
End Sub
Sub P_GravaReferencia()
    Dim strSQL As String
    Dim iCount As Integer
    strSQL = "delete from tab_referencias where cod_cli='" & Me.tag & "' "
    Conn.Execute strSQL
    With msflxgrdRef
        For iCount = 1 To .rows - 1
            
            .row = iCount
            strSQL = "insert into tab_referencias(cod_cli,seq_ref,nome_ref,cod_tip,tel_ref)values("
            strSQL = strSQL & "'" & Me.tag & "',"
            strSQL = strSQL & iCount & ","
            .col = 0
            strSQL = strSQL & "'" & .Text & "',"
            .col = 2
            strSQL = strSQL & "'" & Trim(Left(.Text, 2)) & "',"
            .col = 3
            strSQL = strSQL & "'" & .Text & "')"
            Conn.Execute strSQL
        Next
    End With
End Sub

Sub P_Altera()
    Dim strSQL As String
    strSQL = "update tab_clientes set "
    strSQL = strSQL & "nome_cli='" & txtbxRazao.Text & "',"
    strSQL = strSQL & "cpf_Cli='" & txtbxCNPJ.Text & "',"
    strSQL = strSQL & "rg_cli='" & txtbxRg.Text & "',"
    strSQL = strSQL & "logradouro_cli='" & txtLogradouro.Text & "',"
    strSQL = strSQL & "numero_cli='" & txtNumero.Text & "',"
    strSQL = strSQL & "complemento_cli='" & txtcomplemento.Text & "',"
    strSQL = strSQL & "endereco_Cli='" & txtbxEndereco.Text & "',"
    strSQL = strSQL & "bairro_Cli='" & txtbxBairro.Text & "',"
    strSQL = strSQL & "cidade_Cli='" & txtbxCidade.Text & "',"
    strSQL = strSQL & "uf_Cli='" & cmbxUf.Text & "',"
    strSQL = strSQL & "cep_Cli='" & txtbxCEP.Text & "',"
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
    
         
         
    strSQL = strSQL & " where cod_Cli = '" & Me.tag & "'"
    P_GravaAutorizados
    Conn.Execute strSQL
End Sub

Private Sub cmbxEstadoCiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub chkMens_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxParentesco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxParentRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxStatus_click()
    If cmbxStatus.ListIndex = 1 Then
        txtbxLimite.Enabled = False
        lblLimite.Enabled = False
    Else
        txtbxLimite.Enabled = True
        lblLimite.Enabled = True
    End If
End Sub

Private Sub cmbxStatus_GotFocus()
    sstb.Tab = 0
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

Private Sub cmndbrnExclAut_Click()
    If msflxgrdAut.rows > 1 Then
    
    
        If MsgBox("Confirma Exclusão?", vbYesNo + vbCritical) = vbYes Then
            If msflxgrdAut.rows = 2 Then
                msflxgrdAut.rows = 1
            Else
                msflxgrdAut.RemoveItem msflxgrdAut.row
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
    If cmndbtnAddAut.tag = "" Then
        'If F_InFlxgrd(msflxgrdAut, txtbxAutorizado.Text, 0) > -1 Then
        '    MsgBox "O Nome já se encontra incluído!", vbCritical
        
       '     txtbxAutorizado.SetFocus
       '     Exit Sub
       ' End If
        msflxgrdAut.AddItem txtbxAutorizado.Text & Chr(9) & Mid(cmbxParentesco.Text, 6, Len(cmbxParentesco.Text)) & Chr(9) & Left(cmbxParentesco.Text, 2)
        cmndbtnAddAut.tag = ""
        txtbxAutorizado.SetFocus
    Else
        msflxgrdAut.row = msflxgrdAut.tag
        msflxgrdAut.col = 0: msflxgrdAut.Text = txtbxAutorizado.Text
        msflxgrdAut.col = 1: msflxgrdAut.Text = Mid(cmbxParentesco.Text, 6, Len(cmbxParentesco.Text))
        msflxgrdAut.col = 2: msflxgrdAut.Text = Left(cmbxParentesco.Text, 2)
    End If
End Sub


Private Sub cmdAddAnot_Click()
    If Me.tag = "" Then
        MsgBox "Insira anotações só depois do cliente estar gravado!", vbExclamation
        Exit Sub
    End If
    Clientes.AdicionaAnot Me.tag, txtAnot
    buscaAnot
    txtAnot.SetFocus
End Sub
Sub buscaAnot()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_clientesanot where cod_cli='" & Me.tag & "' and excluido='N' order by id desc "
    abreObjetoAdo rPesq, strSQL

    GridUtil.AjustaGridData gridAnot, 4, 1, 0, "id,anotacao,data,usuario", "id,anotacao,datahora,cod_usu", "0,4000,1800,1000", rPesq, "", False, vbActiveTitleBar
    rPesq.Close
End Sub

Private Sub cmdCEP_Click()
    Dim e As Endereco
     
     e = busca_cep(txtbxCEP.Text)
     If e.cep <> "" Then
        txtLogradouro.Text = e.Logradouro
        
        txtCidade.Text = e.ibge
        cmdCidade_Click
        txtbxBairro.Text = e.bairro
        cmndbtnCobrador_Click
     End If
End Sub

Private Sub cmdCidade_Click()
    Dim c As Cidade
    c = BuscaCidadesIBGE(txtCidade.Text)
    If c.cod_cid <> "" Then
        txtCidade.Text = c.nome_cid
        txtCidade.tag = c.cod_cid
        lblCodCidIBGE = c.cod_cid & "/" & c.uf_cid
    End If
End Sub

Private Sub cmdInternet_Click()
    If Me.tag = "" Then Exit Sub
    frmClienteInternet.txtCPF.Text = F_SoNumero(txtbxCNPJ.Text)
    frmClienteInternet.tag = Me.tag
    frmClienteInternet.Show 1
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
    If cmndbtnAddRef.tag = "" Then
        If F_InFlxgrd(msflxgrdRef, txtbxNomeRef.Text, 0) > -1 Then
            MsgBox "O Nome já se encontra incluído!", vbCritical
        
            txtbxNomeRef.SetFocus
            Exit Sub
        End If
        msflxgrdRef.AddItem txtbxNomeRef.Text & Chr(9) & cmbxParentRef.Text & Chr(9) & Left(cmbxParentRef.Text, 2) & Chr(9) & txtbxTelRef.Text
        cmndbtnAddRef.tag = ""
        txtbxNomeRef.SetFocus
    Else
        msflxgrdRef.row = msflxgrdRef.tag
        msflxgrdRef.col = 0: msflxgrdRef.Text = txtbxNomeRef.Text
        msflxgrdRef.col = 1: msflxgrdRef.Text = Mid(cmbxParentRef.Text, 6, Len(cmbxParentRef.Text))
        msflxgrdRef.col = 2: msflxgrdRef.Text = Left(cmbxParentRef.Text, 2)
        msflxgrdRef.col = 3: msflxgrdRef.Text = txtbxTelRef.Text
    End If

End Sub



Private Sub cmndbtnCobrador_Click()
    buscaSetor
End Sub
Sub buscaSetor()
    Dim s As setor
    s = Setores.F_BuscaSetor(txtbxBairro.Text)
    
    If s.cod <> "" Then
        txtbxBairro.Text = s.nome
        txtbxBairro.tag = s.cod
        
        lblbairro.caption = "&" & s.tipo
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub cmndbtnDeb_Click()
    On Error GoTo trata_Erro
    frmParcelasTer.txtbxNomeCli.Text = txtbxRazao.Text
    frmParcelasTer.txtbxNomeCli.tag = txtbxRazao.tag
    frmParcelasTer.tag = txtbxRazao.tag
    frmParcelasTer.Show 1
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnEditAut_Click()
    With msflxgrdAut
        If .row <= 0 Then
            Exit Sub
        End If
        .col = 0
        txtbxAutorizado.Text = .Text
        .col = 2
        'P_LocalizaOperacao .Text, cmbxParentesco
        cmndbtnAddAut.tag = .row
        msflxgrdAut.tag = .row
    End With
End Sub

Private Sub cmndbtnExclRef_Click()
        If msflxgrdRef.rows > 1 Then
    
    
        If MsgBox("Confirma Exclusão?", vbYesNo + vbCritical) = vbYes Then
            If msflxgrdRef.rows = 2 Then
                msflxgrdRef.rows = 1
            Else
                msflxgrdRef.RemoveItem msflxgrdRef.row
            End If
        End If
    Else
        MsgBox "Não há o que excluir!", vbCritical
    End If

End Sub

Private Sub cmndbtnGrupoCLi_Click()
    Dim g As grupoCli
    g = Clientes.F_BuscaGrupoCli()
    If g.cod_gru <> "" And g.cod_gru <> txtbxGrupoCli.tag Then
        txtbxGrupoCli.Text = g.nome_gru
        txtbxGrupoCli.tag = g.cod_gru
        bAlterado = True
    End If
End Sub

Private Sub cmndbtnHist_Click()
    If Me.tag = "" Then
    Else
        P_histCli Me.tag
    End If
End Sub

Private Sub cmndbtnModifRef_Click()
        With msflxgrdRef
        If .row <= 0 Then
            Exit Sub
        End If
        .col = 0
        txtbxNomeRef.Text = .Text
        .col = 2
        P_LocalizaOperacao .Text, cmbxParentRef
        .col = 3
        txtbxTelRef.Text = .Text
        cmndbtnAddRef.tag = .row
        msflxgrdRef.tag = .row
    End With

End Sub

Private Sub cmndbtnOrgao_Click()
    Dim o As orgao.orgao
    
    o = orgao.F_BuscaOrgao(txtbxOrgao.Text, txtbxOrgao.tag)
    If o.cod_org <> "" Then
        txtbxOrgao.Text = o.nome_org
        txtbxOrgao.tag = o.cod_org
        lblGrupo.caption = o.nome_gru
        lblGrupo.tag = o.cod_gru
        lblEndGrupo.caption = o.endereco_org
        lblBairroGrupo.caption = o.nome_set
        lblCidadeGrupo.caption = o.nome_cid
        'lblUF.caption = o.uf
        txtbxEndTrab.Text = o.endereco_org
        txtbxSetTrab.Text = o.nome_set
        txtbxSetTrab.tag = o.cod_set
        
    End If
End Sub
Sub carregaOrgao()
    Dim o As orgao.orgao
    
    If txtbxOrgao.tag = "" Then
        Exit Sub
    End If
    o = orgao.F_BuscaOrgaoPorCod(txtbxOrgao.tag)
    If o.cod_org <> "" Then
        txtbxOrgao.Text = o.nome_org
        txtbxOrgao.tag = o.cod_org
        lblGrupo.caption = o.nome_gru
        lblEndGrupo.caption = o.endereco_org
        lblBairroGrupo.caption = o.nome_set
        lblCidadeGrupo.caption = o.nome_cid & "/" & o.UF
        
        
    End If
End Sub
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    If strCadcli_usu <> "S" Then
        MsgBox "Você não tem permissão pra cadastrar Cliente!", vbExclamation
        Exit Sub
    End If
    If cmbxSexo.Text = "" Then
        MsgBox "Escolha o sexo!", vbExclamation
        cmbxSexo.SetFocus
        Exit Sub
    End If
    If Left(cmbxCobranca.Text, 1) = "2" Then
        If txtbxOrgao.tag = "" Then
            MsgBox "Para este tipo de cobrança, o cliente deve estar vinculado a um órgão!", vbExclamation
            Exit Sub
        End If
    End If
    If Not bCNPJ Then
        MsgBox "CPF Inválido!!!", vbExclamation
        Exit Sub
    Exit Sub
    End If
    If txtbxRazao.Text = "" Then
        MsgBox "Informe um nome para o Cliente!", vbExclamation
        Exit Sub
        
    ElseIf F_Vazio(txtbxCNPJ.Text) And F_Vazio(txtbxRg.Text) Then
        
            MsgBox "Informe o CPF  ou CI do Cliente!", vbExclamation
            txtbxCNPJ.SetFocus
            Exit Sub
        
    
    
    ElseIf txtbxEndereco.Text = "" Then
        'MsgBox "Informe o endereço do Cliente!", vbExclamation
        'Exit Sub
    ElseIf txtbxBairro.tag = "" Then
        MsgBox "informe o Bairro do Cliente!,vbExclamation  "
        Exit Sub
    
    
    End If
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("cod_cli", "tab_clientes", "cpf_cli='" & txtbxCNPJ.Text & "' or cpf_cli='" & F_SoNumero(txtbxCNPJ.Text) & "'") <> "" Then
            MsgBox "Já existe um cliente com este CPF/CNPJ!", vbCritical
            Exit Sub
            'ElseIf F_BuscaUmCampo("cod_cli", "tab_clientes", "nome_cli='" & txtbxRazao.Text & "'") <> "" Then
            'MsgBox "Já existe um cliente com este Nome/Razão Social!", vbCritical
            'Exit Sub
        
        End If
    
        Me.tag = Clientes.F_GravaCli(txtbxRazao.Text, txtbxDataNasc.Text, txtbxCNPJ.Text, txtbxRg.Text, _
            txtbxEndereco.Text, txtLogradouro.Text, txtNumero.Text, txtcomplemento.Text, txtbxBairro.Text, txtbxCidade.Text, "GO", txtbxCEP.Text, txtbxObs.Text, _
            txtbxMae.Text, txtbxPai.Text, "S", txtbxLimite.Text, txtbxBairro.tag, txtbxTrabalho.Text, _
            txtbxEndTrab.Text, txtbxSetTrab.tag, txtbxTelefoneTrab.Text, txtbxCepTrab.Text, _
            "", Left(cmbxCobranca.Text, 1), txtbxOrgao.tag, txtCidade.tag, Left(cmbxSexo.Text, 1), txtemail.Text)
            
            P_gravatel
            P_GravaReferencia
            API.sincronizar_clientes , Me.tag
        If MsgBox("Cliente Cadastrado com sucesso! Deseja cadastrar outro?", vbYesNo + vbQuestion) = vbYes Then
            P_Limpa
            txtbxRazao.Text = ""
            Me.tag = ""
            'msflxgrdAut.rows = 1
            msflxgrdRef.rows = 1
            msflxgrdTel.rows = 1
        Else
            
        End If
        
        
    Else
        If MsgBox("Confirma alteração do cliente " & txtbxRazao.Text & " ?", vbYesNo + vbQuestion) = vbYes Then
            
            If F_BuscaUmCampo("cod_cli", "tab_clientes", "cpf_cli='" & F_SoNumero(txtbxCNPJ.Text) & "' and cod_cli<>'" & Me.tag & "'") <> "" Then
            MsgBox "Já existe um cliente com este CPF/CNPJ! você está tentando alterar para um CPF que já existe!", vbCritical
            Exit Sub
            'ElseIf F_BuscaUmCampo("cod_cli", "tab_clientes", "nome_cli='" & txtbxRazao.Text & "'") <> "" Then
            'MsgBox "Já existe um cliente com este Nome/Razão Social!", vbCritical
            'Exit Sub
        
        End If
            Dim cpf As String
            cpf = F_SoNumero(txtbxCNPJ.Text)
            If Clientes.F_AlteraCli(Me.tag, txtbxRazao.Text, txtbxDataNasc.Text, txtbxCNPJ.Text, txtbxCEP.Text, CStr(Date), txtbxObs.Text, Left(cmbxStatus.Text, 1), txtbxRg.Text, _
                txtbxMae.Text, txtbxPai.Text, txtbxLimite.Text, _
                txtbxEndTrab.Text, strCod_Usu, IIf(Len(cpf) = 11, "F", "J"), "0", txtbxBairro.tag, lblGrupo.tag, Left(cmbxCobranca.Text, 1), _
                txtbxTrabalho.Text, txtbxEndTrab.Text, txtbxSetTrab.tag, txtbxTelefoneTrab.Text, txtbxCepTrab.Text, txtbxOrgao.tag, CStr(Date), _
                Left(cmbxSexo.Text, 1), txtemail.Text, IIf(chkCreditoSuspenso.Value = 1, "S", "N"), "", txtCidade.tag, txtbxTel.Text, _
                txtLogradouro.Text, txtNumero.Text, txtcomplemento.Text, "9", IIf(chkClienteEspecial.Value = 1, "S", "N")) Then
                
                P_gravatel
                P_GravaReferencia
                MsgBox "Cliente alterado com sucesso!", vbInformation
                sincronizar_clientes "", Me.tag
            Else
                MsgBox "Erro"
            
            End If
                
                        
        End If
    End If
    
End Sub
Sub P_Limpa()
    Dim i As Integer
    For i = 0 To Me.Controls.Count - 1
        If LCase(Left(Me.Controls(i).Name, 3)) = "txt" Then
            Me.Controls(i).Text = ""
        End If
    Next
   
    'cmbxUF.ListIndex = -1
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub





Private Sub cmndbtnSetTrab_Click()
    Dim s As setor
    s = Setores.F_BuscaSetor(txtbxSetTrab.Text)
    
    If s.cod <> "" Then
        txtbxSetTrab.Text = s.nome
        txtbxSetTrab.tag = s.cod
        lblCidadeTrab.caption = s.Cidade & " (" & s.UF & ")"
    
    
    End If
    
    
    
    
End Sub

Private Sub cmndbtnSPC_Click()
    P_ClienteSpc Me.tag
End Sub

Private Sub Form_Activate()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim s As setor
    If bActive Then Exit Sub
    If bUsaCobranca = False Then
        frCob.Visible = False
    End If
    If Me.tag <> "" Then
        strSQL = "select tc.*,cid.nome_cid as ibgenome,cid.uf_cid as ibgeuf, cid.cod_cid as ibgecodigo,"
        strSQL = strSQL & " tci.nome_cid,ts.nome_set,tipo_set,tci.uf_cid "
        strSQL = strSQL & " from tab_clientes as tc right join tab_setores as ts "
        strSQL = strSQL & "  on ts.cod_set=tc.cod_set inner join tab_cidades as tci "
        
        strSQL = strSQL & " on tci.cod_cid=ts.cod_cid "
        strSQL = strSQL & " left join cidades as cid on tc.cod_cid=cid.cod_cid "
        strSQL = strSQL & " where tc.cod_cli = '" & Me.tag & "' "
        'strSQL = strSQL & " and tc.cod_set=ts.cod_set "
        
        abreObjetoAdo rPesq, strSQL
        With rPesq
            
                If !creditosuspenso_cli = "S" Then
                    chkCreditoSuspenso.Value = 1
                End If
                txtbxBairroAntigo.Text = !bairro_cli
                txtbxCidadeAntiga.Text = !cidade_cli
                lblbairro.caption = "&" & !tipo_set
                txtLogradouro.Text = TiraNull(!Logradouro_cli)
                txtcomplemento.Text = Format(!complemento_cli)
                txtNumero.Text = Format(!numero_cli)
                If !especial_cli = "S" Then
                    chkClienteEspecial.Value = 1
                Else
                    chkClienteEspecial.Value = 0
                End If
                lblUfAntigo.caption = !uf_cli
                Me.caption = "Editando Cliente: " & !cod_cli & "-" & !nome_cli
                'txtbxGrupoCli.tag = !cod_gru
                'txtbxGrupoCli.Text = F_BuscaUmCampo("nome_gru", "tab_grupocli", "cod_gru='" & !cod_gru & "'")
                txtbxOrgao.tag = !cod_org
                carregaOrgao
                txtbxOrgao.tag = !cod_org
                P_LocalizaOperacao !cobranca_cli, cmbxCobranca
                txtbxCEP.Text = !cep_cli
                P_LocalizaOperacao !sexo_cli, cmbxSexo
                txtbxCidade.Text = !nome_cid
                txtbxBairro.Text = !nome_set
                txtbxBairro.tag = !cod_set
                s = F_RetornaSetor(!setortrab_cli)
                txtbxTrabalho.Text = !localtrab_cli
                txtbxEndTrab.Text = !endtrab_cli
                txtbxTelefoneTrab.Text = !teltrab_cli
                txtbxCepTrab.Text = !ceptrab_cli
                txtbxSetTrab.tag = !setortrab_cli
                txtbxSetTrab.Text = s.nome
                lblCidadeTrab.caption = s.Cidade & "(" & s.UF & ")"
                txtCidade.Text = TiraNull(!ibgenome)
                txtCidade.tag = TiraNull(!ibgecodigo)
                lblCodCidIBGE = TiraNull(!ibgecodigo) & "/" & TiraNull(!ibgeuf)
                txtemail.Text = Format(!email_cli)
                
                txtbxCNPJ.Text = !cpf_cli
                txtbxRg.Text = !rg_cli
                txtbxEndereco.Text = !endereco_cli
                txtbxMae.Text = Format(!mae_cli)
                txtbxPai.Text = Format(!pai_cli)
                txtbxRazao.Text = !nome_cli
                txtbxDataNasc.Text = Format(!datanasc_cli)
'
                calcIdade
                txtbxObs.Text = F_BuscaUmCampo("obs_cli", "tab_clientes", "cod_cli='" & Me.tag & "'")
                txtbxRg.Text = !rg_cli
                If !mens_cli = "S" Then
 '                   chkMens.Value = 1
                Else
'                    chkMens.Value = 0
                End If
                txtbxLimite.Text = FormatVal(!limite_cli)
                'frfIf !uf_cli <> "" Then cmbxUF.Text = !uf_cli
                
                
                
                
                
                lblDataAlt.caption = (TiraNull(!dataalt_cli))
                lblDataCad.caption = TiraNull(!data_cli)


                
                
                P_LocalizaOperacao !status_cli, cmbxStatus
                If strLiberarCli_usu <> "S" Then
                    cmbxStatus.Enabled = False
                End If
                
                P_CarregaTelefone
                P_CarregaRef
                P_CarregaAut
                buscaAnot
        End With
    End If
    
    
    
    
    sstb.Tab = 0
bActive = True
    
End Sub
Sub P_CarregaAut()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ta.*,tt.nome_tip from tab_autorizados as ta, tab_tipopar as tt "
    strSQL = strSQL & " where ta.cod_tip=tt.cod_tip and cod_cli='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            msflxgrdAut.AddItem !nome_aut & Chr(9) & !nome_tip & Chr(9) & !cod_tip
            .MoveNext
        Loop
        .Close
    End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmndbtnSalvar_Click
    ElseIf KeyCode = vbKeyF3 Then
        sstb.Tab = 0
        sstb.SetFocus
        Sendkeys Chr(9)
    ElseIf KeyCode = vbKeyF4 Then
        sstb.Tab = 1
        sstb.SetFocus
        Sendkeys Chr(9)
    ElseIf KeyCode = vbKeyF5 Then
        sstb.Tab = 2
        sstb.SetFocus
        Sendkeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
    bActive = False
 
    
    
   
    
    
    If Me.tag = "" Then
        bCNPJ = True
    End If
    'stb.Tab = 0
'    With msflxgrdAut
'        .FixedCols = 0
'        .Cols = 3
'        .rows = 1
'        .row = 0
'
'        .Col = 0: .Text = "Nome"
'        .Col = 1: .Text = "Tipo"
'        .Col = 2: .Text = "Cód."
'
'        .ColWidth(0) = 3000
'        .ColWidth(1) = 1000
'        .ColWidth(2) = 500
'
'    End With
    
    With msflxgrdRef
        .FixedCols = 0
        .Cols = 4
        .rows = 1
        .row = 0
        
        .col = 0: .Text = "Nome"
        .col = 1: .Text = "Tipo"
        .col = 2: .Text = "Cód."
        .col = 3: .Text = "Fone."
        
        .ColWidth(0) = 3000
        .ColWidth(1) = 1000
        .ColWidth(2) = 500
        .ColWidth(3) = 1500
    End With
    
    
    'P_PreencheTipoPar cmbxParentesco
    P_PreencheTipoPar cmbxParentRef
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
        .col = 0: .Text = "Fone:"
        .col = 1: .Text = "Tipo:"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1200
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
    End With
    On Error Resume Next
    'P_PreencheUF cmbxUF
    cmbxUf.Text = "GO"
    P_PreencheTipoPar cmbxParentRef
  '  resizeForm Me
  P_CentraForm Me
End Sub

Private Sub Form_Resize()
    'P_CentraObejct frControle, Me
    
End Sub

Private Sub optnPessoa_Click(Index As Integer)
    If Index = 0 Then
        lblcpf.caption = "&CPF:"
        
        txtbxCNPJ.MaxLength = 14
    Else
        lblcpf.caption = "&CNPJ:"
        
        txtbxCNPJ.MaxLength = 18
    End If
    'txtbxCNPJ.Text = ""
End Sub




Private Sub msflxgrdAut_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete Then
        cmndbrnExclAut_Click
    End If
    
End Sub


Private Sub txtbxAutorizado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub gridAnot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If MsgBox("Confirma exclusão da anotação?", vbYesNo + vbQuestion) = vbYes Then
            gridAnot.col = 0
            
            Clientes.MarcaExcluidoAnot gridAnot.Text
            buscaAnot
        End If
    End If
End Sub

Private Sub gridAnot_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnAnot
    End If
End Sub

Private Sub mnExcluirAnot_Click()
    gridAnot_KeyDown vbKeyDelete, 0
End Sub

Private Sub txtAnot_GotFocus()
    P_SelecTexto txtAnot
    
End Sub

Private Sub txtAnot_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAnot_LostFocus()
    P_AjustaTexto txtAnot, 3
End Sub

Private Sub txtbxBairro_GotFocus()
    P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCobrador_Click
    End If
End Sub

Private Sub txtbxBairro_LostFocus()
    P_AjustaTexto txtbxBairro
End Sub


Private Sub txtbxCelular_GotFocus()
    
End Sub

Private Sub txtbxCelular_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCelular_LostFocus()
    txtbxCelular.Text = F_AjustaTelefone(txtbxCelular.Text)
    stb.Tab = 1
    txtbxAutorizado.SetFocus
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


Private Sub txtbxCepTrab_GotFocus()
    P_SelecTexto txtbxCepTrab
End Sub

Private Sub txtbxCepTrab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxCepTrab_LostFocus()
    P_CEPLostFocus txtbxCepTrab
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
    txtbxCNPJ.Text = F_SoNumero(txtbxCNPJ.Text)
    
    If Len(txtbxCNPJ.Text) > 11 Then
        If Not F_CGC(txtbxCNPJ.Text) Then
            MsgBox "CNPJ Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCGC txtbxCNPJ
        If Me.tag = "" Then
            P_ConfereCPF
        End If
    Else
    
    
        If Not F_CPF(txtbxCNPJ.Text) Then
            MsgBox "CPF Inválido! Não é possível concluir a Operação sem um CNPJ válido", vbCritical
            bCNPJ = False
        Else
            bCNPJ = True
        End If
        P_FormataCPF txtbxCNPJ
        If Me.tag = "" Then
            P_ConfereCPF
        End If
    End If
    VerificaCPF
End Sub
Sub VerificaCPF()
    Dim strSQL As String
    If Me.tag <> "" Then
        Exit Sub
    End If
    If txtbxCNPJ.Text = "" Then Exit Sub
    strSQL = F_BuscaUmCampo("cod_cli", "tab_clientes", "cpf_cli='" & txtbxCNPJ.Text & "'")
    
    
    If strSQL <> "" Then
        If MsgBox("Este CPF/CNPJ está cadastrado para o cliente " & strSQL & _
            Chr(13) & "Deseja carregar esses dados?", vbYesNo + vbQuestion) = vbYes Then
            Me.tag = strSQL
            bActive = False
            Form_Activate
        End If
    End If
    
    
End Sub
Sub P_ConfereCPF()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_clientes where cpf_cli"
    
    
    
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
  lblIdade.caption = F_CalcIdade(txtbxDataNasc.Text)
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


Private Sub txtbxEndTrab_GotFocus()
    P_SelecTexto txtbxEndTrab
End Sub

Private Sub txtbxEndTrab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEndTrab_LostFocus()
    P_AjustaTexto txtbxEndTrab
End Sub

Private Sub txtbxGrupoCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnGrupoCLi_Click
    End If
End Sub

Private Sub txtbxLimite_GotFocus()
    P_SelecTexto txtbxLimite
End Sub

Private Sub txtbxLimite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
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
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNomeRef_LostFocus()
    P_AjustaTexto txtbxNomeRef
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then
    '    cmndbtnSalvar.SetFocus
    'End If
End Sub

Private Sub txtbxObs_LostFocus()
    txtbxObs.Text = F_AjustaTexto(txtbxObs.Text, False)
End Sub

Private Sub txtbxOrgao_Change()
    txtbxOrgao.tag = ""
End Sub

Private Sub txtbxOrgao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnOrgao_Click
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

Private Sub txtbxRef_GotFocus()
    
End Sub

Private Sub txtbxRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxRef_LostFocus()
    
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

Private Sub txtbxSetTrab_GotFocus()
    P_SelecTexto txtbxSetTrab
End Sub

Private Sub txtbxSetTrab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnSetTrab_Click
    End If
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.Text = F_AjustaTelefone(txtbxTel.Text)
End Sub

Private Sub txtbxTelCom_GotFocus()
    
End Sub

Private Sub txtbxTelCom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTelCom_LostFocus()
    
End Sub
Sub P_CarregaTelefone()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipo As String
    msflxgrdTel.rows = 1
    strSQL = "select * from tab_telefones where cod_cli='" & Me.tag & "' and tipo_ter='C' "
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
Sub P_CarregaRef()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipo As String
    On Error GoTo trata_Erro
    msflxgrdRef.rows = 1
    strSQL = "select tr.*,tt.nome_tip from tab_referencias as tr, tab_tipopar as tt "
    strSQL = strSQL & " where cod_cli='" & Me.tag & "' "
    strSQL = strSQL & " and tr.cod_tip=tt.cod_tip "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = !nome_ref & Chr(9)
            strSQL = strSQL & !nome_tip & Chr(9)
            strSQL = strSQL & !seq_ref & Chr(9)
            strSQL = strSQL & !tel_ref
            msflxgrdRef.AddItem strSQL
            .MoveNext
        Loop
        .Close
        
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
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
    If cmndbtnAddTel.tag = "" Then
        msflxgrdTel.AddItem txtbxTel.Text & Chr(9) & cmbxTipoTel.Text
    Else
        msflxgrdTel.col = 0: msflxgrdTel.Text = txtbxTel.Text
        msflxgrdTel.col = 1: msflxgrdTel.Text = cmbxTipoTel.Text
    End If
    cmndbtnAddTel.tag = ""
    txtbxTel.SetFocus
End Sub
Sub P_gravatel()
      Dim strSQL As String
      Dim sFone As String
      Dim sTipo As String
      strSQL = "delete from tab_telefones where cod_cli='" & Me.tag & "' "
      strSQL = strSQL & " and tipo_ter='C' "
      Conn.Execute strSQL
      
    
    
    With msflxgrdTel
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0
            sFone = .Text
            .col = 1
            sTipo = Left(.Text, 1)
            If Clientes.F_GravaTelefone(Me.tag, sFone, sTipo, "C") Then
            End If
        Next
    End With
      
      
End Sub
Private Sub cmndbtnEditTel_Click()
    With msflxgrdTel
        If .row <= 0 Then
            Exit Sub
        End If
        .col = 0
        txtbxTel.Text = .Text
        .col = 1
        P_LocalizaOperacao .Text, cmbxTipoTel
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

Private Sub txtbxTelefoneTrab_GotFocus()
    P_SelecTexto txtbxTelefoneTrab
End Sub

Private Sub txtbxTelefoneTrab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTelefoneTrab_LostFocus()
    P_AjustaTelefone txtbxTelefoneTrab
End Sub

Private Sub txtbxTelRef_GotFocus()
    P_SelecTexto txtbxTelRef
End Sub

Private Sub txtbxTelRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTelRef_LostFocus()
    P_AjustaTelefone txtbxTelRef
End Sub
Sub P_ValorGotFocus(objValor As Object)
    'Procedimento de ajuste de foco
    P_SelecTexto objValor
End Sub

Private Sub txtbxTrabalho_GotFocus()
    sstb.Tab = 1
    P_SelecTexto txtbxTrabalho
End Sub

Private Sub txtbxTrabalho_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxTrabalho_LostFocus()
    P_AjustaTexto txtbxTrabalho
End Sub

Private Sub txtCidade_GotFocus()
    P_SelecTexto txtCidade
End Sub

Private Sub txtCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCidade_Click
    End If
End Sub

Private Sub txtComplemento_GotFocus()
    P_SelecTexto txtcomplemento
End Sub

Private Sub txtComplemento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtComplemento_LostFocus()
    P_AjustaTexto txtcomplemento
    If txtLogradouro.Text <> "" Then
        txtbxEndereco.Text = txtLogradouro.Text & " " & txtcomplemento.Text
        If txtNumero.Text <> "" Then
            txtbxEndereco.Text = txtbxEndereco.Text & txtbxEndereco.Text & " N. " & txtNumero.Text
        End If
    End If
End Sub

Private Sub txtEmail_GotFocus()
    P_SelecTexto txtemail
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtEmail_LostFocus()
    P_AjustaTexto txtemail, 2
End Sub

Private Sub txtLogradouro_GotFocus()
    P_SelecTexto txtLogradouro
End Sub

Private Sub txtLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtLogradouro_LostFocus()
    P_AjustaTexto txtLogradouro
    If txtLogradouro.Text <> "" Then
        txtbxEndereco.Text = txtLogradouro.Text & " " & txtcomplemento
        If txtNumero.Text <> "" Then
            txtbxEndereco.Text = txtbxEndereco.Text & txtbxEndereco.Text & " N. " & txtNumero.Text
        End If
    End If
End Sub

Private Sub txtNumero_GotFocus()
    P_SelecTexto txtNumero
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtNumero_LostFocus()
    P_AjustaTexto txtNumero
    If txtLogradouro.Text <> "" Then
        txtbxEndereco.Text = txtLogradouro.Text & " " & txtcomplemento
        If txtNumero.Text <> "" Then
            txtbxEndereco.Text = txtbxEndereco.Text & txtbxEndereco.Text & " N. " & txtNumero.Text
        End If
    End If
End Sub
