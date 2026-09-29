VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadRep 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Representantes:"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14400
   Icon            =   "frmCadRep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   14400
   Begin VB.PictureBox pcBloquear 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   11880
      ScaleHeight     =   3705
      ScaleWidth      =   2385
      TabIndex        =   107
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdBloqImpr 
         Caption         =   "Impressão de Estoque"
         Height          =   855
         Left            =   120
         TabIndex        =   110
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdBloqDev 
         Caption         =   "Devolução"
         Height          =   855
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bloqueio/Liberação"
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
         TabIndex        =   108
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      ScaleHeight     =   225
      ScaleWidth      =   5385
      TabIndex        =   0
      Tag             =   "!"
      Top             =   6120
      Width           =   5415
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         Picture         =   "frmCadRep.frx":058A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações:"
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
         TabIndex        =   2
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox pcAval 
      Appearance      =   0  'Flat
      BackColor       =   &H00F79971&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   6360
      ScaleHeight     =   3705
      ScaleWidth      =   5385
      TabIndex        =   85
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtAvalistaAnotacoes 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   97
         Top             =   3000
         Width           =   5175
      End
      Begin VB.TextBox txtAvalistaTelefone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   95
         Top             =   2400
         Width           =   5175
      End
      Begin VB.TextBox txtAvalistaCpf 
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
         Left            =   135
         MaxLength       =   18
         TabIndex        =   90
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtAvalistaEndereco 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   1680
         Width           =   5175
      End
      Begin VB.TextBox txtAvalistaRg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         MaxLength       =   18
         TabIndex        =   88
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtAvalistaNome 
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
         Left            =   135
         MaxLength       =   50
         TabIndex        =   86
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Anotações 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Obs:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   2760
         Width           =   345
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefones:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   96
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avalista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2280
         TabIndex        =   94
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
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
         TabIndex        =   93
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "&Endereço:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   92
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "&RG:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   91
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
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
         TabIndex        =   87
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   6105
      TabIndex        =   84
      Top             =   7200
      Width           =   6135
      Begin VB.TextBox txtbxSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   79
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "Excluir"
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmndbtnAlterarSenha 
         Caption         =   "Alterar"
         Height          =   435
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   120
         Width           =   915
      End
      Begin VB.TextBox txtbxConfSenha 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   81
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar Senha:"
         Height          =   195
         Left            =   1920
         TabIndex        =   80
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   0
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      TabIndex        =   32
      Top             =   240
      Width           =   6135
      Begin VB.TextBox txtbxLimite 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         MaxLength       =   40
         TabIndex        =   49
         Text            =   "0,00"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.ComboBox cmbxStatus 
         Height          =   315
         ItemData        =   "frmCadRep.frx":0B14
         Left            =   1200
         List            =   "frmCadRep.frx":0B21
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtbxNatural 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   47
         Top             =   3840
         Width           =   2775
      End
      Begin VB.ComboBox cmbxSexo 
         Height          =   315
         ItemData        =   "frmCadRep.frx":0B5C
         Left            =   4800
         List            =   "frmCadRep.frx":0B66
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtbxConjuge 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   45
         Top             =   3480
         Width           =   4815
      End
      Begin VB.TextBox txtbxEndPais 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   44
         Top             =   3120
         Width           =   4815
      End
      Begin VB.TextBox txtbxMae 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   43
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox txtbxRg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   18
         TabIndex        =   42
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtbxDataNasc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   41
         Top             =   2040
         Width           =   990
      End
      Begin VB.TextBox txtbxCEP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   14
         TabIndex        =   40
         Text            =   "75.000-000"
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox cmbxUF 
         Height          =   315
         Left            =   5400
         Style           =   1  'Simple Combo
         TabIndex        =   39
         Text            =   "cmbxUF"
         Top             =   1665
         Width           =   615
      End
      Begin VB.TextBox txtbxCidade 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   38
         Text            =   "Anápolis"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox txtbxBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   37
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtbxEndereco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   36
         Top             =   960
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
         TabIndex        =   35
         Top             =   600
         Width           =   2535
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
         TabIndex        =   34
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtbxPai 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2760
         Width           =   4815
      End
      Begin VB.CommandButton cmndbtnCobrador 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         Picture         =   "frmCadRep.frx":0B7F
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblLimite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Limite($):"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4080
         TabIndex        =   68
         Top             =   4200
         Width           =   735
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
         TabIndex        =   67
         Top             =   4200
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
         TabIndex        =   66
         Top             =   3840
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
         TabIndex        =   65
         Top             =   3840
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
         TabIndex        =   64
         Top             =   3480
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
         TabIndex        =   63
         Top             =   3120
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
         TabIndex        =   62
         Top             =   2040
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
         TabIndex        =   61
         Top             =   2040
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
         TabIndex        =   60
         Top             =   2400
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
         Left            =   3840
         TabIndex        =   59
         Top             =   600
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
         TabIndex        =   58
         Top             =   2040
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
         TabIndex        =   57
         Top             =   2040
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
         TabIndex        =   56
         Top             =   1665
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
         TabIndex        =   55
         Top             =   1680
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
         TabIndex        =   54
         Top             =   1320
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
         TabIndex        =   53
         Top             =   960
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
         TabIndex        =   52
         Top             =   600
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
         TabIndex        =   51
         Top             =   240
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
         TabIndex        =   50
         Top             =   2760
         Width           =   1095
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
      TabIndex        =   29
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
         Picture         =   "frmCadRep.frx":0F09
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   30
         Top             =   0
         Width           =   210
      End
      Begin VB.Label lblNomeRep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dados do Representante"
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
         TabIndex        =   31
         Top             =   0
         Width           =   2040
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
      TabIndex        =   26
      Tag             =   "!"
      Top             =   4920
      Width           =   6135
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         Picture         =   "frmCadRep.frx":1086
         ScaleHeight     =   270
         ScaleWidth      =   285
         TabIndex        =   27
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
         TabIndex        =   28
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.TextBox txtbxObs 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   6360
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   6360
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Foto:"
      Height          =   2775
      Left            =   240
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmndbtnFoto 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   2640
         Picture         =   "frmCadRep.frx":149A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00B4DFE0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   2385
         TabIndex        =   21
         Tag             =   "!"
         Top             =   240
         Width           =   2415
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            Picture         =   "frmCadRep.frx":154F
            ScaleHeight     =   180
            ScaleWidth      =   210
            TabIndex        =   22
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
            TabIndex        =   23
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.PictureBox pc3X4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2055
         Index           =   0
         Left            =   120
         ScaleHeight     =   2025
         ScaleWidth      =   2385
         TabIndex        =   19
         Top             =   480
         Width           =   2415
         Begin VB.PictureBox p 
            BorderStyle     =   0  'None
            Height          =   1695
            Index           =   0
            Left            =   480
            ScaleHeight     =   1695
            ScaleWidth      =   1575
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comissões:"
      Height          =   2175
      Left            =   6360
      TabIndex        =   11
      Top             =   3840
      Width           =   5415
      Begin MSDataGridLib.DataGrid gridCom 
         Height          =   1095
         Left            =   120
         TabIndex        =   111
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1931
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "nome_tip"
            Caption         =   "Nome"
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
            DataField       =   "comissao_tip"
            Caption         =   "Comissão"
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
               ColumnWidth     =   2729,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   975,118
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmndbtnExclCom 
         Height          =   315
         Left            =   4680
         Picture         =   "frmCadRep.frx":1604
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir Comissão"
         Top             =   1770
         Width           =   315
      End
      Begin VB.CommandButton cmndbtnGravaComiss 
         Height          =   315
         Left            =   4320
         Picture         =   "frmCadRep.frx":39CB
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salvar Comissão"
         Top             =   1770
         Width           =   315
      End
      Begin VB.TextBox txtbxComiss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   13
         Text            =   "0"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtbxTipoCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Tecle Enter"
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Comissão(%):"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3030
         TabIndex        =   17
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Co&brança:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture10 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   14370
      TabIndex        =   3
      Top             =   7980
      Width           =   14400
      Begin VB.TextBox txtUltConf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Left            =   120
         TabIndex        =   106
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdLembreteAcerto 
         Caption         =   "Lembrete Acerto"
         Height          =   735
         Left            =   10080
         Picture         =   "frmCadRep.frx":3D55
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdLembrete 
         Caption         =   "Lembrete"
         Height          =   735
         Left            =   8640
         Picture         =   "frmCadRep.frx":41CF
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdUltimaConferencia 
         Caption         =   "Docs a Receber"
         Height          =   735
         Left            =   1440
         Picture         =   "frmCadRep.frx":4649
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Consignados"
         Height          =   735
         Left            =   2880
         Picture         =   "frmCadRep.frx":4733
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Perfil"
         Height          =   735
         Left            =   4320
         Picture         =   "frmCadRep.frx":481D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Histórico"
         Height          =   735
         Left            =   5760
         Picture         =   "frmCadRep.frx":50E7
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "ESC - &Fecha&r"
         Height          =   735
         Left            =   12960
         Picture         =   "frmCadRep.frx":51D1
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSalvar 
         Caption         =   "F7 - Sal&var"
         Height          =   735
         Left            =   11520
         Picture         =   "frmCadRep.frx":533C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   7200
         Picture         =   "frmCadRep.frx":56C6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conferido em:"
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      TabIndex        =   69
      Top             =   5040
      Width           =   6135
      Begin VB.PictureBox pcMotivoBloq 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   2640
         ScaleHeight     =   1305
         ScaleWidth      =   3465
         TabIndex        =   99
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txtMotivoBloq 
            Appearance      =   0  'Flat
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo do Bloqueio"
            Height          =   195
            Left            =   960
            TabIndex        =   100
            Top             =   0
            Width           =   1365
         End
      End
      Begin VB.TextBox txtbxTel 
         Height          =   285
         Left            =   0
         TabIndex        =   74
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox cmbxTipoTel 
         Height          =   315
         ItemData        =   "frmCadRep.frx":5AD3
         Left            =   1560
         List            =   "frmCadRep.frx":5AE6
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton cmndbtnAddTel 
         Height          =   315
         Left            =   4440
         Picture         =   "frmCadRep.frx":5B2C
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1680
         Width           =   315
      End
      Begin VB.CommandButton cmndbtnEditTel 
         Height          =   315
         Left            =   4800
         Picture         =   "frmCadRep.frx":60B6
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1680
         Width           =   315
      End
      Begin VB.CommandButton cmndbtnExclTel 
         Height          =   315
         Left            =   5160
         Picture         =   "frmCadRep.frx":6AB8
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1680
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrdTel 
         Height          =   1335
         Left            =   0
         TabIndex        =   75
         Top             =   120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº:"
         Height          =   195
         Left            =   0
         TabIndex        =   77
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   1560
         TabIndex        =   76
         Top             =   1440
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmCadRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoFotos  As ADODB.Recordset
Attribute adoFotos.VB_VarHelpID = -1
Dim WithEvents ado  As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bCNPJ As Boolean 'para saber se o CNPJ é válido
Dim bActive As Boolean


Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim cFile As New Collection
Dim ai As New ADODB.Stream
Dim WithEvents adoCom As ADODB.Recordset
Attribute adoCom.VB_VarHelpID = -1
Dim Indice As Integer

Sub buscaFoto()
    Dim sFile As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    Set cFile = Nothing
    ai.type = adTypeBinary
    strSQL = "select seq_fot,foto_fot from tab_fotorep where cod_rep='" & Me.tag & "' order by seq_fot desc"
    'c.Open strNomeBanco
    'Set r = Conn_Client.Execute(strSQL)
    adoFotos.Open strSQL, Conn_Client
    If Dir(sFoto_Temp, vbDirectory) = "" Then
        FileSystem.MkDir sFoto_Temp
    End If
    
    With adoFotos
        Do While Not .EOF
            sFile = sFoto_Temp & "\F" & !seq_fot & ".tmp"
            ai.Open
            ai.Write !foto_fot
            ai.SaveToFile (sFile)
            ai.Close
            cFile.Add sFile
            .MoveNext
        Loop
        
        .Close
    
    End With
    
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
    strSQL = strSQL & " and tipo_ter='R' "
    Conn.Execute strSQL
    With msflxgrdTel
        For iCount = 1 To .rows - 1
            .row = iCount
            .col = 0: sFone = .text
            .col = 1: sTipo = Left(.text, 1)
            If Clientes.F_GravaTelefone(Me.tag, sFone, sTipo, "R") Then
            End If
            
        Next
    End With
End Sub



Private Sub cmbxEstadoCiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub


Private Sub chkBloqDev_Click()
    Dim i As Integer
    If Me.tag = "" Then Exit Sub
    If bActive = False Then Exit Sub
    i = chkBloqDev.Value
    
    If i = 0 Then
        i = 1
    Else
        i = 0
    End If
    
    
    
    
    
End Sub

Private Sub chkBloqImpress_Click()
    
End Sub

Private Sub cmbxSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxStatus_click()
    If cmbxStatus.ListIndex <> 0 Then
        txtbxLimite.Enabled = False
        lblLimite.Enabled = False
    Else
        If strAlterarComRepr_Usu = "S" Then
        
            txtbxLimite.Enabled = True
            lblLimite.Enabled = True
        Else
            txtbxLimite.Enabled = False
            lblLimite.Enabled = False
        End If
    End If
    If Left(cmbxStatus.text, 1) = "2" Then
        pcMotivoBloq.Visible = True
        
    Else
        pcMotivoBloq.Visible = False
    End If
End Sub

Private Sub cmbxStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbxTipoCob_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub cmdBloqDev_Click()
    If Me.tag = "" Then Exit Sub
    If strAlterarComRepr_Usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If Not F_senha(True) Then
        Exit Sub
    End If
    
    If cmdBloqDev.tag = "" Then
        AlteraRepresentanteBloqueio Me.tag, "", "S"
        cmdBloqDev.tag = "X"
        cmdBloqDev.caption = "Liberar Devolução"
    Else
        AlteraRepresentanteBloqueio Me.tag, "", "N"
        cmdBloqDev.tag = ""
        cmdBloqDev.caption = "Bloquear Devolução"
    End If
    
    
    
    
End Sub

Private Sub cmdBloqImpr_Click()
    If Me.tag = "" Then Exit Sub
    If strAlterarComRepr_Usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    
    If Not F_senha(True) Then
        Exit Sub
    End If
    
    If cmdBloqImpr.tag = "" Then
        AlteraRepresentanteBloqueio Me.tag, "S", ""
        cmdBloqImpr.tag = "X"
        cmdBloqImpr.caption = "Liberar Impressão"
    Else
        AlteraRepresentanteBloqueio Me.tag, "N", ""
        cmdBloqImpr.tag = ""
        cmdBloqImpr.caption = "Bloquear Impressão"
    End If
End Sub

Private Sub cmdLembrete_Click()
    buscaLembrete Me.tag, True
End Sub

Private Sub cmdLembreteAcerto_Click()
    buscaLembreteAcerto Me.tag, True
End Sub

Private Sub cmndbtnAlterarSenha_Click()
    If Me.tag = "" Then
        MsgBox "Primeiro grave o representante!", vbExclamation
        Exit Sub
    End If
    If strCadcli_usu <> "S" Then
        MsgBox "Não tem permissão!", vbExclamation
        Exit Sub
    End If
    If txtbxSenha.text <> txtbxConfSenha.text Then
        MsgBox "Senhas não conferem!", vbExclamation
        Exit Sub
    End If
    If Not F_SenhaRep(Me.tag) Then
        Exit Sub
    End If
    P_AlteraSenhaRep Me.tag, txtbxSenha.text
    MsgBox "Senha alterada!", vbInformation
    
    
End Sub


Private Sub cmndbtnCobrador_Click()
    Dim s As setor
    s = Setores.F_BuscaSetor(txtbxBairro.text)
    If s.cod <> "" Then
        txtbxBairro.text = s.nome
        txtbxBairro.tag = s.cod
        txtbxCidade.text = s.Cidade
        cmbxUf.text = s.UF
        txtbxBairro.BackColor = vbGreen
        txtbxCEP.SetFocus
    End If
    
End Sub

Private Sub cmndbtnExcl_Click()
    If Me.tag = "" Then
        MsgBox "Primeiro grave o representante!", vbExclamation
        Exit Sub
    End If
    If strCadcli_usu <> "S" Then
        MsgBox "Não tem permissão!", vbExclamation
        Exit Sub
    End If
    If Not F_senha() Then
        Exit Sub
    End If
    
    If MsgBox("Confirma Exclusão da Senha do Consignado?", vbYesNo + vbQuestion) Then
        P_AlteraSenhaRep Me.tag, ""
        MsgBox "Senha alterada!", vbInformation
    End If
End Sub

Private Sub cmndbtnExclCom_Click()
    
    
    If strAlterarComRepr_Usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    If MsgBox("Confirma exclusão da comissão do Representante?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    DeletaRepTip ado!cod_tip, Me.tag
    buscaRepTip
    
End Sub

Private Sub cmndbtnFoto_click()
    frmGravaFoto.sTab = "tab_fotorep"
    frmGravaFoto.sCampo = "cod_rep"
    frmGravaFoto.sTipo = "R"
    
    frmGravaFoto.caption = Me.tag & "-" & Me.caption
    frmGravaFoto.tag = Me.tag
    frmGravaFoto.Show 1
    deletaArquivos
    
    buscaFoto
    
End Sub

Private Sub cmndbtnGravaComiss_Click()
    If strAlterarComRepr_Usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    If txtbxTipoCob.tag = "" Then
        MsgBox "Informe o tipo!", vbExclamation
        Exit Sub
    End If
    If F_GravaRepTip(txtbxTipoCob.tag, Me.tag, txtbxComiss.text) Then
        buscaRepTip
    End If
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnSalvar_Click()
    Dim strSQL As String
    Dim iCount As Integer
    If strCadastrarRepr_Usu <> "S" Then
        MsgBox "Você não tem permissão para cadastrar Representantes", vbInformation
        Exit Sub
    End If
    
    
    If txtbxRazao.text = "" Then
        MsgBox "Informe um nome para o Representante!", vbExclamation
        Exit Sub
        
    ElseIf F_Vazio(txtbxCNPJ.text) And F_Vazio(txtbxRg.text) Then
        
            MsgBox "Informe o CPF  ou CI do Representante !", vbExclamation
            txtbxCNPJ.SetFocus
            Exit Sub
    
    ElseIf txtbxEndereco.text = "" Then
        MsgBox "Informe o endereço do Representante!", vbExclamation
        Exit Sub
    ElseIf txtbxBairro.tag = "" Then
        MsgBox "informe o Bairro do Representante!", vbExclamation
        
        txtbxBairro.SetFocus
        Exit Sub
    ElseIf txtbxCidade.text = "" Then
        MsgBox "Informe a Cidade do Representante!", vbExclamation
        Exit Sub
    ElseIf cmbxUf.text = "" Then
        MsgBox "Informe o Estado de Origem do Representante!", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "" Then
        If F_BuscaUmCampo("cod_rep", "tab_representantes", "cpf_rep='" & txtbxCNPJ.text & "'") <> "" Then
            If MsgBox("Já existe um representante com este CPF/CNPJ!. Deseja continuar?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        ElseIf F_BuscaUmCampo("cod_Rep", "tab_representantes", "nome_rep='" & txtbxRazao.text & "'") <> "" Then
            MsgBox "Já existe um representante com este Nome!", vbCritical
            Exit Sub
        
        End If
    
        strSQL = representante.F_GravaRepresentante(txtbxRazao.text, txtbxLimite.text, txtbxEndereco, _
            txtbxBairro.tag, txtbxCidade.text, cmbxUf.text, txtbxCEP.text, txtbxDataNasc.text, txtbxNatural.text, _
            txtbxRg.text, txtbxCNPJ.text, Left(cmbxSexo.text, 1), txtbxConjuge.text, txtbxPai.text, txtbxMae.text, txtbxEndPais.text, txtbxObs.text, "", _
            txtAvalistaNome.text, txtAvalistaCpf.text, txtAvalistaRg.text, txtAvalistaEndereco.text, txtAvalistaTelefone.text, _
            txtAvalistaAnotacoes.text)
        
        
        
        If strSQL <> "" Then
            MsgBox "Representante Cadastrado com sucesso!", vbInformation
            P_Limpa
            
            txtbxRazao.text = ""
            Me.tag = ""
        Else
            MsgBox "Erro", vbExclamation
            Exit Sub
        End If
    Else
        If MsgBox("Confirma alteração do Representante " & txtbxRazao.text & " ?", vbYesNo + vbQuestion) = vbYes Then
            
            
            If Left(cmbxStatus.text, 1) = "2" And Trim(txtMotivoBloq.text) = "" Then
                MsgBox "Informe o motivo do bloqueio!", vbExclamation
                Exit Sub
            End If
            
            If representante.F_AlteraRepresentante(Me.tag, txtbxRazao.text, txtbxLimite.text, txtbxEndereco.text, txtbxBairro.tag, _
                 txtbxCidade.text, cmbxUf.text, txtbxCEP.text, txtbxDataNasc.text, txtbxNatural.text, _
                 txtbxRg.text, txtbxCNPJ.text, Left(cmbxSexo.text, 1), txtbxConjuge.text, _
                    txtbxPai.text, txtbxMae.text, txtbxEndPais.text, txtbxObs.text, Left(cmbxStatus.text, 1), "", _
                    txtAvalistaNome.text, txtAvalistaCpf.text, txtAvalistaRg.text, txtAvalistaEndereco.text, txtAvalistaTelefone.text, _
            txtAvalistaAnotacoes.text, txtMotivoBloq.text) Then
                P_GravaTelefones
                MsgBox "Representante alterado com sucesso!", vbInformation
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
       ' buscaFoto
    End If
    
    
    

    
bActive = True
    
End Sub
Sub buscaRepTip()
    Dim strSQL As String
    
    strSQL = "select tt.nome_tip,tr.* from tab_tipocob as tt, tab_reptip as tr"
    strSQL = strSQL & " where tt.cod_tip=tr.cod_tip "
    strSQL = strSQL & " and tr.cod_rep='" & Me.tag & "' "
    abreObjetoAdo ado, strSQL, gridCom
    
    
    
End Sub
Sub calcIdade()
    lblIdade.caption = F_CalcIdade(txtbxDataNasc.text)
End Sub
Sub carrega()
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    Dim sLog As String
    
    If strAlterarComRepr_Usu = "S" Then
        pcBloquear.Enabled = True
    Else
        pcBloquear.Enabled = False
    End If
    
    
    strSQL = "select tr.*,ts.nome_set,uf_cid,nome_cid,tc.nome_cid from tab_representantes as tr "
    strSQL = strSQL & " left join tab_setores as ts on ts.cod_set=tr.cod_set "
    strSQL = strSQL & " left join tab_cidades as tc on tc.cod_cid=ts.cod_cid "
    strSQL = strSQL & " where cod_rep = '" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        'Do Until .EOF
            'P_LocalizaOperacao !cod_tip, cmbxTipoCob
            If IsNull(!cod_set) Or !cod_set = "" Or !cod_set = "00001" Then
                txtbxBairro.text = TiraNull(!bairro_rep)
                txtbxCidade.text = !cidade_rep
                If Format(!uf_rep) <> "" Then cmbxUf.text = !uf_rep
                txtbxBairro.BackColor = vbRed
            Else
                txtbxBairro.text = !nome_set
                txtbxBairro.tag = !cod_set
                txtbxCidade.text = !nome_cid
                cmbxUf.text = !uf_cid
                txtbxBairro.BackColor = vbGreen
            End If
            txtbxCEP.text = TiraNull(!cep_rep)
            
            txtbxCNPJ.text = TiraNull(!cpf_rep)
            txtbxRg.text = TiraNull(!rg_rep)
            txtbxEndereco.text = TiraNull(!endereco_rep)
            txtbxMae.text = Format(!mae_rep)
            txtbxPai.text = Format(!pai_rep)
            txtbxRazao.text = TiraNull(!nome_rep)
            txtbxDataNasc.text = Format(!nasc_rep)
            txtbxEndPais.text = TiraNull(!endpais_rep)
            txtbxNatural.text = TiraNull(!natural_rep)
            txtMotivoBloq.text = Format(!motivo_bloqueio)
            
            If !bloqdev = "S" Then
                cmdBloqDev.caption = "Liberar Devolução"
                cmdBloqDev.tag = "X"
            Else
                cmdBloqDev.caption = "Bloquear Devolução"
                cmdBloqDev.tag = ""
            End If
            If !bloqimpr = "S" Then
                cmdBloqImpr.caption = "Liberar Impressão"
                cmdBloqImpr.tag = "X"
            Else
                cmdBloqImpr.caption = "Bloquear Impressão"
                cmdBloqImpr.tag = ""
            End If
            
            
            
            If Not IsNull(!senha_rep) Then
                txtbxSenha.text = (!senha_rep)
                txtbxConfSenha.text = !senha_rep
            End If
            
            P_LocalizaOperacao TiraNull(!sexo_rep), cmbxSexo
            
            txtbxConjuge.text = TiraNull(!conjuge_rep)
            
            calcIdade
            txtbxObs.text = Format(!obs_rep)
            txtbxRg.text = !rg_rep
            lblNomeRep.caption = !cod_Rep & " - " & !nome_rep
            txtbxLimite.text = FormatVal(!limite_rep)
            If strAlterarComRepr_Usu = "S" Then
                txtbxLimite.Enabled = True
            Else
                txtbxLimite.Enabled = False
            End If
            
            txtAvalistaAnotacoes.text = Format(!avalista_anotacoes)
            txtAvalistaCpf.text = Format(!avalista_cpf)
            txtAvalistaRg.text = Format(!avalista_rg)
            txtAvalistaEndereco.text = Format(!avalista_endereco)
            txtAvalistaNome.text = Format(!avalista_nome)
            txtAvalistaTelefone = Format(!avalista_telefones)
            
            
            
            


            
            
            P_LocalizaOperacao !status_rep, cmbxStatus
            If !status_rep = "2" Then
                pcMotivoBloq.Visible = True
            Else
                pcMotivoBloq.Visible = False
            End If
            P_CarregaTelefone
            buscaRepTip
            '.MoveNext
        'Loop
    End With
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
        bActive = False
    'P_PreencheTipoCob cmbxTipoCob
    
    Set adoFotos = New ADODB.Recordset
    Set adoCom = New ADODB.Recordset
    Set ado = New ADODB.Recordset
    resizeForm Me
    
    
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








Private Sub txtbxBairro_Change()
    txtbxBairro.tag = ""
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



Private Sub txtbxConfSenha_GotFocus()
    P_SelecTexto txtbxConfSenha
End Sub

Private Sub txtbxConfSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxConfSenha_LostFocus()
    P_AjustaTexto txtbxConfSenha
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

Private Sub txtbxSenha_GotFocus()
    P_SelecTexto txtbxSenha
End Sub

Private Sub txtbxSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxSenha_LostFocus()
    P_AjustaTexto txtbxSenha
End Sub

Private Sub txtbxTel_LostFocus()
    txtbxTel.text = F_AjustaTelefone(txtbxTel.text)
End Sub

Sub P_CarregaTelefone()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipo As String
    msflxgrdTel.rows = 1
    strSQL = "select * from tab_telefones where cod_cli='" & Me.tag & "' and tipo_ter='R' "
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
        
        If MsgBox("Deseja salvar alterações?", vbYesNo + vbQuestion) = vbYes Then
            P_GravaTelefones
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

Private Sub txtbxTipoCob_GotFocus()
    P_SelecTexto txtbxTipoCob
End Sub
Sub buscaTipoCob()
    Dim t As tipocob
    t = F_BuscaTipoCob("")
    If t.cod_tip <> "" Then
        txtbxTipoCob.text = t.nome_tip
        txtbxTipoCob.tag = t.cod_tip
        txtbxComiss.text = FormatVal(t.comiss_tip)
    End If
End Sub
Private Sub txtbxTipoCob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaTipoCob
    End If
End Sub

Private Sub txtUltConf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
              
        
    End If
End Sub

Private Sub txtUltConf_LostFocus()
    P_AjustaData txtUltConf
    If txtUltConf.text <> "" And Me.tag <> "" Then
            strSQL = "update tab_representantes set ultimaconferencia=" & F_GravaData(txtUltConf.text)
            strSQL = strSQL & " where cod_rep='" & Me.tag & "'"
            Conn.Execute strSQL
        End If
End Sub
