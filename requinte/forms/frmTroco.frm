VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTroco 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Troco"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   13710
   Icon            =   "frmTroco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcCheque 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2520
      ScaleHeight     =   705
      ScaleWidth      =   1665
      TabIndex        =   134
      Top             =   9240
      Visible         =   0   'False
      Width           =   1695
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   0
         ScaleHeight     =   1305
         ScaleWidth      =   7305
         TabIndex        =   136
         Top             =   120
         Width           =   7335
         Begin VB.CommandButton cmndbtnBuscaBanco 
            Height          =   285
            Left            =   3000
            Picture         =   "frmTroco.frx":0CCA
            Style           =   1  'Graphical
            TabIndex        =   143
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox txtbxBanco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   142
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox txtbxContaCorrente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5760
            MaxLength       =   12
            TabIndex        =   141
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtbxNum_Che 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   140
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtbxNome_Che 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   139
            Top             =   360
            Width           =   4815
         End
         Begin VB.TextBox txtbxAgencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4680
            MaxLength       =   8
            TabIndex        =   138
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtbxCPFChe 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   137
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&onta C.:"
            Height          =   195
            Left            =   5760
            TabIndex        =   149
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   148
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Titular:"
            Height          =   195
            Left            =   2040
            TabIndex        =   147
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Cheque Nº:"
            Height          =   195
            Left            =   3360
            TabIndex        =   146
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Agência:"
            Height          =   195
            Left            =   4680
            TabIndex        =   145
            Top             =   720
            Width           =   630
         End
         Begin VB.Label label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&CPF:"
            Height          =   195
            Left            =   120
            TabIndex        =   144
            Top             =   120
            Width           =   345
         End
      End
      Begin VB.CheckBox chkCh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cheque:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   135
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox pcTipoNota 
      Appearance      =   0  'Flat
      BackColor       =   &H0095FAFF&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6960
      ScaleHeight     =   705
      ScaleWidth      =   3225
      TabIndex        =   131
      Top             =   8640
      Width           =   3255
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095FAFF&
         Caption         =   "NFe (55)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   133
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optnCupom 
         Appearance      =   0  'Flat
         BackColor       =   &H0095FAFF&
         Caption         =   "NFCe (65)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de NF"
         Height          =   195
         Left            =   0
         TabIndex        =   150
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox pcDesconto 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      Picture         =   "frmTroco.frx":0DC3
      ScaleHeight     =   1815
      ScaleWidth      =   3120
      TabIndex        =   120
      Top             =   7080
      Width           =   3120
      Begin VB.TextBox txtbxDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   123
         Text            =   "0,00"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtbxDescPerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   122
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmndbtnDesconto 
         Caption         =   "Confirmar"
         Height          =   525
         Left            =   240
         Picture         =   "frmTroco.frx":134F5
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   960
         Width           =   1575
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
         Left            =   240
         TabIndex        =   128
         Top             =   120
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         Height          =   195
         Left            =   240
         TabIndex        =   127
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   1200
         TabIndex        =   126
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Final Previsto:"
         Height          =   195
         Left            =   2040
         TabIndex        =   125
         Top             =   600
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2160
         TabIndex        =   124
         Top             =   960
         Width           =   630
      End
   End
   Begin VB.CheckBox chkVale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vale para funcionário"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   115
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox pcFormasPag 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5520
      Picture         =   "frmTroco.frx":13C5F
      ScaleHeight     =   3015
      ScaleWidth      =   7740
      TabIndex        =   78
      Top             =   1440
      Width           =   7740
      Begin VB.PictureBox pcEntrada 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   0
         Picture         =   "frmTroco.frx":2BF59
         ScaleHeight     =   1785
         ScaleWidth      =   7545
         TabIndex        =   93
         Top             =   1200
         Width           =   7545
         Begin VB.OptionButton optnDep 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEDFD5&
            Caption         =   "Dep"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   29
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton optnPix 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEDFD5&
            Caption         =   "Pix"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.TextBox txtContaPix 
            Appearance      =   0  'Flat
            BackColor       =   &H0073E1F4&
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
            Left            =   1560
            TabIndex        =   32
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtPix 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0073E1F4&
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
            Left            =   240
            TabIndex        =   30
            Tag             =   "PIX"
            Text            =   "0,00"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtDataBaseVale 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8F4F1&
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
            Left            =   5760
            TabIndex        =   40
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtParcelas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F8F4F1&
            Height          =   285
            Left            =   6960
            TabIndex        =   42
            Text            =   "1"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtVale 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F8F4F1&
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
            Left            =   4320
            TabIndex        =   37
            Text            =   "0,00"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtNumeroCartaoPresente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   360
            Width           =   1095
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
            Left            =   2880
            TabIndex        =   26
            Text            =   "0,00"
            Top             =   360
            Width           =   1095
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
            Left            =   240
            TabIndex        =   21
            Text            =   "0,00"
            Top             =   360
            Width           =   1095
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
            Left            =   1440
            TabIndex        =   23
            Text            =   "0,00"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmndbtnDadosCard 
            Height          =   285
            Left            =   2520
            Picture         =   "frmTroco.frx":57E73
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnDadosChv 
            Height          =   285
            Left            =   3960
            Picture         =   "frmTroco.frx":581FD
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdCartaoPresente 
            Height          =   285
            Left            =   5400
            Picture         =   "frmTroco.frx":58587
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdVale 
            Height          =   285
            Left            =   5400
            Picture         =   "frmTroco.frx":58911
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblContaPix 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conta/Chave:"
            Height          =   195
            Left            =   1560
            TabIndex        =   31
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label lblDataBaseVale 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Data Base:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5760
            TabIndex        =   39
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Parc."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6960
            TabIndex        =   41
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblFunVale 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionário"
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
            Height          =   300
            Left            =   4320
            TabIndex        =   116
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vale:"
            Height          =   195
            Left            =   4320
            TabIndex        =   36
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblCartaoPresente 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
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
            Height          =   300
            Left            =   5880
            TabIndex        =   114
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cartão Presente"
            Height          =   195
            Left            =   4320
            TabIndex        =   33
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C&heque:"
            Height          =   195
            Left            =   2880
            TabIndex        =   25
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lblEntrada 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dinheiro:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   120
            Width           =   630
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cartão:"
            Height          =   195
            Left            =   1440
            TabIndex        =   22
            Top             =   120
            Width           =   510
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
         Left            =   5040
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
         Left            =   6120
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtbxInterv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   15
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtbxParcelas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxFormaPag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxDescForma 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   285
         Left            =   2400
         Picture         =   "frmTroco.frx":58C9B
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Prazo:"
         Height          =   195
         Left            =   5040
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
         Left            =   240
         TabIndex        =   94
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label lblDataBase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base:"
         Height          =   195
         Left            =   6120
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
         Left            =   4560
         TabIndex        =   80
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblInternv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo."
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   120
         Width           =   660
      End
      Begin VB.Label lblParcelas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas:"
         Height          =   195
         Left            =   3240
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
         Left            =   2160
         TabIndex        =   79
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pagamento:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   120
         Width           =   120
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   5265
      TabIndex        =   103
      Top             =   4440
      Width           =   5295
      Begin VB.PictureBox pcDescAutorizadoPor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -120
         ScaleHeight     =   735
         ScaleWidth      =   5415
         TabIndex        =   108
         Top             =   1800
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton cmndbtnCancelarDescGer 
            Caption         =   "Cancelar"
            Height          =   255
            Left            =   3840
            TabIndex        =   109
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Autorizado Por:"
            Height          =   195
            Left            =   240
            TabIndex        =   111
            Top             =   120
            Width           =   1815
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
            Left            =   240
            TabIndex        =   110
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   0
         Picture         =   "frmTroco.frx":58D94
         ScaleHeight     =   1155
         ScaleWidth      =   2490
         TabIndex        =   106
         Top             =   480
         Width           =   2490
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
            Left            =   960
            TabIndex        =   107
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
         Height          =   1155
         Left            =   2640
         Picture         =   "frmTroco.frx":6243A
         ScaleHeight     =   1155
         ScaleWidth      =   2490
         TabIndex        =   104
         Top             =   480
         Width           =   2490
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
            Left            =   960
            TabIndex        =   105
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
         Left            =   2760
         TabIndex        =   113
         Top             =   120
         Width           =   1995
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Queima:"
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
         Left            =   360
         TabIndex        =   112
         Top             =   120
         Width           =   1920
      End
   End
   Begin VB.CheckBox chkPromo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cliente Promocional"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   102
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtbxTipoCob 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   53
      Tag             =   "1"
      Text            =   "Pagamento na Loja"
      Top             =   8880
      Width           =   2895
   End
   Begin VB.PictureBox pcAprazo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5520
      ScaleHeight     =   3705
      ScaleWidth      =   8025
      TabIndex        =   77
      Top             =   4800
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtbxMsflx 
         Height          =   285
         Left            =   1200
         TabIndex        =   99
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnEdi 
         Height          =   375
         Left            =   6600
         Picture         =   "frmTroco.frx":6BAE0
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Editar o cheque ou  parcela que você selecionou na grade abaixo..."
         Top             =   120
         Width           =   375
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
         TabIndex        =   46
         Top             =   360
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
         TabIndex        =   48
         Text            =   "0,00"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnAdd 
         Height          =   375
         Left            =   6120
         Picture         =   "frmTroco.frx":6C4E2
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Adiciona o Cheque ou parcela que você digitou.."
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnExcl 
         Height          =   375
         Left            =   7080
         Picture         =   "frmTroco.frx":6CA6C
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Exclui o cheque selecionado na grade abaixo"
         Top             =   120
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid msflxgrd 
         Height          =   2775
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4895
         _Version        =   393216
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
         TabIndex        =   45
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Height          =   195
         Left            =   1800
         TabIndex        =   47
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.CommandButton cmndbtnOk 
      Caption         =   "C&onfirmar (F7)"
      Height          =   735
      Left            =   10320
      Picture         =   "frmTroco.frx":6EE33
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   8640
      Width           =   1575
   End
   Begin VB.OptionButton optnAgora 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lançar Dados Agora"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5640
      TabIndex        =   43
      Top             =   4560
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optnDepois 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lançar Dados Mais Tarde"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7680
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture10 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   13710
      TabIndex        =   81
      Top             =   0
      Width           =   13710
      Begin VB.PictureBox pcLimite 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   7680
         Picture         =   "frmTroco.frx":6EF7D
         ScaleHeight     =   810
         ScaleWidth      =   5295
         TabIndex        =   84
         Top             =   0
         Width           =   5295
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   120
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
            TabIndex        =   91
            Top             =   360
            Width           =   945
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1440
            TabIndex        =   90
            Top             =   120
            Width           =   405
         End
         Begin VB.Label lblLimite 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   89
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   2520
            TabIndex        =   88
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lblEmUso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   2520
            TabIndex        =   87
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   3720
            TabIndex        =   86
            Top             =   120
            Width           =   690
         End
         Begin VB.Label lblLimiteRest 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   3720
            TabIndex        =   85
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   2160
         TabIndex        =   98
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente:"
         Height          =   195
         Left            =   240
         TabIndex        =   97
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
         TabIndex        =   83
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
         Left            =   2160
         TabIndex        =   82
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   5520
      Picture         =   "frmTroco.frx":7CF57
      ScaleHeight     =   810
      ScaleWidth      =   7740
      TabIndex        =   76
      Top             =   720
      Width           =   7740
      Begin VB.TextBox txtbxCNPJ 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   285
         Left            =   4440
         Picture         =   "frmTroco.frx":90C01
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   5040
         TabIndex        =   5
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblCliCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.PictureBox pc 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   -120
      ScaleHeight     =   15
      ScaleWidth      =   14775
      TabIndex        =   73
      Top             =   10200
      Width           =   14775
   End
   Begin VB.CommandButton cmndbtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Ca&ncelar (F12)"
      Height          =   735
      Left            =   12000
      Picture         =   "frmTroco.frx":90CFA
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CheckBox chkEmite 
      Caption         =   "Imprimir Recibo"
      Height          =   615
      Left            =   120
      Picture         =   "frmTroco.frx":91248
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   9000
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1545
      ScaleWidth      =   5265
      TabIndex        =   62
      Top             =   840
      Width           =   5295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   9975
         TabIndex        =   68
         Top             =   1200
         Width           =   9975
         Begin VB.Label lblPExtenso 
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
            Left            =   120
            TabIndex        =   69
            Top             =   0
            Width           =   150
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   9975
         TabIndex        =   63
         Top             =   0
         Width           =   9975
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega:"
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
            Left            =   3600
            TabIndex        =   130
            Top             =   0
            Width           =   795
         End
         Begin VB.Label lblEntrega 
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
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4440
            TabIndex        =   129
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
            Left            =   1800
            TabIndex        =   67
            Top             =   0
            Width           =   930
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
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   2880
            TabIndex        =   66
            Tag             =   "!"
            Top             =   0
            Width           =   105
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
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1200
            TabIndex        =   65
            Tag             =   "!"
            Top             =   0
            Width           =   105
         End
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
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   1110
         End
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   3000
         TabIndex        =   71
         Tag             =   "!"
         Top             =   600
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
         Left            =   240
         TabIndex        =   70
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcDin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   5265
      TabIndex        =   61
      Top             =   2520
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtbxDinheiro 
         Alignment       =   1  'Right Justify
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
         Height          =   540
         Left            =   2040
         TabIndex        =   1
         Text            =   "0,00"
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblDinheiro 
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
         Top             =   0
         Width           =   1350
      End
   End
   Begin VB.PictureBox pcTroco 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   5265
      TabIndex        =   58
      Top             =   3360
      Width           =   5295
      Begin VB.Label lblTroco 
         Alignment       =   1  'Right Justify
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
         Left            =   2955
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.CommandButton cmndbtnCob 
      Height          =   285
      Left            =   6480
      Picture         =   "frmTroco.frx":913AC
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8880
      Width           =   255
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10680
      ScaleHeight     =   375
      ScaleWidth      =   2415
      TabIndex        =   117
      Top             =   4440
      Width           =   2415
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total A Vista:"
         Height          =   195
         Left            =   0
         TabIndex        =   119
         Top             =   120
         Width           =   945
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   1080
         TabIndex        =   118
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Restante:"
      Height          =   195
      Left            =   4560
      TabIndex        =   101
      Top             =   8040
      Width           =   690
   End
   Begin VB.Label lblRestante 
      Alignment       =   1  'Right Justify
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
      Left            =   4560
      TabIndex        =   100
      Top             =   8280
      Width           =   630
   End
   Begin VB.Label lblTipoCob 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cobrança:"
      Height          =   195
      Left            =   3600
      TabIndex        =   52
      Top             =   8640
      Width           =   1125
   End
   Begin VB.Label lblTotalDistribuido 
      Alignment       =   1  'Right Justify
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
      Left            =   4560
      TabIndex        =   96
      Top             =   7560
      Width           =   630
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Distribuído:"
      Height          =   195
      Left            =   3960
      TabIndex        =   95
      Top             =   7320
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
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   74
      Top             =   120
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
Dim cp As CartaoPresente
Private Sub cmndbtnAv_Click()

End Sub


Private Sub chkVale_Click()
    If chkVale.Value = 1 Then
        pcFormasPag.Enabled = False
        txtbxFormaPag.tag = "VAL"
        txtbxFormaPag.text = "VALE"
        lblDinheiro.caption = "Vale"
    Else
        pcFormasPag.Enabled = True
        txtbxFormaPag.tag = "DIN"
        txtbxFormaPag.text = "Dinheiro à vista"
        lblDinheiro.caption = "&Dinheiro"
    End If
End Sub

Private Sub cmdCartaoPresente_Click()
    Busca_CartPresente
End Sub
Sub Busca_FuncVale()
    Dim f As funcionario.funcionario
    If CCur(txtVale.text) <= "0" Then
        lblFunVale.tag = ""
        lblFunVale.caption = "..."
        Exit Sub
    End If
    f = funcionario.F_SelectFuncionario("")
    If f.cod_fun <> "" Then
        lblFunVale.caption = f.guerra_fun
        lblFunVale.tag = f.cod_fun
        txtVale.tag = f.cod_fun
    End If
End Sub
    
Sub Busca_CartPresente()
    cp = BuscaCartaoPresente(txtNumeroCartaoPresente.text)
    If cp.numero <> "" Then
        If cp.status <> "1" Then
            MsgBox "O cartão presente é inválido!", vbExclamation
            lblCartaoPresente.caption = "0,00"
            P_TotalEntrada
            Exit Sub
        End If
        
        txtNumeroCartaoPresente.text = cp.numero
        txtNumeroCartaoPresente.tag = cp.id
        lblCartaoPresente.caption = FormatVal(cp.valor)
        P_TotalEntrada
    End If
End Sub

Private Sub cmdVale_Click()
    Busca_FuncVale
    
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
        txtbxMsflx.top = msflxgrd.top + msflxgrd.CellTop
        txtbxMsflx.Left = msflxgrd.Left + msflxgrd.CellLeft
        If KeyAscii <> 13 Then
            txtbxMsflx.text = Chr(KeyAscii)
        End If
        txtbxMsflx.Visible = True
        
        
        
        txtbxMsflx.SetFocus
        txtbxMsflx.SelStart = Len(txtbxMsflx.text)
        
        
    End If
End Sub


Private Sub optnDep_Click()
    txtPix.tag = "DEP"
End Sub

Private Sub optnPix_Click()
    txtPix.tag = "PIX"
End Sub

Private Sub txtbxAgencia_GotFocus()
    P_SelecTexto txtbxAgencia
End Sub

Private Sub txtbxAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxAgencia_LostFocus()
    P_AjustaTexto txtbxAgencia
End Sub

Private Sub txtbxBanco_GotFocus()
    P_SelecTexto txtbxBanco
End Sub

Private Sub txtbxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaBanco_Click
    End If
End Sub

Private Sub txtbxBanco_LostFocus()
    P_AjustaTexto txtbxBanco
End Sub

Private Sub txtbxContaCorrente_GotFocus()
    P_SelecTexto txtbxContaCorrente
End Sub

Private Sub txtbxContaCorrente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxContaCorrente_LostFocus()
    P_AjustaTexto txtbxContaCorrente
End Sub

Private Sub txtbxDataBase_GotFocus()
    P_SelecTexto txtbxDataBase
End Sub

Private Sub txtbxDataBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDataBase_LostFocus()
    P_AjustaData txtbxDataBase
    distribuiparcelas
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
    
    sDes = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & sUsu & "'")
    If sDes = "" Then
        sDes = "0"
    End If
    If CCur(sDes) < CCur(txtbxDescPerc.text) Then
        MsgBox "O Desconto possível para este usuário é de " & sDes & "%", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Confirma desconto de " & txtbxDesconto.text & "?", vbYesNo + vbCritical) = vbYes Then
        Vendas.P_RecalculaVendaDesconto lblFinalPrev.caption, txtbxDescPerc.text, Me.tag, sUsu
        carrega
    End If
    
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    Dim sVisivel As String
    Dim sSEQ_DOC As String
    
    'On Error GoTo trata_erro
    Dim rest As Currency
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
    Dim erroLancamentoCaixa As Boolean
    Dim msgErro As String
    
    
    If CDate(txtbxDataBase.text) < DateAdd("M", -1, Date) Then
        MsgBox "A data base deve ser igual ou posterior a um mês atrás!", vbExclamation
        Exit Sub
    End If
    If CCur(lblTroco.caption) < 0 Then
        MsgBox "Informe o Valor Pago!", vbExclamation
        'txtbxDinheiro.Text = lblValor.caption
        'txtbxDinheiro.SetFocus
        Exit Sub
       
    End If
    If CCur(txtbxEntradaCart.text < 0) Or CCur(txtbxEntradaChv.text < 0) Or CCur(txtbxEntradaDin.text < 0) Or CCur(txtVale.text) < 0 Then
        MsgBox "Valores inválidos!", vbExclamation
        Exit Sub
    End If
    If CCur(lblRestante.caption) < 0 Then
        MsgBox "Valores lançados inválidos!", vbExclamation
        Exit Sub
    End If
    
    cLancado = CCur(txtPix.text) + CCur(txtbxDinheiro.text) + CCur(txtbxEntradaDin.text) + CCur(txtbxEntradaCart.text) + CCur(txtbxEntradaChv.text) + CCur(lblCartaoPresente.caption) + CCur(txtVale.text)
    If cLancado < CCur(lblValor.caption) And lblTipoCai.caption = "DIN" Then
        MsgBox "O valor não está lançado corretamente!", vbExclamation
        Exit Sub
    End If
    If CCur(txtVale.text) > "0" And lblFunVale.tag = "" Then
        MsgBox "Informe o funcionário responsável pelo vale!", vbExclamation
        Exit Sub
    End If
    If optnAgora.Value Then
        If Not F_confMsflxCH(msflxgrd, 2, 4, 5, 6, 7, 8, 9, 1) Then
            Exit Sub
        End If
    
        
    End If
    
    If CCur(lblRestante.caption) > 0 Then
        MsgBox "Está restando " & lblRestante.caption & " para lançar!", vbExclamation
        Exit Sub
    ElseIf CCur(lblRestante.caption) < 0 Then
        MsgBox "O valor lançado é maior que o precço (" & lblRestante.caption & " )", vbExclamation
        Exit Sub
    End If
    If pc.BackColor = vbGreen Then
        sVisivel = "S"
    Else
        sVisivel = "N"
    End If
        
    If lblTipoCai.caption = "DUP" Or lblTipoCai.caption = "CHP" Or lblTipoCai.caption = "CHV" Then
        
        
        
        If CCur(lblTotalDistribuido.caption) <> CCur(txtbxValorDoc) Then
            MsgBox "O Valor a prazo está diferente do distribuído!", vbExclamation
            Exit Sub
        End If
        If txtbxNomeCli.tag = "" Or txtbxNomeCli.tag = "00000" Then
            MsgBox "Esta opção de pagamento requer que o cliente tenha sido cadastrado!", vbExclamation
            txtbxNomeCli.SetFocus
            Exit Sub
        End If
        If lblStatus_CLi.tag = "2" Then
            MsgBox "Não é possível a realização da venda a crédito pois o cliente se encontra BLOQUEADO!", vbExclamation
            Exit Sub
        End If
        If lblStatus_CLi.tag = "3" Then
            MsgBox "Não é possível a realização da venda a crédito pois o cliente só é liberado para compra com Cheques", vbExclamation
            Exit Sub
        End If
        If txtbxNomeCli.tag = "0" Or txtbxNomeCli.tag = "000000" Or txtbxNomeCli.tag = "00000" Or txtbxNomeCli.tag = "" Then
            MsgBox "Não é possível realizar vendas a prazo para cliente não informado. (Venda a consumidor)", vbExclamation
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
            pc.BackColor = vbGreen
                If Not TefNaoFiscalSemECF(txtbxEntradaCart.text, "", "", "", "", pagcartao.parcelas) Then
                    
                    MsgBox "Erro no TEF! Use outra forma de pagamento ou acione o TEF corretamente", vbExclamation
                    P_Status
                    Exit Sub
                Else
                    strSQL = F_BuscaUmCampo("cod_fin", "tab_financeiras", "sigla_fin='" & TransacaoTEFAtual.rede & "'")
                    If strSQL <> "" Then
                        pagcartao.cod_fin = strSQL
                        pagcartao.parcelas = TransacaoTEFAtual.parcelas
                        pagcartao.CNPJ = F_BuscaUmCampo("cnpj", "tab_financeiras", "cod_fin='" & pagcartao.cod_fin & "'")
                        
                    End If
                    
                    
                End If
            
        Else
            CodigaoAutorizacaoTEF = pagcartao.autorizacao
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
            erroLancamentoCaixa = True
            msgErro = msgErro & "Erro na adição de forma de pagamento usada no fechamento VENFEC" & Chr(13)
            
        End If
        
        
        
        If F_GravaVenFin(Me.tag, pagcartao.cod_fin, CCur(txtbxEntradaCart.text), Date, CodigaoAutorizacaoTEF, _
            pagcartao.parcelas, "V", pagcartao.CNPJ, TransacaoTEFAtual.via) Then
        End If
        
        
        strSQL = F_GravaCai("C", Date, txtbxEntradaCart.text, "PAG. CARTÃO " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, sEspCard, "", txtbxEntradaCart.text, 0, strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
        
        
        If strSQL = "" Then
            erroLancamentoCaixa = True
            msgErro = "Erro no lançamento do caixa pag cartão" & Chr(13)
        End If
            
    End If
            
    If CCur(txtbxEntradaDin.text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtbxEntradaDin.text, "PAG. A VISTA VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, "DIN", "", txtbxEntradaDin.text, 0, strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
        
        If strSQL = "" Then
            erroLancamentoCaixa = True
            msgErro = msgErro & "Houve um erro no lançamento do caixa. Avise o responsável!" & Chr(13)
            
        End If
        
        If Not F_AdicionaVenFec(Me.tag, txtbxEntradaDin.text, "01", _
            strCod_Cai, strCod_Tur, "DIN") Then
            erroLancamentoCaixa = True
            msgErro = msgErro & "Houve um erro no lançamento do caixa. VENFEC!" & Chr(13)
            
        End If
        
        
    End If
    If CCur(txtPix.text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtPix.text, IIf(optnPix.Value, "PG PIX ", "PG DEP ") & txtContaPix.text & " VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, txtPix.tag, "", txtPix.text, txtPix.text, strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
        
        If strSQL = "" Then
            erroLancamentoCaixa = True
            msgErro = msgErro & "Houve um erro no lançamento do caixa. Erro no PIX!" & Chr(13)
            
        End If
        
        If Not F_AdicionaVenFec(Me.tag, txtPix.text, "23", _
            strCod_Cai, strCod_Tur, "PIX") Then
            erroLancamentoCaixa = True
            msgErro = msgErro & "Houve um erro no lançamento do caixa. Erro no PIX!" & Chr(13)
            
        End If
        
         
    End If
            
    If CCur(txtbxEntradaChv.text) > 0 Then
        strSQL = F_GravaCai("C", Date, txtbxEntradaChv.text, "PAG. CHV VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, "CHV", "", txtbxEntradaChv.text, 0, strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
        
        If strSQL = "" Then
            erroLancamentoCaixa = True
            msgErro = msgErro & "Houve um erro no lançamento do caixa. Erro CHV!" & Chr(13)
            
        End If
        
        If Not F_AdicionaVenFec(Me.tag, txtbxEntradaChv.text, "02", _
            strCod_Cai, strCod_Tur, "CHV") Then
            erroLancamentoCaixa = True
            msgErro = msgErro & "Houve um erro no lançamento do caixa. Erro VENFEC!" & Chr(13)
            
        End If
        sSEQ_DOC = F_GravaDoc("VENDA CHV " & txtbxNomeCli.text, Me.tag, txtbxEntradaChv.text, _
            Date, txtbxNomeCli.tag, "R", Me.tag, 1, "CHV", sTipo, "000")
        strSQL = F_GravaCh(sSEQ_DOC, che.num_che, che.cod_bco, che.age_che, che.cco_che, che.nome_che, CStr(che.Valor_che), _
            che.Venc_che, txtbxNomeCli.tag, "R", Me.tag, "", "1", "", che.cpf_che)
            
        
    End If
    If CCur(lblCartaoPresente.caption) > 0 Then
        strSQL = F_GravaCai("C", Date, lblCartaoPresente.caption, "PAG. CART.PRES. " & txtNumeroCartaoPresente.text _
            & " VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
            Me.tag, "CRP", "", CCur(lblCartaoPresente), 0, strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
        If Not F_AdicionaVenFec(Me.tag, lblCartaoPresente.caption, "23", _
            strCod_Cai, strCod_Tur, "CRP") Then
            erroLancamentoCaixa = True
            msgErro = msgErro & "Houve um erro no lançamento do caixa. Erro no CH!" & Chr(13)
        End If
        Cartao_Presente.BaixaCartaoPresente cp.id, lblCartaoPresente.caption, Me.tag
    End If
    If CCur(txtVale.text) > 0 Then
        
            
            If vales.LancaVale(lblFunVale.tag, txtVale.text, txtParcelas.text, txtDataBaseVale.text, "VALE  " & _
                lblFunVale.caption & "-VENDA " & Me.tag & "-" & txtbxNomeCli.text, Me.tag) = False Then
                MsgBox "Erro no lançamento do vale para o funcionário!", vbExclamation
                
            End If
            strSQL = F_GravaCai("C", Date, txtVale.text, "VALE  " & _
                lblFunVale.caption & "-VENDA " & Me.tag & "-" & txtbxNomeCli.text, Me.tag, _
                "VAL", "", CCur(txtVale.text), CCur(txtVale.text), strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
        
            If Not F_AdicionaVenFec(Me.tag, txtVale.text, "01", _
              strCod_Cai, strCod_Tur, "DIN") Then
                erroLancamentoCaixa = True
                msgErro = msgErro & "Houve um erro no lançamento do caixa. Erro no VENFEC!" & Chr(13)
            End If
        
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
                                msgErro = " Erro na gravação da parcela " & iCount & Chr(13)
                                erroLancamentoCaixa = True
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
                                msgErro = " Erro na gravação da parcela " & iCount
                                erroLancamentoCaixa = True
            
                            End If
                        End If
                            
                    Next
                End With
                
                
                If cTotalChp > 0 Then
                    strSQL = F_GravaCai("C", Date, CStr(cTotalChp), "PAG. A CHP VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
                    Me.tag, "CHP", "", CStr(cTotalChp), 0, strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
                    If strSQL = "" Then
                        erroLancamentoCaixa = True
                        msgErro = msgErro & "Houve um erro no lançamento do caixa. Erro no CHEQ!" & Chr(13)
            
                    End If
                End If
                If cTotalDup > 0 Then
                    strSQL = F_GravaCai("C", Date, CStr(cTotalDup), "PAG. A CHP VENDA: " & Me.tag & " " & txtbxNomeCli.text, _
                    Me.tag, "DUP", "", CStr(cTotalDup), 0, strCod_Cai, strCod_Tur, "000", sVisivel, txtbxDesconto.text)
                    If strSQL = "" Then
                        erroLancamentoCaixa = True
                        msgErro = msgErro & "Erro no lançamento do Caixa (DUP)" & Chr(13)
            
                    End If
                End If
            End If
            
        Else
            
        End If
        
        If Not F_AdicionaVenFec(Me.tag, txtbxValorDoc.text, txtbxFormaPag.tag, _
                strCod_Cai, strCod_Tur, lblTipoCai.caption) Then
                
        End If
                
                
    End If
                
                
            
            
            
    strSQL = "update tab_vendas set nomecli_ven = '" & txtbxNomeCli.text & "' "
    strSQL = strSQL & ",valor_ven=" & F_GravaDec(CCur(lblValor.caption) - CCur(lblEntrega.caption))
        
    If optnAgora.Value Then
        strSQL = strSQL & ",situacao_ven='4', " 'efetivada
            
    Else
        strSQL = strSQL & ",situacao_ven='3', " 'a efetivar
    End If
    strSQL = strSQL & " dataefet_ven=" & F_GravaData(Date) & ", "
    strSQL = strSQL & " fechavenda_usu='" & strCod_Usu & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "' "
       
    strSQL = strSQL & " where num_ven='" & Me.tag & "' "
    Conn.Execute strSQL
    
    P_BaixaVenda Me.tag
        
    
    
    

    
    P_DeletaVendaAberta Me.tag
    
        
    strSQL = pacote.F_GravaPacote(Me.tag)
    If sSEQ_DOC <> "" And sTipo <> "U" Then
        imprimeContrato sSEQ_DOC
     End If
    If sSEQ_DOC <> "" Then
        imprimeCarne sSEQ_DOC
    End If
    If chkEmite.Value = 1 Then
        ImprVenda Me.tag
    End If
     
    P_Status " Gerando NFe"
    If chkPromo.Value = 1 And sSEQ_DOC <> "" Then
    
        
     Else
        
        
        If optnCupom.Value Then
            Dim ncm As String
            ncm = VerificaNCMVenda(Me.tag)
            If ncm = "" Then
                strSQL = GravaNFCeVenda(Me.tag, strTipo_IcmsNFe)
             
                 If strSQL = "" Then
                     MsgBox "Houve erro na geração da NFCe da Venda!", vbExclamation
                     
                     Exit Sub
                 End If
                 P_Status ""
                 If MsgBox("NFC-E gerada. Deseja transmitir", vbYesNo + vbQuestion) = vbYes Then
                     strSQL = TransmitirNFcE(strSQL, "")
                     If InStr(1, "CStat=100", strSQL) > 0 Then
                         AlteraStatusNfe Me.tag, "2"
                     Else
            
                     End If
                 End If
            Else
                
            End If
        Else
             strSQL = NFe_DAO.GravaNFCeVenda(Me.tag, strTipo_IcmsNFe, "Venda de Produtos", modelo_Nfe.NFe_55)
             
             If strSQL = "" Then
                 MsgBox "Houve erro na geração da NFCe da Venda!", vbExclamation
                Else
                    frmNFE_Detalhe.tag = strSQL
                    frmNFE_Detalhe.Show 1
                
             End If
        End If
    End If
                
            
        
     If sSEQ_DOC <> "" Then
        P_Status "enviando para internet"
        APIInserirDoc sSEQ_DOC, True
        P_Status ""
     End If
    
    If ncm <> "" Then
        MsgBox "Não é possível gerar a nota, há produtos sem NCM ou com ncm inválido." & Chr(13) & ncm, vbExclamation
    End If
    pagcartao.tipo_fin = ""
    
    P_Status ""
    If erroLancamentoCaixa = True Then
        MsgBox msgErro, vbExclamation
    End If
    
    
    
    Me.tag = "X"
    Me.Hide
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
        
    If bActive = True Then Exit Sub
    che.Valor_che = 0
    che.num_che = ""
    che.nome_che = ""
    che.Venc_che = ""
    che.num_che = ""
    carrega
    bActive = True
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
        lblTotalDistribuido.caption = FormatVal(cTotal)
        .row = iRow
        .col = iCol
    End With
    
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
       ' txtbxEntradaDin.SetFocus
        
        
    End If
    If lblTipoCai.caption = "CHP" Or lblTipoCai.caption = "CHV" Then
        
    End If
    pcEntrada.Visible = True
    P_TotalEntrada
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro
    
    strSQL = F_VerificaVendaAberta(Me.tag)
    If strSQL <> "" And strSQL <> strIP_Usu Then
        
        MsgBox "A Venda se encontra aberta no computador " & strSQL, vbExclamation
        Unload Me
    End If
    
    
    P_GravaVendaAberta Me.tag
    
    
    sTipo = F_BuscaUmCampo("tipocli_ven", "tab_vendas", "num_ven='" & Me.tag & "'")
    
    If sTipo = "U" Then
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "nome_fun"
        sCPF = "cpf_fun"
        sLimite = "limite_fun"
        sStatus = "status_fun"
    Else
        sTab = TAB_CLIENTES
        sCod = "cod_cli"
        sNome = "nome_cli"
        sCPF = "cpf_cli"
        sLimite = "limite_cli"
        sStatus = "status_cli"
        sTipo = "C" 'cliente
    End If
    
    strSQL = "select tc." & sLimite & " as limite_cli ,tc." & sCPF & " as cpf_cli,"
    strSQL = strSQL & " tt.nome_tip,tc." & sStatus & " as status_cli,tv.*,tf.*,tfu.guerra_fun from tab_vendas as tv, "
    strSQL = strSQL & sTab & " as tc, tab_formaspag as tf,tab_funcionarios as tfu,tab_tipopag as tt "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for  and tc." & sCod & "=tv.cod_cli "
    strSQL = strSQL & " and tt.cod_tip=tf.cod_tip "
    strSQL = strSQL & " and num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tv.abrevenda_usu=tfu.cod_fun "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        txtbxNomeCli.text = !nomecli_ven
        txtbxNomeCli.tag = !cod_cli
        
        txtbxCNPJ.text = Format(!cpfcli_ven)
        
        txtbxNomeCli.text = !nomecli_ven
        txtbxFormaPag.text = !nome_for
        lblTipoCai.caption = !tipo_cai
        
        txtbxDescForma.text = !desconto_for 'FormatVal(!desconto_ven / !valor_ven) * 100 '!desconto_for
        txtbxFormaPag.tag = !cod_For
        lblValor.caption = FormatVal(!valor_ven + !valorentrega_ven)
        lblValor.tag = FormatVal(!valor_ven - !desconto_ven)
        lblEntrega.caption = FormatVal(!valorentrega_ven)
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
            lblAutorizaDesconto.caption = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & !autorizadesconto_ven & "'")
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
            lblStatus_CLi.tag = !status_cli
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
Function Restante() As Currency
    Restante = CCur(lblValor.caption) - CCur(lblTotalDistribuido.caption)
End Function
Sub P_TotalEntrada()
    'On Error Resume Next
    lblTotalEntrada.caption = FormatVal(CCur(txtbxEntradaCart.text) + CCur(txtbxEntradaDin.text) + CCur(txtbxEntradaChv.text) + CCur(lblCartaoPresente.caption) + CCur(txtVale.text) + CCur(txtPix.text))
    txtbxValorDoc.text = FormatVal(CCur(lblValor.caption) - CCur(lblTotalEntrada.caption))
    lblRestante.caption = FormatVal(CCur(lblValor.caption) - CCur(lblTotalDistribuido.caption) - CCur(lblTotalEntrada.caption))
    
End Sub
Sub distribuiparcelas()
    msflxgrd.rows = 1
    Dim strSQL As String
    Dim iCount As Integer
    Dim cTotal As Currency
    Dim sTipo As String
    Dim cValor As Currency
    Dim sVlrParcela As String
    
    If lblTipoCai.tag = "01" Then
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
        f = F_SelectFormasPag(txtbxFormaPag.tag, "", txtbxFormaPag.tag)
    Else
        f = F_SelectFormasPag(txtbxFormaPag.text)
    End If
    
    If f.cod_For <> "" Then
        If CCur(txtbxDescForma.text) > f.desconto_for Then
            
                strSQL = "Você optou por mudar a forma de pagamento de " & txtbxFormaPag.text
                strSQL = strSQL & " para " & f.nome_for & Chr(13)
                strSQL = strSQL & "Continuando com esta operação, todo o desconto da venda será cancelado!"
                strSQL = strSQL & "Deseja continuar?"
                If MsgBox(strSQL, vbOKCancel + vbCritical, "CANCELAMENTO DE DESCONTO") = vbCancel Then
                    Exit Sub
                Else
                    P_CancelaDescontoVenda Me.tag, CStr(f.desconto_for)
                    carrega
                End If
            
        End If
        
        
        'txtbxEntradaCart.Text = "0,00"
        'txtbxEntradaChv.Text = "0,00"
        'txtbxEntradaDin.Text = "0,00"
        
        txtbxFormaPag.text = f.nome_for
        txtbxFormaPag.tag = f.cod_For
        txtbxDescForma.text = f.desconto_for
        lblTipoCai.caption = f.tipo_cai
        lblTipoCai.tag = f.cod_tip
        txtbxParcelas.text = f.parcelas_for
        distribuiparcelas
        Vendas.P_AtualizaFormaPag Me.tag, txtbxFormaPag.tag
        
    
        
         
        If lblTipoCai.caption <> "DIN" Then
            'txtbxDinheiro.Text = lblValor.caption
            'txtbxDinheiro_LostFocus
        End If
        If f.tipo_cai = "CRT" Or f.tipo_cai = "CRD" Then
            txtbxEntradaCart.text = lblValor.caption
            Exit Sub
        End If
        P_EscondeOpcoesPagamento
        If f.cod_For = "01" Then
            'avista
            pcTroco.Visible = True
        Else
            pcTroco.Visible = False
        End If
        
        
    End If
    
    
    
    
End Sub
Private Sub cmndbtnCliente_Click()
    Dim c As Clientes.cliente
    
    c = Clientes.F_BuscaCliente(txtbxNomeCli.text)
    If c.cod_cli <> "" Then
        txtbxNomeCli.text = c.nome_cli
        txtbxNomeCli.tag = c.cod_cli
        txtbxCNPJ.text = c.cnpj_Cli
        
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
            MsgBox "O Cliente encontra-se negativado, portanto, só é possível venda a vista!", vbExclamation
            
            
        End If
        
        Vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, "", txtbxCNPJ.text
        mudaCorCliente
        
    End If
    
    
End Sub
Sub mudaCorCliente()
    If txtbxNomeCli.tag = "" Or txtbxNomeCli.tag = "00000" Then
        txtbxNomeCli.BackColor = vbRed
        lblCliCaption.caption = "&Cliente: (Não Cadastrado)"
    Else
        txtbxNomeCli.BackColor = vbGreen
        lblCliCaption.caption = "&Cliente:"
    End If
    
    
End Sub
Private Sub txtbxCPFChe_GotFocus()
    P_SelecTexto txtbxCPFChe
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
    P_SelecTexto txtbxDesconto
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
    txtbxDesconto.text = FormatVal(CCur(txtbxDescPerc.text) / 100 * CCur(lblDemaisProd.caption))
    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.text))
    
End Sub

Private Sub txtbxDinheiro_Change()
    txtbxEntradaDin.text = txtbxDinheiro.text
End Sub

Private Sub txtbxEntradaCart_GotFocus()
    If txtbxEntradaCart.text = "0,00" Then
        txtbxEntradaCart.text = Restante
    End If
    P_SelecTexto txtbxEntradaCart
End Sub

Private Sub txtbxEntradaCart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntradaCart_LostFocus()
    P_ValorLostFocus txtbxEntradaCart
    P_TotalEntrada
    buscaDadosCartao
End Sub

Sub buscaDadosCartao()
    If CCur(txtbxEntradaCart.text) > 0 Then
        pagcartao = F_DadosCartao(pagcartao.cod_fin, pagcartao.nome_fin, pagcartao.parcelas, _
        pagcartao.autorizacao, txtbxEntradaCart.text)
        'pcDadosCartao.Visible = True
        If pagcartao.cod_fin = "" Then
                
            MsgBox "Não foi escolhida uma financeira!", vbExclamation
        Else
            'cmndbtnFormasPag_Click
        End If
    Else
        pagcartao.cod_fin = ""
        pagcartao.nome_fin = ""
        pagcartao.valor = "0"
        
        'pcDadosCartao.Visible = False
    End If
    
    
End Sub

Private Sub txtbxEntradaChv_GotFocus()
    P_SelecTexto txtbxEntradaChv
    If txtbxEntradaChv.text = "0,00" Then
        txtbxEntradaChv.text = Restante
    End If
End Sub

Private Sub txtbxEntradaChv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntradaChv_LostFocus()
    P_ValorLostFocus txtbxEntradaChv
    P_TotalEntrada
    cmndbtnDadosChv_Click
    If msflxgrd.rows > 1 Then
    If MsgBox("Deseja recalcular as parcelas já lançadas?", vbYesNo + vbQuestion) = vbYes Then
        cmndbtnFormasPag_Click
    End If
    End If
End Sub

Private Sub txtbxEntradaDin_GotFocus()
    
    
    If txtbxEntradaDin.text = "0,00" Then
        txtbxEntradaDin.text = Restante
    End If
    P_SelecTexto txtbxEntradaDin
End Sub

Private Sub txtbxEntradaDin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxEntradaDin_LostFocus()
    P_ValorLostFocus txtbxEntradaDin
    P_TotalEntrada
    
    cmndbtnFormasPag_Click
End Sub

Private Sub txtbxFormaPag_Change()
    txtbxFormaPag.tag = ""
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxMsflx_GotFocus()
'    P_SelecTexto txtbxMsflx
End Sub

Private Sub txtbxMsflx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxMsflx_LostFocus()
    If msflxgrd.col = 2 Then
        P_ValorLostFocus txtbxMsflx
    Else
    End If
    msflxgrd.text = txtbxMsflx.text
    If msflxgrd.col <> msflxgrd.Cols - 1 Then
    
        msflxgrd.col = msflxgrd.col + 1
        msflxgrd.SetFocus
    End If
End Sub

Private Sub txtbxNome_Che_GotFocus()
    P_SelecTexto txtbxNome_Che
End Sub

Private Sub txtbxNome_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNome_Che_LostFocus()
    P_AjustaTexto txtbxNome_Che
End Sub

Private Sub txtbxNum_Che_GotFocus()
    P_SelecTexto txtbxNum_Che
End Sub

Private Sub txtbxNum_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNum_Che_LostFocus()
    P_AjustaTexto txtbxNum_Che
End Sub

Private Sub txtbxValor_Che_GotFocus()
    P_SelecTexto txtbxValor_Che
End Sub

Private Sub txtbxValor_Che_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxValor_Che_LostFocus()
    P_ValorLostFocus txtbxValor_Che
End Sub

Private Sub txtbxVencimento_GotFocus()
    P_SelecTexto txtbxVencimento
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
    If KeyCode = vbKeyF8 Then
        pc.BackColor = vbGreen
        'txtbxDinheiro.Text = lblValor.caption
        cmndbtnOk.SetFocus
        
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
        
    End If
    If KeyCode = vbKeyF7 Then
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
    txtDataBaseVale.text = Date
    txtDataBaseVale.Enabled = True
    bActive = False
    If bUsaCaixaSeparado = True Then
        pc.BackColor = vbGreen
    
        pc.ForeColor = vbGreen
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
    P_SelecTexto txtbxCNPJ
End Sub

Private Sub txtbxCNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, "", F_SoNumero(txtbxCNPJ.text)
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtbxCNPJ_LostFocus()
    If Len(txtbxCNPJ.text) > 14 Then
        txtbxCNPJ.text = Format(txtbxCNPJ.text, "0#\.###\.###\/####\-##")
    Else
        txtbxCNPJ.text = Format(txtbxCNPJ.text, "00#\.###\.###\-##")
    End If
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, "", F_SoNumero(txtbxCNPJ.text)
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
   P_ValorLostFocus txtbxDesconto
   On Error Resume Next
   
   txtbxDescPerc.text = FormatVal(CCur(txtbxDesconto.text) * 100 / CCur(lblDemaisProd.caption), "###,##0.0000")

    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.text))
   
   
End Sub

Private Sub txtbxDinheiro_GotFocus()
    P_SelecTexto txtbxDinheiro
End Sub

Private Sub txtbxDinheiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
        
    End If
End Sub

Private Sub txtbxDinheiro_LostFocus()
    
    P_ValorLostFocus txtbxDinheiro
    If CCur(txtbxDinheiro.text) >= CCur(lblValor.caption) Then
        txtbxEntradaDin.text = lblValor.caption
        txtbxEntradaDin_LostFocus
    End If
    
    lblTroco.caption = FormatVal(CCur(txtbxDinheiro.text) - CCur(lblValor.caption))
    If CCur(lblTroco.caption) < 0 Then
        lblTroco.caption = "0,00"
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

Private Sub txtbxParcelas_GotFocus()
    P_SelecTexto txtbxParcelas
End Sub

Private Sub txtbxParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxParcelas_LostFocus()
    P_AjustaTexto txtbxParcelas
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

Private Sub txtContaPix_GotFocus()
    P_SelecTexto txtContaPix
End Sub

Private Sub txtContaPix_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtContaPix_LostFocus()
    P_AjustaTexto txtContaPix, -1
End Sub

Private Sub txtDataBaseVale_GotFocus()
    P_SelecTexto txtDataBaseVale
End Sub

Private Sub txtDataBaseVale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
    
End Sub

Private Sub txtDataBaseVale_LostFocus()
    If txtDataBaseVale.text = "" Then
        txtDataBaseVale.text = Date
    End If
    P_AjustaData txtDataBaseVale
    
End Sub

Private Sub txtNumeroCartaoPresente_Change()
    txtNumeroCartaoPresente.tag = ""
End Sub

Private Sub txtNumeroCartaoPresente_GotFocus()
    P_SelecTexto txtNumeroCartaoPresente
End Sub

Private Sub txtNumeroCartaoPresente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Busca_CartPresente
    End If
End Sub

Private Sub txtParcelas_LostFocus()
    P_ValorLostFocus txtParcelas, "#"
    If CCur(txtParcelas.text) <= 0 Then
        txtParcelas.text = "1"
    End If
End Sub

Private Sub txtPix_GotFocus()
    P_SelecTexto txtPix
    If CCur(txtPix.text) <= 0 Then
        txtPix.text = Restante
    End If
End Sub

Private Sub txtPix_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
    
End Sub

Private Sub txtPix_LostFocus()
    On Error Resume Next
    P_ValorLostFocus txtPix
    If CCur(txtPix.text) > 0 Then
        txtContaPix.Visible = True
        lblContaPix.Visible = True
        txtContaPix.SetFocus
    Else
        txtContaPix.Visible = False
        lblContaPix.Visible = False
    End If
    P_TotalEntrada
End Sub

Private Sub txtVale_GotFocus()
    P_SelecTexto txtVale
    If txtVale.text = "0,00" Then
        txtVale.text = Restante
    End If
End Sub

Private Sub txtVale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Busca_FuncVale
    End If
End Sub

Private Sub txtVale_LostFocus()
    P_ValorLostFocus txtVale
    P_TotalEntrada
    
    If CCur(txtVale.text) <= 0 Then
        lblDataBaseVale.Enabled = False
        txtDataBaseVale.Enabled = False
    Else
        lblDataBaseVale.Enabled = True
        txtDataBaseVale.Enabled = True
    
    End If
    
End Sub
