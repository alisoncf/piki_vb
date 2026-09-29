VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLanChe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lançamento de Cheque:"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EF9C7C&
      Caption         =   "Dados da Venda:"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9375
      Begin VB.Label lblValorVen 
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
         Left            =   6600
         TabIndex        =   28
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         Height          =   195
         Left            =   6120
         TabIndex        =   27
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda a Consumidor"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   3480
         TabIndex        =   25
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblDataVen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
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
         Left            =   2280
         TabIndex        =   24
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         Height          =   195
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblNumVen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000"
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
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C6DADD&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4695
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   720
      Width           =   9375
      Begin VB.CommandButton cmndbtnExcl 
         Height          =   375
         Left            =   7920
         Picture         =   "frmLanChe.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmndbtnAdd 
         Height          =   375
         Left            =   7440
         Picture         =   "frmLanChe.frx":23C7
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Adiciona o Cheque que você digitou.."
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2415
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
         _Version        =   393216
      End
      Begin VB.TextBox txtbxValor_Che 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         MaxLength       =   14
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmbxSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   7680
         Picture         =   "frmLanChe.frx":2523
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtbxContaCorrente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnConfirmar 
         Caption         =   "Con&firmar"
         Height          =   615
         Left            =   6120
         Picture         =   "frmLanChe.frx":268E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtbxVencimento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxNum_Che 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtbxNome_Che 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtbxAgencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   8
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtbxCPF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Va&lor:"
         Height          =   195
         Left            =   5760
         TabIndex        =   17
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C&onta C.:"
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         Height          =   195
         Left            =   6240
         TabIndex        =   15
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   4440
         TabIndex        =   14
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Titular:"
         Height          =   195
         Left            =   2160
         TabIndex        =   13
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cheque Nº:"
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Agência:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lblCPF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&CPF:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmLanChe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
