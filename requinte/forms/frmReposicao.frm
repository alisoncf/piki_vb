VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmReposicao 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16170
   Icon            =   "frmReposicao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10665
   ScaleWidth      =   16170
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   5385
      TabIndex        =   4
      Top             =   5520
      Width           =   5415
      Begin VB.CommandButton cmndbtnInserir 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   2280
         TabIndex        =   69
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3480
         TabIndex        =   64
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtbxFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Left            =   960
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "0,00"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4680
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor Final:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
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
         TabIndex        =   0
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
         Left            =   4680
         TabIndex        =   2
         Top             =   0
         Width           =   300
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
      ScaleWidth      =   16170
      TabIndex        =   30
      Top             =   10620
      Width           =   16170
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10080
      Left            =   0
      ScaleHeight     =   10050
      ScaleWidth      =   5385
      TabIndex        =   33
      Top             =   540
      Width           =   5415
      Begin VB.PictureBox pcStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1185
         ScaleWidth      =   5145
         TabIndex        =   74
         Top             =   8520
         Width           =   5175
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lançamento"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1155
            TabIndex        =   75
            Top             =   120
            Width           =   2625
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         ScaleHeight     =   1065
         ScaleWidth      =   5385
         TabIndex        =   51
         Top             =   0
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
            TabIndex        =   58
            Top             =   720
            Width           =   5415
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
            TabIndex        =   53
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   55
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
               TabIndex        =   54
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
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   2370
            TabIndex        =   52
            Tag             =   "!"
            Top             =   240
            Width           =   630
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
         TabIndex        =   34
         Top             =   1080
         Width           =   5415
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc %"
            Height          =   195
            Left            =   4560
            TabIndex        =   63
            Top             =   2760
            Width           =   540
         End
         Begin VB.Label lblDescSub 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4560
            TabIndex        =   62
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            Height          =   195
            Left            =   1800
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   960
            Width           =   885
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblCodPro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   47
            Top             =   960
            Width           =   1605
         End
         Begin VB.Label lblSubSec 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   46
            Top             =   3000
            Width           =   4365
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Seção:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   2760
            Width           =   840
         End
         Begin VB.Label lblSec 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   44
            Top             =   2280
            Width           =   5085
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seção:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   2040
            Width           =   510
         End
         Begin VB.Label lblFabricante 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   42
            Top             =   1680
            Width           =   5085
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   750
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade:"
            Height          =   195
            Left            =   4320
            TabIndex        =   40
            Top             =   720
            Width           =   870
         End
         Begin VB.Label lblQtd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   39
            Top             =   960
            Width           =   885
         End
         Begin VB.Label lblPreco 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   38
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label lblCapPreco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preço:"
            Height          =   195
            Left            =   2760
            TabIndex        =   37
            Top             =   720
            Width           =   465
         End
         Begin VB.Label lblNomePro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            TabIndex        =   36
            Top             =   360
            Width           =   5085
         End
         Begin VB.Label lblCapPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.Label lblSolicitante 
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
         Left            =   1080
         TabIndex        =   73
         Top             =   7920
         Width           =   270
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitante:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   8040
         Width           =   780
      End
      Begin VB.Label lblAbertura 
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
         Left            =   1080
         TabIndex        =   71
         Top             =   7440
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abertura:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   7560
         Width           =   645
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
         Left            =   120
         TabIndex        =   68
         Top             =   7080
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   6720
         Width           =   630
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
         Left            =   840
         TabIndex        =   66
         Top             =   6720
         Width           =   270
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
      ScaleWidth      =   16140
      TabIndex        =   32
      Top             =   0
      Width           =   16170
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
         TabIndex        =   59
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
            Picture         =   "frmReposicao.frx":08CA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   60
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
            TabIndex        =   61
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
         TabIndex        =   9
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
            Picture         =   "frmReposicao.frx":12CC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   10
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
            TabIndex        =   11
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
         TabIndex        =   12
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
            Picture         =   "frmReposicao.frx":1656
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   13
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
            TabIndex        =   14
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
         TabIndex        =   15
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
            Picture         =   "frmReposicao.frx":18CA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   16
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
            TabIndex        =   17
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
         TabIndex        =   18
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
            Picture         =   "frmReposicao.frx":1A14
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   19
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
            TabIndex        =   20
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
         TabIndex        =   21
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
            Picture         =   "frmReposicao.frx":2416
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   22
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
            TabIndex        =   23
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
         TabIndex        =   24
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmReposicao.frx":47DD
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   25
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
            TabIndex        =   26
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
         TabIndex        =   27
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
            Picture         =   "frmReposicao.frx":4B67
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   28
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
            TabIndex        =   29
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
      Left            =   8520
      TabIndex        =   31
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSDBGrid.DBGrid dbgrdProd 
      Bindings        =   "frmReposicao.frx":4CD2
      Height          =   9855
      Left            =   5520
      OleObjectBlob   =   "frmReposicao.frx":4CE7
      TabIndex        =   65
      Top             =   600
      Width           =   10335
   End
End
Attribute VB_Name = "frmReposicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p As ProdutoSaida

Dim bDirecao As Integer
Dim bActive As Boolean
Dim Preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer
Dim sEdita As String 'se for <> de vazio, estará editando o item




Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    Dim cDesconto As Currency
    Dim cFinal As Currency
    Dim cPq As Currency
    On Error GoTo trata_erro:
    P_Status "Atualizando Valores..."
    
    iPos = dtProd.Recordset.AbsolutePosition
    dtProd.Refresh
    With dtProd.Recordset
        
        
        Do While Not .EOF
            cTotal = cTotal + !total
            
            
            cItens = cItens + !qtd_ite
            cPq = cPq + !pq_pro
            
            .MoveNext
        Loop
        lblValor.caption = FormatVal(cFinal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = dtProd.Recordset.RecordCount
        
        
        
        
    End With
        
    
    
    
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    
    lblProd.caption = dtProd.Recordset.RecordCount
    lblItens.caption = cItens

    
    
    P_Status ""
    Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub carrega()
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    strSQL = "select tr.* from tab_reposicao as tr"
    strSQL = strSQL & " where seq_rep='" & Me.tag & "' "
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        
        lblSolicitante.caption = !solicitante_rep
        lblAbertura.caption = !abertura_rep
        lblStatus.tag = !status_rep
        If !status_rep = "1" Then
            lblStatus.caption = "Para Lançamento"
        ElseIf !status_rep = "2" Then
            lblStatus.caption = "Para Atendimento"
            pcStatus.BackColor = vbRed
        ElseIf !status_rep = "3" Then
            lblStatus.caption = "Para Finalização"
            pcStatus.BackColor = vbBlue
        ElseIf !status_rep = "4" Then
            lblStatus.caption = "Concluída"
            pcStatus.BackColor = vbBlue
        ElseIf !status_rep = "5" Then
            lblStatus.caption = "Cancelada"
            pcStatus.BackColor = vbMagenta
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
    strSQL = Reposicao.F_BuscaItensRep(Me.tag)
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    soma
    
            
            
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbExclamation
End Sub
Sub soma()
    
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
    
    
    If lblStatus.tag <> "1" Then
        MsgBox "Só é possível cancelar esta reposição se ela estiver em aberto!", vbExclamation
        Exit Sub
    
    End If
    
    If dtProd.Recordset.RecordCount > 0 Then
    
    
    End If
    If MsgBox("Tem certeza que deseja cancelar esta reposicao?", vbYesNo + vbCritical) = vbYes Then
        'If operacao = "C" Then
        '    i = F_RetornoVariaveis
        '    If i Mod 2 > 0 Then
        '        ecf_CancelaCupom
        '        iItem = 0
        '    End If
        'End If
        
        strSQL = "delete from tab_reposicao where seq_rep='" & Me.tag & "';"
        Banco.Execute strSQL
        strSQL = strSQL & "delete from tab_itrep where seq_rep='" & Me.tag & "';"
        Banco.Execute strSQL
        
        
        
        
                
        If IsLoad("frmEdiReposicao") Then
            RefreshDataPosition frmEdiVendas.dtVendas
        End If
        
            Unload Me
        
    End If
End Sub




Private Sub cmndbtnExc_Prod_Click()
    excluir True
End Sub
Sub excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_erro
    
    If lblStatus.tag <> "1" Then
        MsgBox "Só é possível excluir um produto desta reposição se ela estiver em aberto!", vbExclamation
        Exit Sub
    
    End If
    
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & !nome_pro & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        Reposicao.F_DeletaitemRep Me.tag, !seq_ite
        P_buscaitens
    End With
    Exit Sub
trata_erro:
    P_MostraErro "Excluir()"
End Sub




Private Sub cmndbtnCancelar_Click()
    sEdita = ""
    txtbxProduto.Text = ""
    txtbxProduto.tag = ""
    txtbxProduto.SetFocus
    
End Sub


Private Sub cmbxFormaPag_Change()

End Sub

Private Sub cmbxFormaPag_DropDown()

End Sub






Private Sub cmndbtnInserir_Click()
    
    insere
End Sub

Private Sub cmndbtnLocaliza_Click()
    buscaProduto
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_erro
    
    
    
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
    
    If lblStatus.tag = "1" Then
        Reposicao.P_AlteraReposicao dtProd.Recordset!seq_rep, "2", False, False
        
    ElseIf lblStatus.tag = "2" Then
        Reposicao.P_AlteraReposicao dtProd.Recordset!seq_rep, "3", False, True
    ElseIf lblStatus.tag = "3" Then
        Reposicao.P_AlteraReposicao dtProd.Recordset!seq_rep, "4", True, False
            
    End If
    
        Unload Me
        
    
    
    
End Sub



Private Sub dbgrdProd_GotFocus()
    
    dbgrdProd.MarqueeStyle = 3
    
    
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error GoTo trata_erro

    If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
        
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


Private Sub dbgrdProd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub dtProd_Reposition()
    With dtProd.Recordset
        
    End With
        
End Sub


Private Sub Form_Activate()
    On Error GoTo trata_erro
    
    
    
    
    If bActive = True Then Exit Sub
    
    
    
    carrega
        
    P_Status "Buscando Itens..."
    txtbxProduto.SetFocus
    P_buscaitens
    P_AtualizaVenda
    P_Status ""
    
    bActive = True
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
        pcBotao_Click 1
    ElseIf KeyCode = vbKeyF1 Then
        pcBotao_Click 0
    ElseIf KeyCode = vbKeyF8 Then
        
    ElseIf KeyCode = vbKeyF9 Then
        
    ElseIf KeyCode = vbKeyF10 Then
        
        
    
        
    ElseIf KeyCode = vbKeyF11 Then
        
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
        frmECF.Show 1
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

    
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", "tab_parametros", "")
    
    resizeForm Me
    
    
    
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    
    
    dbgrdProd.Height = Me.Height - dbgrdProd.Top
    
    'pcCli.Width = pcProd.Width
    
    For iCount = 0 To pcBotao.count - 1
        pcBotao(iCount).Width = pcAtalhos.Width / (pcBotao.count)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).Width + 50
        End If
        
    Next
 
    
    
    
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
    
    mdiFrm.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFrm.Enabled = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub frmCli_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub lblMsg2_Click()

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
        
    ElseIf Index = 1 Then
        
    
    ElseIf Index = 0 Then
        
    ElseIf Index = 7 Then
        
    End If
End Sub
Private Sub pcBotao_GotFocus(Index As Integer)
    P_foco Index
End Sub

Private Sub pcBotao_LostFocus(Index As Integer)
    pcBotao(Index).BorderStyle = 0
End Sub

Private Sub pcMsg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub pcPromo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub




Private Sub pcProd_GotFocus()
    SendKeys Chr(9)
End Sub




Private Sub txtbxDescReais_GotFocus()
    
End Sub




Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    Dim p As Produto
    Dim rPesq As Recordset
    If KeyCode = vbKeyReturn Then
        buscaProduto

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
    On Error GoTo trata_erro
    Dim strSQL As String
    
    
    If txtbxProduto.tag = "" Then
        MsgBox "Primeiro escolha o produto!", vbAbortRetryIgnore
        txtbxProduto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxQtd.Text) = 0 Then
        MsgBox "Informe uma quantidade válida!", vbExclamation
        txtbxQtd.SetFocus
        Exit Sub
    End If
    strSQL = F_BuscaUmCampo("seq_ite", "tab_itrep", "cod_pro='" & txtbxProduto.tag & "' and seq_rep='" & Me.tag & "'")
    If strSQL <> "" Then
        MsgBox "Já existe um produto " & txtbxProduto.tag & " lançado no item " & strSQL, vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    If lblStatus.tag <> "1" Then
        MsgBox "Só é possível lançar um produto nesta reposição se ela estiver em aberto!", vbExclamation
        Exit Sub
    
    End If
    
    strSQL = F_GravaItemRep(Me.tag, txtbxProduto.tag, txtbxQtd.Text)
    P_buscaitens
    Exit Sub
    
trata_erro:
    P_MostraErro "Insere()"
End Sub
Sub buscaProduto()
    
    
    p = Produtos.F_BuscaProdSaida(txtbxProduto.Text, "B")
            
            
            
        If p.cod_pro <> "" Then
                               
            txtbxProduto.Text = p.nome_pro
            txtbxProduto.tag = p.cod_pro
            lblNomePro.caption = p.nome_pro
            
            lblCodPro.caption = p.cod_pro
'            lblSeqPro.caption = p.est.seq_pro
            
            
            lblSec.caption = p.sub.sec.nome_sec
            lblSubSec.caption = p.sub.nome_sub
            
            lblDescSub.caption = p.sub.descmax_sub
            lblPreco.caption = p.pv_pro
            lblRef.caption = p.ref_pro
            
            txtbxUnit.Text = FormatVal(p.pv_pro)
            
            
            
            
            
            SendKeys Chr(9)
        End If
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo trata_erro
    If KeyCode = vbKeyReturn Then
        cmndbtnInserir_Click
      
    End If
Exit Sub
trata_erro:
    P_MostraErro
End Sub
Sub p_totalItem()
    txtbxFinal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text))
    
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
    
End Sub

Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit, "###,##0.000"
    p_totalItem
End Sub
Sub P_GravaItemCancelado()
    On Error Resume Next
    Dim strSQL As String
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
Private Sub pcBotao_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    
    
    P_focoMouse Index
End Sub











