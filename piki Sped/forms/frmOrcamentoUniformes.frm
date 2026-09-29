VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOrcamentoUniformes 
   Caption         =   "Orçamento de Produtos e Serviços"
   ClientHeight    =   9915
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   18960
   Icon            =   "frmOrcamentoUniformes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcTab 
      Height          =   7815
      Left            =   5520
      ScaleHeight     =   7755
      ScaleWidth      =   13395
      TabIndex        =   47
      Top             =   600
      Width           =   13455
      Begin VB.PictureBox pcProd 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   13095
         TabIndex        =   49
         Top             =   120
         Width           =   13095
         Begin VB.TextBox txtProduto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   4695
         End
         Begin VB.TextBox txtQtd 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   9480
            TabIndex        =   59
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   10200
            TabIndex        =   61
            Text            =   "0,00"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtFinal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
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
            Left            =   11040
            TabIndex        =   63
            Text            =   "0,00"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "Inserir"
            Height          =   405
            Left            =   11040
            TabIndex        =   66
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "excluir"
            Height          =   405
            Left            =   11040
            TabIndex        =   68
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtDetalhamento 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   1365
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   65
            Top             =   840
            Width           =   10815
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "Alterar"
            Height          =   405
            Left            =   11040
            TabIndex        =   67
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox cmbBordado 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8040
            Style           =   1  'Simple Combo
            TabIndex        =   57
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbCor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4920
            Style           =   1  'Simple Combo
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox cmbSilk 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6600
            Style           =   1  'Simple Combo
            TabIndex        =   55
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   9480
            TabIndex        =   58
            Top             =   0
            Width           =   300
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
            TabIndex        =   50
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Unitário:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10200
            TabIndex        =   60
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor Final:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   11040
            TabIndex        =   62
            Top             =   0
            Width           =   780
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4920
            TabIndex        =   52
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Silk"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6600
            TabIndex        =   54
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bordado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8040
            TabIndex        =   56
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Detalhamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   600
            Width           =   1035
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   3255
         Left            =   120
         TabIndex        =   69
         Top             =   2760
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   5741
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Produtos Adicionados"
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
         TabIndex        =   70
         Top             =   2520
         Width           =   1860
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1200
         TabIndex        =   48
         Top             =   1320
         Width           =   555
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
      ScaleWidth      =   18930
      TabIndex        =   24
      Top             =   0
      Width           =   18960
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   18000
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   37
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   120
            Picture         =   "frmOrcamentoUniformes.frx":10CA
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   38
            Top             =   120
            Width           =   195
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sair(Esc)"
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
            TabIndex        =   39
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   5280
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   34
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   120
            Picture         =   "frmOrcamentoUniformes.frx":1235
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   35
            Top             =   105
            Width           =   240
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F7 - Efetivar"
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
            TabIndex        =   36
            Top             =   120
            Width           =   945
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   3480
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   31
         Top             =   0
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
            Picture         =   "frmOrcamentoUniformes.frx":15BF
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   32
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
            Index           =   2
            Left            =   480
            TabIndex        =   33
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
         Index           =   1
         Left            =   1920
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   28
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            Picture         =   "frmOrcamentoUniformes.frx":3986
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   29
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblBotao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F4 - Cliente"
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
            TabIndex        =   30
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.PictureBox pcBotao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   1800
         TabIndex        =   25
         Top             =   0
         Width           =   1800
         Begin VB.PictureBox pcIco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   120
            Picture         =   "frmOrcamentoUniformes.frx":3D3C
            ScaleHeight     =   285
            ScaleWidth      =   285
            TabIndex        =   26
            Top             =   120
            Width           =   285
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
            Index           =   0
            Left            =   480
            TabIndex        =   27
            Top             =   120
            Width           =   1050
         End
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   18930
      TabIndex        =   23
      Top             =   7860
      Width           =   18960
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair (Esc)"
         Height          =   735
         Left            =   17760
         Picture         =   "frmOrcamentoUniformes.frx":41F2
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4305
      ScaleWidth      =   5385
      TabIndex        =   12
      Top             =   600
      Width           =   5415
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Width           =   5175
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Tag             =   "00000"
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox txtNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCB683&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Tag             =   "00000"
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txtTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Tag             =   "00000"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtComprador 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   5
         Tag             =   "00000"
         Top             =   840
         Width           =   3375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   825
         ScaleWidth      =   5385
         TabIndex        =   13
         Top             =   3240
         Width           =   5415
         Begin VB.TextBox txtFormasPag 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCB683&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   11
            Tag             =   "01"
            Text            =   "Dinheiro à Vista"
            Top             =   360
            Width           =   4815
         End
         Begin VB.CommandButton cmdFormasPag 
            Height          =   330
            Left            =   4920
            Picture         =   "frmOrcamentoUniformes.frx":4780
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Forma de Pagamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1560
         End
         Begin VB.Label lblTipoPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carteira"
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
            TabIndex        =   14
            Tag             =   "06"
            Top             =   120
            Width           =   660
         End
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Height          =   330
         Left            =   4920
         Picture         =   "frmOrcamentoUniformes.frx":4B0A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblObs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprador/Contato"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   1395
      End
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      ForeColor       =   &H80000008&
      Height          =   7320
      Left            =   0
      ScaleHeight     =   7290
      ScaleWidth      =   5385
      TabIndex        =   15
      Top             =   540
      Width           =   5415
      Begin VB.TextBox txtDesconto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   40
         Text            =   "0,00"
         Top             =   7200
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2385
         ScaleWidth      =   5385
         TabIndex        =   16
         Top             =   4320
         Width           =   5415
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00F8AC67&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   17
            Top             =   2160
            Width           =   5415
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
               TabIndex        =   21
               Top             =   0
               Width           =   900
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
               TabIndex        =   20
               Tag             =   "!"
               Top             =   0
               Width           =   105
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
               Left            =   4200
               TabIndex        =   19
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unidades:"
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
               Left            =   3120
               TabIndex        =   18
               Top             =   0
               Width           =   945
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
               Size            =   63.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1440
            Left            =   1560
            TabIndex        =   22
            Tag             =   "!"
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.Label lblTotalSemDesconto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   105
         TabIndex        =   43
         Tag             =   "!"
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   6960
         Width           =   360
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   41
         Top             =   6960
         Width           =   735
      End
   End
   Begin VB.Menu mnOpt 
      Caption         =   "&Opções"
      Begin VB.Menu mnClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnEfetivar 
         Caption         =   "&Efetivar Orçamento"
      End
      Begin VB.Menu div 
         Caption         =   "-"
      End
      Begin VB.Menu mnInserir 
         Caption         =   "Inserir item"
      End
      Begin VB.Menu mnAlterar 
         Caption         =   "Alterar item"
      End
      Begin VB.Menu mnExcluir 
         Caption         =   "Excluir item selecionado"
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frmOrcamentoUniformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub lblCaixa_Click()

End Sub
Sub buscaItens()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim colunas As String
    Dim ColumData As String
    
    colunas = "Item,Id,Código,Descrição,Qtde,Unitário,Total,Cor,Silk,Bordado,Detalhamento"
    ColumData = "$item,Id,Cod_Pro,produto,quantidade,unitario,total,cor,silk,bordado,detalhamento"
    
    strSQL = "select oi.* from orcamento_itens as oi where id_orcamento=" & Me.tag
    strSQL = strSQL & " order by oi.id "
    
    abreObjetoAdo ado, strSQL
    
    
    AjustaGridData grid, 11, 1, 1, colunas, ColumData, "500,0,1000,4000,800,1000,1000,1000,1000,1000,5000", ado, "4,6"
        
    
    
    Totaliza
    
End Sub
Sub Carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select o.*,c.nome_cli,c.cnpj_cli, f.guerra_fun,fp.* from orcamentos as o inner join clientes as c on o.cod_cli=c.cod_cli "
    strSQL = strSQL & " left join " & TAB_FUNCIONARIOS & " as f on o.cod_usu=f.cod_fun "
    strSQL = strSQL & " inner join " & TAB_FORMASPAG & " as  fp on fp.cod_for=o.cod_for "
    strSQL = strSQL & " where 1>0 "
    strSQL = strSQL & " and o.id='" & Me.tag & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            txtComprador.Text = TiraNull(!contato)
            txtDesconto.Text = F_GravaDec(!desconto)
            txtEmail.Text = TiraNull(!email)
            txtTel.Text = TiraNull(!telefone)
            txtNomeCli.Text = TiraNull(!cliente)
            txtFormasPag.Text = TiraNull(!nome_For)
            txtFormasPag.tag = TiraNull(!cod_For)
            txtObs.Text = TiraNull(!observacao)
        End If
        .Close
    End With
    
    
    buscaItens
    
End Sub
Private Sub txtbxProduto_Change()

End Sub

Private Sub cmdCor_Change()

End Sub

Private Sub cmbCor_GotFocus()
    P_SelecTexto cmbCor
End Sub

Private Sub cmbCor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbCor_LostFocus()
    P_AjustaTexto cmbCor, nada_muda
End Sub

Private Sub cmbSilk_GotFocus()
    P_SelecTexto cmbSilk
End Sub

Private Sub cmbSilk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbSilk_LostFocus()
    P_AjustaTexto cmbSilk, nada_muda
End Sub

Private Sub cmbBordado_GotFocus()
    P_SelecTexto cmbBordado
End Sub

Private Sub cmbBordado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub cmbBordado_LostFocus()
    P_AjustaTexto cmbBordado, nada_muda
End Sub





Private Sub cmdAlterar_Click()
    alterar_item
End Sub

Private Sub cmdExcluir_Click()

    excluir_item

End Sub
Sub excluir_item()
    Dim s As String
    On Error GoTo trata_Erro
    With grid
        If .row > 0 Then
            .col = 1
            s = .Text
            If MsgBox("Confirma remoção do item?", vbYesNo + vbCritical) = vbYes Then
                orcamento.exclui_orcamento_item s
                buscaItens
            End If
        End If
    End With
trata_Erro:
End Sub
Sub alterar_item()
        Dim s As String
    On Error GoTo trata_Erro
    With grid
        
        If cmdInserir.tag <> "" Then
            txtProduto.Text = ""
            txtQtd.Text = "1"
            txtDetalhamento.Text = ""
            cmdInserir.caption = "Inserir"
            cmdAlterar.caption = "Alterar"
            cmdInserir.tag = ""
        ElseIf .row > 0 Then
            .col = 1
            s = .Text
            cmdInserir.tag = s
            cmdInserir.caption = "Confirmar"
            cmdAlterar.caption = "Cancelar"
            .col = 3
            txtProduto.Text = .Text
            .col = 2
            txtProduto.tag = .Text
            .col = 4
            txtQtd.Text = .Text
            .col = 5
            txtQtd.Text = .Text
            .col = 6
            txtFinal.Text = .Text
            .col = 7
            cmbCor.Text = .Text
            
            .col = 8
            cmbSilk.Text = .Text
            
            .col = 9
            cmbBordado.Text = .Text
            .col = 10
            txtDetalhamento.Text = .Text
            
            
            
        End If
    End With
trata_Erro:

End Sub
Private Sub cmdFormasPag_Click()
    buscaFormasPag
    
End Sub
Sub buscaFormasPag()
    Dim f As FormasPag.FormasPag
    f = FormasPag.F_BuscaFormasPag("", txtFormasPag.Text)
    If f.cod_For <> "" Then
        txtFormasPag.Text = f.nome_For
        txtFormasPag.tag = f.cod_For
    End If
End Sub
Private Sub cmdInserir_Click()
    If cmdInserir.tag = "" Then
        insere
    Else
        altera
    End If
    
End Sub
Sub insere()
    Dim s As String
    If txtProduto.Text = "" Then
        MsgBox "Informe o produto", vbExclamation
        txtProduto.SetFocus
        Exit Sub
    End If
    
    If CCur(txtFinal.Text) <= 0 Then
        MsgBox "Informe um valor válido!", vbExclamation
        txtFinal.SetFocus
        Exit Sub
    End If
    
    Iniciar
        
        
    s = insereOrcamento_item(Me.tag, txtProduto.tag, txtProduto.Text, txtQtd.Text, txtUnit.Text, cmbCor.Text, cmbSilk.Text, cmbBordado.Text, txtDetalhamento.Text)
        
    buscaItens
End Sub
Sub altera()

End Sub
Sub Iniciar()
    If Me.tag = "" Then
        Me.tag = orcamento.insereOrcamento(txtNomeCli.tag, txtNomeCli.Text, txtTel.Text, txtComprador.Text, txtEmail.Text, txtFormasPag.tag)
        Me.caption = "Orçamento " & Me.tag & " - " & txtNomeCli.Text
    End If
End Sub


Sub buscaCli()
    Dim c As cliente
    c = Clientes.F_BuscaCliente(txtNomeCli.Text)
    If c.cod_cli <> "" Then
        txtNomeCli.Text = c.nome_Cli
        txtNomeCli.tag = c.cod_cli
        txtTel.Text = c.telefone_cli
        txtEmail.Text = c.email_cli
    End If
    
End Sub



Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmndbtnBuscar_Click()
    buscaCli
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        Carrega
    End If
    bActive = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        sair
    End If
End Sub

Private Sub Form_Load()
    BuscaPic
    
    Set ado = New ADODB.Recordset
    resizeForm Me
End Sub
Sub BuscaPic()
    On Error Resume Next
    Set pcBot.Picture = LoadPicture(App.path & "\pic\logo_topo.jpg")
End Sub
Private Sub pcStatusBar_Click()

End Sub

Sub ajustaGrid()
    grid.Cols = 4
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PosicionaGrade pcTab, Me, pcBot
    PosicionaGrade grid, pcTab, pcBot, 200, True
    grid.height = pcTab.height - grid.Top - 200
    pcProd.width = grid.width
    posicionaADireita cmdInserir, pcProd, 100
    cmdAlterar.Left = cmdInserir.Left
    cmdExcluir.Left = cmdInserir.Left
    txtDetalhamento.width = cmdInserir.Left - txtDetalhamento.Left - 200
    posicionaADireita cmdSair, pcBot
    posicionaAoFim cmdSair, pcBot
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ado.Close
End Sub


Private Sub grid_DblClick()
    cmdAlterar_Click
End Sub

Private Sub grid_RowColChange()
    If grid.rows > 1 Then
        cmdAlterar.Visible = True
        cmdExcluir.Visible = True
    Else
        cmdAlterar.Visible = False
        cmdExcluir.Visible = False
    End If
End Sub


Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub mnSair_Click()
    sair
End Sub
Sub sair()
    On Error Resume Next
    ado.Close
    Unload Me
End Sub


Private Sub pcBotao_Click(Index As Integer)
    If Index = 0 Then
        
        ImprimeOrcamento Me.tag
    ElseIf Index = 1 Then
        frmEdiClientes.Show 1
    ElseIf Index = 2 Then
        If MsgBox("Deseja realmente excluir o orçamento?", vbYesNo + vbQuestion) = vbYes Then
            
        End If
    ElseIf Index = 3 Then
        efetivar
    ElseIf Index = 4 Then
        cmdSair_Click
    
    End If
End Sub
Sub efetivar()
    If MsgBox("Confirma efetivação do orçamento em uma nova venda?", vbYesNo + vbQuestion) Then
            num_ven = orcamento.GeraVenda(Me.tag)
            
            If num_ven <> "" Then
                orcamento.AlteraOrcamentoStatus Me.tag, fechado
            
                frmVenda.tag = num_ven
                Me.Hide
                frmVenda.Show 1
                
                Unload Me
            Else
                MsgBox "Não gerou o orçamento!", vbExclamation
            End If
        End If
End Sub
Private Sub Picture1_Click()

End Sub

Private Sub pcIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub txtDesconto_GotFocus()
    P_SelecTexto txtDesconto
    Totaliza
End Sub

Private Sub txtDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtDesconto_LostFocus()
    P_ValorLostFocus txtDesconto
    Totaliza
End Sub

Private Sub txtDetalhamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        cmdInserir_Click
    End If
End Sub

Private Sub txtFinal_GotFocus()
    P_SelecTexto txtFinal
End Sub

Private Sub txtFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtFinal_LostFocus()
    P_ValorLostFocus txtFinal
    txtUnit.Text = FormatVal(CCur(txtFinal.Text) / CCur(txtQtd.Text))
End Sub

Private Sub txtNomeCli_LostFocus()
    grava
End Sub
Sub grava()
    If Me.tag = "" Then
        Iniciar
    End If
    orcamento.AlteraOrcamento Me.tag, txtNomeCli.tag, txtNomeCli.Text, txtTel.Text, txtFormasPag.tag, lblValor.caption, txtDesconto.Text, txtComprador.Text, txtEmail.Text, txtObs.Text
End Sub

Private Sub txtObs_LostFocus()
    P_AjustaTexto txtObs
End Sub

Private Sub txtProduto_Change()
    txtProduto.tag = ""
End Sub

Sub buscaProduto()
    Dim p As Produto
    
    p = F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        If p.cod_pro = "0" Then
        
        Else
            txtUnit.Text = p.pv_pro
        End If
        cmbCor.SetFocus
    End If
    
    
End Sub

Private Sub txtProduto_DblClick()
    buscaProduto
End Sub

Private Sub txtProduto_GotFocus()
    P_SelecTexto txtProduto
End Sub

Private Sub txtProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProduto
    End If
End Sub

Sub Totaliza()
    
   On Error GoTo trata_Erro
    With ado
        Dim cTotal As Currency
        Dim cUn As Currency
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            cTotal = cTotal + !total
            cUn = cUn + !Quantidade
            .MoveNext
            
        Loop
        
        
        lblTotalSemDesconto.caption = FormatVal(cTotal)
        lblValor.caption = FormatVal(cTotal - (CCur(txtDesconto.Text)))
        lblItens.caption = cUn
        lblProd.caption = .RecordCount
        If CCur(txtDesconto.Text) > CCur(cTotal) Then
            MsgBox "Valor inválido!", vbExclamation
            txtDesconto.Text = "0,00"
            txtDesconto.SetFocus
        End If
    End With
    grava
trata_Erro:
End Sub

Private Sub txtQtd_GotFocus()
    P_SelecTexto txtQtd
End Sub

Private Sub txtQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtQtd_LostFocus()
    txtFinal.Text = FormatVal(CCur(txtUnit.Text) * CCur(txtQtd.Text))
End Sub

Private Sub txtTel_GotFocus()
    P_SelecTexto txtTel
End Sub
Private Sub txtTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtTel_LostFocus()
    P_AjustaTexto txtTel, nada_muda
    grava
End Sub

Private Sub txtComprador_GotFocus()
    P_SelecTexto txtComprador
End Sub
Private Sub txtComprador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtComprador_LostFocus()
    P_AjustaTexto txtComprador, nada_muda
    grava
End Sub
Private Sub txtEmail_GotFocus()
    P_SelecTexto txtEmail
End Sub
Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtEmail_LostFocus()
    P_AjustaTexto txtEmail, nada_muda
    grava
End Sub

Private Sub txtUnit_GotFocus()
    P_SelecTexto txtUnit
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtUnit_LostFocus()
    P_ValorLostFocus txtUnit
    txtFinal.Text = FormatVal(CCur(txtUnit.Text) * CCur(txtQtd.Text))
End Sub
