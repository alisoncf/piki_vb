VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAcerto2 
   Caption         =   "Acerto de:"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16695
   Icon            =   "frmAcerto2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   16695
   Begin MSDBGrid.DBGrid dbgrdAcerto 
      Bindings        =   "frmAcerto2.frx":33E2
      Height          =   2895
      Left            =   0
      OleObjectBlob   =   "frmAcerto2.frx":33F9
      TabIndex        =   99
      Top             =   2040
      Width           =   15135
   End
   Begin VB.Data dtAcerto 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Width           =   1575
   End
   Begin VB.PictureBox pcTotalItens 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   15240
      ScaleHeight     =   3990
      ScaleWidth      =   1425
      TabIndex        =   47
      Top             =   1950
      Width           =   1455
      Begin VB.TextBox txtbxDataBase 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtbxDifAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtbxValorNegociado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   20
         Text            =   "0,00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Negociado:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label lblTotalInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H000040C0&
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Tag             =   "!"
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Inicial:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label44 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total dos Itens:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblDelExcluir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Del> Excluir"
         Height          =   195
         Left            =   240
         MouseIcon       =   "frmAcerto2.frx":412C
         MousePointer    =   4  'Icon
         TabIndex        =   48
         Top             =   3600
         Width           =   930
      End
      Begin VB.Label lblTotalAcerto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   120
         TabIndex        =   26
         Tag             =   "!"
         Top             =   2280
         Width           =   1185
      End
   End
   Begin VB.PictureBox pcCli 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   16665
      TabIndex        =   52
      Top             =   495
      Width           =   16695
      Begin VB.Frame frCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   17415
         Begin VB.OptionButton optnCli 
            BackColor       =   &H00FF8080&
            Caption         =   "&Cliente"
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optnRep 
            BackColor       =   &H00FF8080&
            Caption         =   "&Representante"
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtbxClienteAcerto 
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
            MaxLength       =   50
            TabIndex        =   5
            Top             =   240
            Width           =   4815
         End
         Begin VB.CommandButton cmndbtnClienteAcerto 
            Height          =   285
            Left            =   6600
            Picture         =   "frmAcerto2.frx":46B6
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblCidadeCli 
            BackColor       =   &H007DBFDB&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   16440
            TabIndex        =   70
            Top             =   240
            Width           =   3090
         End
         Begin VB.Label lblBairroCli 
            BackColor       =   &H007DBFDB&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   13680
            TabIndex        =   69
            Top             =   240
            Width           =   2730
         End
         Begin VB.Label lblEnderecoCli 
            BackColor       =   &H007DBFDB&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10560
            TabIndex        =   68
            Top             =   240
            Width           =   2970
         End
         Begin VB.Label lblRgCli 
            BackColor       =   &H007DBFDB&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   9240
            TabIndex        =   67
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label lblCpfCli 
            BackColor       =   &H007DBFDB&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   6960
            TabIndex        =   66
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label lblCli 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "C&liente (F4):"
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
            Left            =   1800
            TabIndex        =   4
            Tag             =   "C"
            Top             =   0
            Width           =   1050
         End
         Begin VB.Label Label40 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "CPF:"
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
            Left            =   6960
            TabIndex        =   65
            Top             =   0
            Width           =   420
         End
         Begin VB.Label Label39 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10560
            TabIndex        =   64
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   13680
            TabIndex        =   63
            Top             =   0
            Width           =   450
         End
         Begin VB.Label Label37 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   16440
            TabIndex        =   62
            Top             =   0
            Width           =   540
         End
         Begin VB.Label label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "RG:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   9240
            TabIndex        =   61
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.PictureBox pcProdAcerto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   14655
         TabIndex        =   53
         Top             =   840
         Width           =   14655
         Begin VB.TextBox txtbxReman 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   10
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtbxVlrFinalAcerto 
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
            Left            =   7080
            TabIndex        =   16
            Text            =   "0,00"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtbxDiferenca 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   8280
            TabIndex        =   18
            Text            =   "0,00"
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtbxUnitAcerto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   6120
            TabIndex        =   14
            Text            =   "0,00"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtbxQtdAcerto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   5520
            TabIndex        =   12
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtbxProdAcerto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label lblSeqConAcerto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Left            =   10320
            TabIndex        =   59
            Top             =   0
            Width           =   120
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seq. Con:"
            ForeColor       =   &H00CF614E&
            Height          =   195
            Left            =   9480
            TabIndex        =   58
            Top             =   0
            Width           =   705
         End
         Begin VB.Label lblDataConAcerto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Left            =   10920
            TabIndex        =   57
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Lançamento:"
            ForeColor       =   &H00CF614E&
            Height          =   195
            Left            =   9480
            TabIndex        =   56
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label lblRefAcerto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Left            =   12480
            TabIndex        =   55
            Top             =   0
            Width           =   120
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referência:"
            ForeColor       =   &H00CF614E&
            Height          =   195
            Left            =   11520
            TabIndex        =   54
            Top             =   0
            Width           =   825
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reman:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4920
            TabIndex        =   9
            Top             =   0
            Width           =   555
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Valor Final:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7080
            TabIndex        =   15
            Top             =   0
            Width           =   780
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diferença:"
            Height          =   195
            Left            =   8280
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Unitário:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6120
            TabIndex        =   13
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Produto (F2)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   870
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "&Qtd:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5640
            TabIndex        =   11
            Top             =   0
            Width           =   300
         End
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   16695
      TabIndex        =   49
      Top             =   0
      Width           =   16695
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   10455
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   495
         Left            =   10440
         Picture         =   "frmAcerto2.frx":47AF
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   495
      End
      Begin VB.Label lblCpf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11640
         TabIndex        =   51
         Top             =   45
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
         Height          =   195
         Left            =   11160
         TabIndex        =   50
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   16695
      TabIndex        =   71
      Top             =   5970
      Width           =   16695
      Begin VB.CommandButton cmndbtnConfAcerto 
         Caption         =   "F7 - Confirmar"
         Height          =   735
         Left            =   10920
         Picture         =   "frmAcerto2.frx":5479
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "F12 - Sair"
         Height          =   735
         Left            =   12720
         Picture         =   "frmAcerto2.frx":55F0
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2040
         Width           =   1695
      End
      Begin VB.PictureBox pcTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H007DB4BD&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   10920
         ScaleHeight     =   705
         ScaleWidth      =   3465
         TabIndex        =   87
         Top             =   1200
         Width           =   3495
         Begin VB.Label lblVlrAcertoDiv 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   825
            Left            =   1200
            TabIndex        =   90
            Tag             =   "!"
            Top             =   -120
            Width           =   1260
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   840
            TabIndex        =   89
            Top             =   240
            Width           =   195
         End
         Begin VB.Label lblParcelas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   825
            Left            =   360
            TabIndex        =   88
            Top             =   -90
            Width           =   360
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   10920
         ScaleHeight     =   945
         ScaleWidth      =   3465
         TabIndex        =   72
         Top             =   120
         Width           =   3495
         Begin VB.Label lblValorFinal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            TabIndex        =   78
            Tag             =   "!"
            Top             =   292
            Width           =   540
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Final:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblCapComiss 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comissão:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   75
            Width           =   720
         End
         Begin VB.Label lblComissao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            TabIndex        =   75
            Tag             =   "!"
            Top             =   0
            Width           =   540
         End
         Begin VB.Label lblTotalDistribuido 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1560
            TabIndex        =   74
            Tag             =   "!"
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Distribuído:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   675
            Width           =   1215
         End
      End
      Begin TabDlg.SSTab sstb 
         Height          =   2775
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "F8 - A Prazo"
         TabPicture(0)   =   "frmAcerto2.frx":575B
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtbxFormaPag"
         Tab(0).Control(1)=   "txtbxDescCob"
         Tab(0).Control(2)=   "cmbxOpc"
         Tab(0).Control(3)=   "txtbxMsf"
         Tab(0).Control(4)=   "txtbxTipoCob"
         Tab(0).Control(5)=   "txtbxVenIniAcerto"
         Tab(0).Control(6)=   "Picture3"
         Tab(0).Control(7)=   "txtbxEntrada"
         Tab(0).Control(8)=   "txtbxTotalAprazo"
         Tab(0).Control(9)=   "msflxgrd"
         Tab(0).Control(10)=   "cmndbtnCob"
         Tab(0).Control(11)=   "cmndbtnFormasPag"
         Tab(0).Control(12)=   "Label4"
         Tab(0).Control(13)=   "lblTipoPag"
         Tab(0).Control(14)=   "Label12"
         Tab(0).Control(15)=   "Label11"
         Tab(0).Control(16)=   "lblTipoCob"
         Tab(0).Control(17)=   "Label43"
         Tab(0).Control(18)=   "Label8"
         Tab(0).ControlCount=   19
         TabCaption(1)   =   "F9 - À Vista"
         TabPicture(1)   =   "frmAcerto2.frx":5777
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "msflxgrdAv"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmndbtnCartao"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmndbtnDin"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmndbtnCh"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Picture5"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtbxFormaPag 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -70800
            TabIndex        =   38
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtbxDescCob 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -71280
            TabIndex        =   36
            Text            =   "0"
            Top             =   600
            Width           =   375
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   9480
            ScaleHeight     =   465
            ScaleWidth      =   945
            TabIndex        =   95
            Top             =   360
            Width           =   975
            Begin VB.CommandButton cmndbnEditAv 
               Height          =   375
               Left            =   480
               Picture         =   "frmAcerto2.frx":5793
               Style           =   1  'Graphical
               TabIndex        =   97
               Top             =   50
               Width           =   375
            End
            Begin VB.CommandButton cmndbtnExclAv 
               Height          =   375
               Left            =   50
               Picture         =   "frmAcerto2.frx":6195
               Style           =   1  'Graphical
               TabIndex        =   96
               Top             =   50
               Width           =   375
            End
         End
         Begin VB.CommandButton cmndbtnCh 
            Caption         =   "Adcionar Pgto em Cheque"
            Height          =   495
            Left            =   2280
            TabIndex        =   94
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cmbxOpc 
            Height          =   315
            ItemData        =   "frmAcerto2.frx":651F
            Left            =   -74400
            List            =   "frmAcerto2.frx":6529
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   1860
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtbxMsf 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   -72360
            TabIndex        =   84
            Top             =   1620
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtbxTipoCob 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -74880
            TabIndex        =   31
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtbxVenIniAcerto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -72480
            TabIndex        =   34
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtbxDescontoAvista 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73800
            TabIndex        =   83
            Text            =   "27"
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox txtbxVlrFinalAvista 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   -65400
            TabIndex        =   82
            Text            =   "0"
            Top             =   2760
            Width           =   735
         End
         Begin VB.CommandButton cmndbtnDin 
            Caption         =   "Adcionar Pgto em Dinheiro"
            Height          =   495
            Left            =   200
            TabIndex        =   81
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnCartao 
            Caption         =   "Adcionar Pgto em Cartão"
            Height          =   495
            Left            =   4440
            TabIndex        =   80
            Top             =   360
            Width           =   2055
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   -65520
            ScaleHeight     =   465
            ScaleWidth      =   945
            TabIndex        =   79
            Top             =   480
            Width           =   975
            Begin VB.CommandButton cmndbtnEditAp 
               Height          =   375
               Left            =   480
               Picture         =   "frmAcerto2.frx":654C
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   50
               Width           =   375
            End
            Begin VB.CommandButton cmndbtnExclAp 
               DragMode        =   1  'Automatic
               Enabled         =   0   'False
               Height          =   375
               Left            =   50
               Picture         =   "frmAcerto2.frx":6F4E
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   50
               Width           =   375
            End
         End
         Begin VB.TextBox txtbxEntrada 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   -68520
            TabIndex        =   41
            Text            =   "0,00"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtbxTotalAprazo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   -67320
            TabIndex        =   43
            Text            =   "0,00"
            Top             =   600
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrd 
            Height          =   1575
            Left            =   -74880
            TabIndex        =   46
            Top             =   1080
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdAv 
            Height          =   1575
            Left            =   240
            TabIndex        =   93
            Top             =   960
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.CommandButton cmndbtnCob 
            Height          =   285
            Left            =   -72840
            Picture         =   "frmAcerto2.frx":72D8
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmndbtnFormasPag 
            Height          =   285
            Left            =   -68880
            Picture         =   "frmAcerto2.frx":7662
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Forma de Pagamento:"
            Height          =   195
            Left            =   -70800
            TabIndex        =   37
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label lblTipoPag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DIN"
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
            Left            =   -69120
            TabIndex        =   98
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada:"
            Height          =   195
            Left            =   -68520
            TabIndex        =   40
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total A Prazo"
            Height          =   195
            Left            =   -67320
            TabIndex        =   42
            Top             =   360
            Width           =   960
         End
         Begin VB.Label lblTipoCob 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrança:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   30
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ven. Ini:"
            Height          =   195
            Left            =   -72480
            TabIndex        =   33
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Final:"
            Height          =   195
            Left            =   -66360
            TabIndex        =   86
            Top             =   2760
            Width           =   780
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(%)"
            Height          =   195
            Left            =   -71280
            TabIndex        =   35
            Top             =   360
            Width           =   210
         End
      End
   End
End
Attribute VB_Name = "frmAcerto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bEnterCell As Boolean
Dim prod As ProdutoSaida
  
Dim bDirecao As Integer
Dim bActive As Boolean
Dim Preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer

Public Atacado As Boolean

Private Type prodPromo
    Nome As String
    Preco As String
End Type
Dim pPromo() As prodPromo
Dim QtdPromo As Integer
Dim iIndPromo As Integer
Dim c As New ADODB.Connection
Dim r As New ADODB.Recordset
Dim cFile As New Collection
Dim ai As New ADODB.Stream
Dim Indice As Integer




Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    
        
   
    
       
    
    
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
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
    If dtProd.Recordset.RecordCount > 0 Then
    
    If strExclVenda_usu <> "S" Then
        strSQL = F_senhaReturnUsu
        If F_BuscaUmCampo("exclvenda_usu", "tab_defusu", "cod_usu='" & strSQL & "'") <> "S" Then
            MsgBox "Usuário Não Autorizado!", vbCritical
            Exit Sub
        End If
    End If
    End If
    If MsgBox("Tem certeza que deseja cancelar esta venda?", vbYesNo + vbCritical) = vbYes Then
        'If operacao = "C" Then
        '    i = F_RetornoVariaveis
        '    If i Mod 2 > 0 Then
        '        ecf_CancelaCupom
        '        iItem = 0
        '    End If
        'End If
        With dtProd.Recordset
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                excluir False
                .MoveNext
            Loop
        End With
        strSQL = "insert into tab_cancelamentos(num_ven,data_can,valor_ven,tipo_ven,"
        strSQL = strSQL & "cod_cai,cod_tur)values("
        strSQL = strSQL & "'" & Me.tag & "',"
        strSQL = strSQL & F_GravaDataTempo(Now) & ","
        strSQL = strSQL & F_GravaDec(lblValor.caption) & ","
        strSQL = strSQL & "'" & operacao & "',"
        strSQL = strSQL & "'" & strCod_Cai & "',"
        strSQL = strSQL & "'" & strCod_Tur & "') "
        Banco.Execute strSQL
        
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
            RefreshDataPosition frmEdiVendas.dtVendas
        End If
        
        If MsgBox("Deseja Gerar uma Nova Venda?" & Chr(13) & "Clique em não caso queira abandonar o modo de vendas!", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            strSQL = F_GeraNovaVenda("00000")
            If strSQL = "" Then
                MsgBox "Erro na geração da venda!", vbExclamation
                Unload Me
            End If
            iItem = 0
            Me.tag = strSQL
            Form_Activate
        Else
            iItem = 0
            Unload Me
        End If
    End If
End Sub



Private Sub cmndbtnEditProd_Click()
    On Error GoTo trata_Erro
    Dim iPos As Integer
    Dim strSQL As String
    Dim cUlt_Qtd As Currency
    
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        iPos = .AbsolutePosition
        frmValoresProd.lblCod_Pro.caption = !cod_pro
        frmValoresProd.lblNome_pro.caption = !nome_pro
        frmValoresProd.txtbxCusto.Text = FormatVal(!custo_pro)
        frmValoresProd.txtbxPreco.Text = FormatVal(!unit_ite)
        frmValoresProd.txtbxQtd.Text = FormatVal(!qtd_ite)
        frmValoresProd.tag = Me.tag
        cUlt_Qtd = !qtd_ite
        On Error Resume Next
        frmValoresProd.Show 1
        If CCur(frmValoresProd.txtbxQtd.Text) <= 0 Then
            MsgBox "Informe um valor decente!", vbExclamation
            Exit Sub
        End If
        strSQL = "Update tab_itvendas set qtd_ite=" & _
        F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & ", unit_ite=" & _
        F_GravaDec(frmValoresProd.txtbxPreco.Text)
        
        strSQL = strSQL & " where num_ven='" & Me.tag & "' "
        strSQL = strSQL & " and cod_pro='" & !cod_pro & "' "
        Banco.Execute strSQL
        'ajustando estoque reservado
        strSQL = "Update tab_produtos set reservado_pro=reservado_pro-" & F_GravaDec(CStr(cUlt_Qtd))
        strSQL = strSQL & "+" & F_GravaDec(frmValoresProd.txtbxQtd.Text)
        strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
        Banco.Execute strSQL
        Unload frmValoresProd
    End With
    dtProd.Refresh
    dtProd.Recordset.Move iPos
    
trata_Erro:
MsgBox "Err: cmndbtnEditProd_Click() "
End Sub

Private Sub cmndbtnExc_Prod_Click()
    excluir True
End Sub
Sub excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    With dtProd.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & !desc_ite & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        iPos = .AbsolutePosition
        strSQL = "delete from tab_itvendas where seq_ite='" & !seq_ite & "' "
        
        strSQL = strSQL & " and num_ven='" & Me.tag & "' "
        Banco.Execute strSQL
        'P_GravaItemCancelado
        'voltando o estoque
'        If pcBaixa.Visible = False Then
           strSQL = "update tab_produtos set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))

            
            

            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            Banco.Execute strSQL
        
        dtProd.Refresh
        If iPos = .RecordCount - 1 Then
            If dtProd.Recordset.RecordCount > 0 Then
                dtProd.Recordset.MoveLast
            End If
        Else
            dtProd.Recordset.Move iPos
        End If
        If bQuestion Then
            
        End If
        If dtProd.Recordset.RecordCount = 0 Then
            txtbxQtd.Enabled = True
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
    Exit Sub
trata_Erro:
    P_MostraErro "Excluir()"
End Sub


Private Sub chkEntrega_Click()
    If chkEntrega.Value = 1 Then
        lblEnd.Enabled = True
        lblBairro.Enabled = True
        lblFone.Enabled = True
        txtbxEndereco.Enabled = True
        txtbxBairro.Enabled = True
        txtbxFone.Enabled = True
        txtbxEndereco.SetFocus
    Else
        lblEnd.Enabled = False
        lblBairro.Enabled = False
        lblFone.Enabled = False
        txtbxEndereco.Enabled = False
        txtbxBairro.Enabled = False
        txtbxFone.Enabled = False
    End If
    
End Sub

Private Sub cmbxFormaspag_Click()
    
    If txtbxTipoCob.tag = "" Then
        MsgBox "Escolha o tipo da cobrança!", vbExclamation
        txtbxTipoCob.SetFocus
        Exit Sub
    End If
    If CCur(lblValorFinal.caption) = 0 Then
        MsgBox "Não há valores para serem divididos!", vbExclamation
        
        Exit Sub
    End If
    P_CalcTotalAprazo
    lblParcelas.caption = F_BuscaUmCampo("parcelas_for", "tab_formaspag", "cod_for='" & txtbxformaspag.tag & "'")
    If lblParcelas.caption = "" Then
        lblParcelas.caption = "1"
    End If
    
    
End Sub

Private Sub cmndbtnBuscaFun_Click()
    Dim f As funcionario.funcionario
    On Error GoTo trata_Erro
    
    
    f = funcionario.F_SelectVendedor("")
    
    
    
    If f.cod_fun <> "" Then
        If F_senhaReturnUsu(True, False, f.guerra_fun, False) Then
        
        
            txtbxVendedor.Text = f.nome_fun
            txtbxVendedor.tag = f.cod_fun
            lblDesc.caption = F_BuscaUmCampo("desconto_usu", "tab_defusu", "cod_usu='" & f.cod_fun & "'")
            
        End If
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Private Sub cmndbtnCliente_Click()
    Dim c As representante.representante
    
    c = representante.F_BuscaRep(lblRepresentante.caption)
    If c.cod_rep <> "" Then
        txtbxRepresentante.tag = c.cod_rep
        txtbxRepresentante.Text = c.nome_rep
        lblLimite.caption = FormatVal(c.limite_rep)
        lblEmUso.caption = FormatVal(F_BuscaLimiteEmUsoRep(c.cod_rep))
        lblLimiteRest.caption = FormatVal(c.limite_rep - CCur(lblEmUso.caption))
        lblStatus_CLi.tag = c.status_rep
        txtbxCNPJ.Text = c.cpf_rep
        txtbxEndereco.Text = c.endereco_rep
        txtbxBairro.Text = c.bairro_rep
        txtbxCidade.Text = c.cidade_rep
        txtbxUf.Text = c.uf_rep
        txtbxFones.Text = c.telefone_rep
        txtbxRg.Text = c.rg_rep
        If c.status_rep = "1" Then
            lblStatus_CLi.caption = "Liberado"
            lblStatus_CLi.ForeColor = vbGreen
        ElseIf c.status_rep = "2" Then
            lblStatus_CLi.caption = "Cheques"
            lblStatus_CLi.ForeColor = vbYellow
        Else
            lblStatus_CLi.caption = "Bloqueado"
            lblStatus_CLi.ForeColor = vbRed
        End If
        If c.status_rep = "2" Then
            MsgBox "O Representante encontra-se negativado, portanto, só é possível venda a vista!", vbExclamation
            
            
        End If
        
        
        
        P_buscaitens
            
            
            
            
        
        
        pcLimite.Visible = True
        
        
        
        'carrega
        
    Else
        
    End If
     
            txtbxProduto.SetFocus
     
    
End Sub

Private Sub cmbxFormaPag_Change()

End Sub

Private Sub cmbxFormaPag_DropDown()

End Sub





Private Sub cmndbtnExclAv_Click()
    If MsgBox("Deseja excluir o lançamento?", vbYesNo + vbCritical) = vbYes Then
        If msflxgrdAv.rows = 2 Then
            msflxgrdAv.rows = 1
        Else
            msflxgrdAv.RemoveItem msflxgrdAv.Row
        End If
    End If
End Sub

Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    Dim strSQL As String
    P_CalcTotalAprazo
    f = F_SelectFormasPag("", "N")
    
    If f.cod_for <> "" Then
        
        txtbxFormaPag.Text = f.nome_for
        txtbxFormaPag.tag = f.cod_for
        
        
        lblParcelas.caption = f.parcelas_for
        lblTipoPag.caption = f.tipo_cai
        lblTipoPag.tag = f.cod_tip
        
        
        distribuiParcela
        somaDesconto
        
        
    End If
    
    
    
End Sub

Private Sub cmndbtnImprPos_Click()
    Venda.P_ImprPosologia Me.tag
End Sub

Private Sub cmbxFormaspag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        msflxgrd.SetFocus
    End If
    
End Sub

Private Sub cmbxOpc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOpc_LostFocus()
    msflxgrd.Text = cmbxOpc.Text
    If Left(msflxgrd.Text, 1) = "1" Then
        editaParcela
    Else
        editaCheque
    End If
    
    cmbxOpc.Visible = False
End Sub
Sub editaParcela()
    Dim p As parcela
    Dim sVen As String
    Dim sVal As String
    msflxgrd.Col = 0
    sVen = msflxgrd.Text
    msflxgrd.Col = 1
    sVal = msflxgrd.Text
    Exit Sub
    p = F_DadosPar(sVen, sVal)
    If p.valor > 0 Then
        If p.vencimento = "" Then
            MsgBox "O vencimento foi informado errado!", vbExclamation
            Exit Sub
        End If
        msflxgrd.Col = 0
        msflxgrd.Text = p.vencimento
        msflxgrd.Col = 1
        msflxgrd.Text = p.valor
                
        
    End If
End Sub
Sub editaCheque()
    Dim c As cheque
    With msflxgrd
        .Col = 0
        c.Venc_che = .Text
        .Col = 1: c.Valor_che = .Text
        .Col = 3: c.num_che = .Text
        .Col = 4: c.cod_bco = .Text
        .Col = 5: c.age_che = .Text
        .Col = 6: c.cco_che = .Text
        .Col = 7: c.nome_che = .Text
        .Col = 8: c.cpf_che = .Text
        
        If c.cpf_che = "" Then
            c.cpf_che = lblCpf.caption
            
        End If
        If c.nome_che = "" Then
            c.nome_che = txtbxClienteAcerto.Text
        End If
        c = F_DadosCheque(True, c.Venc_che, c.nome_che, c.num_che, c.cco_che, c.age_che, c.cod_bco, CStr(c.Valor_che), c.cpf_che)
        If c.num_che <> "" Then
        
            .Col = 1: .Text = c.Valor_che
            .Col = 0: .Text = c.Venc_che
            .Col = 3: .Text = c.num_che
            .Col = 4: .Text = c.cod_bco
            .Col = 5: .Text = c.age_che
            .Col = 6: .Text = c.cco_che
            .Col = 7: .Text = c.nome_che
            .Col = 8: .Text = c.cpf_che
            
        End If
        
        
    End With
            
            
            
End Sub
Private Sub cmndbtnBuscaRep_Click()
    Dim r As representante.representante
    r = F_BuscaRep(txtbxRepresentante.Text)
    
    If r.cod_rep <> "" Then
        
        
        If txtbxRepresentante.tag <> r.cod_rep Then
            msflxgrd.rows = 1
            msflxgrdAcerto.rows = 1
            msflxgrd.rows = 1
            P_SomaItd
        End If
        txtbxRepresentante.tag = r.cod_rep
        txtbxRepresentante.Text = r.nome_rep
        lblCpf.caption = r.cpf_rep
        Me.caption = "Acerto de: " & r.nome_rep
        txtbxClienteAcerto.SetFocus
    End If
End Sub

Private Sub cmndbtnCartao_Click()
    Dim p As pagamentocartao
    p = F_DadosCartao()
    If p.valor > "0" Then
        strSQL = Date & Chr(9)
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        If p.operacao = "C" Then
            strSQL = strSQL & "3 - Cartão de Crédito" & Chr(9) & Chr(9)
        Else
            strSQL = strSQL & "4 - Cartão de Débito" & Chr(9) & Chr(9)
        End If
        strSQL = strSQL & p.cod_fin & Chr(9)
        strSQL = strSQL & p.parcelas & Chr(9)
        
        
        msflxgrdAv.AddItem strSQL
        
    End If
    P_SomaItd
End Sub

Private Sub cmndbtnCh_Click()
    Dim c As cheque
    c = F_DadosCheque(True, , txtbxClienteAcerto.Text, "", "", "", "", "", lblCpf.caption)
    
    If c.Valor_che > 0 Then
        
        strSQL = c.Venc_che & Chr(9)
        strSQL = strSQL & FormatVal(c.Valor_che) & Chr(9)
        strSQL = strSQL & "2 - Cheque-pré" & Chr(9)
        strSQL = strSQL & c.num_che & Chr(9)
        strSQL = strSQL & c.cod_bco & Chr(9)
        strSQL = strSQL & c.age_che & Chr(9)
        strSQL = strSQL & c.cco_che & Chr(9)
        strSQL = strSQL & FormatVal(c.Valor_che) & Chr(9)
        strSQL = strSQL & c.Venc_che & Chr(9)
        strSQL = strSQL & c.nome_che & Chr(9)
        strSQL = strSQL & c.cpf_che
        msflxgrdAv.AddItem strSQL
    End If
    P_SomaItd
End Sub

Private Sub cmndbtnClienteAcerto_Click()
    Dim c As Clientes.Cliente
    If optnRep.Value = True Then
        MsgBox "Você optou por fazer acerto para o próprio representante, não é possível escolher qualquer cliente!", vbExclamation
        optnRep.SetFocus
        Exit Sub
    End If
    c = Clientes.F_BuscaCliente(txtbxClienteAcerto.Text)
    If c.cod_cli <> "" Then
        txtbxClienteAcerto.Text = c.nome_Cli
        txtbxClienteAcerto.tag = c.cod_cli
        lblCpfCli.caption = c.cnpj_Cli
        lblRgCli.caption = c.ie_Cli
        lblEnderecoCli.caption = StrConv(c.endereco_cli, vbProperCase)
        lblBairroCli.caption = StrConv(c.bairro_cli, vbProperCase)
        lblCidadeCli.caption = StrConv(c.cidade_cli, vbProperCase)
        txtbxProdAcerto.SetFocus
    End If
End Sub

Private Sub cmndbtnCob_Click()
    Dim t As tipocob
    t = F_BuscaTipoCobRep(txtbxRepresentante.tag)
    If t.cod_tip <> "" Then
        txtbxTipoCob.tag = t.cod_tip
        txtbxTipoCob.Text = t.nome_tip
        txtbxDescCob.Text = t.comiss_tip
        somaDesconto
    End If
    distribuiParcela
    txtbxVenIniAcerto.SetFocus
End Sub
Sub somaDesconto()
    Dim cCom As Currency
    cCom = CCur(txtbxDescCob.Text) / 100 * CCur(lblTotalInicial.caption)
    cCom = cCom + CCur(txtbxDifAcerto.Text)
    lblComissao.caption = FormatVal(cCom)
    
    If optnRep.Value Then
        'lblCapComiss.caption = "Desconto"
        lblValorFinal.caption = FormatVal(CCur(lblTotalAcerto.caption) - CCur(lblComissao.caption))
    Else
        'lblCapComiss.caption = "Comissão"
        lblValorFinal.caption = lblTotalAcerto.caption
    End If
    
End Sub
Private Sub ConfAprazo()
    
    If lblTotalDistribuido.caption <> lblValorFinal.caption Then
        MsgBox "O Valor distribuído é de " & lblTotalDistribuido.caption _
        & ". No entanto o Deve ser " & lblValorFinal.caption, vbExclamation
        Exit Sub
    End If
    If msflxgrdAcerto.rows <= 1 Then
        MsgBox "Não há o que lançar!", vbExclamation
        Exit Sub
    End If
    If txtbxClienteAcerto.tag = "" Then
        MsgBox "Informe o cliente do acerto!", vbExclamation
        txtbxClienteAcerto.SetFocus
        Exit Sub
    End If
    
    If txtbxRepresentante.tag = "" Then
        MsgBox "Não há um representante escolhido!", vbExclamation
        txtbxRepresentante.SetFocus
        Exit Sub
    End If
    If txtbxFormaPag.tag = "" Then
        MsgBox "Escolha a forma de pagamento!", vbExclamation
        Exit Sub
    End If
    
    If txtbxVenIniAcerto.Text = "" Then
        MsgBox "Informe a data para o vencimento inicial!", vbExclamation
        txtbxVenIniAcerto.SetFocus
        Exit Sub
    End If
    If txtbxTipoCob.tag = "" Or txtbxTipoCob.tag = "0" Then
        MsgBox "Escolha o tipo de cobrança!", vbExclamation
        cmndbtnCob_Click
    End If
    If MsgBox("Confirma lançamento de acerto para o cliente " & txtbxClienteAcerto.Text & "?", vbYesNo + vbQuestion) = vbNo Then
        
        Exit Sub
    End If
    
        
    
    lancaAcerto
    
End Sub
Private Sub ConfAvista()
    
    If lblTotalDistribuido.caption <> lblValorFinal.caption Then
        MsgBox "O Valor distribuído é de " & lblTotalDistribuido.caption _
        & ". No entanto o Deve ser " & lblValorFinal.caption, vbExclamation
        Exit Sub
    End If
    If optnCli.Value = True Then
        MsgBox "Pagamento À Vista apenas para o Representante!", vbExclamation
        optnCli.SetFocus
        Exit Sub
        
    End If
    If msflxgrdAcerto.rows <= 1 Then
        MsgBox "Não há o que lançar!", vbExclamation
        Exit Sub
    End If
    If txtbxClienteAcerto.tag = "" Then
        MsgBox "Informe o cliente do acerto!", vbExclamation
        txtbxClienteAcerto.SetFocus
        Exit Sub
    End If
    
    If txtbxRepresentante.tag = "" Then
        MsgBox "Não há um representante escolhido!", vbExclamation
        
        Exit Sub
    End If
    If msflxgrdAv.rows = 1 Then
        MsgBox "Não há um pagamento adicionado!", vbExclamation
        cmndbtnDin.SetFocus
        Exit Sub
        
    End If
    
    If txtbxVenIniAcerto.Text = "" Then
        MsgBox "Informe a data para o vencimento inicial!", vbExclamation
        txtbxVenIniAcerto.SetFocus
        Exit Sub
    End If
'    If txtbxTipoCob.tag = "" Then
'        MsgBox "Escolha o tipo de cobrança!", vbExclamation
'        cmndbtnCob_Click
'    End If
    If MsgBox("Confirma lançamento de acerto para o cliente " & txtbxClienteAcerto.Text & "?", vbYesNo + vbQuestion) = vbNo Then
        
        Exit Sub
    End If
    
        
    P_Status "Finalizando acerto..."
    
    lancaAcerto
    
    P_Status ""
    
End Sub
Sub lancaAcerto()
    Dim strSQL As String
    Dim iCount As Integer
    Dim sSeq_Doc As String
    Dim sSEQ_ITD As String
    Dim sSEQ As String
    Dim cAux As Currency
    Dim seq_con As String
    Dim unit_ite As String
    Dim qtd_ite As String
    Dim cod_pro As String
    Dim dif_ite As String
    Dim cValorIte As Currency
    Dim cDif As Currency
    Dim sVen As String
    Dim sSeq_Com As String
    Dim sTipoCob As String
'    Dim sDesconto As String
    
    Dim sVal As String, sEsp As String, sNum_Che As String, sBco As String
    Dim sAge As String, sCCO As String, sNome_che As String
    Dim sCpf As String
    Dim sBol As String
    If Not F_confMsflx(msflxgrd) Then Exit Sub
    If Not F_confMsflx(msflxgrdAv) Then Exit Sub
    
    
    If txtbxTipoCob.tag = "" Then
        If sstb.Tab = 0 Then
            MsgBox "Escolha a conta cobrança!", vbExclamation
            cmndbtnCob_Click
            Exit Sub
        Else
            txtbxTipoCob.tag = "0"
        End If
    End If
    
    'If Not F_IsCaixa Then
    '    MsgBox "Este computador não pode fazer acerto, pois não é um Computador Caixa!", vbExclamation
    '    Exit Sub
    'End If
    
    P_Status "Gravando Acerto..."
    sSEQ = F_GravaAcerto(txtbxRepresentante.tag, CStr(txtbxDataBase.Text), CStr(Time), _
    txtbxClienteAcerto.tag, txtbxFormaPag.tag, lblValorFinal.caption, _
    txtbxTipoCob.tag, lblCli.tag, lblComissao.caption, txtbxDifAcerto.Text, txtbxEntrada.Text)
            
    'grava o acerto e guarda o numero em sSEQ pra poder usar no resto
            
    If sSEQ = "" Then
        MsgBox "Erro na gravação do duplicata!", vbExclamation
        Exit Sub
    Else
        With msflxgrdAcerto
            For iCount = 1 To .rows - 1
                .Row = iCount
                .Col = 0: cod_pro = .Text
                .Col = 3: qtd_ite = .Text
                .Col = 5: unit_ite = .Text
                .Col = 7: dif_ite = .Text
                .Col = 9: seq_con = .Text
                
                sSEQ_ITD = F_GravaItAcerto(sSEQ, seq_con, cod_pro, unit_ite, qtd_ite, dif_ite)
                
                
                
                If sSEQ_ITD = "" Then
                    MsgBox "Erro na gravação do item do acerto!", vbExclamation
                    strSQL = "delete from tab_acertos where seq_ace='" & sSEQ & "'"
                    Banco.Execute strSQL
                    strSQL = "delete from tab_itacertos where seq_ace='" & sSEQ & "'"
                    Banco.Execute strSQL
                    
                    Exit Sub
                Else
                    'If Produtos.F_AlteraEstRes(cod_pro, "-" & qtd_ite) Then
                     'já faz na gravaestcons
                    'End If
                    
                    If Not F_GravaEstCons(txtbxRepresentante.tag, cod_pro, "-" & qtd_ite) Then
                        MsgBox "Ocorreu um erro na volta do estoque, informe o responsável!", vbExclamation
                    End If
                    
                    If Not Produtos.F_AlteraEstoquePro(cod_pro, "-" & qtd_ite) Then
                        MsgBox "Erro na baixa do estoque", vbExclamation
                    End If
                    strSQL = "update tab_consignacoes set reman_con=reman_con-" & F_GravaDec(qtd_ite)
                    strSQL = strSQL & " where seq_con='" & seq_con & "'"
                    strSQL = strSQL & " and cod_pro='" & cod_pro & "'"
                    Banco.Execute strSQL
                    
                    'lembrar de fazer o cancelamento e voltar caso não de certo em algum ponto
                    
                    
                    
                End If
            Next
        End With
    End If
    
    
    
    If sstb.Tab = 0 Then
        If optnCli.Value Then
            sSeq_Doc = F_GravaDoc("ACERTO CONSIGN. " & txtbxClienteAcerto.Text, sSEQ, txtbxTotalAprazo.Text, txtbxDataBase.Text, _
                txtbxClienteAcerto.tag, "R", sSEQ, lblParcelas.caption, "A", lblCli.tag, "002")
        Else
    
            sSeq_Doc = F_GravaDoc("ACERTO CONSIGN. " & txtbxRepresentante.Text, sSEQ, txtbxTotalAprazo.Text, txtbxDataBase.Text, _
                txtbxRepresentante.tag, "R", sSEQ, lblParcelas.caption, "A", lblCli.tag, "002")
        End If
        
        
        
        If sSeq_Doc <> "" Then
            With msflxgrd
                For iCount = 1 To .rows - 1
                    .Row = iCount
                    msflxgrd.Col = 0
                    sVen = msflxgrd.Text
                    msflxgrd.Col = 1
                    sVal = msflxgrd.Text
                    msflxgrd.Col = 2
                    sEsp = Left(msflxgrd.Text, 1)
                    msflxgrd.Col = 3
                    sNum_Che = msflxgrd.Text
                    msflxgrd.Col = 4
                    sBco = msflxgrd.Text
                    msflxgrd.Col = 5
                    sAge = msflxgrd.Text
                    msflxgrd.Col = 6
                    sCCO = msflxgrd.Text
                    msflxgrd.Col = 7
                    sNome_che = msflxgrd.Text
                    msflxgrd.Col = 8
                    sCpf = msflxgrd.Text
                    msflxgrd.Col = 9
                    sBol = msflxgrd.Text
                    If sEsp = "1" Then
                        sSEQ_ITD = doc.F_GravaItDoc(sSeq_Doc, sVal, sVen, sVen, CStr(iCount), acerto_consignado, sBol, txtbxTipoCob.tag)
                        strSQL = F_GravaCai("C", txtbxDataBase.Text, sVal, "ACERTO CONSIGN. " & txtbxRepresentante.Text, sSEQ, _
                        "DUP", "", 0, 0, strCod_Cai, strCod_Tur, "000")
                    Else
                        
                        sSEQ_ITD = F_GravaCh(sSeq_Doc, sNum_Che, sBco, sAge, sCCO, _
                            sNome_che, sVal, sVen, txtbxClienteAcerto.tag, lblCli.tag, _
                            "", "", CStr(iCount), "", sCpf)
                        
                        'sSEQ_ITD = doc.F_GravaItDoc(sSeq_Doc, sVal, sVen, sVen, CStr(iCount), acerto_consignado, sBol, txtbxTipoCob.tag)
                        
                        strSQL = F_GravaCai("C", txtbxDataBase.Text, sVal, "ACERTO CONSIGN. " & txtbxRepresentante.Text, sSEQ, _
                        "CHP", "", 0, 0, strCod_Cai, strCod_Tur, "000")
                        
                    End If
                    
                    If optnCli.Value Then
                        If Not F_GravaComissParc(sSEQ_ITD, txtbxRepresentante.tag, sSEQ, _
                            CCur(lblComissao.caption) / (.rows - 1), sVen, sEsp) Then
                            MsgBox "Houve um problema na gravação da comissão, informe o responsável!", vbExclamation
                            
                        End If
             
                    End If
                    
                    
                    If sSEQ_ITD = "" Then
                        MsgBox "Houve um erro na gravação da parcela!", vbExclamation
                        strSQL = "delete from tab_acertos where seq_ace='" & sSEQ & "'"
                        Banco.Execute strSQL
                        strSQL = "delete from tab_itacertos where seq_ace='" & sSEQ & "'"
                        Banco.Execute strSQL
                        strSQL = "delete from tab_doc where seq_doc='" & sSeq_Doc & "'"
                        Banco.Execute strSQL
                        strSQL = "delete from tab_itdoc where seq_doc='" & sSeq_Doc & "'"
                        Banco.Execute strSQL
                        Exit Sub
                    End If
                    
                    
                    
                Next
            End With
        End If
        
        
        
        
        If optnCli.Value Then
            sSeq_Com = Acerto.F_GravaComissProv(txtbxRepresentante.tag, sSEQ, lblComissao.caption, sSeq_Doc)
             
            If sSeq_Com = "" Then
                MsgBox "Houve um erro na gravação da comissão temporária para o representante!", vbExclamation
                'Exit Sub
            End If
        End If
        
    Else
        'é pagamento a vista
        With msflxgrdAv
            For iCount = 1 To .rows - 1
                .Row = iCount
                .Col = 0
                sVen = .Text
                .Col = 1
                sVal = .Text
                .Col = 2
                sEsp = Left(.Text, 1)
                .Col = 3
                sNum_Che = .Text
                .Col = 4
                sBco = .Text
                .Col = 5
                sAge = .Text
                .Col = 6
                sCCO = .Text
                .Col = 7
                sNome_che = .Text
                .Col = 8
                sCpf = .Text
                .Col = 9
                sBol = .Text
                
                
                
                
                If sEsp = "1" Then
                    .Col = 1
                    sVal = .Text
                    .Col = 2
                    If Left(.Text, 1) = "1" Then
                        sEsp = "DIN"
                    ElseIf Left(.Text, 1) = "2" Then
                        sEsp = "CHV"
                    Else
                        sEsp = "CRT"
                    End If
                    sSeq_Doc = F_GravaCai("C", txtbxDataBase.Text, sVal, "ACERTO CONS. " & txtbxRepresentante.Text, sSEQ, _
                     sEsp, "", 0, 0, strCod_Cai, strCod_Tur, "002")
                    
                    If sSeq_Doc <> "" Then
                        strSQL = "insert into tab_acertofin(seq_fin,seq_ace,tipo_fin)values("
                        strSQL = strSQL & "'" & sSeq_Doc & "',"
                        strSQL = strSQL & "'" & sSEQ & "', "
                        strSQL = strSQL & "'C')"
                        Banco.Execute strSQL
                    End If
                    
                ElseIf sEsp = "2" Then
                    sSeq_Doc = F_GravaDoc("ACERTO CONSIGN. " & txtbxClienteAcerto.Text, sSEQ, sVal, txtbxDataBase.Text, _
                    txtbxRepresentante.tag, "R", sSEQ, lblParcelas.caption, "A", lblCli.tag, "002")
                    strSQL = F_GravaCh(sSeq_Doc, sNum_Che, sBco, sAge, sCCO, sNome_che, sVal, sVen, txtbxRepresentante.tag, _
                    "R", sSEQ, "", "1", "", sCpf)
                    
                    
                    
                    'imprimeCarne sSeq_Doc
                    sSeq_Doc = F_GravaCai("C", txtbxDataBase.Text, sVal, "ACERTO CONS. " & txtbxRepresentante.Text, sSEQ, sEsp, "", 0, 0, strCod_Cai, strCod_Tur, "002")
                ElseIf sEsp = "3" Or sEsp = "4" Then
                    If sAge = "" Then sAge = "1"
                    If sEsp = "3" Then
                        sEsp = "CRT"
                    Else
                        sEsp = "CRD"
                    End If
                    strSQL = F_GravaVenFin(sSEQ, sBco, CStr(sVal), txtbxDataBase.Text, "", sAge, "A")
                    sSeq_Doc = F_GravaCai("C", txtbxDataBase.Text, sVal, "ACERTO CONS. " & txtbxRepresentante.Text, sSEQ, sEsp, "", 0, 0, strCod_Cai, strCod_Tur, "002")
                End If
            
            Next
            
        End With
        If CCur(txtbxEntrada.Text) > 0 Then
            strSQL = F_GravaCai("C", txtbxDataBase.Text, txtbxEntrada.Text, "ENTRADA ACERTO CONS " & txtbxRepresentante.Text, "", "DIN", strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, "000")
        End If
    End If
    imprAcertoData txtbxRepresentante.tag, txtbxDataBase.Text
    If sSeq_Doc <> "" Then
        imprimeCarne sSeq_Doc
    End If
    txtbxClienteAcerto.tag = ""
    txtbxClienteAcerto.Text = ""
    txtbxClienteAcerto.Text = ""
    
    msflxgrdAcerto.rows = 1
    msflxgrd.rows = 1
    msflxgrdAv.rows = 1
    
    txtbxValorNegociado.Text = "0,00"
    txtbxDifAcerto.Text = "0,00"
    lblValorFinal.caption = "0,00"
    lblTotalDistribuido.caption = "0,00"
    lblTotalAcerto.caption = "0,00"
    lblVlrAcertoDiv.caption = "0,00"
    txtbxValorNegociado.Text = "0,00"
    txtbxDifAcerto.Text = "0,00"
    lblTotalInicial.caption = "0,00"
    
    txtbxClienteAcerto.SetFocus
End Sub
Sub lanDoc()
    Dim sSeq_Doc As String
    Dim sSEQ_ITD As String
    Dim sSeq_che As String
    Dim iCount As Integer
    Dim sVen As String, sVal As String, sEsp As String, sNum_Che As String, sBco As String
    Dim sAge As String, sCCO As String, sNome_che As String
    Dim sCpf As String
    Dim sBol As String
    
    If msflxgrd.rows = 1 Then
        MsgBox "Os valores não estão distribuídos!", vbExclamation
        Exit Sub
    End If
    
    If F_Soma <> CCur(txtbxValor_Doc.Text) Then
        MsgBox "Os valores não estão corretamente distribuídos!", vbExclamation
        Exit Sub
    ElseIf F_Vazio(txtbxNum_Doc.Text) Then
        MsgBox "Informe o número do Documento!", vbExclamation
        txtbxNum_Doc.SetFocus
        Exit Sub
    End If
    sSeq_Doc = doc.F_GravaDoc(txtbxHist.Text, txtbxNum_Doc.Text, txtbxValor_Doc.Text, _
    CStr(txtbxRef.Text), txtbxNome.tag, Left(lblTipo.caption, 1), _
    "", msflxgrd.rows - 1, "F", lblCli.tag, "002")
    If sSeq_Doc <> "" Then
        With msflxgrd
            For iCount = 1 To .rows - 1
                .Row = iCount
                msflxgrd.Col = 0
                sVen = msflxgrd.Text
                msflxgrd.Col = 1
                sVal = msflxgrd.Text
                msflxgrd.Col = 2
                sEsp = Left(msflxgrd.Text, 1)
                msflxgrd.Col = 3
                sNum_Che = msflxgrd.Text
                msflxgrd.Col = 4
                sBco = msflxgrd.Text
                msflxgrd.Col = 5
                sAge = msflxgrd.Text
                msflxgrd.Col = 6
                sCCO = msflxgrd.Text
                msflxgrd.Col = 9
                sNome_che = msflxgrd.Text
                msflxgrd.Col = 10
                sCpf = msflxgrd.Text
                msflxgrd.Col = 11
                sBol = msflxgrd.Text
                If sEsp = "1" Then
                    sSEQ_ITD = doc.F_GravaItDoc(sSeq_Doc, sVal, sVen, sVen, CStr(iCount), "C", sBol)
                ElseIf sEsp = "2" Then
                    If optnReceb.Value = True Then 'cheque a receber
                        sSeq_che = F_GravaCh(sSeq_Doc, sNum_Che, sBco, sAge, sCCO, _
                        sNome_che, sVal, sVen, txtbxNome.tag, Left(lblTipo.caption, 1), _
                        "", "", CStr(iCount), "", sCpf)
                    Else 'cheque a pagar
                        sSeq_che = F_GravaChePag(sSeq_Doc, sNum_Che, sBco, sVal, sVen, _
                        txtbxNome.tag, CStr(iCount))
                    End If
                End If
                
            Next
        End With
    End If
End Sub
Private Sub cmndbtnLocaliza_Click()
    
    If F_Vazio(txtbxProduto.Text) Then
'        MsgBox "Informe o produto que deseja inserir!", vbExclamation, constNome_sis
'        txtbxProduto.SetFocus
'        Exit Sub
    End If
    
    buscaProduto
End Sub

Private Sub cmndbtnConfAcerto_Click()
    If sstb.Tab = 0 Then
        
        ConfAprazo
    Else
        
        ConfAvista
    End If
    
End Sub

Private Sub cmndbtnDin_Click()
    Dim p As parcela
    p = F_DadosPar(Date)
    If p.valor > 0 Then
        strSQL = Date & Chr(9)
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        strSQL = strSQL & "1 - Dinheiro" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9)
        
        strSQL = strSQL & FormatVal(p.valor) & Chr(9)
        strSQL = strSQL & Date
        msflxgrdAv.AddItem strSQL
        
        
    End If
    P_SomaItd
End Sub

Private Sub cmndbtnEditAp_Click()
    If msflxgrd.Row <= 0 Then
        Exit Sub
    End If
    If Left(msflxgrd.Text, 1) = "1" Then
        editaParcela
    Else
        editaCheque
    End If
End Sub

Sub deletaav()
    If msflxgrdAv.rows = 2 Then
            msflxgrdAv.rows = 1
            
        ElseIf msflxgrdAv.rows > 2 Then
            msflxgrdAv.RemoveItem msflxgrdAv.Row
        End If
        
        P_SomaItd
End Sub
Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    
    Unload Me
    Exit Sub
trata_Erro:
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
    If strSQL <> txtbxVendedor.tag Then
        txtbxVendedor.tag = strSQL
        txtbxVendedor.Text = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & strSQL & "'")
        Venda.P_AtualizaVendaCliente Me.tag, txtbxRepresentante.tag, lblRepresentante.caption, txtbxConveniado.tag
        
    End If
    
    
    
        If lblRepresentante.caption = "" Then
            strSQL = InputBox("Informe um nome para o cliente", "Nome do Cliente", "Venda a consumidor")
            If strSQL = "" Then
                MsgBox "Informe um nome para o cliente!", vbExclamation
                'lblRepresentante.SetFocus
                Exit Sub
            Else
                'lblRepresentante.Text = strSQL
                Venda.P_AtualizaVendaCliente Me.tag, txtbxRepresentante.tag, lblRepresentante.caption, ""
            End If
            
            
            
        End If
        
        If Vendas.F_AlteraStatusVenda(Me.tag, "2") Then
            Me.tag = F_GeraNovaVenda
            bActive = False
            Form_Activate
        End If
    
    
    
End Sub



Private Sub dbgrdProd_DblClick()
    cmndbtnEditProd_Click
End Sub

Private Sub dbgrdProd_GotFocus()
    
    dbgrdProd.MarqueeStyle = 3
    
    
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
    On Error Resume Next
    If Not .EOF Then
        lblProduto.caption = !nome_pro
            
            
                lblCapPreco.caption = "Preço:"
                lblPreco.caption = "R$ " & FormatVal(!pv_pro)
            
            lblCat.caption = !nome_cat
            lblCod.caption = FormatCodigo(!cod_pro)
            lblCodBarr.caption = F_BuscaUmCampo("codbarra_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'")
            lblMarca.caption = !nome_mar
            lblQtd.caption = !qtd_ite
            lblRef.caption = !codfab_pro
            lblBalanca.caption = !balanca_pro
            lblPosologia.caption = F_BuscaUmCampo("descricao_pos", "tab_posologia", "num_ven='" & Me.tag & _
                "' and seq_ite=" & dtProd.Recordset!seq_ite)
            
            
        End If
    End With
        
End Sub



Private Sub Form_Activate()
    On Error GoTo trata_Erro
    
    
    'lblnum_ven.caption = Me.tag
    
    
    lblCaixa = strCod_Cai
    
    If bActive = True Then Exit Sub
    
    
    If txtbxRepresentante.tag = "" Then
        cmndbtnBuscaRep_Click
    End If
    
    P_Status "Buscando Itens..."
    
    P_buscaitens
    
    P_Status ""
    
    bActive = True
    If strCod_Cai = "" Then
        MsgBox "Não existe caixa para este computador, o caixa ficará para o '01' ", vbExclamation
        strCod_Cai = "00"
    End If
    
    Exit Sub
trata_Erro:
    P_MostraErro "Form_Activate"
    
End Sub

    


Sub abreDevolucao()
    
    frmDevolucaoConsign.txtbxRepresentante.Text = txtbxRepresentante.Text
    frmDevolucaoConsign.txtbxRepresentante.tag = txtbxRepresentante.tag
    frmDevolucaoConsign.Show 1
    
    
    
    
    
    
End Sub
Sub mudamodo()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set modo_cai='" & operacao & "' "
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
    Banco.Execute strSQL
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Me.ActiveControl.Name = msflxgrdAcerto.Name Then
            excluiItemAcerto
        End If
    
    End If
    
    If KeyCode = vbKeyF2 Then
        txtbxProdAcerto.SetFocus
    ElseIf KeyCode = vbKeyF3 Then
    ElseIf KeyCode = vbKeyF4 Then
        txtbxClienteAcerto.SetFocus
    ElseIf KeyCode = vbKeyF7 Then
        cmndbtnConfAcerto_Click
    ElseIf KeyCode = vbKeyF8 Then
        sstb.Tab = 0
        txtbxTipoCob.SetFocus
    ElseIf KeyCode = vbKeyF9 Then
        sstb.Tab = 1
    ElseIf KeyCode = vbKeyF12 Then
        cmndbtnSair_Click
    End If
    
End Sub

Private Sub Form_Load()
   
    bActive = False
    sstb.Tab = 0
    resizeForm Me
    ajustaGrid
    
    
   txtbxDataBase = Date
    
    
    'P_PreencheFormasPag cmbxFormaspag, "02"
    
    
'    cmbxBuscar.ListIndex = 0
    
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
    
    
    
    ajustaGridAcerto
End Sub

Function F_QtdProdutoLancadoAcerto(cod_pro As String) As Integer
    Dim i As Integer
    Dim cTotal As Integer
    Dim sSeqCon As String
    With msflxgrdAcerto
        
        For i = 1 To .rows - 1
            .Row = i
            .Col = 9
            sSeqCon = .Text
            .Col = 0
            
            If .Text = cod_pro Then
                .Col = 3
                cTotal = cTotal + CCur(.Text)
            End If
            
        Next
    End With
    F_QtdProdutoLancadoAcerto = cTotal
End Function
Function F_QtdProdutoLancadoAcerto1(cod_pro As String, seq_con As String) As Integer
    Dim i As Integer
    Dim cTotal As Integer
    Dim sSeqCon As String
    With msflxgrdAcerto
        
        For i = 1 To .rows - 1
            .Row = i
            .Col = 9
            sSeqCon = .Text
            .Col = 0
            
            If .Text = cod_pro And seq_con = sSeqCon Then
                .Col = 3
                cTotal = cTotal + CCur(.Text)
            End If
            
        Next
    End With
    F_QtdProdutoLancadoAcerto1 = cTotal
End Function

Sub ajustaGridAcerto()
    With msflxgrdAcerto
        .rows = 1
        .Cols = 10
        .Col = 0: .Text = "Código"
        .Col = 1: .Text = "Referência"
        .Col = 2: .Text = "Produto"
        .Col = 3: .Text = "Qtd"
        .Col = 4: .Text = "Unit. Inicial"
        .Col = 5: .Text = "Unit Vendido"
        .Col = 6: .Text = "Total"
        .Col = 7: .Text = "Diferença"
        .Col = 8: .Text = "Transferido Em:"
        .Col = 9: .Text = "Seq.Con:"
        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 3500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 2500
    End With
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
    On Error Resume Next
    
    
    
    
    
    msflxgrdAcerto.Height = pcBot.Top - msflxgrdAcerto.Top - 100
    msflxgrdAcerto.Width = pcTotalItens.Left - msflxgrdAcerto.Left - 100
    frCli.Width = pcCli.Width
    
    
    
    
    
End Sub

Sub editaItem()
    txtbxQtd.Enabled = True
    txtbxQtd.Text = "1"
    txtbxQtd.SetFocus
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
    On Error Resume Next
    Timer1.Enabled = False
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




Private Sub msflxgrd_KeyPress(KeyAscii As Integer)
    editaCelula KeyAscii
End Sub

Private Sub msflxgrdAv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        deletaav
    End If
End Sub

Private Sub optnCli_Click()
     txtbxClienteAcerto.Text = ""
    txtbxClienteAcerto.tag = ""
    lblCli.tag = "C"
    lblCli.caption = "C&liente (F4):"
End Sub

Private Sub lblDelExcluir_Click()
    excluiItemAcerto
End Sub
Sub excluiItemAcerto()
    With msflxgrdAcerto
        If .Col <= 0 Then
            MsgBox "Selecione um item para excluir", vbExclamation
            Exit Sub
        End If
        .Col = 2
        If MsgBox("Deseja excluir o item " & .Text & "?", vbYesNo + vbCritical) = vbYes Then
            If .rows = 2 Then
                .rows = 1
            Else
                .RemoveItem .Row
            End If
        End If
    End With
    P_SomaAcerto
End Sub
Sub P_SomaAcerto()
    Dim cTotal As Currency
    Dim iCount As Integer
    Dim cDif As Currency
    Dim cIni As Currency
    Dim cQtd As Currency
    With msflxgrdAcerto
    
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 3
            cQtd = CCur(.Text)
            .Col = 6
            cTotal = cTotal + CCur(.Text)
            .Col = 7
            cDif = cDif + CCur(.Text)
            .Col = 4
            cIni = cIni + (CCur(.Text) * cQtd)
        Next
    End With
    
    lblTotalAcerto.caption = FormatVal(cTotal)
    'lblTotalDifAcert0.caption = FormatVal(cDif)
    lblTotalInicial.caption = FormatVal(cIni)
End Sub
Private Sub pcBotao_Click(Index As Integer)
    On Error GoTo trata_Erro
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
        editaItem
    ElseIf Index = 2 Then
        
        lancaPosologia
    ElseIf Index = 0 Then
'        MsgBox "Não há TEF"
    ElseIf Index = 7 Then
        
    ElseIf Index = 8 Then
    ElseIf Index = 9 Then
    ElseIf Index = 10 Then
        ImprConsig dtTrans.Recordset!seq_con
    ElseIf Index = 11 Then
        concluiSaidas
    ElseIf Index = 12 Then
    
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub concluiSaidas()
    Dim strSQL As String
    Dim iCount As Integer
    Dim sSEQ As String
    If dtSai.Recordset.RecordCount = 0 Then
        MsgBox "Não existem transferências para serem concluídas!", vbExclamation
        Exit Sub
        
    End If
    
    'If MsgBox("Confirma conclusão dos lançamentos para transferência?", vbYesNo + vbQuestion) = vbNo Then
    '    Exit Sub
    'End If
    
    If F_senhaReturnUsu() = "" Then
        Exit Sub
    End If
    
    sSEQ = MaxKeySQL("tab_consignacoes", "seq_con", "", "000000#")  'buscando o sequencial q servirá para todos os produtos da transf
    With dtSai.Recordset
        Do While Not .EOF
            If Consignado.F_GravaConsign(sSEQ, !cod_pro, !qtd_con, !cod_rep, !unit_con) = "" Then
                MsgBox "Erro na transferência no item " & .AbsolutePosition, vbCritical
                
            Else
                If Not Consignado.F_DeletaConsignAberta(!seq_con) Then
                
                End If
            End If
            
            .MoveNext
        Loop
        
    End With
    P_buscaitens
    
End Sub
Sub lancaPosologia()
    If dtProd.Recordset.RecordCount = 0 Then
        MsgBox "Não existe itens na sua venda!", vbExclamation
        Exit Sub
    End If
    If dbgrdProd.Row = -1 Then
        MsgBox "Selecione o Item para lançar a posologia!", vbExclamation
        dbgrdProd.SetFocus
        Exit Sub
    End If
    Venda.P_AddPosologia Me.tag, dtProd.Recordset!seq_ite
    
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    
iGif = iGif + 1
If iGif >= 15 Then
    If iIndPromo >= QtdPromo - 1 Then
        
        iIndPromo = 0
    End If
    lblPromo.caption = QuebraLinha(StrConv(pPromo(iIndPromo).Nome, vbProperCase), 30)
    lblPromoPreco.caption = pPromo(iIndPromo).Preco
    iIndPromo = iIndPromo + 1
    iGif = 0
    
End If
    lblBloqueado.caption = F_BuscaUmCampo("bloqueado_cai", "tab_caixas", "cod_cai='" & strCod_Cai & "'")
    
    If lblBloqueado.caption = "S" Then
        If dtProd.Recordset.RecordCount = 0 Then
            
        End If
    End If
    
    
    iDescancoTela = iDescancoTela + 1
    
    If iDescancoTela = 100 Then
        On Error Resume Next
'        frmPropa.Show 0
'        frmPropa.SetFocus
        
        iDescancoTela = 0
    End If
    
    lblHora.caption = Now
    lblHora.Refresh
   End Sub

Sub buscaCodCli()
    Dim strSQL As String
    
    strSQL = F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & _
        txtbxCod_Cli.Text & "'")
    
    If strSQL = "" Then
        MsgBox "Cliente não encontrado!", vbExclamation
        txtbxCod_Cli.Text = ""
        txtbxCod_Cli.SetFocus
    Else
        txtbxRepresentante.Text = strSQL
        txtbxRepresentante.tag = txtbxCod_Cli.Text

    End If
    
    
End Sub


Private Sub pcProd_GotFocus()
    SendKeys Chr(9)
End Sub



Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'editaCelula
    End If
End Sub
Sub editaCelula(KeyAscii As Integer)
    Dim obj As Object
    
    
    
    
    
    
        If msflxgrd.Col = 2 Then
        Set obj = cmbxOpc
    Else
        Set obj = txtbxMsf
    End If
    
    
    
    On Error Resume Next
    
    
    obj.Visible = True
    
    obj.Text = msflxgrd.Text
    obj.SetFocus
    obj.Width = msflxgrd.CellWidth
    obj.Height = msflxgrd.CellHeight
    obj.Top = msflxgrd.CellTop + msflxgrd.Top
    obj.Left = msflxgrd.CellLeft + msflxgrd.Left
    If KeyAscii <> 13 Then
        'If .obj.Text = "" Then
            obj.Text = Chr(KeyAscii)
        'End If
    End If
    
    'obj.Alignment = msflxgrd.CellAlignment
    
End Sub

Private Sub optnRep_Click()
    txtbxClienteAcerto.Text = txtbxRepresentante.Text
    txtbxClienteAcerto.tag = txtbxRepresentante.tag
    lblCli.tag = "R"
    lblCli.caption = "&Representante"
    lblCpfCli.caption = lblCpf.caption
    txtbxProdAcerto.SetFocus
End Sub


Private Sub pcProdAcerto_GotFocus()
    SendKeys Chr(9)
End Sub

Private Sub sstb_Click(PreviousTab As Integer)
    If sstb.Tab = 0 Then
        pcTotal.Visible = True
        msflxgrdAv.rows = 1
        txtbxDescCob.Text = "0"
        txtbxTipoCob.Text = ""
        
        P_SomaItd
        somaDesconto
        lblCapComiss.caption = "Comissão"
    ElseIf sstb.Tab = 1 Then
        pcTotal.Visible = False
        msflxgrd.rows = 1
        txtbxEntrada.Text = "0,00"
        txtbxTipoCob.tag = "0"
        txtbxDescCob.Text = txtbxDescontoAvista.Text
        
        P_SomaItd
        somaDesconto
        lblCapComiss.caption = "Desconto"
    ElseIf sstb.Tab = 2 Then
    ElseIf sstb.Tab = 3 Then
        
    ElseIf sstb.Tab = 4 Then
    End If
End Sub

Private Sub txtbxBairro_GotFocus()
    'P_SelecTexto txtbxBairro
End Sub

Private Sub txtbxBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxClienteAcerto_Change()
    txtbxClienteAcerto.tag = ""
    
End Sub

Private Sub txtbxClienteAcerto_GotFocus()
    P_SelecTexto txtbxClienteAcerto
End Sub

Private Sub txtbxClienteAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnClienteAcerto_Click
    Else
        
    If optnRep.Value = True Then
        MsgBox "Você optou por fazer acerto para o próprio representante, não é possível escolher qualquer cliente!", vbExclamation
        optnRep_Click
        optnRep.SetFocus
        Exit Sub
    End If
        
    End If
End Sub






Private Sub lblrepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    End If
    If KeyCode = vbKeyDown Then
        
    End If
End Sub


Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub


Private Sub txtbxDataBase_GotFocus()
    P_SelecTexto txtbxDataBase
End Sub

Private Sub txtbxDataBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataBase_LostFocus()
    P_AjustaData txtbxDataBase
End Sub

Private Sub txtbxDifAcerto_GotFocus()
    P_SelecTexto txtbxDifAcerto
End Sub

Private Sub txtbxDifAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDifAcerto_LostFocus()
    txtbxDifAcerto.Text = Format(txtbxDifAcerto.Text, "###,##0.00")
    calculaDif True
End Sub
Sub calculaDif(Optional bTotal As Boolean = True)
    If bTotal Then
        txtbxValorNegociado.Text = FormatVal(CCur(txtbxDifAcerto.Text) + CCur(lblTotalInicial.caption))
    Else
    
        txtbxDifAcerto.Text = FormatVal(CCur(txtbxValorNegociado.Text) - CCur(lblTotalInicial.caption))
    End If
    lblTotalAcerto.caption = FormatVal(CCur(lblTotalInicial.caption) + CCur(txtbxDifAcerto.Text))
End Sub

Private Sub txtbxEntrada_GotFocus()
    P_SelecTexto txtbxEntrada
End Sub

Private Sub txtbxEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxEntrada_LostFocus()
    P_ValorLostFocus txtbxEntrada
    P_CalcTotalAprazo
    distribuiParcela
    somaDesconto
End Sub
Sub P_CalcTotalAprazo()
    txtbxTotalAprazo.Text = FormatVal(CCur(lblValorFinal.caption) - CCur(txtbxEntrada.Text))
End Sub

Private Sub txtbxFormaPag_GotFocus()
    P_SelecTexto txtbxFormaPag
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxMsf_GotFocus()
    'P_SelecTexto txtbxMsf
    If msflxgrd.Col = 7 Then
        If txtbxMsf.Text = "" Then
            txtbxMsf.Text = txtbxClienteAcerto.Text
        End If
    End If
    If msflxgrd.Col = 8 Then
        If txtbxMsf.Text = "" Then
            txtbxMsf.Text = lblCpfCli.caption
        End If
    End If
    txtbxMsf.SelStart = Len(txtbxMsf.Text)
End Sub

Private Sub txtbxMsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
    If KeyCode = vbKeyReturn Then
        If msflxgrd.Col = 4 Then
            strSQL = F_SelectBanco("").cod_bco
            msflxgrd.Col = 4
            msflxgrd.Text = strSQL
            msflxgrd.Col = 5
            msflxgrd.SetFocus
            msflxgrd_KeyPress 13
        Else
            msflxgrd.SetFocus
            SendKeys Chr(9)
        End If
    ElseIf KeyCode = vbKeyRight And txtbxMsf.SelStart = Len(txtbxMsf) Then
        SendKeys Chr(9)
    
    End If
End Sub
Sub buscaBanco()
    Dim b As Bancos.TBanco
     
End Sub
Private Sub txtbxMsf_LostFocus()
    
    Dim iCol As Integer
    Dim iRow As Integer
    On Error Resume Next
    If msflxgrd.Col = 1 Then
        P_ValorLostFocus txtbxMsf
    Else
        P_AjustaTexto txtbxMsf
    End If
    iCol = msflxgrd.Col
    iRow = msflxgrd.Row
    msflxgrd.Text = txtbxMsf.Text
    
    P_SomaItd
    
    
    msflxgrd.Col = iCol + 1
    msflxgrd.Row = iRow
    msflxgrd_KeyDown 13, 0
    
    txtbxMsf.Visible = False
    msflxgrd.SetFocus
End Sub

Private Sub txtbxProdAcerto_Change()
    txtbxProdAcerto.tag = ""
End Sub

Private Sub txtbxProdAcerto_GotFocus()
    P_SelecTexto txtbxProdAcerto
End Sub

Private Sub txtbxProdAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaProdAcerto
        
    End If
End Sub
Sub buscaProdAcerto()
    Dim p As Produtos.ProdutoConsig
    
    If txtbxRepresentante.tag = "" Then
        cmndbtnBuscaRep_Click
    End If
    p = Produtos.F_BuscaProdConsign(txtbxRepresentante.tag, txtbxProdAcerto.Text)
    
    If p.reman_con > 0 Then
        txtbxProdAcerto.Text = StrConv(p.pro.nome_pro, vbProperCase)
        txtbxProdAcerto.tag = p.pro.cod_pro
        lblRefAcerto.caption = p.pro.ref_pro
        
        txtbxReman.tag = p.reman_con - F_QtdProdutoLancadoAcerto(p.pro.cod_pro)
        txtbxReman.Text = txtbxReman.tag
        txtbxQtdAcerto.Text = "1"
        txtbxUnitAcerto.Text = FormatVal(p.unit_con)
        txtbxUnitAcerto.tag = FormatVal(p.unit_con)
        lblDataConAcerto.caption = p.data_con
        lblSeqConAcerto.caption = p.seq_con
        txtbxQtdAcerto.SetFocus
                
    End If
    
    
    
End Sub

Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim e As EstadoImp
    
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim sTipoQtd As String
    If CCur(txtbxDescPercent.Text) > CCur(lblDesc.caption) Then
        MsgBox "O desconto máximo permitido para o usuário é de " & lblDesc.caption, vbExclamation
        Exit Sub
    End If
    
    If txtbxProduto.Text = "" Then
        MsgBox "Informe pelo menos parte do nome do produto para que seja localizado!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    On Error GoTo trata_Erro
                
                
                                
                iItem = iItem + 1
                If Consignado.F_GravaConsignAberta(prod.cod_pro, txtbxQtd.Text, txtbxRepresentante.tag, txtbxUnit.Text) = "" Then
                    MsgBox "Erro na inserção do item!", vbExclamation
                End If
                    
                P_buscaitens
                Me.Refresh
                
                txtbxQtd.Text = "1"
                dtProd.Recordset.Move dtProd.Recordset.RecordCount - 1
                txtbxUnit.Text = "0"
                txtbxDescPercent.Text = "0,0"
                txtbxDescReais.Text = "0,00"
                
                txtbxProduto.SetFocus
            
        
        
        On Error Resume Next

        Exit Sub
trata_Erro:
    P_MostraErro "Insere()"
End Sub
Sub buscaProduto()
    
    
    prod = Produtos.F_BuscaProdSaida(txtbxProduto.Text, "B")
            
            
            
        If prod.cod_pro <> "" Then
                               
            txtbxProduto.Text = prod.nome_pro
                
            'lblNomePro.caption = p.nome_pro
            
            'lblCodPro.caption = p.cod_pro
'            lblSeqPro.caption = p.est.seq_pro
            
            
            'lblSec.caption = p.sub.sec.nome_sec
            'lblSubSec.caption = p.sub.nome_sub
            
            'lblDescSub.caption = p.sub.descmax_sub
            'lblPreco.caption = p.pv_pro
            'If prod.pb_pro > 0 Then
            '    txtbxUnit.Text = FormatVal(prod.pb_pro)
            'Else
                txtbxUnit.Text = FormatVal(prod.pv_pro)
            'End If
            'txtbxDescPercent.Text = CCur(txtbxDescForma.Text) + CCur(p.sub.descmax_sub)
            
            SendKeys Chr(9)
        End If
End Sub
Sub buscaDesc()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select * from tab_descontos where cod_pro='" & txtbxProduto.tag & "' "
    strSQL = strSQL & " and cod_car='"
    
    
    
    
    
End Sub
Sub p_totalItem()
    txtbxFinal.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxDescReais.Text))
    
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    Timer1_Timer
End Sub

Private Sub txtbxQtdAcerto_GotFocus()
    P_SelecTexto txtbxQtdAcerto
End Sub

Private Sub txtbxQtdAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxQtdAcerto_LostFocus
        insereItemAcerto
    End If
End Sub

Private Sub txtbxQtdAcerto_LostFocus()
    P_ValorLostFocus txtbxQtdAcerto
        P_TotalAcerto
End Sub

Sub P_GravaItemCancelado()
    On Error Resume Next
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
Sub P_recalculaVenda()
    Venda.P_recalculaVenda txtbxDescForma.Text, Me.tag
    
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
    
    If bBlock = True Then Exit Sub
    
    P_focoMouse Index
End Sub










Private Sub txtbxRepresentante_Change()
    
    '
End Sub

Private Sub txtbxRepresentante_GotFocus()
    P_SelecTexto txtbxRepresentante
End Sub

Private Sub txtbxRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscaRep_Click
    Else
        txtbxRepresentante.tag = ""
    End If
End Sub

Private Sub txtbxRepresentante_LostFocus()
    P_AjustaTexto txtbxRepresentante
End Sub

Private Sub txtbxTipoCob_GotFocus()
    P_SelecTexto txtbxTipoCob
End Sub

Private Sub txtbxTipoCob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCob_Click
    End If
End Sub

Private Sub txtbxUnitAcerto_GotFocus()
    P_SelecTexto txtbxUnitAcerto
End Sub

Private Sub txtbxUnitAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnitAcerto_LostFocus()
    P_ValorLostFocus txtbxUnitAcerto
    P_TotalAcerto
End Sub
Sub P_TotalAcerto()
    'on error resume next
    txtbxVlrFinalAcerto.Text = FormatVal(CCur(txtbxUnitAcerto.Text) * CCur(txtbxQtdAcerto.Text))
    txtbxDiferenca.Text = FormatVal(CCur(txtbxVlrFinalAcerto.Text) - CCur(txtbxUnitAcerto.tag) * CCur(txtbxQtdAcerto.Text))
    
End Sub

Private Sub txtbxValorNegociado_GotFocus()
    P_SelecTexto txtbxValorNegociado
End Sub

Private Sub txtbxValorNegociado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxValorNegociado_LostFocus()
    P_ValorLostFocus txtbxValorNegociado
    calculaDif False
End Sub

Private Sub txtbxVenIniAcerto_GotFocus()
    
    P_SelecTexto txtbxVenIniAcerto
    txtbxVenIniAcerto.tag = txtbxVenIniAcerto.Text
    
    
    
End Sub

Private Sub txtbxVenIniAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
    If KeyCode = vbKeyRight Then
        If txtbxVenIniAcerto.Text = "" Then
            txtbxVenIniAcerto.Text = Date
        End If
        txtbxVenIniAcerto.Text = DateAdd("d", 1, txtbxVenIniAcerto.Text)
    End If
    If KeyCode = vbKeyLeft Then
    If txtbxVenIniAcerto.Text = "" Then
            txtbxVenIniAcerto.Text = Date
        End If
        txtbxVenIniAcerto.Text = DateAdd("d", -1, txtbxVenIniAcerto.Text)
    End If
End Sub

Private Sub txtbxVenIniAcerto_LostFocus()
    P_AjustaData txtbxVenIniAcerto
    If txtbxVenIniAcerto.tag <> txtbxVenIniAcerto.Text Or msflxgrd.rows = 1 Then
        distribuiParcela
    End If
End Sub
Sub distribuiParcela()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iDias As Integer
    Dim bCheque As Boolean
    msflxgrd.rows = 1
    On Error GoTo trata_Erro
    lblVlrAcertoDiv.caption = FormatVal(CCur(txtbxTotalAprazo.Text) / CCur(lblParcelas.caption))
    If txtbxVenIniAcerto.Text = "" Then
        msflxgrd.rows = 1
        P_SomaItd
        Exit Sub
    End If
    If lblTipoPag.caption = "DUP" Then
        bCheque = False
    Else
        bCheque = True
    End If
    For iCount = 1 To Int(lblParcelas.caption)
            
            If strTipoPeriodoPar = "D" Then
                iDias = (30 * (iCount - IIf(chk.Value = 1, 1, 0))) - 30
                strSQL = DateAdd("D", iDias, txtbxVenIniAcerto.Text) & Chr(9)
            Else
                iDias = (1 * iCount) - 1
                strSQL = DateAdd("M", iDias, txtbxVenIniAcerto.Text) & Chr(9)
            End If
            
            strSQL = strSQL & FormatVal(CCur(txtbxTotalAprazo.Text) / CCur(lblParcelas.caption)) & Chr(9)
            
            If bCheque Then
                strSQL = strSQL & "2 - Cheque pré"
            Else
                strSQL = strSQL & "1 - Duplicata"
            End If
            msflxgrd.AddItem strSQL
            msflxgrd.Row = msflxgrd.rows - 1
            msflxgrd.Col = 1
            'msflxgrd.Text = txtbxBoleto.Text
            cTotal = cTotal + Round((CCur(txtbxTotalAprazo.Text) / CCur(lblParcelas.caption)), 2)
            
        Next
            P_SomaItd
            If CCur(lblTotalDistribuido.caption) <> CCur(lblValorFinal.caption) Then
                msflxgrd.Row = msflxgrd.rows - 1
                msflxgrd.Col = 1
                msflxgrd.Text = FormatVal(CCur(msflxgrd.Text) + CCur(txtbxTotalAprazo.Text) - CCur(lblTotalDistribuido.caption))
            End If
            P_SomaItd
            
                    
            Exit Sub
trata_Erro:
          P_MostraErro
            
End Sub
Sub ajustaGrid()
    AjustaGr msflxgrd
    AjustaGr msflxgrdAv
End Sub
Sub AjustaGr(obj As Object)
    bCheque = False
    txtbxVenIniAcerto.Text = Date
    
    'P_CentraForm Me, 600
    obj.rows = 1
    obj.Cols = 10
    obj.Col = 0
    obj.Text = "Vencimento"
    obj.Col = 1
    obj.Text = "Valor"
    obj.Col = 2
    obj.Text = "Espécie"
    obj.Col = 3
    obj.Text = "Cheque Nº"
    obj.Col = 4
    obj.Text = "Ban/Con"
    obj.Col = 5
    obj.Text = "Agencia"
    obj.Col = 6
    obj.Text = "Conta C."
    obj.Col = 7
    obj.Text = "Titular"
    obj.Col = 8
    obj.Text = "CPF"
    obj.Col = 9
    obj.Text = "Boleto"
    obj.ColAlignment(2) = 1
    
    obj.ColWidth(0) = 1000
    
    obj.ColWidth(1) = 1000
    obj.ColWidth(2) = 1500
    obj.ColWidth(7) = 3000
    obj.ColWidth(8) = 1500
End Sub
Private Sub txtbxVlrFinalAcerto_GotFocus()
    P_SelecTexto txtbxVlrFinalAcerto
End Sub

Private Sub txtbxVlrFinalAcerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        insereItemAcerto
    End If
End Sub
Sub insereItemAcerto()
    Dim strSQL As String
    txtbxVlrFinalAcerto_LostFocus
    If CCur(txtbxVlrFinalAcerto.Text) = 0 Then
        Exit Sub
    End If
    If CCur(txtbxQtdAcerto.Text) > CCur(txtbxReman.Text) Then
        MsgBox "A quantidade que está lançando é maior que a restante!", vbExclamation
        txtbxQtdAcerto.Text = txtbxReman.Text
        txtbxQtdAcerto.SetFocus
        Exit Sub
    End If

    If txtbxProdAcerto.tag = "" Then
        MsgBox "Escolha o produto e tecle enter", vbExclamation
        txtbxProdAcerto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxVlrFinalAcerto.Text) = 0 Then
        MsgBox "O Valor final não pode ser zero!", vbExclamation
        Exit Sub
    End If
    strSQL = txtbxProdAcerto.tag & Chr(9)
    strSQL = strSQL & lblRefAcerto.caption & Chr(9)
    strSQL = strSQL & txtbxProdAcerto.Text & Chr(9)
    strSQL = strSQL & txtbxQtdAcerto.Text & Chr(9)
    strSQL = strSQL & txtbxUnitAcerto.tag & Chr(9)
    strSQL = strSQL & txtbxUnitAcerto.Text & Chr(9)
    strSQL = strSQL & txtbxVlrFinalAcerto.Text & Chr(9)
    strSQL = strSQL & txtbxDiferenca.Text & Chr(9)
    strSQL = strSQL & lblDataConAcerto.caption & Chr(9)
    strSQL = strSQL & lblSeqConAcerto.caption
    msflxgrdAcerto.AddItem strSQL
    
    txtbxUnitAcerto.Text = "0"
    txtbxProdAcerto.tag = ""
    txtbxVlrFinalAcerto.Text = "0"
    txtbxQtdAcerto.Text = "1"
    txtbxDiferenca.Text = "0"
    txtbxProdAcerto.SetFocus
    txtbxReman.tag = "0"
    txtbxReman.Text = "0"
    
    txtbxUnitAcerto.tag = "0"
    P_SomaAcerto
    P_SomaItd
    somaDesconto
    msflxgrdAcerto.TopRow = msflxgrdAcerto.rows - 1
End Sub

Private Sub txtbxVlrFinalAcerto_LostFocus()
    On Error Resume Next
    P_ValorLostFocus txtbxVlrFinalAcerto
    txtbxUnitAcerto.Text = FormatVal(CCur(txtbxVlrFinalAcerto.Text) / CCur(txtbxQtdAcerto.Text))
    P_TotalAcerto
End Sub


Sub P_SomaItd()
    Dim cVal As Currency
    Dim iCount As Integer
    Dim iRow As Integer
    Dim cIni As Currency
    Dim obj As Object
    
    If sstb.Tab = 0 Then
        Set obj = msflxgrd
    Else
        Set obj = msflxgrdAv
    End If
    
    iRow = obj.Row
    With obj
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 1
            cVal = cVal + CCur(.Text)
            
        Next
        
    End With
    
    obj.Row = iRow
    lblTotalDistribuido.caption = FormatVal(cVal + CCur(txtbxEntrada.Text))
End Sub

Function F_confMsflx(msf As MSFlexGrid) As Boolean
     Dim iCount As Integer
     Dim sTipo As String
     Dim Retorno As Boolean
     Dim strSQL As String
     
     Retorno = True
     strSQL = "os seguintes dados estão faltando: " & Chr(13) & Chr(13)
     With msf
        For iCount = 1 To .rows - 1
            .Row = iCount
            If 1 = 0 Then
                .Col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .Col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    
                    
                    
                    
                    
                End If
            
            Else
                .Col = 2
                sTipo = Left(.Text, 1)
                If sTipo = "2" Then
                    .Col = 3
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 4
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    
                    If F_BuscaUmCampo("Nome_bco", "tab_bancos", "cod_bco='" & .Text & "'") = "" Then
                        strSQL = strSQL & "linha " & iCount & " não existe o banco de código " & .Text & Chr(13)
                        Retorno = False
                    End If
                    .Col = 5
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " agência" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 6
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " conta" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 7
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " nome do titular" & Chr(13)
                        Retorno = False
                    End If
                    .Col = 8
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " CPF do titular" & Chr(13)
                        Retorno = False
                    End If
                End If
            End If
        Next
     End With
     F_confMsflx = Retorno
     If Retorno = False Then
        MsgBox strSQL, vbExclamation
        
     End If
End Function

