VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVenda 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "-Venda de Produtos:"
   ClientHeight    =   11310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16170
   ControlBox      =   0   'False
   Icon            =   "frmVenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11310
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grade 
      Height          =   1815
      Left            =   5520
      TabIndex        =   93
      Top             =   2280
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmVenda.frx":08CA
      Height          =   2775
      Left            =   5520
      TabIndex        =   91
      Top             =   1800
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4895
      _Version        =   393216
      Appearance      =   0
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "seq_ite"
         Caption         =   "..."
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
         DataField       =   "ref_pro"
         Caption         =   "Referência"
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
      BeginProperty Column02 
         DataField       =   "desc_ite"
         Caption         =   "Descrição do Produto"
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
      BeginProperty Column03 
         DataField       =   "qtd_ite"
         Caption         =   "Qtd"
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
      BeginProperty Column04 
         DataField       =   "unit_ite"
         Caption         =   "Preço"
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
      BeginProperty Column05 
         DataField       =   "total"
         Caption         =   "Total"
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
      BeginProperty Column06 
         DataField       =   "desconto"
         Caption         =   "Desconto"
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
      BeginProperty Column07 
         DataField       =   "final"
         Caption         =   "Final"
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
      BeginProperty Column08 
         DataField       =   "pq"
         Caption         =   "Queima"
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
      BeginProperty Column09 
         DataField       =   "nome_sub"
         Caption         =   "Subseção"
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
      BeginProperty Column10 
         DataField       =   "ncm_pro"
         Caption         =   "NCM"
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
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3509,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   734,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   810,142
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2190,047
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcProd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5520
      ScaleHeight     =   945
      ScaleWidth      =   10545
      TabIndex        =   17
      Top             =   720
      Width           =   10575
      Begin VB.CommandButton cmndbtnInserir 
         Caption         =   "Inserir"
         Height          =   285
         Left            =   9000
         TabIndex        =   88
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtbxPQ_pro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   6480
         TabIndex        =   86
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxFinal 
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
         Left            =   9000
         TabIndex        =   29
         Text            =   "0,00"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtbxDescReais 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   8160
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxDescPercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   7320
         TabIndex        =   25
         Text            =   "0,0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4680
         TabIndex        =   21
         Text            =   "1"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbxProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lblPq_pro 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Promoção."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6480
         TabIndex        =   87
         Top             =   0
         Width           =   765
      End
      Begin VB.Label lblQueima 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "Queima!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   120
         TabIndex        =   85
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(R$)"
         Height          =   195
         Left            =   8160
         TabIndex        =   26
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor Final:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9120
         TabIndex        =   28
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.:(%)"
         Height          =   195
         Left            =   7320
         TabIndex        =   24
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   22
         Top             =   0
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
         TabIndex        =   18
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
         TabIndex        =   20
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox pcCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4665
      ScaleWidth      =   5385
      TabIndex        =   72
      Top             =   600
      Width           =   5415
      Begin VB.CheckBox chkEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Entrega (F10)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   2640
         Width           =   1335
      End
      Begin VB.PictureBox pcEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCB683&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         ScaleHeight     =   1425
         ScaleWidth      =   5385
         TabIndex        =   103
         Top             =   2760
         Visible         =   0   'False
         Width           =   5415
         Begin VB.TextBox txtValorEntrega 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   120
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtEnderecoEntrega 
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
            TabIndex        =   104
            Tag             =   "00000"
            Top             =   360
            Width           =   5175
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total com entrega:"
            Height          =   195
            Left            =   3480
            TabIndex        =   126
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label lblTotalComEntrega 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   540
            Left            =   3720
            TabIndex        =   125
            Tag             =   "!"
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Left            =   120
            TabIndex        =   119
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço da entrega:"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   120
            Width           =   1545
         End
      End
      Begin VB.TextBox txtAniversario 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
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
         Left            =   4320
         TabIndex        =   16
         Tag             =   "00000"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
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
         TabIndex        =   14
         Tag             =   "00000"
         Top             =   2040
         Width           =   5175
      End
      Begin VB.TextBox txtCpf 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
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
         Left            =   2640
         TabIndex        =   12
         Tag             =   "00000"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
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
         TabIndex        =   10
         Tag             =   "00000"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtbxDescForma 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtbxFormaPag 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "A Vista"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtbxNomeCli 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
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
         Top             =   840
         Width           =   4935
      End
      Begin VB.CommandButton cmndbtnFormasPag 
         Height          =   315
         Left            =   2760
         Picture         =   "frmVenda.frx":08DC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmnbtnRecalc 
         Caption         =   "Recalcula Desconto"
         Height          =   435
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnCliente 
         Height          =   330
         Left            =   5040
         Picture         =   "frmVenda.frx":09D5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aniversário:"
         Height          =   195
         Left            =   4320
         TabIndex        =   15
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF"
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   675
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
         Left            =   1800
         TabIndex        =   84
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblNomeCli 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Forma de Pagamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   1560
      End
   End
   Begin VB.PictureBox pcStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EEDFD5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   16170
      TabIndex        =   53
      Top             =   10455
      Width           =   16170
      Begin VB.PictureBox pcDesconto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3480
         ScaleHeight     =   1095
         ScaleWidth      =   5415
         TabIndex        =   106
         Top             =   0
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton cmndbtnDesconto 
            Caption         =   "Confirmar Gerencial"
            Height          =   525
            Left            =   2160
            Picture         =   "frmVenda.frx":0ACE
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtbxDescPerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   108
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtbxDesconto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   107
            Text            =   "0,00"
            Top             =   480
            Width           =   735
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   4200
            TabIndex        =   118
            Top             =   480
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final Previsto:"
            Height          =   195
            Left            =   4080
            TabIndex        =   117
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   1200
            TabIndex        =   116
            Top             =   240
            Width           =   120
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto:"
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Autorizado Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   114
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblAutorizaDesconto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   2040
            TabIndex        =   113
            Top             =   817
            Width           =   510
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Desconto Gerencial"
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
            Left            =   90
            TabIndex        =   112
            Top             =   -30
            Width           =   1860
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Já Utilizado:"
            Height          =   195
            Left            =   2160
            TabIndex        =   111
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblDescontoGerencialUsado 
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
            Left            =   3120
            TabIndex        =   110
            Top             =   0
            Width           =   390
         End
      End
      Begin VB.TextBox txtbxVendedor 
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
         Left            =   120
         TabIndex        =   81
         Text            =   "Nome do Funcionario"
         Top             =   255
         Width           =   3015
      End
      Begin VB.CommandButton cmdBuscarFun 
         Height          =   315
         Left            =   3120
         Picture         =   "frmVenda.frx":1238
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   255
         Width           =   255
      End
      Begin VB.Label lblfracionado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "fracionado"
         Height          =   195
         Left            =   1200
         TabIndex        =   124
         Top             =   600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblTotalQueima 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   10320
         TabIndex        =   122
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblTotalQueimaCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Queima:"
         Height          =   195
         Left            =   9000
         TabIndex        =   121
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDescontoMax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   10560
         TabIndex        =   90
         Top             =   165
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9000
         TabIndex        =   83
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desc:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   82
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lblTur 
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
         Left            =   13560
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblBloqueado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
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
         Left            =   13680
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12960
         TabIndex        =   60
         Top             =   60
         Visible         =   0   'False
         Width           =   465
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
         Left            =   12600
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caixa:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12000
         TabIndex        =   58
         Top             =   60
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lblNomeEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
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
         Left            =   14160
         TabIndex        =   57
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Controle:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9840
         TabIndex        =   56
         Top             =   180
         Width           =   630
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
         Left            =   10800
         TabIndex        =   55
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   15
         Width           =   780
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
      TabIndex        =   52
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
         TabIndex        =   78
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
            Picture         =   "frmVenda.frx":15C2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   79
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
            TabIndex        =   80
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
         TabIndex        =   30
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
            Picture         =   "frmVenda.frx":1FC4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   31
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
            TabIndex        =   32
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
         TabIndex        =   33
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
            Picture         =   "frmVenda.frx":234E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   34
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
            TabIndex        =   35
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
         TabIndex        =   36
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
            Picture         =   "frmVenda.frx":25C2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   37
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
            TabIndex        =   38
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
         TabIndex        =   39
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
            Picture         =   "frmVenda.frx":270C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   40
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
            TabIndex        =   41
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
         TabIndex        =   42
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
            Picture         =   "frmVenda.frx":310E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   43
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
            TabIndex        =   44
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
         TabIndex        =   45
         Top             =   10
         Width           =   1800
         Begin VB.PictureBox pcIco 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "frmVenda.frx":54D5
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   46
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
            TabIndex        =   47
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
         TabIndex        =   48
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
            Picture         =   "frmVenda.frx":585F
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   49
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
            TabIndex        =   50
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmndbtnLocaliza 
      Caption         =   "..."
      Height          =   255
      Left            =   14760
      TabIndex        =   51
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   0
      ScaleHeight     =   9885
      ScaleWidth      =   5385
      TabIndex        =   63
      Top             =   540
      Width           =   5415
      Begin VB.PictureBox pcLimite 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   5415
         TabIndex        =   94
         Top             =   1560
         Visible         =   0   'False
         Width           =   5415
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            Height          =   195
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblStatus_CLi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Liberado"
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
            Left            =   0
            TabIndex        =   101
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Limite"
            Height          =   195
            Left            =   1200
            TabIndex        =   100
            Top             =   0
            Width           =   405
         End
         Begin VB.Label lblLimite 
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
            Left            =   1200
            TabIndex        =   99
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Em Uso:"
            Height          =   195
            Left            =   2160
            TabIndex        =   98
            Top             =   0
            Width           =   600
         End
         Begin VB.Label lblEmUso 
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
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2160
            TabIndex        =   97
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   3120
            TabIndex        =   96
            Top             =   0
            Width           =   690
         End
         Begin VB.Label lblLimiteRest 
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
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   3120
            TabIndex        =   95
            Top             =   240
            Width           =   390
         End
      End
      Begin VB.PictureBox pcLogo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2625
         Left            =   1440
         Picture         =   "frmVenda.frx":59CA
         ScaleHeight     =   2625
         ScaleWidth      =   2280
         TabIndex        =   89
         Top             =   7200
         Width           =   2280
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEDFD5&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   5385
         TabIndex        =   64
         Top             =   4680
         Width           =   5415
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            FillColor       =   &H00C0E0FF&
            ForeColor       =   &H00C0E0FF&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   73
            Top             =   1680
            Width           =   5415
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   360
               TabIndex        =   77
               Top             =   0
               Width           =   570
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   1320
               TabIndex        =   76
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label lblDesconto 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   4440
               TabIndex        =   75
               Tag             =   "!"
               Top             =   0
               Width           =   105
            End
            Begin VB.Label Label6 
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   3360
               TabIndex        =   74
               Top             =   0
               Width           =   930
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5415
            TabIndex        =   66
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   3840
               TabIndex        =   70
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   4440
               TabIndex        =   69
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   1320
               TabIndex        =   68
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
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   360
               TabIndex        =   67
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
               Size            =   63.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1440
            Left            =   1560
            TabIndex        =   65
            Tag             =   "!"
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.Label lblTipoOperacao 
         Alignment       =   2  'Center
         Caption         =   "Venda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         TabIndex        =   127
         Top             =   7920
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label lblPExtenso 
         Alignment       =   2  'Center
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
         Height          =   210
         Left            =   2520
         TabIndex        =   71
         Top             =   6600
         Width           =   150
      End
   End
End
Attribute VB_Name = "frmVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents AdoVendas As ADODB.Recordset
Attribute AdoVendas.VB_VarHelpID = -1
Dim p As produto
Public status As String
Dim bBlock As Boolean
Dim bDirecao As Integer
Dim bActive As Boolean
Dim preco As Currency
Public operacao As String
Public tipo_ven As String 'A - Atacado - V - Venda  - C - Consumo
Dim iGif As Integer
Dim iItem As Integer
Dim bDescancoTela As Boolean
Dim iDescancoTela As Integer
Public Tipo_Operacao As tipo_operacao_venda

Public Atacado As Boolean

Private Type prodPromo
    nome As String
    preco As String
End Type

Dim pPromo() As prodPromo
Dim QtdPromo As Integer
Dim iIndPromo As Integer

Sub carrega()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    If bActive = False Then
        strSQL = F_VerificaVendaAberta(Me.tag)
        If strSQL <> "" Then
            MsgBox "A venda se encontra aberta no computador " & strSQL, vbExclamation
            Unload Me
            Exit Sub
        End If
    End If
    strSQL = "select tv.num_ven,valor_ven,telefonecli_ven,valorentrega_ven,enderecoentrega_ven,"
    strSQL = strSQL & " endereco_ven , bairro_ven, fone_ven, desconto_ven, nomecli_ven, entrega_ven, "
    strSQL = strSQL & " cod_cli,valor_ven,email_ven,aniversario_ven,tipo_operacao,"
    strSQL = strSQL & "tf.*,tpp.nome_tip from tab_vendas as tv, tab_formaspag as tf, tab_tipopag as tpp "
    strSQL = strSQL & " where tf.cod_for=tv.cod_for "
    strSQL = strSQL & " and num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tpp.cod_tip=tf.cod_tip "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If Not IsNull(!nomecli_ven) Then
            txtbxNomeCli.text = !nomecli_ven
        End If
        If Not IsNull(!telefonecli_ven) Then
            txtTel.text = !telefonecli_ven
        End If
        txtbxNomeCli.tag = !cod_cli
        
        If UCase(txtbxNomeCli.text) = "VENDA A CONSUMIDOR" Then
            txtbxNomeCli.text = ""
        End If
        txtbxFormaPag.text = !nome_for
        txtbxDescForma.text = !desconto_for
        txtbxFormaPag.tag = !cod_For
        'txtbxEndereco.Text = Format(!endereco_ven)
        'txtbxBairro.Text = Format(!bairro_ven)
        'txtbxFone.Text = Format(!fone_ven)
        lblTipoPag.tag = !cod_tip
        lblTipoPag.caption = !nome_tip
        lblDescontoMax.caption = !desconto_for + CCur(lblDesc.caption)
        If !entrega_ven = "S" Then
            chkEntrega.Value = 1
        
        End If
        txtValorEntrega.text = F_GravaDec(!valorentrega_ven)
        txtEnderecoEntrega.text = TiraNull(!enderecoentrega_ven)
        txtAniversario.text = TiraNull(!aniversario_ven)
        txtEmail.text = TiraNull(!email_ven)
        Me.Tipo_Operacao = !Tipo_Operacao
        
        P_GravaVendaAberta Me.tag
        .Close
    End With
    
    
    If Tipo_Operacao = 2 Then
        pcLeft.BackColor = vbYellow
        lblTipoOperacao.caption = "Devolução"
        lblNomeCli.caption = "Favorecido"
        txtbxNomeCli.BackColor = vbYellow
    End If
    If Tipo_Operacao = 3 Then
        pcLeft.BackColor = vbBlue
        lblTipoOperacao.caption = "Remessa"
        lblNomeCli.caption = "Destinatário"
        txtbxNomeCli.BackColor = vbBlue
        
    End If
    If Tipo_Operacao = 1 Then
        pcLogo.Visible = True
        lblTipoOperacao.Visible = False
    Else
        pcLogo.Visible = False
        lblTipoOperacao.Visible = True
    End If
    
    
    Exit Sub
trata_Erro:
    
    P_MostraErro
End Sub



Sub p_totalItem()
    txtbxFinal.text = FormatVal(CCur(txtbxUnit.text) * CCur(txtbxQtd.text) - (CCur(txtbxDescReais.text)))
    
End Sub

Sub P_AtualizaVenda()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim iPos As Integer
    Dim cItens As Integer
    Dim cDesc As Currency
    On Error GoTo trata_Erro:
    Dim cPq As Currency
    P_Status "Atualizando Valores..."
    
    dtgrd.Visible = False
    iPos = AdoVendas.AbsolutePosition
    AdoVendas.Requery
    If AdoVendas.RecordCount > 0 Then
        AdoVendas.MoveFirst
    End If
    With AdoVendas
        
        
        Do While Not .EOF
            cTotal = cTotal + !final
            cPq = cPq + !pq_itd * !qtd_ite
            cItens = cItens + !qtd_ite
            cDesc = cDesc + !desconto_ite
            .MoveNext
        Loop
        If chkEntrega.Value = 1 Then
            cTotal = cTotal '+ CCur(txtValorEntrega.Text)
        
        End If
        
        lblValor.caption = FormatVal(cTotal, "###,##0.000")
        lblValor.caption = Mid(lblValor.caption, 1, Len(lblValor.caption) - 1)
        lblItens.caption = AdoVendas.RecordCount
        lblTotalQueima.caption = FormatVal(cPq)
        lblDesconto.caption = FormatVal(cDesc)
        lblTotal.caption = FormatVal(cTotal + cDesc)
        lblTotalComEntrega.caption = FormatVal(cTotal + CCur(txtValorEntrega.text))
    End With
        
    
    
    If AdoVendas.RecordCount > 0 Then
        AdoVendas.MoveFirst
        AdoVendas.Move iPos
    End If
    
    
    lblProd.caption = AdoVendas.RecordCount
    lblItens.caption = cItens

    strSQL = "update tab_vendas set valor_ven =" & F_GravaDec(cTotal) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(cDesc) & ","
    strSQL = strSQL & "nomecli_ven= '" & txtbxNomeCli.text & "', "
    strSQL = strSQL & "telefonecli_ven= '" & txtTel.text & "', "
    strSQL = strSQL & " cod_cli= '" & txtbxNomeCli.tag & "', "
    strSQL = strSQL & " cod_cai='" & strCod_Cai & "', "
    If chkEntrega.Value = 1 Then
        strSQL = strSQL & " entrega_ven='S', "
    Else
        strSQL = strSQL & " entrega_ven='N', "
    End If
    strSQL = strSQL & " valorentrega_ven=" & F_GravaDec(txtValorEntrega.text) & ", "
    strSQL = strSQL & " enderecoentrega_ven='" & txtEnderecoEntrega.text & "', "
    strSQL = strSQL & " abrevenda_usu='" & txtbxVendedor.tag & "' "
'    strSQL = strSQL & " obs_VEN='" & txtbxObs.Text & "' "
    strSQL = strSQL & " where num_Ven='" & Me.tag & "' "
    Conn.Execute strSQL
    
    dtgrd.Visible = True
    P_Status ""
    Exit Sub
trata_Erro:
P_Status ""
'    P_MostraErro
End Sub
Sub P_buscaitens()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = Vendas.F_BuscaItens(Me.tag)
    FechaObjeto AdoVendas
    
    
    
    
    abreObjetoAdo AdoVendas, strSQL, dtgrd
    
    
   ' AjustaGridData grade, 10, 1, 1, "..., Referência, Produto, Qtd, Preço, Total, Desconto, Final, Queima"
    
    
    
            
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbExclamation
End Sub


Private Sub cmndbtnCancel_Click()
    Dim strSQL  As String
    Dim i As Integer
    On Error Resume Next
    If AdoVendas.RecordCount > 0 Then
    
    If strExclVenda_usu <> "S" Then
        strSQL = F_senhaReturnUsu
        If F_BuscaUmCampo("exclvenda_usu", "tab_defusu", "cod_usu='" & strSQL & "'") <> "S" Then
            MsgBox "Usuário Não Autorizado!", vbCritical
            Exit Sub
        End If
    End If
    End If
    If F_BuscaUmCampo("num_ven", "tab_venfec", "num_ven='" & Me.tag & "'") <> "" Then
        MsgBox "Cancelamento impossível", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja cancelar esta venda?", vbYesNo + vbCritical) = vbYes Then
        
        With AdoVendas
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                    
                Excluir False
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
        Conn.Execute strSQL
        
        
        If F_BuscaUmCampo("num_ven", "tab_venfec", "num_ven='" & Me.tag & "'") <> "" Then
            MsgBox "Cancelamento impossível", vbExclamation
            Exit Sub
        End If
        
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
                
        If IsLoad("frmEdiVendas") Then
'            RefreshDataPosition frmEdiVendas.adc
        End If
        
        If MsgBox("Deseja Gerar uma Nova Venda?" & Chr(13) & "Clique em não caso queira abandonar o modo de vendas!", vbYesNo + vbQuestion) = vbYes Then
            strSQL = F_GeraNovaVenda("00000")
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
    


trata_Erro:
MsgBox "Err: cmndbtnEditProd_Click() "
End Sub

Private Sub cmndbtnExc_Prod_Click()
    Excluir True
End Sub
Sub Excluir(bQuestion As Boolean)
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim situacao_ven As String
    On Error GoTo trata_Erro
    With AdoVendas
        If .RecordCount = 0 Then
            Exit Sub
        End If
        
        situacao_ven = F_BuscaUmCampo("situacao_ven", TAB_VENDAS, "num_ven='" & Me.tag & "'")
    
    If situacao_ven = "3" Or situacao_ven = "4" Then
        MsgBox "Não é possível excluir produtos, o status dessa venda não permite!", vbExclamation
        Exit Sub
    End If
    If situacao_ven = "2" Then
        If MsgBox("Essa venda já se encontra fechada, deseja realmente excluir produtos?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
        
        If bQuestion = True Then
            If MsgBox("Confirma Exclusão do Produto " & !desc_ite & "?", _
            vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        
        
        
        iPos = .AbsolutePosition
        strSQL = "delete from tab_itvendas where seq_ite=" & !seq_ite & " "
        
        strSQL = strSQL & " and num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        P_GravaItemCancelado
        
        
            
        
        'AdoVendas.
        If iPos = .RecordCount - 1 Then
            If AdoVendas.RecordCount > 0 Then
                AdoVendas.MoveLast
            End If
        Else
            AdoVendas.Move iPos
        End If
        If bQuestion Then
            P_AtualizaVenda
        End If
        If AdoVendas.RecordCount = 0 Then
            txtbxQtd.Enabled = False
            txtbxProduto.Enabled = True
            txtbxProduto.SetFocus
        End If
    End With
    
    
    P_buscaitens
    If AdoVendas.RecordCount = 0 Then
        lblDescontoGerencialUsado.caption = "0,00"
    End If
    Exit Sub
trata_Erro:
    P_MostraErro "Excluir()"
End Sub


Private Sub AdoVendas_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    With AdoVendas
    
    
    End With
End Sub


Private Sub chkEntrega_Click()
    txtValorEntrega.text = "0"
    If chkEntrega.Value = 1 Then
        txtValorEntrega.text = BuscaParametroEstrutura("taxa_entrega").valor
        txtValorEntrega_LostFocus
        pcEntrega.Visible = True
            
    Else
        pcEntrega.Visible = False
        txtValorEntrega.text = "0"
        txtValorEntrega_LostFocus
        P_AtualizaVenda
        
    End If
End Sub

Private Sub cmnbtnRecalc_Click()
    If MsgBox("Confirma cancelamento do desconto atual?", vbYesNo + vbQuestion) = vbYes Then
        P_CancelaDescontoVenda Me.tag, txtbxDescForma.text
        P_buscaitens
        P_AtualizaVenda
    End If
End Sub

Private Sub cmndbtnCancelar_Click()

End Sub

Private Sub cmdBuscarFun_Click()
    If strVisualizarvenda_Usu = "S" Then
        buscaVendedor
    End If
    
End Sub
Sub buscaVendedor()
    Dim f As funcionario.funcionario
    
    f = funcionario.F_SelectVendedor(txtbxVendedor.text)
    
    If f.cod_fun <> txtbxVendedor.tag Then
        
    End If
    
    
End Sub
Private Sub cmndbtnCliente_Click()
     Dim c As cliente
     c = Clientes.F_BuscaCliente(txtbxNomeCli.text)
     If c.cod_cli <> "" Then
        txtbxNomeCli.text = c.nome_cli
        txtbxNomeCli.tag = c.cod_cli
        txtTel.text = c.telefone_cli
     End If
    
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
    If CCur(sDes) < CCur(txtbxDescPerc.text) + CCur(lblDescontoGerencialUsado.caption) Then
        MsgBox "O Desconto possível para este usuário é de " & sDes & "%", vbExclamation
        txtbxDesconto.SetFocus
        Exit Sub
    End If
    
    
    If MsgBox("Confirma desconto de " & txtbxDesconto.text & "?", vbYesNo + vbCritical) = vbYes Then
        
        Vendas.P_RecalculaVendaDesconto lblFinalPrev.caption, txtbxDescPerc.text, Me.tag, sUsu
        
        P_AtualizaVenda
    End If
    lblDescontoGerencialUsado.caption = CCur(lblDescontoGerencialUsado.caption) + CCur(txtbxDescPerc.text)
    
End Sub

Private Sub cmndbtnFormasPag_Click()
    Dim f As FormasPag.FormasPag
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    f = F_BuscaFormasPag("", "")
    
    If f.cod_For <> "" Then
        If CCur(txtbxDescForma.text) > f.desconto_for Then
            If AdoVendas.RecordCount > 0 Then
                strSQL = "Você optou por mudar a forma de pagamento de " & txtbxFormaPag.text
                strSQL = strSQL & " para " & f.nome_for & Chr(13)
                strSQL = strSQL & "Continuando com esta operação, todo o desconto da venda será cancelado!"
                strSQL = strSQL & "Deseja continuar?"
                If MsgBox(strSQL, vbOKCancel) = vbCancel Then
                    Exit Sub
                Else
                    P_CancelaDescontoVenda Me.tag
                    P_buscaitens
                End If
            End If
        End If
        txtbxFormaPag.text = f.nome_for
        txtbxFormaPag.tag = f.cod_For
        txtbxDescForma.text = f.desconto_for
        txtbxDescForma.tag = f.avista_For
        lblDescontoMax.caption = f.desconto_for + CCur(lblDesc.caption)
        txtbxDescPercent.text = txtbxDescForma.text
        lblTipoPag.caption = f.nome_tip
        lblTipoPag.tag = f.cod_tip
        
        'P_recalculaVenda
        P_AtualizaFormaPag Me.tag, txtbxFormaPag.tag
        P_AtualizaVenda
        
        If Me.ActiveControl.Name = "txtbxNomeCli" Then
            
        Else
            txtbxNomeCli.SetFocus
        End If
        
        
        Exit Sub
    End If
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub cmndbtnInserir_Click()
    insere
End Sub

Private Sub cmndbtnLocaliza_Click()
    
    buscaProd
    
    
End Sub

Private Sub cmndbtnSair_Click()
    Dim strSQL As String
    On Error GoTo trata_Erro
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, txtTel.text
    P_AtualizaVenda
    If AdoVendas.RecordCount > 0 Then
        
            If MsgBox("Confirma Saída da Venda?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            If txtbxNomeCli.text = "" Or txtbxNomeCli.text = "VENDA A CONSUMIDOR" Then
                MsgBox "Informe o nome do cliente!", vbExclamation
                txtbxNomeCli.SetFocus
                Exit Sub
            End If
            
            
    End If
    
    If AdoVendas.RecordCount = 0 And (status = "1" Or status = "") Then
        strSQL = "delete from tab_itvendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
        strSQL = "delete from tab_vendas where num_ven='" & Me.tag & "' "
        Conn.Execute strSQL
        
    End If
    P_DeletaVendaAberta Me.tag
    Unload Me
    Exit Sub
trata_Erro:
    P_MostraErro "cmndbtnSair"
    

End Sub


Private Sub Concluir()
    Dim sTroco As String
    Dim strSQL As String
    Dim sVen As String
    Dim s As String
    If AdoVendas.RecordCount = 0 Then
        MsgBox "Não há itens na sua venda!", vbExclamation
        Exit Sub
    End If
    
    
    Dim ncm As String
    ncm = VerificaNCMVenda(Me.tag)
    If ncm <> "" Then
        If MsgBox("Atenção, essa venda possui produtos com NCM inválido ou não presente." & Chr(13) & ncm & Chr(13) & "Deseja continuar? ", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        
    End If
    
    
    
    
    
    strSQL = F_senhaReturnUsu(False, True)
    If strSQL = "" Then
        MsgBox "não autenticado!", vbExclamation
        Exit Sub
    End If
    If strSQL <> txtbxVendedor.tag Then
        txtbxVendedor.tag = strSQL
        txtbxVendedor.text = F_BuscaUmCampo("nome_fun", "tab_funcionarios", "cod_fun='" & strSQL & "'")
        Vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text
        P_AtualizaVenda
    End If
    
    
    
    
        If txtbxNomeCli.text = "" Then
            strSQL = InputBox("Informe um nome para o cliente", "Nome do Cliente", "Venda a consumidor")
            If strSQL = "" Then
                MsgBox "Informe um nome para o cliente!", vbExclamation
                txtbxNomeCli.SetFocus
                Exit Sub
            Else
                txtbxNomeCli.text = strSQL
                Vendas.P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text
            End If
            
            
            
        End If
        
        
        If Trim(F_SoNumero(txtTel.text)) = "" And Tipo_Operacao = venda Then
            MsgBox "Informe um telefone para o cliente!", vbExclamation
            txtTel.SetFocus
            Exit Sub
        End If
        
        If bUsaCrediario And Me.Tipo_Operacao <= 1 Then
            If lblTipoPag.tag = "02" Then
                If F_GravaCrediarioVen(Me.tag) Then
                    MsgBox "Esta venda venda deve ser liberada pelo setor de Crediário para ser concretizada!" & Chr(13) & "Por favor, encaminhe o cliente ao Crediário.", vbInformation
                End If
            End If
        End If
        
        
        If Me.Tipo_Operacao <= 1 Then
        
            If Vendas.F_AlteraStatusVenda(Me.tag, "2") Then
                
                If Not bUsaCaixaSeparado Then
                    s = F_Troco(Me.tag)
                    If s = "" Then
                        Vendas.F_AlteraStatusVenda Me.tag, "1"
                        Exit Sub
                    End If
                End If
                
                P_DeletaVendaAberta Me.tag
                
              
                
                
                Me.tag = F_GeraNovaVenda
                If Me.tag = "" Then
                    txtbxNomeCli.text = ""
                    txtTel.text = ""
                    
                    MsgBox "Não foi possível gerar uma nova venda!", vbExclamation
                    Unload Me
                    Exit Sub
                End If
                bActive = False
                Form_Activate
            End If
        Else
            If MsgBox("Confirma efetivação da movimentação?", vbYesNo + vbQuestion) = vbYes Then
                If Vendas.FechaRemessaDevolucao(Me.tag, "4") Then
                    Dim id As String
                    Dim descricaoOperacao As String
                    If Me.Tipo_Operacao = tipo_operacao_venda.Devolucao Then
                        descricaoOperacao = "Devolução de Mercadoria"
                    ElseIf Me.Tipo_Operacao = tipo_operacao_venda.retorno_remessa Then
                        descricaoOperacao = "Retorno de Remessa"
                    
                    ElseIf Me.Tipo_Operacao = tipo_operacao_venda.Remessa Then
                        descricaoOperacao = "Remessa para Transporte"
                    End If
                    P_Status "aguarde... gerando Nota Fiscal"
                    
                    strSQL = F_GravaCai("C", Date, "0", descricaoOperacao & Me.tag & "EM " & Date, Me.tag, "DIN", "", 0, 0, strCod_Cai, strCod_Tur, "000", "S")
                    
                    id = NFe_DAO.GravaNFCeVenda(Me.tag, strTipo_IcmsNFe, descricaoOperacao, modelo_Nfe.NFe_55, True, Tipo_Operacao)
                    If id <> "" Then
                            P_Status ""
                            frmNFE_Detalhe.tag = id
                            frmNFE_Detalhe.Show 1
                    End If
                        P_Status ""
                        Unload Me
                End If
        
        
                    
                
            End If
        End If
        
        
                
    
End Sub


Private Sub dtgrd_DblClick()
    editaItem
End Sub

Private Sub dtgrd_GotFocus()
    
    dtgrd.MarqueeStyle = 3
    
    
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo trata_Erro

    '    On Error GoTo trata_erro

    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyDelete Then
        
        If status = "2" Then
            'MsgBox "Não é possível excluir produtos em uma venda fechada.", vbExclamation
            'Exit Sub
        End If
                
                    If dtgrd.row = -1 Then
                        MsgBox "Selecione o item que deseja excluir!", vbCritical
                        dtgrd.SetFocus
                        Exit Sub
                    End If
                    
            Excluir False
            P_AtualizaVenda
    ElseIf KeyCode = vbKeyP Or KeyCode = vbKeyReturn Then
        txtbxProduto_GotFocus
        txtbxQtd.SetFocus
    End If
    
    Exit Sub
trata_Erro:
P_MostraErro

End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub


Private Sub dtgrd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        bDescancoTela = False
    iDescancoTela = 0

End Sub

Private Sub adc_Reposition()
    
        
End Sub


Sub buscaPromo()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
        
    strSQL = "select nome_pro,pv_pro from tab_produtos as tp,"
    strSQL = strSQL & " tab_prodpromo as tpp where tpp.cod_Pro=tp.cod_pro"
    strSQL = strSQL & " and datafin_pro >=" & F_GravaData(Date)
    strSQL = strSQL & " order by nome_pro"
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        QtdPromo = .RecordCount
        If .RecordCount > 0 Then
            ReDim pPromo(.RecordCount) As prodPromo
        End If
        strSQL = ""
        Do While Not .EOF
            pPromo(.AbsolutePosition).nome = !nome_pro
            pPromo(.AbsolutePosition).preco = "R$ " & FormatVal(!pv_pro)
            
            .MoveNext
        Loop
        .Close
        
        
        
    
    End With
    
End Sub
Private Sub Form_Activate()
'    On Error GoTo trata_erro
    If bActive Then Exit Sub
    mdiFrm.Enabled = False
    
    
        
        
        
    lblDesc.caption = crcDesconto_usu
    lblNomeEmp.caption = strFantasia_Emp
    lblnum_ven.caption = Me.tag
    txtbxVendedor.tag = strCod_Usu
    txtbxVendedor.text = strNome_Usu
    lblNomeEmp.caption = strFantasia_Emp
    lblCaixa = strCod_Cai
    lblTur.caption = strCod_Tur
    If bActive = True Then Exit Sub
    carrega
    
    
    P_Status "Buscando Itens..."
On Error Resume Next
    txtbxProduto.SetFocus
    P_buscaitens
    P_AtualizaVenda
    P_Status ""
    If CCur(lblTotalQueima.caption) > 0 Then
        lblTotalQueima.Visible = True
        lblTotalQueimaCaption.Visible = True
    Else
        lblTotalQueima.Visible = False
        lblTotalQueimaCaption.Visible = False
    End If
    
    bActive = True
    buscaLogo
    Form_Resize
    
    bActive = True
    Exit Sub
trata_Erro:
    P_MostraErro "Form_Activate"
    
End Sub
Sub playGif()
    'WebBrowser1.Navigate "c:\market\gifs\maebela2.gif"
    'WebBrowser1.Navigate "about:<html><body scroll='no'><img src='c:\maebela.gif'></img></body></html>"
    
End Sub
Sub buscaLogo()
    On Error Resume Next
    On Error GoTo trata_Erro
    pcLogo.Picture = LoadPicture(App.path & "\pic\logo.jpg")
    pcLogo.Picture = LoadPicture(App.path & "\pic\logovenda.jpg")
    pcLogo.Left = (pcLeft.width / 2) - (pcLogo.width / 2)
    Exit Sub
trata_Erro:
    
End Sub
    


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    bDescancoTela = False
    iDescancoTela = 0
    
    If KeyCode = vbKeyF3 Then
        editaItem
    ElseIf KeyCode = vbKeyF5 Then
        
        dtgrd.SetFocus
    ElseIf KeyCode = vbKeyF4 Then
         dtgrd_KeyDown vbKeyDelete, 0
    ElseIf KeyCode = vbKeyF6 Then
        cmndbtnCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        
        Concluir
    ElseIf KeyCode = vbKeyF2 Then
        ImprVenda Me.tag
    ElseIf KeyCode = vbKeyF8 Then
        pcDesconto.Visible = True
    ElseIf KeyCode = vbKeyF9 Then
        
    ElseIf KeyCode = vbKeyF10 Then
        chkEntrega_Click
    ElseIf KeyCode = vbKeyF1 Then
        If Me.ActiveControl.Name = txtbxNomeCli.Name Then
            cmndbtnFormasPag.SetFocus
        End If
        cmndbtnFormasPag_Click
    ElseIf KeyCode = vbKeyF12 Or KeyCode = vbKeyEscape Then
        cmndbtnSair_Click
    
    End If
End Sub
Sub mudamodo()
    Dim strSQL As String
    On Error Resume Next
    strSQL = "update tab_caixas set modo_cai='" & operacao & "' "
    strSQL = strSQL & " where cod_cai='" & strCod_Cai & "' "
    Conn.Execute strSQL
End Sub

Private Sub Form_Load()
'    buscaPromo
    bActive = False

    'adc.ConnectionString = Conn.ConnectionString

    lblNomeEmp.caption = strNome_Emp
'    lblProp.caption = F_BuscaUmCampo("msgvenda_par", "tab_parametros", "")
    txtbxNomeCli.tag = "00000"
    txtbxNomeCli.text = "VENDA A CONSUMIDOR"
        
        
        
        
    
    Set AdoVendas = New ADODB.Recordset
    
    buscaFormaAvista
    
    
    
'    cmbxBuscar.ListIndex = 0
    'pcECF.BackColor = vbGreen
    'pcECF.BackColor = vbYellow
    operacao = "C"
    'operacao = "B"
    Atacado = False 'mosta se a venda é normal ou atacado
End Sub
Sub buscaFormaAvista()
    Dim f As FormasPag.FormasPag
    
    f = FormasPag.F_BuscaFormasPag("01")
    If f.cod_For <> "" Then
        txtbxDescForma.text = f.desconto_for
        txtbxDescForma.tag = f.cod_For
        On Error Resume Next
        txtbxProduto.SetFocus
    End If
    
    
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
    dtgrd.width = Me.width - dtgrd.Left - 200
    
    
    dtgrd.height = pcStatusBar.top - dtgrd.top - 100
    lblNomeEmp.Left = Me.width - lblNomeEmp.width - 200
    'pcCli.Width = pcProd.Width
    
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).width = pcAtalhos.width / (pcBotao.Count)
        If iCount > 0 Then
            pcBotao(iCount).Left = pcBotao(iCount - 1).Left + pcBotao(iCount - 1).width + 50
        End If
        
    Next
 
    pcProd.width = dtgrd.width
    'pcCli.Width = pcProd.Width
    

End Sub

Sub editaItem()
    'On Error GoTo trata_Erro
    Exit Sub 'para terminar depois
    With AdoVendas
        cmndbtnInserir.tag = !seq_ite
        txtbxProduto.text = !desc_ite
        txtbxProduto.tag = !cod_pro
        txtbxUnit.text = !unit_ite
        txtbxQtd.text = !qtd_ite
        txtbxPQ_pro.text = !pq_itd
        txtbxDescReais.text = !desconto_ite
        txtbxDescReais_LostFocus
        txtbxQtd.SetFocus
        
    End With
    
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub

Private Sub optnCod_Click()
    txtbxProduto.SetFocus
End Sub

Private Sub optnDesc_Click()
    txtbxProduto.SetFocus
End Sub

Private Sub optnEnd_Click()
    
End Sub

Private Sub optnSeq_Click()
    P_buscaitens
End Sub









Sub P_foco(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BackColor = vbWhite
        
    Next
    pcBotao(Indice).BackColor = &HECE3D0
End Sub
Sub P_focoMouse(Indice As Integer)
    Dim iCount As Integer
    For iCount = 0 To pcBotao.Count - 1
        pcBotao(iCount).BorderStyle = 0
        
    Next
    pcBotao(Indice).BorderStyle = 1
'    pcbotao(Indice).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    AdoVendas.Close
End Sub

Private Sub lblBotao_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_Click(Index As Integer)
    pcBotao_Click Index
End Sub

Private Sub pcIco_GotFocus(Index As Integer)
    Sendkeys Chr(9)
End Sub



Private Sub pcBotao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If Index = 0 Then
            pcBotao(pcBotao.Count - 1).SetFocus
        Else
            pcBotao(Index - 1).SetFocus
        End If
        
    ElseIf KeyCode = vbKeyDown Then
        If Index = pcBotao.Count - 1 Then
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

Private Sub txtAniversario_GotFocus()
    P_SelecTexto txtAniversario
End Sub

Private Sub txtAniversario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtAniversario_LostFocus()
    On Error GoTo trata_Erro:
    If Len(txtAniversario.text) = 4 Then
        txtAniversario.text = Left(txtAniversario.text, 2) & "/" & Right(txtAniversario.text, 2)
    End If
    txtAniversario.text = Format(txtAniversario.text, "dd/mm")
    atualiza_venda
    Exit Sub
trata_Erro:
    txtAniversario.text = ""
End Sub

Private Sub txtbxDescForma_LostFocus()
    If CCur(txtbxDescForma.text) > CCur(lblDescontoMax.caption) Then
        MsgBox "O Desconto máximo permitido para o usuário com esta forma de pagamento é de " & lblDescontoMax.caption, vbInformation
        txtbxDescForma.text = lblDescontoMax.caption
        Exit Sub
    End If
End Sub

Private Sub txtbxDesconto_GotFocus()
    P_SelecTexto txtbxDesconto
End Sub

Private Sub txtbxDesconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDesconto_LostFocus()
    On Error Resume Next
    P_ValorLostFocus txtbxDesconto
    If CCur(txtbxDesconto.text) < 0 Then
        MsgBox "Desconto inválido!", vbExclamation
        txtbxDesconto.text = "0"
        Exit Sub
    End If
    
    
   txtbxDescPerc.text = FormatVal(CCur(txtbxDesconto.text) * 100 / CCur(lblValor.caption), "###,##0.0000")

    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.text))

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
    If CCur(txtbxDescPerc.text) < 0 Then
        MsgBox "Desconto inválido!", vbExclamation
        txtbxDescPerc.text = "0"
        Exit Sub
    End If
    
    
    P_ValorLostFocus txtbxDescPerc, "###,##0.0000"
    txtbxDesconto.text = FormatVal(CCur(txtbxDescPerc.text) / 100 * (CCur(lblValor.caption) - CCur(lblTotalQueima.caption)))
    lblFinalPrev.caption = FormatVal(CCur(lblValor.caption) - CCur(txtbxDesconto.text))
    
End Sub


Private Sub txtbxDescPercent_GotFocus()
    P_SelecTexto txtbxDescPercent
End Sub
Private Sub pcBotao_Click(Index As Integer)
    If Index = 7 Then
        cmndbtnSair_Click
    ElseIf Index = 6 Then
        Concluir
    ElseIf Index = 5 Then
        cmndbtnCancel_Click
    ElseIf Index = 4 Then
        dtgrd.SetFocus
    ElseIf Index = 3 Then
        cmndbtnExc_Prod_Click
    ElseIf Index = 2 Then
        editaItem
    ElseIf Index = 0 Then
        
        cmndbtnFormasPag_Click
    ElseIf Index = 0 Then

    ElseIf Index = 7 Then
        
    End If
End Sub

Private Sub txtbxDescPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescPercent_LostFocus()
    P_ValorLostFocus txtbxDescPercent
    
    If CCur(txtbxDescPercent.text) < 0 Then
        MsgBox "Desconto inválido!", vbExclamation
        txtbxDescPercent.text = "0"
        Exit Sub
    End If
    
    On Error GoTo trata_Erro
    txtbxDescReais.text = FormatVal(CCur(txtbxDescPercent.text) * CCur(txtbxUnit.text) * CCur(txtbxQtd.text) / 100)
    
    
    
    
    Exit Sub
trata_Erro:
    txtbxDescReais.text = "0,00"
End Sub

Private Sub txtbxDescReais_GotFocus()
    P_SelecTexto txtbxDescReais
End Sub

Private Sub txtbxDescReais_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxDescReais_LostFocus()
    On Error Resume Next
    P_ValorLostFocus txtbxDescReais
    If CCur(txtbxDescReais.text) < 0 Then
        MsgBox "Desconto inválido!", vbExclamation
        txtbxDescReais.text = "0"
        txtbxFinal.text = "0"
        Exit Sub
    End If
    
    'txtbxDescPercent.Text = FormatVal((CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text))
    txtbxDescPercent.text = FormatVal(CCur(txtbxDescReais.text) / CCur(txtbxQtd.text) / CCur(txtbxUnit.text) * 100)
    
    
    p_totalItem
    
    
End Sub

Private Sub txtbxFinal_GotFocus()
    P_SelecTexto txtbxFinal
End Sub

Private Sub txtbxFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxFinal_LostFocus
        
        insere
        txtbxFinal.text = "0,00"
        txtbxProduto.SetFocus
    End If
End Sub

Private Sub txtbxFinal_LostFocus()
    On Error GoTo trata_Erro
'    On Error Resume Next
    'txtbxDescReais.Text = FormatVal(CCur(txtbxUnit.Text) * CCur(txtbxQtd.Text) - CCur(txtbxFinal.Text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    
    
    
    txtbxDescReais.text = FormatVal((CCur(txtbxUnit.text) * CCur(txtbxQtd.text)) - CCur(txtbxFinal.text))
    
    'txtbxDescPercent.Text = FormatVal(CCur(txtbxDescReais.Text) / CCur(txtbxUnit.Text) * 100) / CCur(txtbxQtd.Text)
    txtbxDescReais_LostFocus
trata_Erro:
    
End Sub


Private Sub txtbxFormaPag_GotFocus()
    P_SelecTexto txtbxFormaPag
End Sub

Private Sub txtbxFormaPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnFormasPag_Click
    End If
End Sub

Private Sub txtbxNomeCli_GotFocus()
    P_SelecTexto txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnCliente_Click
    ElseIf KeyCode = vbKeyF2 Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli, -1
    If txtbxNomeCli.tag <> "" Then
        
    End If
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, txtTel.text
End Sub

Private Sub txtbxObs_GotFocus()
    
End Sub

Private Sub txtbxObs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxObs_LostFocus()
    
End Sub

Private Sub txtbxProduto_Change()
    cmndbtnInserir.tag = ""
    txtbxProduto.tag = ""
    
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub
Sub buscaProd()
    
    p = produtos.F_BuscaProdSaida(txtbxProduto.text)
    If p.cod_pro <> "" Then
    
        txtbxProduto.text = p.nome_pro
        txtbxProduto.tag = p.cod_pro
        txtbxUnit.text = FormatVal(p.pv_pro)
        txtbxUnit.tag = FormatVal(p.pv_pro)
        If Tipo_Operacao = venda Then
        
            txtbxUnit.Enabled = False
        Else
            txtbxUnit.Enabled = True
        End If
        lblfracionado.caption = p.fracionado_pro
        
        
        
        If p.pq_pro > 0 And lblTipoPag.tag = "01" Then
            lblQueima.Visible = True
            txtbxDescReais.Enabled = False
            txtbxDescPercent.Enabled = False
            txtbxPQ_pro.Visible = True
            lblPq_pro.Visible = True
            
            txtbxPQ_pro.text = FormatVal(p.pq_pro)
            txtbxFinal.Locked = True
            txtbxDescReais.text = FormatVal(p.pv_pro - p.pq_pro)
            txtbxPQ_pro.tag = FormatVal(p.pv_pro - p.pq_pro)
            txtbxDescReais_LostFocus
            
        Else
            lblQueima.Visible = False
            lblPq_pro.Visible = False
            txtbxPQ_pro.text = "0,00"
            txtbxPQ_pro.Visible = False
            txtbxDescReais.Enabled = True
            txtbxDescPercent.Enabled = True
            txtbxFinal.Locked = False
            
            On Error Resume Next
            txtbxDescPercent.text = CCur(txtbxDescForma.text) + CCur(p.sub.descmax_sub)
        End If
        
        
        
        txtbxQtd.Enabled = True
        txtbxQtd.SetFocus
        
    End If
    
    
End Sub
Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPos As Integer
    Dim p As produto
    Dim rPesq As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        buscaProd

    ElseIf KeyCode = vbKeyTab Then
        txtbxQtd.Enabled = True
    ElseIf KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    ElseIf KeyCode = vbKeyF6 Then
        'cmndbtnLocaliza_Click
    ElseIf KeyCode = vbKeyF11 Then
        
        
    End If
    
End Sub

Sub insere()
    Dim iPos As Integer
    Dim i As Integer
    Dim e As EstadoImp
    
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sTipoQtd As String
    Dim situacao_ven As String
    On Error GoTo trata_Erro
    If CCur(txtbxFinal.text) <= 0 Then
        MsgBox "Valor inválido!", vbExclamation
        Exit Sub
    End If
    
    If CCur(txtbxUnit.text) < 0 Then
        MsgBox "valor inválido!", vbExclamation
        Exit Sub
    End If
    
    
    If lblQueima.Visible = False And Tipo_Operacao <> tipo_operacao_venda.Devolucao Then
        If CCur(Round(txtbxDescPercent.text)) > CCur(lblDesc.caption) + CCur(txtbxDescForma.text) Then
            MsgBox "O desconto máximo permitido para o usuário é de " & CCur(lblDesc.caption) + CCur(txtbxDescForma.text) & "%", vbExclamation
            Exit Sub
        End If
    Else
        
    End If
    
    situacao_ven = F_BuscaUmCampo("situacao_ven", TAB_VENDAS, "num_ven='" & Me.tag & "'")
    
    If situacao_ven = "3" Or situacao_ven = "4" Then
        MsgBox "Não é possível adicionar mais produtos, o status dessa venda não permite!", vbExclamation
        Exit Sub
    End If
    If situacao_ven = "2" Then
        If MsgBox("Essa venda já se encontra fechada, deseja realmente lançar mais produtos?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    If txtbxProduto.text = "" Then
        MsgBox "Informe pelo menos parte do nome do produto para que seja localizado!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    
                
        If cmndbtnInserir.tag = "" Then
            iItem = iItem + 1
        Else
            
        End If
                
                
                If Vendas.F_AdicionaItem(Me.tag, p, txtbxQtd.text, txtbxDescReais.text, txtbxFinal.text, txtbxPQ_pro.text) = "" Then
                    MsgBox "Erro na inserção do item!", vbExclamation
                    
                Else
                    'If Not Produtos.F_AlteraEstoquePro(p.cod_pro, "-" & txtbxQtd.Text) Then
                    '    MsgBox "Erro na baixa do estoque.", vbExclamation
                    'End If
                    
                    
                End If
                    
                P_buscaitens
                Me.Refresh
                
                txtbxQtd.text = "1"
                'AdoVendas.Move AdoVendas.RecordCount - 1
                txtbxUnit.text = "0"
                txtbxDescPercent.text = "0,0"
                txtbxDescReais.text = "0,00"
                txtbxProduto_GotFocus
                txtbxProduto.SetFocus
                
        
        P_AtualizaVenda
        On Error Resume Next

        Exit Sub
trata_Erro:
    P_MostraErro "Insere()"
End Sub
Private Sub txtbxQtd_GotFocus()
    
    P_SelecTexto txtbxQtd
End Sub

 Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
Exit Sub
trata_Erro:
    P_MostraErro
End Sub

Private Sub txtbxQtd_LostFocus()
    On Error Resume Next
    If lblfracionado = "S" Then
        P_ValorLostFocus txtbxQtd, "###,##0.00"
    Else
        P_ValorLostFocus txtbxQtd, "###"
    End If
    If txtbxPQ_pro.Visible = True Then
        txtbxDescReais.text = FormatVal(CCur(txtbxPQ_pro.tag) * CCur(txtbxQtd.text))
        txtbxDescReais_LostFocus
    End If
    p_totalItem
    
        
    
End Sub


Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit, "###,##0.00"
    
    p_totalItem
End Sub
Sub P_GravaItemCancelado()
    On Error Resume Next
    Dim strSQL As String
    strSQL = "insert into tab_itvndcncl(seq_itv,num_ven,seq_ite,qtd_ite,unit_ite,cod_cai,cod_tur)values("
    strSQL = strSQL & "'" & MaxKeySQL("tab_itvndcncl", "seq_itv", "", "000000#") & "',"
    strSQL = strSQL & "'" & Me.tag & "',"
    strSQL = strSQL & "" & AdoVendas!seq_ite & ","
    strSQL = strSQL & F_GravaDec(AdoVendas!qtd_ite) & ","
    strSQL = strSQL & F_GravaDec(AdoVendas!unit_ite) & ","
    strSQL = strSQL & "'" & strCod_Cai & "',"
    strSQL = strSQL & "'" & strCod_Tur & "') "
    Conn.Execute strSQL
    
    
    
End Sub




Private Sub txtCPF_GotFocus()
    P_SelecTexto txtCpf
End Sub

Private Sub txtCPF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtCPF_LostFocus()
    txtCpf.text = F_SoNumero(txtCpf.text)
    If txtCpf.text <> "" Then
        If Len(txtCpf.text) > 11 Then
            If Not F_CGC(txtCpf.text) Then
                MsgBox "CPF/CNPJ Inválido", vbExclamation
            End If
        Else
            If Not F_CPF(txtCpf.text) Then
                MsgBox "CPF Inválido", vbExclamation
            End If
        End If
        
    
    
    End If
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, txtTel.text, txtCpf.text
End Sub
Sub atualiza_venda()
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, txtTel.text, txtCpf.text, txtEmail.text, txtAniversario.text
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
    P_AjustaTexto txtEmail, 3
    atualiza_venda
    
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
    txtTel.text = F_SoNumero(txtTel, True)
    P_AjustaTexto txtTel
    
    
    
    
    If AdoVendas.RecordCount > 0 Then
    
        
    End If
    
    P_AtualizaVendaCliente Me.tag, txtbxNomeCli.tag, txtbxNomeCli.text, txtTel.text
    
End Sub

Private Sub txtValorEntrega_GotFocus()
    P_SelecTexto txtValorEntrega
End Sub

Private Sub txtValorEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys Chr(9)
    End If
End Sub

Private Sub txtValorEntrega_LostFocus()
    P_ValorLostFocus txtValorEntrega
    P_AtualizaVenda
End Sub
