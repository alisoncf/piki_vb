VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiPro 
   Caption         =   "Edição de Produtos"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   17190
   Icon            =   "frmEdiPro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   17190
   Begin VB.PictureBox pcBotoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00B0B0B0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   12630
      TabIndex        =   51
      Top             =   9090
      Width           =   12630
      Begin VB.CommandButton cmdReservadas 
         Caption         =   "Abrir Reservadas"
         Height          =   615
         Left            =   9480
         Picture         =   "frmEdiPro.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   615
         Left            =   6360
         Picture         =   "frmEdiPro.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnHtml 
         Caption         =   "&Imprimir Lista"
         Height          =   615
         Left            =   4800
         Picture         =   "frmEdiPro.frx":0F43
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSimilares 
         Caption         =   "Similares"
         Height          =   615
         Left            =   7920
         Picture         =   "frmEdiPro.frx":1350
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdHist 
         Caption         =   "&Histórico"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiPro.frx":18DA
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "Fe&char (Esc)"
         Height          =   615
         Left            =   11040
         Picture         =   "frmEdiPro.frx":1A24
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiPro.frx":1B8F
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiPro.frx":1CD9
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox pcPreco 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   11880
      ScaleHeight     =   975
      ScaleWidth      =   2895
      TabIndex        =   20
      Top             =   9000
      Width           =   2895
      Begin VB.PictureBox pcPreco1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   0
         Picture         =   "frmEdiPro.frx":2063
         ScaleHeight     =   975
         ScaleWidth      =   2895
         TabIndex        =   21
         Top             =   0
         Width           =   2895
         Begin VB.Label lblEr 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   2280
            TabIndex        =   25
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblPv 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   960
            TabIndex        =   24
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Preço(R$):"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   240
            TabIndex        =   23
            Top             =   120
            Width           =   780
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Estoque:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1920
            TabIndex        =   22
            Top             =   120
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox pcBusca 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   11400
      TabIndex        =   1
      Top             =   6480
      Width           =   11430
      Begin VB.PictureBox pcInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   945
         ScaleWidth      =   11160
         TabIndex        =   27
         Top             =   1320
         Width           =   11190
         Begin VB.CommandButton cmdCli 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2880
            Picture         =   "frmEdiPro.frx":C16D
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   390
            Width           =   375
         End
         Begin VB.CommandButton cmdVei 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   5880
            Picture         =   "frmEdiPro.frx":C518
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   390
            Width           =   375
         End
         Begin VB.PictureBox pcSimilares 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   6480
            Picture         =   "frmEdiPro.frx":C8C3
            ScaleHeight     =   1695
            ScaleWidth      =   4830
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   4830
            Begin VB.CommandButton cmdAbrirSim 
               Caption         =   "Abrir"
               Height          =   435
               Left            =   1200
               TabIndex        =   33
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdListSim 
               Caption         =   "Listar similares (S)"
               Height          =   435
               Left            =   2040
               TabIndex        =   32
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label10 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Produtos Similares"
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
               Left            =   240
               TabIndex        =   35
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label lblEstoqueSim 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   600
               TabIndex        =   34
               Top             =   480
               Width           =   135
            End
         End
         Begin VB.TextBox txtCliente 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   30
            Top             =   390
            Width           =   2775
         End
         Begin VB.PictureBox pcImg 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   10320
            Picture         =   "frmEdiPro.frx":18905
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox txtVeiculo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3480
            TabIndex        =   28
            Top             =   390
            Width           =   2415
         End
         Begin VB.Label lblDescontoCli 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   1320
            TabIndex        =   41
            Top             =   120
            Width           =   90
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente (Ctrl + L):"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label lblVeic 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Veículo (Alt + V):"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            TabIndex        =   39
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label lblFab 
            AutoSize        =   -1  'True
            Caption         =   "Fabricante"
            Height          =   195
            Left            =   5520
            TabIndex        =   38
            Top             =   120
            Width           =   750
         End
      End
      Begin VB.TextBox txtDetalhe 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiPro.frx":18D7B
         Left            =   9840
         List            =   "frmEdiPro.frx":18D8B
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtbxExpr 
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
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   11
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtCat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkSimilares 
         Caption         =   "Similares"
         Height          =   195
         Left            =   14040
         TabIndex        =   8
         Top             =   -240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkInvisiveis 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Listar invisíveis"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7800
         TabIndex        =   7
         Tag             =   "!"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Listar zeradas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7800
         TabIndex        =   6
         Tag             =   "!"
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkAtacado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Atacado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9480
         TabIndex        =   5
         Tag             =   "!"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtOrigem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkComb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Apenas Combustíveis"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9480
         TabIndex        =   2
         Tag             =   "!"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdEmpresa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9480
         Picture         =   "frmEdiPro.frx":18DB8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdMarca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7440
         Picture         =   "frmEdiPro.frx":19163
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   3480
         TabIndex        =   18
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   9840
         TabIndex        =   17
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar por: (Alt + P)"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   7800
         TabIndex        =   15
         Top             =   120
         Width           =   660
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiPro.frx":1950E
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Lista de Produtos"
      ColumnCount     =   20
      BeginProperty Column00 
         DataField       =   "sinal"
         Caption         =   "*"
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
         DataField       =   "endereco_pro"
         Caption         =   "Endereço"
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
         DataField       =   "cod_pro"
         Caption         =   "Código"
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
         DataField       =   "codfab_pro"
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
      BeginProperty Column04 
         DataField       =   "nome_pro"
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
      BeginProperty Column05 
         DataField       =   "pv_pro"
         Caption         =   "Preço"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "estoque"
         Caption         =   "Em estoque"
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
         DataField       =   "multiplo_pro"
         Caption         =   "Múltiplo"
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
         DataField       =   "nome_mar"
         Caption         =   "Marca"
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
         DataField       =   "nome_cat"
         Caption         =   "Categoria"
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
         DataField       =   "ean_pro"
         Caption         =   "Ean"
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
      BeginProperty Column11 
         DataField       =   "cod_emp"
         Caption         =   "Emp"
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
      BeginProperty Column12 
         DataField       =   "estmin_pro"
         Caption         =   "Min"
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
      BeginProperty Column13 
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
      BeginProperty Column14 
         DataField       =   "genero"
         Caption         =   "Gênero"
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
      BeginProperty Column15 
         DataField       =   "subsecao"
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
      BeginProperty Column16 
         DataField       =   "secao"
         Caption         =   "secao"
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
      BeginProperty Column17 
         DataField       =   "visivel_pro"
         Caption         =   "Habilitada"
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
      BeginProperty Column18 
         DataField       =   "cest_Pro"
         Caption         =   "CEST"
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
      BeginProperty Column19 
         DataField       =   "reservado_pro"
         Caption         =   "Reservado"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   255,118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1335,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   5669,858
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pcRight 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9915
      Left            =   11895
      ScaleHeight     =   9885
      ScaleWidth      =   5265
      TabIndex        =   42
      Top             =   0
      Width           =   5295
      Begin VB.PictureBox pcContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   5025
         TabIndex        =   60
         Top             =   4200
         Width           =   5055
         Begin VB.PictureBox pcImagem 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   240
            ScaleHeight     =   1695
            ScaleWidth      =   1575
            TabIndex        =   61
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid gridVeiculo 
         Bindings        =   "frmEdiPro.frx":19520
         Height          =   2055
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "Veículos"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "veiculo"
            Caption         =   "Veiculo"
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
            DataField       =   "aplicacao"
            Caption         =   "Aplicação"
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
            DataField       =   "fabricante"
            Caption         =   "Fabricante"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   1635,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2009,764
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridConversao 
         Bindings        =   "frmEdiPro.frx":19532
         Height          =   1935
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "Conversão"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "nome_mar"
            Caption         =   "Marca"
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
            DataField       =   "referencia"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2129,953
            EndProperty
         EndProperty
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Veiculo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   4200
         Width           =   525
      End
      Begin VB.Label lblVeiculo 
         AutoSize        =   -1  'True
         Caption         =   "Veículo"
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
         Left            =   720
         TabIndex        =   49
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   48
         Top             =   4200
         Width           =   795
      End
      Begin VB.Label lblFabricante 
         AutoSize        =   -1  'True
         Caption         =   "Fabr"
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
         Left            =   3840
         TabIndex        =   47
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicação:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   4440
         Width           =   750
      End
      Begin VB.Label lblAplicacao 
         Caption         =   "Aplicação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   960
         TabIndex        =   45
         Top             =   4440
         Width           =   4185
      End
   End
   Begin VB.Menu mnOpcoes 
      Caption         =   "&Opções"
      Begin VB.Menu mnEditMult 
         Caption         =   "Editar Múltiplos"
         Begin VB.Menu mnNCM 
            Caption         =   "NCM"
         End
         Begin VB.Menu mnPreco 
            Caption         =   "Preço"
         End
         Begin VB.Menu mnMargem 
            Caption         =   "Margem de Lucro"
         End
         Begin VB.Menu mnCEST 
            Caption         =   "Mudar CEST"
         End
         Begin VB.Menu mnOpcoesNfe 
            Caption         =   "Opções Nfe"
         End
         Begin VB.Menu mnMudarICMS 
            Caption         =   "Mudar ICMS"
         End
         Begin VB.Menu mnTornarInv 
            Caption         =   "Tornar Invisível"
         End
      End
      Begin VB.Menu mnImg 
         Caption         =   "Abrir Imagens"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnEstoqueReservado 
         Caption         =   "Ajustar estoque reservado"
      End
      Begin VB.Menu mnAlterarPrecoLote 
         Caption         =   "Alterar Preços em Lote"
      End
   End
End
Attribute VB_Name = "frmEdiPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opvis As String
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Public WithEvents adoveiculo As ADODB.Recordset
Attribute adoveiculo.VB_VarHelpID = -1
Public WithEvents adoConversao As ADODB.Recordset
Attribute adoConversao.VB_VarHelpID = -1
Dim bActive As Boolean
Dim listaImagens As New Collection
Sub BuscaUltEntrada()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select tc.num_not,entrada_not,tf.nome_for,ti.* "
    strSQL = strSQL & " from tab_cmp as tc, tab_itcmp as ti, "
    strSQL = strSQL & " " & TAB_FORNECEDORES & " as tf "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    strSQL = strSQL & " and tc.seq_not=ti.seq_not "
    strSQL = strSQL & " and ti.cod_pro='" & ado!cod_pro & "' "
    strSQL = strSQL & " order by entrada_not desc limit 1 "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            lblDataUltEntrada.caption = ""
            lblNomeFor.caption = ""
            lblNumNot.caption = ""
            lblQtd.caption = ""
            lblUnit.caption = ""
            
        Else
            lblDataUltEntrada.caption = Format(!entrada_not)
            lblNomeFor.caption = !nome_For
            lblNumNot.caption = Format(!Num_not)
            lblQtd.caption = Format(CCur(!qtd_ite) * CCur(!qtd_emb))
            lblUnit.caption = Format(!unit_ite)
            
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
End Sub

Sub P_TileMsg(msg As String)

End Sub
Private Sub chkEst_Click()
    busca
End Sub

Private Sub adc_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub ado_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If Not bActive Then Exit Sub
   
End Sub

Private Sub ado_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
   If bActive = False Then Exit Sub
    With ado
        If !visivel_pro = "S" Then
            cmdVisivel.caption = "Tornar Invisível"
            opvis = "N"
        Else
            cmdVisivel.caption = "Tornar Visível"
            opvis = "S"
        End If
        VerSim !cod_pro
        lblEstoqueSim.caption = BuscaEstoqueSim(!cod_pro)
        lblPv.caption = FormatVal(ado!pv_pro - (ado!pv_pro * CCur(lblDescontoCli.caption / 100)))
        lblEr.caption = ado!Estoque
        
        If TiraNull(!detalhe_pro) <> "" Then
            txtDetalhe.Text = !detalhe_pro
            txtDetalhe.Visible = True
        Else
            txtDetalhe.Text = ""
            txtDetalhe.Visible = False
        End If
        
        
        
        If bControlaEstoque Then
            If !controlaestoque_pro = "S" Then
                lblEr.Visible = True
            Else
                lblEr.Visible = False
                
            End If
            
            If !reservado_pro > 0 Then
                cmdReservadas.Visible = True
            Else
                cmdReservadas.Visible = False
            End If
             
            buscaVei
            buscaConversao
        Else
            
            
        End If
    End With
    BuscaImg
    
End Sub

Private Sub ado_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    VerSim ado!cod_pro
    lblPv.caption = FormatVal(ado!pv_pro)
    If bControlaEstoque Then
        lblEr.caption = ado!er_pro
    End If
    BuscaImg
End Sub


Sub fitImage(pic As PictureBox, container As PictureBox, Arquivo As String)
    Dim i As Integer
    Dim h As Single
    Dim w As Single
    Dim p As Single
    Dim horizontal As Single
    Dim vertical As Single
    Dim proporcao As Single
    Dim img
    Dim picTemp As PictureBox
    
    
    
    
    On Error GoTo trata_Erro
    
    Set picTemp = pic
    
    pic.AutoSize = True
    Set img = pic.Picture
    h = img.height
    w = img.width
    
    p = h / 10
    
    
    If w > container.ScaleWidth Then
    
        For i = w To container.width Step -p
            horizontal = i
            
            
        Next
        proporcao = i / w
    
    End If
    h = h * proporcao
    
    If h > container.ScaleHeight Then
        For i = h To container.height Step -p
            vertical = i - p
            
            
        Next
        proporcao = i / w
    End If
    pic.Cls
    pic.AutoSize = False
    pic.AutoRedraw = True
    picTemp.PaintPicture img, 0, 0, img.width * proporcao, img.height * proporcao
    pic.Top = -1500
    Set pic.Picture = picTemp.image
    
    Exit Sub
    
    
    
    
trata_Erro:
End Sub
Sub BuscaImg()
    On erro GoTo trata_Erro
    Dim img As Picture
  
    pcImagem.Cls
    
    If Not IsNull(ado!caminho) Then
          
        Set listaImagens = F_SeparaTexto(ado!caminho, ",", False)
        pcImg.Visible = True
        pcImagem.Cls
        
        
        
        'mostraMiniatura  ado!caminho, pcRight, 2000, 600, pcRight.width, pcRight.height
        'Set img = LoadPicture(listaImagens(1))
        'pcImagem.PaintPicture LoadPicture(listaImagens(1)), 0, 0
        'pcImagem.tag = listaImagens(1)
        'pcImagem.Visible = True
        
        carregaImagemFit pcImagem, pcContainer, listaImagens(1)
        pcContainer.Visible = True
        'Me.SetFocus
        
        'fitImage pcImagem, pcContainerImg, listaImagens(1)
    Else
        pcImg.Visible = False
        pcImagem.Cls
        pcImagem.Visible = False
        pcContainer.Visible = False
    End If
trata_Erro:
End Sub
Private Sub chkAtacado_Click()
    busca
End Sub

Private Sub chkComb_Click()
    busca
End Sub

Private Sub chkInvisiveis_Click()
    busca
End Sub

Private Sub chkTodas_Click()
    busca
End Sub

Private Sub cmbxBusca_Click()
    'If Me.Visible = True Then busca
End Sub

Private Sub cmbxDiv_Click()
    lblDiv.caption = F_BuscaUmCampo("Cod_div", "tab_divisoes", "nome_div='" & cmbxDiv.Text & "'")
    busca
End Sub

Private Sub cmbxDiv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxDiv.ListIndex = -1
    End If
End Sub

Private Sub cmbxBusca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxCat.ListIndex = -1
    End If
End Sub



Private Sub cmbxCriterio_Click()
    busca
End Sub

Private Sub cmbxICMS_Click()
    busca
End Sub

Private Sub cmbxICMS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxICMS.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_Click()
    lblMar.caption = F_BuscaUmCampo("Cod_mar", TAB_MARCAS, "nome_mar='" & cmbxMarca.Text & "'")
    busca
End Sub

Private Sub cmbxEsp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxEsp.ListIndex = -1
    End If
End Sub

Private Sub cmbxCat_Click()
    lblCat.caption = F_BuscaUmCampo("Cod_Cat", "tab_categorias", "nome_cat='" & cmbxCat.Text & "'")
    
    busca
End Sub

Private Sub cmbxGenero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmbxGenero.ListIndex = -1
    End If
End Sub

Private Sub cmbxMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
    If KeyCode = vbKeyDelete Then
        cmbxMarca.ListIndex = -1
    End If
End Sub

Private Sub cmbxOrd_Click()
    'If Me.Visible = True Then busca
End Sub

Private Sub cmbxOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxSt_Click()
    busca
End Sub

Private Sub cmbxSt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmdAbrirSim_Click()
    AbrirSimilares ado!cod_pro
End Sub

Private Sub opcoesNfe()
     If strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    Dim strSQL As String
    frmMudarDadosSPEDNFe.Show 1
    If frmMudarDadosSPEDNFe.tag <> "" Then
        If MsgBox("Deseja realmente mudar as opções selecionadas para os " & ado.RecordCount & " produtos listados?", vbYesNo + vbCritical) = vbNo Then
            Unload frmMudarDadosSPEDNFe
            Unload Me
        End If
        With ado
            If .RecordCount > 0 Then
                .MoveFirst
            End If
            Do While Not .EOF
                With frmMudarDadosSPEDNFe
                    AlteraOpcoesNfe ado!cod_pro, .txtGenero.tag, .txtOrigem.tag, .txtbxICMS.tag, .txtCst_Icms.tag, .txtCst_Ipi.tag, .txtCst_Pis.tag, .txtCst_Cofins.tag, ""
                End With
                .MoveNext
            Loop
        End With
    End If
    Unload frmMudarDadosSPEDNFe
    
    
End Sub

Private Sub cmdCat_Click()
    buscaCat
End Sub
Sub buscaCat()
    Dim c As categoria.categoria
    c = categoria.BuscaCategoria(txtCat.Text)
    If c.codigo <> "" Then
        txtCat.Text = c.nome
        txtCat.tag = c.codigo
        busca
        TrocaTabKey 13
    End If
End Sub

Private Sub mudarCest()
    Dim sNCM As String
    On Error GoTo trata_Erro
    Dim strSQL As String
     If strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    sNCM = InputBox("Informe o número CEST para todos os produtos listados sem pontos.(apenas dígitos)", "NCM")
    If sNCM = "" Then
        MsgBox "Não foi informado um CEST válido!", vbExclamation
        Exit Sub
    End If
    
    With ado
        If MsgBox("Confirma alteração do código CEST: " & sNCM & " para os " & .RecordCount & " produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
            .MoveFirst
        
            Do While Not .EOF
                strSQL = "update " & TAB_PRODUTOS & " set cest_pro='" & sNCM & "' "
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                conn.Execute strSQL
                
                .MoveNext
            Loop
            .Requery
            
        End If
        
        
    End With
    
    Exit Sub
trata_Erro:
    MostraErro

End Sub

Private Sub cmdCli_Click()
    buscaCli
End Sub

Private Sub cmdEmpresa_Click()
    buscaEmpresa
End Sub
Sub buscaEmpresa()
    Dim e As sped_empresa.sped_empresa
    
    e = sped_empresa.SelecionarSped_Empresa(txtEmpresa.Text)
    If e.codigo <> "" Then
        txtEmpresa.Text = e.nome
        txtEmpresa.tag = e.codigo
        busca
    End If
        
    
    
End Sub
Private Sub cmdExportar_Click()
    ExportaGridToCSV dtgrd, ado
End Sub

Private Sub cmdHist_Click()
    On Error Resume Next
    P_HistProd ado!cod_pro
End Sub

Private Sub cmdListSim_Click()
    On Error GoTo trata_Erro:
    If Not IsNull(ado!cod_pro) Then
        busca True
    End If
    Exit Sub
trata_Erro:
End Sub


Sub margemLucro()
     If strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    Dim s As String
    s = InputBox("INforme a  margem de lucro desejada")
    If s = "" Then
        Exit Sub
    End If
    
    With ado
        If MsgBox("Confirma alteração da margem de lucro para " & s & "% para os " & .RecordCount & " produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
           P_Status "ajustando margem de lucro..."
            .MoveFirst
            dtgrd.Visible = False
        
            Do While Not .EOF
                strSQL = "update " & TAB_PRODUTOS & " set lucro_pro=" & F_GravaDec(s) & " "
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                conn.Execute strSQL
                
                .MoveNext
            Loop
            .Requery
            
        End If
        dtgrd.Visible = True
        
    End With
    P_Status ""
    MsgBox "Concluído", vbInformation
End Sub

Private Sub mudarNcm()
    Dim sNCM As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    
    If strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    
    sNCM = InputBox("Informe o número NCM para todos os produtos listados sem pontos.(apenas dígitos)", "NCM")
    If sNCM = "" Then
        MsgBox "Não foi informado um NCM válido!", vbExclamation
        Exit Sub
    End If
    
    P_Status "aguarde..."
    With ado
        If MsgBox("Confirma alteração do código NCM: " & sNCM & " para os " & .RecordCount & " produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
            .MoveFirst
        
            Do While Not .EOF
                strSQL = "update " & TAB_PRODUTOS & " set ncm_pro='" & sNCM & "' "
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                conn.Execute strSQL
                
                .MoveNext
            Loop
            .Requery
            
        End If
        
        
    End With
    
    Exit Sub
trata_Erro:
P_Status ""
    MostraErro
End Sub

Private Sub cmdOrigem_Click()
    buscaOrigem
End Sub
Sub buscaOrigem()
    Dim p As Padrao1
    p = Sped_Origem(txtOrigem.Text)
    If p.codigo <> "" Then
        txtOrigem.Text = p.nome
        txtOrigem.tag = p.codigo
        busca
    End If
End Sub

Sub MudarPreco()
     Dim sPreco As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    If strAltpreco_usu <> "S" Or strAltprod_usu <> "S" Then
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        Exit Sub
    End If
    sPreco = InputBox("Informe o valor do reajuste de preço para todos os produtos (%)", "Alterar Preço")
    If sPreco = "" Then
        MsgBox "Não foi informado um valor válido!", vbExclamation
        Exit Sub
    End If
    sPreco = FormatVal(sPreco)
    If CCur(sPreco) >= 15 Then
        MsgBox "Não foi informado um valor válido!", vbExclamation
        Exit Sub
    End If
    If CCur(sPreco) <= -15 Then
        MsgBox "Não foi informado um valor válido!", vbExclamation
        Exit Sub
    End If
    
    dtgrd.Visible = False
    With ado
        If MsgBox("Confirma alteração do preço em  " & sPreco & "% para os " & .RecordCount & " produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
            bActive = False
            .MoveFirst
                
            Do While Not .EOF
                'If CCur(sPreco) > 0 Then
                    strSQL = "update " & TAB_PRODUTOS & " set pv_pro=pv_pro+(pv_pro*" & F_GravaDec(CCur(sPreco / 100)) & ") "
                    strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                    conn.Execute strSQL
                'Else
                '    strSQL = "update " & TAB_PRODUTOS & " set pv_pro=pv_pro-(pv_pro*" & F_GravaDec(CCur(sPreco / 100)) & ") "
                '    strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                '    conn.Execute strSQL
                'End If
                P_Status "Atualização " & .AbsolutePosition & "/" & .RecordCount
                
                .MoveNext
            Loop
            dtgrd.Visible = True
            .Requery
            
        End If
        
        
    End With
    dtgrd.Visible = True
    Exit Sub
trata_Erro:
dtgrd.Visible = True
    MostraErro
End Sub

Private Sub cmdVisivel_Click()
    Dim opcao As String
    With ado
        If MsgBox("Confirma alteração da visibilidade para os " & .RecordCount & " produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
            .MoveFirst
            
            Do While Not .EOF
                strSQL = "update " & TAB_PRODUTOS & " set visivel_pro='" & opvis & "' "
                strSQL = strSQL & " where cod_pro='" & !cod_pro & "'"
                conn.Execute strSQL
                
                .MoveNext
            Loop
            .Requery
            
        End If
        
        
    End With
End Sub

Private Sub cmndbtnCodBarra_Click()
    Dim fr As New frmGeraBarCode
    P_PreencheBarCode ado!cod_pro, fr.lstBarCode
    fr.txtbxDesc.Text = ado!nome_pro
    fr.txtbxPreco.Text = FormatVal(ado!pv_pro)
    fr.Show 0
End Sub

Private Sub deletePro(Optional bMsg As Boolean = True)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim bDeleta As Boolean
On Error GoTo trata_Erro
    
    
    If bMsg Then
        If Not F_senha Then Exit Sub
        If strExclVenda_usu <> "S" Then
            MsgBox "Você não tem permissão!", vbExclamation
        End If
    End If
    bDeleta = True
    With ado
        If bMsg Then
            If MsgBox("Confirma exclusão do Produto?", vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
        End If
        If .RecordCount = 0 Then
            
                Exit Sub
            
        End If
        
        strSQL = "select emissao_not,tc.seq_not from tab_cmp as tc,tab_itcmp as ti where tc.seq_not=ti.seq_not "
        strSQL = strSQL & " and ti.cod_pro='" & !cod_pro & "' "
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.RecordCount > 0 Then
            bDeleta = False
            If bMsg Then
                MsgBox "Há uma compra para este produto. SEQ: " & rPesq!seq_not & " Emissão: " & rPesq!emissao_not, vbCritical
            
                Exit Sub
            End If
        End If
        strSQL = "select emissao_not,tc.seq_not from tab_compras as tc,tab_itcompras as ti where tc.seq_not=ti.seq_not "
        strSQL = strSQL & " and ti.cod_pro='" & !cod_pro & "' "
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.RecordCount > 0 Then
            bDeleta = False
            If bMsg Then
                MsgBox "Há uma compra para este produto. SEQ: " & rPesq!seq_not & " Emissão: " & rPesq!emissao_not, vbCritical
                Exit Sub
            End If
        End If
        
        sAux = F_BuscaUmCampo("cod_pro", "tab_itsai", "cod_pro = '" & !cod_pro & "'")
        
        
        If sAux <> "" Then
            bDeleta = False
            If bMsg Then
                MsgBox "Existem vendas com esse produto!" & Chr(13) & "Não será possível excluir o produto selecionado.", vbCritical
                Exit Sub
            End If
        End If
        
        
        If bDeleta Then
            strSQL = "delete from " & TAB_PRODUTOS & " where cod_pro = '" & !cod_pro & "'"
            conn.Execute strSQL
            strSQL = "delete from estoque where cod_pro = '" & !cod_pro & "'"
            conn.Execute strSQL
            strSQL = "delete from tab_embalagens where cod_pro = '" & !cod_pro & "'"
            conn.Execute strSQL
            strSQL = "delete from tab_codbarras where cod_pro = '" & !cod_pro & "'"
            conn.Execute strSQL
        End If
    End With
        If bMsg Then
        adc.Refresh
        MsgBox "Produto Excluído!", vbExclamation
    End If
Exit Sub
trata_Erro:
MostraErro
    
End Sub

Private Sub cmndbtnDelPro_Click()
    bActive = False
        deletePro True
    bActive = True
End Sub

Sub procurado()
    On Error GoTo trata_Erro
    AbreProdutoProcurado ado!cod_pro
    
    Exit Sub
trata_Erro:
End Sub

Private Sub cmdReservado_Click()
    On Error GoTo trata_Erro
    P_AjustaEstoqueReservado ado!cod_pro
    Exit Sub
trata_Erro:
    
End Sub

Private Sub cmdReservadas_Click()
    frmEdiVendas.txtProd.Text = ado!nome_pro
    frmEdiVendas.txtProd.tag = ado!cod_pro
    frmEdiVendas.txtbxDe.Text = ""
    frmEdiVendas.txtbxAte.Text = ""
    frmEdiVendas.tag = "x"
    frmEdiVendas.Show 1
End Sub

Private Sub cmdSimilares_Click()
    On Error Resume Next
    
    AbrirSimilares ado!cod_pro
End Sub

Private Sub cmdVei_Click()
    selecionaVei
End Sub

Private Sub cmndbtnEdit_Click()
    On Error Resume Next
    
    If strCadprod_usu <> "S" Then
 '       MsgBox "Você não tem permissão para esta operação!", vbExclamation
'        Exit Sub
    End If
    
    frmCadPro.tag = ado!cod_pro
    frmCadPro.Show 1
    RefreshAdoPosition ado
    
    Exit Sub
    Dim fr As New frmCadPro
    
    
    If dtgrd.row = -1 Then
    Exit Sub
    End If
    
    
End Sub

Private Sub cmndbtnFornadas_Click()
    frmDadosFornada.tag = ado!cod_pro
    frmDadosFornada.caption = "Fornadas: " & _
    Format(ado!cod_pro, "00#\.###\-# ") & ado!nome_pro
    frmDadosFornada.frControle.caption = ado!nome_pro
    On Error Resume Next
    frmDadosFornada.Show 1
End Sub

Private Sub cmndbtnHtml_Click()
    P_DtgrdToHtml "LISTAGEM DE PRODUTOS", dtgrd, ado, 0
    
    
End Sub


Private Sub cmndbtnNew_Click()
    If strCadprod_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    Dim fr As New frmCadPro
    fr.Show 0
        
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdMarca_Click()
    buscaMarca
End Sub
Sub buscaMarca()
    Dim c As Marca.Marca
    c = Marca.buscaMarca(txtMarca.Text)
    If c.codigo <> "" Then
        txtMarca.Text = c.nome
        txtMarca.tag = c.codigo
        busca
        TrocaTabKey 13
    End If
End Sub
Sub deletarTodos()
    If Not F_senha Then
        Exit Sub
    End If
    
    If strNome_Usu <> "ALISON" Then
        Exit Sub
    End If
    
    If MsgBox("Deseja deletar todos os produtos listados? " & ado.RecordCount, vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    With ado
        dtgrd.Visible = False
        Do While Not .EOF
            
            deletePro False
            
            If .AbsolutePosition Mod 20 = 0 Then
                P_Status .AbsolutePosition & "?" & .RecordCount
                mdiFrm.lblMsg.Refresh
            End If
            .MoveNext
        Loop
        .Requery
    End With
    P_Status ""
    dtgrd.Visible = True
End Sub



Private Sub dtgrd_DblClick()
    
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        dtgrd.col = 1
        Clipboard.SetText dtgrd.Text
    ElseIf KeyCode = vbKeyS Then
        busca True
        
    ElseIf KeyCode = vbKeyO Then
        busca False
    
    ElseIf KeyCode = vbKeyB Or KeyCode = vbKeyN Or KeyCode = vbKeyM Or _
                KeyCode = vbKeyG Or KeyCode = vbKeyF Then
        
            
            On Error Resume Next
            frmBusca.Show 1
            If frmBusca.tag <> "" Then
                If KeyCode = vbKeyB Then
                    cmbxBusca.ListIndex = 1
                End If
                If KeyCode = vbKeyN Then
                    cmbxBusca.ListIndex = 0
                End If
                If KeyCode = vbKeyF Then
                    cmbxBusca.ListIndex = 2
                End If
                If KeyCode = vbKeyM Then
                    cmbxBusca.ListIndex = 3
                End If
                If KeyCode = vbKeyG Then
                    cmbxBusca.ListIndex = 4
                End If
                
                
                txtbxExpr.Text = frmBusca.txtbxExpr.Text
            End If
            Unload frmBusca
        
    End If
    If KeyCode = vbKeyI And Shift = 1 Then
        If strViscusto_usu <> "S" Then
            MsgBox "Você não tem permissão para isto!", vbCritical
            Exit Sub
        End If
        P_InfPro ado!cod_pro
        RefreshAdoPosition ado
    End If
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        Clipboard.SetText dtgrd.Text
    End If
End Sub


Private Sub dtgrd_LostFocus()
   dtgrd.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    On Error Resume Next
    With ado
        If .RecordCount > 0 Then
            lblEst.caption = FormatVal(!Estoque, "###,##0")
            lblPreco.caption = FormatVal(!pv_pro)
            lblEstRes.caption = FormatVal(!r_pro)
            lblMarca.caption = !nome_mar
            lblCategoria.caption = !nome_cat
            txtbxST.Text = !s_pro
            'P_PreencheBarCode !cod_pro, lstBarCode
            BuscaUltEntrada
            lblEstoqueSim.caption = BuscaEstoqueSim(!cod_pro)
            VerSim !cod_pro
            
        Else
            
        End If
    End With
End Sub

Private Sub dtgrd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnOpcoes
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
     txtbxExpr.SetFocus
        bActive = True
        If Me.tag <> "" Then
        busca
    End If
    'If strViscusto_usu <> "S" Then
    '    dtgrd.Columns(7).Visible = False
    '    dtgrd.Columns(8).Visible = False
    '    dtgrd.Columns(9).Visible = False
     ''   dtgrd.Columns(10).Visible = False
     '   dtgrd.Columns(11).Visible = False
      '  dtgrd.Columns(12).Visible = False
    'End If
    
    If strViscusto_usu = "S" Then
       ' chkEst.Enabled = True
       ' chkEst_Click
    Else
        'chkEst.Enabled = False
    End If
    
    If strCadprod_usu <> "S" Then
        mnEditMult.Enabled = False
        cmdExportar.Enabled = False
    End If
    If Not bControlaEstoque Then
        cmdReservadas.Visible = False
            lblEr.Visible = False
            dtgrd.Columns(6).Visible = False
            dtgrd.Columns(18).Visible = False
    End If
    
End Sub
Sub buscaVei()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = SelectProduto_Veiculo(ado!cod_pro, txtVeiculo.Text)
    abreObjetoAdo adoveiculo, strSQL, gridVeiculo
    
    Dim i As Integer
    Dim j As Integer
    
    
    

    
    
    With adoveiculo
        Do While Not .EOF
            If !id_vei = txtVeiculo.tag Then
                
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
    'gridVeiculo.Columns(1).WrapText = True
    
    
trata_Erro:
End Sub
Sub VerSim(cod_pro As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If cod_pro = "" Then
        pcSimilares.Visible = False
    End If
    strSQL = "select count(*) as total from tab_similares where cod_pro='" & cod_pro & "' OR cod_sim='" & cod_pro & "' "
   
    abreObjetoAdo rPesq, strSQL
   
    If rPesq!total > 0 Then
        pcSimilares.Visible = True
    Else
        pcSimilares.Visible = False
    End If
   
   
   
   
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF11 And Shift = vbCtrlMask Then
         bActive = False
         deletarTodos
    End If
    If KeyCode = vbKeyL And Shift = vbCtrlMask Then
        txtCliente.SetFocus
        
        
        
    End If
    If KeyCode = vbKeyV And Shift = vbAltMask Then
        txtVeiculo.SetFocus
        
        
        
    End If
    If KeyCode = vbKeyP And Shift = vbAltMask Then
        txtbxExpr.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If strViscusto_usu <> "S" Then
   '     dtgrd.Columns(8).Visible = False
    End If
    Set ado = New ADODB.Recordset
    
    
    Set adoveiculo = New ADODB.Recordset
    Set adoConversao = New ADODB.Recordset
    
    
    If Not bUsaVeiculo Then
        pcRight.width = 200
        pcRight.Visible = False
        lblVeic.Visible = False
        cmdVei.Visible = False
        txtVeiculo.Visible = False
        lblFab.Visible = False
    End If

    cmbxOrd.ListIndex = 0
    
    resizeForm frmEdiPro
    'P_PreencheImpostos cmbxICMS
'    txtbxExpr.SetFocus
    'Busca
End Sub
Sub buscaConversao()
    Dim strSQL As String
    strSQL = SelectProdutoConversao(ado!cod_pro)
    abreObjetoAdo adoConversao, strSQL, gridConversao
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu mnEditMult
    End If
End Sub

Private Sub Form_Resize()
   posicionaControles
   
End Sub
Sub posicionaControles()
    On Error Resume Next
    pcBotoes.Top = Me.height - Me.Top - pcBotoes.height - 200
    pcBotoes.width = pcRight.Left - 100
    posicionaADireita cmndbtnSair, pcBotoes
    'posicionaADireita pcRight, Me
    pcBusca.Top = pcBotoes.Top - pcBusca.height
'    pcSimilares.Top = pcBusca.Top - pcSimilares.height
    
    PosicionaGradeLeft dtgrd, pcRight, 100
    PosicionaGradeTop dtgrd, pcInfo
    posicionaADireita pcPreco, pcInfo
    
    'pcImagem.height = pcRight.height
   ' pcImagem.width = pcRight.width
    dtgrd.height = pcBusca.Top - dtgrd.Top
    gridConversao.height = pcBotoes.Top - gridConversao.Top
    pcBusca.width = dtgrd.width
    pcPreco.Top = pcInfo.Top + pcBusca.Top '- 100
    pcPreco.Left = pcBusca.width - pcPreco.width - 100
End Sub

Private Sub mnAlterarPrecoLote_Click()
    If strAltpreco_usu <> "S" Then
        MsgBox "Você não tem permissão para esse acesso!", vbExclamation
        Exit Sub
    End If
    frmAlterarPrecosEmLote.Show 1
End Sub

Private Sub mnCEST_Click()
    mudarCest
End Sub

Private Sub mnEstoqueReservado_Click()
    cmdReservado_Click
End Sub

Private Sub mnImg_Click()
    On Error Resume Next
    AbreImagens ado!cod_pro
End Sub

Private Sub mnMargem_Click()
        margemLucro
End Sub

Private Sub mnMudarICMS_Click()
    frmMudarICMS.Show 1
End Sub

Private Sub mnNCM_Click()
    mudarNcm
End Sub

Private Sub mnOpcoesNfe_Click()
    opcoesNfe
End Sub

Private Sub mnPreco_Click()
    MudarPreco
End Sub

Private Sub mnTornarInv_Click()
    cmdVisivel_Click
End Sub

Private Sub pcBotoes_GotFocus()
    TrocaTabKey 3
End Sub

Sub busca(Optional bSimilares As Boolean = False)
    Dim strSQL As String
    Dim sOrder As String
    Dim sCampo As String
    Dim extra As String
    
    P_Status "Pesquisando..."
    
    
    extra = RetornaValorTagXML(txtbxExpr.Text, "{", "}")
        
    
    
        
    If cmbxOrd.ListIndex = 1 Then
        sOrder = "tp.cod_pro,tp.nome_pro"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrder = "tp.codfab_pro,tp.nome_pro,tp.cod_pro"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrder = "tp.endereco_pro,tp.nome_pro,tp.cod_pro"
    Else
        sOrder = "tp.nome_pro,tp.cod_pro "
    End If
    strSQL = "select distinct if(ef_pro<0,' - ','') as sinal,(e.er-e.reservado) as estoque, tg.nome_cat, te.nome_mar,"
    strSQL = strSQL & " tp.*,((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & "  as liq, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) as lb, "
    strSQL = strSQL & " (((tp.pv_pro/tp.cr_pro-1)*100)-" & F_GravaDec(F_TaxasEntradas) & ") * pv_pro/100  as lr, "
    strSQL = strSQL & " ((tp.pv_pro/tp.cr_pro-1)*100) * pv_pro/100  as bruto, "
    strSQL = strSQL & " e.cod_emp,e.er, ori.descricao as nome_ori,  "
    strSQL = strSQL & "     (SELECT GROUP_CONCAT(caminho SEPARATOR ',') AS caminho FROM imagem_produto AS img where img.cod_pro=tp.cod_pro ORDER BY id desc ) AS caminho"
    strSQL = strSQL & " from " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " inner join estoque as e on tp.cod_pro=e.cod_pro "
    strSQL = strSQL & " left join " & TAB_CATEGORIAS & " as tg on tg.cod_cat=tp.cod_cat  "
    strSQL = strSQL & " left join " & TAB_MARCAS & " as te on te.cod_mar=tp.cod_mar  "
    strSQL = strSQL & " left join sped_origem as ori on tp.cod_ori=ori.codigo"
    strSQL = strSQL & " left join " & TAB_IMPOSTOS & " as ti on ti.cod_imp=tp.cod_imp  "
    strSQL = strSQL & " left join produto_veiculo as pv on tp.cod_pro=pv.cod_pro"
    strSQL = strSQL & " left join veiculo as v on pv.id_vei=v.id "
    strSQL = strSQL & " left join produto_conversao as conv on tp.cod_pro=conv.cod_pro "

        
    If chkAtacado.Value Then
        strSQL = strSQL & " inner join produto_atacado as tpa on tp.cod_pro=tpa.codigo "
    End If
            
            
    If bSimilares = True Then
        strSQL = strSQL & " left join tab_similares as ts on tp.cod_pro=ts.cod_pro or tp.cod_pro=ts.cod_sim "
        strSQL = strSQL & " where ts.cod_pro='" & ado!cod_pro & "' or ts.cod_sim='" & ado!cod_pro & "'"
    Else
        strSQL = strSQL & " where 1> 0 "
        strSQL = strSQL & " and( nome_pro like '" & IIf(bUsaVendaSupermercado, "", "%")
        strSQL = strSQL & txtbxExpr.Text & "%' "
        If txtbxExpr.Text <> "" Then
            strSQL = strSQL & "or tp.cod_pro like '%"
            strSQL = strSQL & txtbxExpr.Text & "%' "
            strSQL = strSQL & " or tp.ncm_pro like '" & txtbxExpr.Text & "%' "
            strSQL = strSQL & " or tp.codfab_pro like '%" & txtbxExpr.Text & "%' "
            strSQL = strSQL & " or tp.detalhe_pro like '%" & txtbxExpr.Text & "%' "
            strSQL = strSQL & " or v.descricao like '%" & txtbxExpr.Text & "%' "
            strSQL = strSQL & " or conv.referencia like '" & txtbxExpr.Text & "%' "
            strSQL = strSQL & " or tp.codbarra_pro like '" & txtbxExpr.Text & "%') "
        Else
            strSQL = strSQL & ")"
        End If
    
        
    End If
    If txtVeiculo.tag <> "" Then
        strSQL = strSQL & " and v.id ='" & txtVeiculo.tag & "' "
    End If
'    If cmbxICMS.text <> "" Then
'        strSQL = strSQL & " and ti.cod_imp='" & Left(cmbxICMS.text, 3) & "' and "
'    End If
    
    If txtOrigem.tag <> "" Then
        strSQL = strSQL & " and tp.cod_ori='" & txtOrigem.tag & "' "
    End If
    If txtEmpresa.tag <> "" Then
        strSQL = strSQL & " and e.cod_emp='" & txtEmpresa.tag & "' "
    End If
    If chkComb.Value = 1 Then
        strSQL = strSQL & " and exists (select cod_pro from produtos_comb where cod_pro=tp.cod_pro) "
    End If
    
    
'    strSQL = strSQL & " and tg.cod_cat=tp.cod_cat and te.cod_mar=tp.cod_mar "
    If chkTodas.Value = 0 Then
        strSQL = strSQL & " and tp.er_pro >0"
    End If
    If chkInvisiveis.Value = 0 Then
        strSQL = strSQL & " and tp.visivel_pro='S'"
    Else
        strSQL = strSQL & " and tp.visivel_pro='N'"
    End If
    
    If txtCat.tag <> "" Then
        strSQL = strSQL & " and tp.cod_cat='" & txtCat.tag & "'  "
    End If
    If txtMarca.tag <> "" Then
        strSQL = strSQL & " and tp.cod_mar='" & txtMarca.tag & "'  "
    End If
    
'    If cmbxCriterio.ListIndex = 1 Then
'        strSQL = strSQL & " and ef_pro<er_pro"
'    ElseIf cmbxCriterio.ListIndex = 2 Then
'        strSQL = strSQL & " and ef_pro>er_pro"
'    ElseIf cmbxCriterio.ListIndex = 3 Then
'        strSQL = strSQL & " and ef_pro<0"
'    ElseIf cmbxCriterio.ListIndex = 4 Then
'        strSQL = strSQL & " and er_pro<estmin_pro"
'    ElseIf cmbxCriterio.ListIndex = 5 Then
'        strSQL = strSQL & " and ef_pro>(er_pro*2)"
'    ElseIf cmbxCriterio.ListIndex = 6 Then
'        strSQL = strSQL & " and er_pro>(ef_pro*2)"
'    End If
    If extra <> "" Then
        strSQL = strSQL & " and " & extra
    End If
    strSQL = strSQL & " group by tp.cod_pro, e.cod_emp "
    
    strSQL = strSQL & " order by " & sOrder
    'adc.RecordSource = strSQL
    'adc.Refresh
    abreObjetoAdo ado, strSQL, dtgrd
    On Error Resume Next
    dtgrd.caption = "[Listadas..: " & ado.RecordCount & _
        "   Cadastradas: " & F_BuscaUmCampo("count(*)", TAB_PRODUTOS, "") & "]"
    P_Status ""
End Sub

Private Sub pcContainer_Click()
    pcImg_Click
End Sub

Private Sub pcImagem_Click()
    pcImg_Click
End Sub

Private Sub pcImg_Click()
    On Error Resume Next
    Shell "explorer " & pcImagem.tag
End Sub


Private Sub strip_Click()
    If strip.SelectedItem.Index = 1 Then
        pcVeiculos.Visible = False
        pcImagem.Visible = True
    Else
        pcVeiculos.Visible = True
        pcImagem.Visible = False
    End If
End Sub

Private Sub txtbxExpr_GotFocus()
    
    ganhaFoco txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        busca
        If ado.RecordCount > 0 Then
            dtgrd.SetFocus
        End If
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dtgrd.SetFocus
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    AjustaTexto txtbxExpr
    P_TileMsg ""
    
End Sub

Private Sub txtCat_Change()
    txtCat.tag = ""
End Sub

Private Sub txtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCat
    End If
End Sub

Private Sub txtCliente_Change()
    txtCliente.tag = ""
End Sub

Private Sub txtCliente_GotFocus()
    P_SelecTexto txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCli
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        dtgrd.SetFocus
    End If
End Sub
Sub buscaCli()
    Dim c As cliente
    On Error GoTo trata_Erro
    c = Clientes.F_BuscaCliente(txtCliente.Text)
    
    If c.cod_cli <> "" Then
        txtCliente.Text = c.nome_Cli
        txtCliente.tag = c.cod_cli
        lblDescontoCli.caption = c.desconto_cli
        lblPv.caption = FormatVal(ado!pv_pro - (ado!pv_pro * CCur(lblDescontoCli.caption / 100)))
    End If
    Exit Sub
trata_Erro:
    
End Sub
Private Sub txtMarca_Change()
    txtMarca.tag = ""
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaMarca
    End If
End Sub

Private Sub txtOrigem_Change()
    txtOrigem.tag = ""
End Sub

Private Sub txtOrigem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaOrigem
    End If
End Sub

Private Sub txtVeiculo_Change()
    txtVeiculo.tag = ""
End Sub

Private Sub txtVeiculo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        selecionaVei
    End If
End Sub
Sub selecionaVei()
    Dim v As Veiculo
    v = Veiculos.BuscaVeiculo(txtVeiculo.Text, "")
    If v.id > 0 Then
        txtVeiculo.Text = v.descricao
        txtVeiculo.tag = v.id
        lblFab.caption = v.fabricante.descricao
        busca
    End If
End Sub
