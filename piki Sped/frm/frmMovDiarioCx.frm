VERSION 5.00
Begin VB.Form frmMovDiarioCx 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimentação Diária do Caixa"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcRestr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   240
      ScaleHeight     =   3225
      ScaleWidth      =   7065
      TabIndex        =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Visualização não disponível clique em IMPRIMIR para impressão."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   1320
         TabIndex        =   46
         Top             =   840
         Width           =   4050
      End
   End
   Begin VB.PictureBox pcECF 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   7455
      TabIndex        =   41
      Top             =   4920
      Width           =   7455
   End
   Begin VB.Frame Frame4 
      Caption         =   "&Turno:"
      Height          =   735
      Left            =   4080
      TabIndex        =   20
      Top             =   3960
      Width           =   1935
      Begin VB.ComboBox cmbxTurno 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "F&undo de Caixa:"
      Height          =   735
      Left            =   6120
      TabIndex        =   22
      Top             =   3960
      Width           =   1455
      Begin VB.TextBox txtbxFundoCaixa 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "0,00"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   5880
      Picture         =   "frmMovDiarioCx.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnLan 
      Caption         =   "&Ver Lançamentos"
      Height          =   735
      Left            =   3720
      Picture         =   "frmMovDiarioCx.frx":054E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmndbtnImprimir 
      Caption         =   "&Imprimir"
      Height          =   735
      Left            =   1920
      Picture         =   "frmMovDiarioCx.frx":0638
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cai&xa"
      Height          =   735
      Left            =   1560
      TabIndex        =   18
      Top             =   3960
      Width           =   2415
      Begin VB.ComboBox cmbxCaixa 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Data:"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
      Begin VB.TextBox txtbxData 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmndbtnFecharCaixa 
      Caption         =   "&Fechar Caixa"
      Height          =   735
      Left            =   120
      Picture         =   "frmMovDiarioCx.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame frControle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmndbtnDet 
         Caption         =   "Detalhar"
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Cancelados:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   3120
         Width           =   1230
      End
      Begin VB.Label lblItensCanc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   1800
         TabIndex        =   47
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multi-Fácil:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   43
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   42
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   1200
         Width           =   360
      End
      Begin VB.Label lblCancela 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   1800
         TabIndex        =   39
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas Canceladas:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   1470
      End
      Begin VB.Line Line10 
         BorderStyle     =   3  'Dot
         X1              =   6720
         X2              =   120
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espécie"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   240
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   840
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   6720
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   6720
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderStyle     =   3  'Dot
         X1              =   6720
         X2              =   120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   6720
         X2              =   120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   6720
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   5040
         X2              =   5040
         Y1              =   240
         Y2              =   2160
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   3480
         X2              =   3480
         Y1              =   240
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   2040
         X2              =   2040
         Y1              =   240
         Y2              =   2160
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   36
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   35
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   34
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3960
         TabIndex        =   33
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblTotalAB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   6120
         TabIndex        =   32
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblTotalSai2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   4440
         TabIndex        =   31
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(D)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   240
         Left            =   3720
         TabIndex        =   30
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(C)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   240
         Left            =   2640
         TabIndex        =   29
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblTotalDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   6120
         TabIndex        =   28
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblTotalSai 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   4440
         TabIndex        =   24
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblTotalAp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   15
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label lblCHP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   14
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label lblTotalAv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6720
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lblFin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   12
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financeira:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label lblCHV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   10
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque à Vista:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque A Prazo:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total.........:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label lblTotalEntradas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   2880
         TabIndex        =   6
         Tag             =   "!"
         ToolTipText     =   "Total de Vendas menos Devoluções."
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label lblDin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   4
         Tag             =   "!"
         ToolTipText     =   "Vendas a Vista"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblCar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2880
         TabIndex        =   3
         Tag             =   "!"
         ToolTipText     =   "Vendas a Prazo"
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carteira:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dinheiro:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmMovDiarioCx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim operacao As String
Private Sub cmbxCaixa_Click()
    busca
End Sub
Function VendasFechadas() As Boolean
    If F_BuscaUmCampo("num_ven", "" & TAB_VENDAS & "", "situacao_ven=2 and cod_cai='" & strCod_Cai & "'") <> "" Then
        VendasFechadas = True
    Else
        VendasFechadas = False
    End If
End Function

Private Sub cmbxTurno_Click()
    busca
End Sub

Private Sub cmbxTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmndbtnFecharCaixa_Click()
    Dim strSQL As String
    Dim cTotal As Currency
    Dim sCod_Con  As String
    If verificaVendaAberta(Left(cmbxCaixa.Text, 2)) Then
        MsgBox "Não é possível fechar o caixa pois existem vendas em aberto! Procure o responsável!", vbCritical
        Exit Sub
    End If
    
    'On Error GoTo trata_erro
    strSQL = F_BuscaUmCampo("count(*)", "" & TAB_VENDAS & "", "situacao_ven=2 and cod_cai='" & strCod_Cai & "'")
    If Int(strSQL) > 0 Then
        MsgBox "Não é possível fechar o caixa pois existem vendas a efetivar!", vbExclamation
        Exit Sub
    End If
'    If F_BuscaUmCampo("data_cai", "tab_fechacaixa", "data_cai=" & _
'        F_PesqData(txtbxData.Text) & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "'") <> "" Then
'        MsgBox "Este caixa já foi fechado neste dia!", vbExclamation
'        Exit Sub
'    End If
    If VendasFechadas = True Then
        MsgBox "Existem vendas condicionais a serem efetivadas!", vbExclamation
        On Error Resume Next
        frmEfetivaVen.Show 1
        Exit Sub
    End If
    If MsgBox("Deseja realmente fechar o Turno do dia " & txtbxData.Text _
        & " Com " & txtbxFundoCaixa.Text & " de Fundo de caixa?", vbYesNo + vbCritical) = vbYes Then
        If CCur(txtbxFundoCaixa.Text) = 0 Then
            
            If MsgBox("Este fechamento não contém fundo de caixa!" & Chr(13) & _
                "Deseja realmente continuar?", vbYesNo + vbQuestion) = vbYesNo Then
                Exit Sub
            End If
        End If
        If F_BuscaUmCampo("status_tur", "tab_turncai", "cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' and cod_tur='" & Left(cmbxTurno.Text, 1) & "' and data_tur=" & F_GravaData(txtbxData.Text)) = "F" Then
            MsgBox "Este turno já se encontra Fechado!", vbCritical
            Exit Sub
         End If
'        strSQL = "insert into tab_fechacaixa(data_Cai,valor_FEC,fundo_cai,cod_usu,COD_CAI)values("
'        strSQL = strSQL & F_GravaData(txtbxData.Text) & ","
'        strSQL = strSQL & F_GravaDec(lblTotalAB.caption) & ","
'        strSQL = strSQL & F_GravaDec(txtbxFundoCaixa.Text) & ","
'        strSQL = strSQL & "'" & strCod_Usu & "','" & Left(cmbxCaixa.Text, 2) & "') "
'        conn.execute strSQL
        
        
        strSQL = "update tab_turncai set logout_usu='" & strCod_Usu & "',"
        strSQL = strSQL & "Fechamento_tur=" & F_GravaHora(Now) & ","
        strSQL = strSQL & "valorfechamento_tur=" & F_GravaDec(lblTotalAB.caption) & ","
        strSQL = strSQL & "status_tur='F' "
        strSQL = strSQL & " where cod_Cai='" & Left(cmbxCaixa.Text, 2) & "' "
        strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
        conn.Execute strSQL
        
        'strSQL = F_GravaCai("D", txtbxData.Text, CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.Text), _
            "TRANSFERÊNCIA DE CAIXA", "", "DIN", "01", 0, 0, Left(cmbxCaixa.Text, 2), strCod_Tur)
        
        strSQL = F_GravaFinanceiro("C", txtbxData.Text, CCur(lblTotalDin.caption) - CCur(txtbxFundoCaixa.Text), _
             "TRANSFERÊNCIA DE CAIXA", "DIN", constContaCaixa, "", "", "", strPlaVendasEfetuadas)
         
        
        
            
            
        strSQL = "Update tab_Caixas set status_Cai='F' where cod_Cai='" & Left(cmbxCaixa.Text, 2) & "'"
        conn.Execute strSQL
        MsgBox "Caixa " & cmbxCaixa.Text & "  Fechado com sucesso!", vbInformation
        
        
    End If
        
    Exit Sub
trata_erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnImprimir_Click()
    If strVisrelfin_usu = "S" Then
        imprime frmVisualiza.Pic
        frmVisualiza.Show 1
        If frmVisualiza.tag <> "" Then
            imprime Printer
            Printer.EndDoc
        End If
        Unload frmVisualiza
    Else
        MsgBox "Você não tem permissão para este acesso!", vbExclamation
        'if str
        'imprime Printer
        'Printer.EndDoc
        'MsgBox "Impressão Concluída!", vbInformation
    End If
End Sub
Sub imprime(obj As Object)
    obj.FontName = "arial"
    Dim yIni As Single
    Dim yFin As Single
    Dim rPesq As New adodb.Recordset
    Dim strSQL As String
    Dim sHora As String
    
    P_CabRelTimbr "Fechamento de Caixa " & txtbxData.Text & " " & cmbxCaixa.Text & " Turno: " & cmbxTurno.Text & " - (Usuário: " & strNome_Usu & ")", obj
    strSQL = "select now();"
    rPesq.Open strSQL, conn_client
    sHora = rPesq.Fields(0).Value
    
    
    rPesq.Close
    
    P_CentraImpressao obj, "Hora no Servidor " & sHora
    obj.Print ""
    
    
    obj.Print ""
    
    obj.Print ""
    yIni = obj.CurrentY
    P_Line obj, 8, obj.CurrentY, 142, obj.CurrentY
    P_ImprX obj, "Espécie: ", 10
    P_ImprX obj, "(C): ", 60, True, True
    P_ImprX obj, "(D): ", 100, True, True
    P_ImprX obj, "(Total): ", 140, False, True
    P_Line obj, 8, obj.CurrentY, 142, obj.CurrentY
    
    P_ImprX obj, "Dinheiro: ", 10
    P_ImprX obj, lblDin.caption, 60, True, True
    P_ImprX obj, lblTotalSai.caption, 100, True, True
    P_ImprX obj, lblTotalDin.caption, 140, False, True
    P_ImprX obj, "Cheque À Vista ", 10
    P_ImprX obj, lblCHV.caption, 60, True, True
    P_ImprX obj, "...", 100, True, True
    P_ImprX obj, "...", 140, False, True
    P_ImprX obj, "Financeira ", 10
    P_ImprX obj, lblFin.caption, 60, True, True
    
    P_ImprX obj, "...", 100, True, True
    P_ImprX obj, "...", 140, False, True
    
    P_ImprX obj, "Multi-Fácil", 10
    P_ImprX obj, lblMulti.caption, 60, True, True
    
    P_ImprX obj, "...", 100, True, True
    P_ImprX obj, "...", 140, False, True
    
    P_Line obj, 8, obj.CurrentY, 142, obj.CurrentY, vbDot
    P_ImprX obj, lblTotalAv.caption, 140, False, True
    P_Line obj, 8, obj.CurrentY, 142, obj.CurrentY, vbDot
    
    P_ImprX obj, "Carteira", 10
    P_ImprX obj, lblCar.caption, 60, True, True
    P_ImprX obj, "...", 100, True, True
    P_ImprX obj, "...", 140, False, True
    P_ImprX obj, "Cheque A Prazo", 10
    P_ImprX obj, lblCHP.caption, 60, True, True
    P_ImprX obj, "...", 100, True, True
    P_ImprX obj, "...", 140, False, True
    P_Line obj, 8, obj.CurrentY, 142, obj.CurrentY, vbDot
    P_ImprX obj, lblTotalAp.caption, 140, False, True
    P_Line obj, 8, obj.CurrentY, 142, obj.CurrentY, vbDot
    
    obj.FontBold = True
    P_ImprX obj, "Total", 10
    P_ImprX obj, lblTotalEntradas.caption, 60, True, True
    P_ImprX obj, lblTotalSai2.caption, 100, True, True
    P_ImprX obj, lblTotalAB.caption, 140, False, True
    obj.FontBold = False
    yFin = obj.CurrentY
    P_Line obj, 8, obj.CurrentY, 142, obj.CurrentY
    P_Line obj, 8, yIni, 8, yFin
    P_Line obj, 142, yIni, 142, yFin
    obj.Print ""
    imprimeCheques obj
    imprimeCHP obj
    imprimeCAR obj
    imprimeFIN obj
    
    
    
    
End Sub

Sub imprimeCheques(obj As Object)
    Dim s 'trSQL As String
    Dim rPesq As New adodb.Recordset
    Dim cTotal As Currency
    strSQL = "select tc.num_ven,tx.* from tab_chv as tx,tab_caixa as tc "
    strSQL = strSQL & " where tc.seq_cai=tx.seq_cai and data_cai=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.Text, 1)
    strSQL = strSQL & " order by tc.seq_cai "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Cheques à Vista"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "CPF", 20
        P_ImprX obj, "Titular", 50
        P_ImprX obj, "Nº Cheque", 110
        P_ImprX obj, "Agência", 135
        P_ImprX obj, "C/C", 155
        P_ImprX obj, "Banco", 173
        P_ImprX obj, "Valor", 198, False, True
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !cpf_che, 20
            P_ImprX obj, Left(StrConv(!nome_che, vbProperCase), 40), 50
            P_ImprX obj, !num_che, 110
            P_ImprX obj, !age_che, 135
            P_ImprX obj, !cco_che, 155
            P_ImprX obj, !cod_bco, 173
            P_ImprX obj, FormatVal(!Valor_che), 198, False, True
            
            cTotal = cTotal + !Valor_che
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 198, False, True
        obj.Print ""
        
    End With


End Sub
Sub imprimeCHP(obj As Object)
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim cTotal As Currency
    strSQL = "select tv.num_ven,tx.* from tab_cheques1 as tx, " & TAB_VENFEC & "  as tf,tab_doc1 as td," & TAB_VENDAS & " as tv "
    strSQL = strSQL & " where tv.num_ven=td.num_ven and dataefet_ven=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " and tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.Text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.Text, 1)
    strSQL = strSQL & " and tx.seq_doc=td.seq_doc "
    strSQL = strSQL & " order by tv.num_ven "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Cheques a Prazo"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "CPF", 22
        P_ImprX obj, "Titular", 50
        P_ImprX obj, "Nº Cheque", 112
        P_ImprX obj, "Agência", 135
        P_ImprX obj, "C/C", 155
        P_ImprX obj, "Banco", 173
        P_ImprX obj, "Valor", 198, False, True
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !cpf_che, 22
            P_ImprX obj, Left(StrConv(!nome_che, vbProperCase), 40), 50
            P_ImprX obj, !num_che, 112
            P_ImprX obj, !age_che, 135
            P_ImprX obj, !cco_che, 155
            P_ImprX obj, !cod_bco, 173
            P_ImprX obj, FormatVal(!Valor_che), 198, False, True
            
            cTotal = cTotal + !Valor_che
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 200, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 198, False, True
        
        
    End With


End Sub
Sub imprimeCAR(obj As Object)
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim cTotal As Currency
    strSQL = "select tv.nomecli_ven,tv.num_ven,ti.* from tab_itdoc1 as ti, " & TAB_VENFEC & "  as tf,tab_doc1 as td," & TAB_VENDAS & " as tv "
    strSQL = strSQL & " where tv.num_ven=td.num_ven and dataefet_ven=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " and tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.Text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.Text, 1)
    strSQL = strSQL & " and ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " order by tv.num_ven "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Vendas em Carteira"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "Cliente", 25
        P_ImprX obj, "Vencimento", 100
        P_ImprX obj, "Valor", 128, False, True
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !nomecli_ven, 25
            P_ImprX obj, !venc_itd, 100
            P_ImprX obj, FormatVal(!valor_itd), 128, False, True
            
            cTotal = cTotal + !valor_itd
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 128, False, True
        
        
    End With


End Sub
Sub imprimeFIN(obj As Object)
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    Dim cTotal As Currency
    strSQL = "select tv.nomecli_ven,tfi.nome_fin,tv.num_ven,ti.* from tab_venfin as ti, " & TAB_VENFEC & "  as tf," & TAB_VENDAS & " as tv, "
    strSQL = strSQL & " tab_financeiras as tfi where tv.num_ven=ti.num_ven and dataefet_ven=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " and ti.cod_fin=tfi.cod_fin "
    strSQL = strSQL & " and tf.num_ven=tv.num_ven "
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.Text, 2) & "'"
    strSQL = strSQL & " and cod_tur=" & Left(cmbxTurno.Text, 1)
    
    strSQL = strSQL & " order by tv.num_ven "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount = 0 Then
            .Close
            Exit Sub
        End If
        
        P_CentraImpressao obj, "Vendas por Financeiras"
        obj.Print
        
        P_ImprX obj, "Nº Venda", 5
        P_ImprX obj, "Cliente", 25
        P_ImprX obj, "Financeira", 85
        P_ImprX obj, "Valor", 128, False, True
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY
        Do While Not .EOF
            If obj.CurrentY >= obj.ScaleHeight - 20 Then
                If TypeName(obj) = "Printer" Then
                    obj.NewPage
                End If
            End If
            P_ImprX obj, !num_ven, 5
            P_ImprX obj, !nomecli_ven, 25
            P_ImprX obj, !nome_fin, 85
            P_ImprX obj, FormatVal(!valor_ven), 128, False, True
            cTotal = cTotal + !valor_ven
            .MoveNext
        Loop
        
        P_Line obj, 4, obj.CurrentY, 130, obj.CurrentY, vbDot
        P_ImprX obj, FormatVal(cTotal), 128, False, True
        
        
    End With


End Sub

Private Sub cmndbtnLan_Click()
    If strVisrelfin_usu = "S" Then
        frmFechaCaixa.Show 1
    Else
        MsgBox "Você não tem permissão para este acesso!", vbCritical
        
    End If
    
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub


Private Sub Command3_Click()
    
End Sub

Private Sub Form_Activate()
    If strVisrelfin_usu = "S" Then
        frControle.Visible = True
        pcRestr.Visible = False
    Else
        frControle.Visible = False
        pcRestr.Visible = True
    End If
    If strFechames_usu <> "S" Then
        cmbxCaixa.Enabled = False
        cmndbtnLan.Enabled = False
    End If
    busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        operacao = "F"
        pcECF.BackColor = vbGreen
        busca
    ElseIf KeyCode = vbKeyF9 Then
        operacao = "R"
        pcECF.BackColor = vbYellow
        busca
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtbxData.Text = Date
    
    
    P_PreencheCaixas cmbxCaixa
    P_PreencheTurnos cmbxTurno
    P_LocalizaOperacao strCod_Cai, cmbxCaixa
    P_LocalizaOperacao strCod_Tur, cmbxTurno
    pcECF.ForeColor = vbGreen
    operacao = "F"
End Sub

Private Sub txtbxData_GotFocus()
    ganhaFoco txtbxData
End Sub

Private Sub txtbxData_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        busca
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxData_LostFocus()
    P_AjustaData txtbxData
    busca
End Sub
Sub buscaR()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    On Error GoTo trata_erro
    lblCar.caption = "0,00"
    lblCHP.caption = "0,00"
    lblCHV.caption = "0,00"
    lblDin.caption = "0,00"
    lblFin.caption = "0,00"
    lblMulti.caption = "0,00"
    lblTotalSai.caption = "0,00"
    strSQL = "select esp_cai,sum(tc.valor_cai) as total "
    strSQL = strSQL & " from tab_caixa as tc "
    strSQL = strSQL & " where tc.data_cai=" & F_PesqData(txtbxData.Text)
    strSQL = strSQL & " and tipo_cai='C' "
    strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    strSQL = strSQL & " group by esp_cai "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If .RecordCount > 0 Then
                If !Esp_cai = "DIN" Then
                    lblDin.caption = FormatVal(!total)
                End If
                If !Esp_cai = "FIN" Then
                    lblFin.caption = FormatVal(!total)
                End If
                If !Esp_cai = "CHP" Then
                    lblCHP.caption = FormatVal(!total)
                End If
                If !Esp_cai = "CHV" Then
                    lblCHV.caption = FormatVal(!total)
                End If
                If !Esp_cai = "CAR" Then
                    lblCar.caption = FormatVal(!total)
                End If
                If !Esp_cai = "MUL" Then
                    lblMulti.caption = FormatVal(!total)
                End If
            End If
            
            .MoveNext
        Loop
        lblTotalAv.caption = FormatVal(CCur(lblDin.caption) + CCur(lblCHV.caption) + CCur(lblFin.caption) + CCur(lblMulti.caption))
        lblTotalAp.caption = FormatVal(CCur(lblCHP.caption) + CCur(lblCar.caption))
        lblTotalEntradas.caption = FormatVal(CCur(lblTotalAv.caption) + CCur(lblTotalAp.caption))
        .Close
        
    End With
    strSQL = "select sum(tc.valor_cai) as total from tab_Caixa as tc "
    strSQL = strSQL & " where tipo_cai='D' and data_cai=" & F_PesqData(txtbxData.Text)
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If Not IsNull(!total) Then
            lblTotalSai.caption = FormatVal(!total)
        End If
        .Close
    End With
    lblTotalAB.caption = FormatVal(CCur(lblTotalEntradas.caption) - CCur(lblTotalSai.caption))
    lblTotalDin.caption = FormatVal(CCur(lblDin.caption) - CCur(lblTotalSai.caption))
    lblTotalSai2.caption = lblTotalSai.caption
    'parei aki
    strSQL = "select sum(tc.valor_ven) as total from tab_cancelamentos  as tc,"
    strSQL = strSQL & "  " & TAB_VENFEC & "  as tf "
    strSQL = strSQL & " where tc.num_ven=tf.num_ven "
    strSQL = strSQL & " and left(data_can,10) = " & F_PesqData(txtbxData.Text)
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and tf.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If Not IsNull(!total) Then
            lblCancela.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    strSQL = "select sum(unit_ite*qtd_ite) as total from tab_itvndcncl as ti,"
    strSQL = strSQL & " " & TAB_VENDAS & " as tv "
    strSQL = strSQL & " where ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and left(tv.data_ven,10) = " & F_PesqData(txtbxData.Text)
    strSQL = strSQL & " and ti.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and ti.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If Not IsNull(!total) Then
            lblItensCanc.caption = FormatVal(!total)
        End If
        .Close
    End With
    
    Exit Sub
trata_erro:
    MostraErro
End Sub
Sub busca()
    On Error GoTo trata_erro
    If operacao = "F" Then
        buscaF
    Else
        buscaR
    End If
    Exit Sub
trata_erro:
    MostraErro
End Sub
Sub buscaF()
    Dim strSQL As String
    Dim rPesq As New adodb.Recordset
    lblCar.caption = "0,00"
    lblCHP.caption = "0,00"
    lblCHV.caption = "0,00"
    lblDin.caption = "0,00"
    lblFin.caption = "0,00"
    lblTotalSai.caption = "0,00"
    strSQL = "select esp_cai,sum(tc.valor_cai) as total "
    strSQL = strSQL & " from tab_caixa as tc,"
    strSQL = strSQL & " tab_cupons as ts "
    strSQL = strSQL & " where ts.num_ven=tc.num_ven"
    strSQL = strSQL & " and tipo_cai='C' "
    strSQL = strSQL & " and cancelado_cup='N' "
    strSQL = strSQL & " and tc.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and tc.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    strSQL = strSQL & " and data_cai=" & F_GravaData(txtbxData.Text)
    strSQL = strSQL & " group by esp_cai "
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            If .RecordCount > 0 Then
                If !Esp_cai = "DIN" Then
                    lblDin.caption = FormatVal(!total)
                End If
                If !Esp_cai = "FIN" Then
                    lblFin.caption = FormatVal(!total)
                End If
                If !Esp_cai = "CHP" Then
                    lblCHP.caption = FormatVal(!total)
                End If
                If !Esp_cai = "CHV" Then
                    lblCHV.caption = FormatVal(!total)
                End If
                If !Esp_cai = "CAR" Then
                    lblCar.caption = FormatVal(!total)
                End If
            End If
            
            .MoveNext
        Loop
        lblTotalAv.caption = FormatVal(CCur(lblDin.caption) + CCur(lblCHV.caption) + CCur(lblFin.caption))
        lblTotalAp.caption = FormatVal(CCur(lblCHP.caption) + CCur(lblCar.caption))
        lblTotalEntradas.caption = FormatVal(CCur(lblTotalAv.caption) + CCur(lblTotalAp.caption))
        .Close
        
    End With
    strSQL = "select sum(tc.valor_cai) as total from tab_Caixa as tc "
    strSQL = strSQL & " where tipo_cai='D' and data_cai=" & F_PesqData(txtbxData.Text)
    strSQL = strSQL & " and cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If Not IsNull(!total) Then
            lblTotalSai.caption = FormatVal(!total)
        End If
        .Close
    End With
    lblTotalAB.caption = FormatVal(CCur(lblTotalEntradas.caption) - CCur(lblTotalSai.caption))
    lblTotalDin.caption = FormatVal(CCur(lblDin.caption) - CCur(lblTotalSai.caption))
    lblTotalSai2.caption = lblTotalSai.caption
    'parei aki
    strSQL = "select sum(tc.valor_ven) as total from tab_cancelamentos  as tc,"
    strSQL = strSQL & "  " & TAB_VENFEC & "  as tf "
    strSQL = strSQL & " where tc.num_ven=tf.num_ven "
    strSQL = strSQL & " and left(data_can,10) = " & F_PesqData(txtbxData.Text)
    strSQL = strSQL & " and tf.cod_cai='" & Left(cmbxCaixa.Text, 2) & "' "
    strSQL = strSQL & " and tf.cod_tur='" & Left(cmbxTurno.Text, 1) & "' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If Not IsNull(!total) Then
            lblCancela.caption = FormatVal(!total)
        End If
        .Close
    End With
End Sub

Private Sub txtbxFundoCaixa_GotFocus()
    ganhaFoco txtbxFundoCaixa
End Sub

Private Sub txtbxFundoCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxFundoCaixa_LostFocus()
    PerdeFoco txtbxFundoCaixa
End Sub
