VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiRepCco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Movimento Corrente de Representantes"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcBotoes 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7155
      Left            =   11655
      ScaleHeight     =   7125
      ScaleWidth      =   1665
      TabIndex        =   14
      Top             =   735
      Width           =   1695
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnAcertos 
         Caption         =   "Acertos"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnDeb 
         Caption         =   "&Débitos"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":0589
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnAcerto 
         Caption         =   "&Acerto"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":06D3
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Devolução"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":0F9D
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":1527
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7920
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnResumoDeb 
         Caption         =   "Resumo Débitos"
         Height          =   695
         Left            =   120
         Picture         =   "frmEdiRepCco.frx":1671
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4320
         Width           =   1455
      End
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   13320
      TabIndex        =   0
      Top             =   0
      Width           =   13350
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   450
         Left            =   12000
         Picture         =   "frmEdiRepCco.frx":17BB
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiRepCco.frx":1B45
         Left            =   10440
         List            =   "frmEdiRepCco.frx":1B58
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtbxAte 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8280
         MaxLength       =   40
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxDe 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbxNat 
         Height          =   315
         ItemData        =   "frmEdiRepCco.frx":1B7F
         Left            =   9360
         List            =   "frmEdiRepCco.frx":1B8C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxRepresentante 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4560
         MaxLength       =   40
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmndbtnBuscaRep 
         Height          =   330
         Left            =   4200
         Picture         =   "frmEdiRepCco.frx":1BAA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   8280
         TabIndex        =   12
         Top             =   0
         Width           =   285
      End
      Begin VB.Label lblDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   7200
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Natureza:"
         Height          =   195
         Left            =   9360
         TabIndex        =   6
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Representante:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   10440
         TabIndex        =   8
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Histórico:"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   0
         Width           =   660
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiRepCco.frx":1CA3
      Height          =   6975
      Left            =   0
      OleObjectBlob   =   "frmEdiRepCco.frx":1CB4
      TabIndex        =   23
      Top             =   720
      Width           =   11535
   End
End
Attribute VB_Name = "frmEdiRepCco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
