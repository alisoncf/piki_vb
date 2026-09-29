VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRelVendasPeriodo 
   Caption         =   "Relatório de Frequência de Vendas"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   14190
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grade 
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   14160
      TabIndex        =   0
      Top             =   0
      Width           =   14190
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxDAtaFIn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Pesquisar"
         Height          =   615
         Left            =   8760
         Picture         =   "frmRelVendasPeriodo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   10080
         Picture         =   "frmRelVendasPeriodo.frx":0350
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   615
         Left            =   12720
         Picture         =   "frmRelVendasPeriodo.frx":083E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   595
         Left            =   11400
         Picture         =   "frmRelVendasPeriodo.frx":0DCC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   3735
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   315
         Left            =   3840
         Picture         =   "frmRelVendasPeriodo.frx":11FB
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data &Final:"
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Categoria"
         Height          =   195
         Left            =   4200
         TabIndex        =   11
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   195
         Left            =   6120
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmRelVendasPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
