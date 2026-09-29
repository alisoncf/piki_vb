VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Compras por Fornecedor"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   7200
         Picture         =   "frmRelCompras.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   5640
         Picture         =   "frmRelCompras.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbxForn 
         Height          =   315
         ItemData        =   "frmRelCompras.frx":0BD4
         Left            =   120
         List            =   "frmRelCompras.frx":0BE1
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   4335
      End
      Begin VB.OptionButton optnNome 
         Caption         =   "Ordernar por &Nome"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ordernar por &Código"
         Height          =   195
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Marca:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Data dtSai 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmRelCompras.frx":0C04
      Height          =   4095
      Left            =   0
      OleObjectBlob   =   "frmRelCompras.frx":0C15
      TabIndex        =   7
      Top             =   1200
      Width           =   8775
   End
   Begin MSDBGrid.DBGrid dbgrdEntr 
      Bindings        =   "frmRelCompras.frx":21BA
      Height          =   855
      Left            =   0
      OleObjectBlob   =   "frmRelCompras.frx":21CE
      TabIndex        =   8
      Top             =   5400
      Width           =   8775
   End
End
Attribute VB_Name = "frmRelCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
