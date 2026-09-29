VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelVendasImpostos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relatório de Vendas de Produtos por ICMS"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frControle 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2175
         Begin VB.TextBox txtbxDataIni 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxdataFin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1200
            TabIndex        =   17
            Top             =   0
            Width           =   285
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Ordenar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2175
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmRelVendasImpostos.frx":0000
            Left            =   120
            List            =   "frmRelVendasImpostos.frx":0010
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Data dtNot 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelVendasImpostos.frx":003C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Im&primir"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelVendasImpostos.frx":0906
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   120
         Picture         =   "frmRelVendasImpostos.frx":0C10
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         ScaleHeight     =   345
         ScaleWidth      =   7545
         TabIndex        =   1
         Top             =   5760
         Width           =   7575
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   38
            Width           =   510
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   720
            TabIndex        =   6
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblJuros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   3480
            TabIndex        =   5
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Juros:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   4
            Top             =   45
            Width           =   525
         End
         Begin VB.Label lblSug 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H80000007&
            Height          =   360
            Left            =   6120
            TabIndex        =   3
            Top             =   0
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sugerido:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5160
            TabIndex        =   2
            Top             =   45
            Width           =   855
         End
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmRelVendasImpostos.frx":0F1A
         Height          =   5295
         Left            =   2400
         OleObjectBlob   =   "frmRelVendasImpostos.frx":0F2E
         TabIndex        =   8
         Top             =   240
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmRelVendasImpostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
