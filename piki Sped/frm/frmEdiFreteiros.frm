VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiFreteiros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Freteiros"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4830
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   3615
      Width           =   4815
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiFreteiros.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1680
         Picture         =   "frmEdiFreteiros.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   3240
         Picture         =   "frmEdiFreteiros.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Data dtServ 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2295
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiFreteiros.frx":085E
      Height          =   3615
      Left            =   0
      OleObjectBlob   =   "frmEdiFreteiros.frx":0873
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmEdiFreteiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadFreteiros
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    dtServ.Connect = banco.Connect
    strSQL = "select * from tab_freteiros  order by nome_fre "
    dtServ.RecordSource = strSQL
    dtServ.Refresh
End Sub
