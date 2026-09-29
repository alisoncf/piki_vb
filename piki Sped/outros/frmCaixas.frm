VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCaixas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caixas Cadastrados"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmndbtnSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   735
      Left            =   6240
      Picture         =   "frmCaixas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnEditar 
      Caption         =   "&Editar"
      Height          =   735
      Left            =   1440
      Picture         =   "frmCaixas.frx":016B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmndbtnNovo 
      Caption         =   "&Novo Caixa"
      Height          =   735
      Left            =   120
      Picture         =   "frmCaixas.frx":0255
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Data dt 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmCaixas.frx":04C9
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frmCaixas.frx":04DA
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmCaixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    busca
End Sub

Private Sub Form_Load()
    dt.Connect = Banco.Connect
    
End Sub
Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_caixas order by cod_cai "
    dt.RecordSource = strSQL
    dt.Refresh
    
    
    
End Sub
