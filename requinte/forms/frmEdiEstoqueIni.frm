VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiEstoqueIni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Ajustes"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Editar"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Data dtPro 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiEstoqueIni.frx":0000
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frmEdiEstoqueIni.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmEdiEstoqueIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    dtPro.Connect = banco.Connect
End Sub
Sub Busca()
    Dim strSQL As String
    strSQL = "select tp*,te.qtd_est from tab_produtos as tp,"
    strSQL = strSQL & "tab_ajest as te where te.cod_pro=tp.cod_pro"
    strSQL = strSQL & "order by nome_pro "
    dtPro.RecordSource = strSQL
    dtPro.Refresh
End Sub

