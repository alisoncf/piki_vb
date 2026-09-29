VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiDes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Descontos e Comissões Decrescentes:"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6900
   Begin VB.Data dtServ 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiDes.frx":0000
      Height          =   3015
      Left            =   0
      OleObjectBlob   =   "frmEdiDes.frx":0015
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   6855
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   5160
         Picture         =   "frmEdiDes.frx":0BAD
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1800
         Picture         =   "frmEdiDes.frx":0CF7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiDes.frx":0E41
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnDesab 
         Caption         =   "E&xcluir"
         Height          =   615
         Left            =   3480
         Picture         =   "frmEdiDes.frx":11CB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEdiDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmDescontos
    With dtServ.Recordset
        fr.txtbxComiss_Des.Text = FormatVal(!comiss_des)
        fr.txtbxValor_Des.Text = FormatVal(!valor_des)
        fr.Tag = !cod_des
        fr.Show 0
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmDescontos
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    P_CentraForm Me
    dtServ.Connect = banco.Connect
    strSQL = "select * from tab_descontos "
    dtServ.RecordSource = strSQL
    dtServ.Refresh
    
End Sub
