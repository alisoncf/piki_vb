VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiGen 
   Caption         =   "Edição de Gêneros"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   Icon            =   "frmEdiGen.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   3450
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   3375
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   2400
         Picture         =   "frmEdiGen.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   1080
         Picture         =   "frmEdiGen.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiGen.frx":0D9E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.Data dtProd 
      Caption         =   "Edição de Gêneros"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmEdiGen.frx":1128
      Height          =   2295
      Left            =   0
      OleObjectBlob   =   "frmEdiGen.frx":113D
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmEdiGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    If dtProd.Recordset.RecordCount > 0 Then
        frmIncGen.Tag = dtProd.Recordset!cod_gen
        frmIncGen.txtbxNome.Text = dtProd.Recordset!nome_gen
        frmIncGen.Show 0
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    frmIncGen.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.MarqueeStyle = 3
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Me.Height = 3870
    Me.Width = 3570
    P_CentraForm Me
    dtProd.Connect = banco.Connect
    strSQL = "select * from tab_generos"
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub
