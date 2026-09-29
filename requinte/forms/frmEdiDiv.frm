VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiDiv 
   Caption         =   "Edição de Divisões"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   Icon            =   "frmEdiDiv.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   4485
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   4455
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiDiv.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1560
         Picture         =   "frmEdiDiv.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   615
         Left            =   3000
         Picture         =   "frmEdiDiv.frx":0D9E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   1335
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmEdiDiv.frx":1128
      Height          =   2400
      Left            =   0
      OleObjectBlob   =   "frmEdiDiv.frx":113D
      TabIndex        =   4
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmEdiDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    If dtProd.Recordset.RecordCount > 0 Then
        frmIncDiv.Tag = dtProd.Recordset!cod_div
        frmIncDiv.txtbxNome.Text = dtProd.Recordset!nome_div
        frmIncDiv.cmbxGenero.Text = dtProd.Recordset!nome_gen
        frmIncDiv.cmbxGenero.Enabled = False
        frmIncDiv.Show 0
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    frmIncDiv.Show 0
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
    Me.Height = 3735
    Me.Width = 4605
    P_CentraForm Me
    
    dtProd.Connect = banco.Connect
    strSQL = "select td.*,tg.nome_gen from tab_divisoes as td,tab_generos as tg "
    strSQL = strSQL & " where tg.cod_gen=td.cod_gen order by nome_gen"
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub

