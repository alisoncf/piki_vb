VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Categorias"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmEdiCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6585
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   4680
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferência"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiCat.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiCat.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiCat.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   735
         Left            =   120
         Picture         =   "frmEdiCat.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1575
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
      Bindings        =   "frmEdiCat.frx":08B2
      Height          =   3375
      Left            =   0
      OleObjectBlob   =   "frmEdiCat.frx":08C7
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmEdiCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmndbtnEdit_Click()
    Dim fr As New frmCadCat
    If dtProd.Recordset.RecordCount > 0 Then
        fr.Tag = dtProd.Recordset!cod_cat
        fr.txtbxNome.Text = dtProd.Recordset!nome_cat
        fr.Show 0
    Else
        MsgBox "Não há o que ser Editado!", vbExclamation
    End If
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmCadCat
    frmCadCat.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    If IsLoad("frmCadPro") Then
        P_PreencheCateg frmCadPro.cmbxCat
        P_PreencheMarca frmCadPro.cmbxMarca
    End If
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
    'Me.Height = 3870
    'Me.Width = 3570
    P_CentraForm Me
    dtProd.Connect = Banco.Connect
    strSQL = "select * from tab_categorias order by nome_cat "
    dtProd.RecordSource = strSQL
    dtProd.Refresh
End Sub

