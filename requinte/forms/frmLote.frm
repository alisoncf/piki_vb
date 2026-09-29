VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLote 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6105
      ScaleWidth      =   7275
      TabIndex        =   5
      Top             =   0
      Width           =   7305
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Buscar:"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   5520
         Width           =   6975
         Begin VB.OptionButton optnTodos 
            Caption         =   "Todos os Produtos"
            Height          =   195
            Left            =   4080
            TabIndex        =   4
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton optnZer 
            Caption         =   "Produtos em Estoque"
            Height          =   195
            Left            =   2160
            TabIndex        =   3
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton optnEst 
            Caption         =   "Produto com Estoque"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Data dtPro 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3360
         Visible         =   0   'False
         Width           =   3135
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmLote.frx":0000
         Height          =   4695
         Left            =   120
         OleObjectBlob   =   "frmLote.frx":0014
         TabIndex        =   0
         Top             =   720
         Width           =   6975
      End
      Begin VB.Label lblProduto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean
Private Sub dbgrd_DblClick()
    Escolha
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Escolha
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    lblProduto.caption = F_BuscaUmCampo("nome_pro", "tab_produtos", "cod_pro='" & Me.tag & "'")
    busca
    bActive = True
End Sub
Sub Escolha()
    On Error Resume Next
     If dbgrd.Row = -1 Then
            Me.tag = ""
        Else
            
        End If
        Me.Hide
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Escolha
    ElseIf KeyCode = vbKeyEscape Then
        Me.tag = ""
        Me.Hide
        
    End If
End Sub

Private Sub Form_Load()
    bActive = False
    dtPro.Connect = Banco.Connect
End Sub

Private Sub optnEst_Click()
    busca
End Sub

Private Sub optnTodos_Click()
    busca
End Sub

Private Sub optnZer_Click()
    busca
End Sub

Sub busca()
    Dim strSQL As String
    strSQL = "select * from tab_Estoque where cod_pro='" & Me.tag & "'"
    dtPro.RecordSource = strSQL
    dtPro.Refresh
    If dtPro.Recordset.RecordCount = 0 Then
'        MsgBox "Existe produto não possui qualquer lote em estoque", vbExclamation
        Me.Hide
    ElseIf dtPro.Recordset.RecordCount = 1 Then
        Me.Hide
    End If
End Sub
