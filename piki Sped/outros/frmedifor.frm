VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiFor 
   Caption         =   "Edição de Fornecedores"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
   Icon            =   "frmedifor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   10890
   Begin VB.Data dtClientes 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10890
      TabIndex        =   0
      Top             =   6030
      Width           =   10890
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmedifor.frx":014A
         Left            =   120
         List            =   "frmedifor.frx":015D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   4
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   285
         Left            =   5640
         Picture         =   "frmedifor.frx":0189
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   780
      End
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados: 0"
         Height          =   195
         Left            =   8400
         TabIndex        =   13
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Fornecedores: 0"
         Height          =   195
         Left            =   8400
         TabIndex        =   12
         Top             =   0
         Width           =   1785
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   10860
      TabIndex        =   6
      Top             =   6765
      Width           =   10890
      Begin VB.CommandButton cmndbtnSelect 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   2880
         Picture         =   "frmedifor.frx":0513
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cheques"
         Height          =   615
         Left            =   5760
         Picture         =   "frmedifor.frx":0A9D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDocs 
         Caption         =   "&Documentos"
         Height          =   615
         Left            =   4320
         Picture         =   "frmedifor.frx":1027
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   7800
         Picture         =   "frmedifor.frx":1CF1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1440
         Picture         =   "frmedifor.frx":1E5C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   0
         Picture         =   "frmedifor.frx":1FA6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   1335
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmedifor.frx":2330
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "frmedifor.frx":2349
      TabIndex        =   14
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmEdiFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmndbtnEdit_Click()
    If dtClientes.Recordset.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    frmIncFor.Tag = dtClientes.Recordset!cod_for
    frmIncFor.Show 0
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderby As String
    
    strSQL = "select * from tab_fornecedores where "
    strSQL = strSQL & "(Nome_for like '%" & txtbxExpr.Text & "%' "
    strSQL = strSQL & "or cod_for like '" & txtbxExpr.Text & "%' "
    strSQL = strSQL & "or cnpj_for like '" & txtbxExpr.Text & "%' "
    strSQL = strSQL & "or ie_for like '" & txtbxExpr.Text & "%') "
        
        
    If cmbxOrd.ListIndex = 0 Then
        sOrderby = "cod_for"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderby = "Nome_for,cnpj_for"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderby = "CNPJ_for"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderby = "Cidade_for,uf_for,NOme_for"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderby = "UF_for,Cidade_for"
    End If
    strSQL = strSQL & " order by " & sOrderby
    dtClientes.RecordSource = strSQL
    dtClientes.Refresh
    
    lblTotal.caption = "Fornecedores Cadastrados...: " & F_BuscaUmCampo("count(cod_for)", "tab_fornecedores", "")
    lblList.caption = "Listados...: " & dtClientes.Recordset.RecordCount
End Sub

Private Sub cmndbtnNew_Click()
    frmIncFor.Show 0
    If frmIncFor.Visible = False Then
        frmIncFor.Visible = True
    End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmndbtnSair_Click()
    If Me.Tag = "" Then
    
        Unload Me
    Else
        Me.Tag = ""
        Me.Hide
    End If
End Sub

Private Sub cmndbtnSelect_Click()
    Me.Tag = "x"
    Me.Hide
End Sub

Private Sub dbgrd_DblClick()
    If Me.Tag = "" Then
        cmndbtnEdit_Click
    Else
        cmndbtnSelect_Click
    End If
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dbgrd_DblClick
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    
    
    cmbxOrd.ListIndex = 0
    cmndbtnLocaliza_Click
    
    If Me.Tag = "" Then
        cmndbtnSelect.Enabled = False
    Else
        On Error Resume Next
        cmndbtnSelect.Enabled = True
        dbgrd.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    resizeForm Me
    
    dtClientes.Connect = banco.Connect
    
    
       
    
End Sub

Private Sub Form_Resize()
    dbgrd.Width = Me.Width - 300
    dbgrd.Height = pcBusca.Top - 300
    cmndbtnSair.Left = Me.Width - cmndbtnSair.Width - 200
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    'If cmbxBusca.Text = "CNPJ" Then
       
    'End If
End Sub

