VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Fornecedores"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "frmEdiFor2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9060
   Begin VB.Frame frControle 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   5880
         TabIndex        =   9
         Top             =   4350
         Width           =   3015
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   735
            Left            =   2040
            Picture         =   "frmEdiFor2.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   735
            Left            =   1080
            Picture         =   "frmEdiFor2.frx":04D4
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo"
            Height          =   735
            Left            =   120
            Picture         =   "frmEdiFor2.frx":061E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   4560
         Width           =   5655
         Begin VB.CommandButton cmndbtnLocaliza 
            Height          =   285
            Left            =   3840
            Picture         =   "frmEdiFor2.frx":09A8
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   255
         End
         Begin VB.ComboBox cmbxOrd 
            Height          =   315
            ItemData        =   "frmEdiFor2.frx":0D32
            Left            =   4200
            List            =   "frmEdiFor2.frx":0D45
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbxBusca 
            Height          =   315
            ItemData        =   "frmEdiFor2.frx":0D71
            Left            =   120
            List            =   "frmEdiFor2.frx":0D81
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtbxExpr 
            Height          =   285
            Left            =   1800
            MaxLength       =   40
            TabIndex        =   5
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Ordenar:"
            Height          =   195
            Left            =   4200
            TabIndex        =   7
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E&xpressão:"
            Height          =   195
            Left            =   1800
            TabIndex        =   4
            Top             =   120
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Buscar Por:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4380
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   8775
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
         Begin MSDBGrid.DBGrid dbgrd 
            Bindings        =   "frmEdiFor2.frx":0DB0
            Height          =   4215
            Left            =   0
            OleObjectBlob   =   "frmEdiFor2.frx":0DC9
            TabIndex        =   14
            Top             =   0
            Width           =   8775
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Fornecedores: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   4200
            Width           =   1785
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Listados: 0"
            Height          =   195
            Left            =   3000
            TabIndex        =   15
            Top             =   4200
            Width           =   765
         End
      End
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
    
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & "Nome_for like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & "cod_for like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & "cnpj_for like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & "ie_for like '" & txtbxExpr.Text & "%' "
    End If
    
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
    If dtClientes.Recordset.RecordCount = 0 Then
        MsgBox "Não foi encontrado nenhum fornecedores na sua busca!", vbExclamation
    End If
    lblTOtal.caption = "Fornecedores Cadastrados...: " & F_BuscaUmCampo("count(cod_for)", "tab_fornecedores", "")
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
    Unload Me
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnEdit_Click
    End If
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    P_CentraForm Me
    dtClientes.Connect = banco.Connect
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0
    cmndbtnLocaliza_Click
       
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    If cmbxBusca.Text = "CNPJ" Then
       
    End If
End Sub

