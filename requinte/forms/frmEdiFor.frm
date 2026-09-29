VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiFor 
   Caption         =   "Edição de Fornecedores"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   Icon            =   "frmEdiFor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   12045
   Begin VB.PictureBox pcBusca 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   5925
      Width           =   12045
      Begin VB.ComboBox cmbxOrd 
         Height          =   315
         ItemData        =   "frmEdiFor.frx":014A
         Left            =   1800
         List            =   "frmEdiFor.frx":015D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbxBusca 
         Height          =   315
         ItemData        =   "frmEdiFor.frx":0189
         Left            =   120
         List            =   "frmEdiFor.frx":0199
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         MaxLength       =   40
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmndbtnLocaliza 
         Height          =   285
         Left            =   5760
         Picture         =   "frmEdiFor.frx":01C8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar:"
         Height          =   195
         Left            =   1800
         TabIndex        =   3
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E&xpressão:"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   825
      End
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listados: 0"
         Height          =   195
         Left            =   8400
         TabIndex        =   15
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Fornecedores: 0"
         Height          =   195
         Left            =   8400
         TabIndex        =   14
         Top             =   120
         Width           =   1785
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12045
      TabIndex        =   8
      Top             =   6660
      Width           =   12045
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "&Selecionar"
         Height          =   615
         Left            =   4440
         Picture         =   "frmEdiFor.frx":0573
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnImprimir 
         Caption         =   "Imprimir Grade"
         Height          =   615
         Left            =   7680
         Picture         =   "frmEdiFor.frx":08B5
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnCheques 
         Caption         =   "&Cheques"
         Height          =   615
         Left            =   6240
         Picture         =   "frmEdiFor.frx":0CC2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnDocs 
         Caption         =   "&Documentos"
         Height          =   615
         Left            =   3000
         Picture         =   "frmEdiFor.frx":16C4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   615
         Left            =   9480
         Picture         =   "frmEdiFor.frx":17AE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   1560
         Picture         =   "frmEdiFor.frx":19F8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   615
         Left            =   120
         Picture         =   "frmEdiFor.frx":1B42
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiFor.frx":1ECC
      Height          =   5895
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "cod_for"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cnpj_for"
         Caption         =   "CNPJ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "fantasia_for"
         Caption         =   "Fantasia"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "nome_for"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "cidade_for"
         Caption         =   "Cidade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3209,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2835,213
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ado As ADODB.Recordset
Attribute ado.VB_VarHelpID = -1
Dim bActive As Boolean
Private Sub cmndbtnEdit_Click()
    If ado.RecordCount <= 0 Then
        MsgBox "Não há o que ser editado!", vbCritical
        Exit Sub
    End If
    frmIncFor.tag = ado!cod_For
    frmIncFor.Show 1
End Sub

Private Sub cmndbtnImprimir_Click()
    P_DbgrdToHtml "Listagem de Fornecedores", dtgrd, ado
End Sub

Private Sub cmndbtnLocaliza_Click()
    Dim strSQL As String
    Dim sOrderBy As String
    
    strSQL = "select * from tab_fornecedores where "
    
    If cmbxBusca.ListIndex = 0 Then
        strSQL = strSQL & "Nome_for like '%" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 1 Then
        strSQL = strSQL & "cod_for like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 2 Then
        strSQL = strSQL & "cnpj_for like '" & txtbxExpr.Text & "%' "
    ElseIf cmbxBusca.ListIndex = 3 Then
        strSQL = strSQL & "ie_for like '" & txtbxExpr.Text & "%' "
    End If
    
    If cmbxOrd.ListIndex = 0 Then
        sOrderBy = "cod_for"
    ElseIf cmbxOrd.ListIndex = 1 Then
        sOrderBy = "Nome_for,cnpj_for"
    ElseIf cmbxOrd.ListIndex = 2 Then
        sOrderBy = "CNPJ_for"
    ElseIf cmbxOrd.ListIndex = 3 Then
        sOrderBy = "Cidade_for,uf_for,Nome_for"
    ElseIf cmbxOrd.ListIndex = 4 Then
        sOrderBy = "UF_for,Cidade_for"
    End If
    strSQL = strSQL & " order by " & sOrderBy
    abreObjetoAdo ado, strSQL, dtgrd
    
    If ado.RecordCount = 0 Then
        MsgBox "Não foi encontrado nenhum fornecedores na sua busca!", vbExclamation
    ElseIf ado.RecordCount = 1 And Me.tag <> "" Then
        cmdSelecionar_Click
    End If
    
    lblTotal.caption = "Fornecedores Cadastrados...: " & F_BuscaUmCampo("count(cod_for)", "tab_fornecedores", "")
    lblList.caption = "Listados...: " & ado.RecordCount
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
    If Me.tag = "" Then
        Unload Me
    Else
        Me.tag = ""
        Me.Hide
    End If
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
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


Private Sub dtgrd_DblClick()
    If Me.tag = "" Then
        cmndbtnEdit_Click
    Else
        cmdSelecionar_Click
    End If
End Sub
Sub cmdSelecionar_Click()
        On Error Resume Next
    If Me.tag <> "" Then
        Me.tag = ado!cod_For
        Me.Hide
    End If

End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub Form_Activate()
    If bActive Then Exit Sub
    If Me.tag <> "" Then
        cmndbtnLocaliza_Click
        
    End If
    bActive = True
    Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    
    resizeForm Me
    bActive = False
    Set ado = New ADODB.Recordset
    
    cmbxBusca.ListIndex = 0
    cmbxOrd.ListIndex = 0

       
End Sub

Private Sub Form_Resize()
    dtgrd.width = Me.width - 300
    dtgrd.height = pcBusca.top - 300
    cmndbtnSair.Left = Me.width - cmndbtnSair.width - 200
    
    
    PosicionaGrade dtgrd, Me, pcBusca
    
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnLocaliza_Click
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        dtgrd.SetFocus
    End If
    
End Sub

Private Sub txtbxExpr_LostFocus()
    P_AjustaTexto txtbxExpr
    If cmbxBusca.Text = "CNPJ" Then
       
    End If
End Sub

