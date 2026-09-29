VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVendaFornadas 
   Caption         =   "Venda de Produtos"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7635
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmndbtnConfirma 
      Cancel          =   -1  'True
      Caption         =   "F&echar"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Frame frControle 
      Caption         =   "Distribuição das Fornadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox cmbxFornada 
         Height          =   315
         ItemData        =   "frmVendaFornadas.frx":0000
         Left            =   120
         List            =   "frmVendaFornadas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   465
         Width           =   2535
      End
      Begin VB.CommandButton cmndbtnOk 
         Caption         =   "&Adicionar"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtbxQtd_Itf 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2730
         TabIndex        =   4
         Text            =   "0,00"
         Top             =   480
         Width           =   735
      End
      Begin VB.Data dtProd 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid dbgrdProd 
         Bindings        =   "frmVendaFornadas.frx":0004
         Height          =   1695
         Left            =   120
         OleObjectBlob   =   "frmVendaFornadas.frx":0019
         TabIndex        =   7
         Top             =   840
         Width           =   7335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fornada:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Label lblEstoque 
      AutoSize        =   -1  'True
      Caption         =   "Estoque Desta Fornada:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   1725
   End
   Begin VB.Label lblVenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota:"
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   2760
      Width           =   390
   End
   Begin VB.Label lblQtd_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   3600
      TabIndex        =   13
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label lblDistr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   195
      Left            =   5040
      TabIndex        =   12
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Lançar:"
      Height          =   195
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lançado:"
      Height          =   195
      Left            =   4320
      TabIndex        =   9
      Top             =   2760
      Width           =   675
   End
End
Attribute VB_Name = "frmVendaFornadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxFornada_Click()
    
    lblEstoque.Tag = _
        F_BuscaUmCampo("Qtd_Pro", "tab_fornadas", "Cod_Pro='" & Me.Tag & _
            "' and fornada_pro='" & cmbxFornada.Text & "'")
    If lblEstoque.Tag = "" Then
        lblEstoque.Tag = "0"
    End If
    lblEstoque.caption = "Estoque Desta Fornada: " & lblEstoque.Tag
    
    
End Sub

Private Sub cmbxFornada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnConfirma_Click()
    Unload Me
End Sub

Private Sub cmndbtnExcl_Click()
    Dim strSQL As String
    
    With dtProd.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que excluir!", vbExclamation
            Exit Sub
        End If
        If dbgrdProd.Row = -1 Then
            MsgBox "Selecione um item!", vbExclamation
            dbgrdProd.SetFocus
            Exit Sub
        End If
        If MsgBox("confirma exclusão do lançamento?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "Delete from tab_venforn "
            strSQL = strSQL & " where num_ven='" & lblVenda.caption & "' "
            strSQL = strSQL & " and cod_loj = '" & strCod_Loj & "' "
            strSQL = strSQL & " and cod_pro = '" & Me.Tag & "' "
            strSQL = strSQL & " and fornada_ite = '" & !fornada_ite & "' "
            banco.Execute strSQL
            strSQL = "UPDATE tab_fornadas set reservado_pro= reservado_pro- "
            strSQL = strSQL & F_GravaDec(!qtd_ite)
            strSQL = strSQL & " where "
            'strSQL = strSQL & "  cod_loj = '" & strCod_Loj & "' and  "
            strSQL = strSQL & " cod_pro = '" & Me.Tag & "' "
            strSQL = strSQL & " and fornada_pro = '" & !fornada_ite & "' "
            banco.Execute strSQL
            
            
            busca
        End If
        
    End With
End Sub

Private Sub cmndbtnOk_Click()
    Dim strSQL As String
    If CCur(txtbxQtd_Itf.Text) + CCur(lblDistr.caption) > CCur(lblQtd_Ite) Then
        MsgBox "Valor está maior que o permitido!", vbExclamation
        txtbxQtd_Itf.SetFocus
        Exit Sub
    End If
    If CCur(txtbxQtd_Itf.Text) > CCur(lblEstoque.Tag) Then
        MsgBox "Valor está maior que o permitido!", vbExclamation
        txtbxQtd_Itf.SetFocus
        Exit Sub
    End If
    If F_Vazio(cmbxFornada.Text) Then
        MsgBox "Informe a Fornada!", vbExclamation
        txtbxForn.SetFocus
        Exit Sub
    End If
    If txtbxQtd_Itf.Text = "0,00" Then
        MsgBox "Informe uma quantidade válida!", vbExclamation
        txtbxQtd_Itf.SetFocus
        Exit Sub
    End If
    
    If F_BuscaUmCampo("Fornada_ite", "tab_venforn", "num_ven='" & lblVenda.caption & "'" _
    & "' and cod_pro='" & Me.Tag & "' and fornada_pro ='" & cmbxFornada.Text & "'") <> "" Then
        MsgBox "Já existe esta fornada para este produto!", vbExclamation
        txtbxForn.SetFocus
        Exit Sub
    End If
    
    
    If MsgBox("Confirma Inclusão?", vbYesNo + vbQuestion) = vbYes Then
        strSQL = "insert into tab_venforn(cod_loj,num_ven,"
        strSQL = strSQL & "cod_pro,fornada_ite,qtd_ite)values("
        strSQL = strSQL & "'" & strCod_Loj & "',"
        strSQL = strSQL & "'" & lblVenda.caption & "',"
        strSQL = strSQL & "'" & Me.Tag & "',"
        strSQL = strSQL & "'" & cmbxFornada.Text & "',"
        strSQL = strSQL & F_GravaDec(txtbxQtd_Itf.Text) & ") "
        banco.Execute strSQL
        strSQL = "UPDATE tab_fornadas set reservado_pro= reservado_pro+ "
        strSQL = strSQL & F_GravaDec(txtbxQtd_Itf.Text)
        strSQL = strSQL & " where "
        'strSQL = strSQL & "  cod_loj = '" & strCod_Loj & "' and  "
        strSQL = strSQL & " cod_pro = '" & Me.Tag & "' "
        strSQL = strSQL & " and fornada_pro = '" & cmbxFornada.Text & "' "
        banco.Execute strSQL
    End If
    busca
    txtbxQtd_Itf.Text = "0,00"
    cmbxFornada.SetFocus
    
End Sub
Sub busca()
    Dim strSQL As String
    Dim cTotal As Currency
    strSQL = "select * from tab_venforn "
    strSQL = strSQL & " where num_ven='" & lblVenda.caption & "' "
    strSQL = strSQL & " and cod_loj = '" & strCod_Loj & "' "
    strSQL = strSQL & " and cod_pro = '" & Me.Tag & "' "
    
    
    
    dtProd.RecordSource = strSQL
    dtProd.Refresh
    With dtProd.Recordset
        Do While Not .EOF
            
            cTotal = cTotal + !qtd_ite
            .MoveNext
        Loop
    End With
    lblDistr.caption = FormatVal(cTotal)
    
End Sub

Private Sub dbgrdProd_GotFocus()
    dbgrdProd.MarqueeStyle = 1
End Sub

Private Sub dbgrdProd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub dbgrdProd_LostFocus()
    dbgrdProd.MarqueeStyle = 0
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
    P_PreencheFornadas cmbxFornada, Me.Tag
    busca
End Sub

Private Sub Form_Load()
    dtProd.Connect = banco.Connect
End Sub


Private Sub txtbxQtd_Itf_GotFocus()
    P_SelecTexto txtbxQtd_Itf
End Sub

Private Sub txtbxQtd_Itf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_Itf_LostFocus()
    P_ValorLostFocus txtbxQtd_Itf
End Sub


