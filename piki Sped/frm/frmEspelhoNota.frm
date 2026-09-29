VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEspelhoNota 
   Caption         =   "Espelho de Nota Fiscal"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   Icon            =   "frmespelhonota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9855
   Begin VB.TextBox txtbxSerie 
      Height          =   285
      Left            =   6960
      TabIndex        =   26
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtbxNumero 
      Height          =   285
      Left            =   8400
      TabIndex        =   24
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtbxData 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtbxGrid 
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Tag             =   "!"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmndbtnFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   8400
      Picture         =   "frmespelhonota.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmndbtnConcluir 
      Caption         =   "&Concluir"
      Height          =   615
      Left            =   6960
      Picture         =   "frmespelhonota.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid msflxGrd 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7011
      _Version        =   393216
      GridColor       =   16777215
      GridColorFixed  =   16777215
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.TextBox txtbxNomeCli 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Tag             =   "00000"
      Text            =   "Venda a Consumidor"
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox txtbxVenda 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   9615
      Begin VB.CommandButton cmndbtnAdd 
         Caption         =   "&Adicionar"
         Height          =   375
         Left            =   7920
         TabIndex        =   17
         Top             =   315
         Width           =   975
      End
      Begin VB.TextBox txtbxTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   16
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   5640
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtbxQtd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4800
         TabIndex        =   12
         Text            =   "1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbxProduto 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblMarca 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         Height          =   195
         Left            =   4080
         TabIndex        =   28
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6720
         TabIndex        =   15
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Unitário:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   13
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Qtd:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4800
         TabIndex        =   11
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Produto"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   555
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Série da Nota:"
      Height          =   195
      Left            =   6960
      TabIndex        =   27
      Top             =   4920
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da Nota:"
      Height          =   195
      Left            =   8400
      TabIndex        =   25
      Top             =   4920
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9840
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data/Hora:"
      Height          =   195
      Left            =   1560
      TabIndex        =   23
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label lbltTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   21
      Top             =   5640
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   20
      Top             =   5640
      Width           =   750
   End
   Begin VB.Label lblItens 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   19
      Top             =   5280
      Width           =   165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Itens:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   18
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   510
   End
End
Attribute VB_Name = "frmEspelhoNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmndbtnAdd_Click()
    Dim strSQL As String
    If txtbxProduto.tag = "" Then
        MsgBox "Informe o produto!", vbExclamation
        txtbxProduto.SetFocus
        Exit Sub
    End If
    If CCur(txtbxTotal.Text) = 0 Then
        MsgBox "informe o valor corretamente!", vbCritical
        txtbxQtd.SetFocus
        Exit Sub
    End If
    If F_InFlxgrd(msflxGrd, txtbxProduto.tag, 0) > 0 Then
        MsgBox "Este produto já existe na nota!", vbExclamation
        Exit Sub
    End If
     strSQL = txtbxProduto.tag & Chr(9) & txtbxProduto.Text & Chr(9)
    strSQL = strSQL & lblMarca.caption & Chr(9)
    strSQL = strSQL & txtbxQtd.Text & Chr(9)
    strSQL = strSQL & (txtbxUnit.Text) & Chr(9)
    strSQL = strSQL & (txtbxTotal.Text) & Chr(9)
    strSQL = strSQL & FormatVal(txtbxUnit.tag) & Chr(9)
    strSQL = strSQL & FormatVal(txtbxTotal.tag)
    
    msflxGrd.AddItem strSQL
     soma
    ZebraGrid msflxGrd, 1
    txtbxProduto.tag = ""
    txtbxProduto.SetFocus
End Sub

Private Sub cmndbtnConcluir_Click()
    Dim iCount As Integer
    Dim strSQL As String
    Dim sQtd As String
    Dim sCod As String
    On Error GoTo trata_erro
    If MsgBox("Confirma conclusão da Nota?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If txtbxNomeCli.tag = "" Then
        MsgBox "Informe um cliente!", vbExclamation
        txtbxNomeCli.SetFocus
        Exit Sub
    End If
    If F_Vazio(txtbxSerie.Text) Then
        MsgBox "Informe a série da nota!", vbCritical
        txtbxSerie.SetFocus
        Exit Sub
    End If
    If F_Vazio(txtbxNumero.Text) Then
        MsgBox "Informe o número da Nota!", vbCritical
        txtbxNumero.SetFocus
        Exit Sub
    End If
    If msflxGrd.rows = 1 Then
        MsgBox "Não há itens em sua Nota!", vbCritical
        Exit Sub
    End If
    strSQL = "insert into tab_saidas(cod_cli,nota_sai,serie_sai,valor_sai,desconto_sai,icms_sai,num_ven,cod_usu,emissao_sai,cod_cai)values("
    strSQL = strSQL & "'" & txtbxNomeCli.tag & "',"
    strSQL = strSQL & "'" & txtbxNumero.Text & "',"
    strSQL = strSQL & "'" & txtbxSerie.Text & "',"
    strSQL = strSQL & F_GravaDec(lbltTotal.caption) & ","
    strSQL = strSQL & "0,"
    strSQL = strSQL & "0,"
    strSQL = strSQL & "'" & txtbxVenda.Text & "',"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'01')"
    Banco.Execute strSQL
    With msflxGrd
        If .rows = 1 Then
            
            MsgBox "Não há itens em sua nota!", vbCritical
            Exit Sub
        End If
        For iCount = 1 To .rows - 1
        .Row = iCount
            strSQL = "insert into tab_itsai(nota_sai,serie_sai,qtd_its,unit_its,"
            strSQL = strSQL & "cod_pro,seq_its,cod_Cai)values("
            strSQL = strSQL & "'" & txtbxNumero.Text & "',"
            strSQL = strSQL & "'" & txtbxSerie.Text & "',"
            .Col = 3: sQtd = .Text
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .Col = 4
            strSQL = strSQL & F_GravaDec(.Text) & ","
            .Col = 0: sCod = .Text
            strSQL = strSQL & "'" & .Text & "',"
            strSQL = strSQL & iCount & ","
            strSQL = strSQL & "'01')"
            Banco.Execute strSQL
            strSQL = "update tab_produtos set ef_pro=ef_pro -" & F_GravaDec(sQtd)
            strSQL = strSQL & " where cod_pro='" & sCod & "' "
            Banco.Execute strSQL
            
            
            
        Next
        
        MsgBox "Nota Fiscal Concluída!", vbCritical
        msflxGrd.rows = 1
        txtbxVenda.Text = ""
        txtbxVenda.SetFocus
    End With
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Private Sub cmndbtnFechar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_Load()
    Me.Width = 9975
    Me.Height = 6990
    P_CentraForm Me
    With msflxGrd
        .FixedCols = 0
        .rows = 1
        .Cols = 8
        .Col = 0: .Text = "Cód:"
        .Col = 1: .Text = "Descrição:"
        .Col = 2: .Text = "Marca:"
        .Col = 3: .Text = "Qtd:"
        .Col = 4: .Text = "Unit.:"
        .Col = 5: .Text = "Total.:"
        .Col = 6: .Text = "Qtd F.:"
        .Col = 7: .Text = "Qtd R.:"
        .ColWidth(0) = 1000
        .ColWidth(1) = 3000
        .ColWidth(2) = 1500
    End With
    txtbxData.Text = Now
End Sub

Private Sub msflxgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With msflxGrd
            If .Col = 3 Or .Col = 4 Then
                txtbxGrid.Text = .Text
                txtbxGrid.Left = msflxGrd.Left + .CellLeft
                txtbxGrid.Top = msflxGrd.Top + .CellTop
                txtbxGrid.Width = .CellWidth
                txtbxGrid.Height = .CellHeight
                txtbxGrid.Visible = True
                txtbxGrid.SetFocus
            End If
        End With
    End If
    If KeyCode = vbKeyDelete Then
        With msflxGrd
            If MsgBox("Confirma Exclusão do Item?", vbYesNo + vbCritical) = vbNo Then
                Exit Sub
            End If
            If .rows = 2 Then
                .rows = 1
            Else
                .RemoveItem .Row
            End If
            soma
            ZebraGrid msflxGrd, 1
        End With
    End If
End Sub

Private Sub txtbxGrid_GotFocus()
    P_SelecTexto txtbxGrid
End Sub

Private Sub txtbxGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        P_ValorLostFocus txtbxGrid
        msflxGrd.Text = txtbxGrid.Text
        somaItem
        soma
        txtbxGrid_LostFocus
        
    End If
End Sub
Sub somaItem()
    Dim sValor As String
    Dim sQtd As String
    With msflxGrd
        .Col = 3
        sQtd = .Text
        .Col = 4: sValor = .Text
        .Col = 5
        .Text = FormatVal(CCur(sValor) * CCur(sQtd))
    End With
    
End Sub
Sub soma()
    Dim strSQL As String
    Dim iCount As Integer
    Dim iRow As Integer
    Dim cTotal As Currency
    With msflxGrd
        iRow = .Row
        For iCount = 1 To .rows - 1
            .Row = iCount
            .Col = 5
            cTotal = cTotal + CCur(.Text)
            
        Next
        lbltTotal.caption = cTotal
        lblItens.caption = .rows - 1
        .Row = iRow
    End With
End Sub

Private Sub txtbxGrid_LostFocus()
    
    txtbxGrid.Visible = False
    msflxGrd.SetFocus
    
End Sub



Private Sub txtbxNomeCli_GotFocus()
    P_SelecTexto txtbxNomeCli
End Sub

Private Sub txtbxNomeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        buscaCliente
        SendKeys Chr(9)
    End If
End Sub
Sub buscaCliente()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxNomeCli.Text)
    txtbxNomeCli.Text = c.nome_Cli
    txtbxNomeCli.tag = c.cod_cli
End Sub
Private Sub txtbxNomeCli_LostFocus()
    P_AjustaTexto txtbxNomeCli
End Sub

Private Sub txtbxNumero_GotFocus()
    P_SelecTexto txtbxNumero
End Sub

Private Sub txtbxNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxNumero_LostFocus()
    P_AjustaNota txtbxNumero
End Sub

Private Sub txtbxProduto_Change()
'    txtbxProduto.Tag = ""
End Sub

Private Sub txtbxProduto_GotFocus()
    P_SelecTexto txtbxProduto
End Sub

Private Sub txtbxProduto_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        buscaProd
    End If
End Sub
Sub buscaProd()
    Dim p As ProdutoSaida
    p = F_BuscaProdSaida(txtbxProduto.Text)
    If p.cod_pro <> "" Then
        txtbxProduto.tag = p.cod_pro
        txtbxProduto.Text = p.nome_pro
        txtbxUnit.Text = p.p2_pro
        txtbxUnit.tag = p.e1_pro
        txtbxTotal.tag = p.e2_pro
        
        TotalItem
        txtbxQtd.SetFocus
    End If
End Sub
Sub TotalItem()
    txtbxTotal.Text = FormatVal(CCur(txtbxQtd.Text) * CCur(txtbxUnit.Text))
    
End Sub

Private Sub txtbxQtd_GotFocus()
    P_SelecTexto txtbxQtd
End Sub

Private Sub txtbxQtd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxQtd_LostFocus()
    P_ValorLostFocus txtbxQtd
    TotalItem
End Sub

Private Sub txtbxSerie_GotFocus()
    P_SelecTexto txtbxSerie
End Sub

Private Sub txtbxSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtbxSerie_LostFocus
        txtbxNumero.Text = MaxKeySQL("tab_saidas", "nota_sai", "serie_sai='" & txtbxSerie.Text & "'", "0000000#")
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxSerie_LostFocus()
    txtbxNumero.Text = MaxKeySQL("tab_saidas", "nota_sai", "serie_sai='" & txtbxSerie.Text & "'", "0000000#")
    P_AjustaTexto txtbxSerie
End Sub

Private Sub txtbxUnit_GotFocus()
    P_SelecTexto txtbxUnit
End Sub

Private Sub txtbxUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxUnit_LostFocus()
    P_ValorLostFocus txtbxUnit
    TotalItem
End Sub

Private Sub txtbxVenda_GotFocus()
    P_SelecTexto txtbxVenda
End Sub

Private Sub txtbxVenda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        txtbxVenda_LostFocus
        Me.tag = txtbxVenda.Text
        busca
    End If
End Sub
Sub busca()
    Dim strSQL As String
    Dim rPesq As Recordset
    Dim cTotal As Currency
    msflxGrd.rows = 1
    strSQL = "select tp.*,tv.cod_cli,tm.nome_mar,"
    strSQL = strSQL & " tv.nomecli_ven,tv.data_ven,ti.* "
    strSQL = strSQL & " from tab_vendas as tv, tab_itvendas as ti, "
    strSQL = strSQL & " tab_marcas as tm, tab_produtos as tp "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and tv.num_ven='" & Me.tag & "' "
    strSQL = strSQL & " and tm.cod_mar=tp.cod_mar "
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    
    Set rPesq = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
    
        If .RecordCount > 0 Then
            txtbxNomeCli.Text = !nomecli_ven
            txtbxNomeCli.tag = !cod_cli
            txtbxData.Text = !data_ven
        End If
        Do While Not .EOF
            strSQL = !cod_pro & Chr(9) & !nome_pro & Chr(9)
            strSQL = strSQL & !nome_mar & Chr(9)
            strSQL = strSQL & !qtd_ite & Chr(9)
            strSQL = strSQL & FormatVal(!unit_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!qtd_ite * !unit_ite) & Chr(9)
            strSQL = strSQL & FormatVal(!ef_pro) & Chr(9)
            strSQL = strSQL & FormatVal(!er_pro)
            msflxGrd.AddItem strSQL
            .MoveNext
        Loop
        .Close
    End With
    soma
    ZebraGrid msflxGrd, 1
End Sub
Private Sub txtbxVenda_LostFocus()
    txtbxVenda.Text = Format(txtbxVenda.Text, "000000#")
End Sub
