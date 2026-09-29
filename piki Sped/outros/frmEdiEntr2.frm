VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiEntr2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edição de Compras"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   Icon            =   "frmEdiEntr2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10515
   Begin VB.Frame frCOntrole 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame1 
         Caption         =   "&Situação:"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2295
         Begin VB.ComboBox cmbxTipo 
            Height          =   315
            ItemData        =   "frmEdiEntr2.frx":014A
            Left            =   120
            List            =   "frmEdiEntr2.frx":015D
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "&Buscar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2295
         Begin VB.ComboBox cmbxBuscar 
            Height          =   315
            ItemData        =   "frmEdiEntr2.frx":01CD
            Left            =   120
            List            =   "frmEdiEntr2.frx":01D7
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   2295
         Begin VB.CommandButton Command1 
            Caption         =   "&Editar"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEntr2.frx":01EF
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2400
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnCancelar 
            Caption         =   "&Cancelar"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEntr2.frx":0339
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnNew 
            Caption         =   "&Novo"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEntr2.frx":0483
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnEdit 
            Caption         =   "&Editar"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEntr2.frx":080D
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   960
            Width           =   2055
         End
         Begin VB.CommandButton cmndbtnSair 
            Cancel          =   -1  'True
            Caption         =   "&Fechar"
            Height          =   685
            Left            =   120
            Picture         =   "frmEdiEntr2.frx":0957
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3120
            Width           =   2055
         End
      End
      Begin VB.Data dtNot 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   1  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
         Begin VB.TextBox txtbxDataIni 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtbxdataFin 
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&De:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Até:"
            Height          =   195
            Left            =   1200
            TabIndex        =   8
            Top             =   0
            Width           =   285
         End
      End
      Begin MSDBGrid.DBGrid dbgrd 
         Bindings        =   "frmEdiEntr2.frx":0AA1
         Height          =   6015
         Left            =   2520
         OleObjectBlob   =   "frmEdiEntr2.frx":0AB5
         TabIndex        =   20
         Top             =   240
         Width           =   7575
      End
      Begin VB.Frame Frame6 
         Caption         =   "E&xpressão:"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2295
         Begin VB.CommandButton cmndbtnTerc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            Picture         =   "frmEdiEntr2.frx":23AE
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtbxExpr 
            Height          =   285
            Left            =   120
            MaxLength       =   40
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Ordenar Por:"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   2295
         Begin VB.ComboBox cmbxOrder 
            Height          =   315
            ItemData        =   "frmEdiEntr2.frx":2738
            Left            =   120
            List            =   "frmEdiEntr2.frx":2748
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "frmEdiEntr2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbxBuscar_Click()
    If cmbxBuscar.ListIndex = 1 Then
        cmndbtnTerc.Visible = True
    Else
        cmndbtnTerc.Visible = False
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBuscar_Click()
    Dim strSQL As String
    Dim sOrd As String
    If cmbxOrder.ListIndex = 0 Then
        sOrd = "Num_Not"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrd = "Nome_For"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrd = "Emissao_Not"
    Else
        sOrd = "Entrada_Not"
    End If
    strSQL = "select tc.*,tf.nome_for,tf.uf_for from tab_fornecedores as tf, tab_compras as tc "
    strSQL = strSQL & " where tc.cod_for=tf.cod_for "
    If cmndbtnTerc.Visible = True Then
        strSQL = strSQL & " and tf.nome_for like '" & txtbxExpr.Text & "%' "
    Else
        strSQL = strSQL & " and tc.num_not like '" & txtbxExpr.Text & "%' "
    End If
    If txtbxdataFin.Text <> "" Then
        strSQL = strSQL & " and tc.entrada_not < " & F_PesqData(txtbxdataFin.Text)
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tc.entrada_not > " & F_PesqData(txtbxDataIni.Text)
    End If
    If cmbxTipo.ListIndex > -1 Then
        strSQL = strSQL & " and tc.status_not='" & Left(cmbxTipo.Text, 1) & "' "
    End If
    strSQL = strSQL & " order by " & sOrd
    dtNot.RecordSource = strSQL
    dtNot.Refresh
End Sub

Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    With dtNot.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "1" Or !status_not = "2" Then
            If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !num_not & _
                " ?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "delete from tab_compras where seq_not = '" & !seq_Not & "'"
            
                banco.Execute strSQL
                strSQL = "delete from tab_itcompras where seq_not = '" & !seq_Not & "'"
            
                banco.Execute strSQL
            
                dtNot.Refresh
            End If
        ElseIf !status_not = "4" Then
            If strExcldoc_usu <> "S" Then
                MsgBox "Você não tem permissão para cancelar esta nota!", vbExclamation
                Exit Sub
                                
            End If
            If MsgBox("Deseja realmente cancelar a nota fiscal e voltar todo o estoque dos produtos envolvidos?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
                End If
                P_cancelanotaFechada !seq_Not
                cmndbtnBuscar_Click
        End If
        
        
    End With
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEntradas
    fr.operacao = "I"
    If dtNot.Recordset.RecordCount <= 0 Or dbgrd.Row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With dtNot.Recordset
        If !status_not <> "1" Then
            MsgBox "A Nota Não está Aberta!", vbExclamation
            Exit Sub
        End If
        
        fr.Tag = !seq_Not
        fr.dbgrdProd.Tag = fr.Tag
        fr.caption = "Nota Fiscal: " & !num_not & " de " & !nome_for
        fr.txtbxDesconto.Text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.Text = !num_not
        fr.txtbxNum_Not.Tag = !num_not
        fr.txtbxSerie_Not.Text = !serie_not
        fr.txtbxSerie_Not.Tag = !serie_not
        fr.txtbxBC_Not.Text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.Text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.Text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.Text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.Text = Format(!emissao_not)
        fr.txtbxEntrada_Not.Text = Format(!entrada_not)
        fr.txtbxICMS_Not.Text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.Text = Format(!ipi_Not, "###,###,##0.00")
        fr.txtbxNome_ter.Text = !nome_for
        fr.txtbxNome_ter.Tag = !cod_for
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.Text = FormatVal(!valor_not)
        
        P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        
        fr.Show 0, mdiFrm
    End With
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmEntradas
    fr.operacao = "I"
    fr.Show 0, mdiFrm
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    dtNot.Recordset.Close
    Unload Me
End Sub

Private Sub cmndbtnTerc_Click()
     
    frmSelectCli.txtbxNome.Tag = "F"
    
    frmSelectCli.txtbxNome.Text = txtbxExpr.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxExpr.Text = frmSelectCli.dtBusca.Recordset!nome
    txtbxExpr.Tag = frmSelectCli.dtBusca.Recordset!cod
    Unload frmSelectCli
End Sub

Private Sub dbgrd_DblClick()
    cmndbtnEdit_Click
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

Private Sub dtNot_Reposition()
    With dtNot.Recordset
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "1" Then
            cmndbtnEdit.Enabled = True
            cmndbtnCancelar.Enabled = True
        ElseIf !status_not = "2" Then
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = True
        ElseIf !status_not = "4" Then
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = True
        
        Else
            
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = False
        
        End If
        
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        
        frmNotaEntrada.Show 0
        
    End If
End Sub

Private Sub Form_Load()
   ' Me.Width = 10560
   ' Me.Height = 6585
    P_CentraForm Me
    dtNot.Connect = banco.Connect
    cmbxBuscar.ListIndex = 0
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmbxBuscar_Click
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Resize()
    P_CentraObejct frControle, Me
End Sub

Private Sub txtbxDataFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataFin_LostFocus()
    P_AjustaData txtbxdataFin
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And cmndbtnTerc.Visible = True Then
        cmndbtnTerc_Click
    End If
    cmndbtnBuscar_Click
End Sub
