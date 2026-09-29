VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEdiEntr 
   Caption         =   "Edição de Compras"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   Icon            =   "frmEdiEntr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   14040
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00B4DFE0&
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   2265
      TabIndex        =   14
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton cmndbtnImprGrade 
         Caption         =   "&Imprimir Grade"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6000
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiEntr.frx":0557
         Left            =   120
         List            =   "frmEdiEntr.frx":0567
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   10
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cmbxBuscar 
         Height          =   315
         ItemData        =   "frmEdiEntr.frx":058F
         Left            =   120
         List            =   "frmEdiEntr.frx":0599
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmEdiEntr.frx":05B1
         Left            =   120
         List            =   "frmEdiEntr.frx":05C4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":077E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0B08
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0C52
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6840
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   685
         Left            =   120
         Picture         =   "frmEdiEntr.frx":0D9C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5280
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         Picture         =   "frmEdiEntr.frx":11A9
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   285
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
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEdiEntr.frx":1533
      Height          =   6015
      Left            =   2400
      OleObjectBlob   =   "frmEdiEntr.frx":1547
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmEdiEntr"
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
    strSQL = "select tc.*,tf.cnpj_for,tf.nome_for,tf.uf_for from tab_fornecedores as tf, tab_compras as tc "
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
        If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !num_not & _
            " ?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "delete from tab_compras where seq_not = '" & !seq_Not & "'"
            
            banco.Execute strSQL
            strSQL = "delete from tab_itcompras where seq_not = '" & !seq_Not & "'"
            
            banco.Execute strSQL
            
            dtNot.Refresh
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

Private Sub cmndbtnImprGrade_Click()
    P_DbgrdToHtml "Entradas de " & txtbxDataIni.Text & " até " & txtbxdataFin.Text, dbgrd, dtNot, 1
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
        On Error Resume Next
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "1" Then
            cmndbtnEdit.Enabled = True
            cmndbtnCancelar.Enabled = True
        ElseIf !status_not = "2" Then
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
    resizeForm Me
    dtNot.Connect = banco.Connect
    cmbxBuscar.ListIndex = 0
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmbxBuscar_Click
    cmndbtnBuscar_Click
End Sub

Private Sub Form_Resize()
    dbgrd.Width = Me.Width - dbgrd.Left - 100
    dbgrd.Height = Me.Height - dbgrd.Top - 100
'    vsc.Height = pcLeft.Height - 100
'    vsc.max = pcLeft.Height
    
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
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

Private Sub vsc_Change()
    'pcLeft.Top = 0 - vsc.Value
End Sub
