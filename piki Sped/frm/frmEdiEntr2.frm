VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiEntr 
   Caption         =   "Edição de Compras"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmEdiEntr2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   15240
   Begin VB.PictureBox pcBuscar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   15210
      TabIndex        =   6
      Top             =   6795
      Width           =   15240
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   450
         Left            =   11760
         Picture         =   "frmEdiEntr2.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEdiEntr2.frx":04F5
         Left            =   10440
         List            =   "frmEdiEntr2.frx":0505
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmEdiEntr2.frx":053C
         Left            =   3480
         List            =   "frmEdiEntr2.frx":0552
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtProduto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton cmbxProd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         Picture         =   "frmEdiEntr2.frx":05C4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdForn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         Picture         =   "frmEdiEntr2.frx":096F
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         Height          =   195
         Left            =   10440
         TabIndex        =   22
         Top             =   0
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor:"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&De:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8280
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "&Até:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9360
         TabIndex        =   17
         Top             =   0
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação:"
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Top             =   0
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto:"
         Height          =   195
         Left            =   4920
         TabIndex        =   15
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   15210
      TabIndex        =   1
      Top             =   7530
      Width           =   15240
      Begin VB.CommandButton cmndbtnReabrir 
         Caption         =   "&Reabrir"
         Height          =   690
         Left            =   7320
         Picture         =   "frmEdiEntr2.frx":0D1A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEfetivar 
         Caption         =   "&Efetivar"
         Height          =   690
         Left            =   5520
         Picture         =   "frmEdiEntr2.frx":0E64
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnTransf 
         Caption         =   "&Transferir"
         Enabled         =   0   'False
         Height          =   690
         Left            =   9120
         Picture         =   "frmEdiEntr2.frx":0FAE
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   690
         Left            =   10920
         Picture         =   "frmEdiEntr2.frx":10F8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCancelar 
         Caption         =   "&Cancelar"
         Height          =   690
         Left            =   3720
         Picture         =   "frmEdiEntr2.frx":1263
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   690
         Left            =   1920
         Picture         =   "frmEdiEntr2.frx":13AD
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   690
         Left            =   120
         Picture         =   "frmEdiEntr2.frx":14F7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid dtgrd 
      Bindings        =   "frmEdiEntr2.frx":1881
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   11880
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
      Caption         =   "Entradas Encontradas"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "num_not"
         Caption         =   "Número"
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
         DataField       =   "nome_for"
         Caption         =   "Fornecedor"
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
         DataField       =   "emissao_not"
         Caption         =   "Emissão"
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
         DataField       =   "data_not"
         Caption         =   "Data Entrada"
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
         DataField       =   "valor_not"
         Caption         =   "Valor"
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
      BeginProperty Column05 
         DataField       =   "icms_not"
         Caption         =   "ICMS"
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
      BeginProperty Column06 
         DataField       =   "status_not"
         Caption         =   "Situação"
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
      BeginProperty Column07 
         DataField       =   "seq_not"
         Caption         =   "ID"
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
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2789,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdiEntr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim operacao As String
Dim tbComp As String
Dim tbItens As String
Dim bActive As Boolean
Dim WithEvents adc As ADODB.Recordset
Attribute adc.VB_VarHelpID = -1
Public Sub setOperacao(op As String)
    operacao = op
    If operacao = "F" Then
        tbComp = "tab_compras"
        tbItens = "tab_itcompras"
        pcBuscar.BackColor = vbGreen
    Else
        tbComp = "tab_cmp"
        tbItens = "tab_itcmp"
        pcBuscar.BackColor = vbYellow
    End If
    
End Sub

Private Sub cmbxBuscar_Click()
    
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub adc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
        On Error Resume Next
    With adc
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "1" Then
            cmndbtnEdit.Enabled = True
            cmndbtnCancelar.Enabled = True
        ElseIf !status_not = "2" And operacao = "F" Then
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = True
            cmndbtnTransf.Enabled = True
        ElseIf !status_not = "4" Then
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = True
        
        Else
            
            cmndbtnEdit.Enabled = False
            cmndbtnCancelar.Enabled = False
        
        End If
        
    End With

End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub cmbxProd_Click()
    buscaProd
End Sub
Sub buscaProd()
    Dim p As Produto
    p = Produtos.F_BuscaProdSaida(txtProduto.Text)
    If p.cod_pro <> "" Then
        txtProduto.Text = p.nome_pro
        txtProduto.tag = p.cod_pro
        cmndbtnBuscar_Click
    End If
End Sub
Private Sub cmbxTipo_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TrocaTabKey KeyCode
    End If
End Sub
Sub busca()
    Dim strSQL As String
    Dim sOrd As String
    
    
    If cmbxOrder.ListIndex = 1 Then
    Else
        sOrd = "data_not desc,nome_for,num_not"
    End If
    strSQL = "select tc.*,tf.nome_for,tf.uf_for from " & TAB_FORNECEDORES & " as tf inner join " & tbComp & " as tc "
    strSQL = strSQL & " on tf.cod_for=tc.cod_for "
    strSQL = strSQL & " left join " & tbItens & " as ti"
    strSQL = strSQL & " on ti.seq_not=tc.cod_for "
    strSQL = strSQL & " left join " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " on ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " where tf.cod_for=tc.cod_for "
    If txtNumero.Text <> "" Then
        strSQL = strSQL & " and tc.num_not like '%" & txtNumero.Text & "' "
    End If
    If txtForn.tag <> "" Then
        strSQL = strSQL & " and tf.cod_for ='" & txtForn.tag & "' "
    End If
    If txtNumero.tag <> "" Then
        strSQL = strSQL & " and tc.num_not like '" & txtNumero.Text & "%' "
    End If
    If txtbxDAtaFIn.Text <> "" Then
        strSQL = strSQL & " and tc.data_not <= " & F_GravaDataTempo(txtbxDAtaFIn.Text & " 23:59:00")
    End If
    If txtbxDataIni.Text <> "" Then
        strSQL = strSQL & " and tc.entrada_not >= " & F_GravaData(txtbxDataIni.Text & "00:00:01")
    End If
    strSQL = strSQL & " and nome_for like '" & txtForn.Text & "%'"
    If cmbxTIpo.Text <> "" Then
        strSQL = strSQL & " and tc.status_not='" & Left(cmbxTIpo.Text, 1) & "' "
    End If
    strSQL = strSQL & " order by " & sOrd
    abreObjetoAdo adc, strSQL, dtgrd
    '
    
    'MsgBox adc.RecordCount & " registros econtrados"
    
    
End Sub
Private Sub cmndbtnBuscar_Click()
    If bActive = False Then Exit Sub
    busca
End Sub

Private Sub cmdBuscar_Click()
    busca
End Sub

Private Sub cmdForn_Click()
    buscaFor
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
    f = fornecedor.F_BuscaFornecedor(txtForn.Text)
    If f.cod_for <> "" Then
        txtForn.Text = f.nome_for
        txtForn.tag = f.cod_for
        cmbxTIpo.SetFocus
    End If
End Sub
Private Sub cmndbtnCancelar_Click()
    Dim strSQL As String
    With adc
        If .RecordCount = 0 Then
            Exit Sub
        End If
        If !status_not = "1" Or !status_not = "2" Then
            If MsgBox("Deseja Realmente Excluir a Nota Fiscal " & !num_not & _
                " ?", vbYesNo + vbCritical) = vbYes Then
                strSQL = "delete from " & tbComp & " where seq_not = '" & !seq_not & "'"
            
                conn.Execute strSQL
                strSQL = "delete from " & tbItens & " where seq_not = '" & !seq_not & "'"
            
                conn.Execute strSQL
                
                adc.MoveFirst
            End If
        ElseIf !status_not = "4" Then
            If strExcldoc_usu <> "S" Then
                MsgBox "Você não tem permissão para cancelar esta nota!", vbExclamation
                Exit Sub
                                
            End If
            If MsgBox("Deseja realmente cancelar a nota fiscal e voltar todo o estoque dos produtos envolvidos?", vbYesNo + vbCritical) = vbNo Then
                    Exit Sub
                End If
                P_cancelanotaFechada !seq_not, operacao
                cmndbtnBuscar_Click
        End If
        
        
    End With
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEntradaSped
    fr.operacao = operacao
    If adc.RecordCount <= 0 Or dtgrd.Row = -1 Then
        MsgBox "Não há notas fiscais nesta pesquisa!", vbExclamation
        Exit Sub
    End If
    
    With adc
        If !status_not <> "1" Then
            MsgBox "A Nota Não está Aberta!", vbExclamation
            Exit Sub
        End If
        
        fr.tag = !seq_not
        fr.dtgrd.tag = fr.tag
        fr.caption = "Nota Fiscal: " & !num_not & " de " & !nome_for
        fr.txtbxDesconto.Text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.Text = !num_not
        fr.txtbxNum_Not.tag = !num_not
        fr.txtbxSerie_Not.Text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.Text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.Text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.Text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.Text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.Text = Format(!emissao_not)
        fr.txtbxEntrada_Not.Text = Format(!entrada_not)
        fr.txtbxICMS_Not.Text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.Text = Format(!ipi_Not, "###,###,##0.00")
        fr.txtForn.Text = !nome_for
        fr.txtForn.tag = !cod_for
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.Text = FormatVal(!valor_not)
        
        'P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        'P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        
        fr.Show 0, mdiFrm
    End With
End Sub

Private Sub cmndbtnEfetivar_Click()
    If strEfetnota_usu <> "S" Then
        MsgBox "Operação não permitida para esse usuário!", vbExclamation
        Exit Sub
    End If
    With adc
        If !status_not <> "2" Then
            MsgBox "Esta nota não pode ser efetivada!", vbExclamation
            Exit Sub
        End If
        If operacao = "E" Then
            If EfetivaNotaR(!seq_not, True) Then
                MsgBox "Nota fiscal efetivada com sucesso!", vbInformation
            End If
        Else
            If EfetivaNotaF(!seq_not, True) Then
                MsgBox "Nota fiscal efetivada com sucesso!", vbInformation
            End If
        End If
    
    adc.Requery
    
    End With
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmEntradaSped
    fr.operacao = operacao
    fr.Show 0, mdiFrm
End Sub

Private Sub cmndbtnReabrir_Click()
    On Error GoTo trata_Erro
    If adc!status_not <> "4" Then
        If MsgBox("Deseja realmente abrir a nota?", vbYesNo + vbCritical) = vbYes Then
            strSQL = "update " & tbComp & " set status_not='1'"
            strSQL = strSQL & " where seq_not='" & adc!seq_not & " ' "
            conn.Execute strSQL
            adc.MoveFirst
        End If
    Else
        MsgBox "Esta nota não pode ser transferida!", vbExclamation
    End If
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    adc.Close
    Unload Me
End Sub



Private Sub cmndbtnTransf_Click()
'    On Error GoTo trata_Erro
    If adc!status_not = "2" Then
    
        transfere
        adc.Requery
    Else
        MsgBox "Esta nota não pode ser transferida!", vbExclamation
    End If
    
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub transfere()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sSeq As String
    
    With adc
        Dim fr As New frmEntradaSped
        fr.operacao = "E"
        fr.tag = !seq_not
        fr.dtgrd.tag = fr.tag
        fr.caption = "Nota Fiscal: " & !num_not & " de " & !nome_for
        fr.txtbxDesconto.Text = FormatVal(!desconto_not)
        fr.txtbxNum_Not.Text = !num_not
        fr.txtbxNum_Not.tag = !num_not
        fr.txtbxSerie_Not.Text = !serie_not
        fr.txtbxSerie_Not.tag = !serie_not
        fr.txtbxBC_Not.Text = Format(!bc_not, "##,###,##0.00")
        fr.txtbxValorProd.Text = FormatVal(!valorprod_not)
        fr.txtbxDepesas_Not.Text = Format(!despesas_not, "###,###,##0.00")
        fr.txtbxFrete.Text = Format(!frete_not, "###,###,##0.00")
        fr.txtbxEmissao_Not.Text = !emissao_not
        fr.txtbxEntrada_Not.Text = !entrada_not
        fr.txtbxICMS_Not.Text = Format(!icms_not, "###,###,##0.00")
        fr.txtbxIPI_Not.Text = Format(!ipi_Not, "###,###,##0.00")
        fr.txtForn.Text = !nome_for
        fr.txtForn.tag = !cod_for
        fr.lblUF_Ter.caption = Format(!uf_for)
        
        fr.txtbxValor_Not.Text = FormatVal(!valor_not)
        
        'P_PreencheOperacoes fr.cmbxTipo_Op, "E"
        
         
        'P_LocalizaOperacao !operacao_not, fr.cmbxTipo_Op
        
        sSeq = MaxKeySQL("tab_cmp", "seq_not", "", "000000#")
        strSQL = "insert into tab_cmp(seq_not,Cod_For,Num_not,Serie_not,desconto_not,valorprod_not,Valor_not,Bc_Not,"
        strSQL = strSQL & "frete_not,icms_not,ipi_not,despesas_not,operacao_not,emissao_not,"
        strSQL = strSQL & "entrada_not,cod_usu,status_not,seq_ent)Values("
        strSQL = strSQL & "'" & sSeq & "',"
        strSQL = strSQL & "'" & !cod_for & "', "
        strSQL = strSQL & "'" & !num_not & "',"
        strSQL = strSQL & "'" & !serie_not & "',"
        strSQL = strSQL & F_GravaDec(!desconto_not) & ","
        strSQL = strSQL & F_GravaDec(!valorprod_not) & ","
        strSQL = strSQL & F_GravaDec(!valor_not) & ","
        strSQL = strSQL & F_GravaDec(!bc_not) & ","
        strSQL = strSQL & F_GravaDec(!frete_not) & ","
        strSQL = strSQL & F_GravaDec(!icms_not) & ","
        strSQL = strSQL & F_GravaDec(!ipi_Not) & ","
        strSQL = strSQL & F_GravaDec(!despesas_not) & ","
        strSQL = strSQL & "'" & (!operacao_not) & "',"
        strSQL = strSQL & F_GravaData(!emissao_not) & ","
        strSQL = strSQL & F_GravaData(!entrada_not) & ","
        strSQL = strSQL & "'" & strCod_Usu & "','1', " 'em aberto
        strSQL = strSQL & "'" & !seq_not & "') "
        conn.Execute strSQL

        
        strSQL = "select * from tab_itcompras where seq_not='" & !seq_not & "' "
        strSQL = strSQL & " order by seq_ite"
        rPesq.Open strSQL, conn_client
        With rPesq
            
            Do While Not .EOF
                
                strSQL = "Insert into tab_itcmp(seq_not,seq_ite,cod_pro,unit_ite,qtd_ite,qtd_emb,ipi_ite,"
                strSQL = strSQL & "gastos_ite,custo_ite,sujerido_ite,precofuturo_ite,lucrosuj_ite,"
                strSQL = strSQL & "lucroliq_ite,custosai_ite,icms_ite)values("
                strSQL = strSQL & "'" & sSeq & "',"
                strSQL = strSQL & !seq_ite & ","
                strSQL = strSQL & "'" & !cod_pro & "',"
                strSQL = strSQL & F_GravaDec(!unit_ite) & ","
                strSQL = strSQL & F_GravaDec(!qtd_ite) & ","
                strSQL = strSQL & F_GravaDec(!qtd_emb) & ","
                strSQL = strSQL & F_GravaDec(!ipi_ite) & ","
                strSQL = strSQL & F_GravaDec(!gastos_ite) & ","
                strSQL = strSQL & F_GravaDec(!custo_ite) & ","
                strSQL = strSQL & F_GravaDec(!sujerido_ite) & ","
                strSQL = strSQL & F_GravaDec(!precofuturo_ite) & ","
                strSQL = strSQL & F_GravaDec(!lucrosuj_ite) & ","
                strSQL = strSQL & F_GravaDec(!lucroliq_ite) & ","
                strSQL = strSQL & F_GravaDec(!custosai_ite) & ","
                strSQL = strSQL & F_GravaDec(!icms_ite) & ")"
                conn.Execute strSQL
                
                .MoveNext
            Loop
            
        End With
        
    End With
    
     With adc
        
        strSQL = "update tab_compras set status_not = '3' where seq_not='" & !seq_not & "' "
        
        conn.Execute strSQL
        
        fr.caption = "Finalização de NF: " & !num_not & " - " & !nome_for
        fr.tag = sSeq
        fr.Show 0
    End With
    
End Sub

Private Sub dtgrd_DblClick()
    cmndbtnEdit_Click
End Sub

Private Sub dtgrd_GotFocus()
    dtgrd.MarqueeStyle = 3
End Sub

Private Sub dtgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtgrd_DblClick
    End If
End Sub

Private Sub dtgrd_LostFocus()
    dtgrd.MarqueeStyle = 0
End Sub

Private Sub adc_Reposition()
    
End Sub

Private Sub Form_Activate()
    If operacao = "" Then
        setOperacao "F"
    End If
    txtbxDAtaFIn.Text = F_UltDiaMes
    txtbxDataIni.Text = F_PrimDiaMes
    
    
    bActive = True
    busca
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Or KeyCode = vbKeyF9 Then
        If operacao = "F" Then
            setOperacao "E"
        Else
            setOperacao "F"
        End If
        busca
        
    End If
End Sub

Private Sub Form_Load()
   ' Me.Width = 10560
   ' Me.Height = 6585
    resizeForm Me
    Set adc = New ADODB.Recordset
    
    
    cmbxTIpo.ListIndex = 0
    cmbxBuscar_Click
    
End Sub

Private Sub Form_Resize()
    dtgrd.Width = Me.Width - 200
    dtgrd.Height = pcBuscar.Top - 100
End Sub


Private Sub txtbxDAtaFIn_GotFocus()
    ganhaFoco txtbxDAtaFIn
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        TrocaTabKey KeyCode
    End If
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxDAtaFIn
End Sub

Private Sub txtbxDataIni_GotFocus()
    ganhaFoco txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey
    
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    TrocaTabKey KeyCode
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        busca
    End If
    
End Sub
