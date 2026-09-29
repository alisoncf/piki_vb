VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEditaMov 
   Caption         =   "Movimentações"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   Icon            =   "frmEditaMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   12360
   Begin VB.PictureBox pcLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   8445
      Left            =   0
      ScaleHeight     =   8415
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEditaMov.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CheckBox chkAberto 
         Caption         =   "Apenas Parc. em Aberto"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":0403
         Left            =   120
         List            =   "frmEditaMov.frx":0413
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEditaMov.frx":043F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5640
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   695
         Left            =   120
         Picture         =   "frmEditaMov.frx":0589
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEditaMov.frx":0913
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4920
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   695
         Left            =   120
         Picture         =   "frmEditaMov.frx":0A5D
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6360
         Width           =   1935
      End
      Begin VB.CommandButton cmndbtnTerc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         Picture         =   "frmEditaMov.frx":0BC8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cmbxBuscar 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":0CC1
         Left            =   120
         List            =   "frmEditaMov.frx":0CCE
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmbxTipo 
         Height          =   315
         ItemData        =   "frmEditaMov.frx":0CEE
         Left            =   120
         List            =   "frmEditaMov.frx":0CF8
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expressão:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Por:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   12330
      TabIndex        =   21
      Top             =   8445
      Width           =   12360
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Não há cheque!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         TabIndex        =   24
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   22
         Top             =   -60
         Width           =   480
      End
   End
   Begin VB.Data dtChe 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data dtNot 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEditaMov.frx":0D2C
      Height          =   6975
      Left            =   2280
      OleObjectBlob   =   "frmEditaMov.frx":0D40
      TabIndex        =   20
      Top             =   120
      Width           =   9975
   End
   Begin MSDBGrid.DBGrid dbgrdChe 
      Bindings        =   "frmEditaMov.frx":2B69
      Height          =   1335
      Left            =   2280
      OleObjectBlob   =   "frmEditaMov.frx":2B7D
      TabIndex        =   19
      Top             =   7080
      Width           =   9975
   End
End
Attribute VB_Name = "frmEditaMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub chkAberto_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
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

Private Sub cmndbtnBaixar_Click()
    Dim sSEQ As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim cJuros As Currency
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para baixar documentos!", vbExclamation
        Exit Sub
    End If
    
'    If IsLoad("frmListaDoc") Then
'        If frmListaDoc.Tag <> Me.Name Then
'            MsgBox "Por favor, escolha as parcelas a serem baixadas!", vbExclamation
'            Exit Sub
'        End If
'    Else
'        MsgBox "Por favor, escolha as parcelas a serem baixadas!", vbExclamation
'        Exit Sub
'    End If
    
    With dtNot.Recordset
    
        If .RecordCount = 0 Then
            MsgBox "Não há o que baixar!", vbExclamation
            Exit Sub
        End If
            cJuros = F_CalculaJuros(!Venc_Itd, Date, !reman_itd)
            
            frmBaixaPar.Tag = dtNot.Recordset!seq_itd
            frmBaixaPar.lblNome_Ter.caption = !nome
            frmBaixaPar.txtbxJuros.Text = FormatVal(cJuros)
            frmBaixaPar.txtbxReman.Text = FormatVal(!reman_itd)
            frmBaixaPar.txtbxValor_Par.Text = FormatVal(!Valor_Itd)
            frmBaixaPar.lblParc.caption = "Parcela: " & !Parc_Itd & " de: " & !Parc_Doc
            frmBaixaPar.cmndbtnConfirma.Tag = Left(cmbxTipo.Text, 1)
            frmBaixaPar.Show 1
            
            If frmBaixaPar.Tag = "" Then
                Unload frmBaixaPar
                Exit Sub
            Else
                If MsgBox("Confirma baixa da Parcela?", vbYesNo + vbQuestion) = vbYes Then
                    With frmBaixaPar.msflxGrd
                        For iCount = 1 To .rows - 1
                            .Row = iCount
                            
                            
                            .Col = 2
                            sEsp = Left(.Text, 1)
                            
                            .Col = 0: c.Venc_che = .Text
                            .Col = 1: c.Valor_che = .Text
                            .Col = 3: c.num_che = .Text
                            .Col = 4: c.cod_bco = .Text
                            .Col = 5: c.age_che = .Text
                            .Col = 6: c.cco_che = .Text
                            .Col = 7: c.nome_che = .Text
                            .Col = 8: c.cpf_che = .Text
                            If sEsp = "1" Then
                                sSEQ = F_GravaFinanceiro(Left(cmbxTipo.Text, 1), Date, CStr(c.Valor_che), "Baixa doc " & _
                                    dtNot.Recordset!num_doc & "-" & dtNot.Recordset!nome, sEsp, frmBaixaPar.cmbxConta.Tag, _
                                    dtNot.Recordset!seq_itd, "", "", IIf(cmbxTipo.ListIndex = 0, strPlaRecebDupl, strPlaPagDupl))
                                
                                If cmbxTipo.ListIndex = 0 Then
                                    If F_GravaClass(sChe, strPlaRecebDupl, CStr(c.Valor_che)) Then
                                    End If
                                Else
                                    If F_GravaClass(sChe, strPlaPagDupl, CStr(c.Valor_che)) Then
                                    End If
                                End If
                                
                            Else
                            
                                sSEQ = F_GravaDoc("BAIXA DOC:" & dtNot.Recordset!num_doc & "-" & dtNot.Recordset!nome, _
                                     dtNot.Recordset!num_doc, CStr(c.Valor_che), Date, _
                                    dtNot.Recordset!Codigo, Left(cmbxTipo.Text, 1), "", "1", IIf(Left(cmbxTipo.Text, 1) = "R", strPlaRecebDupl, strPlaPagDupl))
                                If Left(cmbxTipo.Text, 1) = "R" Then
                                    sChe = F_GravaCh(sSEQ, c.num_che, c.cod_bco, c.age_che, c.cco_che, _
                                            c.nome_che, CCur(c.Valor_che), c.Venc_che, c.cod_bco, Left(cmbxTipo.Text, 1), _
                                            "", "", dtNot.Recordset!Parc_Itd, "", c.cpf_che, Time)
                                Else
                                    sChe = F_GravaChePag(sSEQ, c.num_che, c.cod_bco, _
                                            CCur(c.Valor_che), c.Venc_che, dtNot.Recordset!Codigo, dtNot.Recordset!Parc_Itd, Time)
                                End If
                            End If
                            cValor = cValor + c.Valor_che
                        Next
                        strSQL = "update tab_itdoc1 set reman_itd = reman_itd-" & F_GravaDec(cValor - CCur(frmBaixaPar.txtbxJuros.Text) + CCur(frmBaixaPar.txtbxDesconto.Text))
                        strSQL = strSQL & ",juros_itd=" & F_GravaDec(frmBaixaPar.txtbxJuros.Text)
                        strSQL = strSQL & ",Desconto_itd=" & F_GravaDec(frmBaixaPar.txtbxDesconto.Text)
                        strSQL = strSQL & ", baixa_itd=" & F_GravaData(Date)
                        strSQL = strSQL & " where seq_itd='" & dtNot.Recordset!seq_itd & "' "
                        banco.Execute strSQL
                        Unload frmBaixaPar
                        MsgBox "Documento Baixado!", vbInformation
                        dtNot.Refresh
                    End With
                End If
            End If
        
    End With
End Sub

Private Sub cmndbtnBuscar_Click()
        
    Dim strSQL As String
    Dim sOrderBy As String
    Dim sCampo As String
    On Error GoTo trata_Erro
    
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "venc_Itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "num_doc,venc_itd"
    Else
        sOrderBy = "reman_itd"
    End If
        
    
    If cmbxTipo.ListIndex = 0 Then
        strSQL = "select historico_doc,td.tipo_doc,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, tc.cod_cli as codigo, tc.nome_cli as nome,td.historico_doc,td.valor_Doc,td.parc_doc,td.num_doc,td.data_doc, td.valor_doc,tl.nome_pla,"
        strSQL = strSQL & " ((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, ((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, "
        strSQL = strSQL & "ti.* from tab_clientes as tc, tab_doc1 as td,tab_planocontas as tl, tab_itdoc1 as ti,tab_parametros as tp  "
        strSQL = strSQL & " where tc.cod_cli=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='R' "
        If chkAberto.Value = 1 Then
            strSQL = strSQL & " and ti.reman_itd > 0 "
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
        
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_Cli =  '" & txtbxExpr.Tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            sCampo = "ti.valor_itd = " & str(txtbxExpr.Text)
        
        End If
        strSQL = strSQL & " and tl.cod_pla=td.cod_pla "
        strSQL = strSQL & " order by " & sOrderBy
            
        
    Else
        strSQL = "select  historico_doc,td.tipo_doc,tc.cod_for as codigo,tp.nome_pla,tc.nome_for as nome,td.valor_Doc,td.parc_doc,td.num_doc,td.data_doc, td.valor_doc,"
        strSQL = strSQL & "ti.* from tab_fornecedores as tc, tab_doc1 as td,tab_itdoc1 as ti ,tab_planocontas as tp  "
        strSQL = strSQL & " where tc.cod_for=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & " and td.tipo_doc='P' "
        strSQL = strSQL & " and tp.cod_pla=td.cod_pla "
        If chkAberto.Value = 1 Then
            strSQL = strSQL & " and ti.reman_itd > 0 "
        End If
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and tc.cod_for =  '" & txtbxExpr.Tag & "' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            sCampo = "ti.valor_itd = " & str(txtbxExpr.Text)
        
        End If
        strSQL = strSQL & " order by " & sOrderBy
    End If
    
    
    
    
    
    
    
    dtNot.RecordSource = strSQL
    dtNot.Refresh
    
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnLocaliza_Click()
    
End Sub

Private Sub cmndbtnImpr_Click()

End Sub

Private Sub cmndbtnEdit_Click()
        
    Dim fr As New frmEdiLan
    'On Error Resume Next
    If cmndbtnEdit.Enabled = False Then Exit Sub
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    
    With dtNot.Recordset
    
        If .RecordCount < 0 Then
            Exit Sub
        End If
        fr.Tag = !seq_itd
        fr.lblData.caption = !Data_Doc
        fr.lblNUM_Doc.caption = !num_doc
        fr.lblParc.caption = !Parc_Itd & "/" & !Parc_Doc
        fr.lblSeq_Doc.caption = !seq_doc
        fr.lblSeq_Itd.caption = !seq_itd
        fr.lblValor.caption = FormatVal(!Valor_Itd)
        fr.txtbxHist.Text = !historico_doc
        fr.txtbxNome.Text = !nome
        fr.txtbxNome.Tag = !Codigo
        fr.txtbxVenc.Text = !Venc_Itd
        If !tipo_doc = "R" Then
            P_PreenchePlanoContas fr.cmbxPlanoConta, True
        Else
            P_PreenchePlanoContas fr.cmbxPlanoConta, False
        End If
        fr.cmbxPlanoConta.Text = !nome_pla
        
        fr.txtbxValor.Text = FormatVal(!reman_itd)
        fr.Show 1
        dtNot.Refresh
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnNew_Click()
    If strLandoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbExclamation
        Exit Sub
    End If
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    On Error Resume Next
    
    dtChe.Recordset.Close
    dtNot.Recordset.Close
    Unload Me
    
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    
    If cmbxBuscar.ListIndex <> 1 Then
        cmndbtnBuscar_Click
        Exit Sub
    End If
    If cmbxTipo.ListIndex = 0 Then
        frmSelectCli.txtbxNome.Tag = "C"
    Else
        frmSelectCli.txtbxNome.Tag = "F"
    End If
    frmSelectCli.txtbxNome.Text = txtbxExpr.Text
    
    On Error Resume Next
    
    frmSelectCli.Show 1
    txtbxExpr.Text = frmSelectCli.dtBusca.Recordset!nome
    txtbxExpr.Tag = frmSelectCli.dtBusca.Recordset!cod
    Unload frmSelectCli
    cmndbtnBuscar_Click
End Sub


Private Sub dbgrd_DblClick()
    With dtNot.Recordset
        cmndbtnEdit_Click
        If !reman_itd <= 0 Then
            Exit Sub
        End If
        If IsLoad("frmListaDoc") Then
            If frmListaDoc.Tag <> Me.Name Then
                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
                Exit Sub
            End If
            If F_InFlxgrd(frmListaDoc.msflxGrd, !num_doc, 0) = -1 Then
                frmListaDoc.msflxGrd.AddItem !num_doc & Chr(9) & !reman_itd
            End If
            frmListaDoc.soma
        Else
            frmListaDoc.Tag = Me.Name
            frmListaDoc.caption = "Documentos de Financeiras"
            frmListaDoc.msflxGrd.Cols = 2
            frmListaDoc.msflxGrd.rows = 0
            frmListaDoc.msflxGrd.AddItem !num_doc & Chr(9) & !reman_itd
            frmListaDoc.Show 0, mdiFrm
            frmListaDoc.soma
        End If
    
    End With
End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_LostFocus()
    dbgrd.MarqueeStyle = 0
End Sub

Private Sub dbgrdChe_GotFocus()
    dbgrdChe.MarqueeStyle = 3
End Sub

Private Sub dbgrdChe_LostFocus()
    dbgrdChe.MarqueeStyle = 0
End Sub

Private Sub dtNot_Reposition()
    Dim strSQL As String
    On Error Resume Next
    If bActive = False Then Exit Sub
    With dtNot.Recordset
        If .RecordCount = 0 Then Exit Sub
        strSQL = "select * from tab_cheques1 where seq_doc='" & !seq_doc & "' "
        dtChe.RecordSource = strSQL
        dtChe.Refresh
        
        If dtChe.Recordset.RecordCount > 0 Then
            dbgrdChe.caption = "Cheques referentes ao documento " & !num_doc
            dbgrdChe.Visible = True
        Else
            dbgrdChe.Visible = False
        End If
        If .RecordCount > 0 Then
            If !reman_itd = 0 Then
                cmndbtnEdit.Enabled = False
                cmndbtnBaixar.Enabled = False
            Else
                cmndbtnEdit.Enabled = True
                cmndbtnBaixar.Enabled = True
            End If
            Exit Sub
        
        End If
        
    End With
End Sub

Private Sub Form_Activate()
    MudaAparencia Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Sub soma()
    Dim cTotal As Currency
    
    On Error GoTo trata_Erro
    bActive = False
    With dtNot.Recordset
        .MoveFirst
        Do While Not .EOF
            cTotal = cTotal + !reman_itd
            .MoveNext
        Loop
        .MoveFirst
    End With
    lblTotal.caption = FormatVal(cTotal)
    bActive = True
    Exit Sub
trata_Erro:
    lblTotal.caption = "0,00"
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    On Error Resume Next
    txtbxDataIni.Text = "01" & Mid(Format(Date, "DD/MM/YYYY"), 3, 8)
    strSQL = DateAdd("M", 1, CDate(txtbxDataIni.Text))
    txtbxdataFin.Text = DateAdd("D", -1, strSQL)
    dtNot.Connect = banco.Connect
    cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    dtNot.Connect = banco.Connect
    dtChe.Connect = banco.Connect
    
    resizeForm Me
    
    cmbxBuscar.ListIndex = 1
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmndbtnBuscar_Click
    
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
 
    
End Sub

Private Sub Form_Resize()
    dbgrd.Width = Me.Width - pcLeft.Width - 300
    dbgrdChe.Top = Me.Height - dbgrdChe.Height - 500
    dbgrd.Height = dbgrdChe.Top - 100
    dbgrdChe.Width = dbgrd.Width
End Sub

Private Sub txtbxDAtaFIn_GotFocus()
    P_SelecTexto txtbxdataFin
End Sub

Private Sub txtbxDataFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
        cmndbtnBuscar_Click
    End If
    
End Sub

Private Sub txtbxDAtaFIn_LostFocus()
    P_AjustaData txtbxdataFin
    cmndbtnBuscar_Click
End Sub

Private Sub txtbxDataIni_GotFocus()
    P_SelecTexto txtbxDataIni
End Sub

Private Sub txtbxDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtbxDataIni_LostFocus()
    P_AjustaData txtbxDataIni
End Sub

Private Sub txtbxExpr_GotFocus()
    P_SelecTexto txtbxExpr
End Sub

Private Sub txtbxExpr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnBuscar_Click
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxExpr_LostFocus()
    If cmbxBuscar.ListIndex = 0 Then
        txtbxExpr.Text = F_SoNumero(txtbxExpr.Text)
    Else
        P_AjustaTexto txtbxExpr
    End If
End Sub
