VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEditaMov2 
   Caption         =   "Movimentações"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
   Icon            =   "frmEditaMov2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   14385
   Begin VB.PictureBox pcBot 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   14325
      TabIndex        =   29
      Top             =   8415
      Width           =   14385
      Begin VB.CommandButton cmndbtnContrato 
         Caption         =   "Emitir Contrato"
         Height          =   595
         Left            =   9000
         Picture         =   "frmEditaMov2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnCarne 
         Caption         =   "Emitir &Carnê"
         Height          =   595
         Left            =   7200
         Picture         =   "frmEditaMov2.frx":0406
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnBaixar 
         Caption         =   "&Baixar"
         Height          =   595
         Left            =   3600
         Picture         =   "frmEditaMov2.frx":0502
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnExcl 
         Caption         =   "E&xcluir"
         Height          =   595
         Left            =   10800
         Picture         =   "frmEditaMov2.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnNew 
         Caption         =   "&Novo"
         Height          =   595
         Left            =   1800
         Picture         =   "frmEditaMov2.frx":0796
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnEdit 
         Caption         =   "&Editar"
         Height          =   595
         Left            =   0
         Picture         =   "frmEditaMov2.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnSair 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   595
         Left            =   12600
         Picture         =   "frmEditaMov2.frx":0C6A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmndbtnImpr 
         Caption         =   "&Imprimir"
         Height          =   595
         Left            =   5400
         Picture         =   "frmEditaMov2.frx":0DB4
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox pcTop 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   14385
      TabIndex        =   28
      Top             =   0
      Width           =   14385
      Begin VB.TextBox txtbxTipoCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtbxTerceiro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtbxContaCob 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmndbtnTerc 
         Height          =   285
         Left            =   9600
         Picture         =   "frmEditaMov2.frx":11C1
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmndbtnCob 
         Height          =   285
         Left            =   9600
         Picture         =   "frmEditaMov2.frx":154B
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmbxTipoTer 
         Height          =   315
         ItemData        =   "frmEditaMov2.frx":18D5
         Left            =   2400
         List            =   "frmEditaMov2.frx":18DF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbxOrder 
         Height          =   315
         ItemData        =   "frmEditaMov2.frx":18FB
         Left            =   9960
         List            =   "frmEditaMov2.frx":190B
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtbxExpr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   60
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Height          =   600
         Left            =   9960
         Picture         =   "frmEditaMov2.frx":1937
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtbxdataFin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtbxDataIni 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbxBuscar 
         Height          =   315
         ItemData        =   "frmEditaMov2.frx":2339
         Left            =   120
         List            =   "frmEditaMov2.frx":2349
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmndbtnTipoCob 
         Height          =   285
         Left            =   7080
         Picture         =   "frmEditaMov2.frx":2372
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terceiro:"
         Height          =   195
         Left            =   4800
         TabIndex        =   18
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lblTipoCob 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta Cobrança:"
         Height          =   195
         Left            =   7560
         TabIndex        =   7
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cobrança:"
         Height          =   195
         Left            =   4800
         TabIndex        =   4
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Terceiro:"
         Height          =   195
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar Por:"
         Height          =   195
         Left            =   9960
         TabIndex        =   16
         Tag             =   "!"
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ex&pressão:"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Tag             =   "!"
         Top             =   0
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Até:"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Tag             =   "!"
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Tag             =   "!"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   810
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
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   5640
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
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pcTotal 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   14355
      TabIndex        =   22
      Top             =   8160
      Width           =   14385
      Begin VB.CheckBox chkAberto 
         Caption         =   "Apenas Parc. em Aberto"
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Left            =   2280
         TabIndex        =   25
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
         Left            =   2760
         TabIndex        =   24
         Top             =   -60
         Width           =   480
      End
   End
   Begin MSDBGrid.DBGrid dbgrd 
      Bindings        =   "frmEditaMov2.frx":26FC
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "frmEditaMov2.frx":2710
      TabIndex        =   26
      Top             =   1320
      Width           =   13335
   End
   Begin MSDBGrid.DBGrid dbgrdChe 
      Bindings        =   "frmEditaMov2.frx":4BE5
      Height          =   735
      Left            =   0
      OleObjectBlob   =   "frmEditaMov2.frx":4BF9
      TabIndex        =   27
      Top             =   7200
      Width           =   14295
   End
End
Attribute VB_Name = "frmEditaMov2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bActive As Boolean

Private Sub chkAberto_Click()
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_Click()
    If Not bActive Then Exit Sub
    If cmbxBuscar.ListIndex = 3 Then
        dbgrd.Columns(6).caption = "Boletos"
        dbgrd.Columns(6).Visible = True
        dbgrd.Columns(6).DataField = "num_bol"
    ElseIf cmbxBuscar.ListIndex = 2 Then
        txtbxExpr.Text = "0"
    Else
        dbgrd.Columns(6).Visible = False
    End If
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxOrder_Click()
    If Not bActive Then Exit Sub
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxOrder_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmbxTipo_Click()
    If cmbxTipo.ListIndex = 0 Then
        cmbxTipoTer.ListIndex = 0
        cmbxTipo.Enabled = True
    Else
        cmbxTipo.ListIndex = -1
        cmbxTipo.Enabled = False
    End If
    If Not bActive Then Exit Sub
    
    
    cmndbtnBuscar_Click
End Sub

Private Sub cmbxTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub cmndbtnBaixar_Click()
    Dim sSeq As String
    Dim iCount As Integer
    Dim c As cheque
    Dim sChe As String
    Dim sEsp As String
    Dim cValor As Currency
    Dim cJuros As Currency
    Dim sEspBai As String
    Dim iPar  As Integer
    With dtNot.Recordset
        If .RecordCount = 0 Then
            MsgBox "Não há o que baixar!", vbExclamation
            Exit Sub
        End If
            If IsLoad("frmListaDoc") Then
                cJuros = F_CalculaJuros(!venc_itd, Date, !reman_itd)
                frmBaixaPar.tag = dtNot.Recordset!seq_itd
                frmBaixaPar.lblNome_Ter.caption = !nome
                If cmbxBuscar.ListIndex = 0 Then
                    frmBaixaPar.txtbxJuros.Text = FormatVal(cJuros)
                End If
                frmBaixaPar.txtbxReman.Text = FormatVal(!reman_itd)
                frmBaixaPar.txtbxValor_Par.Text = FormatVal(!valor_itd)
                frmBaixaPar.lblParc.caption = "Parcela: " & !parcela_itd & " de: " & !parcelas_doc
                frmBaixaPar.cmndbtnConfirma.tag = "R"
                frmBaixaPar.Show 1
            Else
                cJuros = 0
                frmBaixaPar.tag = "vários"
                frmBaixaPar.lblNome_Ter.caption = !nome
                'If cmbxBuscar.ListIndex = 0 Then
                '    frmBaixaPar.txtbxJuros.Text = FormatVal(cJuros)
                'End If
                frmBaixaPar.txtbxReman.Text = FormatVal(frmListaDoc.lblTotal.caption)
                frmBaixaPar.txtbxValor_Par.Text = FormatVal(frmListaDoc.lblTotal.caption)
                frmBaixaPar.lblParc.caption = "Parcela:  Várias"
                frmBaixaPar.cmndbtnConfirma.tag = "R"
                frmBaixaPar.Show 1
            
            End If
            If frmBaixaPar.tag = "" Then
                Unload frmBaixaPar
                Exit Sub
            Else
                If MsgBox("Confirma baixa da Parcela?", vbYesNo + vbQuestion) = vbYes Then
                    With frmBaixaPar.msflxgrd
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
                            If sEsp = "1" Then
                                
                                With dtNot.Recordset
                                    If frmBaixaPar.optnFin.Value Then
                                    sSeq = F_GravaFinanceiro("R", CStr(Date), CStr(c.Valor_che), _
                                        "Baixa docs " & !nome, _
                                         frmBaixaPar.cmbxConta.tag, "S", dtNot.Recordset!cod_pla)
                                        sEspBai = "DFI" 'dinheiro financeiro
                                    Else
                                        sSeq = F_GravaCai(Left(cmbxTipo.Text, 1), CStr(Date), CStr(c.Valor_che), _
                                        "Baixa docs " & !nome, "", _
                                            "RDI", strCod_Cai, CStr(c.Valor_che), "0", strCod_Cai, strCod_Tur, dtNot.Recordset!cod_pla)
                                            sEspBai = "DFC" 'dinheiro caixa
                                            
                                    End If
                                
                                End With
                                 'sSeq = F_GravaFinanceiro(Left(cmbxTipo.Text, 1), Date, CStr(c.Valor_che), "HISTORICO", "DIN", _
                                        frmBaixaPar.cmbxConta.Tag, dtNot.Recordset!seq_itd, "1", "")
                                    For iPar = 1 To frmListaDoc.msflxgrd.rows - 1
                                        frmListaDoc.msflxgrd.Row = iPar
                                        
                                        If F_GravaBaixaItd(dtNot.Recordset!seq_itd, CStr(c.Valor_che), "DIN", sSeq) = "" Then
                                            MsgBox "Erro na gravação da baixa!", vbExclamation
                                        End If
                                    Next
                                        
                                'P_Class c.Valor_che, Left(cmbxTipo.Text, 1), sSEQ, "F", dtNot.Recordset!historico_doc
                            ElseIf sEsp = "2" Then
                                sSeq = F_GravaDoc("BAIXA DOC:" & dtNot.Recordset!num_doc & "-" & dtNot.Recordset!nome, _
                                     dtNot.Recordset!num_doc, CStr(c.Valor_che), Date, _
                                    dtNot.Recordset!Codigo, Left(cmbxTipo.Text, 1), "", "1", "F", "C", dtNot.Recordset!cod_pla)
                                    If F_GravaBaixaItd(dtNot.Recordset!seq_itd, CStr(c.Valor_che), "DUP", sSeq) = "" Then
                                            MsgBox "Erro na gravação da baixa!", vbExclamation
                                        End If
                                    
                                If Left(cmbxTipo.Text, 1) = "R" Then
                                    
                                    sChe = F_GravaCh(sSeq, c.num_che, c.cod_bco, c.age_che, c.cco_che, _
                                            c.nome_che, CCur(c.Valor_che), c.Venc_che, c.cod_bco, Left(cmbxTipo.Text, 1), _
                                            "", "", dtNot.Recordset!parcela_itd, "", "")
                                            
                                            
                                    If frmBaixaPar.optnCaixa.Value = True Then
                                        sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa doc " & dtNot.Recordset!num_doc _
                                    & "-" & dtNot.Recordset!nome, "", "RCH", strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, dtNot.Recordset!cod_pla)
                                    End If
                                            
                                    If F_GravaBaixaItd(dtNot.Recordset!seq_itd, CStr(c.Valor_che), "CHP", sChe) = "" Then
                                            MsgBox "Erro na gravação da baixa!", vbExclamation
                                        End If
                                Else
                                    sChe = F_GravaChePag(sSeq, c.num_che, c.cod_bco, _
                                            CCur(c.Valor_che), c.Venc_che, dtNot.Recordset!Codigo, dtNot.Recordset!parcela_itd)
                                End If
                            Else
                                sSeq = MaxKeySQL("tab_venfin", "num_ven", "num_ven like 'B%'", "B00000#")
                                If F_GravaVenFin(sSeq, c.cco_che, c.Valor_che, Date, "", "1", "B") Then
                                End If
                                
                                If sEsp = "3" Then
                                    sEspBai = "RCC"
                                Else
                                    sEspBai = "RCD"
                                End If
                                
                                If frmBaixaPar.optnCaixa.Value Then
                                    sSeqCai = F_GravaCai("C", Date, CStr(c.Valor_che), "Baixa doc " & dtNot.Recordset!num_doc _
                                    & "-" & dtNot.Recordset!nome, sSeq, _
                                     sEspBai, strCod_Cai, 0, 0, strCod_Cai, strCod_Tur, dtNot.Recordset!cod_pla)
                                End If
                                
                                If F_GravaBaixaItd(dtNot.Recordset!seq_itd, CStr(c.Valor_che), "FIN", sSeq) = "" Then
                                            MsgBox "Erro na gravação da baixa!", vbExclamation
                                        End If
                            End If
                            
                            cValor = cValor + c.Valor_che
                        Next
                        
                        
                            strSQL = "update tab_itdoc set reman_itd = reman_itd-" & F_GravaDec(cValor - CCur(frmBaixaPar.txtbxJuros.Text) + CCur(frmBaixaPar.txtbxDesconto.Text))
                            strSQL = strSQL & ",juros_itd=" & F_GravaDec(frmBaixaPar.txtbxJuros.Text)
                            strSQL = strSQL & ",Desconto_itd=" & F_GravaDec(frmBaixaPar.txtbxDesconto.Text)
                            strSQL = strSQL & ", baixa_itd=" & F_GravaData(Date)
                            strSQL = strSQL & " where seq_itd='" & dtNot.Recordset!seq_itd & "' "
                            Banco.Execute strSQL
                        
                        
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
    Dim sTab As String
    Dim sCod As String
    Dim sNome As String
    Dim sTipoTer As String
    
    
    If cmbxOrder.ListIndex = 0 Then
        sOrderBy = "venc_Itd"
    ElseIf cmbxOrder.ListIndex = 1 Then
        sOrderBy = "Nome"
    ElseIf cmbxOrder.ListIndex = 2 Then
        sOrderBy = "num_doc,venc_itd"
    Else
        sOrderBy = "reman_itd"
    End If
    
        If cmbxTipoTer.ListIndex = 0 Then
            sTab = "tab_clientes"
            sCod = "cod_cli"
            sNome = "nome_cli"
            sTipoTer = "C"
        Else
            sTab = "tab_representantes"
            sNome = "nome_rep"
            sCod = "cod_rep"
            sTipoTer = "R"
        End If
        strSQL = "select tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
        strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
        strSQL = strSQL & " tcc.nome_cob,td.cod_ter,"
        strSQL = strSQL & "td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
        If cmbxBuscar.ListIndex = 3 Then
            strSQL = strSQL & ", tb.* "
        End If
        strSQL = strSQL & "from " & sTab & " as tc, tab_doc as td, "
        strSQL = strSQL & "tab_itdoc as ti,tab_parametros as tp, "
        strSQL = strSQL & "tab_tipocob as tt,tab_contacob as tcc,tab_planocontas as tpc "
        If cmbxBuscar.ListIndex = 3 Then
            strSQL = strSQL & ",tab_boletos as tb "
        End If
        strSQL = strSQL & "where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
        strSQL = strSQL & "and td.tipo_doc='R' "
        strSQL = strSQL & " and tt.cod_tip=ti.cod_tip "
        strSQL = strSQL & " and td.cod_pla=tpc.cod_pla and tpc.tipo_pla='C' "
        strSQL = strSQL & " and tcc.cod_cob=ti.cod_cob "
        strSQL = strSQL & " and td.tipo_ter='" & sTipoTer & "' "
        If chkAberto.Value = 1 Then
            strSQL = strSQL & " and ti.reman_itd > 0 "
        End If
        If txtbxTipoCob.tag <> "" Then
            strSQL = strSQL & " and tt.cod_tip='" & txtbxTipoCob.tag & "' "
        End If
        If txtbxContaCob.tag <> "" Then
            strSQL = strSQL & " and cod_cob='" & txtbxContaCob.tag & "' "
        End If
        
        If txtbxdataFin.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd <= " & F_GravaData(txtbxdataFin.Text)
        End If
        If txtbxDataIni.Text <> "" Then
            strSQL = strSQL & " and ti.venc_itd >= " & F_GravaData(txtbxDataIni.Text)
        End If
        If txtbxTerceiro.tag <> "" Then
            strSQL = strSQL & " and tc." & sCod & "  =  '" & txtbxTerceiro.tag & "' "
        End If
        If txtbxContaCob.tag <> "" Then
            strSQL = strSQL & " and ti.cod_cob='" & txtbxContaCob.tag & "' "
        End If
        If cmbxBuscar.ListIndex = 0 Then
            strSQL = strSQL & " and td.num_doc like  '" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 1 Then
            strSQL = strSQL & " and historico_doc like  '%" & txtbxExpr.Text & "%' "
        ElseIf cmbxBuscar.ListIndex = 2 Then
            If txtbxExpr.Text <> "" Then
                strSQL = strSQL & " and ti.valor_itd = " & F_GravaDec(txtbxExpr.Text)
            End If
        ElseIf cmbxBuscar.ListIndex = 3 Then
            strSQL = strSQL & " and tb.seq_itd=ti.seq_itd and tb.num_bol like  '" & txtbxExpr.Text & "%' "
        End If
        strSQL = strSQL & " order by " & sOrderBy
    
    dtNot.RecordSource = strSQL
    dtNot.Refresh
    soma
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnCarne_Click()
    imprimeCarne dtNot.Recordset!seq_doc
End Sub

Private Sub cmndbtnCob_Click()
    buscaContaCob
End Sub
Sub buscaContaCob()
    Dim c As ContaCob
    c = F_BuscaContaCob(txtbxContaCob.Text)
    If c.cod_cob <> "" Then
        txtbxContaCob.Text = c.nome_cob
        txtbxContaCob.tag = c.cod_cob
    End If
    
    
End Sub

Private Sub cmndbtnContrato_Click()
    imprimeContrato dtNot.Recordset!seq_doc
End Sub

Private Sub cmndbtnEdit_Click()
    Dim fr As New frmEdiLan
    On Error GoTo trata_Erro
    If strBaixadoc_usu <> "S" Then
        MsgBox "Você não tem permissão para esta operação!", vbCritical
        Exit Sub
    End If
    With dtNot.Recordset
        If .RecordCount < 0 Then
            Exit Sub
        End If
        fr.tag = !seq_itd
        fr.lblData.caption = !data_doc
        fr.lblNUM_Doc.caption = !num_doc
        fr.lblParc.caption = !parcela_itd & "/" & !parcelas_doc
        fr.lblSeq_Doc.caption = !seq_doc
        fr.lblSeq_Itd.caption = !seq_itd
        fr.lblValor.caption = FormatVal(!valor_itd)
        fr.txtbxHist.Text = !historico_doc
        fr.txtbxNome.Text = !nome
        fr.txtbxNome.tag = !Codigo
        fr.txtbxVenc.Text = !venc_itd
        fr.txtbxValor.Text = FormatVal(!reman_itd)
        fr.txtbxBoleto.Text = !boleto_itd
        fr.txtbxClass.tag = !cod_pla
        fr.txtbxClass.Text = !nome_pla
        fr.lblTipo.caption = "Crédito"
        fr.Show 1
        dtNot.Refresh
    End With
    Exit Sub
trata_Erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmndbtnExcl_Click()
    If strExcldoc_usu <> "S" Then
        MsgBox "Você não pode excluir este documento!", vbExclamation
        Exit Sub
    End If
    
    With dtNot.Recordset
        If !num_ven = "" Then
            If F_BuscaUmCampo("reman_itd", "tab_itdoc", "reman_itd<valor_itd and seq_doc='" & !seq_doc & "'") = "" Then
                If MsgBox("Deseja realmente excluir o documento e suas parcelas?", vbYesNo + vbCritical) = vbYes Then
                    strSQL = "delete from tab_doc where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    strSQL = "delete from tab_itdoc where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "' "
                    Banco.Execute strSQL
                    dtNot.Refresh
                End If
            Else
                MsgBox "Este documento não pode ser excluído por que já teve uma parcela baixada!", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Este documento não pode ser excluído pois se refere a uma venda!", vbExclamation
            
        End If
    End With
    
    
    
    
    
    
End Sub

Private Sub cmndbtnImpr_Click()
    P_DbgrdToHtml "Movimentação " & cmbxTipo.Text & " de " & txtbxDataIni.Text & " Até " & txtbxdataFin.Text, dbgrd, dtNot
End Sub

Private Sub cmndbtnNew_Click()
    Dim fr As New frmLanDoc
    fr.Show 0
End Sub

Private Sub cmndbtnSair_Click()
    Unload Me
End Sub

Private Sub DBGrid1_GotFocus()
    DBGrid1.BorderStyle = 1
End Sub

Private Sub DBGrid1_LostFocus()
    DBGrid1.BorderStyle = 0
End Sub

Private Sub cmndbtnTerc_Click()
    
    
        If cmbxTipoTer.ListIndex = 0 Then
            buscaCli
        Else
            buscaRep
        End If
    
    
    cmndbtnBuscar_Click
End Sub
Sub buscaCli()
    Dim c As Cliente
    c = F_BuscaCliente(txtbxTerceiro.Text)
        If c.cod_cli <> "" Then
            txtbxTerceiro.Text = c.nome_Cli
            txtbxTerceiro.tag = c.cod_cli
        End If
End Sub
Sub buscaFor()
    Dim f As fornecedor.fornecedor
        f = F_SelectForn(txtbxTerceiro.Text)
        If f.cod_for <> "" Then
            txtbxTerceiro.Text = f.nome_for
            txtbxTerceiro.tag = f.cod_for
        End If
End Sub
Sub buscaRep()
    Dim r As representante.representante
        r = representante.F_BuscaRep(txtbxTerceiro.Text)
        If r.cod_rep <> "" Then
            txtbxTerceiro.Text = r.nome_rep
            txtbxTerceiro.tag = r.cod_rep
        End If
    
End Sub

Private Sub cmndbtnTerceiro_Click()
    
End Sub

Private Sub cmndbtnTipoCob_Click()
    buscaTipoCob
End Sub
Sub buscaTipoCob()
    Dim t As tipocob
    t = F_BuscaTipoCob(txtbxTipoCob.Text)
    If t.cod_tip <> "" Then
        txtbxTipoCob.Text = t.nome_tip
        txtbxTipoCob.tag = t.cod_tip
    End If
End Sub


Private Sub dbgrd_DblClick()
        With dtNot.Recordset
        
        If !reman_itd <= 0 Then
            Exit Sub
        End If
        If IsLoad("frmListaDoc") Then
            If frmListaDoc.tag <> Me.Name Then
                MsgBox "A lista está sendo usada por outro formulário!", vbExclamation
                Exit Sub
            End If
            If frmListaDoc.lblCliente.tag <> "" And _
                frmListaDoc.lblCliente.tag <> !cod_ter Then
                MsgBox "Você só pode selecionar parcelas do mesmo Terceiro!", vbExclamation
                Exit Sub
            End If
            frmListaDoc.lblCliente.tag = !cod_ter
            frmListaDoc.lblCliente.caption = !nome
            
            If F_InFlxgrd(frmListaDoc.msflxgrd, !num_doc, 0) = -1 Then
                frmListaDoc.msflxgrd.AddItem !num_doc & Chr(9) & !reman_itd
            End If
            frmListaDoc.soma
        Else
            
            frmListaDoc.tag = Me.Name
            frmListaDoc.lblCliente.tag = !cod_ter
            frmListaDoc.lblCliente.caption = !nome
            frmListaDoc.caption = "Documentos"
            frmListaDoc.msflxgrd.Cols = 2
            frmListaDoc.msflxgrd.rows = 0
            frmListaDoc.msflxgrd.AddItem !num_doc & Chr(9) & !reman_itd
            frmListaDoc.Show 0, mdiFrm
            frmListaDoc.soma
        End If
    
    End With

End Sub

Private Sub dbgrd_GotFocus()
    dbgrd.MarqueeStyle = 3
End Sub

Private Sub dbgrd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        dbgrd_DblClick
    End If
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
        strSQL = "select * from tab_cheques where seq_doc='" & !seq_doc & "' "
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
        If !tipo_doc = "R" Then
            cmndbtnCarne.Enabled = True
        Else
            cmndbtnCarne.Enabled = False
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
    resizeForm Me
    txtbxDataIni.Text = F_PrimDiaMes(Date)
    
    txtbxdataFin.Text = F_UltDiaMes(Date)
    dtNot.Connect = Banco.Connect
    cmbxTipo.ListIndex = 0
    cmbxOrder.ListIndex = 0
    dtNot.Connect = Banco.Connect
    dtChe.Connect = Banco.Connect
    'P_CentraForm Me
'    strSQL = "select tn.*,tt.* from tab_compras as tn, tab_fornecedores as tt "
'    strSQL = strSQL & " where tt.cod_for=tn.cod_for "
'
'    strSQL = strSQL & " order by num_not "
'    dtNot.RecordSource = strSQL
'    dtNot.Refresh
    cmbxBuscar.ListIndex = 1
    cmbxOrder.ListIndex = 0
    cmbxTipo.ListIndex = 0
    cmbxTipoTer.ListIndex = 0
    cmndbtnBuscar_Click
End Sub

Private Sub optnCli_Click()
    frExpr.Enabled = True
End Sub

Private Sub optnTodas_Click()
    frExpr.Enabled = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    dbgrd.Width = Me.Width - dbgrd.Left - 300
    dbgrdChe.Top = pcTotal.Top - dbgrdChe.Height - 100
    dbgrd.Height = dbgrdChe.Top - dbgrd.Top - 100
    dbgrdChe.Width = dbgrd.Width
    
End Sub

Private Sub pcTop_GotFocus()
    SendKeys Chr(9)
End Sub

Private Sub txtbxContaCob_Change()
    txtbxContaCob.tag = ""
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

Private Sub txtbxTerceiro_Change()
    txtbxTerceiro.tag = ""
End Sub

Private Sub txtbxTerceiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmndbtnTerc_Click
    End If
End Sub

Private Sub txtbxTipoCob_Change()
    txtbxTipoCob.tag = ""
End Sub
