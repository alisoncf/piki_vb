VERSION 5.00
Begin VB.Form frmAjuste 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command35 
      Caption         =   "Perdidos"
      Height          =   1575
      Left            =   7440
      TabIndex        =   39
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Clientes CPF"
      Height          =   855
      Left            =   7320
      TabIndex        =   38
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Command33"
      Height          =   735
      Left            =   5520
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Setores Trabalho"
      Height          =   735
      Left            =   3360
      TabIndex        =   36
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Decimal para Double"
      Height          =   615
      Left            =   1080
      TabIndex        =   35
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Visivel"
      Height          =   735
      Left            =   1680
      TabIndex        =   34
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   6360
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   615
         Left            =   2520
         TabIndex        =   33
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Adiciona Estoque"
         Height          =   855
         Left            =   2640
         TabIndex        =   32
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Doc sem itdoc"
         Height          =   615
         Left            =   2760
         TabIndex        =   31
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command26 
         Caption         =   "dif itvenas <> valor venda"
         Height          =   615
         Left            =   0
         TabIndex        =   30
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Descrição dos itvendas"
         Height          =   615
         Left            =   5880
         TabIndex        =   29
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Comissoes dos Consignados"
         Height          =   735
         Left            =   2880
         TabIndex        =   28
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   1335
         Left            =   0
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Busca Cupom Sem Venda"
         Height          =   615
         Left            =   0
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   855
         Left            =   2640
         TabIndex        =   24
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   1095
         Left            =   1200
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Transf"
         Height          =   975
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmndbtnDelCli 
         Caption         =   "Deletar Clientes"
         Height          =   735
         Left            =   2160
         TabIndex        =   21
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Consignados"
         Height          =   855
         Left            =   2400
         TabIndex        =   20
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Itd"
         Height          =   495
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Ajuste COM."
         Height          =   615
         Left            =   2520
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         Caption         =   "BaixaItd"
         Height          =   735
         Left            =   2520
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Data Comiss Perm"
         Height          =   975
         Left            =   2400
         TabIndex        =   16
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Gera Campos Comiss"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Cobrança = 3 onde = R"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmnAjusteCom 
         Caption         =   "Ajuste Comissao"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command21 
         Caption         =   "cobradores com 3 digitos"
         Height          =   735
         Left            =   2400
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Comiss Par"
         Height          =   495
         Left            =   960
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command23 
         Caption         =   "ajusta valor das vendas"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Valor do ItVendas"
         Height          =   615
         Left            =   1680
         TabIndex        =   9
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Código dos produtos em itvendas"
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Command29"
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   3000
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmndbtnVerificaCons 
      Caption         =   "Command30"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Compras"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   1455
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ajuste Venfin"
      Height          =   855
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "busca NF sem complemento"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmnAjusteCom_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from tab_comisscons"
    Conn.Execute strSQL
    strSQL = "select * from tab_baixaitd2 "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            P_VerificaBaixaComissProv !seq_itd, !valor_bai
            .MoveNext
        Loop
    End With
End Sub

Private Sub cmndbtnDelCli_Click()
    Dim strSQL As String
    Dim sCod As String
    Dim rPesq As New ADODB.Recordset
    Dim iCount As Integer
    Text1.Text = ""
    strSQL = "select cod_cli,nome_cli,cpf_cli from tab_clientes order by cod_cli"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
            strSQL = F_BuscaUmCampo("seq_ace", "tab_acertos", "cod_Cli='" & !cod_cli & "' and tipo_ter='C'")
            If strSQL = "" Then
                strSQL = F_BuscaUmCampo("num_Ven", "tab_vendas", "cod_Cli='" & !cod_cli & "'")
                
                If strSQL = "" Then
                    strSQL = F_BuscaUmCampo("seq_doc", "tab_doc", "cod_Ter='" & !cod_cli & "' and tipo_ter='C'")
                    
                    If strSQL = "" Then
                        'strSQL = F_BuscaUmCampo("cod_cli", "tab_clientes", " nome_cli='" & !nome_Cli & "' and cod_Cli<>'" & !cod_cli & "'")
                        strSQL = !cpf_cli
                        
                        If strSQL = "" Then
                            
                            Text1.Text = Text1.Text & !cod_cli & "-" & !nome_Cli & Chr(13) + Chr(10)
                            strSQL = "delete from tab_clientes where cod_cli='" & !cod_cli & "' "
                            Conn.Execute strSQL
                            strSQL = "delete from tab_telefones where cod_cli='" & !cod_cli & "' and tipo_ter='C' "
                            Conn.Execute strSQL
                            
                            iCount = iCount + 1
            
                        End If
                    Else
'                        MsgBox "Doc"
                    End If
                Else
'                    MsgBox "Venda"
                End If
            Else
'                MsgBox " Acerto "
            End If
            
            
            
            Me.caption = .AbsolutePosition + 1 & " Registros Escaniados - " & iCount + 1 & " Registros Econtrados - " & FormatVal((.AbsolutePosition + 1) / (.RecordCount)) * 100 & " % Scaniados - " & FormatVal((iCount + 1) / (.AbsolutePosition + 1)) * 100 & " % Encontrados "
            DoEvents
            .MoveNext
        Loop
        MsgBox iCount + 1 & " Clientes Inativos!", vbExclamation
        .Close
    End With
    
    
End Sub

Private Sub Command1_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "show tables"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "delete from " & .Fields(0).Value
'            conn.execute strSQL
            .MoveNext
        Loop
    End With
End Sub

Private Sub Command10_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sSEQ_DOC As String
    
    Dim sSeq As String
    
    strSQL = "select tv.*,tc.nome_cli from tab_vendas as tv, tab_clientes as tc "
    strSQL = strSQL & " where tc.cod_cli=tv.cod_cli and cod_for='04' and situacao_ven=4"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            sSEQ_DOC = F_GravaDoc("VENDA " & !num_ven & " " & Date & " - " & !nome_Cli, _
            !num_ven, !valor_ven, !dataefet_ven, !cod_cli, "R", !num_ven, "1", "V", "C", "000")
                
                    strSQL = F_GravaItDoc(sSEQ_DOC, !valor_ven, CStr(DateAdd("D", 30, !data_ven)), CStr(DateAdd("D", 30, !data_ven)), "1", "V")
                
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command11_Click()
    
'    Dim strSQL As String
'    Dim rPesq As New ADODB.Recordset
'    Dim iCount As Integer
'    Text1.text = ""
'    strSQL = "select * from tab_doc order by seq_doc "
'    abreObjetoAdo rPesq, strSQL
'    With rPesq
'
'        Do While Not .EOF
'            strSQL = F_BuscaUmCampo("seq_itd", "tab_itdoc", "seq_doc='" & !seq_Doc & "'")
'            If strSQL = "" Then
'                strSQL = F_BuscaUmCampo("seq_che", "tab_cheques", "seq_doc='" & !seq_Doc & "'")
'            End If
'            If strSQL = "" Then
'                Text1.text = Text1.text & !seq_Doc & "-" & !num_doc & Chr(13) + Chr(10)
'
'
'                Me.caption = .RecordCount + 1 & " / " & .AbsolutePosition
'                For iCount = 1 To !parcelas_doc
'                    strSQL = F_GravaItDoc(!seq_Doc, Round(!valor_doc / !parcelas_doc, 2), _
'                    (DateAdd("D", 30 * iCount, !data_doc)), (DateAdd("D", 30 * iCount, !data_doc)), CStr(iCount), "V")
'                Next
'            End If
'            .MoveNext
'                Me.caption = .RecordCount + 1 & " / " & .AbsolutePosition
'            Me.Refresh
'        Loop
'        .Close
    'End With


End Sub

Private Sub Command12_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rVen  As New ADODB.Recordset
    Dim cVendido As Currency
    strSQL = "select ti.*,sum(ti.qtd_ite) as qtd from tab_itcompras as ti group by ti.cod_pro,ti.lote_ite "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "select sum(ti.qtd_ite) as total from tab_itvendas as ti, tab_vendas as tv "
            strSQL = strSQL & " where tv.num_ven=ti.num_ven "
            strSQL = strSQL & " and ti.cod_Pro='" & !cod_pro & "' "
            strSQL = strSQL & " and ti.lote_ite='" & !lote_ite & "'"
            Set rVen = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
            With rVen
                
                    If IsNull(!total) Then
                        cVendido = 0
                    Else
                        cVendido = !total
                    End If
                    
                    
                
                .Close
            End With
            
            
            
            strSQL = "update tab_estoque set er_pro=0 -" & cVendido & "  where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and lote_pro='" & !lote_ite & "' "
            Conn.Execute strSQL
            
            .MoveNext
            
        Loop
        .Close
    End With
    
    
    
    
    
End Sub

Private Sub Command13_Click()
    Dim strSQL As String
    Dim str As String
    Dim rPesq As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    
    
    strSQL = "select te.*,tp.pv_pro from tab_estcons as te,tab_produtos as tp where tp.cod_pro=te.cod_pro and estoque_pro > 0 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            strSQL = "select * from tab_consignacoes where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and reman_con > 0 "
            strSQL = strSQL & " and cod_rep='" & !cod_Rep & "' "
            Set r2 = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
            If r2.RecordCount > 0 Then
            Else
                
                strSQL = "update tab_estcons set estoque_pro=0 where cod_pro='" & !cod_pro & "' "
                strSQL = strSQL & " and cod_rep='" & !cod_Rep & "' "
                Conn.Execute strSQL
                
                str = MaxKeySQL("tab_consignacoes", "seq_con", "", "000000#")
                str = F_GravaConsign(str, !cod_pro, !estoque_pro, !cod_Rep, !pv_pro, True)
                
                
                Text1.Text = Text1.Text & !cod_Rep & "-" & !cod_pro & Chr(13) + Chr(10)
            
            End If
            
            
            .MoveNext
        Loop
        .Close
    End With
    
    
    
    
End Sub

Private Sub Command14_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "update tab_estcons set estoque_pro=0"
    
    Conn.Execute strSQL
    
    strSQL = "select * from tab_consignacoes where reman_con > 0 order by seq_con"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_estcons set estoque_pro=estoque_pro+" & F_GravaDec(!reman_con)
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and cod_rep='" & !cod_Rep & "'"
            Conn.Execute strSQL
            Me.caption = .AbsolutePosition + 1 & " de " & .RecordCount
            DoEvents
            .MoveNext
            
        Loop
        .Close
    End With
    
    
    
End Sub

Private Sub Command15_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim iSeq As Integer
    On Error Resume Next
    strSQL = "select seq_itd,seq_bai from tab_baixaitd order by seq_itd,seq_bai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            iSeq = .AbsolutePosition + 1
            strSQL = "update tab_baixaitd set seq_bai=1,"
            strSQL = strSQL & "seq_cab='" & Format(iSeq, "000000#") & "' "
            strSQL = strSQL & " where seq_itd='" & !seq_itd & "' "
            strSQL = strSQL & " and seq_bai=" & !seq_bai
            
            Conn.Execute strSQL
            .MoveNext
        Loop
        
        strSQL = "insert into tab_codigos(table_cod,field_cod,value_cod)values("
        strSQL = strSQL & "'tab_baixaitd','seq_cab','" & Format(iSeq + 1, "000000#") & "') "
        Conn.Execute strSQL
        .Close
    End With
    
End Sub

Private Sub Command16_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select * from tab_comissprov "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_comissprov set seq_com='" & Format(!seq_com, "000000#") & "'"
            strSQL = strSQL & " where seq_com='" & !seq_com & "' "
            Conn.Execute strSQL
            
            strSQL = "update tab_comisscons set seq_com='" & Format(!seq_com, "000000#") & "'"
            strSQL = strSQL & " where seq_com='" & !seq_com & "' "
            Conn.Execute strSQL
            
            
            .MoveNext
        Loop
    End With
    
End Sub

Private Sub Command17_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_baixaitd"
    Dim sSeqCab As String
    Dim sSeq As String
    Dim iCount As Integer
    abreObjetoAdo rPesq, strSQL
     iCount = 0
    With rPesq
        Do While Not .EOF
            sSeqCab = F_BuscaChaveBaixa
            
            If F_BuscaUmCampo("Seq_itd", "tab_baixaitd2", "seq_itd='" & !seq_itd & "'") = "" Then
                strSQL = F_GravaBaixaItd(sSeqCab, !seq_itd, !valor_bai, !esp_bai, !seq_fin, !data_bai)
                
            End If
           
            If F_BuscaUmCampo("Seq_fin", "tab_baixafin", "seq_fin='" & !seq_fin & "' and valor_bai>0") = "" Then
                strSQL = F_GravaBaixaItdFin(sSeqCab, !seq_fin, !esp_bai, !valor_bai)
            End If
            
            
            .MoveNext
        Loop
        MsgBox iCount
        .Close
    End With
End Sub

Private Sub Command18_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rPesq2 As New ADODB.Recordset
    strSQL = "select distinct tc.*,tb.data_bai from tab_comisscons as tc,tab_baixaitd2 as tb "
    strSQL = strSQL & " where tc.seq_itd=tb.seq_itd "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_comisscons set data_com=" & F_GravaData(!data_bai)
            strSQL = strSQL & " where seq_itd='" & !seq_itd & "' "
            Conn.Execute strSQL
            
            .MoveNext
        Loop
    End With
    
    
    
End Sub

Private Sub Command19_Click()
    Dim strSQL As String
    strSQL = "ALTER TABLE `tab_comisscons` ADD `status_com` CHAR(1)  DEFAULT '1' NOT NULL"
    Conn.Execute strSQL
    strSQL = "ALTER TABLE `tab_comisscons` ADD `data_com` DATE NOT NULL"
    Conn.Execute strSQL
    strSQL = "ALTER TABLE `tab_comisscons` ADD `seq_cai` char(7) not null"
    Conn.Execute strSQL
    
    
End Sub

Private Sub Command2_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_compras order by num_not "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("seq_not", "tab_cmp", "num_not='" & !Num_not & "'") = "" Then
                Text1.Text = Text1.Text & !Num_not & Chr(13) + Chr(10)
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub Command20_Click()
    strSQL = "update tab_acertos set cobranca_ace=3 where cobranca_ace='R'"
    Conn.Execute strSQL
    strSQL = "update tab_acertos set cobranca_ace=1 where cobranca_ace='L'"
    Conn.Execute strSQL
End Sub

Private Sub Command21_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select * from tab_setores "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
            strSQL = "Update tab_setores set cobrador_set='" & Format(!cobrador_set, "00#") & "' "
            strSQL = strSQL & " where cod_set='" & !cod_set & "'"
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command22_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from tab_Comissparc "
    Conn.Execute strSQL
    strSQL = "select td.valor_doc,ta.seq_ace,parcelas_doc,ta.comissao_ace,ta.cod_rep,ti.* "
    strSQL = strSQL & " from tab_acertos as ta,tab_itdoc as ti,tab_doc as td  "
    strSQL = strSQL & " where ta.seq_ace=td.num_ven "
    strSQL = strSQL & " and td.esp_doc='A'"
    strSQL = strSQL & " and ta.tipo_ter='C'"
    strSQL = strSQL & " and ti.seq_doc=td.seq_doc "
    strSQL = strSQL & " and data_ace>'2008-05-01'"
    strSQL = strSQL & " and comissao_ace >0"
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If Not F_GravaComissParc(!seq_itd, !cod_Rep, !seq_ace, !comissao_ace / !parcelas_doc, Format(!venc_itd), "1") Then
                MsgBox "Erro"
            End If
            Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
            DoEvents
            
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Command23_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sFinal As String
    
    strSQL = "select num_ven from tab_vendas where situacao_ven=4 and valor_ven=0"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            Me.caption = .AbsolutePosition + 1 & " de " & .RecordCount
            
            sFinal = F_BuscaUmCampo("sum(final_ite)", "tab_itvendas", "num_ven='" & !num_ven & "'")
            strSQL = "update tab_vendas set valor_ven=" & F_GravaDec(sFinal)
            strSQL = strSQL & " where num_ven='" & !num_ven & "' "
            Conn.Execute strSQL
            
            .MoveNext
            
            DoEvents
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Command24_Click()
    Dim strSQL As String
    strSQL = "update tab_itvendas set unit_ite=(final_ite/qtd_ite) where unit_ite=0"
    Conn.Execute strSQL
End Sub

Private Sub Command25_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select distinct cod_pro from tab_itvendas where length(cod_pro) < 8 "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "Update tab_itvendas set cod_pro='" & Format(!cod_pro, "0000000#") & "' "
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
            Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
            DoEvents
            .MoveNext
        Loop
        
        
        .Close
    End With
    
    
End Sub

Private Sub Command26_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ti.*,tv.valor_ven,tv.desconto_ven "
    strSQL = strSQL & " from tab_itvendas ti,"
    strSQL = strSQL & " tab_vendas as tv "
    strSQL = strSQL & " where tv.num_ven=ti.num_ven "
    strSQL = strSQL & " and tv.situacao_ven=4 "
    strSQL = strSQL & " group by tv.num_ven "
    strSQL = strSQL & " having sum(final_ite) <> valor_ven "
    
    abreObjetoAdo rPesq, strSQL
    rPesq.Close
End Sub

Private Sub Command27_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select distinct tp.cod_pro,tp.nome_pro from tab_produtos as tp, tab_itvendas as ti "
    strSQL = strSQL & " where ti.cod_pro=tp.cod_pro "
    strSQL = strSQL & " and desc_ite=''"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_itvendas set desc_ite='" & !nome_pro & "' "
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            strSQL = strSQL & " and desc_ite='' "
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Private Sub Command28_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ta.seq_ace, ta.valor_ace,"
    strSQL = strSQL & " ta.comissao_ace,"
    strSQL = strSQL & " sum(tp.valor_com) as total,"
    strSQL = strSQL & " count(*) as qtd,"
    strSQL = strSQL & " comissao_ace - sum(tp.valor_com) as dif"
    strSQL = strSQL & " From  tab_acertos as ta,"
    strSQL = strSQL & " tab_comissparc As tp"
    strSQL = strSQL & " where"
    strSQL = strSQL & " ta.seq_ace = tp.seq_ace"
'    strSQL = strSQL & " and status_com=1"
    strSQL = strSQL & " Group By"
    strSQL = strSQL & " ta.seq_ace"
    strSQL = strSQL & " having comissao_Ace>total+.5"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_comissparc set valor_com=" & F_GravaDec(!comissao_ace / !qtd)
            strSQL = strSQL & " where seq_ace='" & !seq_ace & "' "
            'strSQL = strSQL & " and status_com='1' "
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Command3_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from tab_itdoc "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("cod_pro", "tab_codbarras", "cod_pro='" & !cod_pro & "'") = "" Then
                strSQL = "insert into tab_codbarras(cod_pro,codbarra_pro)values('" & !cod_pro & "',"
                strSQL = strSQL & "'" & !cod_pro & "') "
                Conn.Execute strSQL
            End If
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Private Sub Command30_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sVisivel As String
    strSQL = "select * from tab_caixa order by seq_cai "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If !num_ven <> "" And Left(!historico_cai, 6) <> "ACERTO" Then
                If F_BuscaUmCampo("nota_sai", "tab_saidas", "num_ven='" & !num_ven & "'") <> "" Then
                    sVisivel = "S"
                Else
                    sVisivel = "N"
                End If
            ElseIf !num_ven <> "" And Left(!historico_cai, 6) = "ACERTO" Then
                sVisivel = "N"
            Else
                sVisivel = "S"
            End If
            
            strSQL = "update tab_Caixa set visivel_cai='" & sVisivel & "'"
            strSQL = strSQL & " where seq_cai='" & !seq_cai & "'"
            Conn.Execute strSQL
            Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
            Me.Refresh
            .MoveNext
        Loop
        .Close
    End With
    
    
    
End Sub

Private Sub Command31_Click()
    Dim strSQL As String
    Dim rTable As New ADODB.Recordset
    Dim rPesq As New ADODB.Recordset
    strSQL = "show tables"
    Dim iCount As Integer
    Set rTable = Banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rTable
        Do While Not .EOF
            strSQL = "select * from " & .Fields(0).Value & " limit 1"
            abreObjetoAdo rPesq, strSQL
            With rPesq
                Do While Not .EOF
                    For iCount = 0 To .Fields.Count - 1
                        If TypeName(.Fields(iCount).Name) = "Decimal" Then
                            MsgBox ""
                        End If
                    Next
                    .MoveNext
                Loop
                .Close
            End With
            .MoveNext
        Loop
        .Close
    End With
    
End Sub

Private Sub Command32_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    Dim iCount As Integer
    strSQL = "select tc.cod_cli,torg.cod_set from tab_clientes as tc,"
    strSQL = strSQL & " tab_orgaos as torg where tc.cod_org<> ''"
    strSQL = strSQL & " and torg.cod_org=tc.cod_org "
    strSQL = strSQL & " and torg.cod_set<>tc.setortrab_cli "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_clientes set setortrab_cli='" & !cod_set & "' "
            strSQL = strSQL & " where cod_cli='" & !cod_cli & "'"
            Conn.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command33_Click()
    Dim strSQL As String
    Dim b As New ADODB.Connection
    Dim rPesq As New ADODB.Recordset
    Dim r2 As New ADODB.Recordset
    b.Open "requinte_back", "root"
    b.CursorLocation = adUseClient
    
    
    
    strSQL = "select endereco_cli,cod_cli from tab_clientes where length(endereco_cli) <21"
    
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "select endereco_cli from tab_clientes where cod_cli='" & !cod_cli & "'"
            FechaObjeto r2
            r2.Open strSQL, b
            
            If r2.RecordCount <> 0 Then
                If Len(r2!endereco_cli) > Len(rPesq!endereco_cli) Then
                    strSQL = "update tab_clientes set endereco_cli='" & r2!endereco_cli & "' where cod_cli='" & !cod_cli & "'"
                    Conn.Execute strSQL
                End If
                Command33.caption = .RecordCount & " - " & .AbsolutePosition
                Command33.Refresh
                Me.Refresh
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Command34_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim rItens As New ADODB.Recordset
    strSQL = "select cod_cli,nome_cli,cpf_cli from " & TAB_CLIENTES & " where cpf_cli='' order by cod_cli "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            If IsNull(!cpf_cli) Or !cpf_cli = "" Then
            
                strSQL = "select sum(reman_itd) as total from tab_doc as td inner join tab_itdoc as ti on td.seq_doc=ti.seq_doc "
                strSQL = strSQL & " where tipo_ter='C' and cod_ter='" & !cod_cli & "'"
                abreObjetoAdo rItens, strSQL
                strSQL = TiraNull(rItens!total)
                
                If strSQL = "0" Or strSQL = "" Then
                    strSQL = !cod_cli
                    If strSQL <> "0" Then
                        strSQL = "delete from " & TAB_CLIENTES & " where cod_cli=" & !cod_cli & ""
                    Conn.Execute strSQL
                    End If
                Else
                '    MsgBox ""
                End If
                        
                
                    
                    
                    
                
            
            End If
'            strSQL = "Update " & TAB_CLIENTES & " set cpf_cli='" & F_SoNumero(!cpf_cli) & "' where cod_cli='" & !cod_cli & "'"
          '  Conn.Execute strSQL
                
            
            Me.caption = .AbsolutePosition & "/" & .RecordCount
            Me.Refresh
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub

Private Sub Command4_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select cod_pro,codfab_pro from tab_produtos where codfab_pro like '00%'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "Update tab_produtos set codfab_pro=" & !codfab_pro
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub Command5_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    strSQL = "select  * from tab_saidas where emissao_sai >='2006-01-01' order by num_ven "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If F_BuscaUmCampo("situacao_ven", "tab_vendas", "num_ven='" & !num_ven & "'") = "3" Then
                cTotal = cTotal + !valor_sai
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    MsgBox cTotal
End Sub

Private Sub Command6_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sIn As String
    strSQL = "select seq_doc from tab_docback "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            sIn = sIn & "'" & !seq_Doc & "'"
            .MoveNext
            
            If Not .EOF Then
                sIn = sIn & ","
            
                
            End If
        Loop
        sIn = "(" & sIn & ")"
    End With
    strSQL = "select * from tab_doc where seq_doc not in" & sIn
    abreObjetoAdo rPesq, strSQL
    Text1.Text = ""
    With rPesq
        Do While Not .EOF
            Text1.Text = Text1.Text & !seq_Doc & "-" & !data_doc & "-" & !valor_doc & Chr(13)
            .MoveNext
        
        Loop
        MsgBox "Foram encontrados " & .RecordCount & " registros"
        .Close
        
    End With
End Sub

Private Sub Command8_Click()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "delete from tab_estoque"
    Conn.Execute strSQL
    strSQL = "select cod_pro from tab_produtos "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            'strSQL = Produtos.F_AddEstoque(!cod_pro, "01/01/1900", "0", "0", "0", "0")
            .MoveNext
            Me.caption = .AbsolutePosition + 1 & "/" & .RecordCount
            Me.Refresh
            
        Loop
        .Close
    End With
End Sub


