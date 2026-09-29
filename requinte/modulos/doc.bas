Attribute VB_Name = "doc"

Option Explicit
Dim strSQL As String

Public Enum tipo_vencimento
    vencimento_da_parcela = 1
    data_documento = 2
    baixa_cheque = 3
    vencimentoInicial = 4
End Enum




Public Enum operacao_itd
    a_receber_avulso = 0
    a_pagar_avulso = 1
    acerto_consignado = 2
    acerto_loja = 3
End Enum

Type doc
    seq_doc As String
    num_doc As String
    data_doc As String
    historico_doc As String
    valor_doc As Currency
    cod_usu As String
    parcelas_doc As String
    num_ven As String
    cod_ter As String
    tipo_ter As String
    cpf_ter As String
End Type



Type itDoc
    seq_doc As String
    seq_itd As String
    doc As doc
    venc_itd As String
    vencini_itd As String
    reman_itd As Currency
    valor_itd As Currency
    marcado_itd As String
    cod_cob As String
    
End Type


Function F_GravaDocCompleto(hist_doc As String, num_doc As String, valor_doc As String, data_doc As String, _
            cod_ter As String, tipo_doc As String, num_ven As String, _
            parcelas_doc As String, esp_doc As String, tipo_ter As String, cod_pla As String) As String
            
            
    Dim sSeq As String
    Dim sItd As String
    Dim i As Integer
    sSeq = F_GravaDoc(hist_doc, num_doc, valor_doc, data_doc, cod_ter, tipo_doc, num_ven, _
    parcelas_doc, esp_doc, tipo_ter, cod_pla)
    If sSeq <> "" Then
        For i = 1 To Int(parcelas_doc)
            sItd = doc.F_GravaItDoc(sSeq, FormatVal(CCur(valor_doc) / Int(parcelas_doc)), DateAdd("M", i, data_doc), DateAdd("M", i, data_doc), _
            CStr(i), 0)
        Next
    End If
    F_GravaDocCompleto = sSeq
End Function


Function F_GravaDoc(hist_doc As String, num_doc As String, valor_doc As String, data_doc As String, _
            cod_ter As String, tipo_doc As String, num_ven As String, _
            parcelas_doc As String, esp_doc As String, tipo_ter As String, cod_pla As String, Optional cod_emp As String) As String
            
    On Error GoTo trata_Erro
    Dim strSQL As String
    Dim ret As New ADODB.Recordset
    If data_doc = "" Then
        data_doc = Date
    End If
    
    If tipo_doc = "" Then
        tipo_doc = "F" 'financeiro
    End If
    strSQL = "insert into " & TAB_DOC & " (historico_doc, Num_Doc,Valor_doc,Data_Doc,cod_ter, "
    strSQL = strSQL & "tipo_doc,num_ven,parcelas_doc,seq_doc,cod_usu,esp_doc,tipo_ter,cod_pla)Values("
    strSQL = strSQL & "'" & hist_doc & "',"
    strSQL = strSQL & "'" & num_doc & "',"
    strSQL = strSQL & F_GravaDec(valor_doc) & ","
    strSQL = strSQL & F_GravaData(data_doc) & ","
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'" & tipo_doc & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & parcelas_doc & ","
    F_GravaDoc = cod_emp & ChaveSQL("Tab_Doc", "seq_doc", "000000#")
    strSQL = strSQL & "'" & F_GravaDoc & "', "
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & esp_doc & "', "
    strSQL = strSQL & "'" & tipo_ter & "'," ' 'C=Cliente,
    strSQL = strSQL & "'" & cod_pla & "')"

    Set ret = Conn.Execute(strSQL)
    
    
    
    
    Exit Function
trata_Erro:
    If Err.Number = -2147217900 Then
        MsgBox ""
    End If

    F_GravaDoc = ""
    
End Function
Function altera_id_doc_web(seq_doc As String, id_doc_web As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    altera_id_doc_web = True
    strSQL = "Update " & TAB_DOC & " set id_doc_web=" & id_doc_web
    strSQL = strSQL & " where seq_doc='" & seq_doc & "'"
    Conn.Execute strSQL
    altera_id_doc_web = True
    Exit Function
trata_Erro:
    altera_id_doc_web = False

End Function

Function F_GravaItDoc(seq_doc As String, valor_itd As String, venc_itd As String, _
    vencini_itd As String, parcela_itd As String, operacao_itd As operacao_itd, _
    Optional boleto As String = "", Optional cod_tip As String = "1", _
    Optional cod_cob As String = "000", Optional sX As String, Optional cod_emp As String) As String
            
    Dim strSQL As String
    Dim ret As New ADODB.Recordset
    Dim seq_itd As String
    On Error GoTo trata_Erro
    If cod_emp = "" Then cod_emp = "01"
    If cod_cob = "" Then
        cod_cob = "000"
    End If
    'operacao_itd [V=venda,R=doc a receber avulso,P=a pagar avulso,C= Acerto por conta do Consignado, L - A receber por conta da loja
    
    seq_itd = MaxKeySQL(TAB_ITDOC, "seq_itd", "", "000000#")
    
    strSQL = "insert into " & TAB_ITDOC & "(seq_itd,Seq_Doc,Valor_Itd,Venc_Itd,vencini_itd,Reman_Itd,parcela_itd,"
    strSQL = strSQL & "operacao_itd,boleto_itd,cod_tip,cod_cob)Values( "
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & F_GravaDec(valor_itd) & ","
    strSQL = strSQL & F_GravaData(venc_itd) & ","
    strSQL = strSQL & F_GravaData(vencini_itd) & ","
    strSQL = strSQL & F_GravaDec(valor_itd) & ","
    strSQL = strSQL & parcela_itd & ","
    strSQL = strSQL & "'" & operacao_itd & "', "
    strSQL = strSQL & "'" & boleto & "', "
    strSQL = strSQL & "'" & cod_tip & "'," 'TIPO_COB
    strSQL = strSQL & "'" & cod_cob & "') "
    
    Set ret = Conn.Execute(strSQL)
    
    P_gravaDocHist seq_doc, "", cod_cob, cod_tip, cod_tip
    F_GravaItDoc = seq_itd
    
    Exit Function
trata_Erro:
    F_GravaItDoc = ""
    MsgBox Err.Description, vbCritical
    
End Function
Function BuscaTerceiroDoc(seq_doc As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim tipo_ter As String
    On Error Resume Next
    tipo_ter = F_BuscaUmCampo("tipo_ter", TAB_DOC, "seq_doc='" & seq_doc & "'")
    
    If tipo_ter = "C" Then
        strSQL = "select tc.nome_cli as nome from " & TAB_DOC & " as td inner join " & TAB_CLIENTES & " as tc "
        strSQL = strSQL & " on tc.cod_cli=td.cod_ter "
    Else
        strSQL = "select tr.nome_rep from " & TAB_DOC & " as td inner join tab_representantes as tc "
        strSQL = strSQL & " on tc.cod_rep=td.cod_ter "
    
    End If
    strSQL = strSQL & " where seq_doc='" & seq_doc & "' "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        BuscaTerceiroDoc = !nome
        .Close
    End With
    
    
End Function

Sub P_CancelaBaixaiItd(seq_cab As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sMsg As String
    Dim bMesmoDia As Boolean
    Dim bMesmoCaixa As Boolean
    strSQL = "select distinct tb.seq_itd,tb.seq_bai,tb.data_bai,tb.cod_usu,tb.seq_bai,tb.seq_cab,"
    strSQL = strSQL & " tb.hora_bai,tb.cod_cob,tb.cod_cai,tb.valor_bai,f.valor_bai as total,f.seq_fin,f.esp_bai,e.efetivo,td.seq_doc,td.num_doc,"
    strSQL = strSQL & " td.historico_doc from tab_doc as td inner join tab_itdoc as ti "
    strSQL = strSQL & " on td.seq_doc=ti.seq_doc inner join tab_baixaitd2 as tb on ti.seq_itd=tb.seq_itd "
    strSQL = strSQL & " left join tab_baixafin as f on tb.seq_cab=f.seq_cab "
    strSQL = strSQL & " left join " & TAB_ESPCAI & " AS e on e.sigla_cai=f.esp_bai "
    strSQL = strSQL & " where tb.seq_cab='" & seq_cab & "' "
    abreObjetoAdo rPesq, strSQL
    If Not F_IsCaixa Then
        MsgBox "Este computador não é um caixa", vbExclamation
        
        Exit Sub
    End If
    With rPesq
        strSQL = ""
        If CDate(!data_bai) = Date Then
                bMesmoDia = True
                If strCod_Cai <> !cod_cai Then
                    MsgBox "Esta baixa só pode ser cancelada no caixa id: " & !cod_cai, vbExclamation
                    Exit Sub
                Else
                    bMesmoCaixa = True
                End If
            End If
        
        If bMesmoDia Then
            If !efetivo = "S" Then
                    strSQL = F_GravaCai("D", Date, FormatVal(!total), "CANC. MD. BAIXA. DOC: " & !num_doc & BuscaTerceiroDoc(!seq_doc), "", TiraNull(!esp_bai), _
                    strCod_Cai, 0, !total, strCod_Cai, strCod_Tur, "000", "N")
            Else
                    strSQL = F_GravaCai("D", Date, FormatVal(!total), "CANC. MD. BAIXA. DOC: " & !num_doc & BuscaTerceiroDoc(!seq_doc), "", TiraNull(!esp_bai), _
                    strCod_Cai, !total, !total, strCod_Cai, strCod_Tur, "000", "N")
            End If
        Else
                strSQL = F_GravaCai("D", Date, FormatVal(!total), "CANC. DF. BAIXA. DOC: " & !num_doc & BuscaTerceiroDoc(!seq_doc), "", TiraNull(!esp_bai), _
                strCod_Cai, TiraNull(!total), TiraNull(!total), strCod_Cai, strCod_Tur, "000", "N")
        End If
           
            
        
        Do While Not .EOF
            
            
            
            sMsg = sMsg & "Seq = " & !seq_itd & " Data: " & !data_bai & " Valor Baixado: " & FormatVal(!valor_bai) & Chr(13)
            strSQL = "update tab_itdoc set reman_itd=reman_itd+ " & F_GravaDec(!valor_bai)
            strSQL = strSQL & " where seq_itd='" & !seq_itd & "'"
            Conn.Execute strSQL
            
            
            
            
            .MoveNext
        Loop
        
        .Close
    End With
    
    
    strSQL = "delete from tab_baixaitd2 where seq_cab='" & seq_cab & "'"
    Conn.Execute strSQL
    strSQL = "delete from tab_baixafin where seq_cab='" & seq_cab & "'"
    Conn.Execute strSQL
    
    MsgBox sMsg & Chr(13) & "Parcelas Canceladas!", vbExclamation
    
       
    
    
End Sub
Function F_buscaDoc(seq_doc As String) As doc.doc
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select td.* ,"
    strSQL = strSQL & " if(td.tipo_ter='C',tc.cpf_cli ,tr.cpf_rep )as cpf_ter "
    strSQL = strSQL & " from " & TAB_DOC
    strSQL = strSQL & " as td left join " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " on tc.cod_cli = td.cod_ter "
    strSQL = strSQL & " left join tab_representantes tr "
    strSQL = strSQL & " on tr.cod_rep= td.cod_ter "
    strSQL = strSQL & " WHERE td.seq_Doc='" & seq_doc & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            F_buscaDoc.cod_usu = !cod_usu
            F_buscaDoc.data_doc = !data_doc
            F_buscaDoc.historico_doc = !historico_doc
            F_buscaDoc.num_doc = !num_doc
            F_buscaDoc.num_ven = !num_ven
            F_buscaDoc.parcelas_doc = !parcelas_doc
            F_buscaDoc.seq_doc = !seq_doc
            F_buscaDoc.valor_doc = !valor_doc
            F_buscaDoc.cod_ter = !cod_ter
            F_buscaDoc.tipo_ter = !tipo_ter
            F_buscaDoc.cpf_ter = !cpf_ter
        End If
        .Close
    End With
trata_Erro:
    
End Function
Function F_buscaParcela(seq_doc As String) As doc.itDoc()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro
    
    
    
    strSQL = "select ti.* from " & TAB_ITDOC & " as ti where seq_doc='" & seq_doc & "'"
    abreObjetoAdo rPesq, strSQL
    
    ReDim it(rPesq.RecordCount) As doc.itDoc
    
    With rPesq
        If .RecordCount > 0 Then
            it(.AbsolutePosition).seq_itd = !seq_itd
            it(.AbsolutePosition).vencini_itd = !vencini_itd
            it(.AbsolutePosition).reman_itd = !reman_itd
            it(.AbsolutePosition).venc_itd = !venc_itd
            it(.AbsolutePosition).marcado_itd = !marcado_itd
            it(.AbsolutePosition).cod_cob = !cod_cob
        End If
        F_buscaParcela = it
        .Close
    End With
trata_Erro:
    
End Function
Function F_GravaBaixaItd(seq_cab As String, seq_itd As String, valor_bai As String, _
    esp_bai As String, seq_fin As String, cod_cob As String, Optional data_bai As String = "", Optional desconto_bai As String = "0", Optional cod_cai As String = "000") As String
    Dim sSeq As String
    'On Error goto trata_erro
    If data_bai = "" Then
        data_bai = Date
    End If
    Dim strSQL As String
    sSeq = MaxKeySQL("tab_baixaitd2", "seq_itd", "seq_cab='" & seq_cab & "'", "")
    
    strSQL = "insert into tab_baixaitd2(seq_cab,seq_itd,seq_bai,valor_bai,esp_bai,data_bai,hora_bai,seq_fin,cod_usu,cod_cob,desconto_bai,cod_cai)values("
    strSQL = strSQL & "'" & seq_cab & "',"
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & F_GravaDec(valor_bai) & ","
    strSQL = strSQL & "'" & esp_bai & "',"
    strSQL = strSQL & F_GravaData(data_bai) & ","
    strSQL = strSQL & "'" & Format(Time, "hh:mm") & "',"
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & cod_cob & "', "
    strSQL = strSQL & F_GravaDec(desconto_bai) & ","
    strSQL = strSQL & "'" & cod_cai & "') "
    Conn.Execute strSQL
    F_GravaBaixaItd = sSeq
trata_Erro:
End Function
Function F_GravaBaixaItdFin(seq_cab As String, seq_fin As String, esp_bai As String, valor As String, Optional desconto As String = "0") As String
    Dim sSeq As String
    'On Error goto trata_erro
    sSeq = MaxKeySQL("tab_baixafin", "seq_bai", "seq_cab='" & seq_cab & "'", "")
    strSQL = "insert into tab_baixafin(seq_cab,seq_fin,seq_bai,esp_bai,valor_bai,desconto_bai)values("
    strSQL = strSQL & "'" & seq_cab & "',"
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & esp_bai & "',"
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaDec(desconto) & ")"
    Conn.Execute strSQL
    F_GravaBaixaItdFin = sSeq
trata_Erro:
End Function


Function F_SelectDocsReceber(tipo_ter As String, Optional dataI As String = "", _
    Optional dataF As String = "", Optional boleto As String = "", _
    Optional tipo_cob As String = "", Optional cod_cob As String = "", _
    Optional num_doc As String = "", Optional historico_doc As String = "", Optional seq_itd = "", _
    Optional seq_doc = "", Optional cod_gru As String = "", _
    Optional cod_pla As String = "", Optional cod_ter As String = "", _
    Optional status As Integer = 1, Optional sOrderBy As String = "ti.venc_itd,td.cod_ter,ti.parcela_itd", _
    Optional prejuizo As String = "N", _
    Optional tipoData As tipo_vencimento = vencimento_da_parcela, _
    Optional selectCampoEspecifico As String = "", _
        Optional esp_doc As String = "", Optional listaDoc As String = "", Optional cod_cid As String = "", _
        Optional reman As String = "", _
        Optional bloq As String = "", Optional ApenasCliVinculadoEmOrgao As Boolean = False, Optional status1 As String = "", _
        Optional creditosuspenso As String = "") As String
    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim sStatus As String
    Dim sCampoData As String
    If bloq = "0" Then bloq = ""
    
    If tipoData = data_documento Then
        sCampoData = "td.data_doc"
    ElseIf tipoData = vencimentoInicial Then
        sCampoData = "ti.vencini_itd"
    Else
        sCampoData = "ti.venc_itd"
    End If
    
    
    If tipo_ter = "C" Then
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "Cpf_cli"
        sStatus = "status_cli"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
        sStatus = "status_rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
        sStatus = "status_fun"
    End If
    
    If selectCampoEspecifico = "" Then
        strSQL = "select distinct td.cod_usu,if(tc." & sStatus & "=2,'Bloqueado','Liberado') as status,nome_set,tc.cod_set,"
        
        strSQL = strSQL & "nome_cid,"
        
        
        If tipo_ter = "C" Then
            strSQL = strSQL & "nome_org,"
        End If
        
        
        strSQL = strSQL & " ta.diferenca_ace,tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias,td.tipo_ter, "
        strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
        strSQL = strSQL & " tcc.nome_cob,td.cod_ter, tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "td.valor_doc,td.parcelas_doc,td.num_doc,td.data_doc, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido,"
        strSQL = strSQL & " ti.* "
        
    Else
        strSQL = selectCampoEspecifico
    End If
    
    
    
    strSQL = strSQL & " from " & TAB_DOC & " as td inner join " & TAB_ITDOC & " as ti on td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " inner join " & sTab & " as tc on tc." & sCod & "=td.cod_ter and td.tipo_ter='" & tipo_ter & "' "
    
    If tipo_ter = "C" Then
        strSQL = strSQL & " left join tab_orgaos as tor on tc.cod_org= tor.cod_org "
    End If
    strSQL = strSQL & " left join tab_acertos as ta on td.num_doc=ta.seq_ace "
    
    strSQL = strSQL & " left join tab_setores as tset on  tc.cod_set=tset.cod_set "
    strSQL = strSQL & " left join tab_cidades as tcid on tset.cod_cid=tcid.cod_cid "
    
    
    
    
    strSQL = strSQL & " left join tab_contacob as tcc on tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " left join tab_planocontas as tpc on tpc.cod_pla=td.cod_pla and tipo_pla='C'  "
    strSQL = strSQL & " left join tab_tipocob as tt on tt.cod_tip=ti.cod_tip, "
    strSQL = strSQL & " tab_parametros as tp"
    
    
    strSQL = strSQL & " where td.tipo_doc='R' "
    
    
   
    
    If cod_cid <> "" Then
        strSQL = strSQL & " and tcid.cod_cid='" & cod_cid & "' "
    End If
    
    If ApenasCliVinculadoEmOrgao = True Then
        If tipo_ter = "C" Then
            strSQL = strSQL & " and (tc.cod_org<>0 and tc.cod_org <> '' )"
        End If
    End If
    If status = 1 Then
        strSQL = strSQL & " and ti.reman_itd > 0 "
    ElseIf status = 0 Then
        'strSQL = strSQL & " and ti.reman_itd < 0 "
    ElseIf status = -1 Then
        strSQL = strSQL & " and ti.reman_itd < ti.valor_itd "
    End If
    
    If reman <> "" Then
        strSQL = strSQL & " and ti.reman_itd=" & F_GravaDec(reman)
    End If
    
    If tipo_cob <> "" Then
        strSQL = strSQL & " and tt.cod_tip='" & tipo_cob & "' "
    End If
    If cod_cob <> "" Then
        strSQL = strSQL & " and ti.cod_cob='" & cod_cob & "' "
    End If
    
    If dataF <> "" Then
        strSQL = strSQL & " and " & sCampoData & " <= " & F_GravaData(dataF)
    End If
    If dataI <> "" Then
        strSQL = strSQL & " and " & sCampoData & " >= " & F_GravaData(dataI)
    End If
    If cod_ter <> "" Then
        strSQL = strSQL & " and tc." & sCod & "  =  '" & cod_ter & "' "
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and td.num_doc like  '" & num_doc & "%' "
    End If
    
    If historico_doc <> "" Then
        strSQL = strSQL & " and historico_doc like  '%" & historico_doc & "%' "
    End If
    If prejuizo = "" Then
        strSQL = strSQL & " and prejuizo_itd='N'"
    Else
     
        strSQL = strSQL & " and prejuizo_itd='" & prejuizo & "' "
    End If
    
    If boleto <> "" Then
        strSQL = strSQL & " and ti.boleto_itd like  '" & boleto & "%' "
    End If
    If seq_doc <> "" Then
        If Len(seq_doc) < 8 Then
            strSQL = strSQL & " and td.seq_doc ='" & seq_doc & "' "
        
        Else
            strSQL = strSQL & " and td.seq_doc in" & seq_doc
        End If
    End If
    If seq_itd <> "" Then
        strSQL = strSQL & " and ti.seq_itd ='" & seq_itd & "' "
    End If
    If esp_doc <> "" Then
        strSQL = strSQL & " and td.esp_doc='" & esp_doc & "' "
    End If
    If bloq = "1" Then
    
        strSQL = strSQL & " and status_cli=2 "
    ElseIf bloq = "2" Then
        strSQL = strSQL & " and status_cli=1 "
    End If
    If status1 <> "" Then
        strSQL = strSQL & " and status_cli='" & status & "'"
    End If
    If creditosuspenso <> "" Then
        strSQL = strSQL & " and creditosuspenso_cli='" & creditosuspenso & "'"
    End If
    If listaDoc <> "" Then
        
        strSQL = strSQL & " and num_doc in(" & listaDoc & ")"
    
    End If
    
    
    
    If sOrderBy <> "" Then
        strSQL = strSQL & " order by " & sOrderBy
'       strSQL = strSQL & " order by td.seq_doc,ti.seq_itd "
    End If
    F_SelectDocsReceber = strSQL
End Function
Function F_SelectDocsReceberGruposOrgaos(tipo_ter As String, Optional dataI As String = "", _
    Optional dataF As String = "", Optional boleto As String = "", _
    Optional tipo_cob As String = "", Optional cod_cob As String = "", _
    Optional num_doc As String = "", Optional historico_doc As String = "", Optional seq_itd = "", _
    Optional seq_doc = "", Optional cod_gru As String = "", Optional cod_org As String = "", _
    Optional cod_pla As String = "", Optional cod_ter As String = "", _
    Optional status As Integer = 1, Optional sOrderBy As String = "ti.venc_itd,td.cod_ter", _
    Optional prejuizo As String = "N", Optional tipoData As tipo_vencimento, _
    Optional selectCampoEspecifico As String = "", Optional listaDoc As String = "", Optional cod_cid As String = "", Optional reman As String = "", Optional bloq As String = "") As String
    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim sCampoData As String
    If bloq = "0" Then bloq = ""
     If tipoData = data_documento Then
        sCampoData = "td.data_doc"
    ElseIf tipoData = vencimentoInicial Then
        sCampoData = "ti.vencini_itd"
    Else
        sCampoData = "ti.venc_itd"
    End If


    If tipo_ter = "C" Then
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "Cpf_cli"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If

    If selectCampoEspecifico = "" Then
        strSQL = "select distinct impressao_itd,nome_set,nome_cid,"

        strSQL = strSQL & "tipo_ter,tog.nome_org,tg.nome_gru,tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
        strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
        strSQL = strSQL & " tcc.nome_cob,td.cod_ter,tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "

    Else
        strSQL = selectCampoEspecifico
    End If

        strSQL = strSQL & " from " & sTab & " as tc,tab_setores as tset,  tab_cidades as tcid , "

    strSQL = strSQL & "tab_doc as td,tab_itdoc as ti,tab_parametros as tp, "
    strSQL = strSQL & "tab_tipocob as tt,tab_contacob as tcc,tab_planocontas as tpc,tab_orgaos as tog,tab_grupocli as tg "

    strSQL = strSQL & " where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and tc.cod_set=tset.cod_set "
    strSQL = strSQL & " and tset.cod_cid=tcid.cod_cid  "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and tc.cod_org=tog.cod_org "
    strSQL = strSQL & " and tog.cod_gru=tg.cod_gru "
    strSQL = strSQL & " and tt.cod_tip=ti.cod_tip "
    strSQL = strSQL & " and td.cod_pla=tpc.cod_pla and tpc.tipo_pla='C' "
    strSQL = strSQL & " and tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " and td.tipo_ter='" & tipo_ter & "' "
    strSQL = strSQL & " and tpc.cod_pla = td.cod_pla"
    If cod_cid <> "" Then
        strSQL = strSQL & " and tcid.cod_cid='" & cod_cid & "' "
    End If

        If bloq = "1" Then

        strSQL = strSQL & " and status_cli=2 "
    ElseIf bloq = "2" Then
        strSQL = strSQL & " and status_cli=1 "
    End If

    If prejuizo = "" Then
        strSQL = strSQL & " and prejuizo_itd='N'"
    Else

        strSQL = strSQL & " and prejuizo_itd='" & prejuizo & "' "
    End If
    If cod_gru <> "" Then
        strSQL = strSQL & " and tg.cod_gru='" & cod_gru & "'  "
    End If
    If cod_org <> "" Then
        strSQL = strSQL & " and tog.cod_org='" & cod_org & "'  "
    End If
    If status = 1 Then
        strSQL = strSQL & " and ti.reman_itd > 0 "
    ElseIf status = 0 Then
        'strSQL = strSQL & " and ti.reman_itd < 0 "
    End If

    If reman <> "" Then
        strSQL = strSQL & " and ti.reman_itd=" & F_GravaDec(reman)
    End If
    If tipo_cob <> "" Then
        strSQL = strSQL & " and tt.cod_tip='" & tipo_cob & "' "
    End If
    If cod_cob <> "" Then
        strSQL = strSQL & " and cod_cob='" & cod_cob & "' "
    End If

    If dataF <> "" Then
        strSQL = strSQL & " and " & sCampoData & " <= " & F_GravaData(dataF)
    End If
    If dataI <> "" Then
        strSQL = strSQL & " and " & sCampoData & " >= " & F_GravaData(dataI)
    End If
    If cod_ter <> "" Then
        strSQL = strSQL & " and tc." & sCod & "  =  '" & cod_ter & "' "
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and td.num_doc like  '" & num_doc & "%' "
    End If

    If historico_doc <> "" Then
        strSQL = strSQL & " and historico_doc like  '%" & historico_doc & "%' "
    End If

    If listaDoc <> "" Then
        strSQL = strSQL & " and num_doc in(" & listaDoc & ") "
    End If
    If boleto <> "" Then
        strSQL = strSQL & " and tb.seq_itd=ti.seq_itd and tb.num_bol like  '" & boleto & "%' "
    End If
    If seq_doc <> "" Then
        If Len(seq_doc) < 8 Then
            strSQL = strSQL & " and td.seq_doc ='" & seq_doc & "' "

        Else
            strSQL = strSQL & " and td.seq_doc in" & seq_doc
        End If
    End If
    If seq_itd <> "" Then
        strSQL = strSQL & " and td.seq_itd ='" & seq_itd & "' "
    End If
    If sOrderBy <> "" Then
        strSQL = strSQL & " order by " & sOrderBy
    End If
    F_SelectDocsReceberGruposOrgaos = strSQL






End Function

Function F_SelectDocsReceberAcerto(tipo_ter As String, Optional dataI As String = "", _
    Optional dataF As String = "", Optional boleto As String = "", _
    Optional tipo_cob As String = "", Optional cod_cob As String = "", _
    Optional num_doc As String = "", Optional historico_doc As String = "", Optional seq_itd = "", _
    Optional seq_doc = "", Optional cod_gru As String = "", _
    Optional cod_pla As String = "", Optional cod_ter As String = "", _
    Optional status As Integer = 1, Optional seq_ace As String = "", Optional cod_Rep As String = "", _
    Optional sOrderBy As String = "ti.venc_itd,td.cod_ter", Optional prejuizo As String = "N", _
        Optional tipoData As tipo_vencimento, Optional selectCampoEspecifico As String = "", _
            Optional listaDoc As String = "", Optional cod_cid As String = "", Optional reman As String = "", Optional bloq As String = "") As String

    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim sCampoData As String

    If bloq = "0" Then bloq = ""
     If tipoData = data_documento Then
        sCampoData = "td.data_doc"
    ElseIf tipoData = vencimentoInicial Then
        sCampoData = "ti.vencini_itd"
    Else
        sCampoData = "ti.venc_itd"
    End If


    If tipo_ter = "C" Then
        sTab = "tab_clientes"
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "Cpf_cli"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = "tab_funcionarios"
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If
    If selectCampoEspecifico = "" Then
        strSQL = "select  distinct impressao_itd,nome_cid,nome_set,"


        strSQL = strSQL & "diferenca_ace as dif, td.tipo_ter,tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
        strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
        strSQL = strSQL & " tcc.nome_cob,td.cod_ter,tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
    Else
        strSQL = selectCampoEspecifico
    End If


        strSQL = strSQL & " from " & sTab & " as tc, tab_setores as tset,tab_cidades as tcid ,"


    strSQL = strSQL & "tab_doc as td, tab_itdoc as ti,tab_parametros as tp, "
    strSQL = strSQL & "tab_tipocob as tt,tab_contacob as tcc,tab_planocontas as tpc,tab_acertos as ta "

    strSQL = strSQL & " where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and tc.cod_set=tset.cod_set"
    strSQL = strSQL & " and tset.cod_cid=tcid.cod_cid "
    strSQL = strSQL & " and tt.cod_tip=ti.cod_tip "
    strSQL = strSQL & " and td.cod_pla=tpc.cod_pla and tpc.tipo_pla='C' "
    strSQL = strSQL & " and tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " and td.tipo_ter='" & tipo_ter & "' "
    strSQL = strSQL & " and tpc.cod_pla = td.cod_pla "
    strSQL = strSQL & " and tipo_pla='C' "
    strSQL = strSQL & " and ta.seq_ace=td.num_ven "
    strSQL = strSQL & " and td.esp_doc='A' "
    If bloq = "1" Then

        strSQL = strSQL & " and status_cli=2 "
    ElseIf bloq = "2" Then
        strSQL = strSQL & " and status_cli=1 "
    End If

    If cod_cid <> "" Then
        strSQL = strSQL & " and tcid.cod_cid='" & cod_cid & "' "
    End If
    If cod_Rep <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & cod_Rep & "' "
    End If
    If status = 1 Then
        strSQL = strSQL & " and ti.reman_itd > 0 "
    ElseIf status = 0 Then
        strSQL = strSQL & " and ti.reman_itd < 0 "
    End If

    If reman <> "" Then
        strSQL = strSQL & " and ti.reman_itd=" & F_GravaDec(reman)
    End If
    If tipo_cob <> "" Then
        strSQL = strSQL & " and tt.cod_tip='" & tipo_cob & "' "
    End If
    If cod_cob <> "" Then
        strSQL = strSQL & " and ti.cod_cob='" & cod_cob & "' "
    End If

    If dataF <> "" Then
        strSQL = strSQL & " and " & sCampoData & " <= " & F_GravaData(dataF)
    End If
    If dataI <> "" Then
        strSQL = strSQL & " and " & sCampoData & " >= " & F_GravaData(dataI)
    End If
    If cod_ter <> "" Then
        strSQL = strSQL & " and tc." & sCod & "  =  '" & cod_ter & "' "
    End If
    If listaDoc <> "" Then
        strSQL = strSQL & " and num_doc in(" & listaDoc & ")"
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and td.num_doc like  '" & num_doc & "%' "
    End If

    If historico_doc <> "" Then
        strSQL = strSQL & " and historico_doc like  '%" & historico_doc & "%' "
    End If


    If boleto <> "" Then
        strSQL = strSQL & " and tb.seq_itd=ti.seq_itd and tb.num_bol like  '" & boleto & "%' "
    End If
    If seq_doc <> "" Then
        If Len(seq_doc) < 8 Then
            strSQL = strSQL & " and td.seq_doc ='" & seq_doc & "' "

        Else
            strSQL = strSQL & " and td.seq_doc in" & seq_doc
        End If
    End If
    If seq_itd <> "" Then
        strSQL = strSQL & " and td.seq_itd ='" & seq_itd & "' "
    End If
    If sOrderBy <> "" Then
        strSQL = strSQL & " order by " & sOrderBy
    End If
    F_SelectDocsReceberAcerto = strSQL
End Function

Function F_SelectDocsReceberPorReprDeAcerto(Optional dataI As String = "", _
    Optional dataF As String = "", Optional cod_Rep As String = "") As String
    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim boleto As String
    Dim sCampoData As String
    
    sCampoData = "ti.venc_itd"
    sTab = "tab_clientes"
    sCod = "cod_cli"
    sNome = "Nome_cli"
    sCPF = "Cpf_cli"


    
        strSQL = "select distinct tcid.nome_cid,impressao_itd,nome_cli as Cliente,ta.seq_ace as acerto,ti.Valor_itd as valor, "
        strSQL = strSQL & " tipo_ter,ti.Reman_Itd as restante,"
        strSQL = strSQL & "(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
        strSQL = strSQL & "tc." & sCod & " as codigo, tc." & sNome & " as nome,td.historico_doc as Historico, tt.nome_tip as TipoCob,"
        strSQL = strSQL & " tcc.nome_cob as ContaCob,tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido "
    
    strSQL = strSQL & " from " & sTab & " as tc left join tab_setores as tst on tc.cod_set=tst.cod_set"
    strSQL = strSQL & " left join tab_cidades as tcid on tcid.cod_cid=tst.cod_cid, tab_doc as td, "
    strSQL = strSQL & "tab_itdoc as ti,tab_parametros as tp, tab_acertos as ta,"
    strSQL = strSQL & "tab_tipocob as tt,tab_contacob as tcc,tab_planocontas as tpc "

    strSQL = strSQL & "where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and esp_doc='A' "
    strSQL = strSQL & " and ta.seq_ace=td.num_ven "
    strSQL = strSQL & " and tt.cod_tip=ti.cod_tip "
    strSQL = strSQL & " and td.cod_pla=tpc.cod_pla and tpc.tipo_pla='C' "
    strSQL = strSQL & " and tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " and td.tipo_ter='C' "
    strSQL = strSQL & " and tpc.cod_pla = td.cod_pla"
    
    
        strSQL = strSQL & " and prejuizo_itd='N'"
    
    If cod_Rep <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & cod_Rep & "' "
    End If
    strSQL = strSQL & " and ti.reman_itd > 0 "
    
    
    
    If dataF <> "" Then
        strSQL = strSQL & " and " & sCampoData & " <= " & F_GravaData(dataF)
    End If
    If dataI <> "" Then
        strSQL = strSQL & " and " & sCampoData & " >= " & F_GravaData(dataI)
    End If
   
    

    

    
    
        strSQL = strSQL & " order by nome_cli,venc_itd"
    
    F_SelectDocsReceberPorReprDeAcerto = strSQL
End Function

Sub P_VerificaBaixaComissProv(seq_itd As String, valor_bai As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cFator As Currency
    On Error GoTo trata_Erro:
    strSQL = "select distinct valor_doc,seq_com,valor_itd,valor_com from tab_itdoc as ti, tab_comissprov as tcp, tab_doc as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc and ti.seq_itd='" & seq_itd & "' and td.seq_doc=tcp.seq_doc "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            cFator = CCur(valor_bai) / !valor_doc
            strSQL = "insert into tab_comisscons(seq_per,seq_com,gerado_com,valor_bai,seq_itd,data_com,status_com)values("
            strSQL = strSQL & "'',"
            strSQL = strSQL & "'" & !seq_com & "',"
            strSQL = strSQL & F_GravaDec(cFator * CCur(!gerado_com)) & ","
            strSQL = strSQL & F_GravaDec(valor_bai) & ","
            strSQL = strSQL & "'" & seq_itd & "',"
            strSQL = strSQL & F_GravaData(Date) & ","
            strSQL = strSQL & "'1') "
            
            Conn.Execute strSQL
            
            
        End If
                
                
                
                
        .Close
    End With
trata_Erro:
    
End Sub
Sub P_BaixaComissPerm(seq_per As String, seq_com As String, gerado_com As String, seq_fin As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_comisscons set status_com='2',"
    strSQL = strSQL & "seq_fin='" & seq_fin & "'  "
    strSQL = strSQL & " where seq_per=" & seq_per
    Conn.Execute strSQL
    
    strSQL = "update tab_comissprov set reman_com=reman_com-" & F_GravaDec(gerado_com)
    strSQL = strSQL & " where seq_com='" & seq_com & "'"
    Conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_GravaItDocPre(seq_itd As String, anterior_pre As String, atual_pre As String, anotacao_pre As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro:
    
    strSQL = "insert into tab_itdocpre(seq_itd,data_pre,hora_pre,anterior_pre,atual_pre,cod_usu,anotacao_pre)values("
    
    
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Time & "',"
    strSQL = strSQL & F_GravaData(anterior_pre) & ","
    strSQL = strSQL & F_GravaData(atual_pre) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & anotacao_pre & "')"
    Conn.Execute strSQL
    
    strSQL = "update tab_itdoc set venc_itd=" & F_GravaData(atual_pre)
    strSQL = strSQL & " where seq_itd='" & seq_itd & "'"
    Conn.Execute strSQL
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_AlteraDocsAnteriores(seq_doc As String, vencimento As String, novo_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select seq_itd from tab_itdoc where reman_itd>0 "
    strSQL = strSQL & " and venc_itd<=" & F_GravaData(vencimento)
    strSQL = strSQL & " and seq_doc='" & seq_doc & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_itdoc set venc_itd=" & F_GravaData(novo_ven)
            strSQL = strSQL & " where seq_itd='" & !seq_itd & "' "
            Conn.Execute strSQL
            
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Sub P_AlteraDocsAnterioresVariosDocs(SEQ_DOCS As String, vencimento As String, novo_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim seq As String
    Dim c As New Collection
    Dim i As Integer
    Set c = F_SeparaTexto(SEQ_DOCS, " ,;-.", False)
    seq = ""
    For i = 1 To c.Count
        seq = seq & "'" & c(i) & "'"
        If i <> c.Count Then
            seq = seq & ", "
        End If
    Next
    
    
    
    strSQL = "select seq_itd from tab_itdoc where reman_itd>0 "
    strSQL = strSQL & " and venc_itd<=" & F_GravaData(vencimento)
    strSQL = strSQL & " and seq_doc in ( " & seq & ")"
    abreObjetoAdo rPesq, strSQL
    
    
    
    With rPesq
        Do While Not .EOF
            strSQL = "update tab_itdoc set venc_itd=" & F_GravaData(novo_ven)
            strSQL = strSQL & " where seq_itd='" & !seq_itd & "' "
            Conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
End Sub



Sub P_GravaDocImperss(seq_doc As String)
    Dim strSQL As String
    strSQL = "insert into tab_docimpress(seq_imp,seq_doc,data_imp,cod_usu)values("
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "') "
    Conn.Execute strSQL
End Sub
Sub P_gravaDocHist(seq_doc As String, cobout_doc As String, cobin_doc As String, _
    tipocobin_doc As String, tipocobout_doc As String)
    Dim strSQL As String
    strSQL = "insert into " & TAB_DOCHIST & "(seq_doc,cobout_doc,cobin_doc,"
    strSQL = strSQL & "tipocobin_doc,tipocobout_doc,"
    strSQL = strSQL & " data_doc,cod_usu)values("
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & "'" & cobout_doc & "',"
    strSQL = strSQL & "'" & cobin_doc & "',"
    strSQL = strSQL & "'" & tipocobin_doc & "',"
    strSQL = strSQL & "'" & tipocobout_doc & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    Conn.Execute strSQL
End Sub
Function F_SelectDocHist(Optional seq_doc As String = "", Optional cobout As String = "", _
    Optional cobin As String = "", Optional id_doc As String) As String
    Dim strSQL As String
    
    strSQL = "select distinct tcb.nome_cob as cobout,tc.nome_cob as cobin,th.*,"
    strSQL = strSQL & " tf.guerra_fun from tab_dochist as th, tab_contacob as tcb,"
    strSQL = strSQL & " tab_contacob as tc, "
    strSQL = strSQL & " tab_funcionarios as tf where tf.cod_fun=th.cod_usu "
    strSQL = strSQL & " and tcb.cod_cob=cobout_doc"
    strSQL = strSQL & " and tc.cod_cob=cobin_doc"
    If seq_doc <> "" Then
        strSQL = strSQL & " and seq_doc='" & seq_doc & "' "
    End If
'    If id_doc <> "" Then
'        strSQL = strSQL & " and id_doc='" & seq_Doc & "' "
'    End If
    If cobin <> "" Then
        strSQL = strSQL & " and cobin_doc='" & cobin & "' "
    End If
    If cobout <> "" Then
        strSQL = strSQL & " and cobout_doc='" & cobout & "' "
    End If
    F_SelectDocHist = strSQL
End Function
Function F_SelectDocHistUsu(Optional seq_doc As String = "", Optional usuin As String = "") As String
    Dim strSQL As String
    
    strSQL = "select distinct tf.guerra_fun,tf2.guerra_fun as func,th.* from tab_dochistusu as th, "
    strSQL = strSQL & " tab_funcionarios as tf,tab_funcionarios as tf2 "
    strSQL = strSQL & " where tf.cod_fun=th.usuin_doc"
    strSQL = strSQL & " and tf2.cod_fun=th.cod_usu"
    
    
    If seq_doc <> "" Then
        strSQL = strSQL & " and seq_doc='" & seq_doc & "' "
    End If
    If usuin <> "" Then
        strSQL = strSQL & " and usuin_doc='" & usuin & "' "
    End If
    
    F_SelectDocHistUsu = strSQL
End Function
Sub P_gravaDocHistUsu(seq_doc As String, usuin_doc As String)
    Dim strSQL As String
    strSQL = "insert into tab_dochistusu(seq_his,seq_doc,usuin_doc,data_doc,cod_usu)values("
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & seq_doc & "',"
    
    strSQL = strSQL & "'" & usuin_doc & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    Conn.Execute strSQL
End Sub




Function F_buscaItDoc(Optional seq_itd As String = "", Optional seq_doc As String = "", _
    Optional cod_cli As String = "", Optional nomecli As String = "") As itDoc
    frmEditaMov.tag = "X"
    frmEditaMov.txtbxSeqDoc.Text = seq_doc
    frmEditaMov.txtbxSeqItd.Text = seq_itd
    frmEditaMov.txtbxTerceiro.Text = nomecli
    frmEditaMov.txtbxTerceiro.tag = cod_cli
    frmEditaMov.Show 1
    'Do While frmEditaMov.Visible = True
    '    DoEvents
    'Loop
    If frmEditaMov.tag <> "" Then
        With frmEditaMov.ado
            F_buscaItDoc.doc.num_doc = !num_doc
            F_buscaItDoc.reman_itd = !reman_itd
            F_buscaItDoc.seq_doc = !seq_doc
            F_buscaItDoc.seq_itd = !seq_itd
            F_buscaItDoc.valor_itd = !valor_itd
            F_buscaItDoc.venc_itd = !venc_itd
            F_buscaItDoc.vencini_itd = !vencini_itd
        End With
    End If
    Unload frmEditaMov
End Function


Function TotalDebitos(cod_ter As String, tipo_ter As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    On Error GoTo trata_Erro
    strSQL = "select sum(reman_itd) as total from tab_doc as td, tab_itdoc as ti "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and cod_ter='" & cod_ter & "' "
    strSQL = strSQL & " and tipo_ter='" & tipo_ter & "' "
    strSQL = strSQL & " and venc_itd<" & F_GravaData(Date)
    
    
    abreObjetoAdo rPesq, strSQL
    
    TotalDebitos = FormatVal(rPesq!total)
    Exit Function
trata_Erro:
    TotalDebitos = "0,00"
End Function
Function BuscaDocJson(seq_doc As String) As String
    Dim strSQL As String
    strSQL = "SELECT td.seq_doc,td.num_doc,'01' AS cod_emp,"
    strSQL = strSQL & " ti.*"
    strSQL = strSQL & " FROM tab_doc as td inner join tab_itdoc as ti"
    strSQL = strSQL & " ON td.seq_doc=ti.seq_Doc"
    strSQL = strSQL & " WHERE td.seq_Doc='" & seq_doc & "' "


    
    BuscaDocJson = RecordSetToJson(strSQL, , , , , True, , False)
    

End Function
Function BuscaItdocJson(seq_doc As String) As String
    Dim strSQL As String
    strSQL = "select distinct ti.*, "
    strSQL = strSQL & " if(td.tipo_ter='C',tc.cpf_cli ,tr.cpf_rep )as cpf_ter "
    strSQL = strSQL & " from " & TAB_ITDOC & " as ti inner join " & TAB_DOC & " as td on td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " left join " & TAB_CLIENTES & " as tc "
    strSQL = strSQL & " on tc.cod_cli = td.cod_ter "
    strSQL = strSQL & " left join tab_representantes tr "
    strSQL = strSQL & " on tr.cod_rep= td.cod_ter "
    strSQL = strSQL & " WHERE td.seq_Doc='" & seq_doc & "' "


    
    BuscaItdocJson = RecordSetToJson(strSQL, , , , , True, , False)
    

End Function


