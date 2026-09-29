Attribute VB_Name = "Doc"

Option Explicit
Dim strSQL As String
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
End Type



Type itDoc
    seq_doc As String
    seq_itd As String
    doc As doc
    venc_itd As String
    vencini_itd As String
    reman_itd As Currency
    valor_itd As Currency
    
    
End Type




Function F_GravaDoc(hist_doc As String, num_doc As String, valor_doc As String, data_doc As String, _
            cod_ter As String, tipo_doc As String, num_ven As String, _
            parcelas_doc As String, esp_doc As String, tipo_ter As String, cod_pla As String, Optional seq_doc As String = "") As String
    On Error GoTo trata_Erro
    Dim strSQL As String
    If data_doc = "" Then
        data_doc = Date
    End If
    
    If tipo_doc = "" Then
        tipo_doc = "F" 'financeiro
    End If
    
    
    
    strSQL = "insert into " & TAB_DOC & "(historico_doc, Num_Doc,Valor_doc,Data_Doc,cod_ter, "
    strSQL = strSQL & "tipo_doc,num_ven,parcelas_doc,seq_doc,cod_usu,esp_doc,tipo_ter,cod_pla)Values("
    strSQL = strSQL & "'" & hist_doc & "',"
    strSQL = strSQL & "'" & num_doc & "',"
    strSQL = strSQL & F_GravaDec(valor_doc) & ","
    strSQL = strSQL & F_GravaData(data_doc) & ","
    strSQL = strSQL & "'" & cod_ter & "',"
    strSQL = strSQL & "'" & tipo_doc & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & parcelas_doc & ","
    If seq_doc = "" Then
        F_GravaDoc = ChaveSQL(TAB_DOC, "seq_doc", "000000#")
    Else
        F_GravaDoc = seq_doc
    End If
    strSQL = strSQL & "'" & F_GravaDoc & "', "
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'" & esp_doc & "', "
    strSQL = strSQL & "'" & tipo_ter & "'," ' 'C=Cliente,
    strSQL = strSQL & "'" & cod_pla & "') "

    conn.Execute strSQL
    Exit Function
trata_Erro:
    F_GravaDoc = ""
    
End Function

Function F_GravaItDoc(seq_doc As String, valor_itd As String, venc_itd As String, _
    vencini_itd As String, parcela_itd As String, operacao_itd As operacao_itd, _
    Optional boleto As String = "", Optional cod_tip As String = "1", Optional cod_cob As String = "000", _
    Optional seq_itd As String = "", Optional reman_itd As String = "") As String
            
    Dim strSQL As String
    On Error GoTo trata_Erro
    If cod_cob = "" Then
        cod_cob = "000"
    End If
    Dim cod_pla As String
    If cod_pla = "000" Or cod_pla = "" Then
        If operacao_itd = a_receber_avulso Then
            cod_pla = "10001"
        Else
            cod_pla = "20001"
        End If
    End If
    'operacao_itd [V=venda,R=doc a receber avulso,P=a pagar avulso,C= Acerto por conta do Consignado, L - A receber por conta da loja
    
    strSQL = "insert into " & TAB_ITDOC & "(Seq_Doc,Valor_Itd,Venc_Itd,vencini_itd,Reman_Itd,Seq_Itd,parcela_itd,"
    strSQL = strSQL & "operacao_itd,boleto_itd,cod_tip,cod_cob)Values( "
    
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & F_GravaDec(valor_itd) & ","
    strSQL = strSQL & F_GravaData(venc_itd) & ","
    strSQL = strSQL & F_GravaData(vencini_itd) & ","
    If reman_itd = "" Then
        strSQL = strSQL & F_GravaDec(valor_itd) & ","
    Else
        strSQL = strSQL & F_GravaDec(reman_itd) & ","
    End If
    If seq_itd = "" Then
        F_GravaItDoc = MaxKeySQL(TAB_ITDOC, "seq_itd", "", "000000#")
    Else
        F_GravaItDoc = seq_itd
    End If
    strSQL = strSQL & "'" & F_GravaItDoc & "', "
    strSQL = strSQL & parcela_itd & ","
    strSQL = strSQL & "'" & operacao_itd & "', "
    strSQL = strSQL & "'" & boleto & "', "
    strSQL = strSQL & "'" & cod_tip & "'," 'TIPO_COB
    strSQL = strSQL & "'" & cod_cob & "') "
    
    conn.Execute strSQL
    
    P_gravaDocHist seq_doc, "", cod_cob, cod_tip, cod_tip
    
    
    Exit Function
trata_Erro:
    F_GravaItDoc = ""
    'MsgBox Err.Description, vbCritical
    
End Function

Sub P_CancelaBaixaiItd(seq_cab As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sMsg As String
    strSQL = "select * from " & TAB_BAIXAITD & " where seq_cab='" & seq_cab & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        strSQL = ""
        Do While Not .EOF
        sMsg = sMsg & "Seq = " & !seq_itd & " Data: " & !data_bai & " Valor Baixado: " & FormatVal(!valor_bai) & Chr(13)
        strSQL = "update  " & TAB_ITDOC & "  set reman_itd=reman_itd+ " & F_GravaDec(!valor_bai)
        strSQL = strSQL & " where seq_itd='" & !seq_itd & "'"
        conn.Execute strSQL
        
        strSQL = F_GravaCai("D", Date, FormatVal(!valor_bai), "CANCELAMENTO ID. IT. " & !seq_itd, "", "MOV", strCod_Cai, "0", "0", strCod_Cai, strCod_Tur, "000", "N")
        
            .MoveNext
        Loop
        
        .Close
    End With
    MsgBox sMsg & Chr(13) & "Parcelas Canceladas!", vbExclamation
    
    strSQL = "delete from " & TAB_BAIXAITD & " where seq_cab='" & seq_cab & "'"
    conn.Execute strSQL
    strSQL = "delete from " & TAB_BAIXAFIN & " where seq_cab='" & seq_cab & "'"
    conn.Execute strSQL
    
    
       
    
    
End Sub

Function F_GravaBaixaItd(seq_cab As String, seq_itd As String, valor_bai As String, _
    esp_bai As String, seq_fin As String, cod_cob As String, Optional data_bai As String = "", Optional desconto_bai As String = "0", Optional cod_cai As String = "000") As String
    Dim sSeq As String
    If Left(esp_bai, 1) = "1" Then
        esp_bai = "DIN"
    ElseIf Left(esp_bai, 1) = "2" Then
        esp_bai = "CHP"
    Else
        'Esp = "DIN"
    End If
    On Error GoTo trata_Erro
    If data_bai = "" Then
        data_bai = Date
    End If
    Dim strSQL As String
    sSeq = MaxKeySQL("baixaitd", "seq_itd", "seq_cab='" & seq_cab & "'", "")
    
    strSQL = "insert into baixaitd(seq_cab,seq_itd,seq_bai,valor_bai,esp_bai,data_bai,hora_bai,seq_fin,cod_usu,cod_cob,desconto_bai,cod_cai)values("
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
    conn.Execute strSQL
    F_GravaBaixaItd = sSeq
trata_Erro:
End Function
Function F_GravaBaixaItdFin(seq_cab As String, seq_fin As String, esp_bai As String, valor As String, Optional desconto As String = "0") As String
    Dim sSeq As String
    'On Error goto trata_erro
    sSeq = MaxKeySQL("baixafin", "seq_bai", "seq_cab='" & seq_cab & "'", "")
    strSQL = "insert into baixafin(seq_cab,seq_fin,seq_bai,esp_bai,valor_bai,desconto_bai)values("
    strSQL = strSQL & "'" & seq_cab & "',"
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & esp_bai & "',"
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaDec(desconto) & ")"
    conn.Execute strSQL
    F_GravaBaixaItdFin = sSeq
trata_Erro:
End Function



Function F_SelectDocsReceber(tipo_ter As String, Optional dataI As String = "", _
    Optional dataF As String = "", Optional boleto As String = "", _
    Optional tipo_cob As String = "", Optional cod_cob As String = "", _
    Optional num_doc As String = "", Optional historico_doc As String = "", Optional seq_itd = "", _
    Optional seq_doc = "", Optional cod_gru As String = "", _
    Optional cod_pla As String = "", Optional cod_ter As String = "", _
    Optional status As Integer = 1, Optional sOrderBy As String = "ti.venc_itd,td.cod_ter", Optional prejuizo As String = "N", _
    Optional tipodata As tipoVencimentoCheque = Venc_che, Optional selectCampoEspecifico As String = "", _
        Optional esp_doc As String = "", Optional listaDoc As String = "", _
        Optional vinculado As Boolean = False, Optional nNF As String = "", _
        Optional buscaNFe As Boolean = False) As String
    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim sCampoData As String
    Dim sqlNfe As String
    
    If tipodata = data_doc Then
        sCampoData = "td.data_doc"
    Else
        sCampoData = "ti.venc_itd"
    End If
    
    
    If tipo_ter = "C" Then
        sTab = TAB_CLIENTES
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "cnpj_cli"
    Else
        sTab = TAB_FORNECEDORES
        sCod = "cod_for"
        sNome = "Nome_For"
        sCPF = "cnpj_for"
    End If
    
  
    If selectCampoEspecifico = "" Then
        If buscaNFe Then
            strSQL = "select distinct nfe.nnf,nfe.serie,  tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias,tipo_ter, "
        Else
            strSQL = "select distinct nfe.nnf,nfe.serie,  tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias,tipo_ter, "
        End If
        strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
        strSQL = strSQL & " tcc.nome_cob,td.cod_ter,tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
        
    Else
        strSQL = selectCampoEspecifico
    End If
    strSQL = strSQL & " from " & TAB_DOC & " as td inner join " & sTab & " as tc on td.cod_ter=tc." & sCod
    strSQL = strSQL & " left join " & TAB_PLANOCONTAS & " as tpc on tpc.cod_pla=td.cod_pla inner join  "
    
    strSQL = strSQL & TAB_ITDOC & " as ti on td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " left join contacob  as tcc on tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " left join tab_tipocob as tt  on tt.cod_tip=ti.cod_tip "
    strSQL = strSQL & " left join " & TAB_VENDAS & " as tv on tv.num_ven=td.num_doc and esp_doc='V'"
    strSQL = strSQL & " left join nfe  on tv.num_ven=nfe.num_ven "
    strSQL = strSQL & ",tab_parametros as tp  "
    
    
    strSQL = strSQL & " "
    strSQL = strSQL & " where td.tipo_doc='R' "
    
    
    strSQL = strSQL & " and td.tipo_ter='" & tipo_ter & "' "
    'strSQL = strSQL & " and tpc.cod_pla = td.cod_pla"
    'strSQL = strSQL & " and tipo_pla='C' "
    If status = 1 Then
        strSQL = strSQL & " and ti.reman_itd > 0 "
    ElseIf status = 0 Then
        'strSQL = strSQL & " and ti.reman_itd < 0 "
    ElseIf status = -1 Then
        strSQL = strSQL & " and ti.reman_itd < ti.valor_itd "
    End If
    
    If tipo_cob <> "" Then
        strSQL = strSQL & " and tt.cod_tip='" & tipo_cob & "' "
    End If
    If cod_cob <> "" Then
        strSQL = strSQL & " and ti.cod_cob='" & cod_cob & "' "
    End If
    
    If nNF <> "" Then
        strSQL = strSQL & " and nnf ='" & nNF & "'"
    End If
    
    If num_doc <> "" Then
        strSQL = strSQL & " and td.num_doc like  '" & num_doc & "%' "
    Else
        If dataF <> "" Then
            strSQL = strSQL & " and " & sCampoData & " <= " & F_GravaData(dataF)
        End If
        If dataI <> "" Then
            strSQL = strSQL & " and " & sCampoData & " >= " & F_GravaData(dataI)
        End If
        
    End If
    If cod_ter <> "" Then
            strSQL = strSQL & " and tc." & sCod & "  =  '" & cod_ter & "' "
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
        strSQL = strSQL & " and boleto_itd like  '%" & boleto & "%' "
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
    If listaDoc <> "" Then
        
        strSQL = strSQL & " and num_doc in(" & listaDoc & ")"
    
    End If
    
    If vinculado = True Then
        strSQL = strSQL & " and td.seq_doc in (select seq_doc from tab_vinc order by seq_seq )"
    End If
    
    
    If sOrderBy <> "" Then
        strSQL = strSQL & " order by " & sOrderBy
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
    Optional prejuizo As String = "N", Optional tipodata As tipoVencimentoCheque = Venc_che, _
    Optional selectCampoEspecifico As String = "", Optional listaDoc As String = "") As String
    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim sCampoData As String
    
    If tipodata = data_doc Then
        sCampoData = "td.data_doc"
    Else
        sCampoData = "ti.venc_itd"
    End If
    
    If tipo_ter = "C" Then
        sTab = TAB_CLIENTES
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "Cpf_cli"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = TAB_FUNCIONARIOS
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If
    
    If selectCampoEspecifico = "" Then
    
        strSQL = "select distinct impressao_itd,tipo_ter,tog.nome_org,tg.nome_gru,tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
        strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
        strSQL = strSQL & " tcc.nome_cob,td.cod_ter,tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
        If boleto <> "" Then
            strSQL = strSQL & ", tb.* "
        End If
    Else
        strSQL = selectCampoEspecifico
    End If
    strSQL = strSQL & " from " & sTab & " as tc,  " & TAB_DOC & "  as td, "
    strSQL = strSQL & " " & TAB_ITDOC & "  as ti,tab_parametros as tp, "
    strSQL = strSQL & "tab_tipocob as tt,tab_contacob as tcc," & TAB_PLANOCONTAS & " as tpc,tab_orgaos as tog,tab_grupocli as tg "
    If boleto <> "" Then
        strSQL = strSQL & ",tab_boletos as tb "
    End If
    strSQL = strSQL & "where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and tc.cod_org=tog.cod_org "
    strSQL = strSQL & " and tog.cod_gru=tg.cod_gru "
    strSQL = strSQL & " and tt.cod_tip=ti.cod_tip "
    strSQL = strSQL & " and td.cod_pla=tpc.cod_pla and tpc.tipo_pla='C' "
    strSQL = strSQL & " and tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " and td.tipo_ter='" & tipo_ter & "' "
    strSQL = strSQL & " and tpc.cod_pla = td.cod_pla"
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
        strSQL = strSQL & " and ti.reman_itd < 0 "
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
        Optional tipodata As tipoVencimentoCheque = Venc_che, Optional selectCampoEspecifico As String = "", _
            Optional listaDoc As String = "") As String
    
    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim sCampoData As String
    
    If tipodata = data_doc Then
        sCampoData = "td.data_doc"
    Else
        sCampoData = "ti.venc_itd"
    End If
    
    If tipo_ter = "C" Then
        sTab = TAB_CLIENTES
        sCod = "cod_cli"
        sNome = "Nome_cli"
        sCPF = "Cpf_cli"
    ElseIf tipo_ter = "R" Then
        sTab = "tab_representantes"
        sCod = "cod_rep"
        sNome = "Nome_rep"
        sCPF = "Cpf_Rep"
    ElseIf tipo_ter = "U" Then 'funcionário
        sTab = TAB_FUNCIONARIOS
        sCod = "cod_fun"
        sNome = "Nome_fun"
        sCPF = "cpf_fun"
    End If
    If selectCampoEspecifico = "" Then
    
        strSQL = "select distinct impressao_itd,td.tipo_ter,tpc.nome_pla,tpc.cod_pla,(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
        strSQL = strSQL & "tc." & sCod & " as codigo, td.num_ven,tc." & sNome & " as nome,td.esp_doc,td.historico_doc, tt.nome_tip,"
        strSQL = strSQL & " tcc.nome_cob,td.cod_ter,tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "td.valor_Doc,td.parcelas_doc,td.num_doc,td.data_doc, td.valor_doc, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido, ti.* "
    Else
        strSQL = selectCampoEspecifico
    End If
    strSQL = strSQL & " from " & sTab & " as tc,  " & TAB_DOC & "  as td, "
    strSQL = strSQL & " " & TAB_ITDOC & "  as ti,tab_parametros as tp, "
    strSQL = strSQL & "tab_tipocob as tt,tab_contacob as tcc, " & TAB_PLANOCONTAS & " as tpc,tab_acertos as ta "
    
    strSQL = strSQL & "where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and tt.cod_tip=ti.cod_tip "
    strSQL = strSQL & " and td.cod_pla=tpc.cod_pla and tpc.tipo_pla='C' "
    strSQL = strSQL & " and tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " and td.tipo_ter='" & tipo_ter & "' "
    strSQL = strSQL & " and tpc.cod_pla = td.cod_pla "
    strSQL = strSQL & " and tipo_pla='C' "
    strSQL = strSQL & " and ta.seq_ace=td.num_ven "
    strSQL = strSQL & " and td.esp_doc='A' "
    If cod_Rep <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & cod_Rep & "' "
    End If
    If status = 1 Then
        strSQL = strSQL & " and ti.reman_itd > 0 "
    ElseIf status = 0 Then
        strSQL = strSQL & " and ti.reman_itd < 0 "
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
    Optional dataF As String = "", Optional cod_Rep As String = "", _
    Optional tipo_cob As String = "", Optional cod_cob As String = "", _
    Optional num_doc As String = "", Optional historico_doc As String = "", Optional seq_itd = "", _
    Optional seq_doc = "", Optional cod_gru As String = "", _
    Optional cod_pla As String = "", Optional cod_ter As String = "", _
    Optional status As Integer = 1, Optional sOrderBy As String = "ti.venc_itd,td.cod_ter", _
    Optional prejuizo As String = "N", Optional tipodata As tipoVencimentoCheque = Venc_che, _
    Optional selectCampoEspecifico As String = "", Optional listaDoc As String = "") As String
    Dim strSQL As String
    Dim sNome As String
    Dim sCod As String
    Dim sCPF As String
    Dim sTab As String
    Dim boleto As String
    Dim sCampoData As String
    If tipodata = data_doc Then
        sCampoData = "td.data_doc"
    Else
        sCampoData = "ti.venc_itd"
    End If
    sTab = TAB_CLIENTES
    sCod = "cod_cli"
    sNome = "Nome_cli"
    sCPF = "Cpf_cli"
    
    
    
    If selectCampoEspecifico = "" Then
        strSQL = "select distinct impressao_itd,nome_cli as Cliente,ta.seq_ace as acerto,ti.Valor_itd as valor, tipo_ter,ti.Reman_Itd as restante,"
        strSQL = strSQL & "(to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as dias, "
        strSQL = strSQL & "tc." & sCod & " as codigo, tc." & sNome & " as nome,td.historico_doc as Historico, tt.nome_tip as TipoCob,"
        strSQL = strSQL & " tcc.nome_cob as ContaCob,tc." & sCPF & " as cpf,concat(parcela_itd,'/',parcelas_doc) as parcela, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) as valor_juro, "
        strSQL = strSQL & "((tp.juros_par/100*ti.reman_itd)/30) * (to_days(" & F_PesqData(Date) & ")- to_days(venc_itd)) + reman_itd as sugerido "
    Else
        strSQL = selectCampoEspecifico
    End If
    strSQL = strSQL & " from " & sTab & " as tc,  " & TAB_DOC & "  as td, "
    strSQL = strSQL & " " & TAB_ITDOC & "  as ti,tab_parametros as tp, tab_acertos as ta,"
    strSQL = strSQL & "tab_tipocob as tt,tab_contacob as tcc, " & TAB_PLANOCONTAS & " as tpc "
    
    strSQL = strSQL & "where tc." & sCod & "=td.cod_ter and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.tipo_doc='R' "
    strSQL = strSQL & " and esp_doc='A' "
    strSQL = strSQL & " and ta.seq_ace=td.num_ven "
    strSQL = strSQL & " and tt.cod_tip=ti.cod_tip "
    strSQL = strSQL & " and td.cod_pla=tpc.cod_pla and tpc.tipo_pla='C' "
    strSQL = strSQL & " and tcc.cod_cob=ti.cod_cob "
    strSQL = strSQL & " and td.tipo_ter='C' "
    strSQL = strSQL & " and tpc.cod_pla = td.cod_pla"
    If prejuizo = "" Then
        strSQL = strSQL & " and prejuizo_itd='N'"
    Else
     
        strSQL = strSQL & " and prejuizo_itd='" & prejuizo & "' "
    End If
    If cod_Rep <> "" Then
        strSQL = strSQL & " and ta.cod_rep='" & cod_Rep & "' "
    End If
    If status = 1 Then
        strSQL = strSQL & " and ti.reman_itd > 0 "
    ElseIf status = 0 Then
        strSQL = strSQL & " and ti.reman_itd < 0 "
    End If
    If listaDoc <> "" Then
        strSQL = strSQL & " and num_doc in(" & listaDoc & ")"
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
    F_SelectDocsReceberPorReprDeAcerto = strSQL
End Function

Sub P_VerificaBaixaComissProv(seq_itd As String, valor_bai As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cFator As Currency
    On Error GoTo trata_Erro:
    strSQL = "select distinct valor_doc,seq_com,valor_itd,valor_com from  " & TAB_ITDOC & "  as ti, tab_comissprov as tcp,  " & TAB_DOC & "  as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc and ti.seq_itd='" & seq_itd & "' and td.seq_doc=tcp.seq_doc "
    rPesq.Open strSQL, conn
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
            
            conn.Execute strSQL
            
            
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
    conn.Execute strSQL
    
    strSQL = "update tab_comissprov set reman_com=reman_com-" & F_GravaDec(gerado_com)
    strSQL = strSQL & " where seq_com='" & seq_com & "'"
    conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_GravaItDocPre(seq_itd As String, anterior_pre As String, atual_pre As String, anotacao_pre As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro:
    
    strSQL = "insert into  " & TAB_ITDOC & "pre(seq_itd,data_pre,hora_pre,anterior_pre,atual_pre,cod_usu,anotacao_pre)values("
    
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Time & "',"
    strSQL = strSQL & F_GravaData(anterior_pre) & ","
    strSQL = strSQL & F_GravaData(atual_pre) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & anotacao_pre & "')"
    conn.Execute strSQL
    
    strSQL = "update  " & TAB_ITDOC & "  set venc_itd=" & F_GravaData(atual_pre)
    strSQL = strSQL & " where seq_itd='" & seq_itd & "'"
    conn.Execute strSQL
    
    
    
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Sub P_AlteraDocsAnteriores(seq_doc As String, vencimento As String, novo_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select seq_itd from  " & TAB_ITDOC & "  where reman_itd>0 "
    strSQL = strSQL & " and venc_itd<=" & F_GravaData(vencimento)
    strSQL = strSQL & " and seq_doc='" & seq_doc & "' "
    rPesq.Open strSQL, conn
    With rPesq
        Do While Not .EOF
            strSQL = "update  " & TAB_ITDOC & "  set venc_itd=" & F_GravaData(novo_ven)
            strSQL = strSQL & " where seq_itd='" & !seq_itd & "' "
            conn.Execute strSQL
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Sub P_GravaDocImperss(seq_doc As String)
    Dim strSQL As String
    strSQL = "insert into  " & TAB_DOC & " impress(seq_imp,seq_doc,data_imp,cod_usu)values("
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "') "
    conn.Execute strSQL
End Sub
Sub P_gravaDocHist(seq_doc As String, cobout_doc As String, cobin_doc As String, _
    tipocobin_doc As String, tipocobout_doc As String)
    Dim strSQL As String
    strSQL = "insert into tab_dochist(seq_doc,cobout_doc,cobin_doc,"
    strSQL = strSQL & "tipocobin_doc,tipocobout_doc,"
    strSQL = strSQL & " data_doc,cod_usu)values("
    
    strSQL = strSQL & "'" & seq_doc & "',"
    strSQL = strSQL & "'" & cobout_doc & "',"
    strSQL = strSQL & "'" & cobin_doc & "',"
    strSQL = strSQL & "'" & tipocobin_doc & "',"
    strSQL = strSQL & "'" & tipocobout_doc & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    conn.Execute strSQL
End Sub
Function F_SelectDocHist(Optional seq_doc As String = "", Optional cobout As String = "", _
    Optional cobin As String = "") As String
    Dim strSQL As String
    
    strSQL = "select distinct tcb.nome_cob as cobout,tc.nome_cob as cobin,th.*,"
    strSQL = strSQL & " tf.guerra_fun from  TAB_DOChist as th, tab_contacob as tcb,"
    strSQL = strSQL & " tab_contacob as tc, "
    strSQL = strSQL & " tab_funcionarios as tf where tf.cod_fun=th.cod_usu "
    strSQL = strSQL & " and tcb.cod_cob=cobout_doc"
    strSQL = strSQL & " and tc.cod_cob=cobin_doc"
    If seq_doc <> "" Then
        strSQL = strSQL & " and seq_doc='" & seq_doc & "' "
    End If
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
    
    strSQL = "select distinct tf.guerra_fun,tf2.guerra_fun as func,th.* from    TAB_DOChistusu as th, "
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
    strSQL = "insert into   TAB_DOChistusu(seq_his,seq_doc,usuin_doc,data_doc,cod_usu)values("
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & seq_doc & "',"
    
    strSQL = strSQL & "'" & usuin_doc & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "')"
    conn.Execute strSQL
End Sub




Function F_BuscaItDoc(Optional seq_itd As String = "", Optional seq_doc As String = "", _
    Optional cod_cli As String = "", Optional nomecli As String = "") As itDoc
    frmEditaMov.tag = "X"
    frmEditaMov.txtbxSeqDoc.text = seq_doc
    frmEditaMov.txtbxSeqItd.text = seq_itd
    frmEditaMov.txtbxTerceiro.text = nomecli
    frmEditaMov.txtbxTerceiro.tag = cod_cli
    frmEditaMov.Show 1
    'Do While frmEditaMov.Visible = True
    '    DoEvents
    'Loop
    If frmEditaMov.tag <> "" Then
        With frmEditaMov.ado
            F_BuscaItDoc.doc.num_doc = !num_doc
            F_BuscaItDoc.reman_itd = !reman_itd
            F_BuscaItDoc.seq_doc = !seq_doc
            F_BuscaItDoc.seq_itd = !seq_itd
            F_BuscaItDoc.valor_itd = !valor_itd
            F_BuscaItDoc.venc_itd = !venc_itd
            F_BuscaItDoc.vencini_itd = !vencini_itd
        End With
    End If
    Unload frmEditaMov
End Function




Function BuscaListaDoc() As String
    frmListaDocumentos.Show 1
    BuscaListaDoc = frmListaDocumentos.Lista
    
    
End Function

Function F_DadosPar(Optional ven As String = "", _
    Optional valor As String = "0,00", Optional boleto As String = "") As parcela
    
    
    With frmDadosPar
        .txtbxValor_Che.text = valor
        .txtbxVencimento.text = ven
        .txtbxBoleto.text = boleto
        On Error Resume Next
        
        .Show 1
        If .tag <> "" Then
            
            F_DadosPar.valor = .txtbxValor_Che.text
            F_DadosPar.vencimento = .txtbxVencimento.text
            F_DadosPar.boleto = .txtbxBoleto.text
        End If
    End With
        Unload frmDadosPar
        
    
    
End Function





Function F_SelectDocsVinculados(seq_doc As String) As String

    Dim strSQL As String
    strSQL = "select tv.*, td.data_doc, td.historico_doc,td1.historico_doc as hist_original,"
    strSQL = strSQL & " td1.num_doc as doc_original,ti.valor_itd,ti.venc_itd, td.data_doc "
    strSQL = strSQL & " from " & TAB_DOC & " as td inner join tab_vinc as tv on td.seq_doc=tv.seq_doc "
    strSQL = strSQL & " inner join " & TAB_ITDOC & " as ti on tv.seq_itd=ti.seq_itd "
    strSQL = strSQL & " inner join " & TAB_DOC & " as td1 on td1.seq_doc=ti.seq_doc "
    strSQL = strSQL & " where tv.seq_doc='" & seq_doc & "' "
    strSQL = strSQL & " order by ti.venc_itd "
    F_SelectDocsVinculados = strSQL
End Function
Sub P_DocsVinculados(seq_doc As String)
    frmVincListaDocumentos.txtDoc.text = seq_doc
    frmVincListaDocumentos.Show 1

End Sub
Sub EnviarEmailRecibo(seq_cab As String, Optional cod_cli As String = "")
    
    Dim s As String
    Dim anexos(2) As String
    Dim m As String
    Dim p As String
    Dim i As Integer
    Dim email As Mail
    Dim emailCliente As String
    Dim c As cliente
    Dim Arquivo As String
    On Error GoTo trata_Erro
    Arquivo = "Recibo" & seq_cab & ".pdf"
    
    c = F_BuscaClienteCodigo(cod_cli)
    
    emailCliente = c.email_cli
    emailCliente = InputBox("Confirme o e-mail a ser enviado", "Email", emailCliente)
    If emailCliente = "" Then
        Exit Sub
    End If
    email = BuscaMailNome("Doc")
    If email.id = "" Then
        MsgBox "Não foi possível encontrar a configuração de Email.", vbExclamation
        Exit Sub
    End If
    
    ApontarPrinterToPDF Arquivo, , False
    ImprReciboBai seq_cab, False
    Printer.EndDoc
    
    anexos(0) = Arquivo
    SendMailGmail email.sendusername, email.sendpassword, emailCliente, email.subject, email.body, anexos
    
    
    If MsgBox("Email enviado. Deseja visualizar o arquivo PDF Gerado?", vbYesNo + vbQuestion) = vbYes Then
        Shell "explorer " & Arquivo
    End If
Exit Sub
trata_Erro:
    P_MostraErro
    


End Sub

