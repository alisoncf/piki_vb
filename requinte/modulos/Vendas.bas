Attribute VB_Name = "Vendas"
Option Explicit
Dim strSQL As String

Enum tipo_operacao_venda
    venda = 1
    Devolucao = 2
    Remessa = 3
    retorno_remessa = 4
    
End Enum

Function F_GeraNovaVenda(Optional cod_cli As String = "00000", Optional tipo As tipo_operacao_venda = venda) As String
    Dim sNum As String
    Dim sIp As String
    On Error GoTo trata_Erro
    
    'Sleep 100
    
    
    sNum = ChaveSQL("tab_vendas", "num_ven", "000000#")
    
    If F_BuscaUmCampo("num_ven", "tab_Vendas", "num_ven='" & sNum & "'") <> "" Then
        F_GeraNovaVenda = ""
        Exit Function
        
    End If
    
    strSQL = "insert into tab_vendas(num_ven,cod_cli,ip_ven,data_ven,abrevenda_usu,tipo_operacao)values("
    strSQL = strSQL & "'" & sNum & "',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & strIP_Usu & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & tipo & ") "
    Conn.Execute strSQL
    
    F_GeraNovaVenda = sNum
    sIp = F_BuscaUmCampo("ip_ven", "tab_vendas", "num_ven='" & sNum & "'")
    If sIp <> strIP_Usu Then
        MsgBox " A venda " & sNum & " já foi gerada para o IP: " & sIp, vbExclamation
        F_GeraNovaVenda = ""
    End If
    
    
    
    Exit Function
trata_Erro:
    P_MostraErro "Erro na criação da venda, tente novamente!"
End Function
Function F_GeraNovaVendaFuncionario(Optional cod_cli As String = "000") As String
    Dim sNum As String
    Dim sIp As String
    On Error GoTo trata_Erro
    
   
    
    
    sNum = ChaveSQL("tab_vendas", "num_ven", "000000#")
    
    If F_BuscaUmCampo("num_ven", "tab_Vendas", "num_ven='" & sNum & "'") <> "" Then
        F_GeraNovaVendaFuncionario = ""
        Exit Function
        
    End If
    
    strSQL = "insert into tab_vendas(num_ven,cod_cli,data_ven,abrevenda_usu,cod_for,ip_ven,tipocli_ven)values("
    strSQL = strSQL & "'" & sNum & "',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "', "
    strSQL = strSQL & "'06'," ' a prazo 1 x
    strSQL = strSQL & "'" & strIP_Usu & "',"
    strSQL = strSQL & "'U')"
    Conn.Execute strSQL
    
    F_GeraNovaVendaFuncionario = sNum
    
    F_GeraNovaVendaFuncionario = sNum
    sIp = F_BuscaUmCampo("ip_ven", "tab_vendas", "num_ven='" & sNum & "'")
    If sIp <> strIP_Usu Then
        MsgBox " A venda " & sNum & " já foi gerada para o IP: " & sIp, vbExclamation
        F_GeraNovaVendaFuncionario = ""
    End If
    
    
    Exit Function
trata_Erro:
    P_MostraErro "Erro na criação da venda, tente novamente!"
End Function

Function F_MudaQtdEstoque(cod_pro As String, valor, Optional real As Boolean = True) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    F_MudaQtdEstoque = True
    
    strSQL = "update tab_produtos set "
    If real Then
        strSQL = strSQL & " er_pro=er_pro " & F_GravaDec(valor)
    Else
        strSQL = strSQL & " ef_pro=ef_pro " & F_GravaDec(valor)
    End If
    strSQL = strSQL & " where cod_pro='" & cod_pro & "'"
    
    Conn.Execute strSQL
    
    
    Exit Function
trata_Erro:
    F_MudaQtdEstoque = False
End Function
Function F_AdicionaPosologia(num_ven As String, seq_ite As String, descricao_pos As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_posologia(num_ven,seq_ite,descricao_pos)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & seq_ite & ","
    strSQL = strSQL & "'" & descricao_pos & "') "
    Conn.Execute strSQL
    
    
    F_AdicionaPosologia = True
    Exit Function
    
trata_Erro:
    F_AdicionaPosologia = False
End Function
Function F_AlteraPosologia(num_ven As String, seq_ite As String, descricao_pos As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = "UPdate tab_posologia set descricao_pos='" & descricao_pos & "' "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    strSQL = strSQL & " and seq_ite=" & seq_ite
    Conn.Execute strSQL
    F_AlteraPosologia = True
    Exit Function
trata_Erro:
    F_AlteraPosologia = False
End Function
Function F_BuscaItens(num_ven As String, Optional sOrd As String = "seq_ite", Optional sWhere As String = "") As String
    Dim strSQL As String
    strSQL = "select tp.ncm_pro,ts.nome_sec, tu.nome_sub, if(pq_itd>0,'S','N') as pq,"
        
    strSQL = strSQL & "(ti.qtd_ite*ti.unit_ite) as total,desconto_ite as desconto ,final_ite as final, "
    
    strSQL = strSQL & " tp.st_pro,tp.ref_pro,ti.* "
    
    
    
    
    strSQL = strSQL & " from tab_produtos as tp, "
    strSQL = strSQL & " tab_subsec as tu,tab_secoes as ts, "
    strSQL = strSQL & " tab_itvendas as ti "

    
    strSQL = strSQL & " where tp.cod_sub=tu.cod_sub "
    strSQL = strSQL & " and tu.cod_sec=ts.cod_sec "
    
    
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    
    strSQL = strSQL & " and ti.num_ven='" & num_ven & "' "
    If sWhere <> "" Then
        strSQL = strSQL & " and " & sWhere
    End If
    If sOrd <> "" Then
        strSQL = strSQL & " order by " & sOrd
    End If
    F_BuscaItens = strSQL
End Function
Function F_selectVendas(Optional cod_cli As String = "", Optional dataI As String = "", Optional dataF As String = "", _
    Optional situacao_ven As String = "situacao_ven>1", Optional abreVendaUsu As String = "", _
    Optional sOrderBy As String = "num_ven desc", Optional groupBy As String = "") As String
    Dim strSQL As String
    Dim sCampoData As String
    strSQL = "select tfo.nome_for,tc.nome_cli, tv.*, tv.valor_ven+tv.desconto_ven as total, tf.guerra_fun "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " tab_vendas as tv, tab_funcionarios as tf,tab_formaspag as tfo  "
    
    strSQL = strSQL & " where tf.cod_fun=tv.abrevenda_usu "
    
    strSQL = strSQL & " and tfo.cod_for=tv.cod_for "
    If situacao_ven <> "" Then
        strSQL = strSQL & " and " & situacao_ven
    End If
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    If abreVendaUsu <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & abreVendaUsu & "' "
    End If
    
    If cod_cli <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & cod_cli & "' "
    End If
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(dataF & " 23:59:59")
    End If
    If groupBy <> "" Then
        strSQL = strSQL & " group by " & groupBy
    End If
    F_selectVendas = strSQL
End Function

Function F_selectVendasAgrupCli(Optional cod_cli As String = "", Optional dataI As String = "", Optional dataF As String = "", _
    Optional situacao_ven As String = "situacao_ven>4", Optional abreVendaUsu As String = "", _
    Optional formaP As String = "", Optional sOrderBy As String = "nome_cli") As String
    Dim strSQL As String
    Dim sCampoData As String
    strSQL = "select tc.nome_cli, tv.*, sum(tv.valor_ven)as total,"
    strSQL = strSQL & " sum(tv.desconto_ven) as desconto "
    strSQL = strSQL & " from tab_clientes as tc, "
    strSQL = strSQL & " tab_vendas as tv, tab_formaspag as tfo  "
    
    strSQL = strSQL & " where tfo.cod_for=tv.cod_for "
    strSQL = strSQL & " and tv.cod_cli=tc.cod_cli "
    If formaP <> "" Then
    
        strSQL = strSQL & " and tfo.cod_for='" & formaP & "' "
    End If
    If situacao_ven <> "" Then
        strSQL = strSQL & " and " & situacao_ven
    End If
    
    If abreVendaUsu <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & abreVendaUsu & "' "
    End If
    
    If cod_cli <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & cod_cli & "' "
    End If
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(dataF & " 23:59:59")
    End If
    
    strSQL = strSQL & " group by tv.cod_cli"
    
    F_selectVendasAgrupCli = strSQL
End Function


Sub P_RecalculaVendaDesconto(novoValor_ven As String, desconto As String, num_ven As String, autoriza_desconto As String, Optional queima As Boolean = False)
    Dim strSQL As String
    'On Error Resume Next
    Dim rPesq As New ADODB.Recordset
    Dim sTotalVenda As String
    Dim sTotalDesconto As String
    strSQL = "Update tab_itvendas set desconto_ite=desconto_ite+(final_ite/100)*" & F_GravaDec(desconto)
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    If queima = False Then
        strSQL = strSQL & " and pq_itd=0 "
    End If
    Conn.Execute strSQL
    
    strSQL = "update tab_itvendas set final_ite=unit_ite*qtd_ite-desconto_ite "
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    If queima = False Then
        strSQL = strSQL & " and pq_itd=0 "
    End If
    Conn.Execute strSQL
    
    strSQL = "select sum(final_ite) as final, "
    strSQL = strSQL & "sum(desconto_ite) as desconto "
    strSQL = strSQL & "from tab_itvendas where num_ven='" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        strSQL = "update tab_vendas set desconto_ven=" & F_GravaDec(!desconto) & ","
        strSQL = strSQL & "valor_ven=" & F_GravaDec(novoValor_ven) & ", "
        strSQL = strSQL & "autorizadesconto_ven='" & autoriza_desconto & "' "
        strSQL = strSQL & " where num_ven='" & num_ven & "' "
        Conn.Execute strSQL
        .Close
    End With
    
End Sub
Sub P_AtualizaVendaCliente(num_ven As String, cod_cli As String, nomecli_ven As String, _
    Optional telefone As String = "", Optional cpfcli_ven As String = "", Optional email As String, Optional aniversario As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update tab_vendas set cod_cli='" & cod_cli & "' "
    If nomecli_ven <> "" Then
        strSQL = strSQL & ",nomecli_ven='" & nomecli_ven & "'  "
    End If
    If cpfcli_ven <> "" Then
        strSQL = strSQL & ",cpfcli_ven='" & F_SoNumero(cpfcli_ven) & "'  "
    End If

    If telefone <> "" Then
        strSQL = strSQL & ",telefonecli_ven='" & telefone & "' "
    End If
    If email <> "" Then
        strSQL = strSQL & ",email_ven='" & email & "' "
    End If
    If aniversario <> "" Then
        strSQL = strSQL & ",aniversario_ven='" & aniversario & "' "
    End If
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    Conn.Execute strSQL
trata_Erro:
P_MostraErro
End Sub
Sub P_AtualizaVendaValores(num_ven As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cTotal As Currency
    Dim cDesc As Currency
    strSQL = "select final_ite,desconto_ite from tab_itvendas where num_ven='" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            cTotal = cTotal + !final_ite
            cDesc = cDesc + !desconto_ite
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "update tab_vendas set valor_ven=" & F_GravaDec(cTotal) & ","
    strSQL = strSQL & "desconto_ven=" & F_GravaDec(cDesc)
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    Conn.Execute strSQL

End Sub



Sub P_AtualizaFormaPag(num_ven, formaPag As String)
    Dim strSQL As String
    strSQL = "update tab_vendas set cod_For='" & formaPag & "' "
    
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    Conn.Execute strSQL

End Sub

Function F_AlteraStatusVenda(num_ven As String, status As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_vendas set situacao_ven='" & status & "' "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    Conn.Execute strSQL
    
    F_AlteraStatusVenda = True
    
    Exit Function
trata_Erro:
    
End Function

Function FechaRemessaDevolucao(num_ven As String, status As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_vendas set situacao_ven='" & status & "', "
    strSQL = strSQL & "fechavenda_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    Conn.Execute strSQL
    
    FechaRemessaDevolucao = True
    
    Exit Function
trata_Erro:
    
End Function


Function F_BuscaItVendas(num_ven As String) As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select ti.* from tab_itvendas as ti where num_ven='" & num_ven & "' "
    abreObjetoAdo rPesq, strSQL
    strSQL = ""
    With rPesq
        Do While Not .EOF
           strSQL = strSQL & (!cod_pro) & Chr(9)
           'strSQL = strSQL & !desc_ite, 50) & " "
           strSQL = strSQL & F_EsticaTexto(!qtd_ite & "x" & FormatVal(!unit_ite), 50, 2) & " "
           strSQL = strSQL & F_EsticaTexto(FormatVal(!final_ite), 40, 2) & Chr(13) + Chr(10)
            .MoveNext
        Loop
        F_BuscaItVendas = strSQL
        .Close
    End With

End Function
Function F_AdicionaItem(num_ven As String, prod As produto, qtd As String, desconto As String, _
    final As String, Optional pq_itd As String = "0") As String
    Dim sSeq As String
    On Error GoTo trata_Erro
    sSeq = MaxKeySQL("tab_itvendas", "seq_ite", "num_ven='" & num_ven & "'", "")
    strSQL = "insert into tab_itvendas(num_ven,seq_ite,cod_pro,qtd_ite,unit_ite,desc_ite,"
    strSQL = strSQL & " desconto_ite,final_ite,pq_itd)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & prod.cod_pro & "',"
    
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(prod.pv_pro) & ","
    strSQL = strSQL & "'" & prod.nome_pro & "',"
    
    strSQL = strSQL & F_GravaDec(desconto) & ","
    strSQL = strSQL & F_GravaDec(final) & ","
    strSQL = strSQL & F_GravaDec(pq_itd) & ")"
    Conn.Execute strSQL
    F_AdicionaItem = sSeq
    Exit Function
trata_Erro:
End Function

Function F_GravaCrediarioVen(num_ven As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_crediarioven(num_ven,status_cre,data_cre)values('" & num_ven & "',"
    strSQL = strSQL & "'1',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ")"
    Conn.Execute strSQL
    
    
    strSQL = "update tab_vendas set crediario_ven=1"
            
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    Conn.Execute strSQL
    F_GravaCrediarioVen = True
    Exit Function
trata_Erro:
    
End Function
Function F_AlteraCrediarioVen(num_ven As String, status_cre As String, avaliacao_cre As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_crediarioven set status_cre='" & status_cre & "',"
    strSQL = strSQL & "avaliacao_cre='" & avaliacao_cre & "',"
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "',"
    strSQL = strSQL & "fechamento_cre=" & F_GravaDataTempo(Now)
    strSQL = strSQL & " where num_ven='" & num_ven & "'"
    Conn.Execute strSQL
    
    Conn.Execute strSQL
    F_AlteraCrediarioVen = True
    Exit Function
trata_Erro:
    
End Function




Sub P_CancelaDescontoVenda(num_ven As String, Optional desconto As String = "0")
    Dim strSQL As String
    desconto = CCur(desconto / 100)
    strSQL = "update tab_itvendas set desconto_ite=(qtd_ite*unit_ite)*" & F_GravaDec(desconto) & ","
    strSQL = strSQL & "final_ite=(qtd_ite*unit_ite)-((qtd_ite*unit_ite)*" & F_GravaDec(desconto) & "),"
    strSQL = strSQL & "pq_itd=0 "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    Conn.Execute strSQL
    strSQL = "update tab_vendas set autorizadesconto_ven='' where num_ven='" & num_ven & "'"
    Conn.Execute strSQL
    
    P_AtualizaVendaValores num_ven
End Sub
Function F_AdicionaVenFec(num_ven As String, valor_ven As String, _
    cod_For As String, cod_cai As String, cod_tur As String, esp_cai As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "replace into tab_venfec(num_ven,valor_ven,cod_for,cod_cai,cod_tur,esp_cai)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & F_GravaDec(valor_ven) & ","
    strSQL = strSQL & "'" & cod_For & "',"
    strSQL = strSQL & "'" & cod_cai & "',"
    strSQL = strSQL & "'" & cod_tur & "',"
    strSQL = strSQL & "'" & esp_cai & "')"
    Conn.Execute strSQL
    F_AdicionaVenFec = True
Exit Function
trata_Erro:
    
End Function
    


Sub P_AutorizaDescontoGerencial(num_ven As String, desconto As String)
    Dim strSQL As String
    strSQL = "update tab_itvendas set desconto_ite=desconto_ite+(" & F_GravaDec(desconto)
    strSQL = strSQL & ""
End Sub
Sub P_GravaVendaAberta(num_ven As String)
    Dim strSQL As String
    On Error Resume Next
    If bUsaCaixaSeparado = True Then
        strSQL = "insert into tab_vendaaberta(num_ven,ip_ven)values("
        strSQL = strSQL & "'" & num_ven & "',"
        strSQL = strSQL & "'" & strIP_Usu & "') "
        Conn.Execute strSQL
    End If
    
End Sub
Sub P_DeletaVendaAberta(num_ven As String)
    On Error Resume Next
    
    strSQL = "delete from tab_vendaaberta where num_ven='" & num_ven & "'"
    
    Conn.Execute strSQL
    
End Sub
Function F_VerificaVendaAberta(num_ven As String) As String
    Dim strSQL As String
    strSQL = F_BuscaUmCampo("ip_ven", "tab_vendaaberta", "num_ven='" & num_ven & "'")
    If strSQL <> strIP_Usu And strSQL <> strHost_Usu Then
    
        F_VerificaVendaAberta = strSQL
    End If
End Function
Function VerificaVendaAberta(num_ven As String) As String
    VerificaVendaAberta = F_VerificaVendaAberta(num_ven)
End Function
Function F_CancelaVendaAposCupom(sNum_ven As String) As Boolean
    Dim iPos As Integer
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim valorCai As String
    Dim entrada As String
    Dim saida As String
    
    On Error GoTo trata_Erro
    Dim sCodCai As String
    'varre toda a venda voltando o estoque real dos itens depois muda para situacao 5
    F_CancelaVendaAposCupom = False
    
    
    If F_BuscaUmCampo("status_ven", "tab_vendas", "num_ven='" & sNum_ven & "'") = "3" Then
        MsgBox "Não foi possível cancelar os lançamentos da venda, pois ela já se encontra cancelada!", vbExclamation
        F_CancelaVendaAposCupom = False
        Exit Function
    End If
    sCodCai = F_BuscaUmCampo("Cod_Cai", TAB_VENFEC, "num_ven='" & sNum_ven & "'")
    
    'P_cancelaCupon F_BuscaUmCampo("seq_cup", "tab_cupons", "num_ven='" & sNum_ven & "' and cancelado_cup='N'")
   ' strSQL = "update tab_saidas set serie_sai='CANC' where num_ven='" & sNum_ven & "'"
   ' Conn.Execute strSQL
    
    If MsgBox("Confirma exclusão da venda?", vbYesNo + vbQuestion) = vbNo Then
        Exit Function
    End If
    
    strSQL = "select * from " & TAB_ITVENDAS & " where num_ven='" & sNum_ven & "'  "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
        
            strSQL = "update tab_PRODUTOS set er_pro=er_pro + " & F_GravaDec(CStr(!qtd_ite))
            strSQL = strSQL & " where cod_pro='" & !cod_pro & "' "
            
            Conn.Execute strSQL
        
            .MoveNext
        Loop
        .Close
    End With
    
    
    strSQL = "select dataefet_ven,vf.*,avista_for,fp.tipo_cai from " & TAB_VENFEC & "  as vf inner join " & TAB_FORMASPAG & " as fp "
    strSQL = strSQL & " on vf.cod_for=fp.cod_for "
    strSQL = strSQL & " inner join " & TAB_VENDAS & " as tv on tv.num_ven=vf.num_ven "
    strSQL = strSQL & " where vf.num_ven='" & sNum_ven & "' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
       
            
            If !dataefet_ven = Date And !tipo_cai = "DIN" Then
                entrada = 0
                saida = !valor_ven
                
            Else
                entrada = !valor_ven
                saida = !valor_ven
                
            End If
            
                
                  strSQL = F_GravaCai("D", Date, CStr(!valor_ven), "Venda " & !num_ven & " Cancelada", !num_ven, _
                 !tipo_cai, "01", _
                  entrada, saida, sCodCai, strCod_Tur, "000", "S")
                  
       
            
            .MoveNext
        Loop
        .Close
    End With
    strSQL = "select * from tab_doc where num_ven='" & sNum_ven & "' and esp_doc='V' "
                abreObjetoAdo rPesq, strSQL
                With rPesq
                    Do While Not .EOF
                        strSQL = "delete from tab_itdoc where seq_doc='" & !seq_doc & "'"
                        Conn.Execute strSQL
                        strSQL = "delete from tab_cheques where seq_doc='" & !seq_doc & "'"
                        Conn.Execute strSQL
                        strSQL = "delete from tab_doc where seq_doc='" & !seq_doc & "'"
                        Conn.Execute strSQL
                        .MoveNext
                    Loop
                End With
    
    strSQL = "update tab_vendas set situacao_ven=5, "
    strSQL = strSQL & " cancelavenda_usu='" & strCod_Usu & "' "
    strSQL = strSQL & " where num_ven='" & sNum_ven & "' "
    Conn.Execute strSQL '5 = venda cancelada
    F_CancelaVendaAposCupom = True
    Exit Function
trata_Erro:
    P_MostraErro
End Function


Function GravaTrocaAvulsa(cliente As String, vendedor As String, obs As String, _
    valor_in As String, valor_out As String, desconto As String, telefone As String) As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_trocaavulsa", "id", "", "")
    
    strSQL = "insert into tab_trocaavulsa(id,data,cod_usu,cliente,vendedor,obs,"
    
    strSQL = strSQL & "valor_in,valor_out,telefone,desconto)values("
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'" & cliente & "',"
    strSQL = strSQL & "'" & vendedor & "',"
    strSQL = strSQL & "'" & obs & "',"
    
    strSQL = strSQL & F_GravaDec(valor_in) & ","
    strSQL = strSQL & F_GravaDec(valor_out) & ","
    strSQL = strSQL & "'" & telefone & "',"
    strSQL = strSQL & F_GravaDec(desconto) & ")"
    Conn.Execute strSQL



    GravaTrocaAvulsa = sSeq
Exit Function
trata_Erro:

End Function

Sub GravaTrocaAvulsaIt(tipo As String, id_tro As String, codigo As String, qtd As String, unit As String, total As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_trocaavulsait(id_tro,codigo,qtd,unit,total,tipo)values("
    strSQL = strSQL & id_tro & ","
    strSQL = strSQL & "'" & codigo & "',"
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaDec(unit) & ","
    strSQL = strSQL & F_GravaDec(total) & ","
    strSQL = strSQL & "'" & tipo & "')"
    Conn.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
Function GravaVeiculoVenda(venda As String, placa As String, descricao As String, _
    abertura As String, diagnostico As String, status As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "insert into veiculo_venda(venda,placa,modelo,descricao,abertura,diagnostico,status)values("
    strSQL = strSQL & "'" & venda & "',"
    strSQL = strSQL & "'" & placa & "',"
    
    strSQL = strSQL & "'" & descricao & "',"
    strSQL = strSQL & F_GravaDataTempo(abertura) & ","
    strSQL = strSQL & "'" & diagnostico & "',"
    strSQL = strSQL & status & ")"
    Conn.Execute strSQL
    GravaVeiculoVenda = True
    Exit Function
trata_Erro:
    GravaVeiculoVenda = False
End Function

Function AlteraVeiculoVenda(venda As String, placa As String, descricao As String, _
    abertura As String, diagnostico As String, status As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "insert update veiculo_venda set " '(venda,placa,modelo,descricao,abertura,diagnostico,status)values("
    
    strSQL = strSQL & "placa='" & placa & "',"
    
    strSQL = strSQL & "descricao='" & descricao & "',"
    strSQL = strSQL & "abertura=" & F_GravaDataTempo(abertura) & ","
    strSQL = strSQL & "diagnostico='" & diagnostico & "',"
    strSQL = strSQL & "status=" & status & " "
    strSQL = strSQL & " where venda='" & venda & "'"
    Conn.Execute strSQL
    AlteraVeiculoVenda = True
    Exit Function
trata_Erro:
    AlteraVeiculoVenda = False
End Function


Sub P_imprimeTroca(seq_tro As String, obj As Object)
'    On Error Resume Next
'    Dim strSQL As String
'    Dim sSEQ As String
'    Dim iCount As Integer
'    Dim cQtd_Ant As Currency
'    Dim cQtd_Atual As Currency
'    Dim sPro As String
'    Dim sCod_Pro As String
'    Dim cFinal As Currency
'    Dim sSeq_ite As String
'    Dim sUnit_Ite As String
'    Dim rPesq As New ADODB.Recordset
'    Dim rItens As New ADODB.Recordset
'    strSQL = "select tv.*,tt.* from tab_vendas as tv, tab_trocas as tt "
'    strSQL = strSQL & " where tv.num_ven=tt.num_ven "
'    strSQL = strSQL & " and seq_tro='" & seq_tro & "' "
'    abreObjetoAdo rPesq, strSQL
'
'    With rPesq
'
'
'        P_CabRelTimbr "TROCA REFERENTE À VENDA " & !num_ven & " EM " & !dataefet_ven, obj
'        P_ImprX obj, "Venda: " & !num_ven, 0
'        P_ImprX obj, "" & !nomecli_ven, 30, False
'        P_ImprX obj, "" & !dataefet_ven, 0
'        P_ImprX obj, "" & F_BuscaUmCampo("guerra_fun", "tab_funcionarios", "cod_fun='" & !fechavenda_usu & "'"), 45
'        P_ImprX obj, "" & FormatVal(!valor_ven), 100
'        P_ImprX obj, "" & FormatVal(!desconto_ven), 160, False
'        P_ImprLinha obj
'        'P_ImprX obj, "Inicial: " & lblInicial.caption, 0
'        'P_ImprX obj, "À Vista: " & lblAvIsta.caption, 30
'        'P_ImprX obj, "A Prazo: " & lblAPrazo.caption, 60
'        'P_ImprX obj, "Financeira: " & lblFinan.caption, 90
'
'
'        P_ImprX obj, "Desconto: R$ " & FormatVal(!desconto_ven), 130
'        P_ImprX obj, "Final: " & FormatVal(!final_ven), 170, False
'        P_ImprLinha obj, "="
'
'        strSQL = "select tp.*, ti.*, from tab_produtos as tp, tab_ittrocas as ti where tp.cod_pro = ti.cod_pro "
'        strSQL = strSQL & " and ti.seq_tro='" & seq_tro & "'"
'        abreObjetoAdo rItens, strSQL
'
''        With rItens
'            obj.FontBold = True
'            P_ImprX obj, "...", 0
'            P_ImprX obj, "Cód", 5
'            P_ImprX obj, "Descrição", 24
'            P_ImprX obj, "Qtd", 120, True, True
'            P_ImprX obj, "Unit", 135, True, True
'            P_ImprX obj, "Total", 150, True, True
'            P_ImprX obj, "Final", 165, True, True
'            P_ImprX obj, "Inicial", 180, True, True
'            P_ImprX obj, "Ref:", 182, False
'            obj.FontBold = False
'            P_ImprLinha obj
'            Do While Not .EOF
'
'                sPro = !nome_pro
'                cQtd_Atual = !qtdatual_itt
'                cQtd_Ant = !qtdant_itt
'
''                P_ImprX obj, .Text, 0
''                P_ImprX obj, .Text, 5
''                P_ImprX obj, .Text, 24
''                P_ImprX obj, .Text, 120, True, True
''                P_ImprX obj, .Text, 135, True, True
''                P_ImprX obj, .Text, 150, True, True
''                P_ImprX obj, .Text, 165, True, True
''                P_ImprX obj, .Text, 180, True, True
''                P_ImprX obj, .Text, 182, False
'                obj.FontBold = False
'                obj.ForeColor = vbBlack
'            'Next
'            P_ImprLinha obj
'            P_ImprX obj, "A Abater: " & lblAbater.caption, 10
'            P_ImprX obj, "A Acrescer: " & lblAcresc.caption, 60
'            P_ImprX obj, "Desconto: " & txtbxDesconto.Text, 110
'            P_ImprX obj, "Final Previsto: " & lblFinalPrev.caption, 160, False
'            P_ImprLinha obj, "_"
'        End With
'    End With
End Sub

Sub AlteraCabvenda(num_ven As String)
    frmAlteraVendaDadosCliente.tag = num_ven
    frmAlteraVendaDadosCliente.Show 1
    
    
    
    
End Sub
Sub AlteraStatusVen(num_ven As String, status_ven As String)
    Dim strSQL As String
    strSQL = "update " & TAB_VENDAS & " set status_ven='" & status_ven & "' "
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    Conn.Execute strSQL
    
    
End Sub

Sub imprimeLogo(obj As Object, Optional x As Single, Optional y As Single, Optional wid As Single = 0)
    Dim path As String
    Dim Altura As Single
    Dim p As New StdPicture
    On Error Resume Next
    path = App.path & "\pic\logoprint_nfe.jpg"
    If Dir(path) = "" Then
        path = App.path & "\pic\logoprint.jpg"
    End If
    obj.PaintPicture LoadPicture(path), x, y
    Set p = LoadPicture(path)
    Altura = (obj.ScaleY(p.height) / 100) + 5
    P_PrintXY obj, "", 0, Altura, False
    
    
    obj.Print ""
    obj.Print
    
    Exit Sub
trata_Erro:
    MsgBox Err.Description
End Sub

Sub geraManifesto()
    
End Sub
